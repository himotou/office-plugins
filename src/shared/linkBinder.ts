import { getStoredResourcePageUrl } from "./config";
import {
  BindingMode,
  DialogApplyLinkMessage,
  DialogCloseMessage,
  DialogReadyMessage,
  ParentToDialogMessage,
  ResourceSelection,
  SelectionSnapshot,
} from "./types";

/* global Office, PowerPoint, window */

let activeDialog: Office.Dialog | null = null;
let activeSelection: SelectionSnapshot | null = null;
let activeMode: BindingMode = "create-hyperlink";

type ShapeWithParent = PowerPoint.Shape & {
  parentGroup?: PowerPoint.Shape;
  isNullObject?: boolean;
};

function isPowerPointApiSupported(): boolean {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.2");
}

function getUserFacingErrorMessage(error: unknown): string {
  if (typeof OfficeExtension !== "undefined" && error instanceof OfficeExtension.Error) {
    if (error.code === "InvalidArgument" || error.code === "InvalidParam") {
      return "当前文本所在对象已变化或不支持直接设置超链接，请重新选中这段文本后再试。";
    }
  }

  if (error instanceof Error && error.message) {
    if (error.message.includes("InvalidParam passed to GetItem(id)")) {
      return "当前文本所在对象已变化或位于组合形状中，插件未能重新定位到原文本框。请重新选中这段文本后再试。";
    }

    return error.message;
  }

  return "绑定资源失败，请重新选择文本后再试。";
}

function getResourceDisplayText(selection: SelectionSnapshot, resource: ResourceSelection): string {
  return (resource.text || resource.title || selection.text || resource.resolvedUrl || resource.url || "").trim();
}

function buildDialogUrl(pathname: string, params?: Record<string, string>): string {
  const baseUrl = window.location.href;
  const normalizedPath = pathname.startsWith("/") ? `.${pathname}` : pathname;
  const url = new URL(normalizedPath, baseUrl);

  if (params) {
    Object.entries(params).forEach(([key, value]) => {
      url.searchParams.set(key, value);
    });
  }

  return url.toString();
}

function displayDialogAsync(url: string, options: Office.DialogOptions): Promise<Office.Dialog> {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(url, options, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
        return;
      }

      reject(new Error(result.error.message));
    });
  });
}

function postMessageToDialog(message: ParentToDialogMessage): void {
  if (!activeDialog) {
    return;
  }

  activeDialog.messageChild(JSON.stringify(message));
}

function closeActiveDialog(): void {
  if (activeDialog) {
    activeDialog.close();
    activeDialog = null;
  }
}

async function showInfoDialog(title: string, description: string): Promise<void> {
  closeActiveDialog();

  const dialogUrl = buildDialogUrl("/dialog.html", {
    mode: "info",
    title,
    description,
  });

  const dialog = await displayDialogAsync(dialogUrl, {
    width: 28,
    height: 24,
    displayInIframe: true,
  });

  activeDialog = dialog;
  dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
    if (!("message" in args)) {
      return;
    }

    const message = parseDialogMessage(args.message);

    if (message?.type === "close-dialog") {
      closeActiveDialog();
    }
  });

  dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
    activeDialog = null;
  });
}
async function captureSelectionSnapshot(): Promise<SelectionSnapshot> {
  return PowerPoint.run(async (context) => {
    // 检测 API 版本支持
    const has12 = Office.context.requirements.isSetSupported("PowerPointApi", "1.2");
    const has15 = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
    const has18 = Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
    const has110 = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");

    console.log("API版本检测:", { has12, has15, has18, has110 });

    if (!has12) {
      throw new Error("当前 Office 版本太低，需要 PowerPointApi 1.2+（Office 2021 或更高版本）。");
    }

    // 获取选中的文本范围
    const selection = context.presentation.getSelectedTextRangeOrNullObject();
    selection.load(["text", "start", "length"]);

    // 获取文本框和父形状
    const selectionFrame = selection.getParentTextFrame();
    const parentShape = selectionFrame.getParentShape();
    parentShape.load("id");

    await context.sync();

    if ((selection as OfficeExtension.ClientObject & { isNullObject?: boolean }).isNullObject) {
      throw new Error("请先在 PowerPoint 中选中一段文本，再点击\"绑定资源\"。");
    }

    if (selection.length < 0) {
      throw new Error("当前光标位置无法识别，请重新进入文本编辑状态后再试。");
    }

    // 获取父幻灯片（兼容 1.2-1.4 版本）
    let parentSlideId: string;
    
    if (has15) {
      // 1.5+ 可以直接获取父幻灯片
      const parentSlide = parentShape.getParentSlideOrNullObject();
      parentSlide.load("id");
      await context.sync();

      if ((parentSlide as OfficeExtension.ClientObject & { isNullObject?: boolean }).isNullObject) {
        throw new Error("当前选区不在普通幻灯片文本中，暂时无法绑定超链接。");
      }
      parentSlideId = parentSlide.id;
    } else {
      // 1.2-1.4 不支持 getParentSlideOrNullObject，需要遍历所有幻灯片查找
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();

      let foundSlideId: string | null = null;
      for (const slide of slides.items) {
        const shapes = slide.shapes;
        shapes.load("items/id");
        await context.sync();

        const matchingShape = shapes.items.find((s) => s.id === parentShape.id);
        if (matchingShape) {
          foundSlideId = slide.id;
          break;
        }
      }

      if (!foundSlideId) {
        throw new Error("当前选区不在普通幻灯片文本中，暂时无法绑定超链接。");
      }
      parentSlideId = foundSlideId;
    }

    // 获取现有超链接（如果有）
    let existingHyperlinkAddress: string | undefined;
    let existingHyperlinkScreenTip: string | undefined;

    if (has110) {
      // 1.10+ 支持读取超链接
      const hyperlinks = selection.hyperlinks;
      hyperlinks.load("items/address,items/screenTip");
      await context.sync();
      const existingHyperlink = hyperlinks.items[0];
      existingHyperlinkAddress = existingHyperlink?.address;
      existingHyperlinkScreenTip = existingHyperlink?.screenTip;
    }

    // 获取形状路径（1.8+ 支持组合形状）
    let shapePath: string[] | undefined;
    if (has18) {
      shapePath = await getShapePath(parentShape, context);
    }

    // 判断 API 版本
    let apiVersion = "1.2";
    if (has110) {
      apiVersion = "1.10";
    } else if (has18) {
      apiVersion = "1.8";
    } else if (has15) {
      apiVersion = "1.5";
    }

    if (!has110) {
      console.log(`当前版本 ${apiVersion} 不支持自动设置超链接，将使用剪贴板方式`);
      console.log(`当前版本 ${apiVersion} 不支持读取当前选区已有超链接`);
    }

    if (has18 && shapePath && shapePath.length > 1) {
      console.log("检测到组合形状，shapePath:", shapePath);
    }

    return {
      slideId: parentSlideId,
      shapeId: parentShape.id,
      shapePath,
      apiVersion,
      canReadExistingHyperlink: has110,
      existingHyperlinkSupportMessage: has110
        ? undefined
        : "当前版本不支持读取或编辑当前选区已有超链接，如需编辑，请手动复制原链接后处理。",
      start: selection.start,
      length: selection.length,
      text: selection.text,
      isInsertionPoint: selection.length === 0,
      hyperlinkAddress: existingHyperlinkAddress,
      hyperlinkScreenTip: existingHyperlinkScreenTip,
    };
  });
}

async function getShapePath(shape: PowerPoint.Shape, context: PowerPoint.RequestContext): Promise<string[]> {
  const has18 = Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
  
  if (!has18) {
    // 1.8 以下不支持组合形状，只返回单个形状 ID
    return [shape.id];
  }

  const path: string[] = [];
  let currentShape = shape as ShapeWithParent;

  currentShape.load("id");
  await context.sync();

  // 尝试获取 level 属性（1.8+ 支持）
  try {
    currentShape.load("level");
    await context.sync();

    // 如果 level > 0，说明在组合形状内
    while (currentShape.level > 0) {
      path.unshift(currentShape.id);
      
      const parentGroup = currentShape.parentGroup;
      if (!parentGroup) break;
      
      parentGroup.load("id,level");
      await context.sync();
      
      currentShape = parentGroup as ShapeWithParent;
    }
    
    path.unshift(currentShape.id);
  } catch (error) {
    // 如果不支持 level 属性，只返回形状本身
    console.log("不支持 level 属性，返回单个形状 ID");
    return [shape.id];
  }

  return path;
}
async function resolveShapeFromSnapshot(
  context: PowerPoint.RequestContext,
  selection: SelectionSnapshot
): Promise<PowerPoint.Shape> {
  const has13 = Office.context.requirements.isSetSupported("PowerPointApi", "1.3");
  const has18 = Office.context.requirements.isSetSupported("PowerPointApi", "1.8");

  const slide = context.presentation.slides.getItem(selection.slideId);

  // 如果没有 shapePath 或不支持 1.8，直接获取顶层形状
  if (!has18 || !selection.shapePath || selection.shapePath.length === 0) {
    return slide.shapes.getItem(selection.shapeId);
  }

  // 1.8+ 且有 shapePath 时尝试组内定位
  const topLevelShape = has13
    ? slide.shapes.getItemOrNullObject(selection.shapePath[0])
    : slide.shapes.getItem(selection.shapePath[0]);

  topLevelShape.load("id,type");
  await context.sync();

  if (has13 && (topLevelShape as ShapeWithParent).isNullObject) {
    throw new Error("原始文本框已经不存在或已被移动，请重新选中目标文本后再试。");
  }

  // 如果只有一层，直接返回
  if (selection.shapePath.length === 1) {
    return topLevelShape;
  }

  // 遍历组合形状路径
  let currentShape = topLevelShape;

  for (let index = 1; index < selection.shapePath.length; index += 1) {
    currentShape.load("type");
    await context.sync();

    if (currentShape.type !== PowerPoint.ShapeType.group) {
      throw new Error("原始文本框结构已变化，无法恢复到原来的组合层级。请重新选中文本后再试。");
    }

    const childShape = has13
      ? currentShape.group.shapes.getItemOrNullObject(selection.shapePath[index])
      : currentShape.group.shapes.getItem(selection.shapePath[index]);

    childShape.load("id,type");
    await context.sync();

    if (has13 && (childShape as ShapeWithParent).isNullObject) {
      throw new Error("原始文本框结构已变化，无法恢复到原来的组合层级。请重新选中文本后再试。");
    }

    currentShape = childShape;
  }

  return currentShape;
}

async function applyHyperlinkToSelection(selection: SelectionSnapshot, resource: ResourceSelection): Promise<void> {
  const has110 = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");

  const targetUrl = (resource.resolvedUrl || resource.url || "").trim();
  const displayText = getResourceDisplayText(selection, resource);

  if (!targetUrl) {
    throw new Error("资源页没有返回可用链接。");
  }

  if (!displayText) {
    throw new Error("当前没有可插入的超链接文字，请在资源页返回 title 或 text。");
  }

  // 1.10 以下版本：复制到剪贴板 + 提示手动粘贴
  if (!has110) {
    await copyToClipboardAndNotify(targetUrl, selection);
    return;
  }

  // 以下是 1.10+ 的完整自动实现
  await PowerPoint.run(async (context) => {
    const shape = await resolveShapeFromSnapshot(context, selection);
    const textFrame = shape.getTextFrameOrNullObject();
    textFrame.load("hasText");
    await context.sync();

    if ((textFrame as OfficeExtension.ClientObject & { isNullObject?: boolean }).isNullObject || !textFrame.hasText) {
      throw new Error("目标对象当前不再是可编辑文本框，请重新选中目标文本后再试。");
    }

    if (selection.length === 0) {
      const insertionPoint = textFrame.textRange.getSubstring(selection.start, 0);
      insertionPoint.text = displayText;
      await context.sync();

      const insertedRange = textFrame.textRange.getSubstring(selection.start, displayText.length);
      insertedRange.setHyperlink({
        address: targetUrl,
        screenTip: resource.title || resource.text || displayText,
      });
      insertedRange.setSelected();
      await context.sync();
      return;
    }

    const textRange = textFrame.textRange.getSubstring(selection.start, selection.length);
    textRange.setHyperlink({
      address: targetUrl,
      screenTip: resource.title || resource.text || selection.text,
    });
    textRange.setSelected();
    await context.sync();
  });
}

async function copyToClipboardAndNotify(url: string, selection: SelectionSnapshot): Promise<void> {
  let clipboardSuccess = false;
  
  try {
    await navigator.clipboard.writeText(url);
    clipboardSuccess = true;
  } catch (error) {
    console.warn("剪贴板 API 失败，将在对话框中显示链接供手动复制", error);
  }
  
  if (clipboardSuccess) {
    const versionHint = selection.apiVersion === "1.8" 
      ? "（已识别组内文本位置）" 
      : "";
    
    const groupHint = selection.shapePath && selection.shapePath.length > 1
      ? "\n💡 检测到您选中的是组合形状内的文本。"
      : "";

    await showInfoDialog(
      "链接已复制到剪贴板",
      `请按以下步骤完成绑定：\n\n` +
      `1. 确认文本仍处于选中状态\n` +
      `2. 按 Cmd+K (Mac) 或 Ctrl+K (Windows)\n` +
      `3. 在弹出的对话框中粘贴链接\n\n` +
      `💡 自动设置超链接需要 PowerPointApi 1.10（Microsoft 365 最新版）${versionHint}${groupHint}`
    );
  } else {
    await showInfoDialog(
      "请手动复制链接",
      `链接地址：\n${url}\n\n` +
      `请手动复制上面的链接，然后：\n\n` +
      `1. 选中目标文本\n` +
      `2. 按 Cmd+K (Mac) 或 Ctrl+K (Windows)\n` +
      `3. 粘贴链接并确认`
    );
  }
}
function parseDialogMessage(rawMessage: string): DialogReadyMessage | DialogCloseMessage | DialogApplyLinkMessage | null {
  try {
    const message = JSON.parse(rawMessage) as DialogReadyMessage | DialogCloseMessage | DialogApplyLinkMessage;

    if (!message || typeof message !== "object" || !("type" in message)) {
      return null;
    }

    return message;
  } catch (error) {
    return null;
  }
}

function registerDialogHandlers(resourceUrl: string): void {
  if (!activeDialog) {
    return;
  }

  activeDialog.addEventHandler(Office.EventType.DialogMessageReceived, async (args) => {
    if (!("message" in args)) {
      return;
    }

    const message = parseDialogMessage(args.message);

    if (!message) {
      return;
    }

    if (message.type === "dialog-ready") {
      if (activeSelection) {
        postMessageToDialog({
          type: "selection-context",
          selection: activeSelection,
          resourceUrl,
          mode: activeMode,
        });
      }

      return;
    }

    if (message.type === "close-dialog") {
      closeActiveDialog();
      return;
    }

    if (message.type !== "apply-link" || !activeSelection) {
      return;
    }

    try {
      await applyHyperlinkToSelection(activeSelection, message.payload);
      closeActiveDialog();
    } catch (error) {
      postMessageToDialog({
        type: "apply-error",
        message: getUserFacingErrorMessage(error),
      });
    }
  });

  activeDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
    activeDialog = null;
  });
}

export async function openLinkBinderDialog(): Promise<void> {
  if (!isPowerPointApiSupported()) {
    await showInfoDialog(
      "当前环境不支持",
      "这个功能需要 PowerPointApi 1.2+（Office 2021 或更高版本）。\n\n自动设置超链接需要 PowerPointApi 1.10（Microsoft 365 最新版）。"
    );
    return;
  }

  try {
    activeSelection = await captureSelectionSnapshot();
    activeMode =
      activeSelection.canReadExistingHyperlink && activeSelection.hyperlinkAddress
        ? "edit-hyperlink"
        : "create-hyperlink";
  } catch (error) {
    await showInfoDialog("请选择文本", getUserFacingErrorMessage(error));
    return;
  }

  closeActiveDialog();

  const resourceUrl = getStoredResourcePageUrl(window.location.href);
  const dialog = await displayDialogAsync(buildDialogUrl("/dialog.html"), {
    width: 48,
    height: 72,
    displayInIframe: true,
  });

  activeDialog = dialog;
  registerDialogHandlers(resourceUrl);
}
