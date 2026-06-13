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

function isPowerPointApiSupported(): boolean {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
}

function getUserFacingErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message) {
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
    const selection = context.presentation.getSelectedTextRangeOrNullObject();
    selection.load(["text", "start", "length"]);

    const selectionFrame = selection.getParentTextFrame();
    const parentShape = selectionFrame.getParentShape();
    parentShape.load("id");

    const parentSlide = parentShape.getParentSlideOrNullObject();
    parentSlide.load("id");

    const hyperlinks = selection.hyperlinks;
    hyperlinks.load("items/address,screenTip");

    await context.sync();

    if ((selection as OfficeExtension.ClientObject & { isNullObject?: boolean }).isNullObject) {
      throw new Error("请先在 PowerPoint 中选中一段文本，再点击“绑定资源”。");
    }

    if ((parentSlide as OfficeExtension.ClientObject & { isNullObject?: boolean }).isNullObject) {
      throw new Error("当前选区不在普通幻灯片文本中，暂时无法绑定超链接。");
    }

    if (selection.length < 0) {
      throw new Error("当前光标位置无法识别，请重新进入文本编辑状态后再试。");
    }

    const existingHyperlink = hyperlinks.items[0];

    return {
      slideId: parentSlide.id,
      shapeId: parentShape.id,
      start: selection.start,
      length: selection.length,
      text: selection.text,
      isInsertionPoint: selection.length === 0,
      hyperlinkAddress: existingHyperlink?.address,
      hyperlinkScreenTip: existingHyperlink?.screenTip,
    };
  });
}

async function applyHyperlinkToSelection(selection: SelectionSnapshot, resource: ResourceSelection): Promise<void> {
  const targetUrl = (resource.resolvedUrl || resource.url || "").trim();
  const displayText = getResourceDisplayText(selection, resource);

  if (!targetUrl) {
    throw new Error("资源页没有返回可用链接。");
  }

  if (!displayText) {
    throw new Error("当前没有可插入的超链接文字，请在资源页返回 title 或 text。");
  }

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItem(selection.slideId);
    const shape = slide.shapes.getItem(selection.shapeId);

    if (selection.length === 0) {
      const insertionPoint = shape.textFrame.textRange.getSubstring(selection.start, 0);
      insertionPoint.text = displayText;
      await context.sync();

      const insertedRange = shape.textFrame.textRange.getSubstring(selection.start, displayText.length);
      insertedRange.setHyperlink({
        address: targetUrl,
        screenTip: resource.title || resource.text || displayText,
      });
      insertedRange.setSelected();
      await context.sync();
      return;
    }

    const textRange = shape.textFrame.textRange.getSubstring(selection.start, selection.length);
    textRange.setHyperlink({
      address: targetUrl,
      screenTip: resource.title || resource.text || selection.text,
    });
    textRange.setSelected();
    await context.sync();
  });
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
      "这个功能依赖 PowerPointApi 1.10。请在较新的 Microsoft 365 PowerPoint 中使用。"
    );
    return;
  }

  try {
    activeSelection = await captureSelectionSnapshot();
    activeMode = activeSelection.hyperlinkAddress ? "edit-hyperlink" : "create-hyperlink";
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
