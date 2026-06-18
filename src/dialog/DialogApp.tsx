import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import {
  Body1,
  Button,
  Caption1,
  Card,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Spinner,
  Text,
  Textarea,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { ArrowReset24Regular, Checkmark24Regular, Dismiss24Regular } from "@fluentui/react-icons";
import {
  BindingMode,
  DialogApplyLinkMessage,
  DialogCloseMessage,
  DialogReadyMessage,
  ParentToDialogMessage,
  ResourceSelection,
  SelectionSnapshot,
} from "../shared/types";

/* global Office, window */

type FeedbackState =
  | { intent: "error" | "warning"; title: string; message: string }
  | null;

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    backgroundColor: "#f8fafc",
  },
  header: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "20px 20px 16px",
    background: "linear-gradient(135deg, #111827 0%, #1d4ed8 100%)",
    color: "#ffffff",
  },
  toolbar: {
    display: "flex",
    gap: "10px",
    flexWrap: "wrap",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "14px 20px",
    borderBottom: "1px solid #dbe3ef",
    backgroundColor: "#ffffff",
  },
  toolbarButtons: {
    display: "flex",
    gap: "10px",
    flexWrap: "wrap",
  },
  feedbackWrap: {
    padding: "16px 20px 0",
  },
  feedbackMessage: {
    display: "block",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    overflowWrap: "anywhere",
    lineHeight: 1.7,
  },
  infoContent: {
    padding: "20px",
    maxWidth: "800px",
    margin: "0 auto",
    lineHeight: 1.8,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  frameWrap: {
    minHeight: 0,
    flex: "1 1 auto",
    display: "grid",
    gridTemplateColumns: "minmax(260px, 320px) 1fr",
    gap: "16px",
    padding: "16px 20px 20px",
  },
  sideCard: {
    display: "grid",
    alignContent: "start",
    gap: "14px",
    padding: "18px",
    boxShadow: "0 10px 30px rgba(15, 23, 42, 0.08)",
  },
  previewCard: {
    minHeight: 0,
    padding: "0",
    overflow: "hidden",
    boxShadow: "0 10px 30px rgba(15, 23, 42, 0.08)",
  },
  iframe: {
    width: "100%",
    height: "100%",
    minHeight: "540px",
    border: "none",
    backgroundColor: "#ffffff",
  },
  selectedText: {
    padding: "12px 14px",
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: "#eff6ff",
    color: "#1e3a8a",
    lineHeight: 1.6,
  },
  pendingBox: {
    padding: "14px",
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: "#f8fafc",
    border: "1px dashed #cbd5e1",
  },
  loadingState: {
    display: "grid",
    placeItems: "center",
    minHeight: "240px",
    gap: "12px",
    color: tokens.colorNeutralForeground3,
  },
  metaLabel: {
    color: tokens.colorNeutralForeground3,
  },
});

function parseParentMessage(rawMessage: string): ParentToDialogMessage | null {
  try {
    const message = JSON.parse(rawMessage) as ParentToDialogMessage;
    if (!message || typeof message !== "object" || !("type" in message)) {
      return null;
    }
    return message;
  } catch (error) {
    return null;
  }
}

function sendMessageToParent(message: DialogReadyMessage | DialogCloseMessage | DialogApplyLinkMessage): void {
  Office.context.ui.messageParent(JSON.stringify(message));
}

function getUrlQueryParam(name: string): string {
  const params = new URLSearchParams(window.location.search);
  return params.get(name) || "";
}

function buildFallbackResourceSelection(selection: SelectionSnapshot, resourceUrl: string): ResourceSelection {
  return {
    url: resourceUrl,
    resolvedUrl: resourceUrl,
    text: selection.text || "",
    title: selection.text || "",
  };
}

const DialogApp: React.FC = () => {
  const styles = useStyles();
  const infoMode = useMemo(() => getUrlQueryParam("mode") === "info", []);
  const [resourceUrl, setResourceUrl] = useState("");
  const [selection, setSelection] = useState<SelectionSnapshot | null>(null);
  const [resourceSelection, setResourceSelection] = useState<ResourceSelection | null>(null);
  const [feedback, setFeedback] = useState<FeedbackState>(null);
  const [resourceTitle, setResourceTitle] = useState("");
  const [resourceLink, setResourceLink] = useState("");
  const [bindingMode, setBindingMode] = useState<BindingMode>("create-hyperlink");
  const [frameKey, setFrameKey] = useState(0);

  useEffect(() => {
    if (infoMode) {
      return undefined;
    }

    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (args) => {
      const message = parseParentMessage(args.message);

      if (!message) {
        return;
      }

      if (message.type === "selection-context") {
        setSelection(message.selection);
        setResourceUrl(message.resourceUrl);
        setBindingMode(message.mode);
        setFrameKey((currentValue) => currentValue + 1);
        const fallbackSelection = buildFallbackResourceSelection(
          message.selection,
          message.selection.hyperlinkAddress || message.resourceUrl
        );
        setResourceSelection(fallbackSelection);
        setResourceTitle(message.selection.hyperlinkScreenTip || fallbackSelection.title || "");
        setResourceLink(message.selection.hyperlinkAddress || fallbackSelection.resolvedUrl || fallbackSelection.url);
        if (!message.selection.canReadExistingHyperlink && !message.selection.isInsertionPoint) {
          setFeedback({
            intent: "warning",
            title: "当前版本不支持编辑已有超链接",
            message:
              message.selection.existingHyperlinkSupportMessage ||
              "当前版本不支持对当前选区超链接进行编辑，如需编辑，请手动复制原链接后处理。",
          });
        } else {
          setFeedback(null);
        }
        return;
      }

      if (message.type === "apply-error") {
        setFeedback({
          intent: "error",
          title: "设置超链接失败",
          message: message.message,
        });
      }
    });

    sendMessageToParent({ type: "dialog-ready" });
  }, [infoMode]);

  useEffect(() => {
    if (infoMode) {
      return undefined;
    }

    const handleFrameMessage = (event: MessageEvent) => {
      const data = event.data;

      if (!data || (data.type !== "resource-selected" && data.type !== "resource-edit")) {
        return;
      }

      if (typeof data.url !== "string" || !data.url.trim()) {
        setFeedback({
          intent: "warning",
          title: "资源页没有返回有效链接",
          message: "请选择一个资源后再确认，至少需要回传一个 url 字段。",
        });
        return;
      }

      const nextSelection: ResourceSelection = {
        url: data.url.trim(),
        resolvedUrl: typeof data.resolvedUrl === "string" ? data.resolvedUrl.trim() : undefined,
        text: typeof data.text === "string" ? data.text.trim() : undefined,
        title: typeof data.title === "string" ? data.title.trim() : undefined,
      };

      setResourceSelection(nextSelection);
      setResourceTitle(nextSelection.title || nextSelection.text || "");
      setResourceLink(nextSelection.resolvedUrl || nextSelection.url);
      setFeedback(null);
    };

    window.addEventListener("message", handleFrameMessage);
    return () => window.removeEventListener("message", handleFrameMessage);
  }, [infoMode]);

  const pushContextToResourceFrame = (frame: HTMLIFrameElement | null) => {
    if (!frame?.contentWindow || !selection) {
      return;
    }

    frame.contentWindow.postMessage(
      {
        type: bindingMode === "edit-hyperlink" ? "resource-edit" : "resource-selected",
        mode: bindingMode,
        url:
          selection.canReadExistingHyperlink && selection.hyperlinkAddress
            ? selection.hyperlinkAddress
            : resourceLink || resourceUrl,
        resolvedUrl:
          selection.canReadExistingHyperlink && selection.hyperlinkAddress
            ? selection.hyperlinkAddress
            : resourceLink || resourceUrl,
        title:
          selection.canReadExistingHyperlink && selection.hyperlinkScreenTip
            ? selection.hyperlinkScreenTip
            : resourceTitle || selection.text,
        text: selection.text,
      },
      "*"
    );
  };

  const handleClose = () => {
    sendMessageToParent({ type: "close-dialog" });
  };

  const handleApply = () => {
    if (!selection) {
      setFeedback({
        intent: "error",
        title: "没有收到选区信息",
        message: "请关闭窗口后重新选中文本，再点击绑定资源。",
      });
      return;
    }

    const payload: ResourceSelection = {
      url: resourceLink.trim() || resourceSelection?.url || resourceUrl,
      resolvedUrl: resourceLink.trim() || resourceSelection?.resolvedUrl || resourceSelection?.url || resourceUrl,
      title: resourceTitle.trim() || resourceSelection?.title || selection.text,
      text: resourceSelection?.text || selection.text,
    };

    if (!payload.url) {
      setFeedback({
        intent: "warning",
        title: "还没有可绑定的链接",
        message: "请先在资源页里选择一条资源，或者手动填入链接后再确认。",
      });
      return;
    }

    sendMessageToParent({
      type: "apply-link",
      payload,
    });
  };

  const infoTitle = getUrlQueryParam("title") || "提示";
  const infoDescription = getUrlQueryParam("description") || "";

  if (infoMode) {
    return (
      <div className={styles.root}>
        <div className={styles.header}>
          <Text weight="semibold" size={500}>
            {infoTitle}
          </Text>
        </div>
        <div className={styles.toolbar}>
          <div className={styles.toolbarButtons}>
            <Button appearance="primary" icon={<Dismiss24Regular />} onClick={handleClose}>
              关闭
            </Button>
          </div>
        </div>
        <div className={styles.infoContent}>
          <Text>{infoDescription}</Text>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Text weight="semibold" size={500}>
          资源关联窗口
        </Text>
        <Text>
          {selection?.isInsertionPoint
            ? "当前是光标插入模式，确认后会在光标位置插入资源标题并自动设置超链接。"
            : bindingMode === "edit-hyperlink"
              ? "当前选区已经带有超链接，资源页会收到编辑态初始化数据。"
              : selection && selection.canReadExistingHyperlink === false
                ? "当前版本只能读取选中文本，不能读取或编辑该选区原有超链接；你仍可继续绑定新链接。"
                : "选区信息会先被插件缓存，确认后把你选择的资源链接写回 PowerPoint 当前文字。"}
        </Text>
      </div>

      <div className={styles.toolbar}>
        <div>
          <Caption1 className={styles.metaLabel}>当前资源站</Caption1>
          <Body1>{resourceUrl || "等待主页面下发资源地址..."}</Body1>
        </div>
        <div className={styles.toolbarButtons}>
          <Button icon={<ArrowReset24Regular />} appearance="subtle" onClick={() => window.location.reload()}>
            刷新资源页
          </Button>
          <Button icon={<Dismiss24Regular />} appearance="secondary" onClick={handleClose}>
            取消
          </Button>
          <Button icon={<Checkmark24Regular />} appearance="primary" onClick={handleApply}>
            确认
          </Button>
        </div>
      </div>

      {feedback ? (
        <div className={styles.feedbackWrap}>
          <MessageBar intent={feedback.intent}>
            <MessageBarBody>
              <MessageBarTitle>{feedback.title}</MessageBarTitle>
              <Text className={styles.feedbackMessage}>{feedback.message}</Text>
            </MessageBarBody>
          </MessageBar>
        </div>
      ) : null}

      <div className={styles.frameWrap}>
        <Card className={styles.sideCard}>
          <div>
            <Caption1 className={styles.metaLabel}>当前选中文本</Caption1>
            <div className={styles.selectedText}>
              {selection ? selection.text || "当前没有选中文字，将按光标位置插入资源标题。" : "等待宿主返回文本选区..."}
            </div>
          </div>

          <div>
            <Caption1 className={styles.metaLabel}>已选资源</Caption1>
            <div className={styles.pendingBox}>
              <Text>{resourceSelection?.title || resourceSelection?.text || "还没有从资源页收到选择结果"}</Text>
            </div>
          </div>

          <div>
            <Caption1 className={styles.metaLabel}>当前模式</Caption1>
            <div className={styles.pendingBox}>
              <Text>
                {selection?.isInsertionPoint
                  ? "光标插入超链接"
                  : bindingMode === "edit-hyperlink"
                    ? "编辑已有超链接"
                    : selection && selection.canReadExistingHyperlink === false
                      ? "新建超链接（不支持读取已有链接）"
                      : "新建超链接"}
              </Text>
            </div>
          </div>

          <Field label="最终写入的链接">
            <Input value={resourceLink} onChange={(_, data) => setResourceLink(data.value)} />
          </Field>

          <Field label="链接提示文字">
            <Textarea value={resourceTitle} onChange={(_, data) => setResourceTitle(data.value)} resize="vertical" />
          </Field>
        </Card>

        <Card className={styles.previewCard}>
          {resourceUrl ? (
            <iframe
              key={frameKey}
              title="resource-picker"
              src={resourceUrl}
              className={styles.iframe}
              onLoad={(event) => pushContextToResourceFrame(event.currentTarget)}
            />
          ) : (
            <div className={styles.loadingState}>
              <Spinner labelPosition="below" label="正在等待资源页地址和选区上下文..." />
            </div>
          )}
        </Card>
      </div>
    </div>
  );
};

export default DialogApp;
