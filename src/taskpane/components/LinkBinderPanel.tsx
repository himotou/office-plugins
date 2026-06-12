import * as React from "react";
import { useMemo, useState } from "react";
import {
  Body1,
  Button,
  Caption1,
  Card,
  CardHeader,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Text,
  makeStyles,
  shorthands,
  tokens,
} from "@fluentui/react-components";
import { Link24Regular, Open24Regular, Save24Regular, Settings24Regular } from "@fluentui/react-icons";
import { getDefaultResourcePageUrl, normalizeAbsoluteUrl, resetResourcePageUrl, saveResourcePageUrl } from "../../shared/config";
import { openLinkBinderDialog } from "../../shared/linkBinder";

interface LinkBinderPanelProps {
  initialResourceUrl: string;
}

type FeedbackState =
  | { intent: "success" | "warning" | "error"; title: string; message: string }
  | null;

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    background: "linear-gradient(180deg, #e2e8f0 0%, #f8fafc 220px)",
  },
  content: {
    display: "grid",
    gap: "16px",
    padding: "20px 18px 32px",
  },
  card: {
    ...shorthands.borderRadius(tokens.borderRadiusXLarge),
    boxShadow: "0 12px 36px rgba(15, 23, 42, 0.08)",
  },
  fieldActions: {
    display: "flex",
    gap: "10px",
    flexWrap: "wrap",
    marginTop: "16px",
  },
  helperText: {
    color: tokens.colorNeutralForeground3,
    lineHeight: 1.6,
  },
  usageList: {
    display: "grid",
    gap: "10px",
    marginTop: "8px",
  },
  usageItem: {
    display: "grid",
    gridTemplateColumns: "28px 1fr",
    gap: "12px",
    alignItems: "start",
  },
  usageIndex: {
    width: "28px",
    height: "28px",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    borderRadius: "999px",
    backgroundColor: "#dbeafe",
    color: "#1d4ed8",
    fontWeight: tokens.fontWeightSemibold,
  },
  footerHint: {
    color: tokens.colorNeutralForeground3,
  },
});

const LinkBinderPanel: React.FC<LinkBinderPanelProps> = ({ initialResourceUrl }) => {
  const styles = useStyles();
  const [resourceUrl, setResourceUrl] = useState(initialResourceUrl);
  const [feedback, setFeedback] = useState<FeedbackState>(null);
  const [isLaunching, setIsLaunching] = useState(false);

  const defaultResourceUrl = useMemo(() => getDefaultResourcePageUrl(window.location.origin), []);

  const persistResourceUrl = (): string => {
    const normalizedUrl = normalizeAbsoluteUrl(resourceUrl);
    saveResourcePageUrl(normalizedUrl);
    setResourceUrl(normalizedUrl);
    return normalizedUrl;
  };

  const handleSave = () => {
    try {
      persistResourceUrl();
      setFeedback({
        intent: "success",
        title: "资源站地址已保存",
        message: "之后点击“绑定资源”会打开这个页面。页面只要继续通过 postMessage 返回资源链接即可复用。",
      });
    } catch (error) {
      setFeedback({
        intent: "error",
        title: "地址格式不正确",
        message: "请输入完整的绝对地址，例如 https://example.com/resource-picker 。",
      });
    }
  };

  const handleReset = () => {
    const nextUrl = resetResourcePageUrl(window.location.origin);
    setResourceUrl(nextUrl);
    setFeedback({
      intent: "warning",
      title: "已恢复默认地址",
      message: "默认会打开当前插件自带的示例资源页，你可以继续改成自己的资源网站地址。",
    });
  };

  const handleLaunch = async () => {
    setIsLaunching(true);
    setFeedback(null);

    try {
      persistResourceUrl();
      await openLinkBinderDialog();
    } catch (error) {
      setFeedback({
        intent: "error",
        title: "打开绑定窗口失败",
        message: error instanceof Error ? error.message : "请确认当前已在 PowerPoint 中选中文本。",
      });
    } finally {
      setIsLaunching(false);
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.content}>
        {feedback ? (
          <MessageBar intent={feedback.intent}>
            <MessageBarBody>
              <MessageBarTitle>{feedback.title}</MessageBarTitle>
              {feedback.message}
            </MessageBarBody>
          </MessageBar>
        ) : null}

        <Card className={styles.card}>
          <CardHeader
            image={<Settings24Regular />}
            header={
              <Body1>
                <strong>资源站配置</strong>
              </Body1>
            }
            description={<Caption1>这里配置真正打开的资源关联页面地址。</Caption1>}
          />

          <Field
            label="资源页面 URL"
            hint={`默认地址：${defaultResourceUrl}`}
            validationMessage={resourceUrl ? undefined : "请输入要打开的资源页地址"}
          >
            <Input value={resourceUrl} onChange={(_, data) => setResourceUrl(data.value)} appearance="filled-lighter" />
          </Field>

          <div className={styles.fieldActions}>
            <Button icon={<Save24Regular />} appearance="primary" onClick={handleSave}>
              保存地址
            </Button>
            <Button icon={<Open24Regular />} appearance="secondary" onClick={handleLaunch} disabled={isLaunching}>
              绑定资源
            </Button>
            <Button appearance="subtle" onClick={handleReset}>
              恢复默认
            </Button>
          </div>
        </Card>

        <Card className={styles.card}>
          <CardHeader
            image={<Link24Regular />}
            header={
              <Body1>
                <strong>使用流程</strong>
              </Body1>
            }
            description={<Caption1>尽量贴近你现在 OnlyOffice 插件的操作路径。</Caption1>}
          />

          <div className={styles.usageList}>
            <div className={styles.usageItem}>
              <div className={styles.usageIndex}>1</div>
              <Text>先在 PowerPoint 中选中一段文字，再点击功能区里的“绑定资源”。</Text>
            </div>
            <div className={styles.usageItem}>
              <div className={styles.usageIndex}>2</div>
              <Text>插件会弹出一个窗口，里面可以直接嵌你的资源关联网站。</Text>
            </div>
            <div className={styles.usageItem}>
              <div className={styles.usageIndex}>3</div>
              <Text>资源页确认后把链接发回插件，插件会自动给原来的文字选区设置超链接。</Text>
            </div>
          </div>
        </Card>

        <Card className={styles.card}>
          <CardHeader
            header={
              <Body1>
                <strong>消息协议</strong>
              </Body1>
            }
            description={<Caption1>为了兼容你现有页面，插件会监听和 OnlyOffice 类似的数据结构。</Caption1>}
          />
          <Text className={styles.helperText}>
            资源页确认后向父页面发送：
            {" "}
            <code>
              {`{"type":"resource-selected","url":"https://...","resolvedUrl":"https://...","title":"资源标题","text":"展示文字"}`}
            </code>
          </Text>
          <Text className={styles.footerHint}>
            <code>resolvedUrl</code> 可选；如果没有，会直接使用 <code>url</code>。<code>title</code> 或 <code>text</code>{" "}
            会优先作为超链接提示文字。
          </Text>
        </Card>
      </div>
    </div>
  );
};

export default LinkBinderPanel;
