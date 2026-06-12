import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import {
  Badge,
  Body1,
  Button,
  Card,
  CardHeader,
  Field,
  Input,
  Text,
  Textarea,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { bundleIcon, CheckmarkCircleFilled, CheckmarkCircleRegular, Link24Regular } from "@fluentui/react-icons";

const DoneIcon = bundleIcon(CheckmarkCircleFilled, CheckmarkCircleRegular);

interface ResourceRecord {
  id: string;
  title: string;
  summary: string;
  url: string;
  category: string;
}

const MOCK_RESOURCES: ResourceRecord[] = [
  {
    id: "math-001",
    title: "一次函数知识卡",
    summary: "适合课堂讲解中给关键概念加跳转入口。",
    url: "https://resource.example.com/math/linear-function",
    category: "数学",
  },
  {
    id: "chem-002",
    title: "酸碱滴定实验页",
    summary: "包含实验步骤、风险提示和操作视频。",
    url: "https://resource.example.com/chem/titration",
    category: "化学",
  },
  {
    id: "eng-003",
    title: "阅读理解训练包",
    summary: "可跳转到分层阅读与词汇讲解页面。",
    url: "https://resource.example.com/english/reading-pack",
    category: "英语",
  },
];

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    padding: "18px",
    background: "radial-gradient(circle at top, #e0f2fe 0%, #f8fafc 55%)",
    display: "grid",
    gap: "14px",
    alignContent: "start",
  },
  helper: {
    color: tokens.colorNeutralForeground3,
    lineHeight: 1.6,
  },
  resourceList: {
    display: "grid",
    gap: "12px",
  },
  resourceCard: {
    cursor: "pointer",
    border: "1px solid #dbe3ef",
  },
  selectedCard: {
    border: "1px solid #1d4ed8",
    boxShadow: "0 0 0 2px rgba(29, 78, 216, 0.12)",
  },
  resourceMeta: {
    display: "grid",
    gap: "8px",
  },
  actionRow: {
    display: "flex",
    gap: "10px",
    flexWrap: "wrap",
  },
});

function postSelection(selection: { url: string; resolvedUrl: string; title: string; text: string }): void {
  window.parent.postMessage(
    {
      type: "resource-selected",
      ...selection,
    },
    "*"
  );
}

const ResourcePickerApp: React.FC = () => {
  const styles = useStyles();
  const [selectedId, setSelectedId] = useState(MOCK_RESOURCES[0].id);
  const [customUrl, setCustomUrl] = useState(MOCK_RESOURCES[0].url);
  const [customTitle, setCustomTitle] = useState(MOCK_RESOURCES[0].title);
  const [customText, setCustomText] = useState(MOCK_RESOURCES[0].title);
  const [modeLabel, setModeLabel] = useState("新建资源绑定");

  const selectedResource = useMemo(
    () => MOCK_RESOURCES.find((resource) => resource.id === selectedId) || MOCK_RESOURCES[0],
    [selectedId]
  );

  useEffect(() => {
    const handleParentMessage = (event: MessageEvent) => {
      const data = event.data;

      if (!data || (data.type !== "resource-selected" && data.type !== "resource-edit")) {
        return;
      }

      if (typeof data.url === "string" && data.url.trim()) {
        setCustomUrl(data.resolvedUrl || data.url);
      }

      if (typeof data.title === "string" && data.title.trim()) {
        setCustomTitle(data.title);
      }

      if (typeof data.text === "string" && data.text.trim()) {
        setCustomText(data.text);
      }

      setModeLabel(data.type === "resource-edit" ? "编辑已有资源绑定" : "新建资源绑定");
    };

    window.addEventListener("message", handleParentMessage);
    return () => window.removeEventListener("message", handleParentMessage);
  }, []);

  const handleSelectResource = (resource: ResourceRecord) => {
    setSelectedId(resource.id);
    setCustomUrl(resource.url);
    setCustomTitle(resource.title);
    setCustomText(resource.title);
    postSelection({
      url: resource.url,
      resolvedUrl: resource.url,
      title: resource.title,
      text: resource.title,
    });
  };

  const handleConfirmCustomSelection = () => {
    postSelection({
      url: customUrl.trim(),
      resolvedUrl: customUrl.trim(),
      title: customTitle.trim() || selectedResource.title,
      text: customText.trim() || customTitle.trim() || selectedResource.title,
    });
  };

  return (
    <div className={styles.root}>
      <div>
        <Badge appearance="tint">示例资源页</Badge>
        <h2 style={{ marginBottom: 8 }}>这里后面可以替换成你的真实资源网站</h2>
        <Text className={styles.helper}>
          当前页面只是为了先跑通 Office 版链路。你把 taskpane 里的“资源页面 URL”改成自己的站点后，只要站点继续发
          `resource-selected` 消息，这个插件就能直接复用。
        </Text>
        <div style={{ marginTop: 10 }}>
          <Badge color="informative">{modeLabel}</Badge>
        </div>
      </div>

      <div className={styles.resourceList}>
        {MOCK_RESOURCES.map((resource) => {
          const isSelected = selectedId === resource.id;

          return (
            <Card
              key={resource.id}
              className={`${styles.resourceCard} ${isSelected ? styles.selectedCard : ""}`}
              onClick={() => handleSelectResource(resource)}
            >
              <CardHeader
                image={isSelected ? <DoneIcon primaryFill="#1d4ed8" /> : <Link24Regular />}
                header={<Body1><strong>{resource.title}</strong></Body1>}
                description={<Text>{resource.category}</Text>}
              />
              <div className={styles.resourceMeta}>
                <Text>{resource.summary}</Text>
                <Text size={200}>{resource.url}</Text>
              </div>
            </Card>
          );
        })}
      </div>

      <Card>
        <CardHeader
          header={<Body1><strong>自定义返回内容</strong></Body1>}
          description={<Text>可以模拟你的资源站返回不同的标题和链接。</Text>}
        />
        <Field label="资源链接">
          <Input value={customUrl} onChange={(_, data) => setCustomUrl(data.value)} />
        </Field>
        <Field label="资源标题">
          <Input value={customTitle} onChange={(_, data) => setCustomTitle(data.value)} />
        </Field>
        <Field label="展示文字">
          <Textarea value={customText} onChange={(_, data) => setCustomText(data.value)} resize="vertical" />
        </Field>
        <div className={styles.actionRow}>
          <Button appearance="primary" onClick={handleConfirmCustomSelection}>
            发送选中资源
          </Button>
        </div>
      </Card>
    </div>
  );
};

export default ResourcePickerApp;
