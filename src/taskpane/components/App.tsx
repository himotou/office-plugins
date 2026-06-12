import * as React from "react";
import Header from "./Header";
import { makeStyles } from "@fluentui/react-components";
import LinkBinderPanel from "./LinkBinderPanel";

interface AppProps {
  title: string;
  resourceUrl: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <Header title={props.title} subtitle="选择文本后，用资源网站把超链接回填到当前选区。" />
      <LinkBinderPanel initialResourceUrl={props.resourceUrl} />
    </div>
  );
};

export default App;
