import * as React from "react";
import { Text, tokens, makeStyles } from "@fluentui/react-components";

export interface HeaderProps {
  title: string;
  subtitle: string;
}

const useStyles = makeStyles({
  header: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    padding: "28px 24px 18px",
    background: "linear-gradient(145deg, #0f172a 0%, #1d4ed8 100%)",
    color: "#ffffff",
  },
  eyebrow: {
    fontSize: tokens.fontSizeBase200,
    letterSpacing: "0.08em",
    textTransform: "uppercase",
    opacity: 0.8,
  },
  title: {
    margin: 0,
    fontSize: "28px",
    lineHeight: 1.2,
    fontWeight: tokens.fontWeightSemibold,
  },
  subtitle: {
    fontSize: tokens.fontSizeBase300,
    lineHeight: 1.6,
    opacity: 0.9,
  },
});

const Header: React.FC<HeaderProps> = (props: HeaderProps) => {
  const { title, subtitle } = props;
  const styles = useStyles();

  return (
    <section className={styles.header}>
      <Text className={styles.eyebrow}>PowerPoint Resource Link</Text>
      <h1 className={styles.title}>{title}</h1>
      <Text className={styles.subtitle}>{subtitle}</Text>
    </section>
  );
};

export default Header;
