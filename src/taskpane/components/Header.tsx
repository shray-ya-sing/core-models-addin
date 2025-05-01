import * as React from "react";
import { Image, Text, tokens, makeStyles } from "@fluentui/react-components";

export interface HeaderProps {
  title: string;
  logo: string;
}

const useStyles = makeStyles({
  header: {
    display: "flex",
    alignItems: "center",
    padding: "8px 12px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground1,
    height: "40px",
  },
  logo: {
    marginRight: "8px",
    height: "20px",
    width: "20px",
  },
  title: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  spacer: {
    flexGrow: 1,
  },
  mode: {
    display: "flex",
    alignItems: "center",
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
});

const Header: React.FC<HeaderProps> = (props: HeaderProps) => {
  const { title } = props;
  const styles = useStyles();

  return (
    <header className={styles.header}>
      <Image className={styles.logo} src="assets/cori-logo.svg" alt="Cori Logo" />
      <Text className={styles.title}>Cori</Text>
      <div className={styles.spacer} />
    </header>
  );
};

export default Header;
