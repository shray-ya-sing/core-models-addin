import * as React from "react";
import { makeStyles, FluentProvider, tokens, webDarkTheme } from "@fluentui/react-components";
import Header from "./Header";
import { FinancialModelChat } from "../../client/components/FinancialModelChat";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
  },
  content: {
    flexGrow: 1,
    display: "flex",
    flexDirection: "column",
    overflow: "auto",
    padding: tokens.spacingVerticalM,
  }
});

// Create a custom dark theme based on webDarkTheme with deeper background
const customDarkTheme = {
  ...webDarkTheme,
  colorNeutralBackground1: '#1a1a1a',
  colorNeutralBackground2: '#2a2a2a',
  colorBrandForeground1: '#4cc2ff',
};

const App: React.FC<AppProps> = () => {
  const styles = useStyles();

  return (
    <FluentProvider theme={customDarkTheme}>
      <div className={styles.root}>
        <Header title="Financial Model Assistant" logo="assets/logo-filled.png" />
        <div className={styles.content}>
          <FinancialModelChat />
        </div>
      </div>
    </FluentProvider>
  );
};

export default App;
