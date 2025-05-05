import * as React from "react";
import Header from "./Header";
import TailwindFinancialModelChat from "../../client/components/TailwindFinancialModelChat";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App: React.FC<AppProps> = () => {
  // Apply global monospace font style
  const monoStyle = {
    fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace',
  };

  return (
    <div className="h-screen flex flex-col bg-transparent" style={monoStyle}>
      <Header />
      <div className="flex-1 flex flex-col overflow-auto">
        <TailwindFinancialModelChat />
      </div>
    </div>
  );
};

export default App;
