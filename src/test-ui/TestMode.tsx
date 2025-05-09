/**
 * TestMode Component
 * Provides a toggle for the test UI
 */
import * as React from "react";
import { useState, useEffect } from "react";
import { ClientQueryProcessor } from "../client/services/ClientQueryProcessor";
import { ClientExcelCommandInterpreter } from "../client/services/ClientExcelCommandInterpreter";
import TestUI from "./TestUI";

// Keyboard shortcut for toggling test mode (Ctrl+Shift+T)
const TEST_MODE_SHORTCUT = { ctrl: true, shift: true, key: "t" };

interface TestModeProps {
  queryProcessor: ClientQueryProcessor;
  commandInterpreter: ClientExcelCommandInterpreter;
  onClose?: () => void;
}

/**
 * TestMode Component
 * Provides a toggle for the test UI
 */
const TestMode: React.FC<TestModeProps> = ({
  queryProcessor,
  commandInterpreter,
  onClose
}) => {
  const [showTestUI, setShowTestUI] = useState(false);

  // Set up keyboard shortcut listener
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (
        event.ctrlKey === TEST_MODE_SHORTCUT.ctrl &&
        event.shiftKey === TEST_MODE_SHORTCUT.shift &&
        event.key.toLowerCase() === TEST_MODE_SHORTCUT.key
      ) {
        setShowTestUI(prev => !prev);
        event.preventDefault();
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => {
      window.removeEventListener("keydown", handleKeyDown);
    };
  }, []);

  // If we have an external onClose handler, we don't need to manage internal state
  // Just render the TestUI directly
  if (onClose) {
    return (
      <TestUI
        queryProcessor={queryProcessor}
        commandInterpreter={commandInterpreter}
        onClose={onClose}
      />
    );
  }
  
  // Only render the test UI if showTestUI is true when using internal state
  if (!showTestUI) {
    return null;
  }

  return (
    <TestUI
      queryProcessor={queryProcessor}
      commandInterpreter={commandInterpreter}
      onClose={() => setShowTestUI(false)}
    />
  );
};

export default TestMode;
