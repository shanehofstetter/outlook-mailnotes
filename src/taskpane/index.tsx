import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { PartialTheme, ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";

initializeIcons();

let isOfficeInitialized = false;

const title = "Mailnotes";

const lightTheme: PartialTheme = {
};

const darkTheme: PartialTheme = {
  semanticColors: {
    bodyBackground: '#1E1E1E',
    bodyText: 'white',
  },
};

// Office.context.officeTheme seems not to be supported for Outlook
// this would provide us with the colors to use..
// detect if os prefers dark mode via media query as a workaround, probably does not reflect theme configured in outlook itself correctly
const useDarkMode = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;

const render = (Component) => {
  const theme = useDarkMode ? darkTheme : lightTheme;

  ReactDOM.render(
    <AppContainer>
      <div style={{ width: '100%', height: '100%', backgroundColor: theme.semanticColors?.bodyBackground }}>
        <ThemeProvider theme={theme} style={{ padding: '10px 20px' }}>
          <Component
            title={title}
            isOfficeInitialized={isOfficeInitialized}
            itemChangedRegister={itemChangedRegister}
            theme={theme}
          />
        </ThemeProvider>
      </div>
    </AppContainer>,
    document.getElementById("container")
  );
};

let itemChangedHandler: (type: Office.EventType) => void;
const itemChangedRegister = (f: (type: Office.EventType) => void) => {
  itemChangedHandler = f;
};

/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);

  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChangedHandler);
});

render(App);

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
