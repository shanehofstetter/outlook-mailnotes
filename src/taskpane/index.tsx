import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";

initializeIcons();

let isOfficeInitialized = false;

const title = "Mailnotes";

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <Component title={title} isOfficeInitialized={isOfficeInitialized} itemChangedRegister={itemChangedRegister} />
      </ThemeProvider>
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
