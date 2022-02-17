import App from "./components/App/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";

/* global Office */

initializeIcons();

let isOfficeInitialized = false;

const title = "Unsplash Photos Add-in";

const render = (Component: any) => {
  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <Component title={title} isOfficeInitialized={isOfficeInitialized} />
      </ThemeProvider>
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady().then(() => {
  if (Office.context.host != null) {
    isOfficeInitialized = true;
    render(App);
  }
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App/App", () => {
    const NextApp = require("./components/App/App").default;
    render(NextApp);
  });
}
