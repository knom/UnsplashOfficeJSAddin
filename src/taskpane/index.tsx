// "document": "https://microsofteur-my.sharepoint.com/:p:/g/personal/mknor_microsoft_com/EU427BGQCFVEhU8thT4KDFYBLObgxjBqaGor0-ktg8AXGw?e=yMJbnO"
import App from "./components/App/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";

/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Unsplash Photos Add-in";

const render = (Component) => {
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
Office.initialize = () => {
  isOfficeInitialized = true;
  render(App);
};

if ((module as any).hot) {
  (module as any).hot.accept("./components/App/App", () => {
    const NextApp = require("./components/App/App").default;
    render(NextApp);
  });
}