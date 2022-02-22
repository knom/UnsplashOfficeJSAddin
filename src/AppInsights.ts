import { ApplicationInsights } from "@microsoft/applicationinsights-web";
import { ReactPlugin } from "@microsoft/applicationinsights-react-js";
import { createBrowserHistory } from "history";

const browserHistory = createBrowserHistory(<any>{ basename: "" });
const reactPlugin = new ReactPlugin();
const appInsightsId = process.env.REACT_APP_APPINSIGHTS_API_KEY as string;

const appInsights = new ApplicationInsights({
  config: {
    instrumentationKey: appInsightsId,
    extensions: [reactPlugin],
    extensionConfig: {
      [reactPlugin.identifier]: { history: browserHistory },
    },
  },
});
appInsights.loadAppInsights();
export { reactPlugin, appInsights };
