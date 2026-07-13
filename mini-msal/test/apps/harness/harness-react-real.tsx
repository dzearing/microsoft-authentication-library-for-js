import * as msal from "@azure/msal-browser";
import * as msalReact from "@azure/msal-react";
import { setupReactHarness } from "./react-setup.js";

setupReactHarness(msal, msalReact, "real");
