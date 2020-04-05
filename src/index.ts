import { Helper } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { WebPart } from "./wp";
import Strings from "./strings";

// Create the global variable for this solution
window[Strings.GlobalVariable] = {
    Configuration,
    init: WebPart
}

// Notify SharePoint this script is loaded
Helper.SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("sp-docview");