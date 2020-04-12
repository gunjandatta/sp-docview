import { Helper } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    WebPartCfg: [{
        FileName: "sp-docview.webpart",
        Group: "Demo",
        XML: Helper.WebPart.generateScriptEditorXML({
            chromeType: "TitleOnly",
            description: Strings.ProjectDescription,
            title: Strings.ProjectName,
            content: [
                '&lt;div id="`+ Strings.AppElementId + `"&gt;&lt;/div&gt;',
                '&lt;div id="`+ Strings.AppElementId + `Cfg" style="display: none;"&gt;&lt;/div&gt;',
                '&lt;script type="text/javascript" src="`+ Strings.SolutionUrl + `"&gt;&lt;/script&gt;',
                '&lt;script type="text/javascript"&gt;SP.SOD.executeOrDelayUntilScriptLoaded(function() { `+ Strings.GlobalVariable + `.init(); }, \'sp-docview\');&lt;/script&gt;'
            ].join('\n')
        })
    }]
});