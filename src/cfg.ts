import { Helper } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    WebPartCfg: [{
        FileName: "sp-docview.webpart",
        Group: "Demo",
        XML: `<?xml version="1.0" encoding="utf-8"?>
<webParts>
    <webPart xmlns="http://schemas.microsoft.com/WebPart/v3">
        <metaData>
            <type name="Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
            <importErrorMessage>$Resources:core,ImportantErrorMessage;</importErrorMessage>
        </metaData>
        <data>
            <properties>
                <property name="Title" type="string">` + Strings.ProjectName + `</property>
                <property name="Description" type="string">`+ Strings.ProjectDescription + `</property>
                <property name="ChromeType" type="chrometype">TitleOnly</property>
                <property name="Content" type="string">
                    &lt;div id="`+ Strings.AppElementId + `"&gt;&lt;/div&gt;
                    &lt;div id="`+ Strings.AppElementId + `Cfg" style="display: none;"&gt;&lt;/div&gt;
                    &lt;script type="text/javascript" src="`+ Strings.SolutionUrl + `"&gt;&lt;/script&gt;
                    &lt;script type="text/javascript"&gt;SP.SOD.executeOrDelayUntilScriptLoaded(function() { `+ Strings.GlobalVariable + `.init(); }, 'sp-docview');&lt;/script&gt;
                </property>
            </properties>
        </data>
    </webPart>
</webParts>`
    }]
});