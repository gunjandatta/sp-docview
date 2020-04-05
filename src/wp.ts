import { SPTypes, Types, WebParts } from "gd-sprest-bs";
import { DocView } from "./docView";
import Strings from "./strings";

// Library Item
export interface ILibraryItem extends Types.SP.ListItemOData { }

/**
 * WebPart
 */
export const WebPart = () => {
    // Create the list webpart
    return WebParts.WPList({
        elementId: Strings.AppElementId,
        cfgElementId: Strings.AppElementId + "Cfg",
        editForm: {
            // Filter the webpart configuration to only query for document libraries
            listQuery: {
                Filter: "BaseTemplate eq " + SPTypes.ListTemplateType.DocumentLibrary
            }
        },
        // Set the list query to include the appropriate file information
        odataQuery: {
            Expand: ["File"],
            Select: ["Id", "File/Name", "File/ServerRelativeUrl", "File/Title", "File/UniqueId"],
            GetAllItems: true,
            Top: 5000
        },
        onRenderItems: (wpInfo, items: Array<ILibraryItem>) => {
            // Parse the items
            new DocView(wpInfo.el, items)
        }
    });
}