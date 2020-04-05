import { Components, ContextInfo, Helper, Icons, IconTypes } from "gd-sprest-bs";
import { ILibraryItem } from "./wp";

/**
 * Document View
 */
export class DocView {
    private _el: HTMLElement = null;
    private _items: Array<ILibraryItem> = null;

    // Constructor
    constructor(el: HTMLElement, items: Array<ILibraryItem> = []) {
        // Save the parameters
        this._items = items;

        // Create the element
        this._el = document.createElement("div");
        this._el.classList.add("row");
        el.appendChild(this._el);

        // Initialize the callout
        Helper.SP.CalloutManager.init().then(() => {
            // Render the component
            this.render();
        });
    }

    // Create a callout for the card
    private createCallout(item: ILibraryItem, fileName: string, elCard: Element) {
        // Set the view urls
        let downloadUrl = item.File.ServerRelativeUrl + "?d=w" + item.File.UniqueId;
        let previewUrl = ContextInfo.siteServerRelativeUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc={" + item.File.UniqueId + "}&action=interactivepreview&wdSmallView=1";
        let viewUrl = ContextInfo.siteServerRelativeUrl + "/_layouts/15/Doc.aspx?sourcedoc={" + item.File.UniqueId + "}&file={" + fileName + "}action=default&mobileredirect=true";

        // Create a callout for this icon
        let callout = Helper.SP.CalloutManager.createNewIfNecessary({
            content: "<iframe style='width=100%' src='" + previewUrl + "' width='450px' height='300px' frameborder='0'>File Preview</iframe>",
            contentWidth: 500,
            ID: "docView-" + item.File.UniqueId,
            launchPoint: elCard,
            openOptions: { event: "hover" },
            title: fileName
        });

        // Add an action to view the file
        callout.addAction(Helper.SP.CalloutManager.createAction({
            text: "View",
            onClickCallback: (event, action) => {
                // Show the item in a new window
                window.open(viewUrl, "_blank");
            }
        }));

        // Add the action to download the file
        callout.addAction(Helper.SP.CalloutManager.createAction({
            text: "Download",
            onClickCallback: (event, action) => {
                // Show the item in a new window
                window.open(downloadUrl, "_blank");
            }
        }))
    }

    // Creates a card
    private createCard(item: ILibraryItem): Components.ICardProps {
        // Ensure this is a file
        let fileName = item.File.Name || item.File.Title || "";
        if (fileName.length == 0) { return null; }

        // Create an element for this item
        let el = document.createElement("div");
        el.classList.add("col");
        this._el.appendChild(el);

        // Create a card
        let card = Components.Card({
            el,
            header: {
                content: Icons(IconTypes.FileEarmark, 32, 32)
            },
            body: [{ title: fileName }]
        });

        // Create a callout
        this.createCallout(item, fileName, card.el);
    }

    // Render the table
    private render() {
        // Parse the items
        for (let i = 0; i < this._items.length; i++) {
            // Create the card for this item
            this.createCard(this._items[i]);
        }
    }
}