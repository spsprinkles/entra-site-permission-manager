import { List } from "dattatable";
import { Components, Types, Web } from "gd-sprest-bs";
import { Security } from "./security";
import Strings from "./strings";

/**
 * List Item
 * Add your custom fields here
 */
export interface IListItem extends Types.SP.ListItem {
    ClientId: string;
    Owners: { results: { Id: number; EMail: string; Title: string }[] };
    OwnersId: { results: [] };
    SiteUrls: string;
    Status: string;
}

/**
 * Data Source
 */
export class DataSource {
    // List
    private static _list: List<IListItem> = null;
    static get List(): List<IListItem> { return this._list; }
    private static loadList(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Initialize the list
            this._list = new List<IListItem>({
                listName: Strings.Lists.Main,
                itemQuery: {
                    Expand: ["Owners"],
                    GetAllItems: true,
                    OrderBy: ["Title"],
                    Select: [
                        "Title", "ClientId", "OwnersId", "SiteUrls", "Status",
                        "Owners/Title", "Owners/Id", "Owners/EMail"
                    ],
                    Top: 5000
                },
                onInitError: reject,
                onInitialized: () => {
                    // Load the status filters
                    this.loadStatusFilters().then(() => {
                        // Resolve the request
                        resolve();
                    }, reject);
                }
            });
        });
    }

    // List Items
    static get ListItems(): IListItem[] { return this.List.Items; }

    // Status Filters
    private static _statusFilters: Components.ICheckboxGroupItem[] = null;
    static get StatusFilters(): Components.ICheckboxGroupItem[] { return this._statusFilters; }
    static loadStatusFilters(): PromiseLike<Components.ICheckboxGroupItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the status field
            Web(Strings.SourceUrl).Lists(Strings.Lists.Main).Fields("Status").execute((fld: Types.SP.FieldChoice) => {
                let items: Components.ICheckboxGroupItem[] = [];

                // Parse the choices
                for (let i = 0; i < fld.Choices.results.length; i++) {
                    // Add an item
                    items.push({
                        label: fld.Choices.results[i],
                        type: Components.CheckboxGroupTypes.Switch
                    });
                }

                // Set the filters and resolve the promise
                this._statusFilters = items;
                resolve(items);
            }, reject);
        });
    }

    // Initializes the application
    static init(): PromiseLike<any> {
        // Return a promise
        return Promise.all([
            // Load the security
            Security.init(),
            // Load the list
            this.loadList()
        ]);
    }

    // Refreshes the list data
    static refresh(itemId: number): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Refresh the data
            DataSource.List.refreshItem(itemId).then(resolve, reject);
        });
    }
}