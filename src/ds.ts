import { List } from "dattatable";
import { Components, Graph, Site, SPTypes, Types, v2, Web } from "gd-sprest-bs";
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
    // License Status
    private static _hasLicense: boolean = false;
    static HasLicense(): boolean { return this._hasLicense; }
    private static initLicense(): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve) => {
            // Get the graph token
            Graph.getAccessToken().execute(token => {
                // Get the current user licenses
                Graph({
                    accessToken: token.access_token,
                    url: "me?$select=displayName,email,id,assignedLicenses,assignedPlans"
                }).execute(user => {
                    // Parse the plans
                    let assignedPlans: any[] = user["assignedPlans"];
                    for (let i = 0; i < assignedPlans.length; i++) {
                        // See if they have a power apps license assigned
                        if (assignedPlans[i].service.indexOf("Power") >= 0) {
                            // Set the flag
                            this._hasLicense = true;
                            break;
                        }
                    }

                    // Resolve the request
                    resolve(null);
                }, resolve);
            }, resolve);
        });
    }

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
                        "Permission", "Owners/Title", "Owners/Id", "Owners/EMail"
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

    // Loads the current permission for a site
    static loadSitePermissions(url: string): PromiseLike<Types.Microsoft.Graph.permission[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the site id
            Site(url).query({ Select: ["Id"] }).execute(site => {
                // Get the site permissions
                v2.sites(site.Id).permissions().execute(permissions => {
                    // Resolve the request
                    resolve(permissions.results);
                });
            }, reject);
        });
    }

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
            // Init the license status
            this.initLicense(),
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