import { List } from "dattatable";
import { Components, Graph, List as SPList, Site, SPTypes, Types, v2, Web } from "gd-sprest-bs";
import { Security } from "./security";
import Strings from "./strings";

/**
 * Flow Properties
 */
export interface IFlowProps {
    appName: string;
    id: number;
    permission: string;
    token?: string;
    type: "Add" | "Remove";
    url: string;
}

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
            Graph.getAccessToken(Strings.CloudEnv).execute(token => {
                // Get the current user licenses
                Graph({
                    accessToken: token.access_token,
                    cloud: Strings.CloudEnv,
                    url: "me?$select=displayName,email,id,assignedLicenses,assignedPlans"
                }).execute(user => {
                    // Parse the plans
                    let assignedPlans: any[] = user["assignedPlans"];
                    for (let i = 0; i < assignedPlans.length; i++) {
                        // See if they have a power apps license assigned
                        let plan = assignedPlans[i];
                        if (plan.service.indexOf("Power") >= 0 && plan.capabilityStatus == "Enabled") {
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
                        "Id", "Title", "ClientId", "OwnersId", "SiteUrls", "Status",
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

    // Runs a flow
    private static _token: string = null;
    static runFlow(flowProps: IFlowProps): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the site id
            Site(Strings.SourceUrl).query({ Select: ["Id"] }).execute(site => {
                let flowId = flowProps.type == "Add" ? Strings.Flows.Add : Strings.Flows.Remove;

                // Run the flow
                SPList.runFlow({
                    id: flowId,
                    list: Strings.Lists.Main,
                    cloudEnv: Strings.CloudEnv,
                    token: this._token,
                    webUrl: Strings.SourceUrl,
                    data: {
                        AppName: flowProps.appName,
                        ID: flowProps.id,
                        fileName: "",
                        itemUrl: "",
                        Permission: flowProps.permission,
                        SiteId: site.Id,
                        SiteUrl: flowProps.url,
                    }
                }).then(results => {
                    // Save the token
                    this._token = results.flowToken;

                    if (results.executed) {
                        // Resolve w/ the flow token
                        resolve();
                    } else {
                        // Reject the request
                        console.error("Error running the flow: " + results.errorMessage, results.errorDetails);
                        reject(results.errorMessage);
                    }
                });
            });
        });
    }
}