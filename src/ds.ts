import { List } from "dattatable";
import { Components, Graph, List as SPList, Site, Types, v2, Web, ContextInfo } from "gd-sprest-bs";
import { Security } from "./security";
import Strings from "./strings";

/**
 * Flow Properties
 */
export interface IFlowProps {
    appId: string;
    appName: string;
    itemId: number;
    ownerEmails: string;
    permission?: string;
    permissionId?: string;
    requestType: "add" | "remove" | "update";
    siteId?: string;
    siteUrl: string;
}

/**
 * List Item
 * Add your custom fields here
 */
export interface IListItem extends Types.SP.ListItem {
    AppId: string;
    Owners: { results: { Id: number; EMail: string; Title: string }[] };
    OwnersId: { results: [] };
    SiteUrls: string;
}

/**
 * Data Source
 */
export class DataSource {
    // Gets the owner emails for an item
    static getOwnerEmails(item: IListItem): string[] {
        // Parse the owners
        let owners = [];
        for (let i = 0; i < item.Owners.results.length; i++) {
            // Add the email
            owners.push(item.Owners.results[i].EMail);
        }

        // Return the owners
        return owners;
    }

    // Gets the site permission for a client id
    static getSitePermission(siteId: string, permissionId: string): PromiseLike<Types.Microsoft.Graph.permission> {
        // Return a promise
        return new Promise((resolve) => {
            // Get the permissions for this site
            v2.sites(siteId).permissions(permissionId).execute(resolve);
        });
    }

    // Gets the site permission for a client id
    static getSitePermissions(siteUrl: string): PromiseLike<{
        siteId: string;
        permissions: Types.Microsoft.Graph.permissionCollection
    }> {
        // Return a promise
        return new Promise((resolve) => {
            // Get the site id
            Site(siteUrl).query({ Select: ["Id"] }).execute(site => {
                // Get the permissions for this site
                v2.sites(site.Id).permissions().execute(permissions => {
                    resolve({ siteId: site.Id, permissions });
                });
            });
        });
    }

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
                    url: "me?$select=displayName,email,id,assignedPlans"
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
                        "Id", "Title", "AppId", "SiteUrls",
                        "OwnersId", "Owners/Title", "Owners/Id", "Owners/EMail"
                    ],
                    Top: 5000
                },
                onInitError: reject,
                onInitialized: resolve
            });
        });
    }

    // List Items
    static get ListItems(): IListItem[] { return this.List.Items; }

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

    // Sees if the user is an owner of an item
    static isOwner(item: IListItem): boolean {
        let isOwner = false;

        // Parse the owners
        for (let i = 0; i < item.OwnersId.results.length; i++) {
            // See if this is the current user
            if (item.OwnersId.results[i] == ContextInfo.userId) {
                // Set the flag
                isOwner = true;
                break;
            }
        }

        // Return the flag
        return isOwner;
    }

    // Refreshes the list data
    static refresh(itemId?: number): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Refresh the data
            itemId ? DataSource.List.refreshItem(itemId).then(resolve, reject) : DataSource.List.refresh().then(resolve, reject);
        });
    }

    // Runs a flow
    private static _token: string = null;
    static runFlow(flowProps: IFlowProps): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the site id
            Site(Strings.SourceUrl).query({ Select: ["Id"] }).execute(site => {
                // Set the site id
                flowProps.siteId = site.Id;

                // Run the flow
                SPList.runFlow({
                    id: Strings.FlowId,
                    list: Strings.Lists.Main,
                    cloudEnv: Strings.CloudEnv,
                    token: this._token,
                    webUrl: Strings.SourceUrl,
                    data: flowProps
                }).then(results => {
                    // Save the token
                    this._token = results.flowToken;

                    // See if the flow was executed successfully
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