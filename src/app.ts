import { Dashboard } from "dattatable";
import { Components } from "gd-sprest-bs";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { plusSquare } from "gd-sprest-bs/build/icons/svgs/plusSquare";
import * as jQuery from "jquery";
import * as moment from "moment";
import { DataSource, IListItem } from "./ds";
import { Forms } from "./forms";
import { InstallationModal } from "./install";
import { Security } from "./security";
import Strings from "./strings";

/**
 * Main Application
 */
export class App {
    private _dashboard: Dashboard = null;

    // Constructor
    constructor(el: HTMLElement) {
        // Render the dashboard
        this.render(el);
    }

    // Renders the navigation items
    private generateNavItems() {
        // Create the settings menu items
        let itemsEnd: Components.INavbarItem[] = [];
        if (Security.IsAdmin) {
            itemsEnd.push(
                {
                    className: "btn-icon btn-outline-light me-2 p-2 py-1",
                    text: "Settings",
                    iconSize: 22,
                    iconType: gearWideConnected,
                    isButton: true,
                    items: [
                        {
                            text: "App Settings",
                            onClick: () => {
                                // Show the install modal
                                InstallationModal.show(true);
                            }
                        },
                        {
                            text: Strings.Lists.Main + " List",
                            onClick: () => {
                                // Show the FAQ list in a new tab
                                window.open(Strings.SourceUrl + "/Lists/" + Strings.Lists.Main.split(" ").join(""), "_blank");
                            }
                        },
                        {
                            text: Security.AdminGroup.Title + " Group",
                            onClick: () => {
                                // Show the settings in a new tab
                                window.open(Strings.SourceUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + Security.AdminGroup.Id);
                            }
                        },
                        {
                            text: Security.MemberGroup.Title + " Group",
                            onClick: () => {
                                // Show the settings in a new tab
                                window.open(Strings.SourceUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + Security.MemberGroup.Id);
                            }
                        },
                        {
                            text: Security.VisitorGroup.Title + " Group",
                            onClick: () => {
                                // Show the settings in a new tab
                                window.open(Strings.SourceUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + Security.VisitorGroup.Id);
                            }
                        }
                    ]
                }
            );
        }

        // Return the nav items
        return itemsEnd;
    }

    // Returns the Entra icon as an SVG element
    private getEntraIcon(height?, width?, className?) {
        // Set the default values
        if (height === void 0) { height = 32; }
        if (width === void 0) { width = 32; }

        // Get the icon element
        let elDiv = document.createElement("div");
        elDiv.innerHTML = "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='32' viewBox='0 0 18 18'><path d='m17.679,10.013c-.013-.015-3.337-3.818-3.334-3.806-.54-.751-1.484-1.224-2.511-1.224-.028,0-.055.003-.082.004-.856.022-1.553.463-2.085,1.031l3.701,4.182h0s-6.159,3.855-6.159,3.855c0,0-1.168.856-2.495.541.109.068,3.474,2.169,3.474,2.169.499.313,1.146.313,1.645,0,0,0,7.339-4.596,7.489-4.687.826-.501.838-1.512.357-2.066Z'/><path d='m10.113,1.467c-.525-.577-1.598-.683-2.248.043-.189.211-7.325,8.262-7.526,8.489-.568.638-.403,1.615.353,2.081,0,0,1.751,1.097,1.865,1.169.696.436,2.571,1.036,4.454-.096l1.179-.738-3.536-2.213s4.461-5.039,4.467-5.047c1.294-1.416,2.987-1.032,3.613-.735.002.001-2.607-2.937-2.621-2.951Z'/></svg>";
        let icon = elDiv.firstChild as SVGImageElement;
        if (icon) {
            // See if a class name exists
            if (className) {
                // Parse the class names
                let classNames = className.split(' ');
                for (var i = 0; i < classNames.length; i++) {
                    // Add the class name
                    icon.classList.add(classNames[i]);
                }
            } else {
                icon.classList.add("icon-svg");
            }

            // Set the height/width
            height ? icon.setAttribute("height", (height).toString()) : null;
            width ? icon.setAttribute("width", (width).toString()) : null;

            // Hide the icon as non-interactive content from the accessibility API
            icon.setAttribute("aria-hidden", "true");

            // Update the styling
            icon.style.pointerEvents = "none";

            // Support for IE
            icon.setAttribute("focusable", "false");
        }

        // Return the icon
        return icon;
    }

    // Refreshes the dashboard
    private refresh(itemId?: number) {
        // Refresh the data
        DataSource.refresh(itemId).then(() => {
            // Refresh the table
            this._dashboard.refresh(DataSource.ListItems);
        });
    }

    // Renders the dashboard
    private render(el: HTMLElement) {
        // Create the dashboard
        this._dashboard = new Dashboard({
            el,
            hideFooter: !Strings.IsClassic,
            hideHeader: true,
            useModal: true,
            navigation: {
                itemsEnd: this.generateNavItems(),
                showFilter: false,
                title: Strings.ProjectName,
                // Add the branding icon & text
                onRendering: (props) => {
                    // Set the class names
                    props.className = "navbar-expand rounded-top";
                    props.type = Components.NavbarTypes.Primary

                    // Set the brand
                    let brand = document.createElement("div");
                    let text = brand.cloneNode() as HTMLDivElement;
                    brand.className = "d-flex";
                    text.className = "ms-2";
                    text.append(Strings.ProjectName);
                    brand.appendChild(this.getEntraIcon(32, 32, 'brand'));
                    brand.appendChild(text);
                    props.brand = brand;
                },
                // Adjust the brand alignment
                onRendered: (el) => {
                    el.querySelector("nav div.container-fluid a.navbar-brand").classList.add("p-0");
                    el.querySelector("nav div.container-fluid a.navbar-brand").classList.add("pe-none");
                }
            },
            subNavigation: {
                onRendering: props => {
                    props.className = "navbar-sub rounded-bottom";
                },
                onRendered: (el) => {
                    el.querySelector("nav.navbar").classList.remove("bg-light");
                },
                itemsEnd: [
                    {
                        text: "New Application",
                        onRender: (el, item) => {
                            // Clear the existing button
                            el.innerHTML = "";
                            // Create a span to wrap the icon in
                            let span = document.createElement("span");
                            span.className = "bg-white d-inline-flex ms-2 rounded";
                            el.appendChild(span);

                            // Render a tooltip
                            Components.Tooltip({
                                el: span,
                                content: item.text,
                                placement: Components.TooltipPlacements.Left,
                                btnProps: {
                                    // Render the icon button
                                    className: "p-1 pe-2",
                                    iconClassName: "me-1",
                                    iconType: plusSquare,
                                    iconSize: 24,
                                    isSmall: true,
                                    text: "New",
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: () => {
                                        // Show the new form
                                        DataSource.List.newForm({
                                            onSetHeader: (el) => {
                                                el.querySelector("h5") ? el.querySelector("h5").innerHTML = item.text : null;
                                            },
                                            onUpdate: item => {
                                                // Refresh the dashboard
                                                this.refresh(item.Id);
                                            }
                                        });
                                    }
                                },
                            });
                        }
                    }
                ]
            },
            footer: {
                onRendering: props => {
                    // Update the properties
                    props.className = "footer p-0";
                },
                onRendered: (el) => {
                    el.querySelector("nav.footer").classList.remove("bg-light");
                    el.querySelector("nav.footer .container-fluid").classList.add("p-0");
                },
                itemsEnd: [{
                    className: "pe-none text-body",
                    text: "v" + Strings.Version
                }]
            },
            table: {
                rows: DataSource.ListItems,
                dtProps: {
                    dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                    columnDefs: [
                        {
                            "targets": 5,
                            "orderable": false,
                            "searchable": false
                        }
                    ],
                    createdRow: function (row, data, index) {
                        jQuery('td', row).addClass('align-middle');
                    },
                    drawCallback: function (settings) {
                        let api = new jQuery.fn.dataTable.Api(settings) as any;
                        let div = api.table().container() as HTMLDivElement;
                        let table = api.table().node() as HTMLTableElement;
                        div.querySelector(".dataTables_info").classList.add("text-center");
                        div.querySelector(".dataTables_length").classList.add("pt-2");
                        div.querySelector(".dataTables_paginate").classList.add("pt-03");
                        table.classList.remove("no-footer");
                        table.classList.add("tbl-footer");
                        table.classList.add("table-striped");
                    },
                    headerCallback: function (thead, data, start, end, display) {
                        jQuery('th', thead).addClass('align-middle');
                    },
                    // Order by the 3rd column by default; ascending
                    order: [[2, "asc"]]
                },
                columns: [
                    {
                        name: "Title",
                        title: "Application"
                    },
                    {
                        name: "AppId",
                        title: "App ID"
                    },
                    {
                        className: "text-break-spaces",
                        name: "ExpirationDate",
                        title: "Expiration Date",
                        onRenderCell: (el, column, item: IListItem) => {
                            let value = item[column.name];
                            if (value) {
                                // Render the date/time value
                                el.innerHTML = moment(value).format(Strings.TimeFormat);

                                // Set the date/time filter/sort values
                                el.setAttribute("data-filter", moment(value).format("dddd MMMM DD YYYY HH:mm:ss"));
                                el.setAttribute("data-sort", value);
                            }
                        }
                    },
                    {
                        name: "SiteUrls",
                        title: "Site Urls"
                    },
                    {
                        name: "",
                        title: "Owners",
                        onRenderCell: (el, column, item: IListItem) => {
                            // Parse the users
                            let names = [];
                            let users = item.Owners?.results || [];
                            for (let i = 0; i < users.length; i++) {
                                names.push(users[i].Title);
                            }

                            // Render the user names
                            el.innerHTML = names.join(', ');
                        }
                    },
                    {
                        className: "text-end",
                        name: "Actions",
                        isHidden: true,
                        onRenderCell: (el, column, item: IListItem) => {
                            let isOwner = DataSource.isOwner(item);
                            let sitesExist = (item.SiteUrls || "").trim().length > 0;
                            let permLinks: Components.IDropdownItem[] = [];
                            let tooltips: Components.ITooltipProps[] = [
                                {
                                    content: "View Application",
                                    btnProps: {
                                        text: "View",
                                        type: Components.ButtonTypes.OutlineSecondary,
                                        onClick: () => {
                                            // Show the display form
                                            DataSource.List.viewForm({
                                                itemId: item.Id
                                            });
                                        }
                                    }
                                }
                            ];
                            
                            // See if the user is an admin/owner
                            if (Security.IsAdmin || isOwner) {
                                tooltips.push({
                                    content: "Edit Application",
                                    btnProps: {
                                        text: "Edit",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the edit form
                                            DataSource.List.editForm({
                                                itemId: item.Id,
                                                onUpdate: () => {
                                                    // Refresh the dashboard
                                                    this.refresh(item.Id);
                                                }
                                            });
                                        }
                                    }
                                });
                            }

                            // See if the user is an admin/owner and there are no site urls
                            if ((Security.IsAdmin || isOwner) && !sitesExist) {
                                tooltips.push({
                                    content: "Delete Application",
                                    btnProps: {
                                        text: "Delete",
                                        type: Components.ButtonTypes.OutlineDanger,
                                        onClick: () => {
                                            // Show the delete form
                                            Forms.deleteRequest(item, () => {
                                                // Refresh the dashboard
                                                this.refresh();
                                            });
                                        }
                                    }
                                });
                            }

                            // See if the user is a licensed admin
                            if (Security.IsAdmin && DataSource.HasLicense) {
                                // Adds a permission to a new site collection
                                permLinks.push({
                                    text: "Add Permission",
                                    title: "Grant access to a site collection",
                                    onClick: () => {
                                        // Show the add form
                                        Forms.addPermission(item, () => {
                                            // Refresh the dashboard
                                            this.refresh(item.Id);
                                        });
                                    }
                                });

                                // Ensure sites exist
                                if (sitesExist) {
                                    // Views the permission of a site collection
                                    permLinks.push({
                                        text: "View Permission",
                                        title: "View access to a site collection",
                                        onClick: () => {
                                            // Show the view form
                                            Forms.viewPermissions(item);
                                        }
                                    });
                                    
                                    // Edits an existing permission to a site collection
                                    permLinks.push({
                                        text: "Edit Permission",
                                        title: "Edit access to a site collection",
                                        onClick: () => {
                                            // Show the edit form
                                            Forms.editPermission(item);
                                        }
                                    });

                                    // Removes an existing permission from a site collection
                                    permLinks.push({
                                        text: "Remove Permission",
                                        title: "Remove access to a site collection",
                                        onClick: () => {
                                            // Show the remove form
                                            Forms.removePermission(item, () => {
                                                // Refresh the dashboard
                                                this.refresh(item.Id);
                                            });
                                        }
                                    });
                                }
                            }

                            // Render a tooltip group
                            let ttg = Components.TooltipGroup({
                                el,
                                isSmall: true,
                                tooltips
                            });

                            if (permLinks.length > 0) {
                                
                                let ddl = Components.Dropdown({
                                    el: ttg.el,
                                    label: "Permissions",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    items: permLinks,
                                    onMenuRendering: (props) => {
                                        // This is the popover. You can further customize from here

                                        // You must return the props
                                        return props;
                                    },
                                });
                                
                                let btn = ddl.el.querySelector("button");
                                btn ? ttg.el.appendChild(btn) && ddl.el.remove() : null;
                            }
                        }
                    }
                ]
            }
        });
    }
}