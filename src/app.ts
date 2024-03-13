import { Dashboard } from "dattatable";
import { Components } from "gd-sprest-bs";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { plusSquare } from "gd-sprest-bs/build/icons/svgs/plusSquare";
import * as jQuery from "jquery";
import * as moment from "moment";
import { getEntraIcon } from "./common"
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
                    brand.appendChild(getEntraIcon(32, 32, 'brand'));
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
                                    content: "Remove Application",
                                    btnProps: {
                                        text: "Remove",
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
                                tooltips.push({
                                    content: "Set Permissions",
                                    ddlProps: {
                                        label: "Permissions",
                                        type: Components.DropdownTypes.OutlinePrimary,
                                        items: [
                                            {
                                                // Adds a permission to a new site collection
                                                text: "Add Permission",
                                                onClick: () => {
                                                    // Show the add form
                                                    Forms.addPermission(item, () => {
                                                        // Refresh the dashboard
                                                        this.refresh(item.Id);
                                                    });
                                                },
                                                onRender: (el) => {
                                                    Components.Tooltip({
                                                        content: "Grant access to a site collection",
                                                        placement: Components.TooltipPlacements.Left,
                                                        target: el,
                                                        type: Components.TooltipTypes.Primary
                                                    });
                                                }
                                            }
                                        ],
                                        onMenuRendering: (props) => {
                                            props.options = {
                                                theme: "light-border",
                                                trigger: "click"
                                            }
                                            return props;
                                        }
                                    }
                                });

                                // Ensure sites exist
                                if (sitesExist) {
                                    // Insert View Permission first in the dropdown items of the permissions tooltip
                                    tooltips[tooltips.length - 1].ddlProps?.items.unshift(
                                        {
                                            // Views the permission of a site collection
                                            text: "View Permission",
                                            onClick: () => {
                                                // Show the view form
                                                Forms.viewPermissions(item);
                                            },
                                            onRender: (el) => {
                                                Components.Tooltip({
                                                    content: "View access to a site collection",
                                                    placement: Components.TooltipPlacements.Left,
                                                    target: el,
                                                    type: Components.TooltipTypes.Primary
                                                });
                                            }
                                        }
                                    );

                                    // Add more dropdown items to the permissions tooltip
                                    tooltips[tooltips.length - 1].ddlProps?.items.push(
                                        {
                                            // Edits an existing permission to a site collection
                                            text: "Edit Permission",
                                            onClick: () => {
                                                // Show the edit form
                                                Forms.editPermission(item);
                                            },
                                            onRender: (el) => {
                                                Components.Tooltip({
                                                    content: "Edit access to a site collection",
                                                    placement: Components.TooltipPlacements.Left,
                                                    target: el,
                                                    type: Components.TooltipTypes.Primary
                                                });
                                            }
                                        },
                                        {
                                            // Removes an existing permission from a site collection
                                            text: "Remove Permission",
                                            onClick: () => {
                                                // Show the remove form
                                                Forms.removePermission(item, () => {
                                                    // Refresh the dashboard
                                                    this.refresh(item.Id);
                                                });
                                            },
                                            onRender: (el) => {
                                                Components.Tooltip({
                                                    content: "Remove access to a site collection",
                                                    placement: Components.TooltipPlacements.Left,
                                                    target: el,
                                                    type: Components.TooltipTypes.Primary
                                                });
                                            }
                                        }
                                    );
                                }
                            }

                            // Render a tooltip group
                            Components.TooltipGroup({
                                el,
                                className: "bg-white",
                                isSmall: true,
                                tooltips
                            });
                        }
                    }
                ]
            }
        });
    }
}