import { Dashboard, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, Types } from "gd-sprest-bs";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import * as jQuery from "jquery";
import { DataSource, IListItem } from "./ds";
import { Forms } from "./forms";
import { Security } from "./security";
import Strings from "./strings";

/**
 * Main Application
 */
export class App {
    // Constructor
    constructor(el: HTMLElement) {
        // Render the dashboard
        this.render(el);
    }

    // Renders the dashboard
    private render(el: HTMLElement) {
        // Create the dashboard
        let dashboard = new Dashboard({
            el,
            hideHeader: true,
            useModal: true,
            filters: {
                items: [{
                    header: "By Status",
                    items: DataSource.StatusFilters,
                    onFilter: (value: string) => {
                        // Filter the table
                        dashboard.filter(1, value);
                    }
                }]
            },
            navigation: {
                title: Strings.ProjectName,
            },
            subNavigation: {
                onRendering: props => {
                    props.className = "navbar-sub rounded-bottom";
                },
                items: [
                    {
                        className: "btn-outline-light",
                        text: "Create Item",
                        isButton: true,
                        onClick: () => {
                            // Show the new form
                            DataSource.List.newForm({
                                onUpdate: (item: IListItem) => {
                                    // Refresh the data
                                    DataSource.refresh(item.Id).then(() => {
                                        // Refresh the table
                                        dashboard.refresh(DataSource.ListItems);
                                    });
                                }
                            });
                        }
                    }
                ],
                itemsEnd: [
                    {
                        text: "Filter Services",
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
                                placement: Components.TooltipPlacements.Right,
                                btnProps: {
                                    // Render the icon button
                                    className: "p-1 pe-2",
                                    iconClassName: "me-1",
                                    iconType: filterSquare,
                                    iconSize: 24,
                                    isSmall: true,
                                    text: "Filters",
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: () => {
                                        // Show the filter
                                        dashboard.showFilter();
                                    }
                                },
                            });
                        }
                    }
                ]
            },
            footer: {
                itemsEnd: [
                    {
                        text: "v" + Strings.Version
                    }
                ]
            },
            table: {
                rows: DataSource.ListItems,
                dtProps: {
                    dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                    columnDefs: [
                        {
                            "targets": 0,
                            "orderable": false,
                            "searchable": false
                        }
                    ],
                    createdRow: function (row, data, index) {
                        jQuery('td', row).addClass('align-middle');
                    },
                    drawCallback: function (settings) {
                        let api = new jQuery.fn.dataTable.Api(settings) as any;
                        jQuery(api.context[0].nTable).removeClass('no-footer');
                        jQuery(api.context[0].nTable).addClass('tbl-footer');
                        jQuery(api.context[0].nTable).addClass('table-striped');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_info').addClass('text-center');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_length').addClass('pt-2');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_paginate').addClass('pt-03');
                    },
                    headerCallback: function (thead, data, start, end, display) {
                        jQuery('th', thead).addClass('align-middle');
                    },
                    // Order by the 1st column by default; ascending
                    order: [[1, "asc"]]
                },
                columns: [
                    {
                        name: "",
                        title: "Application",
                        onRenderCell: (el, column, item: IListItem) => {
                            // Render the application information
                            el.innerHTML = `
                                <b class="me-2">Name:</b>${item.Title}
                                <br/>
                                <b class="me-2">Client Id:</b>${item.ClientId}
                            `;
                        }
                    },
                    {
                        name: "Status",
                        title: "Status"
                    },
                    {
                        name: "SiteUrls",
                        title: "Site Urls"
                    },
                    {
                        name: "",
                        title: "Owners",
                        onRenderCell: (el, column, item: IListItem) => {
                            let users = item.Owners?.results || [];

                            // Render the user information
                            el.innerHTML = users.join(', ');
                        }
                    },
                    {
                        name: "",
                        title: "Actions",
                        onRenderCell: (el, column, item: IListItem) => {
                            let tooltips = [
                                {
                                    content: "Views the request.",
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
                            ]

                            // See if the user is a licensed admin
                            if (Security.IsAdmin && DataSource.HasLicense) {
                                tooltips.push({
                                    content: "Edits the request.",
                                    btnProps: {
                                        text: "Edit",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the edit form
                                            DataSource.List.editForm({
                                                itemId: item.Id,
                                                onUpdate: () => {
                                                    // Refresh the data
                                                    DataSource.refresh(item.Id).then(() => {
                                                        // Refresh the table
                                                        dashboard.refresh(DataSource.ListItems);
                                                    });
                                                }
                                            });
                                        }
                                    }
                                });

                                tooltips.push({
                                    content: "Grants access to a site collection.",
                                    btnProps: {
                                        text: "Add Permission",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the add form
                                            Forms.addPermission(item);
                                        }
                                    }
                                });

                                tooltips.push({
                                    content: "Edits access to a site collection.",
                                    btnProps: {
                                        text: "Edit Permission",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the add form
                                            Forms.editPermission(item);
                                        }
                                    }
                                });

                                tooltips.push({
                                    content: "Removes access to a site collection.",
                                    btnProps: {
                                        text: "Remove Permission",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the add form
                                            Forms.removePermission(item);
                                        }
                                    }
                                });
                            }

                            // Render a buttons
                            Components.TooltipGroup({
                                el,
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