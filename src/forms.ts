import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, Site } from "gd-sprest-bs";
import { DataSource, IListItem } from "./ds";
import Strings from "./strings";

/**
 * Forms
 */
export class Forms {
    // Adds a permission to a site
    static addPermission(item: IListItem, onUpdate: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Add Site Permission");

        // Set the body
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "siteUrl",
                    label: "Site Url:",
                    type: Components.FormControlTypes.TextField,
                    required: true,
                    errorMessage: "A site url is required"
                } as Components.IFormControlPropsDropdown,
                {
                    name: "permission",
                    label: "Permission:",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    items: [
                        { text: "Read", value: "read" },
                        { text: "Write", value: "write" },
                        { text: "Full Control", value: "fullcontrol" }
                    ]
                } as Components.IFormControlPropsDropdown]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            isSmall: true,
            tooltips: [
                {
                    content: "Add the permission to the site",
                    btnProps: {
                        text: "Add",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let ctrlSiteUrl = form.getControl("siteUrl");
                                let values = form.getValues();

                                // Show a loading dialog
                                LoadingDialog.setHeader("Getting Site Information");
                                LoadingDialog.setBody("Validating the site url...");
                                LoadingDialog.show();

                                // Get the site information
                                Site(values["siteUrl"]).execute(
                                    site => {
                                        // Update the loading dialog
                                        LoadingDialog.setHeader("Calling the Flow");
                                        LoadingDialog.setBody("This will close after the flow is triggered...");

                                        // Run the flow
                                        DataSource.runFlow({
                                            appId: item.AppId,
                                            appName: item.Title,
                                            itemId: item.Id,
                                            ownerEmails: DataSource.getOwnerEmails(item).join(', '),
                                            permission: values["permission"].value,
                                            permissionId: "",
                                            requestType: "add",
                                            siteId: site.Id,
                                            siteUrl: site.ServerRelativeUrl
                                        }).then(
                                            // Success
                                            () => {
                                                // Add the site url to the item
                                                let siteUrls: string[] = (item.SiteUrls || "").split('\n');
                                                if (siteUrls.indexOf(site.ServerRelativeUrl) < 0) {
                                                    // Update the loading dialog
                                                    LoadingDialog.setHeader("Updating List Item");
                                                    LoadingDialog.setBody("This will close after the site url is added to this item...");

                                                    // Append the url
                                                    siteUrls.push(site.ServerRelativeUrl);

                                                    // Update the value
                                                    item.update({
                                                        SiteUrls: siteUrls.join('\n').replace(/^\n|\n$/g, "")
                                                    }).execute(() => {
                                                        // Error getting the site
                                                        ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                                            isValid: true,
                                                            validMessage: "Flow was run and the item was updated successfully"
                                                        });

                                                        // Call the event
                                                        onUpdate();

                                                        // Hide the loading dialog
                                                        LoadingDialog.hide();
                                                    });
                                                } else {
                                                    // Hide the loading dialog
                                                    LoadingDialog.hide();
                                                }
                                            },

                                            // Error
                                            errMessage => {
                                                // Error getting the site
                                                ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                                    isValid: false,
                                                    invalidMessage: errMessage
                                                });

                                                // Hide the loading dialog
                                                LoadingDialog.hide();
                                            }
                                        )
                                    },

                                    () => {
                                        // Error getting the site
                                        ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                            isValid: false,
                                            invalidMessage: "Error getting the site information"
                                        });

                                        // Hide the loading dialog
                                        LoadingDialog.hide();
                                    }
                                );
                            }
                        }
                    }
                },
                {
                    content: "Close the dialog",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => { Modal.hide(); }
                    }
                }
            ]
        });

        // Show the form
        Modal.show();
    }

    // Deletes the application
    static deleteRequest(item: IListItem, onUpdate: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Remove Application");

        // Set the body
        Modal.setBody("Are you sure you want to remove this application?");

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            isSmall: true,
            tooltips: [
                {
                    content: "Remove the Application",
                    btnProps: {
                        text: "Remove",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Show a loading dialog
                            LoadingDialog.setHeader("Removing the Application");
                            LoadingDialog.setBody("This will close after the application is removed...");
                            LoadingDialog.show();

                            // Delete the item
                            item.delete().execute(
                                // Success
                                () => {
                                    // Call the update event
                                    onUpdate();

                                    // Close the dialog
                                    Modal.hide();

                                    // Hide the dialog
                                    LoadingDialog.hide();
                                },

                                // Error
                                () => {
                                    // Set the body
                                    Modal.setBody("Error removing the application. Refresh the page and try again...");

                                    // Hide the dialog
                                    LoadingDialog.hide();
                                }
                            )
                        }
                    }
                },
                {
                    content: "Close the dialog",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Hides the modal
                            Modal.hide();
                        }
                    }
                }
            ]
        });

        // Show the form
        Modal.show();
    }

    // Edits a permission to a site
    static editPermission(item: IListItem) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Edit Site Permission");

        // Parse the site urls
        let items: Components.IDropdownItem[] = [];
        let siteUrls = (item.SiteUrls || "").split('\n');
        for (let i = 0; i < siteUrls.length; i++) {
            let siteUrl = siteUrls[i];
            if (siteUrl) {
                // Add the item
                items.push({ text: siteUrl, value: siteUrl });
            }
        }

        // Set the body
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "siteUrl",
                    label: "Site Url:",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    items,
                    errorMessage: "A site url is required"
                } as Components.IFormControlPropsDropdown,
                {
                    name: "permission",
                    label: "Permission:",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    items: [
                        { text: "Read", value: "read" },
                        { text: "Write", value: "write" },
                        { text: "Full Control", value: "fullcontrol" }
                    ]
                } as Components.IFormControlPropsDropdown]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            isSmall: true,
            tooltips: [
                {
                    content: "Add the permission to the site",
                    btnProps: {
                        text: "Update",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let ctrlSiteUrl = form.getControl("siteUrl");
                                let values = form.getValues();

                                // Show a loading dialog
                                LoadingDialog.setHeader("Calling the Flow");
                                LoadingDialog.setBody("This will close after the flow is triggered...");
                                LoadingDialog.show();

                                // Get the permission id for this site
                                let siteUrl = values["siteUrl"].value;
                                DataSource.getSitePermissions(siteUrl).then(sitePermissions => {
                                    let permissionId = null;

                                    // Find the permission
                                    for (let i = 0; i < sitePermissions.permissions.results.length; i++) {
                                        let permission = sitePermissions.permissions.results[i];

                                        // Parse the identities
                                        let identities = permission.grantedToIdentities || [];
                                        for (let j = 0; j < identities.length; j++) {
                                            if (identities[j].application.id == item.AppId) {
                                                // Set the permission id
                                                permissionId = permission.id;
                                                break;
                                            }
                                        }

                                        // See if the permission id was found
                                        if (permissionId) { break; }
                                    }

                                    // See if the permission was found
                                    if (permissionId) {
                                        // Run the flow
                                        DataSource.runFlow({
                                            appId: item.AppId,
                                            appName: item.Title,
                                            itemId: item.Id,
                                            ownerEmails: DataSource.getOwnerEmails(item).join(', '),
                                            permission: values["permission"].value,
                                            permissionId,
                                            requestType: "update",
                                            siteUrl
                                        }).then(
                                            // Success
                                            () => {
                                                // Successfully triggered the flow
                                                ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                                    isValid: true,
                                                    validMessage: "Flow was run and the item was updated successfully"
                                                });

                                                // Hide the loading dialog
                                                LoadingDialog.hide();
                                            },

                                            // Error
                                            errMessage => {
                                                // Set the error
                                                ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                                    isValid: false,
                                                    invalidMessage: errMessage
                                                });

                                                // Hide the loading dialog
                                                LoadingDialog.hide();
                                            }
                                        )
                                    } else {
                                        // Hide the loading dialog
                                        LoadingDialog.hide();

                                        // Set the error
                                        let ctrl = form.getControl("siteUrl");
                                        ctrl.updateValidation(ctrl.el, {
                                            isValid: false,
                                            invalidMessage: "Error getting the permission for this App Id"
                                        });
                                    }
                                });
                            }
                        }
                    }
                },
                {
                    content: "Close the dialog",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => { Modal.hide(); }
                    }
                }
            ]
        });

        // Show the form
        Modal.show();
    }

    // Removes a permission from a site
    static removePermission(item: IListItem, onUpdate: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Remove Site Permission");

        // Parse the site urls
        let items: Components.IDropdownItem[] = [];
        let siteUrls = (item.SiteUrls || "").split('\n');
        for (let i = 0; i < siteUrls.length; i++) {
            let siteUrl = siteUrls[i];
            if (siteUrl) {
                // Add the item
                items.push({ text: siteUrl, value: siteUrl });
            }
        }

        // Set the body
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "siteUrl",
                    label: "Site Url:",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    items,
                    errorMessage: "A site url is required"
                } as Components.IFormControlPropsDropdown
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            isSmall: true,
            tooltips: [
                {
                    content: "Add the permission to the site",
                    btnProps: {
                        text: "Remove",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let ctrlSiteUrl = form.getControl("siteUrl");
                                let siteUrl = form.getValues()["siteUrl"].value;

                                // Show a loading dialog
                                LoadingDialog.setHeader("Calling the Flow");
                                LoadingDialog.setBody("This will close after the flow is triggered...");
                                LoadingDialog.show();

                                // Get the permission id for this site
                                DataSource.getSitePermissions(siteUrl).then(sitePermissions => {
                                    let permissionId = null;

                                    // Find the permission
                                    for (let i = 0; i < sitePermissions.permissions.results.length; i++) {
                                        let permission = sitePermissions.permissions.results[i];

                                        // Parse the identities
                                        let identities = permission.grantedToIdentities || [];
                                        for (let j = 0; j < identities.length; j++) {
                                            if (identities[j].application.id == item.AppId) {
                                                // Set the permission id
                                                permissionId = permission.id;
                                                break;
                                            }
                                        }

                                        // See if the permission id was found
                                        if (permissionId) { break; }
                                    }

                                    // See if the permission was found
                                    if (permissionId) {
                                        // Run the flow
                                        DataSource.runFlow({
                                            appId: item.AppId,
                                            appName: item.Title,
                                            itemId: item.Id,
                                            ownerEmails: DataSource.getOwnerEmails(item).join(', '),
                                            permission: "",
                                            permissionId,
                                            requestType: "remove",
                                            siteUrl
                                        }).then(
                                            // Success
                                            () => {
                                                // Successfully triggered the flow
                                                ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                                    isValid: true,
                                                    validMessage: "Flow was run and the item was updated successfully"
                                                });

                                                // Add the site url to the item
                                                let siteUrls: string[] = (item.SiteUrls || "").split('\n');
                                                let siteIdx = siteUrls.indexOf(siteUrl);
                                                if (siteIdx >= 0) {
                                                    // Update the loading dialog
                                                    LoadingDialog.setHeader("Updating List Item");
                                                    LoadingDialog.setBody("This will close after the site url is removed from this item...");

                                                    // Remove the url
                                                    siteUrls.splice(siteIdx, 1);

                                                    // Update the value
                                                    item.update({
                                                        SiteUrls: siteUrls.join('\n').replace(/^\n|\n$/g, "")
                                                    }).execute(() => {
                                                        // Call the event
                                                        onUpdate();

                                                        // Hide the loading dialog
                                                        LoadingDialog.hide();
                                                    });
                                                } else {
                                                    // Hide the loading dialog
                                                    LoadingDialog.hide();
                                                }
                                            },

                                            // Error
                                            errMessage => {
                                                // Set the error
                                                ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                                    isValid: false,
                                                    invalidMessage: errMessage
                                                });

                                                // Hide the loading dialog
                                                LoadingDialog.hide();
                                            }
                                        );
                                    } else {
                                        // Set the error
                                        ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                            isValid: false,
                                            invalidMessage: "Error getting the permission for this App Id"
                                        });

                                        // Hide the loading dialog
                                        LoadingDialog.hide();
                                    }
                                });
                            }
                        }
                    }
                },
                {
                    content: "Close the dialog",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => { Modal.hide(); }
                    }
                }
            ]
        });

        // Show the form
        Modal.show();
    }

    // Views the permissions for a site
    static viewPermissions(item: IListItem) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("View Site Permission");

        // Parse the site urls
        let items: Components.IDropdownItem[] = [];
        let siteUrls = (item.SiteUrls || "").split('\n');
        for (let i = 0; i < siteUrls.length; i++) {
            let siteUrl = siteUrls[i];
            if (siteUrl) {
                // Add the item
                items.push({ text: siteUrl, value: siteUrl });
            }
        }

        // Set the body
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "siteUrl",
                    label: "Site Url:",
                    type: Components.FormControlTypes.Dropdown,
                    items
                } as Components.IFormControlPropsDropdown
            ]
        });

        // Permissions element
        let elPermissions = document.createElement("div");
        elPermissions.classList.add("mt-2");
        Modal.BodyElement.appendChild(elPermissions);

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            isSmall: true,
            tooltips: [
                {
                    content: "View the permission to the site",
                    btnProps: {
                        text: "View",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Get the site url
                            let siteUrl = form.getValues()["siteUrl"]?.value;
                            if (siteUrl) {
                                // Get the permissions
                                DataSource.getSitePermissions(siteUrl).then(sitePermissions => {
                                    let permissionInfo: {
                                        appName: string;
                                        AppId: string;
                                        permissionId: string;
                                        siteId: string;
                                    }[] = [];

                                    // Find the permission
                                    for (let i = 0; i < sitePermissions.permissions.results.length; i++) {
                                        let permission = sitePermissions.permissions.results[i];

                                        // Parse the identities
                                        let identities = permission.grantedToIdentities || [];
                                        for (let j = 0; j < identities.length; j++) {
                                            let identity = identities[j];

                                            // Add the permission information
                                            permissionInfo.push({
                                                appName: identity.application.displayName,
                                                AppId: identity.application.id,
                                                permissionId: permission.id,
                                                siteId: sitePermissions.siteId
                                            });
                                        }
                                    }

                                    // Render a table
                                    new DataTable({
                                        el: elPermissions,
                                        rows: permissionInfo,
                                        dtProps: {
                                            dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                                        },
                                        columns: [
                                            {
                                                name: "appName",
                                                title: "Application"
                                            },
                                            {
                                                name: "AppId",
                                                title: "App ID"
                                            },
                                            {
                                                name: "permission",
                                                title: "Permission",
                                                onRenderCell: (el, col, item: {
                                                    appName: string;
                                                    AppId: string;
                                                    permissionId: string;
                                                    siteId: string;
                                                }) => {
                                                    // Get the permission
                                                    DataSource.getSitePermission(item.siteId, item.permissionId).then(permission => {
                                                        // Render the roles
                                                        el.innerHTML = permission.roles.join(', ');
                                                    });
                                                }
                                            }
                                        ]
                                    });
                                });

                                // Clear the element
                                while (elPermissions.firstChild) { elPermissions.removeChild(elPermissions.firstChild); }
                            }
                        }
                    }
                },
                {
                    content: "Close the dialog",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => { Modal.hide(); }
                    }
                }
            ]
        });

        // Show the form
        Modal.show();
    }
}