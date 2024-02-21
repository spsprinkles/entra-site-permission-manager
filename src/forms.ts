import { LoadingDialog, Modal } from "dattatable";
import { Components, Site } from "gd-sprest-bs";
import { DataSource, IListItem } from "./ds";
import Strings from "./strings";

/**
 * Forms
 */
export class Forms {
    // Adds a permission to a site
    static addPermission(item: IListItem) {
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
                    errorMessage: "A site url is required."
                } as Components.IFormControlPropsDropdown,
                {
                    name: "permission",
                    label: "Permission:",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    items: [
                        { text: "Read", value: "read" },
                        { text: "Write", value: "write" },
                        { text: "Owner", value: "owner" }
                    ]
                } as Components.IFormControlPropsDropdown]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            isSmall: true,
            tooltips: [
                {
                    content: "Adds the permission to the site.",
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
                                            appId: item.ClientId,
                                            appName: item.Title,
                                            id: item.Id,
                                            permission: values["permission"].value,
                                            type: "add",
                                            url: site.ServerRelativeUrl
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
                                                        SiteUrls: siteUrls.join('\n')
                                                    }).execute(() => {
                                                        // Error getting the site
                                                        ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                                            isValid: true,
                                                            validMessage: "Flow was run and the item was updated successfully."
                                                        });

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
                                            invalidMessage: "Error getting the site information."
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
                    content: "Closes the dialog",
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
                    errorMessage: "A site url is required."
                } as Components.IFormControlPropsDropdown,
                {
                    name: "permission",
                    label: "Permission:",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    items: [
                        { text: "Read", value: "read" },
                        { text: "Write", value: "write" },
                        { text: "Owner", value: "owner" }
                    ]
                } as Components.IFormControlPropsDropdown]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            isSmall: true,
            tooltips: [
                {
                    content: "Adds the permission to the site.",
                    btnProps: {
                        text: "Update",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let values = form.getValues();

                                // Show a loading dialog
                                LoadingDialog.setHeader("Calling the Flow");
                                LoadingDialog.setBody("This will close after the flow is triggered...");
                                LoadingDialog.show();

                                // Get the permission id for this site
                                DataSource.getSitePermissions(values["siteUrl"].value).then(permissions => {
                                    let permissionId = null;

                                    // Find the permission
                                    for (let i = 0; i < permissions.results.length; i++) {
                                        let permission = permissions.results[i];

                                        // Parse the identities
                                        let identities = permission.grantedToIdentities || [];
                                        for (let j = 0; j < identities.length; j++) {
                                            if (identities[j].application.id == item.ClientId) {
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
                                            appId: item.ClientId,
                                            appName: item.Title,
                                            id: item.Id,
                                            permission: values["permission"].value,
                                            permissionId,
                                            type: "update",
                                            url: Strings.SourceUrl
                                        }).then(
                                            // Success
                                            () => {
                                                // Hide the loading dialog
                                                LoadingDialog.hide();
                                            },

                                            // Error
                                            err => {
                                                // Show the error
                                                // TODO
                                            }
                                        )
                                    } else {
                                        // Hide the loading dialog
                                        LoadingDialog.hide();

                                        // Set the error
                                        let ctrl = form.getControl("siteUrl");
                                        ctrl.updateValidation(ctrl.el, {
                                            isValid: false,
                                            invalidMessage: "Error getting the permission for this client id."
                                        });
                                    }
                                });
                            }
                        }
                    }
                },
                {
                    content: "Closes the dialog",
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
    static removePermission(item: IListItem) {
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
                    errorMessage: "A site url is required."
                } as Components.IFormControlPropsDropdown
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            isSmall: true,
            tooltips: [
                {
                    content: "Adds the permission to the site.",
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
                                DataSource.getSitePermissions(siteUrl).then(permissions => {
                                    let permissionId = null;

                                    // Find the permission
                                    for (let i = 0; i < permissions.results.length; i++) {
                                        let permission = permissions.results[i];

                                        // Parse the identities
                                        let identities = permission.grantedToIdentities || [];
                                        for (let j = 0; j < identities.length; j++) {
                                            if (identities[j].application.id == item.ClientId) {
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
                                            appId: item.ClientId,
                                            appName: item.Title,
                                            id: item.Id,
                                            permissionId,
                                            type: "remove",
                                            url: siteUrl
                                        }).then(
                                            // Success
                                            () => {
                                                // Add the site url to the item
                                                let siteUrls: string[] = (item.SiteUrls || "").split('\n');
                                                let siteIdx = siteUrls.indexOf(siteUrl);
                                                if (siteIdx >= 0) {
                                                    // Update the loading dialog
                                                    LoadingDialog.setHeader("Updating List Item");
                                                    LoadingDialog.setBody("This will close after the site url is added to this item...");

                                                    // Remove the url
                                                    siteUrls.slice(siteIdx, 1);

                                                    // Update the value
                                                    item.update({
                                                        SiteUrls: siteUrls.join('\n')
                                                    }).execute(() => {
                                                        // Error getting the site
                                                        ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                                            isValid: true,
                                                            validMessage: "Flow was run and the item was updated successfully."
                                                        });

                                                        // Hide the loading dialog
                                                        LoadingDialog.hide();
                                                    });
                                                } else {
                                                    // Hide the loading dialog
                                                    LoadingDialog.hide();
                                                }
                                            }

                                            // Error
                                        );
                                    } else {
                                        // Hide the loading dialog
                                        LoadingDialog.hide();

                                        // Set the error
                                        let ctrl = form.getControl("siteUrl");
                                        ctrl.updateValidation(ctrl.el, {
                                            isValid: false,
                                            invalidMessage: "Error getting the permission for this client id."
                                        });
                                    }
                                });
                            }
                        }
                    }
                },
                {
                    content: "Closes the dialog",
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