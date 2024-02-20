import { LoadingDialog, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
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
            controls: [{
                name: "permission",
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
                                // Show a loading dialog
                                LoadingDialog.setHeader("Calling the Flow");
                                LoadingDialog.setBody("This will close after the flow is triggered...");
                                LoadingDialog.show();

                                // Run the flow
                                DataSource.runFlow({
                                    appName: item.Title,
                                    id: item.Id,
                                    permission: form.getValues()["permission"].value,
                                    type: "add",
                                    url: Strings.SourceUrl
                                }).then(
                                    // Success
                                    () => {
                                        // Hide the loading dialog
                                        LoadingDialog.hide();

                                        // Add the site url to the item
                                        // TODO
                                    }

                                    // Error
                                )
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

        // Set the body
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [{
                name: "permission",
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
                                // Show a loading dialog
                                LoadingDialog.setHeader("Calling the Flow");
                                LoadingDialog.setBody("This will close after the flow is triggered...");
                                LoadingDialog.show();

                                // Run the flow
                                DataSource.runFlow({
                                    appName: item.Title,
                                    id: item.Id,
                                    permission: form.getValues()["permission"].value,
                                    type: "update",
                                    url: Strings.SourceUrl
                                }).then(
                                    // Success
                                    () => {
                                        // Hide the loading dialog
                                        LoadingDialog.hide();

                                        // Add the site url to the item
                                        // TODO
                                    }

                                    // Error
                                )
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

        // Set the body
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [{
                name: "permission",
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
                                // Show a loading dialog
                                LoadingDialog.setHeader("Calling the Flow");
                                LoadingDialog.setBody("This will close after the flow is triggered...");
                                LoadingDialog.show();

                                // Run the flow
                                DataSource.runFlow({
                                    appName: item.Title,
                                    id: item.Id,
                                    permission: form.getValues()["permission"].value,
                                    type: "remove",
                                    url: Strings.SourceUrl
                                }).then(
                                    // Success
                                    () => {
                                        // Hide the loading dialog
                                        LoadingDialog.hide();

                                        // Add the site url to the item
                                        // TODO
                                    }

                                    // Error
                                )
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