import { Helper, SPTypes } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                Title: Strings.Lists.Main,
                BaseTemplate: SPTypes.ListTemplateType.GenericList
            },
            TitleFieldDisplayName: "App Name",
            ContentTypes: [
                {
                    Name: "Item",
                    FieldRefs: [
                        "Title",
                        "ClientId",
                        "Status",
                        "Permission",
                        "SiteUrls",
                        "Owners"
                    ]
                }
            ],
            CustomFields: [
                {
                    name: "ClientId",
                    title: "Client ID",
                    description: "The client id of the application.",
                    type: Helper.SPCfgFieldType.Text,
                    required: true,
                },
                {
                    name: "Owners",
                    title: "Owners",
                    description: "The owners of this request to receive notifications.",
                    type: Helper.SPCfgFieldType.User,
                    multi: true,
                    required: true
                } as Helper.IFieldInfoUser,
                {
                    name: "Permission",
                    title: "Permission",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "Read",
                    required: true,
                    choices: [
                        "Read", "Write", "Manage", "Full Control"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "SiteUrls",
                    title: "Site Urls",
                    description: "Each url must be on a separate line.",
                    type: Helper.SPCfgFieldType.Note,
                    required: true,
                    noteType: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "Status",
                    title: "Status",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "Submitted",
                    required: true,
                    showInEditForm: false,
                    showInNewForm: false,
                    choices: [
                        "Submitted", "Approved", "Completed"
                    ]
                } as Helper.IFieldInfoChoice
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "LinkTitle", "Status", "ClientId", "Permission", "SiteUrls", "Owners"
                    ]
                }
            ]
        }
    ]
});