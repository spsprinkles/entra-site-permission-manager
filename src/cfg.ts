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
                        "AppId",
                        "SiteUrls",
                        "Owners"
                    ]
                }
            ],
            CustomFields: [
                {
                    name: "AppId",
                    title: "App ID",
                    description: "The app/client id of the application.",
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
                    name: "SiteUrls",
                    title: "Site Urls",
                    description: "Each url must be on a separate line.",
                    type: Helper.SPCfgFieldType.Note,
                    required: true,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    showInEditForm: false,
                    showInNewForm: false
                } as Helper.IFieldInfoNote
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "LinkTitle", "AppId", "Owners", "SiteUrls"
                    ]
                }
            ]
        }
    ]
});