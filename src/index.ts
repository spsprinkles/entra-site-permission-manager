import { LoadingDialog, waitForTheme } from "dattatable";
import { ContextInfo, ThemeManager } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
import { getEntraIcon } from "./common"
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import { Security } from "./security";
import Strings, { setContext, setFlowId } from "./strings";

// Styling
import "./styles.scss";

// Properties
interface IProps {
    el: HTMLElement,
    context?: any;
    displayMode?: number;
    envType?: number;
    flowId?: string;
    sourceUrl?: string;
    timeFormat?: string;
    title?: string;
}

// Create the global variable for this solution
const GlobalVariable = {
    App: null,
    Configuration,
    description: Strings.ProjectDescription,
    getLogo: () => { return getEntraIcon(28, 28, 'brand logo me-2'); },
    render: (props: IProps) => {
        // See if the page context exists
        if (props.context) {
            // Set the context
            setContext(props.context, props.envType, props.sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(props.sourceUrl || ContextInfo.webServerRelativeUrl);
        }

        // Set the flow id
        props.flowId ? setFlowId(props.flowId) : null;

        // Update the TimeFormat from SPFx value
        props.timeFormat ? Strings.TimeFormat = props.timeFormat : null;

        // Update the ProjectName from SPFx title field
        props.title ? Strings.ProjectName = props.title : null;

        // Initialize the application
        DataSource.init().then(
            // Success
            () => {
                // Show a loading dialog
                LoadingDialog.setHeader("Loading the Theme");
                LoadingDialog.setBody("This will close after the theme is loaded...");
                LoadingDialog.show();

                // Load the current theme and apply it to the components
                waitForTheme().then(() => {
                    // Create the application
                    GlobalVariable.App = new App(props.el);

                    // Hide the loading dialog
                    LoadingDialog.hide();
                });
            },

            // Error
            () => {
                // See if the user has the correct permissions
                Security.hasPermissions().then(hasPermissions => {
                    // See if the user has permissions
                    if (hasPermissions) {
                        // Show the installation modal
                        InstallationModal.show();
                    }
                });
            }
        );
    },
    timeFormat: Strings.TimeFormat,
    title: Strings.ProjectName,
    updateTheme: (themeInfo) => {
        // Set the theme
        ThemeManager.setCurrentTheme(themeInfo);
    }
};

// Make is available in the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Remove the extra border spacing on the webpart in classic mode
    let contentBox = document.querySelector("#contentBox table.ms-core-tableNoSpace");
    contentBox ? contentBox.classList.remove("ms-webpartPage-root") : null;

    // Render the application
    GlobalVariable.render({
        el: elApp,
        flowId: elApp.getAttribute("data-flowId")
    });
}