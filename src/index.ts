import { LoadingDialog } from "dattatable";
import { ContextInfo, ThemeManager } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
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
    flowId?: string;
    url?: string;
}

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    render: (props: IProps) => {
        // See if the page context exists
        if (props.context) {
            // Set the context
            setContext(props.context, props.url);

            // Update the configuration
            Configuration.setWebUrl(props.url || ContextInfo.webServerRelativeUrl);
        }

        // Set the flow id
        props.flowId ? setFlowId(props.flowId) : null;

        // Initialize the application
        DataSource.init().then(
            // Success
            () => {
                // Show a loading dialog
                LoadingDialog.setHeader("Loading the Theme");
                LoadingDialog.setBody("This will close after the theme is loaded...");
                LoadingDialog.show();

                // Load the current theme and apply it to the components
                ThemeManager.load(true).then(() => {
                    // Create the application
                    new App(props.el);

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
    }
};

// Make is available in the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Render the application
    GlobalVariable.render({
        el: elApp,
        flowId: elApp.getAttribute("data-flowId")
    });
}