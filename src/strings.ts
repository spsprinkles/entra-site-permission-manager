import { ContextInfo, SPTypes } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the source url
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "entra-site-manager",
    CloudEnv: SPTypes.CloudEnvironment.Default,
    Flows: {
        Add: "be1edf84-e914-4896-a236-50080778e2bd",
        Remove: "",
        Update: ""
    },
    GlobalVariable: "EntraSiteManager",
    Lists: {
        Main: "Site Manager"
    },
    ProjectName: "Entra ID Site Manager",
    ProjectDescription: "Dashboard to manage site permission requests for apps using the graph api.",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.1"
};
export default Strings;