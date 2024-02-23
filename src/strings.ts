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
    FlowId: "80dc88c1-e064-4cb1-b125-78ee248b1381",
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