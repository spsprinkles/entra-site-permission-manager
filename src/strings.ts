import { ContextInfo, SPTypes } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the source url
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
}

export const setFlowId = (flowId: string) => {
    // Set the flow id
    Strings.FlowId = flowId;
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "entra-site-permission-manager",
    CloudEnv: SPTypes.CloudEnvironment.Default,
    FlowId: "",
    GlobalVariable: "EntraSitePermissionManager",
    Lists: {
        Main: "Site Manager"
    },
    ProjectName: "Entra Site Permission Manager",
    ProjectDescription: "Dashboard to manage site permission requests for apps using the graph api.",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.1"
};
export default Strings;