import { ContextInfo, SPTypes } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, envType?: number, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the properties
    Strings.IsClassic = envType == SPTypes.EnvironmentType.ClassicSharePoint;
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
    FlowId: null,
    GlobalVariable: "EntraSitePermissionManager",
    IsClassic: true,
    Lists: {
        Main: "Site Manager"
    },
    ProjectName: "Entra Site Permission Manager",
    ProjectDescription: "A dashboard to manage site permission requests for apps using Graph API.",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    TimeFormat: "YYYY-MMM-DD[\n]HH:mm:ss",
    Version: "0.0.1"
};
export default Strings;