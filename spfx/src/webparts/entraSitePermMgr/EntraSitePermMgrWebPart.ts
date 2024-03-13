import { DisplayMode, Environment, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneHorizontalRule, PropertyPaneLabel, PropertyPaneLink, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'EntraSitePermMgrWebPartStrings';

export interface IEntraSitePermMgrWebPartProps {
  flowId: string;
  timeFormat: string;
  title: string;
  webUrl: string;
}

// Reference the solution
import "../../../../dist/entra-site-permission-manager.min.js";
declare const EntraSitePermissionManager: {
  description: string;
  getLogo: () => SVGImageElement;
  render: (props: {
    el: HTMLElement,
    context?: WebPartContext;
    displayMode?: DisplayMode;
    envType?: number;
    flowId?: string;
    timeFormat?: string;
    title?: string;
    sourceUrl?: string;
  }) => void;
  timeFormat: string;
  title: string;
  updateTheme: (currentTheme: Partial<IReadonlyTheme>) => void;
};

export default class EntraSitePermMgrWebPart extends BaseClientSideWebPart<IEntraSitePermMgrWebPartProps> {
  private _hasRendered: boolean = false;

  public render(): void {
    // See if have rendered the solution
    if (this._hasRendered) {
      // Clear the element
      while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }
    }

    // Set the default property values
    if (!this.properties.timeFormat) { this.properties.timeFormat = EntraSitePermissionManager.timeFormat; }
    if (!this.properties.title) { this.properties.title = EntraSitePermissionManager.title; }
    if (!this.properties.webUrl) { this.properties.webUrl = this.context.pageContext.web.serverRelativeUrl; }

    // Render the application
    EntraSitePermissionManager.render({
      el: this.domElement,
      context: this.context,
      displayMode: this.displayMode,
      envType: Environment.type,
      flowId: this.properties.flowId,
      timeFormat: this.properties.timeFormat,
      title: this.properties.title,
      sourceUrl: this.properties.webUrl
    });

    // Set the flag
    this._hasRendered = true;
  }

  // Add the logo to the PropertyPane Settings panel
  protected onPropertyPaneRendered(): void {
    const setLogo = setInterval(() => {
      let closeBtn = document.querySelectorAll("div.spPropertyPaneContainer div[aria-label='Entra Site Permission Manager property pane'] button[data-automation-id='propertyPaneClose']");
      if (closeBtn) {
        closeBtn.forEach((el: HTMLElement) => {
          let parent = el.parentElement;
          if (parent && !(parent.firstChild as HTMLElement).classList.contains("logo")) { parent.prepend(EntraSitePermissionManager.getLogo()) }
        });
        clearInterval(setLogo);
      }
    }, 50);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    EntraSitePermissionManager.updateTheme(currentTheme);
  }

  protected get dataVersion(): Version {
    return Version.parse(this.context.manifest.version);
  }

  protected get disableReactivePropertyChanges(): boolean { return true; }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.SettingsGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  description: strings.TitleFieldDescription
                }),
                PropertyPaneTextField('flowId', {
                  label: strings.FlowIdFieldLabel,
                  description: strings.FlowIdFieldDescription
                }),
                PropertyPaneTextField('timeFormat', {
                  label: strings.TimeFormatFieldLabel,
                  description: strings.TimeFormatFieldDescription
                })
              ]
            }
          ]
        },
        {
          groups: [
            {
              groupName: strings.AboutGroupName,
              groupFields: [
                PropertyPaneLabel('version', {
                  text: "Version: " + this.context.manifest.version
                }),
                PropertyPaneLabel('description', {
                  text: EntraSitePermissionManager.description
                }),
                PropertyPaneLabel('about', {
                  text: "We think adding sprinkles to a donut just makes it better! SharePoint Sprinkles builds apps that are sprinkled on top of SharePoint, making your experience even better. Check out our site below to discover other SharePoint Sprinkles apps, or connect with us on GitHub."
                }),
                PropertyPaneLabel('support', {
                  text: "Are you having a problem or do you have a great idea for this app? Visit our GitHub link below to open an issue and let us know!"
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLink('supportLink', {
                  href: "https://www.spsprinkles.com/",
                  text: "SharePoint Sprinkles",
                  target: "_blank"
                }),
                PropertyPaneLink('sourceLink', {
                  href: "https://github.com/spsprinkles/entra-site-permission-manager/",
                  text: "View Source on GitHub",
                  target: "_blank"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
