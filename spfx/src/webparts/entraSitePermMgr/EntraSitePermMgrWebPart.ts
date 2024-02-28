import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'EntraSitePermMgrWebPartStrings';

export interface IEntraSitePermMgrWebPartProps {
  flowId: string;
}

// Reference the solution
import "../../../../dist/entra-site-permission-manager.js";
declare const EntraSitePermissionManager: {
  render: (props: {
    el: HTMLElement,
    context?: WebPartContext;
    flowId?: string;
    url?: string;
  }) => void;
};

export default class EntraSitePermMgrWebPart extends BaseClientSideWebPart<IEntraSitePermMgrWebPartProps> {

  public render(): void {
    // Render the application
    EntraSitePermissionManager.render({
      el: this.domElement,
      context: this.context,
      flowId: this.properties.flowId
    });
  }

  //protected onInit(): Promise<void> { }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      // TODO
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneFlowId
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('flowId', {
                  label: strings.FlowIdFieldLabel,
                  description: "The flow id of the Entra Site Permission Manager flow."
                })
              ]
            }
          ]
        }
      ]
    };
  }
}