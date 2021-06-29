import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import PnPTelemetry from "@pnp/telemetry-js";

import * as strings from 'RolesWebpartWebPartStrings';
import RolesWebpart from './components/RolesWebpart';
import { IRolesWebpartProps } from './components/IRolesWebpartProps';

const telemetry = PnPTelemetry.getInstance();
telemetry.optOut();

export interface IRolesWebpartWebPartProps {
  description: string;
}

export default class RolesWebpartWebPart extends BaseClientSideWebPart<IRolesWebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IRolesWebpartProps> = React.createElement(
      RolesWebpart,
      {
        description: this.properties.description,
        context: this.context,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
