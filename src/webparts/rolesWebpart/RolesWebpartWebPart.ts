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
import { PropertyPaneAsyncDropdown } from './controls/PropertyPaneAsyncDropdown';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { update, get } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';

const telemetry = PnPTelemetry.getInstance();
telemetry.optOut();

export interface IRolesWebpartWebPartProps {
  description: string;
  listName: string;
  unique: string;
}

export default class RolesWebpartWebPart extends BaseClientSideWebPart<IRolesWebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IRolesWebpartProps> = React.createElement(
      RolesWebpart,
      {
        description: this.properties.description,
        listName: this.properties.listName,
        context: this.context,
        unique: this.properties.unique
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private async loadLists(): Promise<IDropdownOption[]> {
    return new Promise<IDropdownOption[]>(async(resolve: (options: IDropdownOption[]) => void, reject: (error: any) => void) => {
      
      setTimeout(async ()=> {
        var a = new Array();
        const restApi = `${this.context.pageContext.web.absoluteUrl}/_api/lists`;
        await this.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
          .then(resp => { return resp.json(); })
          .then(items => {
            items.value.forEach(element => {
              if(element.Hidden === false && element.IsCatalog === false) {
                a.push({key: element.Title, text: element.Title});
              }
            });
            resolve(a);
          });
      }, 1500);
    });
  }
  
  
  private onListChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    this.context.propertyPane.refresh();
    // refresh web part
    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('2.0');
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
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                new PropertyPaneAsyncDropdown('listName', {
                  label: "Select List",
                  loadOptions: this.loadLists.bind(this),
                  onPropertyChange: this.onListChange.bind(this),
                  selectedKey: this.properties.listName
                }),
                PropertyPaneTextField('unique', {
                  label: "Unique Column to Filter"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
