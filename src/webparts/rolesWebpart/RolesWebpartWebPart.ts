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
  columns: string;
  filterList: string;
  uniqueFilter: string;
  optionalColumnFilter: string;
  optionalColumnFilterValue: string;
  removeColumns: string;
}

export default class RolesWebpartWebPart extends BaseClientSideWebPart<IRolesWebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IRolesWebpartProps> = React.createElement(
      RolesWebpart,
      {
        description: this.properties.description,
        listName: this.properties.listName,
        context: this.context,
        unique: this.properties.unique,
        columns: this.properties.columns,
        filterList: this.properties.filterList,
        uniqueFilter: this.properties.uniqueFilter,
        optionalColumnFilter: this.properties.optionalColumnFilter,
        optionalColumnFilterValue: this.properties.optionalColumnFilterValue,
        removeColumns: this.properties.removeColumns
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
  
  private async loadLists1(): Promise<IDropdownOption[]> {
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
    // refresh web part
    this.render();
    // this.context.propertyPane.refresh();
  }

  
  private onListChange1(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    // this.context.propertyPane.refresh();
    // refresh web part
    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('2.0');
  }

  // protected get disableReactivePropertyChanges(): boolean {
  //   return true;
  // }
  
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
                  label: "Column to use in Dropdown"
                }),
                PropertyPaneTextField('columns', {
                  label: "Columns to Include in Grid"
                })
              ]
            },
            {
              groupFields: [
                new PropertyPaneAsyncDropdown('filterList', {
                  label: "Select List for dynamic filtering",
                  loadOptions: this.loadLists1.bind(this),
                  onPropertyChange: this.onListChange1.bind(this),
                  selectedKey: this.properties.filterList
                }),
                PropertyPaneTextField('uniqueFilter', {
                  label: "Column to Filter from Dropdown"
                })
              ]
            },
            {
              groupFields: [
                PropertyPaneTextField('optionalColumnFilter', {
                  label: "Optional filter (Select Column to Filter Based on Page)"
                }),
                PropertyPaneTextField('optionalColumnFilterValue', {
                  label: "Page Name"
                })
              ]
            },
            {
              groupFields: [
                PropertyPaneTextField('removeColumns', {
                  label: "Remove Columns in Card List (Seperate by Comma. Ex. Column1,Column2..)"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
