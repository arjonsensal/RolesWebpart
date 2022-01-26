import * as React from 'react';
import { IRolesWebpartProps } from './IRolesWebpartProps';
import { ListView, IViewField } from "@pnp/spfx-controls-react/lib/ListView";
import { SPHttpClient } from '@microsoft/sp-http';
import { ComboBox, IComboBoxOption, IComboBox } from 'office-ui-fabric-react/lib/index';
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IRolesViewState {
  choice: any[];
  items: any[];
  Name?: String | null;
  SingleSelect: any;
  unique?: String | null;
  list: any[];
}

export default class RolesWebpart extends React.Component<IRolesWebpartProps,IRolesViewState, {}> {
  constructor(props: IRolesWebpartProps, state: IRolesViewState) {
    super(props);
    this.state = {
      items: [],
      choice: [],
      Name: "",
      SingleSelect: "",
      unique: "",
      list: []
    };
  }

  public componentDidMount() {
    this._loadChoices();
  }

  public componentDidUpdate() {
  }

  public _loadList(title):void { 
    const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/items?$filter=${this.props.unique} eq '${title}'`;
    this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); })
      .then(items => {
        this.setState({
          items: items.value ? items.value : []
        });
      });
  }
  
  public _loadChoices():void { 
    const restApi2 = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/fields?$filter=ReadOnlyField eq false and Hidden eq false`;
    this.props.context.spHttpClient.get(restApi2, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); })
      .then(items => {
        items.value.forEach((item, index) => {
          if (index > 2) {
            var joined = this.state.list.concat(
              {
                name: item.StaticName, 
                displayName: item.Title, 
                maxWidth: (item.Title === this.state.unique) ? 150 : 200, 
                render: (itemx) =>{
                  return <div dangerouslySetInnerHTML={{__html: itemx[item.StaticName]}} />;
                }
              });
            this.setState({ list: joined });
          }
        });
      });
    const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/items`;
    this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); })
      .then(items => {
        let tempIt = [];
        items.value.forEach((item, index) => {
          tempIt.push({key: item[this.props.unique], text: item[this.props.unique]});
        });
        this.setState({
          choice: tempIt ? tempIt : []
        });
      });
  }
  public onComboBoxChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    this._loadList(option.key);
    this.setState({ SingleSelect: option.key });
  }

  public render(): React.ReactElement<IRolesWebpartProps> {
    const viewFields: IViewField[] = this.state.list;
    return (
      <div>
      <h2>{this.props.description}</h2>
      <ComboBox
        placeholder="Select Role"
        selectedKey={this.state.SingleSelect}
        options={this.state.choice}
        onChange={this.onComboBoxChange}
      />
        <br/>
        <ListView
                items={this.state.items}
                compact={true}
                viewFields={viewFields} />
      </div>
    );
  }
}
