import * as React from 'react';
import { IRolesWebpartProps } from './IRolesWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping} from "@pnp/spfx-controls-react/lib/ListView";
import { SPHttpClient } from '@microsoft/sp-http';
import { ComboBox, IComboBoxOption, IComboBox, PrimaryButton } from 'office-ui-fabric-react/lib/index';
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IRolesViewState {
  items: any[];
  Name?: String | null;
  SingleSelect: any;
}

export default class RolesWebpart extends React.Component<IRolesWebpartProps,IRolesViewState, {}> {
  constructor(props: IRolesWebpartProps, state: IRolesViewState) {
    super(props);
    this.state = {
      items: [],
      Name: "",
      SingleSelect: "Product Management"
    };
  }

  public componentDidMount() {
    this._loadList("Product Management");
  }

  public componentDidUpdate() {
  }

  public _loadList(title):void { 
    const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Release Planning Step 1 - Responsibilities by Role')/items?$filter=Role eq '${title}'`;
    this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); })
      .then(items => {
        this.setState({
          items: items.value ? items.value : []
        });
      });
  }
  public onComboBoxChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    this._loadList(option.key);
    this.setState({ SingleSelect: option.key });
  }

  public render(): React.ReactElement<IRolesWebpartProps> {
    const viewFields: IViewField[] = [
      {
        name: 'Role',
        displayName: 'Role',
        sorting: true,
        maxWidth: 100
      },
      {
        name: 'Level',
        displayName: 'Level',
        sorting: true,
        maxWidth: 100
      },
      {
        name: 'Prerequisites',
        displayName: 'Prerequisites',
        sorting: true,
        maxWidth: 130
      },
      {
        name: 'DeliverablesandActionItems',
        displayName: 'Deliverables and Action Items',
        sorting: true,
        maxWidth: 130
      }
    ];
    return (
      <div>
      <ComboBox
        placeholder="Select Role"
        selectedKey={this.state.SingleSelect}
        options={
          [ 
            { key: 'Product Management', text: 'Project Management' },  
            { key: 'Quality Engineering', text: 'Quality Engineering' },  
            { key: 'User Experience', text: 'User Experience' },  
          ] }
        onChange={this.onComboBoxChange}
      />
        <br/>
        <ListView
                items={this.state.items}
                viewFields={viewFields} />
      </div>
    );
  }
}
