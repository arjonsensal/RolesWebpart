import * as React from 'react';
import { IRolesWebpartProps } from './IRolesWebpartProps';
import { ListView, IViewField } from "@pnp/spfx-controls-react/lib/ListView";
import { SPHttpClient } from '@microsoft/sp-http';
import { ComboBox, IComboBoxOption, IComboBox } from 'office-ui-fabric-react/lib/index';
import "@pnp/sp/lists";
import "@pnp/sp/items";
import './Card/Card.css';
import styles from './RolesWebpart.module.scss';

export interface IRolesViewState {
  choice: any[];
  items: any[];
  Name?: String | null;
  SingleSelect: any;
  unique?: String | null;
  list: any[];
  cardItems: any[];
  cardCol: any[];
  filteredCardItems: any[];
  listParentLink?: String | null;
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
      list: [],
      cardItems: [],
      cardCol: [],
      listParentLink: "",
      filteredCardItems: []
    };
  }

  public componentDidMount() {
    // this._loadList("Activities", "Title", "Title", 'Card');
    this._loadChoices(this.props.filterList, "Card");
    this._loadChoices(this.props.listName, "Table");
    this._loadListItemUrl(this.props.filterList);
    // this._loadChoices("Activities");
  }

  public componentDidUpdate() {
  }

  public _loadList(list, unique, title, item):void { 
    const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${list}')/items?$filter=${unique} eq '${title}'`;
    this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); })
      .then(items => {
          this.setState({
            items: items.value ? items.value : []
          });
      });
  }
  

  public _loadListItemUrl(list):void { 
    const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${list}')/Forms?$select=ServerRelativeUrl&$filter=FormType eq 4`;
    this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); })
      .then(data => {
        var url = data.value[0].ServerRelativeUrl;
        this.setState({listParentLink: url})
        return url;
      });
  }
  
  public _loadChoices(listName, itemType):void { 
    const restApi2 = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/fields?$filter=ReadOnlyField eq false and Hidden eq false`;
    this.props.context.spHttpClient.get(restApi2, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); })
      .then(items => {
        items.value.forEach((item, index) => {
          if (index > 2) {
            if (itemType !== "Card") {
              console.log(item)
              var joined = this.state.list.concat(
                {
                  name: item.StaticName, 
                  displayName: item.Title, 
                  maxWidth: (item.Title === this.state.unique) ? 150 : 200, 
                  render: (itemx) =>{
                  //   // console.log(itemx)
                  //   // const par = (disp) => {
                  //   // // if(item.TypeDisplayName.includes('lines of text')) {
                  //   //   // console.log("true")
                  console.log(itemx[item.StaticName])
                  if(itemx[item.StaticName].includes("<div")) {
                    console.log("aaa")
                    return <div dangerouslySetInnerHTML={{__html: itemx[item.StaticName]}}style={{display: "table-cell", whiteSpace: "pre-wrap"}} />;
                  } else {
                    return (
                      <p style={{width: 190, 
                        display: "table-cell",
                        whiteSpace: "pre-wrap"}}>{itemx[item.StaticName]}</p>
                    )
                  }
                  //   //   // }
                  //   // }
                  //   //   </div>;
                  }
                });
              this.setState({ list: joined });
            } else {
              var joined = this.state.cardCol.concat({stat: item.StaticName, title: item.Title});
              this.setState({ cardCol: joined });
            }
          }
        });
        
      const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items`;
      this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
        .then(resp => { return resp.json(); })
        .then(items => {
          if (itemType !== "Card") {
            let tempIt = [];
            items.value.forEach((item, index) => {
              tempIt.push({key: item[this.props.unique], text: item[this.props.unique]});
            });
            this.setState({
              choice: tempIt ? tempIt : []
            });
          } else {
            items.value.forEach(element => {
              var myKeys = Object.keys(element);
              var obj = {};
              this.state.cardCol.forEach(el => {
                var matchingKeys = myKeys.filter(function(key){ return key.indexOf(el.stat) !== -1 });
                obj[matchingKeys[0]] = element[matchingKeys[0]];
                // if (myKeys.indexOf(el) !== -1) {
                //   console.log("true")
                // }
              })
              obj['id'] = element['Id'];
              var joined3 = this.state.cardItems.concat(obj);
              this.setState({ cardItems: joined3 });
            });
            var output = this.state.cardItems.map(card => {
              this.state.cardCol.forEach(cardCol => {
                // if (cardCol.stat !== cardCol.title) {
                  if (card.hasOwnProperty(cardCol.stat) ){
                    var tmp = card[cardCol.title];
                    card[cardCol.title] = card[cardCol.stat];
                    delete card[cardCol.stat];
                    if(cardCol.stat === cardCol.title) { 
                      card[cardCol.title] = tmp;
                    }
                  } 
                // }
              })
              return card;
            });
            this.setState({ cardItems: output, filteredCardItems: output });
          }
        });
      });
    // if (itemType === "Card") return;
  }
  public onComboBoxChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    this._loadList(this.props.listName, this.props.unique, option.key, 'Table');
    this.setState({ SingleSelect: option.key });
    
    var myKeys = Object.keys(this.state.cardItems[0]);
    console.log(myKeys)
    console.log(this.props.uniqueFilter)
    var string = this.props.uniqueFilter;
    var cardKey = myKeys.filter(function(key){ return key.indexOf(string) !== -1 });
    this.setState({
      filteredCardItems: this.state.cardItems.filter(function(card) {
        return card[cardKey[0]].includes(option.key);
      })
    });
  }
  public render(): React.ReactElement<IRolesWebpartProps> {
    const viewFields: IViewField[] = this.state.list;
    console.log(this.state.choice)
    var getImageUrl = (image) => {
      if (image === null) return "https://genesisairway.com/wp-content/uploads/2019/05/no-image.jpg";
      var imageObj = JSON.parse(image)
      var url = imageObj.serverUrl + imageObj.serverRelativeUrl;
      return url;
    };

    const handleContainerClick = (e, i) =>{
      var makeAbsUrl = (strUrl) => {
        var url = strUrl.split("/sites");
        return url[0];
      };
      var link = encodeURI(makeAbsUrl(this.props.context.pageContext.web.absoluteUrl) + this.state.listParentLink + '?ID=' + i);
      window.open(link, "_blank")
    }
    return (
      <div>
        <h3>{this.props.description}</h3>
        <ComboBox
          placeholder="Select Role"
          selectedKey={this.state.SingleSelect}
          options={this.state.choice}
          onChange={this.onComboBoxChange}
          style={{width: '45%'}}
        />
          <br/>
          
        <ListView
                items={this.state.items}
                compact={true}
                viewFields={viewFields} />
        <h3>{this.props.filterList}</h3>
        <div className="wrapper">
          {this.state.filteredCardItems.map((card, i) => {
            var myKeys = Object.keys(card);
            return (
              <div className="card-container" onClick={(e) => {handleContainerClick(e, card.id)}}>
                <div className="img-container">
                    <img src={(myKeys.indexOf("Image") !== -1) ? getImageUrl(card.Image): "https://genesisairway.com/wp-content/uploads/2019/05/no-image.jpg"} alt="" style={{objectFit: 'fill'}}/>
                </div>
                <div className="body-container">
                {Object.keys(card).map((keyName, i) => (
                  (keyName !== "Image" && keyName !== "id") && 
                  <div>
                    <p className="key-name">{keyName}</p>
                    <div style={{marginTop: '-15px'}}><div className="key-desc" dangerouslySetInnerHTML={{__html: (i === 1) ? <strong>{card[keyName] !== null ? card[keyName] : "--"}</strong> : (card[keyName] !== null) ? ((typeof card[keyName] === "object") ? card[keyName].join(",") : card[keyName]) : "--"}}style={{display: "table-cell", whiteSpace: "pre-wrap", margin: '20px'}} /></div>
                    {/* <p className="key-desc">{(i === 1) ? <strong>{card[keyName] !== null ? card[keyName] : "--"}</strong> : (card[keyName] !== null) ? ((typeof card[keyName] === "object") ? card[keyName].join(",") : card[keyName]) : "--"}</p> */}
                  </div>)
                )}
                </div>
              </div>
            )
          })}
        </div>
      </div>
    );
  }
}
