import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from './BoLearningCommonToolsWebpart.module.scss';
import { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse, ODataVersion, ISPHttpClientConfiguration } from '@microsoft/sp-http';
import { IBoLearningCommonToolsWebpartProps, IBoLearningCommonToolsWebpartState } from './IBoLearningCommonToolsWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class BoLearningCommonToolsWebpart extends React.Component<IBoLearningCommonToolsWebpartProps, IBoLearningCommonToolsWebpartState> {

  public constructor(props: IBoLearningCommonToolsWebpartProps, state: IBoLearningCommonToolsWebpartState) {
    super(props);

    this.state = {
      items: [
        {
          "Title": "",
          "NavigateURL": "",
          "Logo": {
            "Url": "",
            "Description": "",
          }
        }
      ]
    };
    this._OpenPrptyPane = this._OpenPrptyPane.bind(this);
  }

  public componentDidMount(): void {
    if (this.props.ListTitle) {
      this.fetchCategoryfromSharePointList();
    }
  }
  public componentDidUpdate(prevProps, prevState): void {
    //check is there any change  on properties
    //&& this.state.items.length !== prevState.items.length
    if (this.props.ListTitle !== prevProps.ListTitle) {
      if (this.props.ListTitle)
        this.fetchCategoryfromSharePointList();
    }
  }

  private fetchCategoryfromSharePointList() {
    let currentWebUrl = this.props.SiteUrl;
    let listname = this.props.ListTitle;
    let requestUrl = currentWebUrl + "/_api/web/lists/getbytitle('" + listname + "')/items";
    this.props.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      if (response.ok) {
        response.json().then((responseJSON) => {
          if (responseJSON != null && responseJSON.value != null) {
            //  console.log(responseJSON.value);
            this.setState({
              items: responseJSON.value
            });
          }
        });
      }
    });
  }

  public render(): React.ReactElement<IBoLearningCommonToolsWebpartProps> {
    if (this.needsConfiguration()) {
      return (
        <div className={styles.configtab}>
          <div className={styles.Innerconfigtab}>
            <span>Please Configure this Webpart</span>
            <div><button onClick={this._OpenPrptyPane}>Configure</button></div>
          </div>
        </div>);
    }
    else {
      return (<div className={styles.ParentCatogoryDiv}>
        <div className={styles.ParentHeadline} >
          {this.state.items.map((item, key) => {
            return (
              <div className={styles.topHeadline}>
                <a className={styles.categoryatag} href={item.NavigateURL}>
                  <div className={styles.categorybox} key={key}>
                    <div className={styles.categoryTitle}>{item.Title}</div> 
                    <img className={styles.categoryIcon} src={item.Logo.Url}/>
                  </div>
                </a>
              </div>
            );
          })}
        </div>
      </div>
      );
    }
  }
  private _OpenPrptyPane(): void {
    this.props.context.propertyPane.open();
  }

  private needsConfiguration(): boolean {
    return BoLearningCommonToolsWebpart.isEmpty(this.props.ListTitle);
  }

  private static isEmpty(value: string): boolean {
    return value === undefined ||
      value === null ||
      value.length === 0;
  }
}
