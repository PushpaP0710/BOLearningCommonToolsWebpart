import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IODataList } from '@microsoft/sp-odata-types';
import { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse, ODataVersion, ISPHttpClientConfiguration } from '@microsoft/sp-http';
import { BaseClientSideWebPart, IPropertyPaneDropdownOption } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';

import * as strings from 'BoLearningCommonToolsWebpartWebPartStrings';
import BoLearningCommonToolsWebpart from './components/BoLearningCommonToolsWebpart';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IBoLearningCommonToolsWebpartProps } from './components/IBoLearningCommonToolsWebpartProps';

export interface IBoLearningCommonToolsWebpartWebPartProps {
  description: string;
  dropdownProperty: string;
  context: WebPartContext;
}

export default class BoLearningCommonToolsWebpartWebPart extends BaseClientSideWebPart<IBoLearningCommonToolsWebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IBoLearningCommonToolsWebpartProps> = React.createElement(
      BoLearningCommonToolsWebpart,
      {
        ListTitle: this.properties.dropdownProperty,
        SiteUrl: this.context.pageContext.web.absoluteUrl,
        context: this.context
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

  private dropdownOptions: IPropertyPaneDropdownOption[];
  private listsFetched: boolean;

  // these methods are split out to go step-by-step, but you could refactor and be more direct if you choose..

  private fetchLists(url: string): Promise<any> {
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      if (response.ok) {
        return response.json();
      } else {
        console.log("WARNING - failed to hit URL " + url + ". Error = " + response.statusText);
        return null;
      }
    });
  }


  private fetchOptions(): Promise<IPropertyPaneDropdownOption[]> {
    var url = this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`;

    return this.fetchLists(url).then((response) => {
      var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
      response.value.map((list: IODataList) => {
        console.log("Found list with title = " + list.Title);
        options.push({ key: list.Title, text: list.Title });
      });

      return options;
    });
  }

  protected onPropertyPaneConfigurationStart(): void {
    // loads list name into list dropdown  
    this.fetchOptions();

  }

  protected get disableReactivePropertyChanges(): boolean {   
    return true;   
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    if (!this.listsFetched) {
      this.fetchOptions().then((response) => {
        this.dropdownOptions = response;
        this.listsFetched = true;
        // now refresh the property pane, now that the promise has been resolved..
        this.context.propertyPane.refresh();
        this.onDispose();
      });
   }
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
                
                PropertyPaneDropdown('dropdownProperty', {
                  label: strings.ListNameFieldLabel,
                  options: this.dropdownOptions
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
