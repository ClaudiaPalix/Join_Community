import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import {
  BaseClientSideWebPart,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
} from '@microsoft/sp-webpart-base';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from "JoinCommunityWebPartStrings";
import JoinCommunity from "./components/JoinCommunity";
import { IJoinCommunityProps } from "./components/IJoinCommunityProps";

import { app } from "@microsoft/teams-js";

import { AadTokenProvider } from "@microsoft/sp-http";
import { IYammerProvider } from "../../utils/yammer/IYammerProvider";
import YammerProvider from "../../utils/yammer/YammerProvider";


export interface IJoinCommunityWebPartProps {
  description: string;
  listSelected: string;
  listSelected2: string;
  seeAll: string;
}

export default class JoinCommunityWebPart extends BaseClientSideWebPart<IJoinCommunityWebPartProps> {
  private aadToken: string = "";
  private availableLists: IPropertyPaneDropdownOption[] = [];

  private isRunningOnTeams = false;
  private isBodyEmbedded = false;

  public async onInit(): Promise<void> {
    try {
      await app.initialize();
      const context = await app.getContext();
      console.log("Context:", context);
      if(context.app.host.name.includes("teams") || context.app.host.name.includes("Teams")){
        console.log("The webpart is running inside Microsoft Teams");
        this.isRunningOnTeams = true;
      }else{
        console.log("The webpart is running outside Microsoft Teams");
      }
    } catch (exp) {
        console.log("The webpart is running outside Microsoft Teams");
    }
    this.isBodyEmbedded = document.body.classList.contains('embedded');
    if (this.isBodyEmbedded) {
      console.log('Body has the embedded class');
    } else {
      console.log('Body does not have the embedded class');
    }
    const tokenProvider: AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
    await tokenProvider
      .getToken("https://api.yammer.com")
      .then((token) => {
        this.aadToken = token;
      })
      .catch((err) => console.log(err));
  }

  public render(): void {
    let yammerProvider: IYammerProvider = new YammerProvider(
      this.aadToken,
      this.context.pageContext.user.email
    );

    const element: React.ReactElement<IJoinCommunityProps> =
      React.createElement(JoinCommunity, {
        context: this.context,
        yammerProvider,
        userEmail: this.context.pageContext.user.email,
        selectedList: this.properties.listSelected,
        selectedList2: this.properties.listSelected2,
        seeAllUrl: this.properties.seeAll,
        description: this.properties.description,
        isTeams: this.isRunningOnTeams,
        isEmbedded: this.isBodyEmbedded
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
  protected onPropertyPaneConfigurationStart(): void {
    this._loadLists();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'listSelected') {
      this.setListTitle(newValue);
    }
    if(propertyPath === 'listSelected2'){
      this.setListTitle2(newValue);
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private _loadLists(): void {
    const listsUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`;
    //SPHttpClient is a class provided by Microsoft that allows developers to perform HTTP requests to SharePoint REST APIs or other endpoints within SharePoint or the host environment. It is used for making asynchronous network requests to SharePoint or other APIs in SharePoint Framework web parts, extensions, or other components.
    this.context.spHttpClient.get(listsUrl, SPHttpClient.configurations.v1)
    //SPHttpClientResponse is the response object returned after making a request using SPHttpClient. It contains information about the response, such as status code, headers, and the response body.
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: { value: any[] }) => {
        this.availableLists = data.value.map((list) => {
          return { key: list.Title, text: list.Title };
        });
        this.context.propertyPane.refresh();
      })
      .catch((error) => {
        console.error('Error fetching lists:', error);
      });
  }

  private setListTitle(listSelected: string): void {
    this.properties.listSelected = listSelected;
    this.context.propertyPane.refresh();
  }
  private setListTitle2(listSelected: string): void {
    this.properties.listSelected2 = listSelected;
    this.context.propertyPane.refresh();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.DescriptionFieldLabel,
          },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Title For The Application"
                }),
                PropertyPaneDropdown('listSelected', {
                  label: "Select A List to Display User's Current Communities",
                  options: this.availableLists,
                }),
                PropertyPaneDropdown('listSelected2', {
                  label: "Select A List to Display User's Current Community Requests",
                  options: this.availableLists,
                }),
                PropertyPaneTextField('seeAll',{
                  label: 'Url for See All button',
                })
              ],
            },
          ],
        }
      ]
    };
  }}
