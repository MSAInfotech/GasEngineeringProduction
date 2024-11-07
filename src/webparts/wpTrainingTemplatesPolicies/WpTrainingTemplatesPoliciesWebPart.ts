import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WpTrainingTemplatesPoliciesWebPartStrings';
import WpTrainingTemplatesPolicies from './components/WpTrainingTemplatesPolicies';
import { IWpTrainingTemplatesPoliciesProps } from './components/IWpTrainingTemplatesPoliciesProps';
import { SPFI, spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { stringIsNullOrEmpty } from '@pnp/core';
export interface IWpTrainingTemplatesPoliciesWebPartProps {
  ListName: string;
  FilterCols: string;
  DataCols: string;
}

export default class WpTrainingTemplatesPoliciesWebPart extends BaseClientSideWebPart<IWpTrainingTemplatesPoliciesWebPartProps> {

  private _sp: SPFI;
  private _currentUserID: number = 0;
  private _currentUserName: string = "";
  private _currentUserEmail: string = "";
  private _webRelativeUrl: string = "";
  private _folderUrl: string = "";

  public render(): void {
    const element: React.ReactElement<IWpTrainingTemplatesPoliciesProps> = React.createElement(
      WpTrainingTemplatesPolicies,
      {
        sp: this._sp,
        currentUserID: this._currentUserID,
        currentUserName: this._currentUserName,
        currentUserEmail: this._currentUserEmail,
        WebServerRelativeURL: this._webRelativeUrl,
        LibName: this.properties.ListName,
        FilterColumns: this.properties.FilterCols,
        DataColumns: this.properties.DataCols,
        FolderURL: this._folderUrl,
        context:this.context,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._webRelativeUrl = this.context.pageContext.web.serverRelativeUrl;
    this._sp = spfi().using(SPFx(this.context));
    this._currentUserEmail = this.context.pageContext.user.loginName;
    return this._getCurrentUserDetails().then(() => {
    });
  }

  private async _getCurrentUserDetails(): Promise<string> {
    const url = new URL(window.location.href);
    const params = new URLSearchParams(url.search);
    var folderUrl: any = params.get("Folder");

    if (!stringIsNullOrEmpty(folderUrl)) {
      this._folderUrl = folderUrl;
    }

    let user = await this._sp.web.currentUser();
    if (user) {
      this._currentUserID = user.Id;
      this._currentUserName = user.Title;
    }
    return Promise.resolve("")
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
                PropertyPaneTextField('ListName', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('DataCols', {
                  label: strings.DataColumnsFieldsLabel
                }),
                PropertyPaneTextField('FilterCols', {
                  label: strings.FilterColumnsFieldsLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
