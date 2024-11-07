import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WpInteractiveChecklistWebPartStrings';
import WpInteractiveChecklist from './components/WpInteractiveChecklist';
import { IWpInteractiveChecklistProps } from './components/IWpInteractiveChecklistProps';

import { SPFI, spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";

export interface IWpInteractiveChecklistWebPartProps {
  ListName: string;
  FilterPageURL: string;
}

export default class WpInteractiveChecklistWebPart extends BaseClientSideWebPart<IWpInteractiveChecklistWebPartProps> {

  private _sp: SPFI;
  private _currentUserID: number = 0;
  private _currentUserName: string = "";
  private _currentUserEmail: string = "";
  private _webRelativeUrl: string = "";

  public render(): void {
    const element: React.ReactElement<IWpInteractiveChecklistProps> = React.createElement(
      WpInteractiveChecklist,
      {
        sp: this._sp,
        currentUserID: this._currentUserID,
        currentUserName: this._currentUserName,
        currentUserEmail: this._currentUserEmail,
        WebServerRelativeURL: this._webRelativeUrl,
        LibName: this.properties.ListName,
        filterPageURL: this.properties.FilterPageURL,
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
    //this._webRelativeUrl = this.context.pageContext.web.serverRelativeUrl;
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
                PropertyPaneTextField('FilterPageURL', {
                  label: strings.FilterPageURLFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
