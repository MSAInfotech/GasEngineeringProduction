import * as React from 'react';
import styles from './WpFooterComponent.module.scss';
import type { IWpFooterComponentProps } from './IWpFooterComponentProps';
import { SPFI, spfi, SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";


export interface IWpMainFooterState {
  FooterItem: any[];
}

export default class WpFooterComponent extends React.Component<IWpFooterComponentProps, IWpMainFooterState> {
  private _sp: SPFI;
  constructor(props: any) {
    super(props);
    sp: this._sp,
      //this._sp = spfi("https://sempra.sharepoint.com/sites/gasopscon/eng/").using(SPFx(this.props.context));
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      FooterItem: []
    }
  }

  public componentDidMount(): void {
    debugger;
    this.fetchData();

    setTimeout(function () {
      const headerText: any = document.querySelector('#O365_AppName > span');
      console.log(headerText);
      if (headerText) {
        console.log('body onload - ' + headerText);
        headerText.innerHTML = "Gas Engineering - Project Engineering Resource Hub";
      }
    }, 2500);
  }

  protected async fetchData(): Promise<string> {
    try {
      let _FooterItem: any[] = [];
      let tempFooterMain = await this._sp.web.lists.getByTitle("FooterDetail").items.select("ID", "Title", "PoweredBy").orderBy("ID", false).top(1)();
      debugger;
      if (tempFooterMain && tempFooterMain.length > 0) {
        _FooterItem.push({
          "Id": tempFooterMain[0].ID,
          "Title": tempFooterMain[0].Title,
          "PoweredBy": tempFooterMain[0].PoweredBy
        });
      } else {
        _FooterItem = [];
      }
      this.setState({
        FooterItem: _FooterItem
      });
      return Promise.resolve("");
    } catch (error) {
      // Handle error
      return Promise.resolve("");
    }
  }

  public render(): React.ReactElement<IWpFooterComponentProps> {

    return (
      <section className={`${styles.wpFooterComponent}`}>
        {this.state.FooterItem.map((nhItem, index) => (
          <div className="row">
            <div className="firstPart col-md-12 col-lg-12 col-sm-12">{nhItem.Title}</div>
            {/* <div className="secondPart col-md-5 col-lg-5">{nhItem.PoweredBy}</div> */}
          </div>
        ))
        }
      </section>
    );
  }
}
