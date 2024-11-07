import * as React from 'react';
import styles from './WpFrequentlyAccessedSites.module.scss';
import type { IWpFrequentlyAccessedSitesProps } from './IWpFrequentlyAccessedSitesProps';
import type { IWpFrequentlyAccessedSitesState } from './IWpFrequentlyAccessedSitesState';
import { SPFI, spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { stringIsNullOrEmpty } from '@pnp/core';

export default class WpFrequentlyAccessedSites extends React.Component<IWpFrequentlyAccessedSitesProps, IWpFrequentlyAccessedSitesState> {
  private _sp: SPFI;


  constructor(props: any) {
    super(props);
    sp: this._sp,
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      error: null,
      listAllItems: [],
      graphApiItems: [],
      displayHTMLArray: [],
      filterArray: [],
      image: false,
    }
  }
  public componentDidMount(): void {
    this._getFrequentlyUsedSites();
    // this.getTopAccessedPages();
  }

  // Method to get top accessed pages
  // private async getTopAccessedPages(): Promise<any[]> {
  //   const user = await this._sp.web.currentUser();
  //   const list = this._sp.web.lists.getByTitle("FrequentlyAccessSiteData");

  //   const items = await list.items.filter(`User eq '${user.LoginName}'`).orderBy("AccessCount", false).top(5)();

  //   return items;
  // }

  public async _getFrequentlyUsedSites() {
    let _listItems: any[] = [];
    const sp = spfi().using(SPFx(this.props.context));
    const user = await this._sp.web.currentUser();
    try {
      if (!stringIsNullOrEmpty(this.props.ListName)) {
        const pinnedItems: any[] = await sp.web.lists.getByTitle(this.props.ListName).items
          .select("*")
          .filter(`AccessSite eq 1 and User/Id eq '${user.Id}'`)
          .orderBy("AccessCount", false)
          .top(3)();

        const userItems: any[] = await sp.web.lists.getByTitle(this.props.ListName).items
          .select("*")
          .filter(`User/Id eq '${user.Id}' and AccessSite eq 0`)
          .orderBy("AccessCount", false)
          .top(3)();

        const combinedItems: any[] = [...pinnedItems, ...userItems];
        const limitedItems: any[] = combinedItems.slice(0, 3);

        console.log(limitedItems);
        if (limitedItems && limitedItems.length > 0) {
          limitedItems.map((item: any) => {
            _listItems.push({ "Id": item.Id, 'siteName': item.Title, 'webUrl': item.PageURL, 'accessCount': item.AccessCount, 'accessSite': item.AccessSite });
          });
          console.log(_listItems);
        }
      }
      if (!stringIsNullOrEmpty(this.props.ListName)) {
        this.setState({ listAllItems: _listItems });
        console.log('No Data in Frequently List')
      }
    }
    catch (e) {
      console.error(e.error);
    }
  }

  // public async addItem(site: any) {
  //   // add an item to the list
  //   const sp = spfi().using(SPFx(this.props.context));
  //   const item = await sp.web.lists.getByTitle(this.props.ListName).items.add({
  //     Title: site.siteName,
  //     WebUrl: site.PageURL,
  //     AccessSite: true,
  //   });
  //   console.log(item);
  //   console.log("Create item succressfully");
  //   this._getFrequentlyUsedSites();
  // }

  // public async deleteItem(site: any) {
  //   if (site.Id != 0) {
  //     const sp = spfi().using(SPFx(this.props.context));
  //     await sp.web.lists.getByTitle(this.props.ListName).items.getById(site.Id).delete();
  //     this._getFrequentlyUsedSites();
  //   }
  // }

  public async addItem(site: any) {
    // add an item to the list
    const sp = spfi().using(SPFx(this.props.context));

    await sp.web.lists.getByTitle(this.props.ListName).items.getById(site.Id).update({
      AccessSite: true,
    });
    console.log(`Updated item with ID: ${site.Id}`);

    this._getFrequentlyUsedSites();
  }

  public async deleteItem(site: any) {
    if (site.Id != 0) {
      const sp = spfi().using(SPFx(this.props.context));
      await sp.web.lists.getByTitle(this.props.ListName).items.getById(site.Id).update({
        AccessSite: false,
      });
      console.log(`Updated item with ID: ${site.Id}`);
      this._getFrequentlyUsedSites();
    }
  }

  public render(): React.ReactElement<IWpFrequentlyAccessedSitesProps> {
    return (
      <section className={`${styles.wpFrequentlyAccessedSites} col-md-4 col-sm-4 col-lg-4`}>
        <div className={styles.wpFrequentlyAccessedSites}>
          <div>
            <div>
              <div>
                <span className={styles.titleFaq}>Frequently Accessed</span>
                {this.state.error ? (
                  <div >{this.state.error}</div>
                ) : (
                  <ul>
                    {this.state.listAllItems.map((site: any) => (
                      <li key={site.id}>
                        <img src={this.props.webUrl + "/SiteAssets/PortalImages/globe.png"} className={styles.icFaqone} width="35px" />
                        <a href={site.webUrl}>
                          {site.siteName}
                        </a>
                        {site.accessSite == true &&
                          <img onClick={() => { this.deleteItem(site) }} src={this.props.webUrl + "/SiteAssets/PortalImages/pin.png"} className={styles.icFaq} width="35px" />
                        }
                        {site.accessSite == false &&
                          <img onClick={() => { this.addItem(site) }} src={this.props.webUrl + "/SiteAssets/PortalImages/pinoutline.png"} className={styles.icFaq} width="35px" />
                        }
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            </div>
          </div>
        </div>
      </section>
    );
  }
}
