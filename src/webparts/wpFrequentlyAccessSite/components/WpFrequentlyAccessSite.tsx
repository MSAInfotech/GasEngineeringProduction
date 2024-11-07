import * as React from 'react';
import styles from './WpFrequentlyAccessSite.module.scss';
import type { IWpFrequentlyAccessSiteProps } from './IWpFrequentlyAccessSiteProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { SPFI, spfi, SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";

export interface IMyWebPartState {
  topPages: any[];
}

export default class WpFrequentlyAccessSite extends React.Component<IWpFrequentlyAccessSiteProps, IMyWebPartState> {
  private _sp: SPFI;



  constructor(props: any) {
    super(props);
    sp: this._sp,
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      topPages: [],
    }
  }

  // Fetch and track page access when the component mounts
  public async componentDidMount(): Promise<void> {
    const pageUrl = window.location.href;
    await this.trackPageAccess(pageUrl);
  }

  private async trackPageAccess(pageUrl: string): Promise<void> {
    const user = await this._sp.web.currentUser();
    const userId = user.Id;  // Use user.Id for Person/Group fields
    let pageTitle = this.props.pageTitle;
    const list = this._sp.web.lists.getByTitle("FrequentlyAccessSiteData");

    const items = await list.items.select("*").filter(`User/Id eq '${userId}' and PageURL eq '${pageUrl}'`).top(1)();

    if (!items || items.length > 0) {
      const item = items[0];
      await list.items.getById(item.Id).update({
        Title: pageTitle,
        AccessCount: item.AccessCount + 1
      });
    } else {
      await list.items.add({
        Title: pageTitle,
        UserId: userId,
        PageURL: pageUrl,
        AccessCount: 1,
        AccessSite: false
      });
    }
    console.log(list);
  }

  // Method to get top accessed pages
  // private async getTopAccessedPages(): Promise<any[]> {
  //   const user = await this._sp.web.currentUser();
  //   const list = this._sp.web.lists.getByTitle("FrequentlyAccessSiteData");

  //   const items = await list.items.filter(`User eq '${user.LoginName}'`).orderBy("AccessCount", false).top(5)();

  //   return items;
  // }


  public render(): React.ReactElement<IWpFrequentlyAccessSiteProps> {
    const {

      hasTeamsContext,

    } = this.props;

    return (
      <section className={`${styles.wpFrequentlyAccessSite} ${hasTeamsContext ? styles.teams : ''}`}>
      </section>
    );
  }
}
