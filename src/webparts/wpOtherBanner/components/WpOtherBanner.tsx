import * as React from 'react';
import styles from './WpOtherBanner.module.scss';
import type { IWpOtherBannerProps } from './IWpOtherBannerProps';

import { SPFI, spfi, SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
//import { IconButton } from '@fluentui/react/lib/Button';
// import { TooltipHost, ITooltipHostStyles } from '@fluentui/react/lib/Tooltip';

export interface IWpOtherBannerState {
  BannerItemsState: any[];
  linkUrl: string;
  setLinkUrl: string;
  Title: string;
}

export default class WpOtherBanner extends React.Component<IWpOtherBannerProps, IWpOtherBannerState> {

  private _sp: SPFI;

  constructor(props: any) {
    super(props);
    sp: this._sp,
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      BannerItemsState: [],
      linkUrl: "",
      setLinkUrl: "",
      Title: ""

    }
  }

  public componentDidMount(): void {
    //this.checkUserInGroup("PortalAdmins");

    this.fetchData();
  }

  protected async fetchData() {
    try {
      await this._getBannerItems();
    } catch (error) {
      // Handle error
    }
  }

  private async _getBannerItems(): Promise<string> {
    let _BannerItem: any[] = [];
    //let _QuickLinksArrItems = [];
    this.setState({ BannerItemsState: [] });
    let categoryfilter = this.props.category;
    debugger;
    let tmpBanner = await this._sp.web.lists.getByTitle("Banners").items.select("ID", "Title", "Category", "BannerImg").filter("Category eq '" + categoryfilter + "'").orderBy("ID", true).top(1)();
    if (tmpBanner && tmpBanner.length > 0) {
      tmpBanner.forEach(element => {
        var imageURL = "";
        if (element.BannerImg != null && element.BannerImg != "") {
          let filename = JSON.parse(element.BannerImg).fileName;
          console.log(this.props.context.pageContext.web.absoluteUrl);
          imageURL = "https://sempra.sharepoint.com/sites/gasopscon/eng/resource%20hub/Lists/Banners/Attachments/" + element.ID + "/" + filename;
        }
        _BannerItem.push(
          {
            "ID": element.ID,
            "Title": element.Title,
            //"RedirectURL": element.RedirectURL,
            //"ImageURL":element.ImageURL,
            "ImageURL": imageURL,
          });
      })

    } else {
      _BannerItem = [];
    }
    this.setState({ BannerItemsState: _BannerItem })
    console.log(this.state.BannerItemsState);
    return Promise.resolve("")
  }

  public render(): React.ReactElement<IWpOtherBannerProps> {
    // const iconButtonStyles = {
    //   root: {
    //     fontSize: '24px', // Increase the size of the icon
    //     color: 'white', // Change the color of the icon
    //   },
    //   rootHovered: {
    //     color: 'black', // Change the color of the icon on hover
    //   },
    // };

    return (
      <section className={`${styles.wpMainBanner}`}>
        <div className={`${styles.Banner_img}`}>
          {/* {this.props.category === 'HomePage' ?
            <>
              <div className={`${styles.Small_logo_img}`}>
                <img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/img/SoCalGas_logo_01_white%201.png"} />
                <img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/img/sdge%20copy2%201.png"} className='mt-2 ps-2' />
              </div>
              <div className={`${styles.User_new}`}>
                <p>Hi {this.props.userDisplayName}!</p>
              </div>
            </>
            : ''} */}

          {this.state.BannerItemsState &&
            this.state.BannerItemsState.map((nhItem, index) => (
              <div className={`${styles.User_dec}`}  style={{ marginLeft: '4.5%', marginRight: '4.5%'}}>
                <img src={nhItem.ImageURL} />
                {/* <p>{nhItem.Title}
                  {this.props.category === 'HomePage' ? '' : <TooltipHost
                    content={this.state.toolTipBanner}
                    styles={tooltipStyles}
                  >
                    <IconButton
                      iconProps={{ iconName: 'Info' }}
                      title="Info"
                      ariaLabel="Info"
                      styles={iconButtonStyles} // Apply custom styles here
                    />
                  </TooltipHost>}
                </p> */}
              </div>
            ))
          }
        </div>
      </section>
    );
  }
}
