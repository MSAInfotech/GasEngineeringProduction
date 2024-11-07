import * as React from 'react';
import styles from './WpMainBanner.module.scss';
import type { IWpMainBannerProps } from './IWpMainBannerProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPFI, spfi, SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
import { IconButton } from '@fluentui/react/lib/Button';
import { TooltipHost, ITooltipHostStyles } from '@fluentui/react/lib/Tooltip';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faGear } from '@fortawesome/free-solid-svg-icons';
import { PrimaryButton, TextField, Dialog, DialogType, DialogFooter, DefaultButton } from '@fluentui/react';

const tooltipStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
export interface IWpMainBannerState {
  BannerItemsState: any[];
  linkUrl: string;
  setLinkUrl: string;
  Title: string;
  hoverDescription: string;
  toolTipBanner: string;
  authorGuideToolTipBanner: string;
  isUserInGroup: boolean;
  _showFeedbackDialogEvents: boolean;
  FeedbackTitle: string;
  siteUrl: string;
}

export default class WpMainBanner extends React.Component<IWpMainBannerProps, IWpMainBannerState> {

  private _sp: SPFI;

  constructor(props: any) {
    super(props);
    sp: this._sp,
      //this._sp = spfi("https://sempra.sharepoint.com/sites/gasopscon/eng/").using(SPFx(this.props.context));
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      BannerItemsState: [],
      linkUrl: "",
      setLinkUrl: "",
      Title: "",
      hoverDescription: "",
      toolTipBanner: "",
      authorGuideToolTipBanner: "",
      isUserInGroup: false,
      _showFeedbackDialogEvents: false,
      FeedbackTitle: "",
      siteUrl: "",
    }
  }

  public componentDidMount(): void {
    this.checkUserInGroup("PortalAdmins");

    this.fetchData();
  }

  private async checkUserInGroup(groupName: string): Promise<void> {
    this.setState({
      siteUrl: (await this._sp.web.getContextInfo()).SiteFullUrl
    });

    const currentUser = await this._sp.web.currentUser();
    const groups = await this._sp.web.siteGroups.getByName(groupName).users();
    const isUserInGroupExists = groups.some(user => user.Id === currentUser.Id);
    this.setState({ isUserInGroup: isUserInGroupExists });
  }

  protected async fetchData() {
    try {
      await this._getBannerItems();
      await this._getToolTip();
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
    let tmpBanner = await this._sp.web.lists.getByTitle("Banners").items.select("ID", "Title", "Category", "BannerImg", "HoverDescription").filter("Category eq '" + categoryfilter + "'").orderBy("ID", true).top(1)();
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
            "HoverDescription": element.HoverDescription,
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

  private async _getToolTip(): Promise<string> {
    let _bannerTooTip = "";
    let _autherGuideTooTip = "";
    let category = this.props.category;
    let autherguidecategory = "HomePageAutherGuide";
    let tmpBannerToolTip = await this._sp.web.lists.getByTitle("ToolTip").items.select("ID", "Title", "Category").filter("Category eq '" + category + "'").orderBy("ID", false).top(1)();
    if (tmpBannerToolTip && tmpBannerToolTip.length > 0) {
      _bannerTooTip = tmpBannerToolTip[0].Title;
    } else {
      _bannerTooTip = "";
    }

    let tmpAutherGuideToolTip = await this._sp.web.lists.getByTitle("ToolTip").items.select("ID", "Title", "Category").filter("Category eq '" + autherguidecategory + "'").orderBy("ID", false).top(1)();
    if (tmpAutherGuideToolTip && tmpAutherGuideToolTip.length > 0) {
      _autherGuideTooTip = tmpAutherGuideToolTip[0].Title.replace(
        "Over J",
        "<b>Over J</b>"
      );
    } else {
      _autherGuideTooTip = "";
    }
    this.setState({
      toolTipBanner: _bannerTooTip,
      authorGuideToolTipBanner: _autherGuideTooTip
    })

    return Promise.resolve("")
  }

  private UserDisplayNameForBanner(displayName: string): string {
    let procesedName: string = '';
    if (displayName && displayName.indexOf(',') != -1) {
      procesedName = displayName.split(',')[1].trim();
      if (procesedName && procesedName.indexOf(' ') != -1) {
        procesedName = procesedName.split(' ')[0];
      }
    }
    return procesedName;
  }

  private _onFeedbackIconClick(): void {
    this.setState({
      _showFeedbackDialogEvents: !this.state._showFeedbackDialogEvents,
    });
  }


  private _onDismissEvent(): void {
    this.setState({
      _showFeedbackDialogEvents: !this.state._showFeedbackDialogEvents,
      FeedbackTitle: ""
    });
  }

  private async submitFeedback() {
    const { FeedbackTitle } = this.state;
    if (!FeedbackTitle || FeedbackTitle.trim() == '') {
      alert('Comment is required to send feedback.');
      return;
    }

    if (FeedbackTitle.length > 500) {
      alert('Feedback comment exceeds the allowed length of 500 characters.');
      return;
    }

    const currentUser = await this._sp.web.currentUser();
    try {
      await this._sp.web.lists.getByTitle("UserFeedback").items.add({
        UserComment: this.state.FeedbackTitle,
        SiteUserId: currentUser.Id,
      });
    } catch (error) {
      console.error('Failed to submit your feedback to the list.');
      return;
    }

    let portalAdmins;
    try {
      portalAdmins = await this._sp.web.siteGroups.getByName('FeedbackEmailAdmins').users();
    } catch (error) {
      console.error('Failed to retrieve admin user information.');
      return;
    }

    const adminEmailAddress = portalAdmins.map(email => email.Email);
    if (!adminEmailAddress || adminEmailAddress.length === 0) {
      console.error('No admin email addresses found to send the feedback.');
      return;
    }

    const siteInfo = await this._sp.web.select('Title')();
    const siteName = siteInfo.Title;

    this.SendAnEmail(adminEmailAddress, siteName)
    this._onDismissEvent()
  }

  private async SendAnEmail(adminEmailAddress: any, siteName: string) {
    try {
      await this._sp.utility.sendEmail({
        To: adminEmailAddress,
        Subject: `Feedback - ${siteName}`,
        Body: this.state.FeedbackTitle,
        AdditionalHeaders: {
          "content-type": "text/plain"
        },
      });
      alert('Feedback sent successfully!');
    } catch (error) {
      console.error('Error sending email: ', error);
      alert('There was an issue sending the feedback. Please try again later.');
    }
  }

  public render(): React.ReactElement<IWpMainBannerProps> {
    const iconButtonStyles = {
      root: {
        fontSize: '24px', // Increase the size of the icon
        color: 'white', // Change the color of the icon
      },
      rootHovered: {
        color: 'black', // Change the color of the icon on hover
      },
    };

    return (
      <section className={`${styles.wpMainBanner}`}>
        <div className={`${styles.Banner_img}`}>
          {this.props.category === 'HomePage' ?
            <>
              <div className={`${styles.Small_logo_img}`}>
                <img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/img/SoCalGas_logo_01_white%201.png"} />
                <img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/img/sdge%20copy2%201.png"} className='mt-2 ps-2' />
              </div>
              {this.state.isUserInGroup && (
                <div className={`${styles.gearIcon}`}>
                  <FontAwesomeIcon icon={faGear} onClick={() => {
                    window.open('https://sempra.sharepoint.com.mcas.ms/sites/gasopscon/eng/resource%20hub/_layouts/15/groups.aspx', '_blank');
                  }}
                    style={{ color: '#ffffff', fontSize: '24px', cursor: 'pointer' }} />
                </div>
              )}



              <div className={`${styles.User_new}`}>
                <p>Hi {this.UserDisplayNameForBanner(this.props.userDisplayName)}!</p>
              </div>
            </>
            : ''}

          {this.state.BannerItemsState &&
            this.state.BannerItemsState.map((nhItem, index) => (
              <>
                <div className={`${styles.User_dec}`}>
                  <img width="100px" height="100px" src={nhItem.ImageURL} />
                  <p>{nhItem.Title}
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
                  </p>
                  {this.props.category === 'HomePage' &&
                    <div className={styles.Main_Banner_Icons} /*style={{ position: 'absolute', right: '10px', width: '8%', bottom: '8%' }}*/>
                      <a style={{ marginRight: '15px' }} onClick={() => { this._onFeedbackIconClick() }} href='javascript:void(0);'>
                        <svg width="21" height="24" viewBox="0 0 21 23" fill="none" xmlns="http://www.w3.org/2000/svg">
                          <path d="M18.1617 6.1355H2.59353C1.71232 6.1355 1 6.82616 1 7.68059V17.9931C1 18.8475 1.71232 19.5406 2.59353 19.5406H12.6443L15.5474 21.8119C15.7506 21.9709 16.0541 21.8404 16.0663 21.5864L16.1618 19.5406H18.1617C19.0429 19.5406 19.7552 18.8475 19.7552 17.9931V7.68059C19.7552 6.82616 19.0429 6.1355 18.1617 6.1355ZM10.2258 16.8277L7.08284 17.4164L7.64584 14.3594L14.1448 7.96778L16.7248 10.4361L10.2258 16.8277Z" stroke="white" stroke-miterlimit="10" />
                          <path d="M5.32496 4.50178L4.3238 3.7945C4.24057 3.73517 4.12797 3.73517 4.04475 3.7945L3.04114 4.49703C2.79636 4.66792 2.4708 4.43532 2.56381 4.16001L2.94812 3.01839C2.97994 2.92346 2.94567 2.81903 2.86245 2.75969L1.86129 2.05241C1.61895 1.87915 1.74379 1.50653 2.04488 1.50653H3.28347C3.38628 1.5089 3.47685 1.44482 3.50867 1.34988L3.89298 0.208267C3.986 -0.0694225 4.39234 -0.0694225 4.48291 0.208267L4.86232 1.34988C4.89414 1.44482 4.98471 1.5089 5.08752 1.5089H6.32612C6.6272 1.51127 6.75204 1.88627 6.50726 2.05716L5.50365 2.75969C5.42042 2.81903 5.38615 2.92108 5.41798 3.01602L5.79739 4.15763C5.89041 4.43532 5.5624 4.66555 5.31761 4.49229L5.32496 4.50178Z" fill="white" />
                          <path d="M11.408 4.53498L10.4068 3.8277C10.3236 3.76837 10.211 3.76837 10.1278 3.8277L9.12415 4.53024C8.87936 4.70112 8.5538 4.46853 8.64682 4.19321L9.03113 3.0516C9.06295 2.95666 9.02868 2.85223 8.94546 2.79289L7.9443 2.08562C7.70196 1.91236 7.8268 1.53973 8.12788 1.53973H9.36648C9.46929 1.5421 9.55986 1.47802 9.59168 1.38308L9.97599 0.241471C10.069 -0.0362193 10.4753 -0.0362193 10.5659 0.241471L10.9453 1.38308C10.9771 1.47802 11.0677 1.5421 11.1705 1.5421H12.4091C12.7102 1.54448 12.835 1.91948 12.5903 2.09036L11.5867 2.79289C11.5034 2.85223 11.4692 2.95429 11.501 3.04922L11.8804 4.19084C11.9734 4.46853 11.6454 4.69875 11.4006 4.52549L11.408 4.53498Z" fill="white" />
                          <path d="M17.4541 4.56355L16.4529 3.85627C16.3697 3.79693 16.2571 3.79693 16.1739 3.85627L15.1703 4.5588C14.9255 4.72969 14.5999 4.49709 14.693 4.22178L15.0773 3.08016C15.1091 2.98523 15.0748 2.88079 14.9916 2.82146L13.9904 2.11418C13.7481 1.94092 13.8729 1.56829 14.174 1.56829H15.4126C15.5154 1.57067 15.606 1.50659 15.6378 1.41165L16.0221 0.270035C16.1151 -0.00765489 16.5215 -0.00765489 16.6121 0.270035L16.9915 1.41165C17.0233 1.50659 17.1139 1.57067 17.2167 1.57067H18.4553C18.7563 1.57304 18.8812 1.94804 18.6364 2.11893L17.6328 2.82146C17.5496 2.88079 17.5153 2.98285 17.5471 3.07779L17.9265 4.2194C18.0196 4.49709 17.6915 4.72731 17.4468 4.55405L17.4541 4.56355Z" fill="white" />
                        </svg>
                      </a>
                      <a style={{ marginRight: '15px' }} onClick={() => { window.open(this.state.siteUrl + '/eng/resource%20hub/UserManualDocument', '_blank') }} href='javascript:void(0);'>
                        <svg width="27" height="23" viewBox="0 0 27 22" fill="none" xmlns="http://www.w3.org/2000/svg">
                          <path d="M8.1 0C4.806 0 2.53463 1.12468 2.48063 1.18109C2.26463 1.29391 2.16 1.46667 2.16 1.69231V3.38462H1.62C0.702 3.38462 0 4.11795 0 5.07692V19.7436C0 20.7026 0.702 21.4359 1.62 21.4359H12.3019C12.5949 21.7766 13.01 22 13.5 22C13.9901 22 14.4051 21.7766 14.6981 21.4359H25.38C26.298 21.4359 27 20.7026 27 19.7436V5.07692C27 4.11795 26.298 3.38462 25.38 3.38462H24.84V1.69231C24.84 1.46667 24.7354 1.29391 24.5194 1.18109C24.4654 1.12468 22.194 0 18.9 0C16.146 0 14.148 0.793269 13.5 1.07532C12.852 0.793269 10.854 0 8.1 0ZM8.1 1.12821C10.476 1.12821 12.258 1.74167 12.96 2.08013V18.3862C12.042 18.0478 10.314 17.5401 8.1 17.5401C5.886 17.5401 4.158 18.0478 3.24 18.3862V2.08013C3.942 1.74167 5.724 1.12821 8.1 1.12821ZM18.9 1.12821C21.276 1.12821 23.058 1.74167 23.76 2.08013V18.3862C22.842 18.0478 21.114 17.5401 18.9 17.5401C16.686 17.5401 14.958 18.0478 14.04 18.3862V2.08013C14.742 1.74167 16.524 1.12821 18.9 1.12821ZM18.9 3.94872C18.3035 3.94872 17.82 4.45383 17.82 5.07692C17.82 5.70001 18.3035 6.20513 18.9 6.20513C19.4965 6.20513 19.98 5.70001 19.98 5.07692C19.98 4.45383 19.4965 3.94872 18.9 3.94872ZM17.1788 7.89744C17.0355 7.91146 16.9035 7.98435 16.8117 8.10005C16.72 8.21576 16.6759 8.36481 16.6894 8.51442C16.7028 8.66403 16.7726 8.80194 16.8833 8.89781C16.9941 8.99369 17.1368 9.03967 17.28 9.02564H17.82V14.6667H19.98V8.46154C19.98 8.12308 19.764 7.89744 19.44 7.89744H18.9844H18.9H17.82H17.28C17.2631 7.89661 17.2462 7.89661 17.2294 7.89744C17.2125 7.89661 17.1956 7.89661 17.1788 7.89744Z" fill="white" />
                        </svg>
                      </a>
                      <TooltipHost
                        content={
                          <span dangerouslySetInnerHTML={{ __html: this.state.authorGuideToolTipBanner }} />
                        }
                      >
                        <svg width="24.42" height="24" viewBox="0 0 25 25" fill="none" xmlns="http://www.w3.org/2000/svg">
                          <path fill-rule="evenodd" clip-rule="evenodd" d="M15.8742 9.77529C15.8742 9.77529 16.2245 7.9051 15.0319 6.04276C14.2309 4.79334 11.923 3.63036 12.7241 4.75405C13.2341 5.46651 14.4678 6.06371 14.4678 6.06371C14.4678 6.06371 10.0453 8.2718 8.37367 4.07564C8.37367 4.07564 7.47732 3.29509 6.77672 5.71796C6.06324 8.18013 6.5243 9.68362 6.50112 9.77529L5.92673 8.85591C5.91643 8.68566 4.98917 7.53577 5.39356 4.96098C5.67689 3.15888 6.5449 1.65539 8.09549 1.28344C8.09549 1.28344 8.64412 0.0602194 10.1586 0.00783292C13.8394 -0.123133 16.5928 1.84136 16.7345 4.42139C16.8916 7.32099 15.8742 9.77792 15.8742 9.77792V9.77529Z" fill="white"></path>
                          <path d="M7 11C7.42037 13.4705 9.12568 15 11.0867 15C13.0477 15 14.6165 13.4705 15 11" stroke="white" stroke-width="0.5" stroke-miterlimit="10" stroke-linecap="round"></path>
                          <path d="M15.7763 8.50749H15.4647C15.2432 7.6693 14.4911 7.04852 13.5999 7.04852C12.7087 7.04852 12.0055 7.62739 11.7582 8.42629C11.7247 8.40795 11.6912 8.38962 11.6552 8.3739C11.4775 8.29008 11.2869 8.24032 11.0885 8.22722C11.0834 8.22722 11.0756 8.22722 11.0705 8.22722C11.0653 8.22722 11.0602 8.22722 11.055 8.22722C11.055 8.22722 11.0499 8.22722 11.0473 8.22722C11.0447 8.22722 11.0422 8.22722 11.0396 8.22722C11.0344 8.22722 11.0293 8.22722 11.0241 8.22722C11.019 8.22722 11.0113 8.22722 11.0061 8.22722C10.8103 8.24032 10.6197 8.29008 10.442 8.3739C10.4085 8.38962 10.3725 8.40795 10.339 8.42629C10.0917 7.63001 9.35764 7.04852 8.49734 7.04852C7.58296 7.04852 6.81539 7.70074 6.61706 8.57035H6.31313C6.19207 8.57035 6.09161 8.66989 6.09161 8.79561C6.09161 8.92134 6.18949 9.02087 6.31313 9.02087H6.56297C6.56297 10.1079 7.43099 10.9932 8.49734 10.9932C9.56369 10.9932 10.4317 10.1079 10.4317 9.02087C10.4317 8.98158 10.4291 8.94491 10.4266 8.90562C10.491 8.85586 10.5579 8.81395 10.63 8.77728C10.7614 8.71441 10.9031 8.68036 11.0499 8.66988C11.1967 8.67774 11.3384 8.71441 11.4697 8.77728C11.5419 8.81133 11.6088 8.85586 11.6732 8.90562C11.6732 8.94229 11.6681 8.98158 11.6681 9.02087C11.6681 10.1079 12.5361 10.9932 13.6024 10.9932C14.6688 10.9932 15.5368 10.1079 15.5368 9.02087C15.5368 8.99992 15.5368 8.97897 15.5342 8.95801H15.7815C15.9026 8.95801 16.0004 8.85848 16.0004 8.73275C16.0004 8.60702 15.9026 8.50749 15.7815 8.50749H15.7763ZM8.49734 10.5427C7.67311 10.5427 7.00342 9.85906 7.00342 9.02087C7.00342 8.18269 7.67311 7.49905 8.49734 7.49905C9.25976 7.49905 9.88824 8.08316 9.97839 8.83752C9.98611 8.89777 9.99126 8.96063 9.99126 9.02349C9.99126 9.8643 9.32158 10.5453 8.49734 10.5453V10.5427ZM13.5999 10.5427C12.7756 10.5427 12.1059 9.85906 12.1059 9.02087C12.1059 8.95801 12.1111 8.89515 12.1188 8.8349C12.209 8.08316 12.8374 7.49643 13.5999 7.49643C14.4241 7.49643 15.0938 8.18007 15.0938 9.01826C15.0938 9.85644 14.4241 10.5401 13.5999 10.5401V10.5427Z" fill="white"></path>
                          <path fill-rule="evenodd" clip-rule="evenodd" d="M19.2484 15.5797L15.0809 20.4726H0C0 20.4726 1.12044 15.5745 4.0748 15.1842C7.37173 14.7494 7.9796 14.5634 7.9796 14.5634C8.36596 15.9648 9.17216 18.3588 11.0215 18.4845C12.8709 18.3588 13.6771 15.9648 14.0635 14.5634C14.0635 14.5634 14.6688 14.7494 17.9683 15.1842C18.4422 15.2471 18.8698 15.3806 19.251 15.5797H19.2484Z" fill="white"></path>
                          <path d="M22.9033 14.6588L23.6537 15.3044C24.0759 15.6676 24.1288 16.3126 23.7716 16.742L17.8367 23.8763L15.5547 21.9131L21.4896 14.7788C21.8468 14.3494 22.4811 14.2956 22.9033 14.6588Z" fill="white"></path>
                          <path fill-rule="evenodd" clip-rule="evenodd" d="M15.1896 22.3525L17.4722 24.3146L15.4903 24.8482C15.2279 24.918 14.9756 24.7025 15.0019 24.4275L15.1916 22.3504L15.1896 22.3525Z" fill="white"></path>
                        </svg>
                      </TooltipHost>
                    </div>
                  }
                </div>
                {/* {nhItem.Title === "Project Engineering Teams" ?
                  <div className={`${styles.bannerTextOverlay}`}>
                    <p>{nhItem.HoverDescription}</p>
                  </div> : ''} */}
              </>
            ))
          }
        </div>
        <div>
          <Dialog
            hidden={!this.state._showFeedbackDialogEvents}
            onDismiss={this._onDismissEvent}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: 'Submit Feedback',
            }}
            modalProps={{
              isBlocking: false,
              styles: { main: { minWidth: '520px !important' } }
            }}
          >
            <div style={{ paddingTop: '20px' }}>
              <TextField multiline rows={7} placeholder='Please type your feedback.' value={this.state.FeedbackTitle} onChange={(e, newValue) => this.setState({ FeedbackTitle: (newValue || '') })} style={{ border: '1px solid rgb(96, 94, 92) !important' }} />
            </div>
            <DialogFooter>
              <PrimaryButton onClick={() => this.submitFeedback()} text="Submit" style={{ background: '#009bda', color: '#fff', border: 0 }} />
              <DefaultButton onClick={() => { this._onDismissEvent() }} text="Cancel" />
            </DialogFooter>
          </Dialog>
        </div>
      </section >
    );
  }
}

