import * as React from 'react';
import styles from './WpNewsMain.module.scss';
import type { IWpNewsMainProps } from './IWpNewsMainProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPFI, spfi, SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
import { IconButton, IButtonStyles } from '@fluentui/react/lib/Button';
import { TooltipHost, ITooltipHostStyles } from '@fluentui/react/lib/Tooltip';
import "bootstrap/dist/css/bootstrap.min.css"
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faEye, faPen, faTrashAlt } from '@fortawesome/free-solid-svg-icons';
// import { faPenToSquare } from '@fortawesome/free-regular-svg-icons';

const tooltipStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
const iconButtonStyles: Partial<IButtonStyles> = {
  root: {
    fontSize: '22.4px', // Increases the size of the icon
    color: 'gray', // Sets the default icon color
    verticalAlign: 'middle'
  },
  rootHovered: {
    color: 'rgb(181, 73, 15)', // Changes the icon color on hover
    backgroundColor: 'rgb(240, 240, 240)', // Sets the background color on hover
  },
  icon: {
    fontSize: '22.4px', // Increases the size of the icon
    color: 'gray',
    verticalAlign: 'middle',
  },

};
import { PrimaryButton, TextField, Label, Dialog, DialogType, DialogFooter, DefaultButton } from '@fluentui/react';
export interface IWpNewsMainState {
  NewsItemsfirstRowState: any[];
  NewsItemsSecondRowState: any[];
  _showDialogEvents: boolean,
  _showDialogDelete: boolean,
  _showDialogDeleteConfirm: boolean,
  _showDialogDeleteConfirmYes: boolean,
  _showDialogDeleteConfirmNo: boolean,
  _deleteItem: any,
  linkUrl: string,
  setLinkUrl: string,
  imageFile: any,
  setImageFile: any,
  setNewsFile: any,
  file: any,
  filename: string,
  sortorder: string,
  Title: string,
  EditImageURL: string,
  showEditImageURL: boolean,
  saveButtonAdd: boolean,
  savButtonEdit: boolean,
  _EditItem: any;
  isUserInGroup: boolean;
  toolTipNews: string;
  newsviewFirst: string;
  newsviewSecond: string;
  newsviewThird: string;
  newsviewFourth: string;
  newsAllItems: any;
  isItemInEditMode: boolean,
  siteUrl: string,
  subSiteUrl: string
}

export default class WpNewsMain extends React.Component<IWpNewsMainProps, IWpNewsMainState> {

  private _sp: SPFI;

  constructor(props: any) {
    super(props);
    debugger;
    sp: this._sp,
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      NewsItemsfirstRowState: [],
      NewsItemsSecondRowState: [],
      _showDialogEvents: false,
      _showDialogDelete: false,
      _showDialogDeleteConfirm: false,
      _showDialogDeleteConfirmYes: false,
      _showDialogDeleteConfirmNo: false,
      linkUrl: "",
      setLinkUrl: "",
      imageFile: null,
      setImageFile: null,
      file: null,
      filename: "",
      sortorder: "",
      Title: "",
      EditImageURL: "",
      showEditImageURL: false,
      saveButtonAdd: false,
      savButtonEdit: false,
      _EditItem: {},
      _deleteItem: {},
      isUserInGroup: false,
      toolTipNews: "",
      newsviewFirst: "",
      newsviewSecond: "",
      newsviewThird: "",
      newsviewFourth: "",
      newsAllItems: [],
      setNewsFile: null,
      isItemInEditMode: false,
      siteUrl: "",
      subSiteUrl: "/sites/gasopscon/eng/resource%20hub"
    }
  }

  public componentDidMount(): void {
    this.checkUserInGroup("PortalAdmins");
    debugger;
    this.fetchData();
    const title: any = document.getElementById('leftRegion');
    console.log(title);
    if (title) {
      console.log('body onload - ' + title);
      if (title.className.indexOf('col-') <= -1) {
        title.className += " col-md-9 col-lg-9";
      }
    }
    const headerText: any = document.querySelector('#O365_AppName > span');
    console.log(headerText);
    if (headerText) {
      console.log('body onload - ' + headerText);
      headerText.innerHTML = "Gas Engineering - Project Engineering Resource Hub";
    }
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
      await this._getItems();
      await this._getToolTip();
    } catch (error) {
      // Handle error
    }
  }

  private async _getToolTip(): Promise<string> {
    let _NewsTooTip = "";

    debugger;
    let tmpNewsToolTip = await this._sp.web.lists.getByTitle("ToolTip").items.select("ID", "Title", "Category").filter("Category eq 'News'").orderBy("ID", false).top(1)();
    if (tmpNewsToolTip && tmpNewsToolTip.length > 0) {
      _NewsTooTip = tmpNewsToolTip[0].Title;
    } else {
      _NewsTooTip = "";
    }
    this.setState({
      toolTipNews: _NewsTooTip
    })

    return Promise.resolve("")
  }

  private async _getItems(): Promise<string> {
    let _NewsMainItems: any[] = [];
    let tmpNewsMain = await this._sp.web.lists.getByTitle("News").items.select("ID", "Title", "ImageURL", "RedirectURL").orderBy("ID", false).top(4)();
    const currentUserId = await this._sp.web.currentUser();
    const userViews = await this._sp.web.lists.getByTitle('NewsViews').items.select("Id,Title,Author/Id").filter(`Author/Id eq '${currentUserId.Id}'`).expand("Author").orderBy("ID", false)();
    if (tmpNewsMain && tmpNewsMain.length > 0) {
      //tmpNewsMain.forEach(element => {
      for (var i = 0; i < tmpNewsMain.length; i++) {
        const newsCount = await this.fetchNewsViewsCount(tmpNewsMain[i].Id);
        let _currentUserClicked = false;
        if (userViews && userViews.length > 0) {
          userViews.forEach(userelement => {
            if (tmpNewsMain[i].Id == userelement.Id) {
              _currentUserClicked = true;
            }
          });
        }
        _NewsMainItems.push({
          "NewsID": tmpNewsMain[i].ID,
          "NewsTitle": tmpNewsMain[i].Title,
          "NewsViews": newsCount,
          "currentUserClicked": _currentUserClicked,
          "ImageURL": tmpNewsMain[i].ImageURL,
          "RedirectURL": tmpNewsMain[i].RedirectURL
        });
      }

    } else {
      _NewsMainItems = [];
    }
    this.setState({
      newsAllItems: _NewsMainItems
    });
    if (_NewsMainItems.length >= 4) {
      this.setState({ NewsItemsfirstRowState: _NewsMainItems.slice(0, 2) });
      this.setState({ NewsItemsSecondRowState: _NewsMainItems.slice(2, 4) });
      //this.state.NewsItemsfirstRowState = _NewsMainItems.slice(0, 2);
    }
    else if (_NewsMainItems.length < 4 && _NewsMainItems.length > 0) {
      this.setState({ NewsItemsfirstRowState: _NewsMainItems.slice(0, 2) });
      this.setState({ NewsItemsSecondRowState: _NewsMainItems.slice(2, 4) });
    }
    else {
      this.setState({ NewsItemsfirstRowState: [] });
      this.setState({ NewsItemsSecondRowState: [] });
    }
    //this.setState({ NewsItemsfirstRowState: _NewsMainItems })
    //console.log(this.state.NewsItemsfirstRowState);
    return Promise.resolve("");
  }

  private async fetchNewsViewsCount(itemid: any): Promise<number> {
    try {
      const items = await this._sp.web.lists.getByTitle('NewsViews').items.select("ID,Title,NewsLookup/Id").filter(`NewsLookup/Id eq '${itemid}'`).expand("NewsLookup")();
      return items.length;
    } catch (error) {
      console.error("Error fetching data:", error);
      //Promise.resolve("");
      return 0;
    }

  }


  // private async getCurrentUser(): Promise<any>  {
  //   try {
  //     const currentUser = await this._sp.web.currentUser();
  //     return currentUser.Id;
  //   } catch (error) {
  //     console.error('Error fetching current user:', error);
  //     throw error;
  //   }
  //   return Promise.resolve("");
  // }

  // private async _getViewsCount(queryString:string): Promise<string> {
  //   try{
  //     let _NewsMainItems = [];
  //     const currentUser = await this.getCurrentUser();
  //     let tmpNewsViewsMain = await this._sp.web.lists.getByTitle("NewsViews").items.select("ID", "Title","NewsReader/","NewsLookup").filter(queryString).expand('NewsLookup','NewsReader').orderBy("ID", false)();
  //     const filteredItems = tmpNewsViewsMain.filter(item => item.AssignedTo && item.AssignedTo.Id === currentUser);
  //     if (tmpNewsViewsMain && tmpNewsViewsMain.length > 0) {
  //       _NewsMainItems = tmpNewsViewsMain;
  //       let _queryString = "";
  //       let cnt = 1;
  //       let _newsArray = this.state.newsAllItems;
  //       // _newsArray.forEach(element => {
  //       //   tmpNewsViewsMain.forEach(elementViews => {



  //       //   });


  //       // });
  //       this._getViewsCount(_queryString);
  //     } else {
  //       _NewsMainItems = [];
  //     }
  //   }
  //   catch(e){

  //   }
  //   return Promise.resolve("")

  // }


  private async handleNewsFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    if (event.target.files && event.target.files.length > 0) {
      this.setState({ setNewsFile: event.target.files[0] });
      this.setState({
        file: event.target.files[0],
        filename: event.target.files[0].name,
      });
    }
  };

  private async handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    if (event.target.files && event.target.files.length > 0) {
      this.setState({ setImageFile: event.target.files[0] });
      this.setState({
        file: event.target.files[0],
        filename: event.target.files[0].name,
      });
    }
  };

  private saveItem() {
    if (this.state.saveButtonAdd) {
      if (this.validateAddItem()) {
        this.uploadImageAndAddItem();
      } else {
        let message: string = "";
        if (!this.state.Title) {
          message = message + "News Title";
        }
        if (!this.state.setImageFile) {
          message = message.length > 0 ? message + ", Icon" : message + "Icon";
        }
        if (!this.state.setNewsFile && !this.state.setLinkUrl) {
          message = message.length > 0 ? message + ", News File/News Link" : message + "News File/News Link";
        }
        if (message.length > 0) {
          alert('Please provide ' + message + '.');
        }
      }
    }
    else {
      if (this.validateEditItem()) {
        this.uploadImageAndUpdateItem();
      } else {
        let message: string = "";
        if (!this.state.Title) {
          message = message + "News Title";
        }
        if (!this.state.setImageFile) {
          message = message.length > 0 ? message + ", Icon" : message + "Icon";
        }
        if (!this.state.setNewsFile && !this.state.setLinkUrl) {
          message = message.length > 0 ? message + ", News File/News Link" : message + "News File/News Link";
        }
        if (message.length > 0) {
          alert('Please provide ' + message + '.');
        }
      }
    }
  }

  private validateAddItem(): boolean {
    let isValid: boolean = (this.state.Title && this.state.setImageFile && (this.state.setNewsFile || this.state.setLinkUrl));
    return isValid;
  }

  private validateEditItem(): boolean {
    let isValid: boolean = (this.state.Title && (this.state.setImageFile || this.state.showEditImageURL) && (this.state.setNewsFile || this.state.setLinkUrl));
    return isValid;
  }

  private async uploadImageAndAddItem() {
    //if (this.state.setImageFile) {
    let isMaxNews: boolean = false;
    let tmpNewsMain = await this._sp.web.lists.getByTitle("News").items.select("ID")();
    if (tmpNewsMain.length >= 4) {
      isMaxNews = true;
    }

    if (isMaxNews) {
      alert('There are 4 News items already, please delete 1 to proceed !');
    }
    else {
      const folderServerRelativeUrl = decodeURI(this.state.subSiteUrl) + "/ImagesLibrary";
      const fileAddResult = await this._sp.web.getFolderByServerRelativePath(`/${folderServerRelativeUrl}`).files.addUsingPath(this.state.setImageFile.name, this.state.setImageFile, { Overwrite: true });
      const imageUrl = this.state.siteUrl.split('/sites/')[0] + fileAddResult.data.ServerRelativeUrl;

      let documentUrl: string = '';
      if (this.state.setNewsFile) {
        const libraryServerRelativeUrl = decodeURI(this.state.subSiteUrl) + "/NewsDocument";
        const uploadAndResult = await this._sp.web.getFolderByServerRelativePath(`/${libraryServerRelativeUrl}`).files.addUsingPath(this.state.setNewsFile.name, this.state.setNewsFile, { Overwrite: true });
        documentUrl = this.state.siteUrl.split('/sites/')[0] + uploadAndResult.data.ServerRelativeUrl;
      }

      const i = await this._sp.web.lists.getByTitle("News").items.add({
        Title: this.state.Title,
        ImageURL: imageUrl,
        RedirectURL: documentUrl ? documentUrl : this.state.setLinkUrl,
        //SortOrder: parseInt(this.state.sortorder),
        Category: this.props.category
      });

      if (i) {
        console.log("successfully created");
      }

      alert('News added successfully!');
      this._getItems();
      this._onDismissEvent();
    }
  };

  private async uploadImageAndUpdateItem() {
    //if (this.state.setImageFile || this.state.showEditImageURL) {
    debugger;
    let fileAddResult: any = null;
    let imageUrl: string = '';
    if (this.state.setImageFile) {
      const folderServerRelativeUrl = decodeURI(this.state.subSiteUrl) + "/ImagesLibrary";
      fileAddResult = await this._sp.web.getFolderByServerRelativePath(`/${folderServerRelativeUrl}`).files.addUsingPath(this.state.setImageFile.name, this.state.setImageFile, { Overwrite: true });
      imageUrl = this.state.siteUrl.split('/sites/')[0] + fileAddResult.data.ServerRelativeUrl;
    } else {
      imageUrl = this.state.EditImageURL;
    }

    let newsDocUrl: string = '';
    if (this.state.setNewsFile) {
      const folderServerRelativeUrl = decodeURI(this.state.subSiteUrl) + "/NewsDocument";
      fileAddResult = await this._sp.web.getFolderByServerRelativePath(`/${folderServerRelativeUrl}`).files.addUsingPath(this.state.setNewsFile.name, this.state.setNewsFile, { Overwrite: true });
      newsDocUrl = this.state.siteUrl.split('/sites/')[0] + fileAddResult.data.ServerRelativeUrl;
    } else {
      newsDocUrl = this.state.setLinkUrl;
    }


    debugger;
    let updateID = this.state._EditItem.NewsID;
    const i = await this._sp.web.lists.getByTitle("News").items.getById(updateID).update({
      Title: this.state.Title,
      ImageURL: imageUrl,
      RedirectURL: newsDocUrl,
      Category: this.props.category,
      //SortOrder: parseInt(this.state.sortorder),
    });

    if (i) {
      console.log("successfully Updated");
    }

    // const list = this._sp.web.lists.getByTitle("QuickLinks"); // Change this to your list title
    // await list.items.add({


    // });

    alert('News Updated successfully!');
    this._getItems();
    this._onDismissEvent();
    // } else {
    //   alert('Please select an image file to upload.');
    // }
  };

  private _onDismissEvent(): void {
    debugger;
    this.setState({
      _showDialogEvents: !this.state._showDialogEvents,
      showEditImageURL: false,
      saveButtonAdd: true,
      savButtonEdit: false,
      Title: "",
      setLinkUrl: "",
      sortorder: "",
      EditImageURL: "",
      isItemInEditMode: false,
      setImageFile: null
    });
    //return false;
  }

  private _onDismissDelete(): void {
    this.setState({ _showDialogDelete: !this.state._showDialogDelete });
  }
  private _onEditItem(tmpEditItem: any): void {
    debugger;
    this.setState({
      _showDialogEvents: !this.state._showDialogEvents,
      showEditImageURL: tmpEditItem.ImageURL ? true : false,
      Title: tmpEditItem.NewsTitle,
      setLinkUrl: tmpEditItem.RedirectURL,
      sortorder: tmpEditItem.SortOrder,
      EditImageURL: tmpEditItem.ImageURL,
      saveButtonAdd: false,
      savButtonEdit: true,
      isItemInEditMode: true,
      _EditItem: tmpEditItem

    });

  }

  private _onDeleteItem(tmpDeleteItem: any): void {
    debugger;
    this.setState({
      _showDialogDelete: !this.state._showDialogDelete,
      showEditImageURL: false,
      saveButtonAdd: false,
      savButtonEdit: false,
      _deleteItem: tmpDeleteItem
    });
    //return false;
  }

  private async deleteItem() {

    let itemId = this.state._deleteItem.NewsID;
    try {

      await this._sp.web.lists.getByTitle("News").items.getById(itemId).delete();
      alert(`News deleted successfully.`);
      this._getItems();
      this.setState({ _showDialogDelete: !this.state._showDialogDelete });
    } catch (error) {
      console.error(`Error deleting item with ID ${itemId}: `, error);
    }
  }

  private async onNewsClick(onClickNews: any, isFirstRowState: boolean) {
    const currentUserId = await this._sp.web.currentUser();
    debugger;
    const items = await this._sp.web.lists.getByTitle('NewsViews').items.select("ID,Title,NewsReader/Id,NewsReader/EMail,NewsLookup/Id").filter(`NewsLookup/Id eq '${onClickNews.NewsID}' and NewsReader/Id eq '${currentUserId.Id}'`).expand("NewsLookup,NewsReader")();
    if (!items || items?.length == 0) {
      debugger;
      await this._sp.web.lists.getByTitle("NewsViews").items.add({
        NewsReaderId: currentUserId.Id,
        NewsLookupId: onClickNews.NewsID
      });
      if (isFirstRowState) {
        let firstRowData: any[] = this.state.NewsItemsfirstRowState;
        firstRowData.forEach(element => {
          if (element.NewsID == onClickNews.NewsID) {
            element.NewsViews = element.NewsViews + 1;
          }
        });
        this.setState({
          NewsItemsfirstRowState: firstRowData
        });
      } else {
        let secondRowData: any[] = this.state.NewsItemsSecondRowState;
        secondRowData.forEach(element => {
          if (element.NewsID == onClickNews.NewsID) {
            element.NewsViews = element.NewsViews + 1;
          }
        });
        this.setState({
          NewsItemsSecondRowState: secondRowData
        });
      }
    }
    debugger;

    const url = onClickNews.RedirectURL.startsWith('http') ? onClickNews.RedirectURL : 'https://' + onClickNews.RedirectURL;
    window.open(url);
  }


  public render(): React.ReactElement<IWpNewsMainProps> {
    const {
      hasTeamsContext
    } = this.props;

    return (
      <section className={`${styles.wpNewsMain} ${hasTeamsContext ? styles.teams : ''} col-md-8 col-sm-8 col-lg-8`}>
        <div className={`${styles.Internal_content}`}>
          <div className={`${styles.News_quicklinks_section} col-md-12 col-sm-12 col-lg-12`}>
            <div className={`${styles.w_news} col-md-1 col-sm-1 col-lg-1`}></div>
            <div className={`${styles.News_section} col-md-10 col-sm-10 co-lg-10`}>
              <div className={`${styles.News_title_section} col-md-12 col-sm-12 col-lg-12`}>
                <div className={`${styles.News_title} col-md-10 col-sm-10 col-lg-10`}>
                  <h3>News
                    <TooltipHost
                      content={this.state.toolTipNews}
                      styles={tooltipStyles}
                    >
                      <IconButton
                        iconProps={{ iconName: 'Info' }}
                        title="Info"
                        ariaLabel="Info"
                        styles={iconButtonStyles}
                      />
                    </TooltipHost>
                  </h3>
                </div>

                {this.state.isUserInGroup && (
                  <div className={`${styles.News_oprt} col-md-1 col-sm-1 col-lg-1`}>
                    <ul className={`${styles.News_Add_Option}`}>
                      <li><a onClick={() => { this._onDismissEvent() }} href='javascript:void(0);'><img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/plus-symbol.png"} width="20px" /></a></li>
                    </ul>
                  </div>
                )}
              </div>
              <div className={`col-md-12 col-sm-12 col-lg-12`}>
                <div className={`${styles.News_tiles}`}>
                  {this.state.NewsItemsfirstRowState &&
                    this.state.NewsItemsfirstRowState.map((nhItem, index) => (
                      <div className={`${styles.News_tiles_in} mb-5`}>
                        <div style={{ float: 'inline-end', marginTop: '5px', marginRight: '5px' }}>
                          {this.state.isUserInGroup && (
                            <span>
                              <a onClick={() => { this._onEditItem(nhItem) }} href='javascript:void(0);' style={{ marginRight: '8px' }}><FontAwesomeIcon icon={faPen} style={{ color: 'white' }} /></a>
                              <a onClick={() => { this._onDeleteItem(nhItem) }} href='javascript:void(0);' style={{ marginRight: '5px' }}><FontAwesomeIcon icon={faTrashAlt} style={{ color: 'white' }} /></a>
                            </span>
                          )}
                        </div>
                        <div className={`${styles.Edit_del}`}>
                          <ul>
                            <li>
                              <a onClick={() => { this.onNewsClick(nhItem, true) }} href='javascript:void(0);'><img src={decodeURI(nhItem.ImageURL)} className={`${styles.Card_img_tile}`} />
                                <h2>{nhItem.NewsTitle}</h2></a>
                              <div>
                                <FontAwesomeIcon icon={faEye} style={{ color: '#FFFF', fontSize: '16px' }} />&nbsp;&nbsp;<label style={{ color: '#FFFF', fontFamily: 'Montserrat', fontStyle: 'normal', fontWeight: '400', fontSize: '16px', lineHeight: '20px' }}> {nhItem.NewsViews} views</label>
                              </div>
                            </li>
                          </ul>
                        </div>
                      </div>
                    ))
                  }
                </div>
                <div className={`${styles.News_tiles}`}>
                  {this.state.NewsItemsSecondRowState &&
                    this.state.NewsItemsSecondRowState.map((nhItem, index) => (
                      <div className={`${styles.News_tiles_in} mb-5`}>
                        <div style={{ float: 'inline-end', marginTop: '5px', marginRight: '5px' }}>
                          {this.state.isUserInGroup && (
                            <span>
                              <a onClick={() => { this._onEditItem(nhItem) }} href='javascript:void(0);' style={{ marginRight: '8px' }}><FontAwesomeIcon icon={faPen} style={{ color: 'white' }} /></a>
                              <a onClick={() => { this._onDeleteItem(nhItem) }} href='javascript:void(0);' style={{ marginRight: '5px' }}><FontAwesomeIcon icon={faTrashAlt} style={{ color: 'white' }} /></a>
                            </span>
                          )}
                        </div>
                        <div className={`${styles.Edit_del}`}>
                          <ul>
                            <li>
                              <a onClick={() => { this.onNewsClick(nhItem, false) }} href='javascript:void(0);'><img src={decodeURI(nhItem.ImageURL)} className={`${styles.Card_img_tile}`} />
                                <h2>{nhItem.NewsTitle}</h2></a>
                              <div>
                                <FontAwesomeIcon icon={faEye} style={{ color: '#FFFF', fontSize: '16px' }} />&nbsp;&nbsp;<label style={{ color: '#FFFF', fontFamily: 'Montserrat', fontStyle: 'normal', fontWeight: '400', fontSize: '16px', lineHeight: '20px' }}> {nhItem.NewsViews} views</label>
                              </div>
                            </li>
                          </ul>
                        </div>
                      </div>
                    ))
                  }
                </div>
              </div>
            </div>
          </div>
        </div>
        <div>
          <Dialog
            hidden={!this.state._showDialogEvents}
            onDismiss={this._onDismissEvent}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: this.state.isItemInEditMode == true ? 'Edit News' : 'Add News',
              subText: 'Provide the below information.'
            }}
            modalProps={{
              isBlocking: false,
              styles: { main: { minWidth: '520px !important' } }
            }}
          >
            <TextField label="News Title" value={this.state.Title} onChange={(e, newValue) => this.setState({ Title: (newValue || '') })} style={{ border: '1px solid rgb(96, 94, 92) !important' }} />
            <TextField label="Link URL" value={this.state.setLinkUrl} onChange={(e, newValue) => this.setState({ setLinkUrl: (newValue || '') })} style={{ border: '1px solid rgb(96, 94, 92) !important' }} />
            {/* <TextField label="Sort Order" value={this.state.sortorder} onChange={(e, newValue) => this.setState({ sortorder: (newValue || '') })} /> */}

            {this.state.showEditImageURL && (
              <img src={this.state.EditImageURL} width="30" height="30" style={{ padding: '5px', background: '#009bda' }}></img>
            )}

            <Label>Icon Upload</Label>
            <input type="file" accept="image/*" onChange={(event) => { this.handleFileChange(event) }} />
            <Label>File Upload</Label>
            <input type="file" onChange={(event) => { this.handleNewsFileChange(event) }} />
            <DialogFooter>
              <PrimaryButton onClick={() => this.saveItem()} text="Save" style={{ background: '#009bda', color: '#fff', border: 0 }} />
              {/* <PrimaryButton onClick={() => this.uploadImageAndAddItem()} text="Add Item" /> */}
              <DefaultButton onClick={() => { this._onDismissEvent() }} text="Cancel" />
            </DialogFooter>
          </Dialog>
          <Dialog
            hidden={!this.state._showDialogDelete}
            onDismiss={this._onDismissDelete}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: 'Delete Item',
              //subText: 'Upload an image and provide a link URL.'
            }}
            modalProps={{
              isBlocking: false,
              styles: { main: { maxWidth: 450 } }
            }}
          >

            <Label>Are you sure want to delete this news ?</Label>

            <DialogFooter>
              <PrimaryButton onClick={() => this.deleteItem()} text="Confirm" style={{ background: '#009bda', color: '#fff', border: 0 }} />
              {/* <PrimaryButton onClick={() => this.uploadImageAndAddItem()} text="Add Item" /> */}
              <DefaultButton onClick={() => { this._onDismissDelete() }} text="Cancel" />
            </DialogFooter>
          </Dialog>
        </div>
      </section>
    );
  }
}
