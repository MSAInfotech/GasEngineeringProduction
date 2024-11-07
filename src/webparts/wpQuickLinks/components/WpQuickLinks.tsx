import * as React from 'react';
import styles from './WpQuickLinks.module.scss';
import type { IWpQuickLinksProps } from './IWpQuickLinksProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPFI, spfi, SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
// import "@pnp/sp/site-users/web";
// import "@pnp/sp/sitememberships";
import "bootstrap/dist/css/bootstrap.min.css"
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPen, faTrashAlt } from '@fortawesome/free-solid-svg-icons';
import { PrimaryButton, TextField, Label, Dialog, DialogType, DialogFooter, DefaultButton } from '@fluentui/react';

//import { nth } from 'lodash';


export interface IWpQuickLinksState {
  QuickLinksItemsState: any[];
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
  isItemInEditMode: boolean
}

//const [isModalOpen, setIsModalOpen] = React.useState<boolean>(false);
//const [imageUrl, setImageUrl] = React.useState<string>('');
//const [ setImageUrl] = React.useState<string>('');
// const [linkUrl, setLinkUrl] = React.useState<string>('');
// const [imageFile, setImageFile] = React.useState<File | null>(null);

// const openModal = () => this.setIsModalOpen(true);
// const closeModal = () => this.setIsModalOpen(false);


export default class WpQuickLinks extends React.Component<IWpQuickLinksProps, IWpQuickLinksState> {

  private _sp: SPFI;
  //private isUserInGroup: boolean = false;

  constructor(props: any) {
    super(props);
    sp: this._sp,
      //this._sp = spfi("https://sempra.sharepoint.com/sites/gasopscon/eng/").using(SPFx(this.props.context));
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      QuickLinksItemsState: [],
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
      isItemInEditMode: false
    }
  }


  public componentDidMount(): void {
    this.checkUserInGroup("PortalAdmins");

    this.fetchData();
  }

  private async checkUserInGroup(groupName: string): Promise<void> {
    const currentUser = await this._sp.web.currentUser();
    const groups = await this._sp.web.siteGroups.getByName(groupName).users();
    const isUserInGroupExists = groups.some(user => user.Id === currentUser.Id);
    this.setState({ isUserInGroup: isUserInGroupExists });
  }

  protected async fetchData() {
    try {
      await this._getQuickLinksItems();

    } catch (error) {
      // Handle error
    }
  }

  private async handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    if (event.target.files && event.target.files.length > 0) {
      debugger;
      this.setState({ setImageFile: event.target.files[0] });
      this.setState({
        file: event.target.files[0],
        filename: event.target.files[0].name,
      });
    }
  };


  private async _getQuickLinksItems(): Promise<string> {
    let _QuickLinksItems: any[] = [];
    //let _QuickLinksArrItems = [];
    this.setState({ QuickLinksItemsState: [] });
    let categoryfilter = this.props.category;
    debugger;
    let tmpQuickLinks = await this._sp.web.lists.getByTitle("QuickLinks").items.select("ID", "Title", "SortOrder", "RedirectURL", "Category", "Icon", "ImageURL").filter("Category eq '" + categoryfilter + "'").orderBy("SortOrder", true)();
    if (tmpQuickLinks && tmpQuickLinks.length > 0) {
      tmpQuickLinks.forEach(element => {
        //const linkTitle = element.Title.length > 25 ? element.Title.substring(0, 25) + '...' : element.Title;
        _QuickLinksItems.push(
          {
            "ID": element.ID,
            "Title": element.Title,
            "RedirectURL": element.RedirectURL,
            "ImageURL": element.ImageURL,
            "SortOrder": element.SortOrder
          });

      })

    } else {
      _QuickLinksItems = [];
    }
    this.setState({ QuickLinksItemsState: _QuickLinksItems })
    console.log(this.state.QuickLinksItemsState);
    return Promise.resolve("")
  }

  private saveItem() {
    if (this.state.saveButtonAdd) {
      this.uploadImageAndAddItem();
    }
    else {
      this.uploadImageAndUpdateItem();
    }
  }

  private async uploadImageAndAddItem() {
    const { Title, setLinkUrl } = this.state;
    if (!Title || Title.trim() == '') {
      alert('Link Title cannot be empty.');
      return;
    }
    if (!setLinkUrl || setLinkUrl.trim() == '') {
      alert('Link Url cannot be empty.');
      return;
    }

    if (this.state.setImageFile) {
      const folderServerRelativeUrl = "/sites/gasopscon/eng/resource%20hub/ImagesLibrary";
      const fileAddResult = await this._sp.web.getFolderByServerRelativePath(`/${decodeURI(folderServerRelativeUrl)}`).files.addUsingPath(this.state.setImageFile.name, this.state.setImageFile, { Overwrite: true });
      const imageUrl = fileAddResult.data.ServerRelativeUrl;
      //const setImageUrl = imageUrl;

      const i = await this._sp.web.lists.getByTitle("QuickLinks").items.add({
        Title: this.state.Title,
        ImageURL: imageUrl,
        RedirectURL: this.state.setLinkUrl,
        SortOrder: parseInt(this.state.sortorder),
        Category: this.props.category
      });

      if (i) {
        console.log("successfully created");
      }

      alert('Quick link added successfully!');
      this._getQuickLinksItems();
      this._onDismissEvent();
    } else {
      alert('Please select an image file to upload.');
    }
  };

  private async uploadImageAndUpdateItem() {
    const { Title, setLinkUrl } = this.state;
    if (!Title || Title.trim() == '') {
      alert('Link Title cannot be empty.');
      return;
    }
    if (!setLinkUrl || setLinkUrl.trim() == '') {
      alert('Link Url cannot be empty.');
      return;
    }

    if (this.state.setImageFile || this.state.showEditImageURL) {
      let fileAddResult: any = null;
      let imageUrl: string = '';
      if (this.state.setImageFile) {
        const folderServerRelativeUrl = "/sites/gasopscon/eng/resource%20hub/ImagesLibrary";
        fileAddResult = await this._sp.web.getFolderByServerRelativePath(`/${decodeURI(folderServerRelativeUrl)}`).files.addUsingPath(this.state.setImageFile.name, this.state.setImageFile, { Overwrite: true });
        imageUrl = fileAddResult.data.ServerRelativeUrl;
      } else {
        imageUrl = this.state.EditImageURL;
      }

      let updateID = this.state._EditItem.ID;
      const i = await this._sp.web.lists.getByTitle("QuickLinks").items.getById(updateID).update({
        Title: this.state.Title,
        ImageURL: imageUrl,
        RedirectURL: this.state.setLinkUrl,
        SortOrder: parseInt(this.state.sortorder),
      });

      if (i) {
        console.log("Quick link updated successfully.");
      }

      // const list = this._sp.web.lists.getByTitle("QuickLinks"); // Change this to your list title
      // await list.items.add({


      // });

      alert('Quick link updated successfully!');
      this._getQuickLinksItems();
      this._onDismissEvent();
    } else {
      alert('Please select an image file to upload.');
    }
  };

  private _onDismissEvent(): void {
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
    this.setState({
      _showDialogEvents: !this.state._showDialogEvents,
      showEditImageURL: tmpEditItem.ImageURL ? true : false,
      Title: tmpEditItem.Title,
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

    let itemId = this.state._deleteItem.ID;
    try {

      await this._sp.web.lists.getByTitle("QuickLinks").items.getById(itemId).delete();
      alert(`Quick link deleted successfully.`);
      this._getQuickLinksItems();
      this.setState({ _showDialogDelete: !this.state._showDialogDelete });
    } catch (error) {
      console.error(`Error deleting item with ID ${itemId}: `, error);
    }

    //this.deleteListItem();
  }


  public render(): React.ReactElement<IWpQuickLinksProps> {
    const {
      //description,            
      hasTeamsContext
    } = this.props;

    return (
      <section className={`${styles.wpQuickLinks} ${hasTeamsContext ? styles.teams : ''} col-md-12 col-sm-12 col-lg-12`}>
        <div className={`${styles.Internal_content}`}>
          <div className={`${styles.News_quicklinks_section}`}>
            <div className={`${styles.quick_section}`}>
              <div className={`${styles.News_title_section}`}>
                <div className={`${this.props.category === 'HomePage' ? styles.News_title : styles.News_title_h} col-md-9 col-sm-9 col-lg-9`}>
                  <h3>Quick Links<img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/info.png"} width="18px" /></h3>
                </div>
                {this.state.isUserInGroup && (
                  <div className={`${styles.News_oprt} col-md-2 col-sm-2 col-lg-2`}>
                    <ul>
                      <li><a onClick={() => { this._onDismissEvent() }} href='javascript:void(0);'><img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/plus-symbol.png"} width="20px" /></a></li>
                    </ul>
                  </div>
                )}

              </div>
              <div className={`${styles.Quick_list} col-md-12 col-sm-12 col-lg-12`}>
                <div className={`${styles.Q_list}`}>
                  <div className={`${styles.Edit_del}`}>
                    <ul style={{ padding: 0 }}>
                      {this.state.QuickLinksItemsState &&
                        this.state.QuickLinksItemsState.map((nhItem, index) => (
                          <li className={`${styles.Link_List}`}><a href='javascript:void(0);' onClick={() => {
                            let url = nhItem.RedirectURL.startsWith('http') ? nhItem.RedirectURL : 'https://' + nhItem.RedirectURL;
                            window.open(url);
                          }} title={nhItem.Title}><img src={decodeURI(nhItem.ImageURL)} width="20px" /><span className={`${styles.Link_Title}`}>{nhItem.Title}</span></a>
                            {this.state.isUserInGroup && (
                              <div style={{ width: "10%" }}>
                                <a onClick={() => { this._onEditItem(nhItem) }} href='javascript:void(0);' style={{ marginRight: '15%' }}><FontAwesomeIcon icon={faPen} /></a>
                                <a onClick={() => { this._onDeleteItem(nhItem) }} href='javascript:void(0);'><FontAwesomeIcon icon={faTrashAlt} /></a>
                              </div>
                            )}
                          </li>
                        ))
                      }
                    </ul>
                  </div>
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
              title: this.state.isItemInEditMode == true ? 'Edit Quick Link' : 'Add New Quick Link',
              subText: 'Provide the below information.'
            }}
            modalProps={{
              isBlocking: false,
              styles: { main: { minWidth: '520px !important' } }
            }}
          >
            <TextField label="Quick Link Title" value={this.state.Title} onChange={(e, newValue) => this.setState({ Title: (newValue || '') })} />
            <TextField label="Link URL" value={this.state.setLinkUrl} onChange={(e, newValue) => this.setState({ setLinkUrl: (newValue || '') })} />
            <TextField label="Sort Order" value={this.state.sortorder} onChange={(e, newValue) => this.setState({ sortorder: (newValue || '') })} />

            {this.state.showEditImageURL && (
              <img src={this.state.EditImageURL} width="30" height="30" style={{ padding: '5px', color: '#009bda', background: '#ffff' }}></img>
            )}

            <Label>Icon Upload</Label>
            <input type="file" accept="image/*" onChange={(event) => { this.handleFileChange(event) }} />
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

            <Label>Are you sure want to delete this quick link ?</Label>

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
