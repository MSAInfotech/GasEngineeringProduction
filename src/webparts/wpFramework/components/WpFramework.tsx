import * as React from 'react';
import styles from './WpFramework.module.scss';
import type { IWpFrameworkProps } from './IWpFrameworkProps';
// import { escape } from '@microsoft/sp-lodash-subset';
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


export interface IWpFrameworkState {
  FrameworkItemsState: any[];
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
  Title: string,
  EditImageURL: string,
  showEditImageURL: boolean,
  saveButtonAdd: boolean,
  savButtonEdit: boolean,
  _EditItem: any;
  isUserInGroup: boolean;
  isItemInEditMode: boolean
}

export default class WpFramework extends React.Component<IWpFrameworkProps, IWpFrameworkState> {

  private _sp: SPFI;
  //private isUserInGroup: boolean = false;

  constructor(props: any) {
    super(props);
    sp: this._sp,
      //this._sp = spfi("https://sempra.sharepoint.com/sites/gasopscon/eng/").using(SPFx(this.props.context));
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      FrameworkItemsState: [],
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
      await this._getFrameworkItems();

    } catch (error) {
      // Handle error
    }
  }

  private async handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    if (event.target.files && event.target.files.length > 0) {
      this.setState({ setImageFile: event.target.files[0] });
      this.setState({
        file: event.target.files[0],
        filename: event.target.files[0].name,
      });
    }
  };


  private async _getFrameworkItems(): Promise<string> {
    let _FrameworkItems: any[] = [];
    //let _FrameworkArrItems = [];
    this.setState({ FrameworkItemsState: [] });
    let categoryfilter = this.props.category;
    let tmpFramework = await this._sp.web.lists.getByTitle("PEFramework")
      .items.select("ID", "Title", "RedirectUrl", "Category", "ImageUrl")
      .filter("Category eq '" + categoryfilter + "'")
      .orderBy("ID", false)();

    if (tmpFramework && tmpFramework.length > 0) {
      tmpFramework.forEach(element => {
        // var ImageUrl = "";
        // if (element.Icon != null && element.Icon != ""){
        //   ImageUrl = JSON.parse(element.Icon).fileName;
        // }
        _FrameworkItems.push(
          {
            "ID": element.ID,
            "Title": element.Title,
            "RedirectUrl": element.RedirectUrl,
            "ImageUrl": element.ImageUrl,
            //"ImageUrl": "https://sempra.sharepoint.com/sites/gasopscon/eng/resource%20hub/Lists/PEFramework/Attachments/" + element.ID + "/" + ImageUrl,
          });

      })

    } else {
      _FrameworkItems = [];
    }
    this.setState({ FrameworkItemsState: _FrameworkItems })
    console.log(this.state.FrameworkItemsState);
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
      alert('Title cannot be empty.');
      return;
    }
    if (!setLinkUrl || setLinkUrl.trim() == '') {
      alert('Link URL cannot be empty.');
      return;
    }
    if (this.state.setImageFile) {
      //const { file, filename } = this.state;
      //console.log(`Generated URL: ${this.context.pageContext.web.serverRelativeUrl}/ImagesLibrary`);
      const folderServerRelativeUrl = "/sites/gasopscon/eng/resource%20hub/ImagesLibrary/PEFramework";
      //const folderServerRelativeUrl = `${this.context.web.serverRelativeUrl}/ImagesLibrary`;
      //const libraryName = 'ImagesLibra  ry'; // Change this to your library name
      //const fileAddResult = await this._sp.web.getFolderByServerRelativeUrl(`/${libraryName}`).files.add(imageFile.name, imageFile, true);

      //await this._sp.web.getFolderByServerRelativePath(folderServerRelativeUrl).files();
      // const folder = await this._sp.web.getFolderByServerRelativePath(folderServerRelativeUrl).files();
      // if (!folder) {
      //   throw new Error(`Folder '${folderServerRelativeUrl}' does not exist.`);
      // }
      //let result: IFileAddResult;
      const fileAddResult = await this._sp.web.getFolderByServerRelativePath(`/${decodeURI(folderServerRelativeUrl)}`).files.addUsingPath(this.state.setImageFile.name, this.state.setImageFile, { Overwrite: true });
 
      const imageUrl = fileAddResult.data.ServerRelativeUrl;
      //const setImageUrl = imageUrl;

      const i = await this._sp.web.lists.getByTitle("PEFramework").items.add({
        Title: this.state.Title,
        ImageUrl: imageUrl,
        RedirectUrl: this.state.setLinkUrl,
        Category: this.props.category
      });

      if (i) {
        console.log("successfully created");
      }

      // const list = this._sp.web.lists.getByTitle("PEFramework"); // Change this to your list title
      // await list.items.add({


      // });

      alert('Framework added successfully!');
      this._getFrameworkItems();
      this._onDismissEvent();
    } else {
      alert('Please select an image file to upload.');
    }
  };

  private async uploadImageAndUpdateItem() {
    const { Title, setLinkUrl } = this.state;
    if (!Title || Title.trim() == '') {
      alert('Title cannot be empty.');
      return;
    }
    if (!setLinkUrl || setLinkUrl.trim() == '') {
      alert('Link URL cannot be empty.');
      return;
    }

    if (this.state.setImageFile || this.state.showEditImageURL) {
      let fileAddResult: any = null;
      let imageUrl: string = '';
      if (this.state.setImageFile) {
        const folderServerRelativeUrl = "/sites/gasopscon/eng/resource%20hub/ImagesLibrary/PEFramework";
        fileAddResult = await this._sp.web.getFolderByServerRelativePath(`/${decodeURI(folderServerRelativeUrl)}`).files.addUsingPath(this.state.setImageFile.name, this.state.setImageFile, { Overwrite: true });
        imageUrl = fileAddResult.data.ServerRelativeUrl;
      } else {
        imageUrl = this.state.EditImageURL;
      }

      let updateID = this.state._EditItem.ID;
      const i = await this._sp.web.lists.getByTitle("PEFramework").items.getById(updateID).update({
        Title: this.state.Title,
        ImageUrl: imageUrl,
        RedirectUrl: this.state.setLinkUrl,
      });

      if (i) {
        console.log("Framework updated successfully.");
      }

      // const list = this._sp.web.lists.getByTitle("PEFramework"); // Change this to your list title
      // await list.items.add({


      // });

      alert('Framework updated successfully!');
      this._getFrameworkItems();
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
      EditImageURL: "",
      isItemInEditMode: false,
      setImageFile: null
    });
    //return false;
  }

  private _onclickbutton(url: any): void {
    url = url.startsWith('http') ? url : 'https://' + url;
    window.open(url);
  }

  private _onDismissDelete(): void {
    this.setState({ _showDialogDelete: !this.state._showDialogDelete });
  }

  private _onEditItem(tmpEditItem: any): void {
    this.setState({
      _showDialogEvents: !this.state._showDialogEvents,
      showEditImageURL: tmpEditItem.ImageUrl ? true : false,
      Title: tmpEditItem.Title,
      setLinkUrl: tmpEditItem.RedirectUrl,
      EditImageURL: tmpEditItem.ImageUrl,
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

      await this._sp.web.lists.getByTitle("PEFramework").items.getById(itemId).delete();
      alert(`Framework deleted successfully.`);
      this._getFrameworkItems();
      this.setState({ _showDialogDelete: !this.state._showDialogDelete });
    } catch (error) {
      console.error(`Error deleting item with ID ${itemId}: `, error);
    }

    //this.deleteListItem();
  }

  public render(): React.ReactElement<IWpFrameworkProps> {
    const {
      hasTeamsContext,
    } = this.props;

    return (
      <section className={`${styles.wpFramework} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={`${styles.frameworkheader} d-flex justify-content-between align-items-center`}>
          <h2>Framework</h2>
          {this.state.isUserInGroup && (
            <div className={styles.iconactions}>
              <a onClick={() => { this._onDismissEvent() }} href='javascript:void(0);'><img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/plus-symbol.png"} width="20px" /></a>
            </div>
          )}
        </div>

        <div className={`${styles.buttongroup}`}>
          {this.state.FrameworkItemsState &&
            this.state.FrameworkItemsState.map((nhItem, index) => (
              <div key={index} className={`${styles.textcenter} d-flex`}>
                <a href='javascript:void(0);' onClick={() => {
                  if (nhItem.RedirectUrl) {
                    this._onclickbutton(nhItem.RedirectUrl);
                  } else {
                    console.log('Redirect URL is not available');
                    // alert('No valid redirect URL available.');
                  }
                }} className={`${styles.frameworkbutton} align-items-center justify-content-between ${this.state.isUserInGroup ? '' : 'w-100'}`}>
                  <img src={nhItem.ImageUrl} width="43" height="43" style={{ color: '#009bda', background: '#ffff' }}></img>
                  {/* <FontAwesomeIcon icon={nhItem.ImageUrl} className="mr-2" /> */}
                  <span className="ms-3">{nhItem.Title}</span>
                </a>
                {this.state.isUserInGroup && (
                  <div>
                    <a onClick={() => { this._onEditItem(nhItem) }} href='javascript:void(0);' className="mx-2" ><FontAwesomeIcon icon={faPen} /></a>
                    <a onClick={() => { this._onDeleteItem(nhItem) }} href='javascript:void(0);' className="mx-1"><FontAwesomeIcon icon={faTrashAlt} /></a>
                  </div>
                )}
              </div>
            ))}
        </div>
        <div>
          <Dialog
            hidden={!this.state._showDialogEvents}
            onDismiss={this._onDismissEvent}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: this.state.isItemInEditMode == true ? 'Edit Framework' : 'Add New Framework',
              subText: 'Provide the below information.'
            }}
            modalProps={{
              isBlocking: false,
              styles: { main: { minWidth: '520px !important' } }
            }}
          >
            <TextField label="Framework Title" value={this.state.Title} onChange={(e, newValue) => this.setState({ Title: (newValue || '') })} />
            <TextField label="Link URL" value={this.state.setLinkUrl} onChange={(e, newValue) => this.setState({ setLinkUrl: (newValue || '') })} />

            {this.state.showEditImageURL && (
              <img src={this.state.EditImageURL} width="43" height="43" style={{ marginTop: '2%', color: '#009bda', background: '#ffff' }}></img>
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

            <Label>Are you sure want to delete this framework ?</Label>

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
