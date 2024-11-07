import * as React from 'react';
import styles from './WpComponentsLinksForProjectEng.module.scss';
import type { IWpComponentsLinksForProjectEngProps } from './IWpComponentsLinksForProjectEngProps';
import { SPFI, spfi, SPFx }  from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/items/get-all";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
import "bootstrap/dist/css/bootstrap.min.css"

import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faChevronDown, faChevronUp, faPen, faTrashAlt } from '@fortawesome/free-solid-svg-icons';
import { PrimaryButton, TextField, Label, Dialog, DialogType, DialogFooter, DefaultButton } from '@fluentui/react';


export interface IWpComponentsLinksForProjectEngState {
  ComponentLinksItemsState: any[];
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
  hoverdescription: string,
  Title: string,
  EditImageURL: string,
  showEditImageURL: boolean,
  saveButtonAdd: boolean,
  savButtonEdit: boolean,
  _EditItem: any;
  isUserInGroup: boolean;
  isItemInEditMode: boolean;
  isCollapsed: boolean;
}

export default class WpComponentsLinksForProjectEng extends React.Component<IWpComponentsLinksForProjectEngProps, IWpComponentsLinksForProjectEngState> {

  private _sp: SPFI;

  constructor(props: any) {
    super(props);
    sp: this._sp,
      //this._sp = spfi("https://sempra.sharepoint.com/sites/gasopscon/eng/").using(SPFx(this.props.context));
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      ComponentLinksItemsState: [],
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
      hoverdescription: "",
      Title: "",
      EditImageURL: "",
      showEditImageURL: false,
      saveButtonAdd: false,
      savButtonEdit: false,
      _EditItem: {},
      isUserInGroup: false,
      isItemInEditMode: false,
      _deleteItem: {},
      isCollapsed: false
    }
    this.toggleCollapse = this.toggleCollapse.bind(this);
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
      await this._getItems();

    } catch (error) {
      // Handle error
    }
  }


  private async _getItems(): Promise<string> {
    let _componentLinksItems = [];
    let categoryfilter = this.props.category;
    debugger;
    let tmpcomponentLinks = await this._sp.web.lists.getByTitle("DepartmentsForProjectEng").items.select("ID", "Title", "Category", "ImageURL", "RedirectURL", "SortOrder", "HoverDescription").filter("Category eq '" + categoryfilter + "'").orderBy("SortOrder", true)();
    if (tmpcomponentLinks && tmpcomponentLinks.length > 0) {
      _componentLinksItems = tmpcomponentLinks;
    } else {
      _componentLinksItems = [];
    }
    this.setState({ ComponentLinksItemsState: _componentLinksItems })
    console.log(this.state.ComponentLinksItemsState);
    return Promise.resolve("")
  }

  toggleCollapse() {
    this.setState((prevState) => ({ isCollapsed: !prevState.isCollapsed }));
  }

  private _onDismissEvent(): void {
    this.setState({
      _showDialogEvents: !this.state._showDialogEvents,
      showEditImageURL: false,
      saveButtonAdd: true,
      savButtonEdit: false,
      Title: "",
      setLinkUrl: "",
      sortorder: "",
      hoverdescription: "",
      EditImageURL: "",
      isItemInEditMode: false,
      setImageFile: null
    });
    //return false;
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

  private saveItem() {
    debugger;
    if (this.state.saveButtonAdd) {
      this.uploadImageAndAddItem();
    }
    else {
      this.uploadImageAndUpdateItem();
    }
  }


  private async uploadImageAndAddItem() {
    debugger;
    if (this.state.setImageFile) {
      debugger;
      //const { file, filename } = this.state;
      //console.log(`Generated URL: ${this.context.pageContext.web.serverRelativeUrl}/ImagesLibrary`);
      const folderServerRelativeUrl = "/sites/gasopscon/eng/resource%20hub/ImagesLibrary";
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

      debugger;
      const imageUrl = fileAddResult.data.ServerRelativeUrl;
      //const setImageUrl = imageUrl;

      const i = await this._sp.web.lists.getByTitle("DepartmentsForProjectEng").items.add({
        Title: this.state.Title,
        ImageURL: imageUrl,
        RedirectURL: this.state.setLinkUrl,
        SortOrder: parseInt(this.state.sortorder),
        Category: this.props.category,
        HoverDescription: this.state.hoverdescription
      });

      if (i) {
        console.log("successfully created");
      }

      // const list = this._sp.web.lists.getByTitle("QuickLinks"); // Change this to your list title
      // await list.items.add({


      // });

      alert('Component link added successfully!');
      this._getItems();
      this._onDismissEvent();
    } else {
      alert('Please select an image file to upload.');
    }
  };

  private async uploadImageAndUpdateItem() {
    debugger;
    if (this.state.setImageFile || this.state.showEditImageURL) {
      debugger;
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
      const i = await this._sp.web.lists.getByTitle("DepartmentsForProjectEng").items.getById(updateID).update({
        Title: this.state.Title,
        ImageURL: imageUrl,
        RedirectURL: this.state.setLinkUrl,
        SortOrder: parseInt(this.state.sortorder),
        HoverDescription: this.state.hoverdescription
      });

      if (i) {
        console.log("Component link updated successfully.");
      }

      // const list = this._sp.web.lists.getByTitle("QuickLinks"); // Change this to your list title
      // await list.items.add({


      // });

      alert('Component link updated successfully!');
      this._getItems();
      this._onDismissEvent();
    } else {
      alert('Please select an image file to upload.');
    }
  };

  private _onDismissDelete(): void {
    this.setState({ _showDialogDelete: !this.state._showDialogDelete});
  }

  private _onEditItem(tmpEditItem: any): void {
    this.setState({
      _showDialogEvents: !this.state._showDialogEvents,
      showEditImageURL: tmpEditItem.ImageURL ? true : false,
      Title: tmpEditItem.Title,
      setLinkUrl: tmpEditItem.RedirectURL,
      sortorder: tmpEditItem.SortOrder,
      hoverdescription: tmpEditItem.HoverDescription,
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

  private async deleteItem(){
    
    let itemId = this.state._deleteItem.ID;
    try {
      
      await this._sp.web.lists.getByTitle("DepartmentsForProjectEng").items.getById(itemId).delete();
      alert(`Component link deleted successfully.`);
      this._getItems();
      this.setState({ _showDialogDelete: !this.state._showDialogDelete});
    } catch (error) {
      console.error(`Error deleting item with ID ${itemId}: `, error);
    } 

//this.deleteListItem();
}



  public render(): React.ReactElement<IWpComponentsLinksForProjectEngProps> {
    const {
      hasTeamsContext,
    } = this.props;

    const {
      isCollapsed,
    } = this.state

    return (
      <section className={`${styles.wpComponentsLinksForProjectEng} ${hasTeamsContext ? styles.teams : ''}`}>
        <div style={{ display: 'inline-flex' }} className='col-md-12 col-sm-12 col-lg-12'>
          <div style={{ marginLeft: '6.5%', width: '4%' }}>
            <button onClick={this.toggleCollapse} className={styles.toggleButton}>
              {isCollapsed ? <FontAwesomeIcon icon={faChevronUp} /> : <FontAwesomeIcon icon={faChevronDown} />}
            </button>
          </div>
          <div style={{ width: '96%' }}>
            <h3 style={{ fontFamily: 'Montserrat', fontStyle: 'normal', fontWeight: 700, fontSize: '24px', lineHeight: '29px', color: '#004693' }}>List of Departments we work with</h3>
          </div>
          {this.state.isUserInGroup && (
            <div className={`${styles.News_oprt} col-md-2 col-sm-2 col-lg-2`}>
              <ul>
                <li><a onClick={() => { this._onDismissEvent() }} href='javascript:void(0);'><img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/plus-symbol.png"} width="20px" /></a></li>
              </ul>
            </div>
          )}
        </div>
        <div className={`${styles.Internal_content} ${isCollapsed ? styles.collapsed : ''}`}>
          <div className={`${styles.Img_content}`}>
            <div className={`${styles.Photos_cont}`}>
              {/* <div style={{ width: '4%' }}>&nbsp;</div> */}
              {this.state.ComponentLinksItemsState &&
                this.state.ComponentLinksItemsState.map((nhItem, index) => (
                  <>
                    {(index) % 4 == 0 ? <div style={{ width: '7%', float: 'inline-start' }}>&nbsp;</div> : ''}
                    <div className={`${styles.Thumbnails_cont}`}>
                      <a href={nhItem.RedirectURL}>
                        <img src={nhItem.ImageURL} alt="" />
                        <div className={`${styles.Black_cont}`}></div>
                        <div className={`${styles.Title_cont}`}>{nhItem.Title}</div>
                        <h5 className={`${styles.Mission}`}>Mission</h5>
                        <div className={`${styles.Hover_Desc}`}>{nhItem.HoverDescription}</div>
                      </a>

                      {this.state.isUserInGroup && (
                        <div className={`${styles.Edit_del} col-md-3 col-sm-3 col-lg-3`}>
                          <a onClick={() => { this._onDeleteItem(nhItem) }} href='javascript:void(0);' style={{ float: 'right' }}><FontAwesomeIcon icon={faTrashAlt} /></a>
                          <a onClick={() => { this._onEditItem(nhItem) }} href='javascript:void(0);' style={{ marginRight: '5px', float: 'right' }}><FontAwesomeIcon icon={faPen} /></a>
                        </div>
                      )}
                    </div>
                  </>
                ))
              }
            </div>
          </div>
        </div>

        <div>
          <Dialog
            hidden={!this.state._showDialogEvents}
            onDismiss={this._onDismissEvent}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: this.state.isItemInEditMode == true ? 'Edit Component Link' : 'Add New Component Link',
              subText: 'Provide the below information.'
            }}
            modalProps={{
              isBlocking: false,
              styles: { main: { minWidth: '520px !important' } }
            }}
          >
            <TextField label="Component Link Title" value={this.state.Title} onChange={(e, newValue) => this.setState({ Title: (newValue || '') })} />
            <TextField label="Link URL" value={this.state.setLinkUrl} onChange={(e, newValue) => this.setState({ setLinkUrl: (newValue || '') })} />
            <TextField label="Sort Order" value={this.state.sortorder} onChange={(e, newValue) => this.setState({ sortorder: (newValue || '') })} />
            <TextField label="Hover Description" value={this.state.hoverdescription} onChange={(e, newValue) => this.setState({ hoverdescription: (newValue || '') })} />

            {this.state.showEditImageURL && (
              <img src={this.state.EditImageURL} width="30" height="30" style={{ padding: '5px', color: '#009bda', background: '#ffff' }}></img>
            )}

            <Label>Image Upload</Label>
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

            <Label>Are you sure want to delete this component link ?</Label>

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
