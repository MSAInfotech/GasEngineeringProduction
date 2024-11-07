import * as React from 'react';
import styles from './WpReportsDashboard.module.scss';
import type { IWpReportsDashboardProps } from './IWpReportsDashboardProps';
import { SPFI, spfi, SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/items/get-all";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
import "bootstrap/dist/css/bootstrap.min.css"

import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPen, faTrashAlt, faSearch } from '@fortawesome/free-solid-svg-icons';
import { PrimaryButton, TextField, Label, Dialog, DialogType, DialogFooter, DefaultButton } from '@fluentui/react';

export interface IWpReportsDashboardState {
  ReportsDashboardItemsState: any[];
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
  isItemInEditMode: boolean,
  search: string,
  filteredItems: any[],
}

export default class WpReportsDashboard extends React.Component<IWpReportsDashboardProps, IWpReportsDashboardState> {

  private _sp: SPFI;

  constructor(props: any) {
    super(props);
    sp: this._sp,
      //this._sp = spfi("https://sempra.sharepoint.com/sites/gasopscon/eng/").using(SPFx(this.props.context));
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      ReportsDashboardItemsState: [],
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
      isUserInGroup: false,
      isItemInEditMode: false,
      _deleteItem: {},
      search: "",
      filteredItems: [],
    }
    this.handleGlobalSearch = this.handleGlobalSearch.bind(this);
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
    let _reportsdashboardItems = [];
    let tmpreportsdashboard = await this._sp.web.lists.getByTitle("ReportsDashboard").items.select("ID", "Title", "ImageURL", "RedirectURL", "SortOrder").orderBy("SortOrder", true).top(4)();
    if (tmpreportsdashboard && tmpreportsdashboard.length > 0) {
      _reportsdashboardItems = tmpreportsdashboard;
    } else {
      _reportsdashboardItems = [];
    }
    this.setState({ ReportsDashboardItemsState: _reportsdashboardItems, filteredItems: _reportsdashboardItems })
    console.log(this.state.ReportsDashboardItemsState);
    return Promise.resolve("")
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
      EditImageURL: "",
      isItemInEditMode: false,
      setImageFile: null
    });
    //return false;
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

  private saveItem() {

    if (this.state.saveButtonAdd) {
      this.uploadImageAndAddItem();
    }
    else {
      this.uploadImageAndUpdateItem();
    }
  }


  private async uploadImageAndAddItem() {
    const { Title, setLinkUrl} = this.state;
    if (!Title || Title.trim() == '') {
      alert('Title cannot be empty.');
      return;
    }
    if (!setLinkUrl || setLinkUrl.trim() == '') {
      alert('Redirect URL cannot be empty.');
      return;
    }
    if (this.state.setImageFile) {
      const folderServerRelativeUrl = "/sites/gasopscon/eng/resource%20hub/ImagesLibrary/ReportsDashboard";
      const fileAddResult = await this._sp.web.getFolderByServerRelativePath(`/${decodeURI(folderServerRelativeUrl)}`).files.addUsingPath(this.state.setImageFile.name, this.state.setImageFile, { Overwrite: true });
      const imageUrl = fileAddResult.data.ServerRelativeUrl;

      const i = await this._sp.web.lists.getByTitle("ReportsDashboard").items.add({
        Title: this.state.Title,
        ImageURL: imageUrl,
        RedirectURL: this.state.setLinkUrl,
        SortOrder: parseInt(this.state.sortorder)
      });

      if (i) {
        console.log("successfully created");
      }

      alert('Reports/Dashboard added successfully!');
      this._getItems();
      this._onDismissEvent();
    } else {
      alert('Please select an image file to upload.');
    }
  };

  private async uploadImageAndUpdateItem() {
    const { Title, setLinkUrl} = this.state;
    if (!Title || Title.trim() == '') {
      alert('Title cannot be empty.');
      return;
    }
    if (!setLinkUrl || setLinkUrl.trim() == '') {
      alert('Redirect URL cannot be empty.');
      return;
    }
    if (this.state.setImageFile || this.state.showEditImageURL) {
      let fileAddResult: any = null;
      let imageUrl: string = '';
      if (this.state.setImageFile) {
        const folderServerRelativeUrl = "/sites/gasopscon/eng/resource%20hub/ImagesLibrary/ReportsDashboard";
        fileAddResult = await this._sp.web.getFolderByServerRelativePath(`/${decodeURI(folderServerRelativeUrl)}`).files.addUsingPath(this.state.setImageFile.name, this.state.setImageFile, { Overwrite: true });
        imageUrl = fileAddResult.data.ServerRelativeUrl;
      } else {
        imageUrl = this.state.EditImageURL;
      }

      let updateID = this.state._EditItem.ID;
      const i = await this._sp.web.lists.getByTitle("ReportsDashboard").items.getById(updateID).update({
        Title: this.state.Title,
        ImageURL: imageUrl,
        RedirectURL: this.state.setLinkUrl,
        SortOrder: parseInt(this.state.sortorder),
      });

      if (i) {
        console.log("Reports/Dashboard updated successfully.");
      }

      alert('Reports/Dashboard updated successfully!');
      this._getItems();
      this._onDismissEvent();
    } else {
      alert('Please select an image file to upload.');
    }
  };

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

      await this._sp.web.lists.getByTitle("ReportsDashboard").items.getById(itemId).delete();
      alert(`Reports/Dashboard deleted successfully.`);
      this._getItems();
      this.setState({ _showDialogDelete: !this.state._showDialogDelete });
    } catch (error) {
      console.error(`Error deleting item with ID ${itemId}: `, error);
    }

    //this.deleteListItem();
  }

  public handleGlobalSearch = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    this.setState({ search: newValue || '' }, this.applyFilters);
  }

  public applyFilters = () => {
    const { search } = this.state;
    this._searchitems(search);
  }

  private async _searchitems(searchTerm: string): Promise<any> {
    try {

      this.setState({ filteredItems: [] })
      let data: any[] = this.state.ReportsDashboardItemsState;

      data = data.filter(x => x.Title.toLowerCase().includes(searchTerm.toLowerCase()));

      if (data.length === 0) {

        this.setState({ filteredItems: [] })
      }
      else {
        this.setState({ filteredItems: data })
      }

    } catch (error) {
      console.error("Error retrieving folder hierarchy:", error);
      throw error;
    }
  }

  public render(): React.ReactElement<IWpReportsDashboardProps> {
    const {
      hasTeamsContext,

    } = this.props;

    return (
      <section className={`${styles.wpReportsDashboard} ${hasTeamsContext ? styles.teams : ''}`}>
        {/* <div className="container"> */}
        <div className="row">
          <div className={styles.divReportsTitleSection}>
            <div className={styles.ReportsTitle}>
              <h3>Reports/ Dashboard</h3>
            </div>
            <div className={styles.searchContainer}>
              <TextField
                className={styles.SearchnewTextBox}
                name="search"
                value={this.state.search}
                onChange={this.handleGlobalSearch}
                placeholder="Search Reports/Dashboards"
                styles={{
                  field: { outline: 'none' },
                }}
              />
              <FontAwesomeIcon icon={faSearch} className={styles.searchIcon} />
            </div>
          </div>
          {/* <div className="col-md-12 col-sm-12 col-lg-12 justify-content-end"> */}
          {/* <div className={`${styles.Searchnew} col-md-4 col-sm-4 col-lg-4 float-end`}> 
            <div className={styles.SearchWrapper}>
              <TextField
                className={styles.SearchnewTextBox}
                name='search'
                value={this.state.search}
                placeholder="Search Reports/Dashboards"
                onChange={this.handleGlobalSearch}
              />
              <div className={styles.SearchIcon}>
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <g clipPath="url(#clip0)">
                    <path d="M15.5 14H14.71L14.43 13.73C15.41 12.59 16 11.11 16 9.5C16 5.91 13.09 3 9.5 3C5.91 3 3 5.91 3 9.5C3 13.09 5.91 16 9.5 16C11.11 16 12.59 15.41 13.73 14.43L14 14.71V15.5L19 20.49L20.49 19L15.5 14ZM9.5 14C7.01 14 5 11.99 5 9.5C5 7.01 7.01 5 9.5 5C11.99 5 14 7.01 14 9.5C14 11.99 11.99 14 9.5 14Z" fill="#899FB5" />
                  </g>
                  <defs>
                    <clipPath id="clip0">
                      <rect width="24" height="24" fill="white" />
                    </clipPath>
                  </defs>
                </svg>
              </div>
            </div>
          </div> */}
          {/* <TextField
              name='search'
              value={this.state.search}
              placeholder="Search Reports/Dashboards"
              className={`${styles.searchInput} col-md-3 col-sm-2 col-lg-4 float-end`}
              onChange={this.handleGlobalSearch}
            /> */}
          {/* </div> */}
        </div>

        <div className="row">
          {this.state.isUserInGroup && (
            <div className="col-md-12 col-sm-12 col-lg-12 d-flex justify-content-end mb-3 mt-3">
              <a onClick={() => { this._onDismissEvent() }} href="javascript:void(0);">
                <img
                  src={`${this.props.WebServerRelativeURL}/SiteAssets/PortalImages/plus-symbol.png`}
                  width="20px"
                  alt="Add"
                />
              </a>
            </div>
          )}
        </div>

        {/* Dashboard Items */}
        <div className="row">
          {this.state.filteredItems && this.state.filteredItems.map((nhItem, index) => (
            <div className="col-md-5 col-lg-5 col-sm-5 mb-4" key={index}>
              <div className={styles.dashboardItem}>
                <div className={styles.dashboardHeader}>
                  <h3 className={styles.dashboardTitle}>{nhItem.Title}</h3>
                  {this.state.isUserInGroup && (
                    <div className={styles.actionIcons}>
                      <a
                        onClick={() => { this._onEditItem(nhItem) }}
                        href="javascript:void(0);"
                        className={styles.editIcon}
                      >
                        <FontAwesomeIcon icon={faPen} />
                      </a>
                      <a
                        onClick={() => { this._onDeleteItem(nhItem) }}
                        href="javascript:void(0);"
                        className={styles.deleteIcon}
                      >
                        <FontAwesomeIcon icon={faTrashAlt} />
                      </a>
                    </div>
                  )}
                </div>
                <a href='javascript:void(0);' 
                onClick={() => {
                  let url = nhItem.RedirectURL.startsWith('http') ? nhItem.RedirectURL : 'https://' + nhItem.RedirectURL;
                  window.open(url);
                }}
                className={styles.thumbnailLink}>
                  <img src={nhItem.ImageURL} alt={nhItem.Title} className={styles.thumbnailImage} />
                </a>
              </div>
            </div>
          ))}
        </div>
        {/* </div> */}

        <div>
          <Dialog
            hidden={!this.state._showDialogEvents}
            onDismiss={this._onDismissEvent}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: this.state.isItemInEditMode == true ? 'Edit Reports/Dashboard' : 'Add New Reports/Dashboard',
              subText: 'Provide the below information.'
            }}
            modalProps={{
              isBlocking: false,
              styles: { main: { minWidth: '520px !important' } }
            }}
          >
            <TextField label="Reports/Dashboard Title" value={this.state.Title} onChange={(e, newValue) => this.setState({ Title: (newValue || '') })} />
            <TextField label="Redirect URL" value={this.state.setLinkUrl} onChange={(e, newValue) => this.setState({ setLinkUrl: (newValue || '') })} />
            <TextField label="Sort Order" value={this.state.sortorder} onChange={(e, newValue) => this.setState({ sortorder: (newValue || '') })} />

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

            <Label>Are you sure want to delete this Reports/Dashboard ?</Label>

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
