import * as React from 'react';
import styles from './WpIntranetAnnouncements.module.scss';
import "bootstrap/dist/css/bootstrap.min.css";
import type { IWpIntranetAnnouncementsProps } from './IWpIntranetAnnouncementsProps';
import Pagination from '../components/Pagination';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  TextField,
  DatePicker,
  defaultDatePickerStrings,
  mergeStyles,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Stack,
  PrimaryButton,
  DialogFooter,
  Dialog,
  DialogType,
  Label,
  Checkbox,
  Link,
} from "@fluentui/react";
import { SPFI, spfi, SPFx } from '@pnp/sp/presets/all';
//import { IFieldInfo } from "@pnp/sp/fields/types";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/fields";
import { stringIsNullOrEmpty } from '@pnp/core';
import { RotatingLines } from 'react-loader-spinner';
import * as moment from 'moment';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPen, faSearch, faStar as faStarSolid, faTimes, faTrashAlt } from '@fortawesome/free-solid-svg-icons';
import { faStar as faStarRegular } from '@fortawesome/free-regular-svg-icons';

const rootClass = mergeStyles({ maxWidth: 300, selectors: { '> *': { marginBottom: 15 } } });

export interface IWpAnnouncementState {
  _isLoading: boolean;
  _showDialogEvents: boolean;
  _showDialogDelete: boolean;
  _showDialogDeleteConfirm: boolean;
  _showDialogDeleteConfirmYes: boolean;
  _showDialogDeleteConfirmNo: boolean;
  _allData: any;
  currentUser: any;
  filteredItems: any;
  search: string;
  filterSearch: string;
  categoryOption: any;
  catfilterDrop: any;
  autfilterDrop: any;
  categorySelected: any[];
  authorOption: any;
  authorSelected: any[];
  startDate: any;
  endDate: any;
  selectedItem: any;
  openDialog: boolean;
  isAdd: boolean;
  category: string | number;
  publishDate: any;
  fileVersion: number;
  selectedId: string;
  categoryLstChoice: any;
  EditImageURL: string;
  showEditImageURL: boolean;
  selectedFile: any;
  docName: any;
  hideDelModal: boolean;
  saveButtonAdd: boolean;
  savButtonEdit: boolean;
  _EditItem: any;
  _deleteItem: any;
  isItemInEditMode: boolean;
  favoritesAnnouncementItemsState: any[];
  favoritesOnly: boolean;
  currentPage: number;
  itemsPerPage: number;
  isUserInGroup: boolean;
  siteUrl: string;
}
export default class WpIntranetAnnouncements extends React.Component<IWpIntranetAnnouncementsProps, IWpAnnouncementState> {
  public _columns: IColumn[];
  public _sp: SPFI;
  constructor(props: any) {
    super(props);
    sp: this._sp,
      this._sp = spfi().using(SPFx(this.props.context));
    this._columns = [
      {
        key: 'column1',
        name: 'Name',
        fieldName: 'Name',
        minWidth: 365,
        maxWidth: 465,
        isResizable: true,
        styles: { root: { textAlign: 'center', alignItems: 'center', justifyContent: 'center', color: '#75787B', fontStyle: 'normal' } },
        onRender: (item) => (
          <div className={`${styles.customCellName}`}>
            <a
              href='javascript:void(0);'
              onClick={() => { window.open(item.ServerRelativeUrl, '_blank') }}
              rel='noopener noreferrer'
              className={`${styles.customAnchor}`}
              title={item.Name}
            >
              {item.Name}
            </a>
          </div>
        ),
      },
      {
        key: 'column2',
        name: 'Category',
        fieldName: 'Category',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        styles: { root: { textAlign: 'center', alignItems: 'center', justifyContent: 'center', color: '#75787B', fontStyle: 'normal' } },
        onRender: (item) => <span className={`${styles.customCellPadding}`}>{item.Category}</span>,
      },
      {
        key: 'column3',
        name: 'Publish Date',
        fieldName: 'PublishDate',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        styles: { root: { textAlign: 'center', alignItems: 'center', justifyContent: 'center', color: '#75787B', fontStyle: 'normal' } },
        onRender: (item) => <span className={`${styles.customCellCenter}`}>{item.PublishDate}</span>,
      },
      {
        key: 'column4',
        name: 'Author',
        fieldName: 'Author',
        minWidth: 140,
        maxWidth: 200,
        isResizable: true,
        styles: { root: { textAlign: 'center', alignItems: 'center', justifyContent: 'center', color: '#75787B', fontStyle: 'normal' } },
        onRender: (item) => <span className={`${styles.customCell}`}>{item.Author}</span> && item.Author ? item.Author.Title : ''
      },
      {
        key: 'column5',
        name: 'Version',
        fieldName: 'FileVersion',
        minWidth: 80,
        maxWidth: 200,
        isResizable: true,
        styles: { root: { textAlign: 'center', alignItems: 'center', justifyContent: 'center', color: '#75787B', fontStyle: 'normal' } },
        onRender: (item) => <span className={`${styles.customCellCenter}`}>{item.FileVersion}</span>,
      },
      {
        key: 'column6',
        name: 'Size',
        fieldName: 'Length',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        styles: { root: { textAlign: 'center', alignItems: 'center', justifyContent: 'center', color: '#75787B', fontStyle: 'normal' } },
        onRender: (item) => <span className={`${styles.customCellCenter}`}>{item.Length}</span>,
      },
      {
        key: 'column8',
        name: 'Action',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        styles: { root: { textAlign: 'center', alignItems: 'center', justifyContent: 'center', color: '#75787B', fontStyle: 'normal' } },
        onRender: (item) => (<span className={`${styles.customCellCenter}`}>{item.Action}</span> &&
          <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', width: '100%' }}>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              {this.state.isUserInGroup && (
                <>
                  <a
                    onClick={() => { this._onEditItem(item) }}
                    title='Edit'
                    href='javascript:void(0);'
                    style={{
                      color: '#009bda',
                      fontSize: '16px',
                      display: 'inline-flex',
                      alignItems: 'center',
                      justifyContent: 'center'
                    }}
                  >
                    <FontAwesomeIcon icon={faPen} />
                  </a>
                  <a
                    onClick={() => { this._onDeleteItem(item) }}
                    title='Delete'
                    href='javascript:void(0);'
                    style={{
                      color: '#009bda',
                      fontSize: '16px',
                      display: 'inline-flex',
                      alignItems: 'center',
                      justifyContent: 'center'
                    }}
                  >
                    <FontAwesomeIcon icon={faTrashAlt} />
                  </a>
                </>
              )}
              <a
                onClick={() => { this._onFavoriteItem(item) }}
                title={item.IsFavorite ? 'Favorite' : 'Unfavorite'}
                href='javascript:void(0);'
                style={{
                  color: '#009bda',
                  fontSize: '16px',
                  display: 'inline-flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                }}
              >
                <FontAwesomeIcon icon={item.IsFavorite ? faStarSolid : faStarRegular} />
              </a>
            </Stack>
          </div>
        ),
      },
    ]
    this.state = {
      _isLoading: false,
      _showDialogEvents: false,
      _showDialogDelete: false,
      _showDialogDeleteConfirm: false,
      _showDialogDeleteConfirmYes: false,
      _showDialogDeleteConfirmNo: false,
      _allData: [],
      currentUser: {},
      filteredItems: [],
      search: '',
      filterSearch: '',
      startDate: null,
      catfilterDrop: [],
      autfilterDrop: [],
      categoryOption: [],
      authorOption: [],
      categorySelected: [],
      authorSelected: [],
      endDate: null,
      selectedItem: {},
      openDialog: false,
      isAdd: false,
      category: '',
      publishDate: null,
      fileVersion: 0,
      selectedId: '',
      categoryLstChoice: [],
      EditImageURL: "",
      showEditImageURL: false,
      selectedFile: null,
      docName: null,
      hideDelModal: true,
      saveButtonAdd: false,
      savButtonEdit: false,
      _EditItem: {},
      _deleteItem: {},
      isItemInEditMode: false,
      favoritesAnnouncementItemsState: [],
      favoritesOnly: false,
      currentPage: 1,
      itemsPerPage: 10,
      isUserInGroup: false,
      siteUrl: "",
    }
    this.handleGlobalSearch = this.handleGlobalSearch.bind(this);
    this.getCategoryChoiceFields = this.getCategoryChoiceFields.bind(this);
  }
  public componentDidMount(): void {
    this.setState({ _isLoading: true });
    this.checkUserInGroup("PortalAdmins");
    this.getAnnouncementLib();
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
  public async getAnnouncementLib() {
    let _allItems: any[] = [];
    let _catfil: any[] = [];
    let _autfil: any[] = [];
    let filteredCatArr: any[] = [];
    let filteredAutArr: any[] = [];
    let folderUrl = this.props.webUrl + '/' + this.props.LibName;
    try {
      const CurrentUser = await this._sp.web.currentUser();
      this.setState({ currentUser: CurrentUser });
      this._getFavoriteAnnouncementsItems();
      if (!stringIsNullOrEmpty(folderUrl)) {
        let items: any[] = await this._sp.web.getFolderByServerRelativePath(folderUrl).files.select('*').expand('ListItemAllFields, Author')();
        console.log(items);
        if (items && items.length > 0) {
          items.map((item: any, i: number) => {
            _catfil.push(item.ListItemAllFields.Category);
            _autfil.push(item.Author.Title);
            filteredCatArr = Array.from(new Set(_catfil));
            filteredAutArr = Array.from(new Set(_autfil));
            //let length = (Math.round(item.Length) / 1024).toFixed(2);
            let length = this.formatFileSize(item.Length);
            let dotIndex = item.Name.lastIndexOf('.');
            let fileNameWithoutExtension = item.Name.substring(0, dotIndex);
            _allItems.push({
              'Name': fileNameWithoutExtension, 'Category': item.ListItemAllFields.Category,
              'PublishDate': moment(item.ListItemAllFields.PublishDate).format('MM/DD/YYYY'),
              'Author': item.Author,
              'FileVersion': item.ListItemAllFields.FileVersion,
              'Length': length,
              "ID": item.ListItemAllFields.Id,
              "ServerRelativeUrl": item.ServerRelativeUrl,
              "IsFavorite": this.state.favoritesAnnouncementItemsState != null && this.state.favoritesAnnouncementItemsState.length > 0 && this.state.favoritesAnnouncementItemsState.filter(x => x.AnnouncementID == item.ListItemAllFields.Id && x.UserId == this.state.currentUser.Id && x.IsActive == true).length > 0
            });
          });

          _allItems = this.sortAnnouncement(_allItems);
          //_allItems = _allItems.sort((a, b) => b.IsFavorite - a.IsFavorite);
          console.log(_allItems);
          this.setState({ _isLoading: false, _allData: _allItems, filteredItems: _allItems, catfilterDrop: filteredCatArr, autfilterDrop: filteredAutArr });
          this.applyFilters();
          this.bindCategoryDropdown();
        }
      }
      console.log('No Data in Announcement Library/List')
    }
    catch (e) {
      console.error(e.error);
      this.setState({ _isLoading: false });
    }
  }
  private async _getFavoriteAnnouncementsItems(): Promise<any> {
    let favoritesAnnouncementItems: any[] = [];
    this.setState({ favoritesAnnouncementItemsState: [] });
    let tmpFavoritesAnnouncementItems: any = await this._sp.web.lists.getByTitle("Favorite Announcements").items.select("*").filter("UserId eq '" + this.state.currentUser.Id + "'")();
    if (tmpFavoritesAnnouncementItems && tmpFavoritesAnnouncementItems.length > 0) {
      tmpFavoritesAnnouncementItems.forEach((element: { ID: any; AnnouncementID: any; UserId: any; IsActive: any; }) => {
        favoritesAnnouncementItems.push(
          {
            "ID": element.ID,
            "AnnouncementID": element.AnnouncementID,
            "UserId": element.UserId,
            "IsActive": element.IsActive
          });
      })

    } else {
      favoritesAnnouncementItems = [];
    }
    this.setState({ favoritesAnnouncementItemsState: favoritesAnnouncementItems })
    console.log(this.state.favoritesAnnouncementItemsState);
    return Promise.resolve("")
  }
  public sortAnnouncement(_allItems: any[]) {
    _allItems.sort((a, b) => {
      if (b.IsFavorite !== a.IsFavorite) {
        return b.IsFavorite - a.IsFavorite;
      }
      const dateA = new Date(a.PublishDate);
      const dateB = new Date(b.PublishDate);
      return dateB.getTime() - dateA.getTime();
    });
    return _allItems;
  }
  public async bindCategoryDropdown() {
    let _cat: any[] = [];
    let _aut: any[] = [];
    this.state.catfilterDrop.map((item: any) => {
      _cat.push({ key: item, text: item });
    });
    this.state.autfilterDrop.map((item: any) => {
      _aut.push({ key: item, text: item });
    });
    const catMulOption: IDropdownOption[] = _cat;
    const autMulOption: IDropdownOption[] = _aut;
    this.setState({ categoryOption: catMulOption, authorOption: autMulOption });
    this.getCategoryChoiceFields();
  }
  public async getCategoryChoiceFields() {
    //let resultStatusarr: any = [];
    let _catchoice: any = [];
    const sp = spfi().using(SPFx(this.props.context));
    const list = await sp.web.lists.getByTitle("AnnouncementCategory");

    // const rootFolder = await list.rootFolder();
    // const rootFolderPath = rootFolder.ServerRelativeUrl;

    // Get all top-level folders in the document library
    // const folders = await list.items.filter(`FSObjType eq 1 and FileDirRef eq '${rootFolderPath}'`).select('Title', 'FileLeafRef')();
    const announcementCategory = await list.items.select('*')();
    //const folders = await list.items.filter('FSObjType eq 1').select('Title', 'FileLeafRef')();
    announcementCategory.map((item: any) => {
      // _catchoice.push({ key: item.FileLeafRef, text: item.FileLeafRef });
      _catchoice.push({ key: item.Category, text: item.Category });
    });
    const categoryOption: IDropdownOption[] = _catchoice;
    this.setState({ categoryLstChoice: categoryOption });

    /*
    const field1: IFieldInfo = await sp.web.lists.getByTitle(this.props.LibName).fields.getByInternalNameOrTitle("Category")();
    console.log(field1.Choices);
    resultStatusarr = field1.Choices;
    resultStatusarr.map((item: any) => {
      _catchoice.push({ key: item, text: item });
    });
    const categoryOption: IDropdownOption[] = _catchoice;
    this.setState({ categoryLstChoice: categoryOption });
    */
  }
  public formatFileSize = (sizeInBytes: number): string => {
    if (sizeInBytes < 1024) {
      return `${sizeInBytes} Bytes`;
    } else if (sizeInBytes < 1024 * 1024) {
      return `${(sizeInBytes / 1024).toFixed(2)} KB`;
    } else if (sizeInBytes < 1024 * 1024 * 1024) {
      return `${(sizeInBytes / (1024 * 1024)).toFixed(2)} MB`;
    } else {
      return `${(sizeInBytes / (1024 * 1024 * 1024)).toFixed(2)} GB`;
    }
  }
  public _onFormatDate = (date: Date): string => {
    return (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear();
  };
  public changeCategory = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    console.log('here is the things updating...' + item.key + ' ' + item.text);
    this.setState({ category: item.key });
  }
  public async handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    if (event.target.files && event.target.files.length > 0) {
      if (this.state.saveButtonAdd) {
        this.setState({ selectedFile: event.target.files[0] });
      } else if (this.state.savButtonEdit) {
        const fileName = event.target.files[0].name.split('.')[0];
        if (fileName == this.state.docName) {
          this.setState({ selectedFile: event.target.files[0] });
        } else {
          event.target.value = '';
          alert('Please upload file with name ' + this.state.docName);
          return;
        }
      }
    }
  }
  public _onDismissEvent(): void {
    this.setState({
      _showDialogEvents: !this.state._showDialogEvents,
      showEditImageURL: false,
      saveButtonAdd: true,
      savButtonEdit: false,
      publishDate: new Date(),
      fileVersion: 0,
      category: '',
      isItemInEditMode: false
    });
  }
  public _onDismissDelete(): void {
    this.setState({ _showDialogDelete: !this.state._showDialogDelete });
  }
  public async _onEditItem(tmpEditItem: any): Promise<void> {
    //const fileAddResult = this._sp.web.getFolderByServerRelativePath(tmpEditItem.ServerRelativeUrl).files;
    this.setState({
      _showDialogEvents: !this.state._showDialogEvents,
      showEditImageURL: tmpEditItem.ImageURL ? true : false,
      category: tmpEditItem.Category,
      publishDate: new Date(tmpEditItem.PublishDate),
      fileVersion: tmpEditItem.FileVersion,
      docName: tmpEditItem.Name,
      saveButtonAdd: false,
      savButtonEdit: true,
      isItemInEditMode: true,
      _EditItem: tmpEditItem
    });
  }
  public _onDeleteItem(tmpDeleteItem: any): void {
    this.setState({
      _showDialogDelete: !this.state._showDialogDelete,
      showEditImageURL: false,
      saveButtonAdd: false,
      savButtonEdit: false,
      _deleteItem: tmpDeleteItem
    });
  }
  private async _onFavoriteItem(item: any): Promise<void> {
    if (this.state.favoritesAnnouncementItemsState != null && this.state.favoritesAnnouncementItemsState.length > 0 && this.state.favoritesAnnouncementItemsState.filter(x => x.AnnouncementID == item.ID && x.UserId == this.state.currentUser.Id).length > 0) {
      var update = this.state.favoritesAnnouncementItemsState.filter(x => x.AnnouncementID == item.ID && x.UserId == this.state.currentUser.Id)[0];
      await this._sp.web.lists.getByTitle("Favorite Announcements").items.getById(update.ID).update({
        IsActive: update.IsActive == true ? false : true
      }).then((x) => {
        if (update.IsActive == true) {
          alert('Announcement favorite removed successfully!');
          this.getFavoriteAnnouncementLib();
        } else {
          alert('Announcement favorite added successfully!');
          this.getAnnouncementLib();
        }
      });
    } else {
      await this._sp.web.lists.getByTitle("Favorite Announcements").items.add({
        AnnouncementID: item.ID,
        UserId: this.state.currentUser.Id,
        IsActive: true
      }).then((x) => {
        alert('Announcement favorite added successfully!');
        this.getAnnouncementLib();
      });
    }
  }
  public async getFavoriteAnnouncementLib() {
    let _allItems: any[] = [];
    let _catfil: any[] = [];
    let _autfil: any[] = [];
    let filteredCatArr: any[] = [];
    let filteredAutArr: any[] = [];
    let folderUrl = this.props.webUrl + '/' + this.props.LibName;
    try {
      const CurrentUser = await this._sp.web.currentUser();
      this.setState({ currentUser: CurrentUser });
      this._getFavoriteAnnouncementsItems();
      if (!stringIsNullOrEmpty(folderUrl)) {
        let items: any[] = await this._sp.web.getFolderByServerRelativePath(folderUrl).files.select('*').expand('ListItemAllFields, Author')();
        console.log(items);
        if (items && items.length > 0) {
          items.map((item: any, i: number) => {
            _catfil.push(item.ListItemAllFields.Category);
            _autfil.push(item.Author.Title);
            filteredCatArr = Array.from(new Set(_catfil));
            filteredAutArr = Array.from(new Set(_autfil));
            //let length = (Math.round(item.Length) / 1024).toFixed(2);
            let length = this.formatFileSize(item.Length);
            let dotIndex = item.Name.lastIndexOf('.');
            let fileNameWithoutExtension = item.Name.substring(0, dotIndex);
            _allItems.push({
              'Name': fileNameWithoutExtension, 'Category': item.ListItemAllFields.Category,
              'PublishDate': moment(item.ListItemAllFields.PublishDate).format('MM/DD/YYYY'),
              'Author': item.Author,
              'FileVersion': item.ListItemAllFields.FileVersion,
              'Length': length,
              "ID": item.ListItemAllFields.Id,
              "ServerRelativeUrl": item.ServerRelativeUrl,
              "IsFavorite": this.state.favoritesAnnouncementItemsState != null && this.state.favoritesAnnouncementItemsState.length > 0 && this.state.favoritesAnnouncementItemsState.filter(x => x.AnnouncementID == item.ListItemAllFields.Id && x.UserId == this.state.currentUser.Id && x.IsActive == true).length > 0
            });
          });
          _allItems = this.sortAnnouncement(_allItems);
          console.log(_allItems);
          this.setState({ _isLoading: false, _allData: _allItems, filteredItems: _allItems, catfilterDrop: filteredCatArr, autfilterDrop: filteredAutArr });
          this.applyFilters();
          this.bindCategoryDropdown();
        }
      }
      console.log('No Data in Announcement Library/List')
    }
    catch (e) {
      console.error(e.error);
      this.setState({ _isLoading: false });
    }
  }
  public checkSaveValidation() {
    let IsValid: boolean = true;
    if (!this.state.selectedFile) {
      IsValid = false;
      alert('Please select an file to upload.');
    }
    if (this.state.selectedFile && !this.state.category) {
      IsValid = false;
      alert('Please select category.');
    }
    if (this.state.selectedFile && this.state.category && !this.state.publishDate) {
      IsValid = false;
      alert('Please select Publish Date.');
    }
    return IsValid;
  }
  public checkUpdateValidation() {
    let IsValid: boolean = true;
    if (!this.state.category) {
      IsValid = false;
      alert('Please select category.');
    }
    if (this.state.category && !this.state.publishDate) {
      IsValid = false;
      alert('Please select Publish Date.');
    }
    return IsValid;
  }
  public ActionAnnouncement() {
    if (this.state.saveButtonAdd) {
      const IsValidSaveItem: boolean = this.checkSaveValidation();
      if (IsValidSaveItem) {
        this.SaveFileItem();
      }
    }
    else {
      const IsValidUpdateItem: boolean = this.checkUpdateValidation();
      if (IsValidUpdateItem) {
        this.UpdateFileItem();
      }
    }
  }
  public async SaveFileItem() {
    const folderServerRelativeUrl = "/sites/gasopscon/eng/resource%20hub/Announcement";
    const fileAddResult = await this._sp.web.getFolderByServerRelativePath(`/${decodeURI(folderServerRelativeUrl)}`).files.addUsingPath(this.state.selectedFile.name, this.state.selectedFile, { Overwrite: true });
    const item = await fileAddResult.file.getItem();
    await item.update({
      Category: this.state.category,
      PublishDate: this.state.publishDate,
      FileVersion: this.state.fileVersion + 1
    });
    alert('Announcement saved successfully!');
    this.getAnnouncementLib();
    this._onDismissEvent();
  }
  public async UpdateFileItem() {
    if (this.state.selectedFile == null) {
      await this._sp.web.lists.getByTitle('Announcement').items.getById(this.state._EditItem.ID).update({
        Category: this.state.category,
        PublishDate: this.state.publishDate,
        FileVersion: this.state.fileVersion + 1
      });
    } else {
      const folderServerRelativeUrl = "/sites/gasopscon/eng/resource%20hub/Announcement";
      const fileAddResult = await this._sp.web.getFolderByServerRelativePath(`/${decodeURI(folderServerRelativeUrl)}`).files.addUsingPath(this.state.selectedFile.name, this.state.selectedFile, { Overwrite: true });
      const item = await fileAddResult.file.getItem();
      await item.update({
        Category: this.state.category,
        PublishDate: this.state.publishDate,
        FileVersion: this.state.fileVersion + 1
      });
    }
    alert('Announcement updated successfully!');
    this.getAnnouncementLib();
    this._onDismissEvent();
  }
  public async deleteItem() {
    let itemId = this.state._deleteItem.ID;
    try {
      await this._sp.web.lists.getByTitle("Announcement").items.getById(itemId).delete();
      alert('Announcement deleted successfully!');
      this.getAnnouncementLib();
      this.setState({ _showDialogDelete: !this.state._showDialogDelete });
    } catch (error) {
      console.error(`Error deleting item with ID ${itemId}: `, error);
    }
  }
  public handleGlobalSearch = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    this.setState({ search: newValue || '' }, this.applyFilters);
  }
  public handleFilterSearch = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    this.setState({ filterSearch: newValue || '' }, this.applyFilterSearch);
  }
  public handleAuthorChange = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption): void => {
    const authorSelected = option.selected
      ? [...this.state.authorSelected, option.key]
      : this.state.authorSelected.filter(key => key !== option.key);
    this.setState({ authorSelected }, this.applyFilters);
  }
  public handleCategoryChange = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption): void => {
    const categorySelected = option.selected
      ? [...this.state.categorySelected, option.key]
      : this.state.categorySelected.filter(key => key !== option.key);
    this.setState({ categorySelected }, this.applyFilters);
  }
  public handleFavoritesChange = (event: any, checked: any) => {
    this.setState({ favoritesOnly: checked }, this.applyFilters);
  };
  public handleStartDateChange = (date: any) => {
    this.setState({ startDate: date }, this.applyFilters);
  }
  public handleEndDateChange = (date: any) => {
    this.setState({ endDate: date }, this.applyFilters);
  }
  // public clearFilters = () => {
  //   this.setState({
  //     filterSearch: '',
  //     authorSelected: [],
  //     categorySelected: [],
  //     startDate: null,
  //     endDate: null,
  //     favoritesOnly: false,
  //     currentPage: 1,
  //     filteredItems: this.state._allData
  //   }, () => {
  //     this.applyFilters();
  //     this.applyFilterSearch();
  //   });
  // }
  public clearFilters = () => {
    this.setState({
      search: '',
      authorSelected: [],
      categorySelected: [],
      startDate: null,
      endDate: null,
      favoritesOnly: false,
      currentPage: 1,
      filteredItems: this.state._allData
    }, () => {
      this.applyFilters();
      this.applyFilterSearch();
    });
  }
  // public applyFilters = () => {
  //   const { search, _allData, authorSelected, categorySelected, startDate, endDate, favoritesOnly } = this.state;
  //   const filteredItems = _allData.filter((item: { Name: string; Category: string; Author: { Title: string; }; PublishDate: string | number | Date; IsFavorite: boolean; }) => {
  //     const fileName = item.Name ? item.Name.toLowerCase() : '';
  //     const publishDate = new Date(item.PublishDate);
  //     const matchesSearchQuery = !search || search.length < 3 || fileName.includes(search.toLowerCase());
  //     const matchesAuthors = !authorSelected.length || (authorSelected.some(author => author === item.Author.Title));
  //     const matchesCategories = !categorySelected.length || categorySelected.some(category => category === item.Category);
  //     const matchesStartDate = !startDate || publishDate >= startDate;
  //     const matchesEndDate = !endDate || publishDate <= endDate;
  //     const matchesFavorites = !favoritesOnly || item.IsFavorite;

  //     return matchesSearchQuery && matchesAuthors && matchesCategories && matchesStartDate && matchesEndDate && matchesFavorites;
  //   });
  //   this.setState({ filteredItems });
  // }
  public applyFilters = () => {
    const { search, _allData, authorSelected, categorySelected, startDate, endDate, favoritesOnly } = this.state;
    const filteredItems = _allData.filter((item: { Name: string; Category: string; Author: { Title: string; }; PublishDate: string | number | Date; IsFavorite: boolean; }) => {
      const fileName = item.Name ? item.Name.toLowerCase() : '';
      const category = item.Category ? item.Category.toLowerCase() : '';
      const author = item.Author ? item.Author.Title.toLowerCase() : '';
      const publishDate = new Date(item.PublishDate);
      const matchesSearchQuery = !search || search.length < 3 ||
        fileName.includes(search.toLowerCase()) ||
        category.includes(search.toLowerCase()) ||
        author.includes(search.toLowerCase());
      const matchesAuthors = !authorSelected.length || (authorSelected.some(author => author === item.Author.Title));
      const matchesCategories = !categorySelected.length || categorySelected.some(category => category === item.Category);
      const matchesStartDate = !startDate || publishDate >= startDate;
      const matchesEndDate = !endDate || publishDate <= endDate;
      const matchesFavorites = !favoritesOnly || item.IsFavorite;

      return matchesSearchQuery && matchesAuthors && matchesCategories && matchesStartDate && matchesEndDate && matchesFavorites;
    });
    this.setState({ filteredItems, currentPage: 1 });
    this.applyFilterSearch();
  }
  // public applyFilterSearch = () => {
  //   const { autfilterDrop, catfilterDrop, filterSearch } = this.state;
  //   const filteredAuthors: any[] = autfilterDrop
  //     .filter((x: any) => x && x.toLowerCase().includes(filterSearch.toLowerCase()))
  //     .map((author: string) => ({ key: author, text: author }));

  //   const filteredCategories: any[] = catfilterDrop
  //     .filter((x: any) => x && x.toLowerCase().includes(filterSearch.toLowerCase()))
  //     .map((category: string) => ({ key: category, text: category }));

  //   const catMulOption: IDropdownOption[] = filteredCategories;
  //   const autMulOption: IDropdownOption[] = filteredAuthors;
  //   this.setState({
  //     categoryOption: catMulOption,
  //     authorOption: autMulOption
  //   });
  // }
  public applyFilterSearch = () => {
    const { autfilterDrop, catfilterDrop, search } = this.state;
    const filteredAuthors: any[] = autfilterDrop
      .filter((x: any) => x && x.toLowerCase().includes(search.toLowerCase()))
      .map((author: string) => ({ key: author, text: author }));

    const filteredCategories: any[] = catfilterDrop
      .filter((x: any) => x && x.toLowerCase().includes(search.toLowerCase()))
      .map((category: string) => ({ key: category, text: category }));

    const catMulOption: IDropdownOption[] = filteredCategories;
    const autMulOption: IDropdownOption[] = filteredAuthors;
    this.setState({
      categoryOption: catMulOption,
      authorOption: autMulOption
    });
  }
  public handlePageChange = (page: number) => {
    this.setState({ currentPage: page });
  }
  public render(): React.ReactElement<IWpIntranetAnnouncementsProps> {
    const { filteredItems, currentPage, itemsPerPage } = this.state;
    const indexOfLastItem = currentPage * itemsPerPage;
    const indexOfFirstItem = indexOfLastItem - itemsPerPage;
    const currentItems = filteredItems.slice(indexOfFirstItem, indexOfLastItem);
    const totalPages = Math.ceil(filteredItems.length / itemsPerPage);

    return (
      <section >
        {this.state._isLoading == true &&
          <div
            style={{
              width: "100%",
              height: "100",
              display: "flex",
              justifyContent: "center",
              alignItems: "center"
            }}
          >
            <RotatingLines strokeWidth="5" animationDuration="0.75" />
          </div>
        }
        {this.state._isLoading == false &&
          <div className={`${styles.wpIntranetAnnouncements}`}>
            <div className={styles.divAnnouncementsTitleSection}>
              <div className={styles.AnnouncementTitle}>
                <h3>Announcements</h3>
              </div>
              {this.state.isUserInGroup && (
                <div className={`${styles.plusIc}`}>
                  <a onClick={() => { this._onDismissEvent() }}
                    title='Add New Announcement'
                    href='javascript:void(0);'>
                    <img src={this.props.webUrl + "/SiteAssets/PortalImages/plus-symbol.png"} width="20px" />
                  </a>
                </div>
              )}
              {/* <div className={styles.searchContainer}>
                <TextField
                  className={styles.SearchnewTextBox}
                  name="search"
                  value={this.state.search}
                  onChange={this.handleGlobalSearch}
                  placeholder="Search Announcements"
                  onRenderSuffix={() =>
                    this.state.search ? (
                      <FontAwesomeIcon
                        icon={faTimes}
                        className={styles.clearIcon}
                        onClick={() => this.setState({
                          search: '',
                          filteredItems: this.state._allData
                        })}
                        style={{ cursor: 'pointer', marginRight: '8px' }}
                        title="Clear Current Search"
                      />
                    ) : null
                  }
                />
                <FontAwesomeIcon icon={faSearch} className={styles.searchIcon} />
              </div> */}
            </div>
            <div className={`${styles.desingcard} col-lg-12 col-md-12 col-sm-12`}>
              <div className={`${styles.filterLeft} col-lg-2 col-md-2 col-sm-2`}>
                <div className={styles.searchFilterContainer}>
                  {/* <TextField
                    className={styles.SearchnewTextBox}
                    name="filterSearch"
                    value={this.state.filterSearch}
                    onChange={this.handleFilterSearch}
                    placeholder="Search"
                  /> */}
                  <TextField
                    className={styles.SearchnewTextBox}
                    name="search"
                    value={this.state.search}
                    onChange={this.handleGlobalSearch}
                    placeholder="Search"
                    onRenderSuffix={() =>
                      this.state.search ? (
                        <FontAwesomeIcon
                          icon={faTimes}
                          className={styles.clearIcon}
                          onClick={() => this.setState({
                            search: '',
                            filteredItems: this.state._allData
                          }, () => {
                            this.bindCategoryDropdown();
                            this.applyFilters();
                          })}
                          style={{ cursor: 'pointer' }}
                          title="Clear Current Search"
                        />
                      ) : null
                    }
                  />
                  <FontAwesomeIcon icon={faSearch} className={styles.searchIcon} />
                </div>
                <Label className={styles.filterLable_filter}>Filter</Label>
                <br />
                <Label className={styles.filterLable} style={{ marginTop: '0px' }}>Author</Label>
                <Dropdown
                  style={{ width: "100% !important" }}
                  placeholder="Select Author"
                  selectedKeys={this.state.authorSelected}
                  onChange={this.handleAuthorChange}
                  multiSelect
                  options={this.state.authorOption}
                />
                <Label className={styles.filterLable}>Category</Label>
                <Dropdown
                  style={{ width: "100% !important" }}
                  placeholder="Select Category"
                  selectedKeys={this.state.categorySelected}
                  onChange={this.handleCategoryChange}
                  multiSelect
                  options={this.state.categoryOption}
                />
                <div className={styles.filterFavorites}>
                  <Checkbox
                    label="Favorites"
                    checked={this.state.favoritesOnly}
                    onChange={this.handleFavoritesChange}
                  />
                </div>
                <Label className={styles.filterLable}>Publish Date</Label>
                <div className={rootClass}>
                  <DatePicker
                    style={{ width: "100% !important" }}
                    placeholder="Start Date"
                    value={this.state.startDate}
                    onSelectDate={this.handleStartDateChange}
                    strings={defaultDatePickerStrings}
                    formatDate={this._onFormatDate}
                  />
                  <DatePicker
                    style={{ width: "100% !important" }}
                    placeholder="End Date"
                    value={this.state.endDate}
                    onSelectDate={this.handleEndDateChange}
                    strings={defaultDatePickerStrings}
                    formatDate={this._onFormatDate}
                  />
                  <div className={styles.clearbtn}>
                    <DefaultButton text="Clear All Filters" onClick={() => this.clearFilters()} />
                  </div>
                </div>
              </div>
              <div className={`${styles.filterRight} col-lg-10 col-md-10 col-sm-10`}>
                <DetailsList
                  compact={true}
                  items={currentItems}
                  columns={this._columns}
                  setKey="set"
                  layoutMode={DetailsListLayoutMode.justified}
                  selectionPreservedOnEmptyClick={true}
                  selectionMode={SelectionMode.none}
                  onShouldVirtualize={() => false}
                />
                <Pagination
                  currentPage={currentPage}
                  totalPages={totalPages === 0 ? 1 : totalPages}
                  onPageChange={this.handlePageChange}
                />
              </div>
              <Dialog
                hidden={!this.state._showDialogEvents}
                onDismiss={this._onDismissEvent}
                dialogContentProps={{
                  type: DialogType.largeHeader,
                  title: this.state.isItemInEditMode == true ? 'Edit Announcement' : 'Add New Announcement',
                  subText: 'Provide the below information.'
                }}
                modalProps={{
                  isBlocking: false,
                  styles: { main: { minWidth: '520px !important' } }
                }}
              >
                <Label>File Upload</Label>
                <input
                  type="file"
                  onChange={(event) => { this.handleFileChange(event) }}
                  style={{ marginBottom: '10px' }}
                />
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                  <Label>Select Category -</Label>
                  <Link className="text-primary ms-1" href='javascript:void(0);' rel='noopener noreferrer' onClick={() => { window.open(this.state.siteUrl + "/eng/resource%20hub/Lists/AnnouncementCategory/AllItems.aspx", '_blank') }}>
                    Create Category
                  </Link>
                </Stack>
                <Dropdown
                  placeholder="Select Category"
                  options={this.state.categoryLstChoice}
                  defaultSelectedKey={this.state.category}
                  selectedKey={this.state.category}
                  onChange={this.changeCategory}
                  onFocus={this.getCategoryChoiceFields}
                  style={{ marginBottom: '10px' }}
                />
                <DatePicker
                  placeholder="Select Publish Date"
                  label="Publish Date"
                  value={this.state.publishDate}
                  onSelectDate={(e) => { this.setState({ publishDate: e }); }}
                  strings={defaultDatePickerStrings}
                  formatDate={this._onFormatDate}
                  style={{ marginBottom: '10px' }}
                />
                {/* {this.state.isItemInEditMode == true ? <p>Current file name: <b>{this.state.docName}</b></p> : ""} */}
                <DialogFooter>
                  <PrimaryButton onClick={() => this.ActionAnnouncement()} text='Save' style={{ background: '#009bda', color: '#fff', border: 0 }} />
                  <DefaultButton onClick={() => { this._onDismissEvent() }} text="Cancel" />
                </DialogFooter>
              </Dialog>
              <Dialog
                hidden={!this.state._showDialogDelete}
                onDismiss={this._onDismissDelete}
                dialogContentProps={{
                  type: DialogType.largeHeader,
                  title: 'Delete Item',
                }}
                modalProps={{
                  isBlocking: false,
                  styles: { main: { maxWidth: 450 } }
                }}
              >
                <Label>Are you sure want to delete this Announcement?</Label>
                <DialogFooter>
                  <PrimaryButton onClick={() => this.deleteItem()} text="Confirm" style={{ background: '#009bda', color: '#fff', border: 0 }} />
                  <DefaultButton onClick={() => { this._onDismissDelete() }} text="Cancel" />
                </DialogFooter>
              </Dialog>
            </div>
          </div>
        }
      </section>
    );
  }
}


