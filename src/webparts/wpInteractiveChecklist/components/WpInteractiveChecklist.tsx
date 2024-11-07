import * as React from 'react';
import styles from './WpInteractiveChecklist.module.scss';
import type { IWpInteractiveChecklistProps } from './IWpInteractiveChecklistProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/items/get-all";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
import "bootstrap/dist/css/bootstrap.min.css";
import { DragDropContext, Droppable, Draggable } from 'react-beautiful-dnd';

import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPen, faTrashAlt, faChevronDown, faChevronRight, /*faDownload*/ } from '@fortawesome/free-solid-svg-icons';
import { faFolder } from '@fortawesome/free-regular-svg-icons';
import { PrimaryButton, TextField, Label, Dialog, DialogType, DialogFooter, DefaultButton } from '@fluentui/react';


import { spfi, SPFI, SPFx } from '@pnp/sp/presets/all';

export interface IWpInteractiveChecklistsState {
  categoriesState: any[];
  _showDialogEvents: boolean,
  _showDialogEventsForStage: boolean,
  _isLoading: boolean;
  _allStages: any[],
  _stageArray: any[];
  _selectedCategory: any;
  isUserInGroup: boolean;
  showEditImageURL: boolean,
  saveButtonAdd: boolean,
  savButtonEdit: boolean,
  _EditItem: any;
  isItemInEditMode: boolean;
  isCollapsed: boolean;
  _deleteItem: any,
  Title: string,
  _showDialogDelete: boolean,
  _showDialogDeleteForStage: boolean,
  subSiteUrl: string,
  allcategoriesInfo: any[],
  stagesInfo: any[],
  deliverablesInfo: any[],
  _selectedCategoryForStage: { name: string, url: string } | null;
  _collapsedCategories: { [key: string]: boolean };
  fileToUpload: any;
  existingFileName: string | null;
  search: string;
  filteredItems: any,
}


export default class WpInteractiveChecklist extends React.Component<IWpInteractiveChecklistProps, IWpInteractiveChecklistsState> {

  private _sp: SPFI;

  constructor(props: any) {
    super(props);
    sp: this._sp,
      //this._sp = spfi("https://sempra.sharepoint.com/sites/gasopscon/eng/").using(SPFx(this.props.context));
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      categoriesState: [],
      _isLoading: true,
      _allStages: [],
      _stageArray: [],
      _selectedCategory: "",
      isUserInGroup: false,
      _showDialogEvents: false,
      _showDialogEventsForStage: false,
      showEditImageURL: false,
      saveButtonAdd: false,
      savButtonEdit: false,
      _EditItem: {},
      isItemInEditMode: false,
      _deleteItem: {},
      isCollapsed: false,
      Title: "",
      _showDialogDelete: false,
      _showDialogDeleteForStage: false,
      subSiteUrl: "/sites/gasopscon/eng/resource%20hub",
      allcategoriesInfo: [],
      stagesInfo: [],
      deliverablesInfo: [],
      _selectedCategoryForStage: null,
      _collapsedCategories: {},
      fileToUpload: null,
      existingFileName: null,
      search: "",
      filteredItems: [],
    }
    this.handleGlobalSearch = this.handleGlobalSearch.bind(this);
    this.onDragEnd = this.onDragEnd.bind(this);
  }

  public componentDidMount(): void {
    this.checkUserInGroup("PortalAdmins");
    this._getCategories();
    this._getInteractiveChecklistInfo();
  }


  private async checkUserInGroup(groupName: string): Promise<void> {
    const currentUser = await this._sp.web.currentUser();
    const groups = await this._sp.web.siteGroups.getByName(groupName).users();
    const isUserInGroupExists = groups.some(user => user.Id === currentUser.Id);
    this.setState({ isUserInGroup: isUserInGroupExists });
  }

  private _onDismissEvent(): void {
    this.setState({
      _showDialogEvents: !this.state._showDialogEvents,
      showEditImageURL: false,
      saveButtonAdd: true,
      savButtonEdit: false,
      Title: "",
      isItemInEditMode: false,
    });
    //return false;
  }

  private _onDismissEventForStage(selectedCategory: { name: string, url: string }): void {
    this.setState({
      _showDialogEventsForStage: !this.state._showDialogEventsForStage,
      showEditImageURL: false,
      saveButtonAdd: true,
      savButtonEdit: false,
      Title: "",
      isItemInEditMode: false,
      _selectedCategoryForStage: selectedCategory,
      existingFileName: null
    });
    //return false;
  }

  private _onDismissEventForStageWrapper = () => {
    const { _selectedCategoryForStage } = this.state;
    this._onDismissEventForStage(_selectedCategoryForStage || { name: "", url: "" });
  }


  private saveItemStage() {
    if (this.state.saveButtonAdd) {
      this.addStage();
    }
    else {
      this.updateStage();
    }
  }

  private _onDismissDeleteForStage(): void {
    this.setState({ _showDialogDeleteForStage: !this.state._showDialogDeleteForStage });
  }

  private async _onEditItemStage(tmpEditItem: any): Promise<void> {
    this.setState({
      _showDialogEventsForStage: !this.state._showDialogEventsForStage,
      showEditImageURL: tmpEditItem.ImageURL ? true : false,
      Title: tmpEditItem.name,
      saveButtonAdd: false,
      savButtonEdit: true,
      isItemInEditMode: true,
      _EditItem: tmpEditItem,
    });

    // Fetch the existing file details
    const folderUrl = tmpEditItem.url;
    const folder = await this._sp.web.getFolderByServerRelativePath(folderUrl);
    const files = await folder.files.select("Name", "TimeLastModified", "ServerRelativeUrl").orderBy("TimeLastModified", false).top(1)();
    if (files.length > 0) {
      this.setState({ existingFileName: files[0].Name });
    }
  }

  private async _downloadFile(tmpDeleteItem: any): Promise<void> {
    try {
      const folderUrl = tmpDeleteItem.url;
      const folder = await this._sp.web.getFolderByServerRelativePath(folderUrl);
      const files = await folder.files.select("Name", "TimeLastModified", "ServerRelativeUrl").orderBy("TimeLastModified", false).top(1)();
      if (files && files.length > 0) {
        const file = await this._sp.web.getFileByServerRelativePath(files[0].ServerRelativeUrl).getBlob();
        const url = window.URL.createObjectURL(file);
        const link = document.createElement('a');
        link.href = url;
        link.download = files[0].ServerRelativeUrl.substring(files[0].ServerRelativeUrl.lastIndexOf('/') + 1);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
      }
    } catch (error) {
      console.error('Error downloading file:', error);
    }
    //return false;
  }
  private _onDeleteItemStage(tmpDeleteItem: any): void {
    this.setState({
      _showDialogDeleteForStage: !this.state._showDialogDeleteForStage,
      showEditImageURL: false,
      saveButtonAdd: false,
      savButtonEdit: false,
      _deleteItem: tmpDeleteItem
    });
    //return false;
  }

  handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    if (event.target.files && event.target.files.length > 0) {
      this.setState({ fileToUpload: event.target.files[0] });
    }
  }

  private async addStage() {
    const { Title, _selectedCategoryForStage, fileToUpload } = this.state;

    if (!Title || Title.trim() == '' || !_selectedCategoryForStage) {
      alert('Stage Title cannot be empty.');
      return;
    }

    if (this.state.fileToUpload) {
      try {
        // Get the selected category's URL
        const categoryFolderUrl = _selectedCategoryForStage.url;

        // Construct the path for the new stage folder
        const libraryPath = decodeURI(this.state.subSiteUrl) + "/InteractiveChecklists";
        const newStageFolderUrl = `${categoryFolderUrl}/${Title}`;

        // Add the stage folder inside the selected category
        await this._sp.web.getFolderByServerRelativePath(libraryPath).folders.addUsingPath(newStageFolderUrl, true);

        if (fileToUpload) {
          await this._sp.web.getFolderByServerRelativePath(newStageFolderUrl).files.addUsingPath(fileToUpload.name, fileToUpload, { Overwrite: true });
        }

        alert('Stage added successfully!');
        this._getCategories(); // Refresh the list of categories
        this._getInteractiveChecklistInfo(); // Refresh the list of checklist
        this._onDismissEventForStageWrapper(); // Close the dialog
      } catch (error) {
        console.error('Error adding stage:', error);
        alert('Error adding stage. Please try again.');
      }
    } else {
      alert('Please select an file to upload.');
    }
  };

  private async updateStage() {
    const { Title, fileToUpload } = this.state;

    if (!Title || Title.trim() == '') {
      alert('Stage Title cannot be empty.');
      return;
    }

    try {

      let newFolderUrl: string = this.state._EditItem.url;
      if (this.state._EditItem.name != this.state.Title) {
        const currentFolderUrl = this.state._EditItem.url;
        // Extract the parent folder URL and construct the new folder URL
        const parentFolderUrl = currentFolderUrl.substring(0, currentFolderUrl.lastIndexOf('/'));
        newFolderUrl = `${parentFolderUrl}/${this.state.Title}`;
        // Move the folder to the new URL
        await this._sp.web.getFolderByServerRelativePath(currentFolderUrl).moveByPath(newFolderUrl);
      }
      if (fileToUpload) {
        await this._sp.web.getFolderByServerRelativePath(newFolderUrl).files.addUsingPath(fileToUpload.name, fileToUpload, { Overwrite: true });

        /*const targetFolder = this._sp.web.getFolderByServerRelativePath(newFolderUrl);
        const files = await targetFolder.files(); // Get list of files in the target folder
        // Check if file already exists
        const fileExists = files.some(file => file.Name === fileToUpload.name);
        if (!fileExists) {
          await this._sp.web.getFolderByServerRelativePath(newFolderUrl).files.addUsingPath(fileToUpload.name, fileToUpload, { Overwrite: true });
        }*/
      }

      alert('Stage updated successfully!');
      this._getCategories(); // Refresh the list of categories
      this._getInteractiveChecklistInfo(); // Refresh the list of checklist
      this._onDismissEventForStageWrapper(); // Close the dialog
    } catch (error) {
      console.error('Error updating stage:', error);
      alert('Error updating stage. Please try again.');
    }
  };

  private async deleteStage() {
    debugger;
    try {
      // Get the folder's current URL
      const currentFolderUrl = this.state._deleteItem.url;

      // Extract the parent folder URL and construct the new folder URL
      //const parentFolderUrl = currentFolderUrl.substring(0, currentFolderUrl.lastIndexOf('/'));
      //const newFolderUrl = `${parentFolderUrl}/${this.state.Title}`;

      // Delete the folder
      await this._sp.web.getFolderByServerRelativePath(currentFolderUrl).delete();

      alert('stage deleted successfully!');
      this._getCategories(); // Refresh the list of categories
      this._getInteractiveChecklistInfo(); // Refresh the list of checklist
      this.setState({ _showDialogDeleteForStage: !this.state._showDialogDeleteForStage });
    } catch (error) {
      console.error('Error deleted stage:', error);
      alert('Error deleted stage. Please try again.');
    }
  }

  private async _getCategories(): Promise<string> {
    let _categories: { Title: string; }[] = [];

    const libraryPath = decodeURI(this.state.subSiteUrl) + "/InteractiveChecklists"; // Server-relative URL of your library
    const folders = await this._sp.web.getFolderByServerRelativePath(libraryPath).folders();

    // Assuming folders are returned in the format { Name: 'FolderName' }
    let categoryList = folders.map(folder => ({ Title: folder.Name }));

    if (categoryList && categoryList.length > 0) {
      _categories = categoryList;
    } else {
      _categories = [];
    }
    this.setState({ categoriesState: _categories })
    console.log(this.state.categoriesState);
    return Promise.resolve("")
  }

  private async _getInteractiveChecklistInfo(): Promise<string> {
    debugger;
    const libraryPath = decodeURI(this.state.subSiteUrl) + "/InteractiveChecklists"; // Server-relative URL of your library
    let counter: number = 1;
    let data = await this._getFolderHierarchy(libraryPath, counter);
    console.log(data);
    if (data) {
      this.setState({
        allcategoriesInfo: data.folders,
        filteredItems: data.folders,
      })
    }
    return Promise.resolve("")
  }


  private async _getFolderHierarchy(folderUrl: any, counter: number): Promise<any> {
    try {
      const folder = this._sp.web.getFolderByServerRelativePath(folderUrl);

      const files = await folder.files();
      let subfolders:any[] = await folder.folders.select("Name", "ServerRelativeUrl","ItemCount", "ListItemAllFields/SortDeliverable").expand("ListItemAllFields")();
      // if (subfolders.length > 0 && subfolders[0].Name === 'Forms') {
      //   subfolders.shift(); // Remove the first element
      // }
      subfolders = subfolders.filter(subfolder => {
        // Check if the folder is named 'Forms' and has properties that match the system folder
        const isSystemFormsFolder = subfolder.Name === 'Forms' && subfolder.ItemCount === 0;
        return !isSystemFormsFolder;
      });

      subfolders = subfolders.sort((a, b) => a.ListItemAllFields?.SortDeliverable - b.ListItemAllFields?.SortDeliverable);
      const fileList = files.map(file => file.Name);
      const folderList = subfolders.map(subfolder => ({
        name: subfolder.Name,
        url: subfolder.ServerRelativeUrl,
        sortDeliverable:subfolder.ListItemAllFields?.SortDeliverable
      }));
      //let subfolderHierarchy: any[] = [];
      //if (counter <= 3) {
      const subfolderHierarchy = await Promise.all(
        folderList.map(async subfolder => ({
          ...subfolder,
          children: await this._getFolderHierarchy(subfolder.url, counter++)
        }))
      );
      //}

      return {
        files: fileList,
        folders: subfolderHierarchy
      };
    } catch (error) {
      console.error("Error retrieving folder hierarchy:", error);
      throw error;
    }
  }

  private saveItem() {
    if (this.state.saveButtonAdd) {
      this.addCategory();
    }
    else {
      this.updateCategory();
    }
  }

  private _onDismissDelete(): void {
    this.setState({ _showDialogDelete: !this.state._showDialogDelete });
  }

  private _onEditItem(tmpEditItem: any): void {
    this.setState({
      _showDialogEvents: !this.state._showDialogEvents,
      showEditImageURL: tmpEditItem.ImageURL ? true : false,
      Title: tmpEditItem.name,
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

  private async addCategory() {
    const { Title } = this.state;
    if (!Title || Title.trim() == '') {
      alert('Category Title cannot be empty.');
      return;
    }

    try {
      const libraryPath = decodeURI(this.state.subSiteUrl) + "/InteractiveChecklists";
      const folderPath = `${libraryPath}/${Title}`;

      // Check if a folder with the same name already exists
      const folderExists = await this._sp.web.getFolderByServerRelativePath(folderPath).select('Exists')();

      // Check if it's the system "Forms" folder
      if (folderExists.Exists && Title.toLowerCase() === 'forms') {
        const folderDetails = await this._sp.web.getFolderByServerRelativePath(folderPath).select('ItemCount')();

        // If it's the system "Forms" folder, do not overwrite
        if (folderDetails.ItemCount === 0) {
          alert("A system folder named 'Forms' already exists and cannot be overwritten. Please insert a different category title.");
          return;
        }
      }

      await this._sp.web.getFolderByServerRelativePath(libraryPath).folders.addUsingPath(Title, true);

      alert('Category added successfully!');
      this._getCategories(); // Refresh the list of categories
      this._getInteractiveChecklistInfo(); // Refresh the list of checklist
      this._onDismissEvent(); // Close the dialog
    } catch (error) {
      console.error('Error adding category:', error);
      alert('Error adding category. Please try again.');
    }
  };

  private async updateCategory() {
    const { Title } = this.state;
    if (!Title || Title.trim() == '') {
      alert('Category Title cannot be empty.');
      return;
    }

    try {
      // Get the folder's current URL
      const currentFolderUrl = this.state._EditItem.url;
      if (this.state._EditItem.name != this.state.Title) {
        // Extract the parent folder URL and construct the new folder URL
        const parentFolderUrl = currentFolderUrl.substring(0, currentFolderUrl.lastIndexOf('/'));
        const newFolderUrl = `${parentFolderUrl}/${this.state.Title}`;

        // Check if a folder with the new name already exists
        const folderExists = await this._sp.web.getFolderByServerRelativePath(newFolderUrl).select('Exists')();

        // Check if it's the system "Forms" folder
        if (folderExists.Exists && Title.toLowerCase() === 'forms') {
          const folderDetails = await this._sp.web.getFolderByServerRelativePath(newFolderUrl).select('ItemCount')();

          // If it's the system "Forms" folder, do not overwrite
          if (folderDetails.ItemCount === 0) {
            alert("A system folder named 'Forms' already exists and cannot be overwritten. Please insert a different category title.");
            return;
          }
        }

        // Move the folder to the new URL
        await this._sp.web.getFolderByServerRelativePath(currentFolderUrl).moveByPath(newFolderUrl);
      }

      alert('Category updated successfully!');
      this._getCategories(); // Refresh the list of categories
      this._getInteractiveChecklistInfo(); // Refresh the list of checklist
      this._onDismissEvent(); // Close the dialog
    } catch (error) {
      console.error('Error updating category:', error);
      alert('Error updating category. Please try again.');
    }
  };

  private async deleteCategory() {
    debugger;
    try {
      // Get the folder's current URL
      const currentFolderUrl = this.state._deleteItem.url;

      // Extract the parent folder URL and construct the new folder URL
      //const parentFolderUrl = currentFolderUrl.substring(0, currentFolderUrl.lastIndexOf('/'));
      //const newFolderUrl = `${parentFolderUrl}/${this.state.Title}`;

      // Delete the folder
      await this._sp.web.getFolderByServerRelativePath(currentFolderUrl).delete();

      alert('Category deleted successfully!');
      this._getCategories(); // Refresh the list of categories
      this._getInteractiveChecklistInfo(); // Refresh the list of checklist
      this.setState({ _showDialogDelete: !this.state._showDialogDelete });
    } catch (error) {
      console.error('Error deleted category:', error);
      alert('Error deleted category. Please try again.');
    }
  }

  private _toggleCategoryCollapse = (categoryName: string): void => {
    this.setState((prevState) => ({
      _collapsedCategories: {
        ...prevState._collapsedCategories,
        [categoryName]: !prevState._collapsedCategories[categoryName],
      },
    }));
  }

  public handleGlobalSearch = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    this.setState({ search: newValue || '' }, this.applyFilters);
  }

  public applyFilters = () => {
    const { search } = this.state;

    // if (search.length < 3) {
    //   //this.setState({ allcategoriesInfo: [] });
    //   return;
    // }
    const libraryPath = decodeURI(this.state.subSiteUrl) + "/InteractiveChecklists"; // Server-relative URL of your library
    this._searchAndBuildFolderHierarchy(libraryPath, search);
  }

  private async _searchAndBuildFolderHierarchy(folderUrl: string, searchTerm: string): Promise<any> {
    try {
      this.setState({ filteredItems: [] })
      let data = JSON.stringify(this.state.allcategoriesInfo);
      let result: any[] = [];
      for (let item of JSON.parse(data)) {
        if (item.name.toLowerCase().indexOf(searchTerm.toLowerCase()) > -1) {
          result.push(item);

          if (result.length == 0) {

            this.setState({ filteredItems: [] })
          }
          else {
            this.setState({ filteredItems: result })
          }
        } else {
          let clonnedItem: any = item;
          if (item.children && item.children.folders && item.children.folders.length > 0) {
            let matchingStage: any[] = [];
            for (let sitem of item.children.folders) {
              if (sitem.name.toLowerCase().indexOf(searchTerm.toLowerCase()) > -1) {
                matchingStage.push(sitem);
              } else {
                let clonnedStage: any = sitem;
                if (sitem.children && sitem.children.folders && sitem.children.folders.length > 0) {
                  let deliverable: any[] = [];
                  for (let ditem of sitem.children.folders) {
                    if (ditem.name.toLowerCase().indexOf(searchTerm.toLowerCase()) > -1) {
                      deliverable.push(ditem);
                    }
                  }
                  if (deliverable.length > 0) {
                    clonnedStage.children.folders = deliverable;
                    matchingStage.push(clonnedStage);
                  }
                  else {
                    if (result.length === 0) {

                      this.setState({ filteredItems: [] })
                    }
                    else {
                      this.setState({ filteredItems: result })
                    }
                  }
                }
              }
            }
            if (matchingStage.length > 0) {
              clonnedItem.children.folders = matchingStage;
              result.push(item);
              debugger;
              if (result.length === 0) {

                this.setState({ filteredItems: [] })
              }
              else {
                this.setState({ filteredItems: result })
              }
            }
          }
        }
      }
      // data = data.filter(x => x.name.toLowerCase().indexOf(searchTerm.toLowerCase()) > -1
      //   || x.children.folders.filter((y: any) => y.name.toLowerCase().indexOf(searchTerm.toLowerCase()) > -1));
      // debugger;

    } catch (error) {
      console.error("Error retrieving folder hierarchy:", error);
      throw error;
    }
  }

  onDragEnd(result: any) {
    const { destination, source } = result;

    // Dropped outside the list
    if (!destination) {
      return;
    }

    if (source.droppableId === destination.droppableId) {
      const categoryId = source.droppableId.split('-')[1];
      const stagesId = source.droppableId.split('-')[2];
      const deliverable = Array.from(this.state.filteredItems[categoryId].children.folders[stagesId].children.folders);

      // Rearrange the stages array
      const [movedStage] = deliverable.splice(source.index, 1);
      deliverable.splice(destination.index, 0, movedStage);

      // Update state with the new order
      const newFilteredItems = [...this.state.filteredItems];
      newFilteredItems[categoryId].children.folders[stagesId].children.folders = deliverable;

      this.setState({ filteredItems: newFilteredItems });

      // Call the update function
      const category = this.state.filteredItems[categoryId].name;
      const stageFolder = newFilteredItems[categoryId].children.folders[stagesId].name;
      this.updateSortOrder(category, stageFolder, deliverable);
    }
  }

  async updateSortOrder(category: string, stageFolder: string, deliverables: any[]) {
    const sp = spfi().using(SPFx(this.props.context));
    // const list = await sp.web.lists.getByTitle("InteractiveChecklists");
  
    for (let i = 0; i < deliverables.length; i++) {
      const deliverableFolderName = deliverables[i].name; // Assuming deliverables array contains folder names
  
      const folderUrl = `${this.props.WebServerRelativeURL}/InteractiveChecklists/${category}/${stageFolder}/${deliverableFolderName}`;
  
      // Get folder item
      const folderItem = await sp.web.getFolderByServerRelativePath(folderUrl).listItemAllFields();
      
      // Update sort order
      await sp.web.lists.getByTitle("InteractiveChecklists").items.getById(folderItem.Id).update({
        SortDeliverable: i + 1 // i+1 to start numbering from 1
      });
    }
  }
  

  public render(): React.ReactElement<IWpInteractiveChecklistProps> {
    const {
      hasTeamsContext,
    } = this.props;

    return (
      <section className={`${styles.wpInteractiveChecklist} ${hasTeamsContext ? styles.teams : ''}`}>
        <div style={{ display: 'inline-flex', marginBottom: '5%' }} className='col-md-12 col-sm-12 col-lg-12'>
          <div className='col-md-7 col-sm-7 col-lg-7'>
            <h3 style={{ fontFamily: 'Montserrat', fontStyle: 'normal', fontWeight: 700, fontSize: '24px', lineHeight: '29px', color: '#004693' }}>Interactive Checklists</h3>
          </div>
          <div className={`${styles.Searchnew} col-md-4 col-sm-4 col-lg-4`}>
            {/* <TextField
              className={styles.SearchnewTextBox}
              name='search'
              value={this.state.search}
              placeholder="Search deliverable by any category or any stage"
              onChange={this.handleGlobalSearch} /> */}
            <div className={styles.SearchWrapper}>
              <TextField
                className={styles.SearchnewTextBox}
                name='search'
                value={this.state.search}
                placeholder="Search deliverable by any category or any stage"
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
          </div>
          {this.state.isUserInGroup && (
            <div className={`${styles.News_oprt} col-md-1 col-sm-1 col-lg-1`}>
              <ul>
                <li><a onClick={() => { this._onDismissEvent() }} href='javascript:void(0);'><img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/plus-symbol.png"} width="20px" /></a></li>
              </ul>
            </div>
          )}
        </div>
        {this.state.search && this.state.search.length > 0 &&
          <p style={{ fontFamily: 'Montserrat', fontStyle: 'normal', fontWeight: '700', fontSize: '28px', lineHeight: '34px', color: '#004693' }}>
            Search result for {this.state.search}</p>}
        <div className='category'>
          {this.state.filteredItems && this.state.filteredItems.length > 0 &&

            this.state.filteredItems.map((nhItem: any, nhIndex: number) => (
              <div className='categories'>
                <div className='row'>
                  <div style={{ float: 'left', width: '1%', color: '#009bda', padding: '0px', marginLeft: '1%' }}><FontAwesomeIcon icon={faFolder} /></div>
                  <div className='col-md-10 col-sm-10 col-lg-10' style={{ float: 'left' }}>
                    <h3 style={{ fontFamily: 'Montserrat', fontStyle: 'normal', fontWeight: 700, fontSize: '20px', lineHeight: '24px', color: '#009BDA', float: 'left' }}>
                      <a className={`${styles.linkanchorStyle}`} href={`${this.props.WebServerRelativeURL}${this.props.filterPageURL}?Category=${nhItem.name}`}
                        target='_blank' style={{ textDecoration: "none" }}>{nhItem.name}
                      </a>
                    </h3>
                    <a onClick={() => this._toggleCategoryCollapse(nhItem.name)} href='javascript:void(0);' style={{ marginLeft: '10px', float: 'left', color: '#009BDA' }}>
                      <FontAwesomeIcon icon={this.state._collapsedCategories[nhItem.name] ? faChevronRight : faChevronDown} />
                    </a>
                    {this.state.isUserInGroup && (
                      <div>
                        <a onClick={() => { this._onDismissEventForStage(nhItem) }} href='javascript:void(0);' style={{ marginLeft: '3%' }}><img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/plus-symbol.png"} width="20px" /></a>
                      </div>
                    )}
                  </div>
                  {this.state.isUserInGroup && (
                    <div className={`${styles.Edit_del} col-md-1 col-sm-1 col-lg-1`}>
                      <a onClick={() => { this._onEditItem(nhItem) }} href='javascript:void(0);' style={{ marginRight: '15%' }}><FontAwesomeIcon icon={faPen} /></a>
                      <a onClick={() => { this._onDeleteItem(nhItem) }} href='javascript:void(0);'><FontAwesomeIcon icon={faTrashAlt} /></a>
                    </div>
                  )}
                </div>

                {!this.state._collapsedCategories[nhItem.name] && (
                  <div className={`${styles.Stage_card_section} col-md-12 col-lg-12 col-sm-12`}>
                    {
                      nhItem.children && nhItem.children.folders && nhItem.children.folders.length > 0 && nhItem.children.folders.map((stItem: any, stIndex: number) => (
                        <DragDropContext onDragEnd={this.onDragEnd}>
                          <Droppable droppableId={`droppable-${nhIndex}-${stIndex}`} type="DELIVERABLE">
                            {(provided) => (
                              <div ref={provided.innerRef} {...provided.droppableProps} className={`${styles.divItems}`}>
                                <h2 className={`${styles.headerName} col-md-10 col-lg-10 col-sm-10`}>
                                  <a className={`${styles.linkanchorStyle}`} href={`${this.props.WebServerRelativeURL}${this.props.filterPageURL}?Category=${nhItem.name}&Stage=${stItem.name}`}
                                    target='_blank' style={{ textDecoration: "none" }}>{stItem.name}
                                  </a>
                                </h2>
                                <div className={`${styles.Edit_del} col-md-2 col-sm-2 col-lg-2`}>
                                  <a style={{ marginRight: '10%' }} onClick={() => { this._downloadFile(stItem) }} href='javascript:void(0);'>
                                    <svg width="15" height="20" viewBox="0 0 22 26" fill="none" xmlns="http://www.w3.org/2000/svg">
                                      <path d="M9.59971 24.2H3.9997C2.4533 24.2 1.1997 22.9464 1.19971 21.4L1.19982 4.60003C1.19983 3.05364 2.45343 1.80005 3.99982 1.80005H16.6001C18.1465 1.80005 19.4001 3.05365 19.4001 4.60005V12.3M20.8001 21.3504L17.9368 24.2M17.9368 24.2L15.2001 21.4792M17.9368 24.2V17.2M6.10014 7.40005H14.5001M6.10014 11.6H14.5001M6.10014 15.8H10.3001" stroke="#00A3E0" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" />
                                    </svg>
                                  </a>
                                  {this.state.isUserInGroup && (
                                    <>
                                      <a onClick={() => { this._onEditItemStage(stItem); }} href='javascript:void(0);' style={{ marginRight: '10%' }}><FontAwesomeIcon icon={faPen} /></a>
                                      <a onClick={() => { this._onDeleteItemStage(stItem); }} href='javascript:void(0);'><FontAwesomeIcon icon={faTrashAlt} /></a>
                                    </>
                                  )}
                                </div>
                                <ol className={`${styles.listingolItems}`}>
                                  {
                                    stItem.children && stItem.children.folders && stItem.children.folders.length > 0 && stItem.children.folders.map((chItem: any, chIndex: number) => (
                                      <Draggable key={`draggable-${nhIndex}-${stIndex}-${chIndex}`} draggableId={`draggable-${nhIndex}-${stIndex}-${chIndex}`} index={chIndex}>
                                        {(provided, snapshot) => (
                                          <li ref={provided.innerRef} {...provided.draggableProps} {...provided.dragHandleProps} className={`${snapshot.isDragging ? styles.draggedliItems : styles.listingliItems}`}>
                                            <a className={`${styles.linkanchorStyle}`} href={`${this.props.WebServerRelativeURL}${this.props.filterPageURL}?Category=${nhItem.name}&Stage=${stItem.name}&DeliverableFolder=${chItem.name}`}
                                              target='_blank' style={{ textDecoration: "none" }}>{chItem.name}
                                            </a>
                                          </li>
                                        )}
                                      </Draggable>
                                    ))
                                  }
                                  {provided.placeholder}
                                </ol>
                              </div>
                            )
                            }
                          </Droppable>
                        </DragDropContext>
                      ))
                    }
                  </div>
                )}

              </div>
            ))
          }
          {this.state.filteredItems.length > 0}
        </div>

        <div>
          <Dialog
            hidden={!this.state._showDialogEvents}
            onDismiss={this._onDismissEvent}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: this.state.isItemInEditMode == true ? 'Edit Category' : 'Add New Category',
              subText: 'Provide the below information.'
            }}
            modalProps={{
              isBlocking: false,
              styles: { main: { minWidth: '520px !important' } }
            }}
          >
            <TextField label="Category Title" value={this.state.Title} onChange={(e, newValue) => this.setState({ Title: (newValue || '') })} />

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

            <Label>Are you sure want to delete this category ?</Label>

            <DialogFooter>
              <PrimaryButton onClick={() => this.deleteCategory()} text="Confirm" style={{ background: '#009bda', color: '#fff', border: 0 }} />
              {/* <PrimaryButton onClick={() => this.uploadImageAndAddItem()} text="Add Item" /> */}
              <DefaultButton onClick={() => { this._onDismissDelete() }} text="Cancel" />
            </DialogFooter>
          </Dialog>
        </div>


        <div>
          <Dialog
            hidden={!this.state._showDialogEventsForStage}
            onDismiss={this._onDismissEventForStageWrapper}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: this.state.isItemInEditMode == true ? 'Edit Stage' : 'Add New Stage',
              subText: 'Provide the below information.'
            }}
            modalProps={{
              isBlocking: false,
              styles: { main: { minWidth: '520px !important' } }
            }}
          >
            <TextField label="Stage Title" value={this.state.Title} onChange={(e, newValue) => this.setState({ Title: (newValue || '') })} />

            <Label>File Upload</Label>
            <input type="file" onChange={(event) => { this.handleFileChange(event) }} />

            {this.state.existingFileName && (
              <div>
                <Label>Existing File: {this.state.existingFileName}</Label>
              </div>
            )}

            <DialogFooter>
              <PrimaryButton onClick={() => this.saveItemStage()} text="Save" style={{ background: '#009bda', color: '#fff', border: 0 }} />
              {/* <PrimaryButton onClick={() => this.uploadImageAndAddItem()} text="Add Item" /> */}
              <DefaultButton onClick={() => { this._onDismissEventForStageWrapper() }} text="Cancel" />
            </DialogFooter>
          </Dialog>

          <Dialog
            hidden={!this.state._showDialogDeleteForStage}
            onDismiss={this._onDismissDeleteForStage}
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

            <Label>Are you sure want to delete this stage ?</Label>

            <DialogFooter>
              <PrimaryButton onClick={() => this.deleteStage()} text="Confirm" style={{ background: '#009bda', color: '#fff', border: 0 }} />
              {/* <PrimaryButton onClick={() => this.uploadImageAndAddItem()} text="Add Item" /> */}
              <DefaultButton onClick={() => { this._onDismissDeleteForStage() }} text="Cancel" />
            </DialogFooter>
          </Dialog>
        </div>
      </section>
    );
  }
}
