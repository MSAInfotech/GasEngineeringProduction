import * as React from 'react';
import styles from './WpTrainingTemplatesPolicies.module.scss';
import "bootstrap/dist/css/bootstrap.min.css";
import type { IWpTrainingTemplatesPoliciesProps } from './IWpTrainingTemplatesPoliciesProps';
import Pagination from '../components/Pagination';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  TextField,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Stack,
  PrimaryButton,
  DialogFooter,
  Dialog,
  DialogType,
  Label,
} from "@fluentui/react";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/fields";
import { spfi, SPFI, SPFx } from '@pnp/sp/presets/all';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPen, faTrashAlt, faSearch } from '@fortawesome/free-solid-svg-icons';
import { RotatingLines } from 'react-loader-spinner';

export interface IWpTrainingTemplatesPoliciesState {
  _isLoading: boolean;
  _allFiles: any[];
  _filtColsLeft: any[];
  _filteredFiles: any[];
  _columns: IColumn[];
  _filterValues: any[];
  _filterDelTitle: any;
  isUserInGroup: boolean;
  _showDialogEvents: boolean,
  _showDialogEventsEdit: boolean;
  showEditImageURL: boolean;
  saveButtonAdd: boolean;
  savButtonEdit: boolean;
  isItemInEditMode: boolean;
  _EditItem: any;
  _deleteItem: any;
  _showDialogDelete: boolean;
  _showDialogDeleteConfirm: boolean;
  _showDialogDeleteConfirmYes: boolean;
  _showDialogDeleteConfirmNo: boolean;
  categoryLstChoice: any;
  stageLstChoice: IDropdownOption[];
  category: string | number;
  deliverablefileToUpload: File[];
  policiesfileToUpload: File[];
  trainingfileToUpload: File[];
  examplefileToUpload: File[];
  Title: string,
  stage: string[];
  subSiteUrl: string,
  siteUrl: string,
  itemsPerPage: number;
  currentPage: number;
  allcategoriesInfo: any[];
  dropdownSets: { category: string; stage: string[]; }[];
  stageLists: IDropdownOption[][];
  filterSearch: string;
  categorySelected: any[];
  categoryOption: any;
  stageSelected: any[];
  stageOption: any;
  deliverableTitleSelected: any[];
  deliverableTitleOption: any;
  search: string;
  deliverabletitlefilterDrop: any;
  catfilterDrop: any;
  stagefilterDrop: any;
  rootUrl: string;
}

export default class WpTrainingTemplatesPolicies extends React.Component<IWpTrainingTemplatesPoliciesProps, IWpTrainingTemplatesPoliciesState> {
  public _columns: IColumn[];
  private _sp: SPFI;

  constructor(props: any) {
    super(props);
    sp: this._sp,
      //this._sp = spfi("https://sempra.sharepoint.com/sites/gasopscon/eng/").using(SPFx(this.props.context));
      this._sp = spfi().using(SPFx(this.props.context));
    this._columns = [
      {
        key: 'column1', name: 'Deliverable Title', fieldName: 'title', minWidth: 410, maxWidth: 410, isResizable: true,
        styles: { root: { color: '#75787B', fontStyle: 'normal', textAlign: 'center', marginTop: '6px' } },
        onRender: (item) => <span className={`${styles.customCellTitle}`}>{item.title}</span>,
      },
      {
        key: 'column2', name: 'Templates', fieldName: 'Templates', minWidth: 110, maxWidth: 130, isResizable: true,
        styles: { root: { color: '#75787B', fontStyle: 'normal', textAlign: 'center', marginTop: '6px' } },
        onRender: (item) => (
          <span className={`${styles.customCellTitle}`}>
            {item.Templates ? (
              <a onClick={() => { window.open(this.state.siteUrl.split('/sites/')[0] + item.Templates.url, '_blank') }} href="javascript:void(0);">
                <img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/Deliverables.png"} width="30px" />
              </a>
            ) : ''}
          </span>
        ),
      },
      {
        key: 'column3', name: 'Policies', fieldName: 'Policies', minWidth: 110, maxWidth: 130, isResizable: true,
        styles: { root: { color: '#75787B', fontStyle: 'normal', textAlign: 'center', marginTop: '6px' } },
        onRender: (item) => (
          <span className={`${styles.customCellTitle}`}>
            {item.Policies ? (
              <a onClick={() => { window.open(this.state.siteUrl.split('/sites/')[0] + item.Policies.url, '_blank') }} href="javascript:void(0);">
                <img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/Policies.png"} width="30px" />
              </a>
            ) : ''}
          </span>
        ),
      },
      {
        key: 'column4', name: 'Trainings', fieldName: 'Training', minWidth: 110, maxWidth: 130, isResizable: true,
        styles: { root: { color: '#75787B', fontStyle: 'normal', textAlign: 'center', marginTop: '6px' } },
        onRender: (item) => (
          <span className={`${styles.customCellTitle}`}>
            {item.Training ? (
              <a onClick={() => { window.open(this.state.siteUrl.split('/sites/')[0] + item.Training.url, '_blank') }} href="javascript:void(0);">
                <img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/Trainings.png"} width="30px" />
              </a>
            ) : ''}
          </span>
        ),
      },
      {
        key: 'column5', name: 'Examples', fieldName: 'Example', minWidth: 110, maxWidth: 130, isResizable: true,
        styles: { root: { color: '#75787B', fontStyle: 'normal', textAlign: 'center', marginTop: '6px' } },
        onRender: (item) => (
          <span className={`${styles.customCellTitle}`}>
            {item.Example ? (
              <a onClick={() => { window.open(this.state.siteUrl.split('/sites/')[0] + item.Example.url, '_blank') }} href="javascript:void(0);">
                <img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/Example.png"} width="30px" />
              </a>
            ) : ''}
          </span>
        ),
      },
      {
        key: 'column6', name: 'Category', fieldName: 'Category', minWidth: 110, maxWidth: 130, isResizable: true,
        styles: { root: { color: '#75787B', fontStyle: 'normal', textAlign: 'center', marginTop: '6px' } },
        onRender: (item) => <span className={`${styles.customCell}`}>{item.Category}</span>,
      },
      {
        key: 'column7', name: 'CDM Stage Gate', fieldName: 'Stage', minWidth: 130, maxWidth: 150, isResizable: true,
        styles: { root: { color: '#75787B', fontStyle: 'normal', textAlign: 'center', marginTop: '6px' } },
        onRender: (item) => <span className={`${styles.customCell}`}>{item.Stage}</span>,
      },
      {
        key: 'column8', name: 'Actions', minWidth: 100, maxWidth: 200, isResizable: true,
        styles: { root: { color: '#75787B', fontStyle: 'normal', textAlign: 'center', marginTop: '6px' } },
        onRender: (item) => (this.state.isUserInGroup && (<span className={`${styles.customCell}`}>{item.Action}</span> &&
          <div className={`${styles.Edit_del} col-lg-2 col-md-2 col-sm-2`}>
            <Stack horizontal tokens={{ childrenGap: 0 }}>
              <a
                onClick={() => this._onEditItem(item)}
                title='Edit Deliverable'
                href='javascript:void(0);'
                style={{ marginRight: '15px', color: '#009bda', fontSize: '20px', alignContent: 'center' }}
              >
                <FontAwesomeIcon icon={faPen} />
              </a>
              <a
                onClick={() => this._onDeleteItem(item)}
                title='Delete Deliverable'
                href='javascript:void(0);'
                style={{ color: '#ffff', fontSize: '20px', backgroundColor: '#DC3B0C', alignContent: 'center', padding: '0% 80%', borderRadius: '50%' }}
              >
                <FontAwesomeIcon icon={faTrashAlt} />
              </a>
            </Stack>
          </div>
        )
        ),
      },
    ]
    this.state = {
      _isLoading: true,
      _allFiles: [],
      _filtColsLeft: [],
      _filteredFiles: [],
      _columns: [],
      _filterValues: [],
      _filterDelTitle: "",
      isUserInGroup: false,
      _showDialogEvents: false,
      _showDialogEventsEdit: false,
      showEditImageURL: false,
      saveButtonAdd: false,
      savButtonEdit: false,
      isItemInEditMode: false,
      _EditItem: {},
      _deleteItem: {},
      _showDialogDelete: false,
      _showDialogDeleteConfirm: false,
      _showDialogDeleteConfirmYes: false,
      _showDialogDeleteConfirmNo: false,
      categoryLstChoice: [],
      stageLstChoice: [],
      category: '',
      stage: [],
      deliverablefileToUpload: [],
      policiesfileToUpload: [],
      trainingfileToUpload: [],
      examplefileToUpload: [],
      Title: "",
      siteUrl: "",
      subSiteUrl: "/sites/gasopscon/eng/resource%20hub",
      itemsPerPage: 10,
      currentPage: 1,
      allcategoriesInfo: [],
      dropdownSets: [], // Initialize with one set of dropdowns
      stageLists: [[]], // Initialize stage lists for each dropdown set
      filterSearch: '',
      categorySelected: [],
      categoryOption: [],
      stageSelected: [],
      stageOption: [],
      deliverableTitleSelected: [],
      deliverableTitleOption: [],
      search: '',
      deliverabletitlefilterDrop: [],
      catfilterDrop: [],
      stagefilterDrop: [],
      rootUrl: '',
    }
  }

  public componentDidMount(): void {
    this.setState({ _isLoading: true });
    this.checkUserInGroup("PortalAdmins");
    this.getDeliverableFolders();
  }

  private async getDeliverableFolders(): Promise<string> {
    debugger;
    const libraryPath = decodeURI(this.state.subSiteUrl) + "/InteractiveChecklists"; // Server-relative URL of your library
    let _catfil: any[] = [];
    let _stagefil: any[] = [];
    let _deliverableTitlefil: any[] = [];
    let filteredCatArr: any[] = [];
    let filteredStageArr: any[] = [];
    let filteredDeliverableTitleArr: any[] = [];
    let counter: number = 1;
    let data = await this._getFolderHierarchy(libraryPath, counter);
    let result: any[] = [];
    for (let category of data.folders) {
      if (category.children && category.children.folders && category.children.folders.length > 0) {
        _catfil.push(category.name);
        filteredCatArr = Array.from(new Set(_catfil));
        for (let stage of category.children.folders) {
          if (stage.children && stage.children.folders && stage.children.folders.length > 0) {
            _stagefil.push(stage.name);
            filteredStageArr = Array.from(new Set(_stagefil));
            for (let deliverable of stage.children.folders) {
              _deliverableTitlefil.push(deliverable.name);
              filteredDeliverableTitleArr = Array.from(new Set(_deliverableTitlefil));
              let resultItem: any = {};
              resultItem.title = deliverable.name;
              resultItem.Category = category.name;
              resultItem.Stage = stage.name;
              resultItem.url = deliverable.url;
              resultItem.timeCreated = deliverable.timeCreated;
              if (deliverable.children && deliverable.children.folders && deliverable.children.folders.length > 0) {
                for (let deliItem of deliverable.children.folders) {
                  resultItem[deliItem.name] = deliItem;
                }
              }
              result.push(resultItem);
            }
          }
        }
      }
    }

    console.log(result);
    if (result) {
      result.sort((a, b) => new Date(b.timeCreated).getTime() - new Date(a.timeCreated).getTime());
      this.setState({ _isLoading: false, _allFiles: result, _filteredFiles: result, catfilterDrop: filteredCatArr, deliverabletitlefilterDrop: filteredDeliverableTitleArr, stagefilterDrop: filteredStageArr });
      this.bindCategoryDropdown();

      const urlParams = new URLSearchParams(window.location.search);
      const category = urlParams.get('Category');
      const stage = urlParams.get('Stage');
      const deliverableFolder = urlParams.get('DeliverableFolder');
      if (category) {
        const categorySelected = true
          ? [...this.state.categorySelected, category]
          : this.state.categorySelected.filter(key => key !== category);
        this.setState({ categorySelected }, this.applyFilters);
      }
      if (stage) {
        const stageSelected = true
          ? [...this.state.stageSelected, stage]
          : this.state.stageSelected.filter(key => key !== stage);
        this.setState({ stageSelected }, this.applyFilters);
      }
      if (deliverableFolder) {
        const deliverableTitleSelected = true
          ? [...this.state.deliverableTitleSelected, deliverableFolder]
          : this.state.deliverableTitleSelected.filter(key => key !== deliverableFolder);
        this.setState({ deliverableTitleSelected }, this.applyFilters);
      }
    }
    return Promise.resolve("")
  }

  public async bindCategoryDropdown() {
    let _cat: any[] = [];
    let _delTitle: any[] = [];
    let _stage: any[] = [];
    this.state.catfilterDrop.map((item: any) => {
      _cat.push({ key: item, text: item });
    });
    this.state.deliverabletitlefilterDrop.map((item: any) => {
      _delTitle.push({ key: item, text: item });
    });
    this.state.stagefilterDrop.map((item: any) => {
      _stage.push({ key: item, text: item });
    });
    const catMulOption: IDropdownOption[] = _cat;
    const stageMulOption: IDropdownOption[] = _stage;
    const delTitleMulOption: IDropdownOption[] = _delTitle;
    this.setState({ categoryOption: catMulOption, stageOption: stageMulOption, deliverableTitleOption: delTitleMulOption });
    this.getCategoryChoiceFields();
  }

  private async _getFolderHierarchy(folderUrl: any, counter: number): Promise<any> {
    try {
      const folder = this._sp.web.getFolderByServerRelativePath(folderUrl);

      const files = await folder.files();
      const subfolders = await folder.folders();

      const fileList = files.map(file => ({ fileName: file.Name, fileUrl: file.ServerRelativeUrl }));
      const folderList = subfolders.map(subfolder => ({
        name: subfolder.Name,
        url: subfolder.ServerRelativeUrl,
        timeCreated: subfolder.TimeCreated
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

  private async checkUserInGroup(groupName: string): Promise<void> {
    this.setState({
      siteUrl: (await this._sp.web.getContextInfo()).SiteFullUrl
    });

    const currentUser = await this._sp.web.currentUser();
    const groups = await this._sp.web.siteGroups.getByName(groupName).users().then((response) => {
      const isUserInGroupExists = response.some(user => user.Id === currentUser.Id);
      this.setState({ isUserInGroup: isUserInGroupExists });
      console.log("Is User in Gropup = ", isUserInGroupExists);
      if (!isUserInGroupExists) {
        this._columns = this._columns.filter(column => column.key !== 'column8');
      }
    });
    console.log(groups);
  }

  public _onDismissEvent(): void {
    this.setState({
      _showDialogEvents: !this.state._showDialogEvents,
      showEditImageURL: false,
      saveButtonAdd: true,
      savButtonEdit: false,
      Title: "",
      deliverablefileToUpload: [],
      policiesfileToUpload: [],
      trainingfileToUpload: [],
      examplefileToUpload: [],
      category: "",
      stage: [],
      dropdownSets: [],
      stageLstChoice: [],
    });
  }

  public _onDismissEventEdit(): void {
    this.setState({
      _showDialogEventsEdit: !this.state._showDialogEventsEdit,
      showEditImageURL: false,
      saveButtonAdd: true,
      savButtonEdit: false,
      Title: "",
      deliverablefileToUpload: [],
      policiesfileToUpload: [],
      trainingfileToUpload: [],
      examplefileToUpload: [],
      category: "",
      stage: [],
      stageLstChoice: [],
    });
  }

  public async _onEditItem(tmpEditItem: any): Promise<void> {
    //const fileAddResult = this._sp.web.getFolderByServerRelativePath(tmpEditItem.ServerRelativeUrl).files;
    const list = await this._sp.web.lists.getByTitle("InteractiveChecklists");
    const rootFolder = await list.rootFolder();
    const rootFolderPath = rootFolder.ServerRelativeUrl;
    this.setState({
      _showDialogEventsEdit: !this.state._showDialogEventsEdit,
      showEditImageURL: tmpEditItem.ImageURL ? true : false,
      Title: tmpEditItem.title,
      deliverablefileToUpload: tmpEditItem.Deliverable,
      policiesfileToUpload: tmpEditItem.Policies,
      trainingfileToUpload: tmpEditItem.Training,
      examplefileToUpload: tmpEditItem.Example,
      category: tmpEditItem.Category,
      stage: tmpEditItem.Stage,
      saveButtonAdd: false,
      savButtonEdit: true,
      isItemInEditMode: true,
      _EditItem: tmpEditItem,
      rootUrl: rootFolderPath
    });
    this.getStagesForCategory(tmpEditItem.Category);
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

  public async getCategoryChoiceFields() {
    let _catchoice: any = [];
    const list = await this._sp.web.lists.getByTitle("InteractiveChecklists");

    const rootFolder = await list.rootFolder();
    const rootFolderPath = rootFolder.ServerRelativeUrl;

    // Get all top-level folders in the document library
    const folders = await list.items.filter(`FSObjType eq 1 and FileDirRef eq '${rootFolderPath}'`).select('Title', 'FileLeafRef')();


    //const folders = await list.items.filter('FSObjType eq 1').select('Title', 'FileLeafRef')();
    folders.map((item: any) => {
      _catchoice.push({ key: item.FileLeafRef, text: item.FileLeafRef });
    });
    const categoryOption: IDropdownOption[] = _catchoice;
    this.setState({ categoryLstChoice: categoryOption });
  }

  public _onDismissDelete(): void {
    this.setState({ _showDialogDelete: !this.state._showDialogDelete });
  }

  public changeCategory = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    console.log('here is the things updating...' + item.key + ' ' + item.text);
    if (this.state._showDialogEvents) {
      this.setState({ stage: [] })
    }
    this.setState({ category: item.key, stageLstChoice: [] });

    this.getStagesForCategory(item.key as string);
  }

  public async getStagesForCategory(category: string) {
    let _catchoice: any = [];
    const list = await this._sp.web.lists.getByTitle("InteractiveChecklists");

    const rootFolder = await list.rootFolder();
    const rootFolderPath = rootFolder.ServerRelativeUrl;

    // Construct the folder path for the selected category
    const stageFolderPath = `${rootFolderPath}/${category}`;

    // Get folders within the selected category folder
    const folders = await list.items.filter(`FSObjType eq 1 and FileDirRef eq '${stageFolderPath}'`).select('Title', 'FileLeafRef')();

    _catchoice = folders.map((item: any) => ({
      key: item.FileLeafRef,
      text: item.FileLeafRef,
    }));

    const stageOption: IDropdownOption[] = _catchoice;
    this.setState({ stageLstChoice: stageOption });
  }

  public changeStage = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, isMultiSelect?: boolean): void => {
    debugger;
    if (option) {
      const selectedStages = this.state.stage;
      const selectedKey = option.key as string;
      if (isMultiSelect) {
        if (option.selected) {
          // Add selected stage to the state
          if (selectedStages.indexOf(selectedKey) === -1) {
            this.setState({ stage: [...selectedStages, selectedKey] });
          }
        } else {
          // Remove deselected stage from the state
          this.setState({ stage: selectedStages.filter(key => key !== selectedKey) });
        }
      } else {
        this.setState({ stage: [selectedKey] });
      }
    }
  }

  public async handleFileChangeForDeliverable(event: React.ChangeEvent<HTMLInputElement>) {
    if (event.target.files && event.target.files.length > 0) {
      this.setState({ deliverablefileToUpload: Array.from(event.target.files) });
    }
  }

  public async handleFileChangeForPolicies(event: React.ChangeEvent<HTMLInputElement>) {
    if (event.target.files && event.target.files.length > 0) {
      this.setState({ policiesfileToUpload: Array.from(event.target.files) });
    }
  }

  public async handleFileChangeForTraining(event: React.ChangeEvent<HTMLInputElement>) {
    if (event.target.files && event.target.files.length > 0) {
      this.setState({ trainingfileToUpload: Array.from(event.target.files) });
    }
  }

  public async handleFileChangeForExample(event: React.ChangeEvent<HTMLInputElement>) {
    if (event.target.files && event.target.files.length > 0) {
      this.setState({ examplefileToUpload: Array.from(event.target.files) });
    }
  }

  private saveItem() {
    if (this.state.saveButtonAdd) {
      this.addDeliverable();
    }
    else {
      this.updateDeliverable();
    }
  }

  public async addDeliverable() {
    const { category, stage, Title, dropdownSets, deliverablefileToUpload, policiesfileToUpload, trainingfileToUpload, examplefileToUpload } = this.state;

    if (!Title) {
      alert('Deliverable Title cannot be empty.');
      return;
    }
    if (!category) {
      alert('Deliverable Category cannot be empty.');
      return;
    }
    if (!stage || (Array.isArray(stage) && stage.length === 0)) {
      alert('Deliverable Stage cannot be empty.');
      return;
    }

    // Validate each dynamically added dropdown set
    for (let i = 0; i < dropdownSets.length; i++) {
      const { category, stage } = dropdownSets[i];

      if (!category) {
        alert(`Deliverable Category cannot be empty for set ${i + 1}.`);
        return;
      }
      if (!stage || (Array.isArray(stage) && stage.length === 0)) {
        alert(`Deliverable Stage cannot be empty for set ${i + 1}.`);
        return;
      }
    }

    try {
      const list = await this._sp.web.lists.getByTitle("InteractiveChecklists");
      const rootFolder = await list.rootFolder();
      const rootFolderPath = rootFolder.ServerRelativeUrl;

      // Create folder under category folder for each selected stage
      let folderCreationPromises = stage.map(async (stageFolder) => {
        const stageFolderPath = `${rootFolderPath}/${category}/${stageFolder}/${Title}`;

        // Get all folder from ${rootFolderPath}/${category}/${stageFolder}
        let folderUrl = `${rootFolderPath}/${category}/${stageFolder}`;
        const folder = this._sp.web.getFolderByServerRelativePath(folderUrl);
        let sfolders: any[] = await folder.folders.select("Name", "ServerRelativeUrl", "ListItemAllFields/SortDeliverable").expand("ListItemAllFields")();
        let maxNumber: number = 0;
        if (sfolders.length > 0) {
          sfolders = sfolders.sort((a, b) => b.ListItemAllFields?.SortDeliverable - a.ListItemAllFields?.SortDeliverable);
          maxNumber = sfolders[0].ListItemAllFields?.SortDeliverable;
        }

        await this.createFolder(stageFolderPath);

        const folderItem = await this._sp.web.getFolderByServerRelativePath(stageFolderPath).listItemAllFields();
        await this._sp.web.lists.getByTitle("InteractiveChecklists").items.getById(folderItem.Id).update({
          SortDeliverable: maxNumber + 1
        });

        // Create subfolders
        const subfolders = ["Templates", "Policies", "Training", "Example"];
        for (const subfolder of subfolders) {
          await this.createFolder(`${stageFolderPath}/${subfolder}`);
        }

        return stageFolderPath;
      });

      let createdStageFolders = await Promise.all(folderCreationPromises);

      // Log the created folder paths for debugging
      console.log('Created Stage Folders:', createdStageFolders);

      // Upload files to the deliverable folder under each stage folder
      let fileUploadPromises: Promise<void>[] = createdStageFolders.reduce((acc: Promise<void>[], stageFolderPath) => {
        let deliverablePromises = deliverablefileToUpload.map((file) => this.uploadFileToFolder(`${stageFolderPath}/Templates`, file));
        let policiesPromises = policiesfileToUpload.map((file) => this.uploadFileToFolder(`${stageFolderPath}/Policies`, file));
        let trainingPromises = trainingfileToUpload.map((file) => this.uploadFileToFolder(`${stageFolderPath}/Training`, file));
        let examplePromises = examplefileToUpload.map((file) => this.uploadFileToFolder(`${stageFolderPath}/Example`, file));

        return acc.concat(deliverablePromises, policiesPromises, trainingPromises, examplePromises);
      }, []);


      // Log the number of upload promises for debugging
      console.log('Number of File Upload Promises:', fileUploadPromises.length);

      await Promise.all(fileUploadPromises);

      // Create folders for each dynamically added dropdown set
      for (let i = 0; i < dropdownSets.length; i++) {
        const { category, stage } = dropdownSets[i];

        let folderCreationPromises = stage.map(async (stageFolder) => {
          const stageFolderPath = `${rootFolderPath}/${category}/${stageFolder}/${Title}`;

          // Get all folder from ${rootFolderPath}/${category}/${stageFolder}
          let folderUrl = `${rootFolderPath}/${category}/${stageFolder}`;
          const folder = this._sp.web.getFolderByServerRelativePath(folderUrl);
          let sfolders: any[] = await folder.folders.select("Name", "ServerRelativeUrl", "ListItemAllFields/SortDeliverable").expand("ListItemAllFields")();
          let maxNumber: number = 0;
          if (sfolders.length > 0) {
            sfolders = sfolders.sort((a, b) => b.ListItemAllFields?.SortDeliverable - a.ListItemAllFields?.SortDeliverable);
            maxNumber = sfolders[0].ListItemAllFields?.SortDeliverable;
          }

          await this.createFolder(stageFolderPath);

          const folderItem = await this._sp.web.getFolderByServerRelativePath(stageFolderPath).listItemAllFields();
          await this._sp.web.lists.getByTitle("InteractiveChecklists").items.getById(folderItem.Id).update({
            SortDeliverable: maxNumber + 1
          });

          // Create subfolders
          const subfolders = ["Templates", "Policies", "Training", "Example"];
          for (const subfolder of subfolders) {
            await this.createFolder(`${stageFolderPath}/${subfolder}`);
          }

          return stageFolderPath;
        });

        let createdStageFolders = await Promise.all(folderCreationPromises);

        // Log the created folder paths for debugging
        console.log('Created Stage Folders:', createdStageFolders);

        // Upload files to the deliverable folder under each stage folder
        let fileUploadPromises: Promise<void>[] = createdStageFolders.reduce((acc: Promise<void>[], stageFolderPath) => {
          let deliverablePromises = deliverablefileToUpload.map((file) => this.uploadFileToFolder(`${stageFolderPath}/Templates`, file));
          let policiesPromises = policiesfileToUpload.map((file) => this.uploadFileToFolder(`${stageFolderPath}/Policies`, file));
          let trainingPromises = trainingfileToUpload.map((file) => this.uploadFileToFolder(`${stageFolderPath}/Training`, file));
          let examplePromises = examplefileToUpload.map((file) => this.uploadFileToFolder(`${stageFolderPath}/Example`, file));

          return acc.concat(deliverablePromises, policiesPromises, trainingPromises, examplePromises);
        }, []);

        // Log the number of upload promises for debugging
        console.log('Number of File Upload Promises:', fileUploadPromises.length);

        await Promise.all(fileUploadPromises);
      }

      alert('Deliverable added successfully!');
      this.getDeliverableFolders(); // Refresh the list of deliverables
      this._onDismissEvent();

    } catch (error) {
      console.error('Error adding Deliverable:', error);
      alert('Error adding Deliverable. Please try again.');
    }
  }

  // Method to create a folder
  private async createFolder(folderPath: string): Promise<void> {
    try {
      await this._sp.web.folders.addUsingPath(folderPath, true);
      console.log(`Folder created at ${folderPath}`);
    } catch (error) {
      console.error(`Error creating folder at ${folderPath}:`, error);
    }
  }

  public async updateDeliverable() {
    const { category, stage, Title, _EditItem } = this.state;
    if (!Title) {
      alert('Category Title cannot be empty.');
      return;
    }
    if (!category) {
      alert('Deliverable Category cannot be empty.');
      return;
    }
    if (!stage) {
      alert('Deliverable Stage cannot be empty.');
      return;
    }

    try {
      const list = await this._sp.web.lists.getByTitle("InteractiveChecklists");
      const rootFolder = await list.rootFolder();
      const rootFolderPath = rootFolder.ServerRelativeUrl;

      // Construct the folder path for the selected category and stage
      const newFolderPath = `${rootFolderPath}/${category}/${stage}/${Title}`;
      const oldFolderPath = `${rootFolderPath}/${_EditItem.Category}/${_EditItem.Stage}/${_EditItem.title}`;
      let moveFolder: boolean = false;

      // Rename the folder if the title has changed
      if (_EditItem.Category !== category || _EditItem.Stage !== stage) {
        moveFolder = true;
      }

      // Rename the folder if the title has changed
      let renameDeliverable: boolean = false;
      if (_EditItem.title !== Title) {
        renameDeliverable = true;
      }

      if (renameDeliverable) {
        await this.renameFolder(oldFolderPath, newFolderPath);
      }
      if (moveFolder) {
        await this.renameFolder(oldFolderPath, newFolderPath);
      }

      // Get all folder from ${rootFolderPath}/${category}/${stageFolder}
      let folderUrl = `${rootFolderPath}/${category}/${stage}`;
      const folder = this._sp.web.getFolderByServerRelativePath(folderUrl);
      let sfolders: any[] = await folder.folders.select("Name", "ServerRelativeUrl", "ListItemAllFields/SortDeliverable").expand("ListItemAllFields")();
      let maxNumber: number = 0;
      if (sfolders.length > 0) {
        sfolders = sfolders.sort((a, b) => b.ListItemAllFields?.SortDeliverable - a.ListItemAllFields?.SortDeliverable);
        maxNumber = sfolders[0].ListItemAllFields?.SortDeliverable;
      }

      const folderItem = await this._sp.web.getFolderByServerRelativePath(newFolderPath).listItemAllFields();
      await this._sp.web.lists.getByTitle("InteractiveChecklists").items.getById(folderItem.Id).update({
        SortDeliverable: maxNumber + 1
      });

      alert('Deliverable updated successfully!');
      this.getDeliverableFolders(); // Refresh the list of deliverables
      this._onDismissEventEdit();

    } catch (error) {
      console.error('Error adding stage:', error);
      alert('Error adding stage. Please try again.');
    }
  }

  // Helper method to rename a folder
  private async renameFolder(oldFolderPath: string, newFolderPath: string): Promise<void> {
    const oldFolder = this._sp.web.getFolderByServerRelativePath(oldFolderPath);
    await oldFolder.moveByPath(newFolderPath, false);
    console.log(`Folder renamed from ${oldFolderPath} to ${newFolderPath}`);
  }

  // Helper method to upload file to a specific folder
  public async uploadFileToFolder(folderPath: string, file: File): Promise<void> {
    const folder = this._sp.web.getFolderByServerRelativePath(folderPath);
    const fileName = file.name;

    // Upload the file to the specified folder
    try {
      await folder.files.addUsingPath(`${folderPath}/${fileName}`, file, { Overwrite: true });
      console.log(`File uploaded to ${folderPath}`);
    } catch (error) {
      console.error(`Error uploading file to ${folderPath}:`, error);
    }
  }

  public async deleteDeliverable() {
    debugger;
    try {
      const currentFolderUrl = this.state._deleteItem.url;
      // Delete the folder
      await this._sp.web.getFolderByServerRelativePath(currentFolderUrl).delete();

      alert('Deliverable deleted successfully!');
      this.getDeliverableFolders(); // Refresh the list of deliverables
      this.setState({ _showDialogDelete: !this.state._showDialogDelete });
    } catch (error) {
      console.error('Error deleted Deliverable:', error);
      alert('Error deleted Deliverable. Please try again.');
    }
  }

  public handlePageChange = (page: number) => {
    this.setState({ currentPage: page });
  }

  addDropdownSet = () => {
    this.setState(prevState => ({
      dropdownSets: [...prevState.dropdownSets, { category: '', stage: [], title: '' }],
      stageLists: [...prevState.stageLists, []]
    }));
  };


  handleCategoryChange = async (index: number, event: React.FormEvent<HTMLDivElement> | undefined, option: IDropdownOption | undefined) => {
    const newDropdownSets = [...this.state.dropdownSets];
    if (option) {
      newDropdownSets[index].category = option.key as string;
      newDropdownSets[index].stage = [];  // Reset the selected stages
      this.setState({ dropdownSets: newDropdownSets, stageLists: [[]] });

      // Fetch stages for the selected category
      const stages = await this.getStages(option.key as string);
      // Update the state for stages specific to the dropdown set
      const updatedStageLists = [...this.state.stageLists];
      updatedStageLists[index] = stages;
      this.setState({ stageLists: updatedStageLists });
    }
  };

  public async getStages(category: string): Promise<IDropdownOption[]> {
    let _catchoice: IDropdownOption[] = [];
    const list = await this._sp.web.lists.getByTitle("InteractiveChecklists");

    const rootFolder = await list.rootFolder();
    const rootFolderPath = rootFolder.ServerRelativeUrl;

    // Construct the folder path for the selected category
    const stageFolderPath = `${rootFolderPath}/${category}`;

    // Get folders within the selected category folder
    const folders = await list.items.filter(`FSObjType eq 1 and FileDirRef eq '${stageFolderPath}'`).select('Title', 'FileLeafRef')();

    _catchoice = folders.map((item: any) => ({
      key: item.FileLeafRef,
      text: item.FileLeafRef,
    }));

    return _catchoice;

  }

  handleStageChange = (index: number, event: React.FormEvent<HTMLDivElement> | undefined, options: IDropdownOption[] | undefined) => {
    const newDropdownSets = [...this.state.dropdownSets];
    if (options) {
      let updatedOption = options[0];
      let selectedStage: any = [...newDropdownSets[index].stage];
      if (updatedOption && updatedOption.selected) {
        let key: string = updatedOption.key as string;
        if (selectedStage.indexOf(key) == -1) {
          let updatedStage: string[] = [...newDropdownSets[index].stage, key];
          newDropdownSets[index].stage = updatedStage;
          this.setState({ dropdownSets: newDropdownSets });
        }
      } else {
        let key: string = updatedOption.key as string;
        let updatedStage: string[] = selectedStage.filter((x: any) => x !== key);
        newDropdownSets[index].stage = updatedStage;
        this.setState({ dropdownSets: newDropdownSets });
      }
    }
  };

  public handleFilterSearch = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    this.setState({ filterSearch: newValue || '' }, this.applyFilterSearch);
  }

  public handleDeliverableTitle = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption): void => {
    const deliverableTitleSelected = option.selected
      ? [...this.state.deliverableTitleSelected, option.key]
      : this.state.deliverableTitleSelected.filter(key => key !== option.key);
    this.setState({ deliverableTitleSelected, currentPage: 1 }, this.applyFilters);
  }

  public handleCategory = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption): void => {
    const categorySelected = option.selected
      ? [...this.state.categorySelected, option.key]
      : this.state.categorySelected.filter(key => key !== option.key);
    this.setState({ categorySelected, currentPage: 1 }, this.applyFilters);
  }

  public handleStage = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption): void => {
    const stageSelected = option.selected
      ? [...this.state.stageSelected, option.key]
      : this.state.stageSelected.filter(key => key !== option.key);
    this.setState({ stageSelected, currentPage: 1 }, this.applyFilters);
  }

  public clearFilters = () => {
    this.setState({
      filterSearch: '',
      deliverableTitleSelected: [],
      categorySelected: [],
      stageSelected: [],
      currentPage: 1,
      _filteredFiles: this.state._allFiles
    }, this.applyFilters);
    this.bindCategoryDropdown();
  }
  public applyFilters = () => {
    const { _allFiles, deliverableTitleSelected, categorySelected, stageSelected } = this.state;
    const _filteredFiles = _allFiles.filter((item: { Stage: string; Category: string; title: string; }) => {
      const matchesDeliverabletitle = !deliverableTitleSelected.length || deliverableTitleSelected.some(deliverableTitle => deliverableTitle === item.title);
      const matchesCategories = !categorySelected.length || categorySelected.some(category => category === item.Category);
      const matchesStages = !stageSelected.length || stageSelected.some(stage => stage === item.Stage);
      return matchesDeliverabletitle && matchesCategories && matchesStages;
    });
    this.setState({ _filteredFiles });
  }
  public applyFilterSearch = () => {
    const { deliverabletitlefilterDrop, catfilterDrop, stagefilterDrop, filterSearch } = this.state;
    const filteredDeliverableTitle: any[] = deliverabletitlefilterDrop
      .filter((x: any) => x && x.toLowerCase().includes(filterSearch.toLowerCase()))
      .map((deliverableTitle: string) => ({ key: deliverableTitle, text: deliverableTitle }));

    const filteredCategories: any[] = catfilterDrop
      .filter((x: any) => x && x.toLowerCase().includes(filterSearch.toLowerCase()))
      .map((category: string) => ({ key: category, text: category }));

    const filteredStages: any[] = stagefilterDrop
      .filter((x: any) => x && x.toLowerCase().includes(filterSearch.toLowerCase()))
      .map((stage: string) => ({ key: stage, text: stage }));

    const catMulOption: IDropdownOption[] = filteredCategories;
    const deliverableTitleMulOption: IDropdownOption[] = filteredDeliverableTitle;
    const stageMulOption: IDropdownOption[] = filteredStages;
    this.setState({
      categoryOption: catMulOption,
      stageOption: stageMulOption,
      deliverableTitleOption: deliverableTitleMulOption
    });
  }


  public render(): React.ReactElement<IWpTrainingTemplatesPoliciesProps> {
    const { _filteredFiles, currentPage, itemsPerPage } = this.state;
    const indexOfLastItem = currentPage * itemsPerPage;
    const indexOfFirstItem = indexOfLastItem - itemsPerPage;
    const currentItems = _filteredFiles.slice(indexOfFirstItem, indexOfLastItem);
    const totalPages = Math.ceil(_filteredFiles.length / itemsPerPage);
    return (
      <section>
        {
          this.state._isLoading == true &&
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

        {
          this.state._isLoading == false && (
            <div className={`${styles.wpTrainingTemplatesPolicies}`}>
              {this.state.isUserInGroup && (
                <div className={`${styles.plusIc}`}>
                  <a onClick={() => { this._onDismissEvent() }}
                    title='Add New Deliverable'
                    href='javascript:void(0);'>
                    <img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/plus-symbol.png"} width="20px" />
                  </a>
                </div>
              )}
              <div className={`${styles.desingcard} col-lg-12 col-md-12 col-sm-12`}>
                <div className={`${styles.filterLeft} col-lg-2 col-md-2 col-sm-2`}>
                  <Label className={styles.filterLable_filter}>Filter by</Label>
                  <div className={styles.searchFilterContainer}>
                    < TextField
                      className={styles.SearchnewTextBox}
                      name='search'
                      value={this.state.filterSearch}
                      placeholder="Search"
                      onChange={this.handleFilterSearch} />
                    <FontAwesomeIcon icon={faSearch} className={styles.searchIcon} />
                  </div>
                  <div className={styles.clearbtn}>
                    <DefaultButton text="Clear All Filters" onClick={() => this.clearFilters()} />
                  </div>
                  <Label className={styles.filterLable}>Deliverable Title</Label>
                  <Dropdown
                    style={{ width: "100% !important" }}
                    placeholder="Select Deliverable"
                    selectedKeys={this.state.deliverableTitleSelected}
                    onChange={this.handleDeliverableTitle}
                    multiSelect
                    options={this.state.deliverableTitleOption}
                  />
                  <Label className={styles.filterLable}>Category</Label>
                  <Dropdown
                    style={{ width: "100% !important" }}
                    placeholder="Select Category"
                    selectedKeys={this.state.categorySelected}
                    onChange={this.handleCategory}
                    multiSelect
                    options={this.state.categoryOption}
                  />
                  <Label className={styles.filterLable}>CDM Stage Gate</Label>
                  <Dropdown
                    style={{ width: "100% !important" }}
                    placeholder="Select Stage"
                    selectedKeys={this.state.stageSelected}
                    onChange={this.handleStage}
                    multiSelect
                    options={this.state.stageOption}
                  />
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
                    title: 'Add New Deliverable',
                    subText: 'Provide the below information.'
                  }}
                  modalProps={{
                    isBlocking: false,
                    styles: { main: { minWidth: '520px !important' } }
                  }}
                >

                  <TextField label="Deliverable Title" value={this.state.Title} onChange={(e, newValue) => this.setState({ Title: (newValue || '') })} />

                  <Label>Template Upload</Label>
                  <input
                    type="file"
                    multiple
                    onChange={(event) => { this.handleFileChangeForDeliverable(event) }}
                    style={{ marginBottom: '10px' }}
                  />

                  <Label>Policies Upload</Label>
                  <input
                    type="file"
                    multiple
                    onChange={(event) => { this.handleFileChangeForPolicies(event) }}
                    style={{ marginBottom: '10px' }}
                  />

                  <Label>Training Upload</Label>
                  <input
                    type="file"
                    multiple
                    onChange={(event) => { this.handleFileChangeForTraining(event) }}
                    style={{ marginBottom: '10px' }}
                  />

                  <Label>Example Upload</Label>
                  <input
                    type="file"
                    multiple
                    onChange={(event) => { this.handleFileChangeForExample(event) }}
                    style={{ marginBottom: '10px' }}
                  />

                  <Dropdown
                    placeholder="Select Category"
                    label="Select Category"
                    options={this.state.categoryLstChoice}
                    defaultSelectedKey={this.state.category}
                    selectedKey={this.state.category}
                    onChange={this.changeCategory}
                    style={{ marginBottom: '10px' }}
                  />

                  <Dropdown
                    placeholder="Select Stage"
                    label="Select Stage"
                    options={this.state.stageLstChoice}
                    multiSelect
                    selectedKeys={this.state.stage} // Bind selected keys to state
                    onChange={(event, option) => this.changeStage(event, option, true)}
                    style={{ marginBottom: '10px' }}
                  />

                  {this.state.dropdownSets.map((dropdownSet, index) => (
                    <div key={index}>
                      <Dropdown
                        placeholder="Select Category"
                        label="Select Category"
                        options={this.state.categoryLstChoice}
                        defaultSelectedKey={dropdownSet.category}
                        selectedKey={dropdownSet.category}
                        onChange={(e, option) => this.handleCategoryChange(index, e, option)}
                        style={{ marginBottom: '10px' }}
                      />
                      <Dropdown
                        placeholder="Select Stage"
                        label="Select Stage"
                        options={this.state.stageLists[index] || []} // Use updated stages for each dropdown set
                        multiSelect
                        selectedKeys={dropdownSet.stage}
                        onChange={(e, options) => {
                          debugger;
                          if (options) {
                            const selectedOptions = [options];
                            this.handleStageChange(index, e, selectedOptions);
                          }
                        }}
                        style={{ marginBottom: '10px' }}
                      />
                    </div>
                  ))}

                  <PrimaryButton
                    text="Add More"
                    onClick={this.addDropdownSet}
                    style={{ marginBottom: '10px' }}
                  />

                  <DialogFooter>
                    <PrimaryButton onClick={() => this.saveItem()} text='Save' style={{ background: '#009bda', color: '#fff', border: 0 }} />
                    <DefaultButton onClick={() => { this._onDismissEvent() }} text="Cancel" />
                  </DialogFooter>
                </Dialog>

                <Dialog
                  hidden={!this.state._showDialogEventsEdit}
                  onDismiss={this._onDismissEventEdit}
                  dialogContentProps={{
                    type: DialogType.largeHeader,
                    title: 'Edit Deliverable',
                    subText: 'Provide the below information.'
                  }}
                  modalProps={{
                    isBlocking: false,
                    styles: { main: { minWidth: '520px !important' } }
                  }}
                >

                  <TextField label="Deliverable Title" value={this.state.Title} onChange={(e, newValue) => this.setState({ Title: (newValue || '') })} />

                  <Label>Template Upload</Label>
                  <a onClick={() => { window.open(this.state.siteUrl.split('/sites/')[0] + this.state._EditItem.Deliverable?.url, '_blank') }} href="javascript:void(0);">
                    <img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/Templates.png"} width="20px" />
                  </a>

                  <Label>Policies Upload</Label>
                  <a onClick={() => { window.open(this.state.siteUrl.split('/sites/')[0] + this.state._EditItem.Policies?.url, '_blank') }} href="javascript:void(0);">
                    <img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/Policies.png"} width="20px" />
                  </a>

                  <Label>Training Upload</Label>
                  <a onClick={() => { window.open(this.state.siteUrl.split('/sites/')[0] + this.state._EditItem.Training?.url, '_blank') }} href="javascript:void(0);">
                    <img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/Trainings.png"} width="20px" />
                  </a>

                  <Label>Example Upload</Label>
                  <a onClick={() => { window.open(this.state.siteUrl.split('/sites/')[0] + this.state._EditItem.Example?.url, '_blank') }} href="javascript:void(0);">
                    <img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/Example.png"} width="20px" />
                  </a>

                  <Dropdown
                    placeholder="Select Category"
                    label="Select Category"
                    options={this.state.categoryLstChoice}
                    defaultSelectedKey={this.state.category}
                    selectedKey={this.state.category}
                    onChange={this.changeCategory}
                    style={{ marginBottom: '10px' }}
                  />

                  <Dropdown
                    placeholder="Select Stage"
                    label="Select Stage"
                    options={this.state.stageLstChoice}
                    defaultSelectedKey={this.state.stage}
                    selectedKey={this.state.stage} // Bind selected keys to state
                    onChange={(event, option) => this.changeStage(event, option, false)}
                    style={{ marginBottom: '10px' }}
                  />

                  <DialogFooter>
                    <PrimaryButton onClick={() => this.saveItem()} text='Save' style={{ background: '#009bda', color: '#fff', border: 0 }} />
                    <DefaultButton onClick={() => { this._onDismissEventEdit() }} text="Cancel" />
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
                  <Label>Are you sure want to delete this Deliverable?</Label>
                  <DialogFooter>
                    <PrimaryButton onClick={() => this.deleteDeliverable()} text="Confirm" style={{ background: '#009bda', color: '#fff', border: 0 }} />
                    <DefaultButton onClick={() => { this._onDismissDelete() }} text="Cancel" />
                  </DialogFooter>
                </Dialog>
              </div>
            </div>
          )
        }
      </section >
    );
  }
}
