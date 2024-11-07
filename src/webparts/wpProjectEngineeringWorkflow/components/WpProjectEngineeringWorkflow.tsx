import * as React from 'react';
import styles from './WpProjectEngineeringWorkflow.module.scss';
import type { IWpProjectEngineeringWorkflowProps } from './IWpProjectEngineeringWorkflowProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { SPFI, spfi, SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
import "bootstrap/dist/css/bootstrap.min.css"
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPen, faTrashAlt } from '@fortawesome/free-solid-svg-icons';
import { PrimaryButton, TextField, Label, Dialog, DialogType, DialogFooter, DefaultButton } from '@fluentui/react';

export interface IWpProjectEngineeringWorkflowState {
  ProjectEngineeringWorkflowItemsState: any[];
  _showDialogEvents: boolean,
  _showDialogDelete: boolean,
  _showDialogDeleteConfirm: boolean,
  _showDialogDeleteConfirmYes: boolean,
  _showDialogDeleteConfirmNo: boolean,
  _deleteItem: any,
  linkUrl: string,
  setLinkUrl: string,
  file: any,
  filename: string,
  Title: string,
  saveButtonAdd: boolean,
  savButtonEdit: boolean,
  _EditItem: any;
  isUserInGroup: boolean;
  isItemInEditMode: boolean
}

export default class WpProjectEngineeringWorkflow extends React.Component<IWpProjectEngineeringWorkflowProps, IWpProjectEngineeringWorkflowState> {

  private _sp: SPFI;

  constructor(props: any) {
    super(props);
    sp: this._sp,
      //this._sp = spfi("https://sempra.sharepoint.com/sites/gasopscon/eng/").using(SPFx(this.props.context));
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      ProjectEngineeringWorkflowItemsState: [],
      _showDialogEvents: false,
      _showDialogDelete: false,
      _showDialogDeleteConfirm: false,
      _showDialogDeleteConfirmYes: false,
      _showDialogDeleteConfirmNo: false,
      linkUrl: "",
      setLinkUrl: "",
      file: null,
      filename: "",
      Title: "",
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
      await this._getProjectEngineeringWokflowItems();

    } catch (error) {
      // Handle error
    }
  }

  private async _getProjectEngineeringWokflowItems(): Promise<string> {
    let _ProjectEngineeringWorkflowItems: any[] = [];
    //let _ProjectEngineeringWorkflowArrItems = [];
    this.setState({ ProjectEngineeringWorkflowItemsState: [] });
    let categoryfilter = this.props.category;
    try {
      let tmpProjectEngineeringWorkflow = await this._sp.web.lists.getByTitle("ProjectEngineeringWorkFlow")
        .items.select("ID", "Title", "RedirectUrl", "Category")
        .filter("Category eq '" + categoryfilter + "'")
        .orderBy("ID", false)();

      if (tmpProjectEngineeringWorkflow && tmpProjectEngineeringWorkflow.length > 0) {
        tmpProjectEngineeringWorkflow.forEach(element => {
          _ProjectEngineeringWorkflowItems.push({
            "ID": element.ID,
            "Title": element.Title,
            "RedirectUrl": element.RedirectUrl,
          });
        });
      } else {
        _ProjectEngineeringWorkflowItems = [];
      }

      this.setState({ ProjectEngineeringWorkflowItemsState: _ProjectEngineeringWorkflowItems });
      console.log(this.state.ProjectEngineeringWorkflowItemsState);
    } catch (error) {
      console.error('Error fetching ProjectEngineeringWorkFlow items:', error);
      // Handle the error (e.g., show a message to the user)
    }
    return Promise.resolve("")
  }

  private saveItem() {
    if (this.state.saveButtonAdd) {
      this.AddItem();
    }
    else {
      this.updateItem();
    }
  }

  private async AddItem() {
    const { Title, setLinkUrl } = this.state;
    if (!Title || Title.trim() == '') {
      alert('Title cannot be empty.');
      return;
    }
    if (!setLinkUrl || setLinkUrl.trim() == '') {
      alert('Link URL cannot be empty.');
      return;
    }
    const i = await this._sp.web.lists.getByTitle("ProjectEngineeringWorkFlow").items.add({
      Title: this.state.Title,
      RedirectUrl: this.state.setLinkUrl,
      Category: this.props.category
    });

    if (i) {
      console.log("successfully created");
    }

    alert('Project engineering workFlow added successfully!');
    this._getProjectEngineeringWokflowItems();
    this._onDismissEvent();

  };

  private async updateItem() {
    const { Title, setLinkUrl } = this.state;
    if (!Title || Title.trim() == '') {
      alert('Title cannot be empty.');
      return;
    }
    if (!setLinkUrl || setLinkUrl.trim() == '') {
      alert('Link URL cannot be empty.');
      return;
    }

    let updateID = this.state._EditItem.ID;
    const i = await this._sp.web.lists.getByTitle("ProjectEngineeringWorkFlow").items.getById(updateID).update({
      Title: this.state.Title,
      RedirectUrl: this.state.setLinkUrl,
    });

    if (i) {
      console.log("Project engineering workFlow updated successfully.");
    }
    alert('Project engineering workFlow updated successfully!');
    this._getProjectEngineeringWokflowItems();
    this._onDismissEvent();
  };

  private _onDismissEvent(): void {
    this.setState({
      _showDialogEvents: !this.state._showDialogEvents,
      saveButtonAdd: true,
      savButtonEdit: false,
      Title: "",
      setLinkUrl: "",
      isItemInEditMode: false
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

      Title: tmpEditItem.Title,
      setLinkUrl: tmpEditItem.RedirectUrl,
      saveButtonAdd: false,
      savButtonEdit: true,
      isItemInEditMode: true,
      _EditItem: tmpEditItem

    });

  }

  private _onDeleteItem(tmpDeleteItem: any): void {
    this.setState({
      _showDialogDelete: !this.state._showDialogDelete,
      saveButtonAdd: false,
      savButtonEdit: false,
      _deleteItem: tmpDeleteItem
    });
    // return false;
  }

  private async deleteItem() {

    let itemId = this.state._deleteItem.ID;
    try {

      await this._sp.web.lists.getByTitle("ProjectEngineeringWorkFlow").items.getById(itemId).delete();
      alert(`Project engineering workFlow deleted successfully.`);
      this._getProjectEngineeringWokflowItems();
      this.setState({ _showDialogDelete: !this.state._showDialogDelete });
    } catch (error) {
      console.error(`Error deleting item with ID ${itemId}: `, error);
    }

    //this.deleteListItem();
  }

  public render(): React.ReactElement<IWpProjectEngineeringWorkflowProps> {
    // const {
    //   description,
    //   isDarkTheme,
    //   environmentMessage,
    //   hasTeamsContext,
    //   userDisplayName
    // } = this.props;
    const {
      hasTeamsContext
    } = this.props;
    return (
      <section className={`${styles.wpProjectEngineeringWorkflow} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={`${styles.workflowheader} d-flex justify-content-between align-items-center`}>
          <h2>Project Engineering Workflow</h2>
          {this.state.isUserInGroup && (
            <div className={styles.iconactions}>
              <a onClick={() => { this._onDismissEvent() }} href='javascript:void(0);'><img src={this.props.WebServerRelativeURL + "/SiteAssets/PortalImages/plus-symbol.png"} width="20px" /></a>
            </div>
          )}
        </div>

        <div className={`${styles.buttongroup}`}>
          {this.state.ProjectEngineeringWorkflowItemsState &&
            this.state.ProjectEngineeringWorkflowItemsState.map((nhItem, index) => (
              <div key={index} className={`${styles.textcenter} d-flex`}>
                <a href='javascript:void(0);' onClick={() => {
                  if (nhItem.RedirectUrl) {
                    this._onclickbutton(nhItem.RedirectUrl);
                  } else {
                    console.log('Redirect URL is not available');
                    // alert('No valid redirect URL available.');
                  }
                }} className={`${styles.workflowbutton} align-items-center justify-content-between ${this.state.isUserInGroup ? '' : 'w-100'}`}>
                  <span>{nhItem.Title}</span>
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
              title: this.state.isItemInEditMode == true ? 'Edit Project Engineering Workflow' : 'Add New Project Engineering Workflow',
              subText: 'Provide the below information.'
            }}
            modalProps={{
              isBlocking: false,
              styles: { main: { minWidth: '520px !important' } }
            }}
          >
            <TextField label="Project Engineering Workflow Title" value={this.state.Title} onChange={(e, newValue) => this.setState({ Title: (newValue || '') })} />
            <TextField label="Link URL" value={this.state.setLinkUrl} onChange={(e, newValue) => this.setState({ setLinkUrl: (newValue || '') })} />

            <DialogFooter>
              <PrimaryButton onClick={() => this.saveItem()} text="Save" style={{ background: '#009bda', color: '#fff', border: 0 }} />
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

            <Label>Are you sure want to delete this project engineering workflow ?</Label>

            <DialogFooter>
              <PrimaryButton onClick={() => this.deleteItem()} text="Confirm" style={{ background: '#009bda', color: '#fff', border: 0 }} />
              <DefaultButton onClick={() => { this._onDismissDelete() }} text="Cancel" />
            </DialogFooter>
          </Dialog>
        </div>
      </section>
    );
  }
}
