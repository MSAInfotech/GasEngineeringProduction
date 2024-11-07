import * as React from 'react';
import styles from './WpProjectEngineeringTeam.module.scss';
import type { IWpProjectEngineeringTeamProps } from './IWpProjectEngineeringTeamProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPFI, spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
//import { PrimaryButton, TextField, Label, Dialog, DialogType, DialogFooter, DefaultButton } from '@fluentui/react';
export interface IWpProjectEngineeringTeamState {
  ProjectEngineeringTeamState: any[];
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
}


export default class WpProjectEngineeringTeam extends React.Component<IWpProjectEngineeringTeamProps, IWpProjectEngineeringTeamState> {

  private _sp: SPFI;

  constructor(props: any) {
    super(props);
    sp: this._sp,
      //this._sp = spfi("https://sempra.sharepoint.com/sites/gasopscon/eng/").using(SPFx(this.props.context));
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      ProjectEngineeringTeamState: [],
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
      _deleteItem: {}
    }
  }

  public componentDidMount(): void {
    this.fetchData();
  }

  protected async fetchData() {
    try {
      await this._getItems();

    } catch (error) {
      // Handle error
    }
  }

  private async _getItems(): Promise<string> {
    let _ProjectEngineeringTeamItems = [];
    let categoryfilter = this.props.category;
    let _designation = this.props.designation;
    debugger;
    let tmpProjectEngineeringTeam = await this._sp.web.lists.getByTitle("ProjectTeams").items.select("ID", "Title", "Designation", "Category", "Description", "ImageURL").filter("Category eq '" + categoryfilter + "' and Designation eq '" + _designation + "'").orderBy("Title", true)();
    if (tmpProjectEngineeringTeam && tmpProjectEngineeringTeam.length > 0) {
      _ProjectEngineeringTeamItems = tmpProjectEngineeringTeam;
    } else {
      _ProjectEngineeringTeamItems = [];
    }
    this.setState({ ProjectEngineeringTeamState: _ProjectEngineeringTeamItems })
    console.log(this.state.ProjectEngineeringTeamState);
    return Promise.resolve("")
  }

  public render(): React.ReactElement<IWpProjectEngineeringTeamProps> {
    const {
      hasTeamsContext
    } = this.props;

    return (
      <section className={`${this.props.category === 'Overall Project Engineering' ? styles.wpProjectEngineeringTeam : `${styles.wpProjectEngineeringTeamOtherPages} ${hasTeamsContext ? styles.teams : ''}`} col-md-3 col-lg-3 col-sm-3 float-start`}>
        <div>
          <h3 className={`${styles.projectEngineeringTeamDesignation}`}>{this.props.designation}</h3>
        </div>
        <div className={`${styles.Team_lead}`}>

          {this.state.ProjectEngineeringTeamState &&
            this.state.ProjectEngineeringTeamState.map((nhItem, index) => (
              <div className={`${styles.profile_wrapper}`}>
                <div className={`${this.props.designation === 'Contractors' ? styles.profile_body_contractors : styles.profile_body}`}>
                  <img src={nhItem.ImageURL} width="20px" height="20px" alt="" />
                  <div className={`${this.props.designation === 'Contractors' ? styles.profile_details_contractors : styles.profile_details}`}>
                    <h1>{nhItem.Title}</h1>
                    <p className={`${styles.designation}`}>{nhItem.Designation}</p>
                    {this.props.designation === 'Contractors' ? '' :
                      <>
                        {<p className={`${styles.description}`}>{nhItem.Description}</p>}
                      </>
                    }
                  </div>

                </div>
              </div>
            ))
          }
        </div>
      </section>
    );
  }
}
