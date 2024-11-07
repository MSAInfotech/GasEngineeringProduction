import * as React from 'react';
import styles from './WpTabs.module.scss';
import type { IWpTabsProps } from './IWpTabsProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPFI, spfi, SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";

export interface IWpTabsState {
  TabsItemsState: any[];
  linkUrl: string;
  setLinkUrl: string;
  Title: string;
  isOverallEngineering: boolean;
  inTakeControls: boolean;
  executionStorageTransmission: boolean;
  executionSpecialtyProjects: boolean;
  executionMajorProjects: boolean;
  announcements: boolean;
  reportsdashboard: boolean;
}


export default class WpTabs extends React.Component<IWpTabsProps, IWpTabsState> {
  private _sp: SPFI;

  constructor(props: any) {
    super(props);
    sp: this._sp,
      //this._sp = spfi("https://sempra.sharepoint.com/sites/gasopscon/eng/").using(SPFx(this.props.context));
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      TabsItemsState: [],
      linkUrl: "",
      setLinkUrl: "",
      Title: "",
      isOverallEngineering: false,
      inTakeControls: false,
      executionStorageTransmission: false,
      executionSpecialtyProjects: false,
      executionMajorProjects: false,
      announcements: false,
      reportsdashboard: false
    }


  }
  public componentDidMount(): void {

    debugger;
    if (this.props.category == "Overall Project Engineering") {
      console.log(this.props.category);
      this.setState({
        isOverallEngineering: !this.state.isOverallEngineering,
        inTakeControls: this.state.inTakeControls,
        executionStorageTransmission: this.state.executionStorageTransmission,
        executionSpecialtyProjects: this.state.executionSpecialtyProjects,
        executionMajorProjects: this.state.executionMajorProjects,
        // announcements: this.state.announcements,
        // reportsdashboard: this.state.reportsdashboard
      });
    }
    else if (this.props.category == "Intake and Controls") {
      console.log(this.props.category);
      this.setState({
        isOverallEngineering: this.state.isOverallEngineering,
        inTakeControls: !this.state.inTakeControls,
        executionStorageTransmission: this.state.executionStorageTransmission,
        executionSpecialtyProjects: this.state.executionSpecialtyProjects,
        executionMajorProjects: this.state.executionMajorProjects
      });
    }
    else if (this.props.category == "Execution Storage and Transmission") {
      console.log(this.props.category);
      this.setState({
        isOverallEngineering: this.state.isOverallEngineering,
        inTakeControls: this.state.inTakeControls,
        executionStorageTransmission: !this.state.executionStorageTransmission,
        executionSpecialtyProjects: this.state.executionSpecialtyProjects,
        executionMajorProjects: this.state.executionMajorProjects
      });
    }
    else if (this.props.category == "Execution Specialty Projects") {
      console.log(this.props.category);
      this.setState({
        isOverallEngineering: this.state.isOverallEngineering,
        inTakeControls: this.state.inTakeControls,
        executionStorageTransmission: this.state.executionStorageTransmission,
        executionSpecialtyProjects: !this.state.executionSpecialtyProjects,
        executionMajorProjects: this.state.executionMajorProjects
      });
    }
    else if (this.props.category == "Execution Major Projects") {
      console.log(this.props.category);
      this.setState({
        isOverallEngineering: this.state.isOverallEngineering,
        inTakeControls: this.state.inTakeControls,
        executionStorageTransmission: this.state.executionStorageTransmission,
        executionSpecialtyProjects: this.state.executionSpecialtyProjects,
        executionMajorProjects: !this.state.executionMajorProjects
      });
    }
    else if (this.props.category == "Announcements") {
      console.log(this.props.category);
      this.setState({
        announcements: !this.state.announcements,
        reportsdashboard: this.state.reportsdashboard
      });
    }
    else if (this.props.category == "Reports/ Dashboard") {
      console.log(this.props.category);
      this.setState({
        announcements: this.state.announcements,
        reportsdashboard: !this.state.reportsdashboard
      });
    }
    console.log(this.state.isOverallEngineering);
    console.log(this.state.announcements);
  }
  public render(): React.ReactElement<IWpTabsProps> {

    const { isOverallEngineering, inTakeControls, executionStorageTransmission, executionSpecialtyProjects, executionMajorProjects, announcements, reportsdashboard } = this.state;

    return (
      <section>

        <div className={`${styles.pagecontent}`}>
          <ul className={`${this.props.ListName === "Communications" ? styles.tabs_cm : styles.tabs}`}>
            {(this.props.ListName === "Project Engineering Team") && (
              <>
                <li className={`${styles['tab-li']}`}>
                  <a href={this.props.WebServerRelativeURL + "/SitePages/Overall-Project-Engineering.aspx"} className={`${styles.tablilink} ${isOverallEngineering ? styles.active : ''}`}>Overall Project <br />Engineering</a>
                </li>
                <li className={`${styles['tab-li']}`}>
                  <a href={this.props.WebServerRelativeURL + "/SitePages/Intake-&-Controls.aspx"} className={`${styles.tablilink} ${inTakeControls ? styles.active : ''}`}>Intake & Controls</a>
                </li>
                <li className={`${styles['tab-li']}`}>
                  <a href={this.props.WebServerRelativeURL + "/SitePages/Execution-Storage-&-Transmis.aspx"} className={`${styles.tablilink} ${executionStorageTransmission ? styles.active : ''}`}>Execution -<br />Storage & Transmission </a>
                </li>
                <li className={`${styles['tab-li']}`}>
                  <a href={this.props.WebServerRelativeURL + "/SitePages/Execution.aspx"} className={`${styles.tablilink} ${executionSpecialtyProjects ? styles.active : ''}`}>Execution -<br />Specialty Projects </a>
                </li>
                <li className={`${styles['tab-li']}`}>
                  <a href={this.props.WebServerRelativeURL + "/SitePages/Execution---Major-Projects.aspx"} className={`${styles.tablilink} ${executionMajorProjects ? styles.active : ''}`}>Execution -<br />Major Projects</a>
                </li>
              </>
            )}

            {(this.props.ListName === "Communications") && (
              <>
                <li className={`${styles['tab-li']}`}>
                  <a href={this.props.WebServerRelativeURL + "/SitePages/Announcements.aspx"} className={`${styles.tablilink} ${announcements ? styles.active : ''}`}>Announcements </a>
                </li>
                <li className={`${styles['tab-li']}`}>
                  <a href={this.props.WebServerRelativeURL + "/SitePages/Reports-Dashboard.aspx"} className={`${styles.tablilink} ${reportsdashboard ? styles.active : ''}`}>Reports/ Dashboard</a>
                </li>
              </>
            )}
          </ul>
        </div>
      </section>
    );
  }
}
