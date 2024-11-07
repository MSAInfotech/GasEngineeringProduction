import * as React from 'react';
import styles from './WpTopLevelTabs.module.scss';
import type { IWpTopLevelTabsProps } from './IWpTopLevelTabsProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { SPFI, spfi, SPFx } from '@pnp/sp/presets/all';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faHouseChimney, faGear } from '@fortawesome/free-solid-svg-icons';

interface TabComponentState {
  activeTab: string;
  TabsItemsState: any[];
  linkUrl: string;
  setLinkUrl: string;
  Title: string;
  isUserInGroup: boolean;
}


export default class WpTopLevelTabs extends React.Component<IWpTopLevelTabsProps, TabComponentState> {

  private _sp: SPFI;
  constructor(props: any) {
    super(props);
    this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      activeTab: '',
      linkUrl: "",
      setLinkUrl: "",
      Title: "",
      TabsItemsState: [],
      isUserInGroup: false,
    };
  }

  handleTabClick = (tab: string, url: string) => {
    this.setState({ activeTab: tab });
    window.location.href = url;
  }
  public componentDidMount(): void {
    this.checkUserInGroup("PortalAdmins");

    if (this.props.category == "HomePage") {
      console.log(this.props.category);
      this.setState({
        activeTab: "tab1"
      });
    }
    else if (this.props.category == "Project Engineering Team") {
      console.log(this.props.category);
      this.setState({
        activeTab: "tab2"

      });
    }
    else if (this.props.category == "Communications") {
      console.log(this.props.category);
      this.setState({
        activeTab: "tab3"

      });
    }
    else if (this.props.category == "Framework") {
      console.log(this.props.category);
      this.setState({
        activeTab: "tab4"

      });
    }
    else if (this.props.category == "Training") {
      console.log(this.props.category);
      this.setState({
        activeTab: "tab5"

      });
    }
    //console.log(this.state.isOverallEngineering);
  }

  private async checkUserInGroup(groupName: string): Promise<void> {
    const currentUser = await this._sp.web.currentUser();
    const groups = await this._sp.web.siteGroups.getByName(groupName).users();
    const isUserInGroupExists = groups.some(user => user.Id === currentUser.Id);
    this.setState({ isUserInGroup: isUserInGroupExists });
  }

  public render(): React.ReactElement<IWpTopLevelTabsProps> {

    const { activeTab } = this.state;
    return (
      <section className={`${styles.wpTopLevelTabs}`}>
        <div>
          <ul className={styles.tabs}>
            <li className={`${styles.tab} ${activeTab === 'tab1' ? styles.active : ''}`}>
              <a
                href={this.props.WebServerRelativeURL + "/SitePages/hub.aspx"}
                onClick={() => {
                  this.setState({ activeTab: 'tab1' })
                }}
                style={{ textDecoration: 'inherit', color: 'inherit' }}>
                <FontAwesomeIcon icon={faHouseChimney} style={{ color: "#ffffff", fontSize: '24px' }} />
              </a>
            </li>

            <li className={`${styles.tab} ${activeTab === 'tab2' ? styles.active : ''}`}>
              <a
                href={`${activeTab === 'tab2' ? "#" : this.props.WebServerRelativeURL + "/SitePages/Overall-Project-Engineering.aspx"}`}
                onClick={() => {
                  this.setState({ activeTab: 'tab2' })
                }}
                style={{ textDecoration: 'inherit', color: 'inherit' }}>
                Project Engineering Teams
              </a>
            </li>

            <li className={`${styles.tab} ${activeTab === 'tab3' ? styles.active : ''}`}>
              <a
                href={`${activeTab === 'tab3' ? "#" : this.props.WebServerRelativeURL + "/SitePages/Announcements.aspx"}`}
                onClick={() => {
                  this.setState({ activeTab: 'tab3' })
                }}
                style={{ textDecoration: 'inherit', color: 'inherit' }}>
                Communications
              </a>
            </li>

            <li className={`${styles.tab} ${activeTab === 'tab4' ? styles.active : ''}`}>
              <a
                href={`${activeTab === 'tab4' ? "#" : this.props.WebServerRelativeURL + "/SitePages/Interactive-Checklist.aspx"}`}
                onClick={() => {
                  this.setState({ activeTab: 'tab4' })
                }}
                style={{ textDecoration: 'inherit', color: 'inherit' }}>
                Framework
              </a>
            </li>

            <li className={`${styles.tab} ${activeTab === 'tab5' ? styles.active : ''}`}>
              <a
                href={`${activeTab === 'tab5' ? "#" : this.props.WebServerRelativeURL + "/SitePages/Interactive-Checklist-Filters.aspx"}`}
                onClick={() => {
                  this.setState({ activeTab: 'tab5' })
                }}
                style={{ textDecoration: 'inherit', color: 'inherit' }}>
                Training, Templates, & Policies
              </a>
            </li>

            {this.state.isUserInGroup && (
              <li className={`${styles.tab}`} >
                <FontAwesomeIcon icon={faGear} onClick={() => {
                  window.open('https://sempra.sharepoint.com.mcas.ms/sites/gasopscon/eng/resource%20hub/_layouts/15/groups.aspx', '_blank');
                }}
                  style={{ color: '#ffffff', fontSize: '24px' }} />
              </li>
            )}
          </ul>

          {/* <div className={styles.tabContent}>
          {activeTab === 'tab1' && <div>Content for Tab 1</div>}
          {activeTab === 'tab2' && <div>Content for Tab 2</div>}
          {activeTab === 'tab3' && <div>Content for Tab 3</div>}
          {activeTab === 'tab3' && <div>Content for Tab 4</div>}
        </div> */}
        </div>
      </section>
    );
  }
}
