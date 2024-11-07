import * as React from 'react';
import styles from './WpComponentsLinks.module.scss';
import type { IWpComponentsLinksProps } from './IWpComponentsLinksProps';
import { SPFI, spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "bootstrap/dist/css/bootstrap.min.css"
//import { PrimaryButton, TextField, Label, Dialog, DialogType, DialogFooter, DefaultButton } from '@fluentui/react';

export interface IWpComponentsLinksState {
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
  Title: string,
  EditImageURL: string,
  showEditImageURL: boolean,
  saveButtonAdd: boolean,
  savButtonEdit: boolean,
  _EditItem: any;
}

export default class WpComponentsLinks extends React.Component<IWpComponentsLinksProps, IWpComponentsLinksState> {

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
    let _componentLinksItems = [];
    let categoryfilter = this.props.category;
    debugger;
    let tmpcomponentLinks = await this._sp.web.lists.getByTitle("Departments").items.select("ID", "Title", "Category", "ImageURL", "RedirectURL", "SortOrder", "HoverDescription").filter("Category eq '" + categoryfilter + "'").orderBy("SortOrder", true).top(4)();
    if (tmpcomponentLinks && tmpcomponentLinks.length > 0) {
      _componentLinksItems = tmpcomponentLinks;
    } else {
      _componentLinksItems = [];
    }
    this.setState({ ComponentLinksItemsState: _componentLinksItems })
    console.log(this.state.ComponentLinksItemsState);
    return Promise.resolve("")
  }



  public render(): React.ReactElement<IWpComponentsLinksProps> {
    const {
      hasTeamsContext,
    } = this.props;

    return (
      <section className={`${styles.wpComponentsLinks} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={`${styles.Internal_content}`}>
          <div className={`${styles.Img_content}`}>
            <div className={`${styles.Photos_cont}`}>
              <div>&nbsp;</div>
              {this.state.ComponentLinksItemsState &&
                this.state.ComponentLinksItemsState.map((nhItem, index) => (
                  <div className={`${styles.Thumbnails_cont}`}>
                    <a href={nhItem.RedirectURL}>
                      <img src={decodeURI(nhItem.ImageURL)} alt="" />
                      <div className={`${styles.Black_cont}`}></div>
                      <div className={`${styles.Title_cont}`}>{nhItem.Title}</div>
                      <h5 className={`${styles.Mission}`}>Mission</h5>
                      <div className={`${styles.Hover_Desc}`}>{nhItem.HoverDescription}</div>
                    </a>
                  </div>
                ))
              }
              <div>&nbsp;</div>
            </div>
          </div>
        </div>
      </section>
    );
  }
}
