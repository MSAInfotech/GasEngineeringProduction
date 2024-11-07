import * as React from 'react';
import styles from './WpWhatWeDo.module.scss';
import type { IWpWhatWeDoProps } from './IWpWhatWeDoProps';
import { SPFI, spfi, SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";

export interface IWpWhatWeDoState {
  WhatWeDoItemsState: any[];
  linkUrl: string;
  setLinkUrl: string;
  Title: string;
}

export default class WpWhatWeDo extends React.Component<IWpWhatWeDoProps, IWpWhatWeDoState> {

  private _sp: SPFI;

  constructor(props: any) {
    super(props);
    sp: this._sp,
      this._sp = spfi().using(SPFx(this.props.context));
    this.state = {
      WhatWeDoItemsState: [],
      linkUrl: "",
      setLinkUrl: "",
      Title: "",

    }
  }

  public componentDidMount(): void {
    this.fetchData();
  }

  protected async fetchData() {
    try {
      await this._getWhatWeDoItems();

    } catch (error) {
      // Handle error
    }
  }

  private async _getWhatWeDoItems(): Promise<string> {
    let _WhatWeDoItem: any[] = [];
    this.setState({ WhatWeDoItemsState: [] });
    let categoryfilter = this.props.category;
    debugger;
    let tmpWhatWeDo = await this._sp.web.lists.getByTitle("WhatWeDo").items.select("ID", "Title", "Category", "Description").filter("Category eq '" + categoryfilter + "'").orderBy("ID", true).top(1)();
    if (tmpWhatWeDo && tmpWhatWeDo.length > 0) {
      tmpWhatWeDo.forEach(element => {
        _WhatWeDoItem.push(
          {
            "ID": element.ID,
            "Title": element.Title,
            "Description": element.Description
          });
      })

    } else {
      _WhatWeDoItem = [];
    }
    this.setState({ WhatWeDoItemsState: _WhatWeDoItem })
    console.log(this.state.WhatWeDoItemsState);
    return Promise.resolve("")
  }

  public render(): React.ReactElement<IWpWhatWeDoProps> {

    return (
      <section className={`${this.props.category === 'Overall Project Engineering' ? styles.wpWhatWeDoOverallProjectEng : `${styles.wpWhatWeDoOtherPages} col-md-7 col-sm-7 col-lg-7 float-start`}`}>
        <div className={`${styles.tabaricalpara}`}>
          <h2>What We Do:</h2>
          {this.state.WhatWeDoItemsState &&
            this.state.WhatWeDoItemsState.map((nhItem, index) => (
              <div className={`${styles.para}`}>
                <p>
                  {nhItem.Description}
                </p>
              </div>
            ))
          }
        </div>
      </section>
    );
  }
}
