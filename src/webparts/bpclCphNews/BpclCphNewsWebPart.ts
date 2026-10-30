import * as React from "react";
import * as ReactDOM from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import MyViewNews from "./components/MyViewNews";
import { SPFI, spfi, SPFx } from "@pnp/sp";
 
export interface IMyViewNewsWebPartProps {}

export default class MyViewNewsWebPart extends BaseClientSideWebPart<IMyViewNewsWebPartProps> {
  public render(): void {
    const sp: SPFI = spfi().using(SPFx(this.context));

    const element = React.createElement(MyViewNews, {
      sp: sp,
      context: this.context
    });

    ReactDOM.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
}
  