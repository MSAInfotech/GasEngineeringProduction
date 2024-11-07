import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IWpTrainingTemplatesPoliciesProps {
  sp: any;
  currentUserID: number;
  currentUserName: string;
  currentUserEmail: string;
  WebServerRelativeURL: string;
  LibName: string;
  FolderURL: string;
  FilterColumns: string;
  DataColumns: string;
  context: WebPartContext;
}
