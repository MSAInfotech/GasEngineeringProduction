import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IWpInteractiveChecklistProps {
  currentUserID: number;
  currentUserName: string;
  currentUserEmail: string;
  WebServerRelativeURL: string;
  LibName: string;
  filterPageURL: string;
  context: WebPartContext;
  hasTeamsContext: boolean;
}
