import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IWpReportsDashboardProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  WebServerRelativeURL: string;
}
