import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IWpFrameworkProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  ListName: string;
  context: WebPartContext;
  category: string;
  WebServerRelativeURL: string;
}
