import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IWpFrequentlyAccessSiteProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext;
  pageTitle: string;
}
