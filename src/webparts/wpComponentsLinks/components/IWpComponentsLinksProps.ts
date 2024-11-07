import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IWpComponentsLinksProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  category: string;
  WebServerRelativeURL: string;
}
