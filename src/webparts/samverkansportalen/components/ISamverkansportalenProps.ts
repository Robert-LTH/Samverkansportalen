import { SPHttpClient } from '@microsoft/sp-http';

export interface ISamverkansportalenProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userLoginName: string;
  siteUrl: string;
  spHttpClient: SPHttpClient;
}
