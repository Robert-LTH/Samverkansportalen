import { SPHttpClient } from '@microsoft/sp-http';

export const DEFAULT_SUGGESTIONS_LIST_TITLE: string = 'SamverkansportalenSuggestions';

export interface ISamverkansportalenProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userLoginName: string;
  siteUrl: string;
  spHttpClient: SPHttpClient;
  listTitle?: string;
}
