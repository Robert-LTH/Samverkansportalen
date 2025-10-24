import GraphSuggestionsService from '../services/GraphSuggestionsService';

export const DEFAULT_SUGGESTIONS_LIST_TITLE: string = 'SamverkansportalenSuggestions';

export interface ISamverkansportalenProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userLoginName: string;
  isCurrentUserAdmin: boolean;
  graphService: GraphSuggestionsService;
  listTitle?: string;
  useTableLayout?: boolean;
}
