import GraphSuggestionsService from '../services/GraphSuggestionsService';

export const DEFAULT_SUGGESTIONS_LIST_TITLE: string = 'SamverkansportalenSuggestions';
export const DEFAULT_SUGGESTIONS_HEADER_TITLE: string = 'Suggestion board';
export const DEFAULT_SUGGESTIONS_HEADER_SUBTITLE: string = 'Share ideas, cast your votes and celebrate what has been delivered.';
export const DEFAULT_STATUS_DEFINITIONS: string = 'Active\nIn Progress\nDone';

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
  voteListTitle?: string;
  useTableLayout?: boolean;
  subcategoryListTitle?: string;
  categoryListTitle?: string;
  statusListTitle?: string;
  commentListTitle?: string;
  headerTitle: string;
  headerSubtitle: string;
  statuses: string[];
  completedStatus: string;
}
