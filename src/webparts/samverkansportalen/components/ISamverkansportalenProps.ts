import * as strings from 'SamverkansportalenWebPartStrings';
import GraphSuggestionsService from '../services/GraphSuggestionsService';

export const DEFAULT_SUGGESTIONS_LIST_TITLE: string = 'SamverkansportalenSuggestions';
export const DEFAULT_SUGGESTIONS_HEADER_TITLE: string = strings.DefaultSuggestionsHeaderTitle;
export const DEFAULT_SUGGESTIONS_HEADER_SUBTITLE: string = strings.DefaultSuggestionsHeaderSubtitle;
export const DEFAULT_STATUS_DEFINITIONS: string = strings.DefaultStatusDefinitions;
export const DEFAULT_TOTAL_VOTES_PER_USER: number = 5;

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
  showMetadataInIdColumn?: boolean;
  headerTitle: string;
  headerSubtitle: string;
  statuses: string[];
  completedStatus: string;
  defaultStatus: string;
  totalVotesPerUser: number;
}
