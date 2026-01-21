import type { SuggestionCategory } from '../services/GraphSuggestionsService';

export type MainTabKey = 'add' | 'active' | 'completed' | 'myVotes' | 'admin';

export interface ISuggestionItem {
  id: number;
  title: string;
  description: string;
  votes: number;
  status: string;
  voters: string[];
  category: SuggestionCategory;
  subcategory?: string;
  createdByLoginName?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  completedDateTime?: string;
  voteEntries: IVoteEntry[];
  commentCount: number;
  comments: ISuggestionComment[];
  areCommentsLoaded: boolean;
}

export interface IVoteEntry {
  id: number;
  username: string;
  votes: number;
}

export interface ISuggestionComment {
  id: number;
  text: string;
  author?: string;
  createdDateTime?: string;
}

export interface ISubcategoryDefinition {
  key: string;
  title: string;
  category?: SuggestionCategory;
}

export interface ISimilarSuggestionsQuery {
  title: string;
  description: string;
}

export interface ISamverkansportalenState {
  activeSuggestions: IPaginatedSuggestionsState;
  completedSuggestions: IPaginatedSuggestionsState;
  activePageSize: number;
  completedPageSize: number;
  activeSuggestionsTotal?: number;
  completedSuggestionsTotal?: number;
  isLoading: boolean;
  isActiveSuggestionsLoading: boolean;
  isCompletedSuggestionsLoading: boolean;
  newTitle: string;
  newDescription: string;
  newCategory: SuggestionCategory;
  newSubcategoryKey?: string;
  subcategories: ISubcategoryDefinition[];
  categories: SuggestionCategory[];
  availableVotesByCategory: Record<string, number>;
  isUnlimitedVotes: boolean;
  statuses: string[];
  completedStatus: string;
  deniedStatus?: string;
  defaultStatus: string;
  activeFilter: IFilterState;
  completedFilter: IFilterState;
  similarSuggestions: IPaginatedSuggestionsState;
  isSimilarSuggestionsLoading: boolean;
  similarSuggestionsQuery: ISimilarSuggestionsQuery;
  selectedSimilarSuggestion?: ISuggestionItem;
  isSelectedSimilarSuggestionLoading: boolean;
  myVoteSuggestions: ISuggestionItem[];
  isMyVotesLoading: boolean;
  adminSuggestions: ISuggestionItem[];
  isAdminSuggestionsLoading: boolean;
  adminFilter: IFilterState;
  selectedMainTab: MainTabKey;
  error?: string;
  success?: string;
  expandedCommentIds: number[];
  loadingCommentIds: number[];
  commentDrafts: Record<number, string>;
  commentComposerIds: number[];
  submittingCommentIds: number[];
}

export interface IFilterState {
  searchQuery: string;
  category?: SuggestionCategory;
  subcategory?: string;
  suggestionId?: number;
  status?: string;
  includeDenied?: boolean;
}

export interface IPaginatedSuggestionsState {
  items: ISuggestionItem[];
  page: number;
  currentToken?: string;
  nextToken?: string;
  previousTokens: (string | undefined)[];
  totalCount?: number;
}

export interface ISuggestionInteractionState {
  hasVoted: boolean;
  disableVote: boolean;
  canAddComment: boolean;
  canAdvanceSuggestionStatus: boolean;
  canDeleteSuggestion: boolean;
  isVotingAllowed: boolean;
}

export interface ISuggestionCommentState {
  isExpanded: boolean;
  isLoading: boolean;
  hasLoaded: boolean;
  resolvedCount: number;
  comments: ISuggestionComment[];
  canAddComment: boolean;
  canDeleteComments: boolean;
  regionId: string;
  toggleId: string;
  isComposerVisible: boolean;
  draftText: string;
  isSubmitting: boolean;
}

export interface ISuggestionViewModel {
  item: ISuggestionItem;
  interaction: ISuggestionInteractionState;
  comment: ISuggestionCommentState;
}

export type SuggestionAction = (item: ISuggestionItem) => void | Promise<void>;
export type CommentAction = (item: ISuggestionItem, comment: ISuggestionComment) => void | Promise<void>;
