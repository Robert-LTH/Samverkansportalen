export type SuggestionStatus = 'Proposed' | 'InProgress' | 'Completed' | 'Removed';

export const activeStatuses: SuggestionStatus[] = ['Proposed', 'InProgress'];

export const suggestionStatusOptions: { key: SuggestionStatus; text: string }[] = [
  { key: 'Proposed', text: 'Föreslagen' },
  { key: 'InProgress', text: 'Pågående' },
  { key: 'Completed', text: 'Genomförd' },
  { key: 'Removed', text: 'Avslutad' }
];

export interface IUserInfo {
  id?: number;
  title: string;
  email?: string;
}

export interface ISuggestionItem {
  id: number;
  title: string;
  description: string;
  status: SuggestionStatus;
  created: string;
  createdBy: IUserInfo;
}

export interface ISuggestionWithVotes extends ISuggestionItem {
  totalVotes: number;
  activeVotes: number;
  userHasActiveVote: boolean;
  userVoteId?: number;
  userHasAnyVote: boolean;
}
