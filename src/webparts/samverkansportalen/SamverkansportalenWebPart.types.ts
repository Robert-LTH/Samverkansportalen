import { type IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';

export type ListCreationType =
  | 'suggestions'
  | 'votes'
  | 'comments'
  | 'subcategories'
  | 'categories'
  | 'statuses';

export interface IStatusDropdownOption extends IPropertyPaneDropdownOption {
  data?: {
    sortOrder?: number;
    isCompleted?: boolean;
  };
}

export interface IStatusDefinition {
  id: number;
  title: string;
  order?: number;
  isCompleted: boolean;
}

export interface ISamverkansportalenWebPartProps {
  description: string;
  listTitle?: string;
  useTableLayout?: boolean;
  subcategoryListTitle?: string;
  categoryListTitle?: string;
  statusListTitle?: string;
  voteListTitle?: string;
  commentListTitle?: string;
  selectedSubcategoryKey?: string;
  newSubcategoryTitle?: string;
  selectedCategoryKey?: string;
  newCategoryTitle?: string;
  selectedStatusKey?: string;
  newStatusTitle?: string;
  headerTitle: string;
  headerSubtitle: string;
  statusDefinitions?: string;
  completedStatus?: string;
  deniedStatus?: string;
  defaultStatus?: string;
  totalVotesPerUser?: string;
  showMetadataInIdColumn?: boolean;
}
