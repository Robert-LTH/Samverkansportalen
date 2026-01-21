import * as React from 'react';
import { DefaultButton, Dropdown, Spinner, SpinnerSize, TextField, Toggle, type IDropdownOption } from '@fluentui/react';
import styles from '../Samverkansportalen.module.scss';
import * as strings from 'SamverkansportalenWebPartStrings';
import type { CommentAction, ISuggestionItem, ISuggestionViewModel, SuggestionAction } from '../types';
import PaginationControls from '../common/PaginationControls';
import SuggestionList from './SuggestionList';

interface ISuggestionSectionProps {
  title: string;
  titleId: string;
  contentId: string;
  isLoading: boolean;
  isSectionLoading: boolean;
  searchValue: string;
  onSearchChange: (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ) => void;
  categoryOptions: IDropdownOption[];
  selectedCategoryKey: IDropdownOption['key'] | undefined;
  onCategoryChange: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void;
  disableCategoryDropdown: boolean;
  subcategoryOptions: IDropdownOption[];
  selectedSubcategoryKey: IDropdownOption['key'] | undefined;
  onSubcategoryChange: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void;
  disableSubcategoryDropdown: boolean;
  subcategoryPlaceholder: string;
  showDeniedFilter?: boolean;
  isDeniedFilterOn?: boolean;
  onDeniedFilterChange?: (event: React.MouseEvent<HTMLElement>, checked?: boolean) => void;
  onClearFilters: () => void;
  isClearFiltersDisabled: boolean;
  pageSizeOptions: number[];
  selectedPageSize: number;
  onPageSizeChange: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void;
  viewModels: ISuggestionViewModel[];
  useTableLayout: boolean;
  showMetadataInIdColumn: boolean;
  totalPages?: number;
  onToggleVote: SuggestionAction;
  onChangeStatus: (item: ISuggestionItem, status: string) => void;
  onDeleteSuggestion: SuggestionAction;
  onSubmitComment: SuggestionAction;
  onCommentDraftChange: (item: ISuggestionItem, value: string) => void;
  onDeleteComment: CommentAction;
  onToggleComments: (itemId: number) => void;
  onToggleCommentComposer: (itemId: number) => void;
  formatDateTime: (value: string) => string;
  statuses: string[];
  page: number;
  hasPrevious: boolean;
  hasNext: boolean;
  onPrevious: () => void;
  onNext: () => void;
}

const SuggestionSection: React.FC<ISuggestionSectionProps> = ({
  title,
  titleId,
  contentId,
  isLoading,
  isSectionLoading,
  searchValue,
  onSearchChange,
  categoryOptions,
  selectedCategoryKey,
  onCategoryChange,
  disableCategoryDropdown,
  subcategoryOptions,
  selectedSubcategoryKey,
  onSubcategoryChange,
  disableSubcategoryDropdown,
  subcategoryPlaceholder,
  showDeniedFilter,
  isDeniedFilterOn,
  onDeniedFilterChange,
  onClearFilters,
  isClearFiltersDisabled,
  pageSizeOptions,
  selectedPageSize,
  onPageSizeChange,
  viewModels,
  useTableLayout,
  showMetadataInIdColumn,
  totalPages,
  onToggleVote,
  onChangeStatus,
  onDeleteSuggestion,
  onSubmitComment,
  onCommentDraftChange,
  onDeleteComment,
  onToggleComments,
  onToggleCommentComposer,
  formatDateTime,
  statuses,
  page,
  hasPrevious,
  hasNext,
  onPrevious,
  onNext
}) => {
  const normalizedPageSizeOptions: number[] = React.useMemo(() => {
    const seen: Set<number> = new Set();
    const items: number[] = [];
    const addSize = (size: number | undefined): void => {
      if (typeof size !== 'number' || !Number.isFinite(size) || size <= 0) {
        return;
      }

      const normalized: number = Math.floor(size);

      if (seen.has(normalized)) {
        return;
      }

      seen.add(normalized);
      items.push(normalized);
    };

    pageSizeOptions.forEach((size) => addSize(size));
    addSize(selectedPageSize);

    return items;
  }, [pageSizeOptions, selectedPageSize]);

  return (
    <div className={styles.suggestionSection}>
      <div className={styles.sectionHeader}>
        <h3 id={titleId} className={styles.sectionTitle}>
          {title}
        </h3>
      </div>
      <div id={contentId} role="region" aria-labelledby={titleId} className={styles.sectionContent}>
        <div className={styles.filters}>
          <div className={styles.filterControls}>
            <TextField
              label={strings.SearchLabel}
              value={searchValue}
              onChange={onSearchChange}
              disabled={isLoading}
              placeholder={strings.SearchPlaceholder}
              className={styles.filterSearch}
            />
            <Dropdown
              label={strings.CategoryLabel}
              options={categoryOptions}
              selectedKey={selectedCategoryKey}
              onChange={onCategoryChange}
              disabled={isLoading || isSectionLoading || disableCategoryDropdown}
              className={styles.filterDropdown}
            />
            <Dropdown
              label={strings.SubcategoryLabel}
              options={subcategoryOptions}
              selectedKey={selectedSubcategoryKey}
              onChange={onSubcategoryChange}
              disabled={isLoading || isSectionLoading || disableSubcategoryDropdown}
              className={styles.filterDropdown}
              placeholder={subcategoryPlaceholder}
            />
            {showDeniedFilter && (
              <Toggle
                label={strings.ShowDeniedSuggestionsLabel}
                checked={isDeniedFilterOn === true}
                onChange={onDeniedFilterChange}
                disabled={isLoading || isSectionLoading}
                className={styles.filterToggle}
              />
            )}
            <Dropdown
              label={strings.ItemsPerPageLabel}
              options={normalizedPageSizeOptions.map((size) => ({
                key: size,
                text: size.toString()
              }))}
              selectedKey={selectedPageSize}
              onChange={onPageSizeChange}
              disabled={isLoading || isSectionLoading}
              className={styles.filterDropdown}
            />
            <DefaultButton
              text={strings.ClearFiltersButtonText}
              className={styles.filterButton}
              onClick={onClearFilters}
              disabled={isLoading || isSectionLoading || isClearFiltersDisabled}
            />
          </div>
        </div>
        {isLoading || isSectionLoading ? (
          <Spinner label={strings.LoadingSuggestionsLabel} size={SpinnerSize.large} />
        ) : (
          <div className={styles.suggestionResults}>
            <div className={styles.suggestionListScroll}>
              <SuggestionList
                viewModels={viewModels}
                useTableLayout={useTableLayout}
                showMetadataInIdColumn={showMetadataInIdColumn}
                isLoading={isLoading}
                onToggleVote={onToggleVote}
                onChangeStatus={onChangeStatus}
                onDeleteSuggestion={onDeleteSuggestion}
                onSubmitComment={onSubmitComment}
                onCommentDraftChange={onCommentDraftChange}
                onDeleteComment={onDeleteComment}
                onToggleComments={onToggleComments}
                onToggleCommentComposer={onToggleCommentComposer}
                formatDateTime={formatDateTime}
                statuses={statuses}
              />
            </div>
            <PaginationControls
              page={page}
              hasPrevious={hasPrevious}
              hasNext={hasNext}
              totalPages={totalPages}
              onPrevious={onPrevious}
              onNext={onNext}
            />
          </div>
        )}
      </div>
    </div>
  );
};

export default SuggestionSection;
