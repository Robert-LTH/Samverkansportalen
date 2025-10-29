/* eslint-disable max-lines */
import * as React from 'react';
import {
  PrimaryButton,
  DefaultButton,
  IconButton,
  ActionButton,
  Icon,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  TextField,
  Dropdown,
  type IDropdownOption
} from '@fluentui/react';
import { debounce } from '@microsoft/sp-lodash-subset';
import styles from './Samverkansportalen.module.scss';
import { DEFAULT_SUGGESTIONS_LIST_TITLE, type ISamverkansportalenProps } from './ISamverkansportalenProps';
import {
  type SuggestionCategory,
  type IGraphSuggestionItem,
  type IGraphSuggestionItemFields,
  type IGraphVoteItem,
  type IGraphSubcategoryItem,
  type IGraphCategoryItem,
  type IGraphCommentItem
} from '../services/GraphSuggestionsService';
import * as strings from 'SamverkansportalenWebPartStrings';

interface ISuggestionItem {
  id: number;
  title: string;
  description: string;
  votes: number;
  status: 'Active' | 'Done';
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

interface IVoteEntry {
  id: number;
  username: string;
  votes: number;
}

interface ISuggestionComment {
  id: number;
  text: string;
  author?: string;
  createdDateTime?: string;
}

interface ISubcategoryDefinition {
  key: string;
  title: string;
  category?: SuggestionCategory;
}

interface ISimilarSuggestionsQuery {
  title: string;
  description: string;
}

interface ISamverkansportalenState {
  activeSuggestions: IPaginatedSuggestionsState;
  completedSuggestions: IPaginatedSuggestionsState;
  isLoading: boolean;
  isActiveSuggestionsLoading: boolean;
  isCompletedSuggestionsLoading: boolean;
  newTitle: string;
  newDescription: string;
  newCategory: SuggestionCategory;
  newSubcategoryKey?: string;
  subcategories: ISubcategoryDefinition[];
  categories: SuggestionCategory[];
  availableVotes: number;
  activeFilter: IFilterState;
  completedFilter: IFilterState;
  similarSuggestions: IPaginatedSuggestionsState;
  isSimilarSuggestionsLoading: boolean;
  similarSuggestionsQuery: ISimilarSuggestionsQuery;
  selectedSimilarSuggestion?: ISuggestionItem;
  isSelectedSimilarSuggestionLoading: boolean;
  error?: string;
  success?: string;
  isAddSuggestionExpanded: boolean;
  isActiveSuggestionsExpanded: boolean;
  isCompletedSuggestionsExpanded: boolean;
  expandedCommentIds: number[];
  loadingCommentIds: number[];
}

interface IFilterState {
  searchQuery: string;
  category?: SuggestionCategory;
  subcategory?: string;
  suggestionId?: number;
}

interface IPaginatedSuggestionsState {
  items: ISuggestionItem[];
  page: number;
  currentToken?: string;
  nextToken?: string;
  previousTokens: (string | undefined)[];
}

interface ISuggestionInteractionState {
  hasVoted: boolean;
  disableVote: boolean;
  canAddComment: boolean;
  canMarkSuggestionAsDone: boolean;
  canDeleteSuggestion: boolean;
  isVotingAllowed: boolean;
}

interface ISuggestionCommentState {
  isExpanded: boolean;
  isLoading: boolean;
  hasLoaded: boolean;
  resolvedCount: number;
  comments: ISuggestionComment[];
  canAddComment: boolean;
  canDeleteComments: boolean;
  regionId: string;
  toggleId: string;
}

interface ISuggestionViewModel {
  item: ISuggestionItem;
  interaction: ISuggestionInteractionState;
  comment: ISuggestionCommentState;
}

type SuggestionAction = (item: ISuggestionItem) => void | Promise<void>;
type CommentAction = (item: ISuggestionItem, comment: ISuggestionComment) => void | Promise<void>;

interface ISectionHeaderProps {
  title: string;
  titleId: string;
  contentId: string;
  isExpanded: boolean;
  onToggle: () => void;
}

const SectionHeader: React.FC<ISectionHeaderProps> = ({ title, titleId, contentId, isExpanded, onToggle }) => (
  <div className={styles.sectionHeader}>
    <h3 id={titleId} className={styles.sectionTitle}>
      {title}
    </h3>
    <ActionButton
      className={styles.sectionToggleButton}
      iconProps={{ iconName: isExpanded ? 'ChevronUpSmall' : 'ChevronDownSmall' }}
      onClick={onToggle}
      aria-expanded={isExpanded}
      aria-controls={contentId}
    >
      {isExpanded ? strings.HideSectionLabel : strings.ShowSectionLabel}
    </ActionButton>
  </div>
);

interface IPaginationControlsProps {
  page: number;
  hasPrevious: boolean;
  hasNext: boolean;
  onPrevious: () => void;
  onNext: () => void;
}

const PaginationControls: React.FC<IPaginationControlsProps> = ({
  page,
  hasPrevious,
  hasNext,
  onPrevious,
  onNext
}) => {
  if (!hasPrevious && !hasNext && page <= 1) {
    return null;
  }

  return (
    <div className={styles.paginationControls}>
      <DefaultButton text={strings.PreviousButtonText} onClick={onPrevious} disabled={!hasPrevious} />
      <span className={styles.paginationInfo} aria-live="polite">
        {strings.PaginationPageLabel.replace('{0}', page.toString())}
      </span>
      <DefaultButton text={strings.NextButtonText} onClick={onNext} disabled={!hasNext} />
    </div>
  );
};

interface ISuggestionTimestampsProps {
  item: ISuggestionItem;
  formatDateTime: (value: string) => string;
}

const SuggestionTimestamps: React.FC<ISuggestionTimestampsProps> = ({ item, formatDateTime }) => {
  const entries: { label: string; value: string }[] = [];
  const { createdDateTime, lastModifiedDateTime, completedDateTime } = item;

  if (createdDateTime) {
    entries.push({ label: strings.CreatedLabel, value: createdDateTime });
  }

  const shouldShowLastModified: boolean = !!lastModifiedDateTime && !completedDateTime;

  if (shouldShowLastModified && lastModifiedDateTime) {
    entries.push({ label: strings.LastModifiedLabel, value: lastModifiedDateTime });
  }

  if (completedDateTime) {
    entries.push({ label: strings.CompletedLabel, value: completedDateTime });
  }

  if (entries.length === 0) {
    return null;
  }

  return (
    <div className={styles.timestampRow}>
      {entries.map((entry) => (
        <span key={entry.label} className={styles.timestampEntry}>
          <span className={styles.timestampLabel}>{entry.label}:</span>
          <span className={styles.timestampValue}>{formatDateTime(entry.value)}</span>
        </span>
      ))}
    </div>
  );
};

interface IActionButtonsProps {
  interaction: ISuggestionInteractionState;
  containerClassName: string;
  isLoading: boolean;
  onToggleVote: () => void;
  onMarkSuggestionAsDone: () => void;
  onDeleteSuggestion: () => void;
}

const ActionButtons: React.FC<IActionButtonsProps> = ({
  interaction,
  containerClassName,
  isLoading,
  onToggleVote,
  onMarkSuggestionAsDone,
  onDeleteSuggestion
}) => (
  <div className={containerClassName}>
    {interaction.isVotingAllowed ? (
      <PrimaryButton
        text={interaction.hasVoted ? strings.RemoveVoteButtonText : strings.VoteButtonText}
        onClick={onToggleVote}
        disabled={interaction.disableVote}
      />
    ) : (
      <DefaultButton text={strings.VotesClosedText} disabled />
    )}
    {interaction.canMarkSuggestionAsDone && (
      <DefaultButton text={strings.MarkAsDoneButtonText} onClick={onMarkSuggestionAsDone} disabled={isLoading} />
    )}
    {interaction.canDeleteSuggestion && (
      <IconButton
        iconProps={{ iconName: 'Delete' }}
        title={strings.RemoveSuggestionButtonLabel}
        ariaLabel={strings.RemoveSuggestionButtonLabel}
        onClick={onDeleteSuggestion}
        disabled={isLoading}
      />
    )}
  </div>
);

interface ICommentSectionProps {
  item: ISuggestionItem;
  comment: ISuggestionCommentState;
  onToggle: () => void;
  onAddComment: () => void;
  onDeleteComment: (comment: ISuggestionComment) => void;
  formatDateTime: (value: string) => string;
  isLoading: boolean;
}

const CommentSection: React.FC<ICommentSectionProps> = ({
  item,
  comment,
  onToggle,
  onAddComment,
  onDeleteComment,
  formatDateTime,
  isLoading
}) => (
  <div className={styles.commentSection}>
    <div className={styles.commentHeader}>
      <button
        type="button"
        id={comment.toggleId}
        className={styles.commentToggleButton}
        onClick={onToggle}
        aria-expanded={comment.isExpanded}
        aria-controls={comment.regionId}
      >
        <Icon iconName={comment.isExpanded ? 'ChevronDownSmall' : 'ChevronRightSmall'} className={styles.commentToggleIcon} />
        <span className={styles.commentHeading}>{strings.CommentsLabel}</span>
        <span className={styles.commentCount}>({comment.resolvedCount})</span>
      </button>
      {comment.canAddComment && (
        <DefaultButton
          className={styles.commentAddButton}
          text={strings.AddCommentButtonText}
          onClick={onAddComment}
          disabled={isLoading}
        />
      )}
    </div>
    <div
      id={comment.regionId}
      role="region"
      aria-labelledby={comment.toggleId}
      className={`${styles.commentContent} ${comment.isExpanded ? '' : styles.commentContentCollapsed}`}
      hidden={!comment.isExpanded}
    >
      {comment.isExpanded && (
        comment.isLoading ? (
          <Spinner label={strings.LoadingCommentsLabel} size={SpinnerSize.small} />
        ) : !comment.hasLoaded ? null : comment.comments.length === 0 ? (
          <p className={styles.commentEmpty}>{strings.NoCommentsLabel}</p>
        ) : (
          <ul className={styles.commentList}>
            {comment.comments.map((commentItem) => {
              const hasMeta: boolean = !!commentItem.author || !!commentItem.createdDateTime;

              return (
                <li key={commentItem.id} className={styles.commentItem}>
                  {(hasMeta || comment.canDeleteComments) && (
                    <div className={styles.commentItemTopRow}>
                      {hasMeta ? (
                        <div className={styles.commentMeta}>
                          {commentItem.author && <span className={styles.commentAuthor}>{commentItem.author}</span>}
                          {commentItem.createdDateTime && (
                            <span className={styles.commentTimestamp}>{formatDateTime(commentItem.createdDateTime)}</span>
                          )}
                        </div>
                      ) : (
                        <span className={styles.commentMetaPlaceholder} aria-hidden={true} />
                      )}
                      {comment.canDeleteComments && (
                        <IconButton
                          className={styles.commentDeleteButton}
                          iconProps={{ iconName: 'Delete' }}
                          title={strings.DeleteCommentButtonLabel}
                          ariaLabel={strings.DeleteCommentButtonLabel}
                          onClick={() => onDeleteComment(commentItem)}
                          disabled={isLoading}
                        />
                      )}
                    </div>
                  )}
                  <p className={styles.commentText}>{commentItem.text}</p>
                </li>
              );
            })}
          </ul>
        )
      )}
    </div>
  </div>
);

interface ISuggestionCardsProps {
  viewModels: ISuggestionViewModel[];
  onToggleVote: SuggestionAction;
  onMarkSuggestionAsDone: SuggestionAction;
  onDeleteSuggestion: SuggestionAction;
  onAddComment: SuggestionAction;
  onDeleteComment: CommentAction;
  onToggleComments: (itemId: number) => void;
  formatDateTime: (value: string) => string;
  isLoading: boolean;
}

const SuggestionCards: React.FC<ISuggestionCardsProps> = ({
  viewModels,
  onToggleVote,
  onMarkSuggestionAsDone,
  onDeleteSuggestion,
  onAddComment,
  onDeleteComment,
  onToggleComments,
  formatDateTime,
  isLoading
}) => (
  <ul className={styles.suggestionList}>
    {viewModels.map(({ item, interaction, comment }) => (
      <li key={item.id} className={styles.suggestionCard}>
        <div className={styles.cardHeader}>
          <div className={styles.cardText}>
            <div className={styles.cardMeta}>
              <span
                className={styles.entryId}
                aria-label={strings.EntryAriaLabel.replace('{0}', item.id.toString())}
              >
                #{item.id}
              </span>
              <span className={styles.categoryBadge}>{item.category}</span>
              {item.subcategory && <span className={styles.subcategoryBadge}>{item.subcategory}</span>}
            </div>
            <h4 className={styles.suggestionTitle}>{item.title}</h4>
            <SuggestionTimestamps item={item} formatDateTime={formatDateTime} />
            {item.description && <p className={styles.suggestionDescription}>{item.description}</p>}
          </div>
          <div
            className={styles.voteBadge}
            aria-label={`${item.votes} ${item.votes === 1 ? strings.VoteSingularLabel : strings.VotesLabel}`}
          >
            <span className={styles.voteNumber}>{item.votes}</span>
            <span className={styles.voteText}>{item.votes === 1 ? strings.VoteSingularLabel : strings.VotesLabel}</span>
          </div>
        </div>
        <ActionButtons
          interaction={interaction}
          containerClassName={styles.cardActions}
          isLoading={isLoading}
          onToggleVote={() => onToggleVote(item)}
          onMarkSuggestionAsDone={() => onMarkSuggestionAsDone(item)}
          onDeleteSuggestion={() => onDeleteSuggestion(item)}
        />
        <CommentSection
          item={item}
          comment={comment}
          onToggle={() => onToggleComments(item.id)}
          onAddComment={() => onAddComment(item)}
          onDeleteComment={(commentItem) => onDeleteComment(item, commentItem)}
          formatDateTime={formatDateTime}
          isLoading={isLoading}
        />
      </li>
    ))}
  </ul>
);

interface ISuggestionTableProps {
  viewModels: ISuggestionViewModel[];
  onToggleVote: SuggestionAction;
  onMarkSuggestionAsDone: SuggestionAction;
  onDeleteSuggestion: SuggestionAction;
  onAddComment: SuggestionAction;
  onDeleteComment: CommentAction;
  onToggleComments: (itemId: number) => void;
  formatDateTime: (value: string) => string;
  isLoading: boolean;
}

const SuggestionTable: React.FC<ISuggestionTableProps> = ({
  viewModels,
  onToggleVote,
  onMarkSuggestionAsDone,
  onDeleteSuggestion,
  onAddComment,
  onDeleteComment,
  onToggleComments,
  formatDateTime,
  isLoading
}) => (
  <div className={styles.tableWrapper}>
    <table className={styles.suggestionTable}>
      <thead>
        <tr>
          <th scope="col" className={styles.tableHeaderId}>
            #
          </th>
          <th scope="col" className={styles.tableHeaderSuggestion}>
            {strings.SuggestionTableSuggestionColumnLabel}
          </th>
          <th scope="col" className={styles.tableHeaderCategory}>
            {strings.CategoryLabel}
          </th>
          <th scope="col" className={styles.tableHeaderSubcategory}>
            {strings.SubcategoryLabel}
          </th>
          <th scope="col" className={styles.tableHeaderVotes}>
            {strings.VotesLabel}
          </th>
          <th scope="col" className={styles.tableHeaderActions}>
            {strings.SuggestionTableActionsColumnLabel}
          </th>
        </tr>
      </thead>
      <tbody>
        {viewModels.map(({ item, interaction, comment }) => (
          <React.Fragment key={item.id}>
            <tr className={styles.suggestionRow}>
              <td className={styles.tableCellId} data-label={strings.SuggestionTableEntryColumnLabel}>
                <span
                  className={styles.entryId}
                  aria-label={strings.EntryAriaLabel.replace('{0}', item.id.toString())}
                >
                  #{item.id}
                </span>
              </td>
              <td
                className={styles.tableCellSuggestion}
                data-label={strings.SuggestionTableSuggestionColumnLabel}
              >
                <h4 className={styles.suggestionTitle}>{item.title}</h4>
                <SuggestionTimestamps item={item} formatDateTime={formatDateTime} />
                {item.description && <p className={styles.suggestionDescription}>{item.description}</p>}
              </td>
              <td className={styles.tableCellCategory} data-label={strings.CategoryLabel}>
                <span className={styles.categoryBadge}>{item.category}</span>
              </td>
              <td className={styles.tableCellSubcategory} data-label={strings.SubcategoryLabel}>
                {item.subcategory ? (
                  <span className={styles.subcategoryBadge}>{item.subcategory}</span>
                ) : (
                  <span className={styles.subcategoryPlaceholder}>—</span>
                )}
              </td>
              <td className={styles.tableCellVotes} data-label={strings.VotesLabel}>
                <div
                  className={styles.voteBadge}
                  aria-label={`${item.votes} ${item.votes === 1 ? strings.VoteSingularLabel : strings.VotesLabel}`}
                >
                  <span className={styles.voteNumber}>{item.votes}</span>
                  <span className={styles.voteText}>
                    {item.votes === 1 ? strings.VoteSingularLabel : strings.VotesLabel}
                  </span>
                </div>
              </td>
              <td
                className={styles.tableCellActions}
                data-label={strings.SuggestionTableActionsColumnLabel}
              >
                <ActionButtons
                  interaction={interaction}
                  containerClassName={styles.tableActions}
                  isLoading={isLoading}
                  onToggleVote={() => onToggleVote(item)}
                  onMarkSuggestionAsDone={() => onMarkSuggestionAsDone(item)}
                  onDeleteSuggestion={() => onDeleteSuggestion(item)}
                />
              </td>
            </tr>
            <tr className={styles.metaRow}>
              <td
                className={styles.metaCell}
                colSpan={6}
                data-label={strings.SuggestionTableDetailsColumnLabel}
              >
                <div className={styles.metaContent}>
                  <SuggestionTimestamps item={item} formatDateTime={formatDateTime} />
                  <CommentSection
                    item={item}
                    comment={comment}
                    onToggle={() => onToggleComments(item.id)}
                    onAddComment={() => onAddComment(item)}
                    onDeleteComment={(commentItem) => onDeleteComment(item, commentItem)}
                    formatDateTime={formatDateTime}
                    isLoading={isLoading}
                  />
                </div>
              </td>
            </tr>
          </React.Fragment>
        ))}
      </tbody>
    </table>
  </div>
);

interface ISuggestionListProps {
  viewModels: ISuggestionViewModel[];
  useTableLayout: boolean;
  isLoading: boolean;
  onToggleVote: SuggestionAction;
  onMarkSuggestionAsDone: SuggestionAction;
  onDeleteSuggestion: SuggestionAction;
  onAddComment: SuggestionAction;
  onDeleteComment: CommentAction;
  onToggleComments: (itemId: number) => void;
  formatDateTime: (value: string) => string;
}

const SuggestionList: React.FC<ISuggestionListProps> = ({
  viewModels,
  useTableLayout,
  isLoading,
  onToggleVote,
  onMarkSuggestionAsDone,
  onDeleteSuggestion,
  onAddComment,
  onDeleteComment,
  onToggleComments,
  formatDateTime
}) => {
  if (viewModels.length === 0) {
    return <p className={styles.emptyState}>{strings.NoSuggestionsLabel}</p>;
  }

  return useTableLayout ? (
    <SuggestionTable
      viewModels={viewModels}
      onToggleVote={onToggleVote}
      onMarkSuggestionAsDone={onMarkSuggestionAsDone}
      onDeleteSuggestion={onDeleteSuggestion}
      onAddComment={onAddComment}
      onDeleteComment={onDeleteComment}
      onToggleComments={onToggleComments}
      formatDateTime={formatDateTime}
      isLoading={isLoading}
    />
  ) : (
    <SuggestionCards
      viewModels={viewModels}
      onToggleVote={onToggleVote}
      onMarkSuggestionAsDone={onMarkSuggestionAsDone}
      onDeleteSuggestion={onDeleteSuggestion}
      onAddComment={onAddComment}
      onDeleteComment={onDeleteComment}
      onToggleComments={onToggleComments}
      formatDateTime={formatDateTime}
      isLoading={isLoading}
    />
  );
};

interface ISuggestionSectionProps {
  title: string;
  titleId: string;
  contentId: string;
  isExpanded: boolean;
  onToggle: () => void;
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
  viewModels: ISuggestionViewModel[];
  useTableLayout: boolean;
  onToggleVote: SuggestionAction;
  onMarkSuggestionAsDone: SuggestionAction;
  onDeleteSuggestion: SuggestionAction;
  onAddComment: SuggestionAction;
  onDeleteComment: CommentAction;
  onToggleComments: (itemId: number) => void;
  formatDateTime: (value: string) => string;
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
  isExpanded,
  onToggle,
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
  viewModels,
  useTableLayout,
  onToggleVote,
  onMarkSuggestionAsDone,
  onDeleteSuggestion,
  onAddComment,
  onDeleteComment,
  onToggleComments,
  formatDateTime,
  page,
  hasPrevious,
  hasNext,
  onPrevious,
  onNext
}) => (
  <div className={styles.suggestionSection}>
    <SectionHeader
      title={title}
      titleId={titleId}
      contentId={contentId}
      isExpanded={isExpanded}
      onToggle={onToggle}
    />
    <div
      id={contentId}
      role="region"
      aria-labelledby={titleId}
      className={`${styles.sectionContent} ${isExpanded ? '' : styles.sectionContentCollapsed}`}
      hidden={!isExpanded}
    >
      {isExpanded && (
        <>
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
          </div>
          {isLoading || isSectionLoading ? (
            <Spinner label={strings.LoadingSuggestionsLabel} size={SpinnerSize.large} />
          ) : (
            <>
              <SuggestionList
                viewModels={viewModels}
                useTableLayout={useTableLayout}
                isLoading={isLoading}
                onToggleVote={onToggleVote}
                onMarkSuggestionAsDone={onMarkSuggestionAsDone}
                onDeleteSuggestion={onDeleteSuggestion}
                onAddComment={onAddComment}
                onDeleteComment={onDeleteComment}
                onToggleComments={onToggleComments}
                formatDateTime={formatDateTime}
              />
              <PaginationControls
                page={page}
                hasPrevious={hasPrevious}
                hasNext={hasNext}
                onPrevious={onPrevious}
                onNext={onNext}
              />
            </>
          )}
        </>
      )}
    </div>
  </div>
);

interface ISimilarSuggestionsProps {
  viewModels: ISuggestionViewModel[];
  isLoading: boolean;
  query: ISimilarSuggestionsQuery;
  page: number;
  hasPrevious: boolean;
  hasNext: boolean;
  onPrevious: () => void;
  onNext: () => void;
  onToggleVote: SuggestionAction;
  onMarkSuggestionAsDone: SuggestionAction;
  onDeleteSuggestion: SuggestionAction;
  onAddComment: SuggestionAction;
  onDeleteComment: CommentAction;
  onToggleComments: (itemId: number) => void;
  formatDateTime: (value: string) => string;
  isProcessing: boolean;
}

const SimilarSuggestions: React.FC<ISimilarSuggestionsProps> = ({
  viewModels,
  isLoading,
  query,
  page,
  hasPrevious,
  hasNext,
  onPrevious,
  onNext,
  onToggleVote,
  onMarkSuggestionAsDone,
  onDeleteSuggestion,
  onAddComment,
  onDeleteComment,
  onToggleComments,
  formatDateTime,
  isProcessing
}) => {
  const hasTitleQuery: boolean = query.title.length > 0;
  const hasDescriptionQuery: boolean = query.description.length > 0;

  if (!hasTitleQuery && !hasDescriptionQuery) {
    return null;
  }

  const querySegments: { key: string; content: React.ReactNode }[] = [];

  if (hasTitleQuery) {
    querySegments.push({
      key: 'title',
      content: (
        <>
          {strings.SimilarSuggestionsQueryTitleLabel}{' '}
          <span className={styles.similarSuggestionsQueryValue}>“{query.title}”</span>
        </>
      )
    });
  }

  if (hasDescriptionQuery) {
    querySegments.push({
      key: 'description',
      content: (
        <>
          {strings.SimilarSuggestionsQueryDescriptionLabel}{' '}
          <span className={styles.similarSuggestionsQueryValue}>“{query.description}”</span>
        </>
      )
    });
  }

  const hasResults: boolean = viewModels.length > 0;

  return (
    <div className={styles.similarSuggestions} aria-live="polite">
      <div className={styles.similarSuggestionsHeader}>
        <h4 className={styles.similarSuggestionsTitle}>{strings.SimilarSuggestionsTitle}</h4>
        {!isLoading && hasResults && (
          <span className={styles.similarSuggestionsSummary}>
            {viewModels.length === 1
              ? strings.SingleMatchingSuggestionLabel
              : strings.MultipleMatchingSuggestionsLabel.replace('{0}', viewModels.length.toString())}
          </span>
        )}
      </div>
      <p className={styles.similarSuggestionsQuery}>
        {strings.SimilarSuggestionsQueryPrefix}{' '}
        {querySegments.map((segment, index) => (
          <React.Fragment key={segment.key}>
            {index > 0 && (
              <>
                {' '}
                {strings.SimilarSuggestionsQuerySeparator}
                {' '}
              </>
            )}
            {segment.content}
          </React.Fragment>
        ))}
      </p>
      {isLoading ? (
        <Spinner label={strings.SimilarSuggestionsLoadingLabel} size={SpinnerSize.small} />
      ) : hasResults ? (
        <>
          <div className={styles.similarSuggestionsResults}>
            <SuggestionList
              viewModels={viewModels}
              useTableLayout={false}
              isLoading={isProcessing}
              onToggleVote={onToggleVote}
              onMarkSuggestionAsDone={onMarkSuggestionAsDone}
              onDeleteSuggestion={onDeleteSuggestion}
              onAddComment={onAddComment}
              onDeleteComment={onDeleteComment}
              onToggleComments={onToggleComments}
              formatDateTime={formatDateTime}
            />
          </div>
          <PaginationControls
            page={page}
            hasPrevious={hasPrevious}
            hasNext={hasNext}
            onPrevious={onPrevious}
            onNext={onNext}
          />
        </>
      ) : (
        <p className={styles.noSimilarSuggestions}>{strings.NoSimilarSuggestionsLabel}</p>
      )}
    </div>
  );
};

const MAX_VOTES_PER_USER: number = 5;
const FALLBACK_CATEGORIES: SuggestionCategory[] = [
  strings.DefaultCategoryChangeRequest,
  strings.DefaultCategoryWebinar,
  strings.DefaultCategoryArticle
];
const DEFAULT_SUGGESTION_CATEGORY: SuggestionCategory = FALLBACK_CATEGORIES[0];
const ALL_CATEGORY_FILTER_KEY: string = '__all_categories__';
const ALL_SUBCATEGORY_FILTER_KEY: string = '__all_subcategories__';
const SUGGESTIONS_PAGE_SIZE: number = 5;
const SIMILAR_SUGGESTIONS_DEBOUNCE_MS: number = 500;
const MIN_SIMILAR_SUGGESTION_QUERY_LENGTH: number = 3;
const MAX_SIMILAR_SUGGESTIONS: number = 5;
const EMPTY_SIMILAR_SUGGESTIONS_QUERY: ISimilarSuggestionsQuery = { title: '', description: '' };

export default class Samverkansportalen extends React.Component<ISamverkansportalenProps, ISamverkansportalenState> {
  private _isMounted: boolean = false;
  private _currentListId?: string;
  private _currentVotesListId?: string;
  private _currentCommentsListId?: string;
  private _currentSubcategoryListId?: string;
  private _currentCategoryListId?: string;
  private readonly _sectionIds: {
    add: { title: string; content: string };
    active: { title: string; content: string };
    completed: { title: string; content: string };
  };
  private readonly _commentSectionPrefix: string;
  private readonly _debouncedSimilarSuggestionsSearch: ReturnType<typeof debounce>;
  private _pendingSimilarSuggestionsQuery?: ISimilarSuggestionsQuery;

  public constructor(props: ISamverkansportalenProps) {
    super(props);

    const uniquePrefix: string = `samverkansportalen-${Math.random().toString(36).slice(2, 10)}`;
    this._sectionIds = {
      add: { title: `${uniquePrefix}-add-title`, content: `${uniquePrefix}-add-content` },
      active: { title: `${uniquePrefix}-active-title`, content: `${uniquePrefix}-active-content` },
      completed: {
        title: `${uniquePrefix}-completed-title`,
        content: `${uniquePrefix}-completed-content`
      }
    };
    this._commentSectionPrefix = `${uniquePrefix}-comment`;
    this._debouncedSimilarSuggestionsSearch = debounce((query: ISimilarSuggestionsQuery) => {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this._searchSimilarSuggestions(query);
    }, SIMILAR_SUGGESTIONS_DEBOUNCE_MS);

    this.state = {
      activeSuggestions: { items: [], page: 1, currentToken: undefined, nextToken: undefined, previousTokens: [] },
      completedSuggestions: { items: [], page: 1, currentToken: undefined, nextToken: undefined, previousTokens: [] },
      isLoading: false,
      isActiveSuggestionsLoading: false,
      isCompletedSuggestionsLoading: false,
      newTitle: '',
      newDescription: '',
      newCategory: DEFAULT_SUGGESTION_CATEGORY,
      newSubcategoryKey: undefined,
      subcategories: [],
      categories: [...FALLBACK_CATEGORIES],
      availableVotes: MAX_VOTES_PER_USER,
      activeFilter: {
        searchQuery: '',
        category: undefined,
        subcategory: undefined,
        suggestionId: undefined
      },
      completedFilter: {
        searchQuery: '',
        category: undefined,
        subcategory: undefined,
        suggestionId: undefined
      },
      similarSuggestions: {
        items: [],
        page: 1,
        currentToken: undefined,
        nextToken: undefined,
        previousTokens: []
      },
      isSimilarSuggestionsLoading: false,
      similarSuggestionsQuery: { ...EMPTY_SIMILAR_SUGGESTIONS_QUERY },
      selectedSimilarSuggestion: undefined,
      isSelectedSimilarSuggestionLoading: false,
      isAddSuggestionExpanded: true,
      isActiveSuggestionsExpanded: true,
      isCompletedSuggestionsExpanded: true,
      expandedCommentIds: [],
      loadingCommentIds: []
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._initialize();
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
    this._debouncedSimilarSuggestionsSearch.cancel();
  }

  public componentDidUpdate(prevProps: ISamverkansportalenProps): void {
    const listChanged: boolean = this._normalizeListTitle(prevProps.listTitle) !== this._listTitle;
    const voteListChanged: boolean =
      this._normalizeVoteListTitle(prevProps.voteListTitle, prevProps.listTitle) !== this._voteListTitle;
    const commentListChanged: boolean =
      this._normalizeCommentListTitle(prevProps.commentListTitle, prevProps.listTitle) !== this._commentListTitle;
    const subcategoryListChanged: boolean =
      this._normalizeOptionalListTitle(prevProps.subcategoryListTitle) !== this._subcategoryListTitle;
    const categoryListChanged: boolean =
      this._normalizeOptionalListTitle(prevProps.categoryListTitle) !== this._categoryListTitle;

    if (listChanged || voteListChanged || commentListChanged || subcategoryListChanged || categoryListChanged) {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this._initialize();
    }
  }

  public render(): React.ReactElement<ISamverkansportalenProps> {
    const {
      activeSuggestions,
      completedSuggestions,
      similarSuggestions,
      isLoading,
      isActiveSuggestionsLoading,
      isCompletedSuggestionsLoading,
      isSimilarSuggestionsLoading,
      availableVotes,
      newTitle,
      newDescription,
      newCategory,
      newSubcategoryKey,
      subcategories,
      categories,
      activeFilter,
      completedFilter,
      similarSuggestionsQuery,
      error,
      success,
      isAddSuggestionExpanded,
      isActiveSuggestionsExpanded,
      isCompletedSuggestionsExpanded
    } = this.state;

    const subcategoryOptions: IDropdownOption[] = this._getSubcategoryOptions(newCategory, subcategories);
    const categoryOptions: IDropdownOption[] = this._getCategoryOptions(categories);
    const filterCategoryOptions: IDropdownOption[] = this._getFilterCategoryOptions(categories);
    const activeFilterSubcategoryOptions: IDropdownOption[] = this._getFilterSubcategoryOptions(
      activeFilter.category,
      subcategories
    );
    const completedFilterSubcategoryOptions: IDropdownOption[] = this._getFilterSubcategoryOptions(
      completedFilter.category,
      subcategories
    );

    const isFilterCategoryLimited: boolean = filterCategoryOptions.length <= 1;
    const isActiveFilterSubcategoryLimited: boolean = activeFilterSubcategoryOptions.length <= 1;
    const isCompletedFilterSubcategoryLimited: boolean = completedFilterSubcategoryOptions.length <= 1;
    const activeFilterSubcategoryPlaceholder: string = isActiveFilterSubcategoryLimited
      ? strings.NoSubcategoriesAvailablePlaceholder
      : strings.SelectSubcategoryPlaceholder;
    const completedFilterSubcategoryPlaceholder: string = isCompletedFilterSubcategoryLimited
      ? strings.NoSubcategoriesAvailablePlaceholder
      : strings.SelectSubcategoryPlaceholder;

    const activeSuggestionViewModels: ISuggestionViewModel[] = this._createSuggestionViewModels(
      activeSuggestions.items,
      false
    );
    const completedSuggestionViewModels: ISuggestionViewModel[] = this._createSuggestionViewModels(
      completedSuggestions.items,
      true
    );
    const similarSuggestionViewModels: ISuggestionViewModel[] = this._createSuggestionViewModels(
      similarSuggestions.items,
      true,
      { allowVoting: true }
    );

    return (
      <section className={`${styles.samverkansportalen} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        <header className={styles.header}>
          <div>
            <h2 className={styles.title}>{this.props.headerTitle}</h2>
            <p className={styles.subtitle}>{this.props.headerSubtitle}</p>
          </div>
          <div className={styles.voteSummary} aria-live="polite">
            <span className={styles.voteLabel}>{strings.VotesRemainingLabel}</span>
            <span className={styles.voteValue}>{availableVotes} / {MAX_VOTES_PER_USER}</span>
          </div>
        </header>

        {error && (
          <MessageBar
            className={styles.messageBar}
            messageBarType={MessageBarType.error}
            isMultiline={false}
            onDismiss={this._dismissError}
          >
            {error}
          </MessageBar>
        )}

        {success && (
          <MessageBar
            className={styles.messageBar}
            messageBarType={MessageBarType.success}
            isMultiline={false}
            onDismiss={this._dismissSuccess}
          >
            {success}
          </MessageBar>
        )}

        <div className={styles.addSuggestion}>
          <SectionHeader
            title={strings.AddSuggestionSectionTitle}
            titleId={this._sectionIds.add.title}
            contentId={this._sectionIds.add.content}
            isExpanded={isAddSuggestionExpanded}
            onToggle={this._toggleAddSuggestionSection}
          />
          <div
            id={this._sectionIds.add.content}
            role="region"
            aria-labelledby={this._sectionIds.add.title}
            className={`${styles.sectionContent} ${
              isAddSuggestionExpanded ? '' : styles.sectionContentCollapsed
            }`}
            hidden={!isAddSuggestionExpanded}
          >
            {isAddSuggestionExpanded && (
              <div className={styles.addForm}>
                <TextField
                  label={strings.AddSuggestionTitleLabel}
                  required
                  value={newTitle}
                  onChange={this._onTitleChange}
                  disabled={isLoading}
                />
                <TextField
                  label={strings.AddSuggestionDetailsLabel}
                  multiline
                  rows={3}
                  value={newDescription}
                  onChange={this._onDescriptionChange}
                  disabled={isLoading}
                />
                <SimilarSuggestions
                  viewModels={similarSuggestionViewModels}
                  isLoading={isSimilarSuggestionsLoading}
                  query={similarSuggestionsQuery}
                  page={similarSuggestions.page}
                  hasPrevious={similarSuggestions.previousTokens.length > 0}
                  hasNext={!!similarSuggestions.nextToken}
                  onPrevious={this._goToPreviousSimilarPage}
                  onNext={this._goToNextSimilarPage}
                  onToggleVote={(item) => this._toggleVote(item)}
                  onMarkSuggestionAsDone={(item) => this._markSuggestionAsDone(item)}
                  onDeleteSuggestion={(item) => this._deleteSuggestion(item)}
                  onAddComment={(item) => this._addCommentToSuggestion(item)}
                  onDeleteComment={(item, comment) => this._deleteCommentFromSuggestion(item, comment)}
                  onToggleComments={(id) => this._toggleCommentsSection(id)}
                  formatDateTime={(value) => this._formatDateTime(value)}
                  isProcessing={isLoading}
                />
                <Dropdown
                  label={strings.CategoryLabel}
                  options={categoryOptions}
                  selectedKey={newCategory}
                  onChange={this._onCategoryChange}
                  disabled={isLoading || categoryOptions.length === 0}
                />
                <Dropdown
                  label={strings.SubcategoryLabel}
                  options={subcategoryOptions}
                  selectedKey={newSubcategoryKey}
                  onChange={this._onSubcategoryChange}
                  disabled={isLoading || subcategoryOptions.length === 0}
                  placeholder={
                    subcategoryOptions.length === 0
                      ? strings.NoSubcategoriesAvailablePlaceholder
                      : strings.SelectSubcategoryPlaceholder
                  }
                />
                <PrimaryButton
                  text={strings.SubmitSuggestionButtonText}
                  onClick={this._addSuggestion}
                  disabled={isLoading || newTitle.trim().length === 0}
                />
              </div>
            )}
          </div>
        </div>

        <SuggestionSection
          title={strings.ActiveSuggestionsSectionTitle}
          titleId={this._sectionIds.active.title}
          contentId={this._sectionIds.active.content}
          isExpanded={isActiveSuggestionsExpanded}
          onToggle={this._toggleActiveSection}
          isLoading={isLoading}
          isSectionLoading={isActiveSuggestionsLoading}
          searchValue={activeFilter.searchQuery}
          onSearchChange={this._onActiveSearchChange}
          categoryOptions={filterCategoryOptions}
          selectedCategoryKey={activeFilter.category ?? ALL_CATEGORY_FILTER_KEY}
          onCategoryChange={this._onActiveFilterCategoryChange}
          disableCategoryDropdown={isFilterCategoryLimited}
          subcategoryOptions={activeFilterSubcategoryOptions}
          selectedSubcategoryKey={activeFilter.subcategory ?? ALL_SUBCATEGORY_FILTER_KEY}
          onSubcategoryChange={this._onActiveFilterSubcategoryChange}
          disableSubcategoryDropdown={isActiveFilterSubcategoryLimited}
          subcategoryPlaceholder={activeFilterSubcategoryPlaceholder}
          viewModels={activeSuggestionViewModels}
          useTableLayout={this.props.useTableLayout === true}
          onToggleVote={(item) => this._toggleVote(item)}
          onMarkSuggestionAsDone={(item) => this._markSuggestionAsDone(item)}
          onDeleteSuggestion={(item) => this._deleteSuggestion(item)}
          onAddComment={(item) => this._addCommentToSuggestion(item)}
          onDeleteComment={(item, comment) => this._deleteCommentFromSuggestion(item, comment)}
          onToggleComments={(id) => this._toggleCommentsSection(id)}
          formatDateTime={(value) => this._formatDateTime(value)}
          page={activeSuggestions.page}
          hasPrevious={activeSuggestions.previousTokens.length > 0}
          hasNext={!!activeSuggestions.nextToken}
          onPrevious={this._goToPreviousActivePage}
          onNext={this._goToNextActivePage}
        />

        <SuggestionSection
          title={strings.CompletedSuggestionsSectionTitle}
          titleId={this._sectionIds.completed.title}
          contentId={this._sectionIds.completed.content}
          isExpanded={isCompletedSuggestionsExpanded}
          onToggle={this._toggleCompletedSection}
          isLoading={isLoading}
          isSectionLoading={isCompletedSuggestionsLoading}
          searchValue={completedFilter.searchQuery}
          onSearchChange={this._onCompletedSearchChange}
          categoryOptions={filterCategoryOptions}
          selectedCategoryKey={completedFilter.category ?? ALL_CATEGORY_FILTER_KEY}
          onCategoryChange={this._onCompletedFilterCategoryChange}
          disableCategoryDropdown={isFilterCategoryLimited}
          subcategoryOptions={completedFilterSubcategoryOptions}
          selectedSubcategoryKey={completedFilter.subcategory ?? ALL_SUBCATEGORY_FILTER_KEY}
          onSubcategoryChange={this._onCompletedFilterSubcategoryChange}
          disableSubcategoryDropdown={isCompletedFilterSubcategoryLimited}
          subcategoryPlaceholder={completedFilterSubcategoryPlaceholder}
          viewModels={completedSuggestionViewModels}
          useTableLayout={this.props.useTableLayout === true}
          onToggleVote={(item) => this._toggleVote(item)}
          onMarkSuggestionAsDone={(item) => this._markSuggestionAsDone(item)}
          onDeleteSuggestion={(item) => this._deleteSuggestion(item)}
          onAddComment={(item) => this._addCommentToSuggestion(item)}
          onDeleteComment={(item, comment) => this._deleteCommentFromSuggestion(item, comment)}
          onToggleComments={(id) => this._toggleCommentsSection(id)}
          formatDateTime={(value) => this._formatDateTime(value)}
          page={completedSuggestions.page}
          hasPrevious={completedSuggestions.previousTokens.length > 0}
          hasNext={!!completedSuggestions.nextToken}
          onPrevious={this._goToPreviousCompletedPage}
          onNext={this._goToNextCompletedPage}
        />
      </section>
    );
  }

  private _createSuggestionViewModels(
    items: ISuggestionItem[],
    readOnly: boolean,
    options: { allowVoting?: boolean } = {}
  ): ISuggestionViewModel[] {
    const noVotesRemaining: boolean = this.state.availableVotes <= 0;
    const normalizedUser: string | undefined = this._normalizeLoginName(this.props.userLoginName);
    const allowVoting: boolean = options.allowVoting === true;

    return items.map((item) => {
      const interaction: ISuggestionInteractionState = this._getInteractionState(
        item,
        readOnly,
        normalizedUser,
        noVotesRemaining,
        allowVoting
      );
      const isExpanded: boolean = this._isCommentSectionExpanded(item.id);
      const isLoadingComments: boolean = this.state.loadingCommentIds.indexOf(item.id) !== -1;
      const hasLoadedComments: boolean = item.areCommentsLoaded;
      const resolvedCommentCount: number = hasLoadedComments ? item.comments.length : item.commentCount;
      const renderedComments: ISuggestionComment[] = hasLoadedComments ? item.comments : [];
      const regionId: string = `${this._commentSectionPrefix}-${item.id}`;
      const toggleId: string = `${regionId}-toggle`;

      return {
        item,
        interaction,
        comment: {
          isExpanded,
          isLoading: isLoadingComments,
          hasLoaded: hasLoadedComments,
          resolvedCount: resolvedCommentCount,
          comments: renderedComments,
          canAddComment: interaction.canAddComment,
          canDeleteComments: this.props.isCurrentUserAdmin,
          regionId,
          toggleId
        }
      };
    });
  }

  private _goToPreviousActivePage = async (): Promise<void> => {
    const { activeSuggestions, activeFilter } = this.state;

    if (activeSuggestions.previousTokens.length === 0) {
      return;
    }

    const tokens: (string | undefined)[] = [...activeSuggestions.previousTokens];
    const previousToken: string | undefined = tokens.pop();

    this._updateState({ isActiveSuggestionsLoading: true, error: undefined, success: undefined });

    try {
      await this._fetchActiveSuggestions({
        page: Math.max(activeSuggestions.page - 1, 1),
        previousTokens: tokens,
        skipToken: previousToken,
        filter: activeFilter
      });
    } catch (error) {
      this._handleError('We could not load the previous page of active suggestions.', error);
    } finally {
      this._updateState({ isActiveSuggestionsLoading: false });
    }
  };

  private _goToNextActivePage = async (): Promise<void> => {
    const { activeSuggestions, activeFilter } = this.state;

    if (!activeSuggestions.nextToken) {
      return;
    }

    const tokens: (string | undefined)[] = [
      ...activeSuggestions.previousTokens,
      activeSuggestions.currentToken
    ];

    this._updateState({ isActiveSuggestionsLoading: true, error: undefined, success: undefined });

    try {
      await this._fetchActiveSuggestions({
        page: activeSuggestions.page + 1,
        previousTokens: tokens,
        skipToken: activeSuggestions.nextToken,
        filter: activeFilter
      });
    } catch (error) {
      this._handleError('We could not load more active suggestions. Please try again.', error);
    } finally {
      this._updateState({ isActiveSuggestionsLoading: false });
    }
  };

  private _goToPreviousCompletedPage = async (): Promise<void> => {
    const { completedSuggestions, completedFilter } = this.state;

    if (completedSuggestions.previousTokens.length === 0) {
      return;
    }

    const tokens: (string | undefined)[] = [...completedSuggestions.previousTokens];
    const previousToken: string | undefined = tokens.pop();

    this._updateState({ isCompletedSuggestionsLoading: true, error: undefined, success: undefined });

    try {
      await this._fetchCompletedSuggestions({
        page: Math.max(completedSuggestions.page - 1, 1),
        previousTokens: tokens,
        skipToken: previousToken,
        filter: completedFilter
      });
    } catch (error) {
      this._handleError('We could not load the previous page of completed suggestions.', error);
    } finally {
      this._updateState({ isCompletedSuggestionsLoading: false });
    }
  };

  private _goToNextCompletedPage = async (): Promise<void> => {
    const { completedSuggestions, completedFilter } = this.state;

    if (!completedSuggestions.nextToken) {
      return;
    }

    const tokens: (string | undefined)[] = [
      ...completedSuggestions.previousTokens,
      completedSuggestions.currentToken
    ];

    this._updateState({ isCompletedSuggestionsLoading: true, error: undefined, success: undefined });

    try {
      await this._fetchCompletedSuggestions({
        page: completedSuggestions.page + 1,
        previousTokens: tokens,
        skipToken: completedSuggestions.nextToken,
        filter: completedFilter
      });
    } catch (error) {
      this._handleError('We could not load more completed suggestions. Please try again.', error);
    } finally {
      this._updateState({ isCompletedSuggestionsLoading: false });
    }
  };

  private _goToPreviousSimilarPage = async (): Promise<void> => {
    const { similarSuggestions, similarSuggestionsQuery } = this.state;

    if (similarSuggestions.previousTokens.length === 0) {
      return;
    }

    const tokens: (string | undefined)[] = [...similarSuggestions.previousTokens];
    const previousToken: string | undefined = tokens.pop();

    await this._fetchSimilarSuggestions({
      page: Math.max(similarSuggestions.page - 1, 1),
      previousTokens: tokens,
      skipToken: previousToken,
      query: { ...similarSuggestionsQuery }
    });
  };

  private _goToNextSimilarPage = async (): Promise<void> => {
    const { similarSuggestions, similarSuggestionsQuery } = this.state;

    if (!similarSuggestions.nextToken) {
      return;
    }

    const tokens: (string | undefined)[] = [
      ...similarSuggestions.previousTokens,
      similarSuggestions.currentToken
    ];

    await this._fetchSimilarSuggestions({
      page: similarSuggestions.page + 1,
      previousTokens: tokens,
      skipToken: similarSuggestions.nextToken,
      query: { ...similarSuggestionsQuery }
    });
  };

  private _getInteractionState(
    item: ISuggestionItem,
    readOnly: boolean,
    normalizedUser: string | undefined,
    noVotesRemaining: boolean,
    allowVoting: boolean
  ): {
    hasVoted: boolean;
    disableVote: boolean;
    canAddComment: boolean;
    canMarkSuggestionAsDone: boolean;
    canDeleteSuggestion: boolean;
    isVotingAllowed: boolean;
  } {
    const hasVoted: boolean = !!normalizedUser && item.voters.indexOf(normalizedUser) !== -1;
    const isVotingAllowed: boolean = allowVoting || !readOnly;
    const disableVote: boolean =
      this.state.isLoading || !isVotingAllowed || item.status === 'Done' || (!hasVoted && noVotesRemaining);
    const canMarkSuggestionAsDone: boolean = this.props.isCurrentUserAdmin && !readOnly && item.status !== 'Done';
    const canDeleteSuggestion: boolean = this._canCurrentUserDeleteSuggestion(item);
    const canAddComment: boolean = !readOnly && item.status !== 'Done';

    return {
      hasVoted,
      disableVote,
      canAddComment,
      canMarkSuggestionAsDone,
      canDeleteSuggestion,
      isVotingAllowed
    };
  }

  private _formatDateTime(value: string): string {
    try {
      const parsed: Date = new Date(value);

      if (!Number.isNaN(parsed.getTime())) {
        return parsed.toLocaleString();
      }
    } catch (error) {
      console.warn('Failed to parse completion date.', error);
    }

    return value;
  }

  private _getSubcategoryOptions(
    category: SuggestionCategory,
    definitions: ISubcategoryDefinition[]
  ): IDropdownOption[] {
    return this._getSubcategoriesForCategory(category, definitions).map((definition) => ({
      key: definition.key,
      text: definition.title
    }));
  }

  private _getFilterSubcategoryOptions(
    category: SuggestionCategory | undefined,
    definitions: ISubcategoryDefinition[]
  ): IDropdownOption[] {
    const options: IDropdownOption[] = this._getSubcategoriesForCategory(category, definitions).map(
      (definition) => ({
        key: definition.title,
        text: definition.title
      })
    );

    return [{ key: ALL_SUBCATEGORY_FILTER_KEY, text: strings.AllSubcategoriesOptionLabel }, ...options];
  }

  private _getCategoryOptions(categories: SuggestionCategory[]): IDropdownOption[] {
    return categories.map((category) => ({ key: category, text: category }));
  }

  private _getFilterCategoryOptions(categories: SuggestionCategory[]): IDropdownOption[] {
    return [{ key: ALL_CATEGORY_FILTER_KEY, text: strings.AllCategoriesOptionLabel }, ...this._getCategoryOptions(categories)];
  }

  private _getSubcategoriesForCategory(
    category: SuggestionCategory | undefined,
    definitions: ISubcategoryDefinition[] = this.state.subcategories
  ): ISubcategoryDefinition[] {
    return definitions.filter((definition) => !definition.category || !category || definition.category === category);
  }

  private _normalizeFilterSubcategory(
    category: SuggestionCategory | undefined,
    preferredSubcategory: string | undefined,
    definitions: ISubcategoryDefinition[]
  ): string | undefined {
    if (!preferredSubcategory) {
      return undefined;
    }

    const availableTitles: string[] = this._getSubcategoriesForCategory(category, definitions).map((definition) =>
      definition.title.trim()
    );

    return availableTitles.indexOf(preferredSubcategory) !== -1 ? preferredSubcategory : undefined;
  }

  private _getValidSubcategoryKeyForCategory(
    category: SuggestionCategory,
    preferredKey: string | undefined,
    definitions: ISubcategoryDefinition[] = this.state.subcategories
  ): string | undefined {
    const options: ISubcategoryDefinition[] = this._getSubcategoriesForCategory(category, definitions);

    if (preferredKey && options.some((option) => option.key === preferredKey)) {
      return preferredKey;
    }

    return options.length > 0 ? options[0].key : undefined;
  }

  private _getSelectedSubcategoryDefinition(): ISubcategoryDefinition | undefined {
    const { newSubcategoryKey, subcategories } = this.state;

    if (!newSubcategoryKey) {
      return undefined;
    }

    return subcategories.find((definition) => definition.key === newSubcategoryKey);
  }

  private async _initialize(): Promise<void> {
    this._currentListId = undefined;
    this._currentVotesListId = undefined;
    this._currentCommentsListId = undefined;
    this._currentSubcategoryListId = undefined;
    this._currentCategoryListId = undefined;
    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      await this._ensureLists();
      await this._ensureCategoryList();
      await this._ensureSubcategoryList();
      await this._loadSuggestions();
    } catch (error) {
      const message: string =
        error instanceof Error && error.message.includes('category list')
          ? 'We could not load the configured category list. Please verify the configuration or reset it to use the default categories.'
          : error instanceof Error && error.message.includes('subcategory list')
          ? 'We could not load the configured subcategory list. Please verify the configuration or remove it.'
          : 'We could not load the suggestions list. Please refresh the page or contact your administrator.';
      this._handleError(message, error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private async _ensureLists(): Promise<void> {
    const listTitle: string = this._listTitle;
    const voteListTitle: string = this._voteListTitle;
    const commentListTitle: string = this._commentListTitle;
    const result = await this.props.graphService.ensureList(listTitle);
    this._currentListId = result.id;
    this._flushPendingSimilarSuggestionsSearch();

    const votesResult = await this.props.graphService.ensureVoteList(voteListTitle);
    this._currentVotesListId = votesResult.id;

    const commentsResult = await this.props.graphService.ensureCommentList(commentListTitle);
    this._currentCommentsListId = commentsResult.id;
  }

  private async _ensureCategoryList(): Promise<void> {
    this._currentCategoryListId = undefined;

    const listTitle: string | undefined = this._categoryListTitle;

    if (!listTitle) {
      this._applyCategories(FALLBACK_CATEGORIES);
      return;
    }

    const listInfo = await this.props.graphService.getListByTitle(listTitle);

    if (!listInfo) {
      throw new Error(`Failed to load the category list "${listTitle}".`);
    }

    this._currentCategoryListId = listInfo.id;
    await this._loadCategories();
  }

  private async _ensureSubcategoryList(): Promise<void> {
    this._currentSubcategoryListId = undefined;
    this._updateState({ subcategories: [], newSubcategoryKey: undefined });

    const listTitle: string | undefined = this._subcategoryListTitle;

    if (!listTitle) {
      return;
    }

    const listInfo = await this.props.graphService.getListByTitle(listTitle);

    if (!listInfo) {
      throw new Error(`Failed to load the subcategory list "${listTitle}".`);
    }

    this._currentSubcategoryListId = listInfo.id;
    await this._loadSubcategories();
  }

  private async _loadSuggestions(): Promise<void> {
    this._updateState({
      isActiveSuggestionsLoading: true,
      isCompletedSuggestionsLoading: true
    });

    try {
      await Promise.all([
        this._fetchActiveSuggestions({
          page: 1,
          previousTokens: [],
          skipToken: undefined,
          filter: this.state.activeFilter
        }),
        this._fetchCompletedSuggestions({
          page: 1,
          previousTokens: [],
          skipToken: undefined,
          filter: this.state.completedFilter
        }),
        this._loadAvailableVotes()
      ]);
    } finally {
      this._updateState({
        isActiveSuggestionsLoading: false,
        isCompletedSuggestionsLoading: false
      });
    }
  }

  private async _loadCategories(): Promise<void> {
    const listId: string = this._getResolvedCategoryListId();
    const itemsFromGraph: IGraphCategoryItem[] = await this.props.graphService.getCategoryItems(listId);

    const definitions: SuggestionCategory[] = itemsFromGraph
      .map((item) => {
        const rawTitle: unknown = item.fields?.Title;
        if (typeof rawTitle !== 'string') {
          return undefined;
        }

        const trimmed: string = rawTitle.trim();
        return trimmed.length > 0 ? trimmed : undefined;
      })
      .filter((value): value is SuggestionCategory => !!value);

    this._applyCategories(definitions);
  }

  private async _loadSubcategories(): Promise<void> {
    const listId: string = this._getResolvedSubcategoryListId();
    const itemsFromGraph: IGraphSubcategoryItem[] = await this.props.graphService.getSubcategoryItems(listId);

    const definitions: ISubcategoryDefinition[] = itemsFromGraph
      .map((item) => {
        const fields = item.fields ?? {};
        const rawTitle: unknown = (fields as { Title?: unknown }).Title;

        if (typeof rawTitle !== 'string') {
          return undefined;
        }

        const trimmedTitle: string = rawTitle.trim();

        if (!trimmedTitle) {
          return undefined;
        }

        const rawCategory: unknown = (fields as { Category?: unknown }).Category;
        const normalizedCategory: SuggestionCategory | undefined = this._tryNormalizeCategory(rawCategory);

        return {
          key: item.id.toString(),
          title: trimmedTitle,
          category: normalizedCategory
        } as ISubcategoryDefinition;
      })
      .filter((definition): definition is ISubcategoryDefinition => !!definition)
      .sort((a, b) => a.title.localeCompare(b.title));

    const nextSubcategoryKey: string | undefined = this._getValidSubcategoryKeyForCategory(
      this.state.newCategory,
      this.state.newSubcategoryKey,
      definitions
    );

    const nextActiveFilterSubcategory: string | undefined = this._normalizeFilterSubcategory(
      this.state.activeFilter.category,
      this.state.activeFilter.subcategory,
      definitions
    );

    const nextCompletedFilterSubcategory: string | undefined = this._normalizeFilterSubcategory(
      this.state.completedFilter.category,
      this.state.completedFilter.subcategory,
      definitions
    );

    this._updateState({
      subcategories: definitions,
      newSubcategoryKey: nextSubcategoryKey,
      activeFilter: { ...this.state.activeFilter, subcategory: nextActiveFilterSubcategory },
      completedFilter: { ...this.state.completedFilter, subcategory: nextCompletedFilterSubcategory }
    });
  }

  private _applyCategories(definitions: SuggestionCategory[]): void {
    const normalized: SuggestionCategory[] = this._normalizeCategoryList(definitions);
    const categories: SuggestionCategory[] = normalized.length > 0 ? normalized : [...FALLBACK_CATEGORIES];

    const nextCategory: SuggestionCategory =
      this._findCategoryMatch(this.state.newCategory, categories) ?? this._getDefaultCategory(categories);

    const nextActiveFilterCategory: SuggestionCategory | undefined = this._findCategoryMatch(
      this.state.activeFilter.category,
      categories
    );
    const nextCompletedFilterCategory: SuggestionCategory | undefined = this._findCategoryMatch(
      this.state.completedFilter.category,
      categories
    );

    const nextSubcategoryKey: string | undefined = this._getValidSubcategoryKeyForCategory(
      nextCategory,
      this.state.newSubcategoryKey,
      this.state.subcategories
    );

    const nextActiveFilterSubcategory: string | undefined = this._normalizeFilterSubcategory(
      nextActiveFilterCategory,
      this.state.activeFilter.subcategory,
      this.state.subcategories
    );

    const nextCompletedFilterSubcategory: string | undefined = this._normalizeFilterSubcategory(
      nextCompletedFilterCategory,
      this.state.completedFilter.subcategory,
      this.state.subcategories
    );

    this._updateState({
      categories,
      newCategory: nextCategory,
      newSubcategoryKey: nextSubcategoryKey,
      activeFilter: {
        ...this.state.activeFilter,
        category: nextActiveFilterCategory,
        subcategory: nextActiveFilterSubcategory
      },
      completedFilter: {
        ...this.state.completedFilter,
        category: nextCompletedFilterCategory,
        subcategory: nextCompletedFilterSubcategory
      }
    });
  }

  private _normalizeCategoryList(values: SuggestionCategory[]): SuggestionCategory[] {
    const seen: Set<string> = new Set();
    const normalized: SuggestionCategory[] = [];

    values.forEach((value) => {
      const trimmed: string = value.trim();

      if (!trimmed) {
        return;
      }

      const key: string = trimmed.toLowerCase();

      if (seen.has(key)) {
        return;
      }

      seen.add(key);
      normalized.push(trimmed);
    });

    normalized.sort((a, b) => a.localeCompare(b));
    return normalized;
  }

  private _findCategoryMatch(
    value: SuggestionCategory | undefined,
    categories: SuggestionCategory[]
  ): SuggestionCategory | undefined {
    if (!value) {
      return undefined;
    }

    const normalized: string = value.trim();

    if (!normalized) {
      return undefined;
    }

    const lower: string = normalized.toLowerCase();
    return categories.find((category) => category.toLowerCase() === lower);
  }

  private _getDefaultCategory(categories: SuggestionCategory[]): SuggestionCategory {
    return categories[0] ?? DEFAULT_SUGGESTION_CATEGORY;
  }

  private async _fetchActiveSuggestions(options: {
    page: number;
    previousTokens: (string | undefined)[];
    skipToken?: string;
    filter?: IFilterState;
  }): Promise<void> {
    const filter: IFilterState = options.filter ?? this.state.activeFilter;
    const hasSpecificSuggestion: boolean = typeof filter.suggestionId === 'number';
    const effectiveSkipToken: string | undefined = hasSpecificSuggestion ? undefined : options.skipToken;
    const effectivePreviousTokens: (string | undefined)[] = hasSpecificSuggestion
      ? []
      : options.previousTokens;

    const { items, nextToken } = await this._getSuggestionsPage('Active', effectiveSkipToken, filter);

    if (!hasSpecificSuggestion && items.length === 0 && effectivePreviousTokens.length > 0) {
      const tokens: (string | undefined)[] = [...effectivePreviousTokens];
      const previousToken: string | undefined = tokens.pop();

      await this._fetchActiveSuggestions({
        page: Math.max(options.page - 1, 1),
        previousTokens: tokens,
        skipToken: previousToken,
        filter
      });
      return;
    }

    this._updateState({
      activeSuggestions: {
        items,
        page: hasSpecificSuggestion ? 1 : options.page,
        currentToken: hasSpecificSuggestion ? undefined : effectiveSkipToken,
        nextToken: hasSpecificSuggestion ? undefined : nextToken,
        previousTokens: hasSpecificSuggestion ? [] : effectivePreviousTokens
      },
      activeFilter: filter,
      isActiveSuggestionsLoading: false
    });
  }

  private async _fetchCompletedSuggestions(options: {
    page: number;
    previousTokens: (string | undefined)[];
    skipToken?: string;
    filter?: IFilterState;
  }): Promise<void> {
    const filter: IFilterState = options.filter ?? this.state.completedFilter;
    const hasSpecificSuggestion: boolean = typeof filter.suggestionId === 'number';
    const effectiveSkipToken: string | undefined = hasSpecificSuggestion ? undefined : options.skipToken;
    const effectivePreviousTokens: (string | undefined)[] = hasSpecificSuggestion
      ? []
      : options.previousTokens;

    const { items, nextToken } = await this._getSuggestionsPage('Done', effectiveSkipToken, filter);

    if (!hasSpecificSuggestion && items.length === 0 && effectivePreviousTokens.length > 0) {
      const tokens: (string | undefined)[] = [...effectivePreviousTokens];
      const previousToken: string | undefined = tokens.pop();

      await this._fetchCompletedSuggestions({
        page: Math.max(options.page - 1, 1),
        previousTokens: tokens,
        skipToken: previousToken,
        filter
      });
      return;
    }

    this._updateState({
      completedSuggestions: {
        items,
        page: hasSpecificSuggestion ? 1 : options.page,
        currentToken: hasSpecificSuggestion ? undefined : effectiveSkipToken,
        nextToken: hasSpecificSuggestion ? undefined : nextToken,
        previousTokens: hasSpecificSuggestion ? [] : effectivePreviousTokens
      },
      completedFilter: filter,
      isCompletedSuggestionsLoading: false
    });
  }

  private async _refreshActiveSuggestions(): Promise<void> {
    const { activeSuggestions, activeFilter } = this.state;

    this._updateState({ isActiveSuggestionsLoading: true });

    try {
      await this._fetchActiveSuggestions({
        page: activeSuggestions.page,
        previousTokens: activeSuggestions.previousTokens,
        skipToken: activeSuggestions.currentToken,
        filter: activeFilter
      });
    } finally {
      this._updateState({ isActiveSuggestionsLoading: false });
    }
  }

  private async _refreshCompletedSuggestions(): Promise<void> {
    const { completedSuggestions, completedFilter } = this.state;

    this._updateState({ isCompletedSuggestionsLoading: true });

    try {
      await this._fetchCompletedSuggestions({
        page: completedSuggestions.page,
        previousTokens: completedSuggestions.previousTokens,
        skipToken: completedSuggestions.currentToken,
        filter: completedFilter
      });
    } finally {
      this._updateState({ isCompletedSuggestionsLoading: false });
    }
  }

  private async _getSuggestionsPage(
    status: 'Active' | 'Done',
    skipToken: string | undefined,
    filter: IFilterState
  ): Promise<{ items: ISuggestionItem[]; nextToken?: string }> {
    const listId: string = this._getResolvedListId();
    const response = await this.props.graphService.getSuggestionItems(listId, {
      status,
      top: SUGGESTIONS_PAGE_SIZE,
      skipToken,
      category: filter.category,
      subcategory: filter.subcategory,
      searchQuery: filter.searchQuery,
      suggestionIds:
        typeof filter.suggestionId === 'number' ? [filter.suggestionId] : undefined,
      orderBy: status === 'Done' ? 'fields/CompletedDateTime desc' : 'createdDateTime desc'
    });

    const suggestionIds: number[] = response.items
      .map((entry) => this._parseNumericId(entry.fields.id ?? (entry.fields as { Id?: unknown }).Id))
      .filter((value): value is number => typeof value === 'number');

    let votesBySuggestion: Map<number, IVoteEntry[]> = new Map();

    if (status === 'Active' && suggestionIds.length > 0) {
      const voteListId: string = this._getResolvedVotesListId();
      const voteItems: IGraphVoteItem[] = await this.props.graphService.getVoteItems(voteListId, {
        suggestionIds
      });
      votesBySuggestion = this._groupVotesBySuggestion(voteItems);
    }

    let commentCounts: Map<number, number> = new Map();

    if (suggestionIds.length > 0 && this._currentCommentsListId) {
      const commentListId: string = this._getResolvedCommentsListId();
      commentCounts = await this.props.graphService.getCommentCounts(commentListId, {
        suggestionIds
      });
    }

    const items: ISuggestionItem[] = this._mapGraphItemsToSuggestions(
      response.items,
      votesBySuggestion,
      commentCounts
    );
    return { items, nextToken: response.nextToken };
  }

  private _mapGraphItemsToSuggestions(
    graphItems: IGraphSuggestionItem[],
    votesBySuggestion: Map<number, IVoteEntry[]>,
    commentCounts: Map<number, number>
  ): ISuggestionItem[] {
    return graphItems
      .map((entry) => {
        const fields: IGraphSuggestionItemFields = entry.fields;
        const rawId: unknown = fields.id ?? (fields as { Id?: unknown }).Id;
        const suggestionId: number | undefined = this._parseNumericId(rawId);

        if (typeof suggestionId !== 'number') {
          return undefined;
        }

        const voteEntries: IVoteEntry[] = votesBySuggestion.get(suggestionId) ?? [];
        const storedVotes: number = this._parseVotes(fields.Votes);
        const status: 'Active' | 'Done' = fields.Status === 'Done' ? 'Done' : 'Active';
        const liveVotes: number = voteEntries.reduce((total, vote) => total + vote.votes, 0);
        const votes: number = status === 'Done' ? Math.max(liveVotes, storedVotes) : liveVotes;
        const createdDateTime: string | undefined =
          typeof entry.createdDateTime === 'string' && entry.createdDateTime.trim().length > 0
            ? entry.createdDateTime.trim()
            : undefined;
        const lastModifiedDateTime: string | undefined =
          typeof entry.lastModifiedDateTime === 'string' && entry.lastModifiedDateTime.trim().length > 0
            ? entry.lastModifiedDateTime.trim()
            : undefined;
        const completedDateTime: string | undefined =
          typeof fields.CompletedDateTime === 'string' && fields.CompletedDateTime.trim().length > 0
            ? fields.CompletedDateTime.trim()
            : undefined;
        const commentCount: number = commentCounts.get(suggestionId) ?? 0;

        return {
          id: suggestionId,
          title:
            typeof fields.Title === 'string' && fields.Title.trim().length > 0
              ? fields.Title
              : 'Untitled suggestion',
          description: typeof fields.Details === 'string' ? fields.Details : '',
          votes,
          status,
          category: this._normalizeCategory(fields.Category),
          subcategory:
            typeof fields.Subcategory === 'string' && fields.Subcategory.trim().length > 0
              ? fields.Subcategory.trim()
              : undefined,
          voters: voteEntries.map((vote) => vote.username),
          createdByLoginName: this._normalizeLoginName(entry.createdByUserPrincipalName),
          createdDateTime,
          lastModifiedDateTime,
          completedDateTime,
          voteEntries,
          commentCount,
          comments: [],
          areCommentsLoaded: commentCount === 0
        } as ISuggestionItem;
      })
      .filter((item): item is ISuggestionItem => !!item);
  }

  private _groupVotesBySuggestion(voteItems: IGraphVoteItem[]): Map<number, IVoteEntry[]> {
    const votesBySuggestion: Map<number, IVoteEntry[]> = new Map();

    voteItems.forEach((entry: IGraphVoteItem) => {
      const fields = entry.fields ?? {};
      const suggestionId: number | undefined = this._parseNumericId(
        (fields as { SuggestionId?: unknown }).SuggestionId
      );
      const rawUsername: unknown = (fields as { Username?: unknown }).Username;
      const normalizedUsername: string | undefined = this._normalizeLoginName(
        typeof rawUsername === 'string' ? rawUsername : undefined
      );
      const votes: number = this._parseVotes((fields as { Votes?: unknown }).Votes);

      if (!suggestionId || !normalizedUsername || votes <= 0) {
        return;
      }

      const entriesForSuggestion: IVoteEntry[] = votesBySuggestion.get(suggestionId) ?? [];
      entriesForSuggestion.push({
        id: entry.id,
        username: normalizedUsername,
        votes
      });
      votesBySuggestion.set(suggestionId, entriesForSuggestion);
    });

    return votesBySuggestion;
  }

  private _groupCommentsBySuggestion(commentItems: IGraphCommentItem[]): Map<number, ISuggestionComment[]> {
    const commentsBySuggestion: Map<number, ISuggestionComment[]> = new Map();

    commentItems.forEach((entry) => {
      const fields = entry.fields ?? {};
      const suggestionId: number | undefined = this._parseNumericId(
        (fields as { SuggestionId?: unknown }).SuggestionId
      );
      const rawComment: unknown =
        (fields as { Comment?: unknown }).Comment ?? (fields as { Title?: unknown }).Title;
      const commentText: string | undefined =
        typeof rawComment === 'string' && rawComment.trim().length > 0 ? rawComment.trim() : undefined;

      if (!suggestionId || !commentText) {
        return;
      }

      const createdDateTime: string | undefined =
        typeof entry.createdDateTime === 'string' && entry.createdDateTime.trim().length > 0
          ? entry.createdDateTime.trim()
          : undefined;

      const displayName: string | undefined =
        typeof entry.createdByUserDisplayName === 'string' && entry.createdByUserDisplayName.trim().length > 0
          ? entry.createdByUserDisplayName.trim()
          : undefined;
      const principalName: string | undefined =
        typeof entry.createdByUserPrincipalName === 'string' && entry.createdByUserPrincipalName.trim().length > 0
          ? entry.createdByUserPrincipalName.trim()
          : undefined;
      const author: string | undefined = displayName ?? principalName;

      const existing: ISuggestionComment[] = commentsBySuggestion.get(suggestionId) ?? [];
      existing.push({
        id: entry.id,
        text: commentText,
        author,
        createdDateTime
      });
      commentsBySuggestion.set(suggestionId, existing);
    });

    commentsBySuggestion.forEach((comments, key) => {
      const sorted: ISuggestionComment[] = [...comments].sort((a, b) => {
        if (!a.createdDateTime && !b.createdDateTime) {
          return a.id - b.id;
        }

        if (!a.createdDateTime) {
          return -1;
        }

        if (!b.createdDateTime) {
          return 1;
        }

        return new Date(a.createdDateTime).getTime() - new Date(b.createdDateTime).getTime();
      });
      commentsBySuggestion.set(key, sorted);
    });

    return commentsBySuggestion;
  }

  private async _loadAvailableVotes(): Promise<void> {
    const normalizedUser: string | undefined = this._normalizeLoginName(this.props.userLoginName);

    if (!normalizedUser) {
      this._updateState({ availableVotes: MAX_VOTES_PER_USER });
      return;
    }

    const voteListId: string = this._getResolvedVotesListId();
    const voteItems: IGraphVoteItem[] = await this.props.graphService.getVoteItems(voteListId, {
      username: normalizedUser
    });

    const usedVotes: number = voteItems.reduce((total, entry) => {
      const votes: number = this._parseVotes(entry.fields?.Votes);
      return total + votes;
    }, 0);

    const availableVotes: number = Math.max(MAX_VOTES_PER_USER - usedVotes, 0);
    this._updateState({ availableVotes });
  }

  private _onTitleChange = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this._updateState({ newTitle: newValue ?? '' }, () => {
      this._handleSimilarSuggestionsInput(this.state.newTitle, this.state.newDescription);
    });
  };

  private _onDescriptionChange = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this._updateState({ newDescription: newValue ?? '' }, () => {
      this._handleSimilarSuggestionsInput(this.state.newTitle, this.state.newDescription);
    });
  };

  private _areSimilarSuggestionQueriesEqual(
    left: ISimilarSuggestionsQuery,
    right: ISimilarSuggestionsQuery
  ): boolean {
    return left.title === right.title && left.description === right.description;
  }

  private _handleSimilarSuggestionsInput(title: string, description: string): void {
    const normalizedTitle: string = (title ?? '').replace(/\s+/g, ' ').trim();
    const normalizedDescription: string = (description ?? '').replace(/\s+/g, ' ').trim();
    const hasTitleQuery: boolean = normalizedTitle.length >= MIN_SIMILAR_SUGGESTION_QUERY_LENGTH;
    const hasDescriptionQuery: boolean =
      normalizedDescription.length >= MIN_SIMILAR_SUGGESTION_QUERY_LENGTH;

    if (!hasTitleQuery && !hasDescriptionQuery) {
      this._debouncedSimilarSuggestionsSearch.cancel();
      const previousSelectedId: number | undefined = this.state.selectedSimilarSuggestion?.id;
      const nextExpandedCommentIds: number[] =
        typeof previousSelectedId === 'number'
          ? this.state.expandedCommentIds.filter((id) => id !== previousSelectedId)
          : this.state.expandedCommentIds;
      const nextLoadingCommentIds: number[] =
        typeof previousSelectedId === 'number'
          ? this.state.loadingCommentIds.filter((id) => id !== previousSelectedId)
          : this.state.loadingCommentIds;
      this._updateState({
        similarSuggestions: {
          items: [],
          page: 1,
          currentToken: undefined,
          nextToken: undefined,
          previousTokens: []
        },
        similarSuggestionsQuery: { ...EMPTY_SIMILAR_SUGGESTIONS_QUERY },
        isSimilarSuggestionsLoading: false,
        selectedSimilarSuggestion: undefined,
        isSelectedSimilarSuggestionLoading: false,
        expandedCommentIds: nextExpandedCommentIds,
        loadingCommentIds: nextLoadingCommentIds
      });
      this._pendingSimilarSuggestionsQuery = undefined;
      return;
    }

    const nextQuery: ISimilarSuggestionsQuery = {
      title: hasTitleQuery ? normalizedTitle : '',
      description: hasDescriptionQuery ? normalizedDescription : ''
    };

    if (
      this._areSimilarSuggestionQueriesEqual(nextQuery, this.state.similarSuggestionsQuery) &&
      !this.state.isSimilarSuggestionsLoading &&
      this.state.similarSuggestions.items.length > 0
    ) {
      return;
    }

    this._pendingSimilarSuggestionsQuery = nextQuery;

    if (!this._currentListId) {
      return;
    }

    this._pendingSimilarSuggestionsQuery = undefined;
    this._debouncedSimilarSuggestionsSearch(nextQuery);
  }

  private async _searchSimilarSuggestions(query: ISimilarSuggestionsQuery): Promise<void> {
    if (!this._isMounted) {
      return;
    }

    this._pendingSimilarSuggestionsQuery = undefined;

    const normalizedTitle: string = (query.title ?? '').replace(/\s+/g, ' ').trim();
    const normalizedDescription: string = (query.description ?? '').replace(/\s+/g, ' ').trim();
    const effectiveQuery: ISimilarSuggestionsQuery = {
      title:
        normalizedTitle.length >= MIN_SIMILAR_SUGGESTION_QUERY_LENGTH ? normalizedTitle : '',
      description:
        normalizedDescription.length >= MIN_SIMILAR_SUGGESTION_QUERY_LENGTH
          ? normalizedDescription
          : ''
    };

    if (!effectiveQuery.title && !effectiveQuery.description) {
      this._updateState({
        similarSuggestions: {
          items: [],
          page: 1,
          currentToken: undefined,
          nextToken: undefined,
          previousTokens: []
        },
        similarSuggestionsQuery: { ...EMPTY_SIMILAR_SUGGESTIONS_QUERY },
        isSimilarSuggestionsLoading: false
      });
      return;
    }

    await this._fetchSimilarSuggestions({
      page: 1,
      previousTokens: [],
      skipToken: undefined,
      query: effectiveQuery
    });
  }

  private async _fetchSimilarSuggestions(options: {
    page: number;
    previousTokens: (string | undefined)[];
    skipToken?: string;
    query: ISimilarSuggestionsQuery;
  }): Promise<void> {
    if (!this._isMounted) {
      return;
    }

    const listId: string | undefined = this._currentListId;

    if (!listId) {
      return;
    }

    const normalizedTitle: string = (options.query.title ?? '').replace(/\s+/g, ' ').trim();
    const normalizedDescription: string = (options.query.description ?? '').replace(/\s+/g, ' ').trim();
    const effectiveQuery: ISimilarSuggestionsQuery = {
      title: normalizedTitle,
      description: normalizedDescription
    };

    if (!effectiveQuery.title && !effectiveQuery.description) {
      this._updateState({
        similarSuggestions: {
          items: [],
          page: 1,
          currentToken: undefined,
          nextToken: undefined,
          previousTokens: []
        },
        similarSuggestionsQuery: { ...EMPTY_SIMILAR_SUGGESTIONS_QUERY },
        isSimilarSuggestionsLoading: false
      });
      return;
    }

    this._updateState({
      similarSuggestionsQuery: { ...effectiveQuery },
      isSimilarSuggestionsLoading: true
    });

    try {
      const response = await this.props.graphService.getSuggestionItems(listId, {
        top: MAX_SIMILAR_SUGGESTIONS,
        skipToken: options.skipToken,
        titleSearchQuery: effectiveQuery.title || undefined,
        descriptionSearchQuery: effectiveQuery.description || undefined,
        orderBy: 'createdDateTime desc'
      });

      const suggestionIds: number[] = response.items
        .map((entry) => {
          const fields: IGraphSuggestionItemFields = entry.fields;
          const rawId: unknown = fields.id ?? (fields as { Id?: unknown }).Id;
          return this._parseNumericId(rawId);
        })
        .filter((value): value is number => typeof value === 'number');

      let votesBySuggestion: Map<number, IVoteEntry[]> = new Map();
      let commentCounts: Map<number, number> = new Map();
      let commentsBySuggestion: Map<number, ISuggestionComment[]> = new Map();

      if (suggestionIds.length > 0) {
        if (this._currentVotesListId) {
          const voteListId: string = this._getResolvedVotesListId();
          const voteItems: IGraphVoteItem[] = await this.props.graphService.getVoteItems(
            voteListId,
            { suggestionIds }
          );
          votesBySuggestion = this._groupVotesBySuggestion(voteItems);
        }

        if (this._currentCommentsListId) {
          const commentListId: string = this._getResolvedCommentsListId();
          const [counts, commentItems] = await Promise.all([
            this.props.graphService.getCommentCounts(commentListId, { suggestionIds }),
            this.props.graphService.getCommentItems(commentListId, { suggestionIds })
          ]);
          commentCounts = counts;
          commentsBySuggestion = this._groupCommentsBySuggestion(commentItems);
        }
      }

      const baseItems: ISuggestionItem[] = this._mapGraphItemsToSuggestions(
        response.items,
        votesBySuggestion,
        commentCounts
      );

      const enrichedItems: ISuggestionItem[] = this._currentCommentsListId
        ? baseItems.map((item) => {
            const loadedComments: ISuggestionComment[] = commentsBySuggestion.get(item.id) ?? [];
            const mappedCount: number = Math.max(
              loadedComments.length,
              commentCounts.get(item.id) ?? 0
            );

            return {
              ...item,
              comments: loadedComments,
              commentCount: mappedCount,
              areCommentsLoaded: true
            };
          })
        : baseItems;

      if (
        !this._areSimilarSuggestionQueriesEqual(this.state.similarSuggestionsQuery, effectiveQuery)
      ) {
        return;
      }

      const limited: ISuggestionItem[] = enrichedItems.slice(0, MAX_SIMILAR_SUGGESTIONS);

      const currentSelectedId: number | undefined = this.state.selectedSimilarSuggestion?.id;
      const shouldKeepSelection: boolean =
        typeof currentSelectedId === 'number' &&
        limited.some((entry) => entry.id === currentSelectedId);
      const nextExpandedCommentIds: number[] =
        !shouldKeepSelection && typeof currentSelectedId === 'number'
          ? this.state.expandedCommentIds.filter((id) => id !== currentSelectedId)
          : this.state.expandedCommentIds;
      const nextLoadingCommentIds: number[] =
        !shouldKeepSelection && typeof currentSelectedId === 'number'
          ? this.state.loadingCommentIds.filter((id) => id !== currentSelectedId)
          : this.state.loadingCommentIds;

      const nextSelectedSimilarSuggestion: ISuggestionItem | undefined = shouldKeepSelection
        ? limited.find((entry) => entry.id === currentSelectedId) ?? this.state.selectedSimilarSuggestion
        : undefined;

      const nextIsSelectedSimilarSuggestionLoading: boolean = shouldKeepSelection
        ? this.state.isSelectedSimilarSuggestionLoading
        : false;

      this._updateState({
        similarSuggestions: {
          items: limited,
          page: options.page,
          currentToken: options.skipToken,
          nextToken: response.nextToken,
          previousTokens: options.previousTokens
        },
        isSimilarSuggestionsLoading: false,
        selectedSimilarSuggestion: nextSelectedSimilarSuggestion,
        isSelectedSimilarSuggestionLoading: nextIsSelectedSimilarSuggestionLoading,
        expandedCommentIds: nextExpandedCommentIds,
        loadingCommentIds: nextLoadingCommentIds
      });
    } catch (error) {
      console.error('Failed to load similar suggestions.', error);

      if (
        !this._areSimilarSuggestionQueriesEqual(this.state.similarSuggestionsQuery, effectiveQuery)
      ) {
        return;
      }

      const staleSelectedId: number | undefined = this.state.selectedSimilarSuggestion?.id;
      const nextExpandedCommentIds: number[] =
        typeof staleSelectedId === 'number'
          ? this.state.expandedCommentIds.filter((id) => id !== staleSelectedId)
          : this.state.expandedCommentIds;
      const nextLoadingCommentIds: number[] =
        typeof staleSelectedId === 'number'
          ? this.state.loadingCommentIds.filter((id) => id !== staleSelectedId)
          : this.state.loadingCommentIds;

      this._updateState({
        similarSuggestions: {
          items: [],
          page: 1,
          currentToken: undefined,
          nextToken: undefined,
          previousTokens: []
        },
        isSimilarSuggestionsLoading: false,
        selectedSimilarSuggestion: undefined,
        isSelectedSimilarSuggestionLoading: false,
        expandedCommentIds: nextExpandedCommentIds,
        loadingCommentIds: nextLoadingCommentIds
      });
    }
  }

  private async _loadSelectedSimilarSuggestion(
    suggestionId: number,
    status: 'Active' | 'Done'
  ): Promise<void> {
    if (!this._isMounted) {
      return;
    }

    this._updateState({ isSelectedSimilarSuggestionLoading: true });

    try {
      const { items } = await this._getSuggestionsPage(status, undefined, {
        searchQuery: '',
        category: undefined,
        subcategory: undefined,
        suggestionId
      });

      if (!this._isMounted || this.state.selectedSimilarSuggestion?.id !== suggestionId) {
        return;
      }

      const nextSuggestion: ISuggestionItem | undefined = items.find((entry) => entry.id === suggestionId);

      if (!nextSuggestion) {
        const nextExpanded: number[] = this.state.expandedCommentIds.filter((id) => id !== suggestionId);
        const nextLoading: number[] = this.state.loadingCommentIds.filter((id) => id !== suggestionId);

        this._updateState({
          selectedSimilarSuggestion: undefined,
          isSelectedSimilarSuggestionLoading: false,
          expandedCommentIds: nextExpanded,
          loadingCommentIds: nextLoading
        });
        return;
      }

      let comments: ISuggestionComment[] = [];
      let areCommentsLoaded: boolean = nextSuggestion.commentCount === 0;
      let commentCount: number = nextSuggestion.commentCount;

      if (nextSuggestion.commentCount > 0) {
        const commentListId: string = this._getResolvedCommentsListId();
        const commentItems: IGraphCommentItem[] = await this.props.graphService.getCommentItems(
          commentListId,
          {
            suggestionIds: [suggestionId]
          }
        );
        const commentsBySuggestion: Map<number, ISuggestionComment[]> = this._groupCommentsBySuggestion(
          commentItems
        );
        comments = commentsBySuggestion.get(suggestionId) ?? [];
        commentCount = comments.length;
        areCommentsLoaded = true;
      }

      this._updateState(
        {
          selectedSimilarSuggestion: {
            ...nextSuggestion,
            comments,
            commentCount,
            areCommentsLoaded
          },
          isSelectedSimilarSuggestionLoading: false
        },
        () => {
          this._ensureCommentSectionExpanded(suggestionId);
        }
      );
    } catch (error) {
      if (!this._isMounted || this.state.selectedSimilarSuggestion?.id !== suggestionId) {
        return;
      }

      this._handleError('We could not load the selected suggestion. Please try again.', error);

      const nextExpanded: number[] = this.state.expandedCommentIds.filter((id) => id !== suggestionId);
      const nextLoading: number[] = this.state.loadingCommentIds.filter((id) => id !== suggestionId);

      this._updateState({
        selectedSimilarSuggestion: undefined,
        isSelectedSimilarSuggestionLoading: false,
        expandedCommentIds: nextExpanded,
        loadingCommentIds: nextLoading
      });
    }
  }

  private _onActiveSearchChange = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    const nextFilter: IFilterState = {
      ...this.state.activeFilter,
      searchQuery: newValue ?? '',
      suggestionId: undefined
    };
    this._applyActiveFilter(nextFilter);
  };

  private _onCategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key);
    const normalized: string = key.trim();
    const nextCategory: SuggestionCategory =
      this._findCategoryMatch(normalized, this.state.categories) ?? this._getDefaultCategory(this.state.categories);
    const nextSubcategoryKey: string | undefined = this._getValidSubcategoryKeyForCategory(
      nextCategory,
      this.state.newSubcategoryKey
    );

    this._updateState({ newCategory: nextCategory, newSubcategoryKey: nextSubcategoryKey });
  };

  private _onSubcategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key);
    const definition: ISubcategoryDefinition | undefined = this.state.subcategories.find(
      (item) => item.key === key
    );

    if (!definition) {
      return;
    }

    this._updateState({ newSubcategoryKey: definition.key });
  };

  private _onActiveFilterCategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key).trim();
    let nextCategory: SuggestionCategory | undefined;

    if (key !== ALL_CATEGORY_FILTER_KEY) {
      nextCategory =
        this._findCategoryMatch(key, this.state.categories) ?? (key.length > 0 ? key : undefined);
    }

    const nextFilter: IFilterState = {
      ...this.state.activeFilter,
      category: nextCategory,
      subcategory: this._normalizeFilterSubcategory(
        nextCategory,
        this.state.activeFilter.subcategory,
        this.state.subcategories
      ),
      suggestionId: undefined
    };

    this._applyActiveFilter(nextFilter);
  };

  private _onActiveFilterSubcategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key);
    const nextFilter: IFilterState =
      key === ALL_SUBCATEGORY_FILTER_KEY
        ? { ...this.state.activeFilter, subcategory: undefined, suggestionId: undefined }
        : { ...this.state.activeFilter, subcategory: key, suggestionId: undefined };

    this._applyActiveFilter(nextFilter);
  };

  private _onCompletedSearchChange = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    const nextFilter: IFilterState = {
      ...this.state.completedFilter,
      searchQuery: newValue ?? '',
      suggestionId: undefined
    };
    this._applyCompletedFilter(nextFilter);
  };

  private _onCompletedFilterCategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key).trim();
    let nextCategory: SuggestionCategory | undefined;

    if (key !== ALL_CATEGORY_FILTER_KEY) {
      nextCategory =
        this._findCategoryMatch(key, this.state.categories) ?? (key.length > 0 ? key : undefined);
    }

    const nextFilter: IFilterState = {
      ...this.state.completedFilter,
      category: nextCategory,
      subcategory: this._normalizeFilterSubcategory(
        nextCategory,
        this.state.completedFilter.subcategory,
        this.state.subcategories
      ),
      suggestionId: undefined
    };

    this._applyCompletedFilter(nextFilter);
  };

  private _onCompletedFilterSubcategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key);
    const nextFilter: IFilterState =
      key === ALL_SUBCATEGORY_FILTER_KEY
        ? { ...this.state.completedFilter, subcategory: undefined, suggestionId: undefined }
        : { ...this.state.completedFilter, subcategory: key, suggestionId: undefined };

    this._applyCompletedFilter(nextFilter);
  };

  private _dismissError = (): void => {
    this._updateState({ error: undefined });
  };

  private _dismissSuccess = (): void => {
    this._updateState({ success: undefined });
  };

  private _applyActiveFilter(nextFilter: IFilterState): void {
    this._updateState({ isActiveSuggestionsLoading: true, error: undefined, success: undefined });

    this._fetchActiveSuggestions({
      page: 1,
      previousTokens: [],
      skipToken: undefined,
      filter: nextFilter
    })
      .then(() => {
        this._updateState({ isActiveSuggestionsLoading: false });
      })
      .catch((error) => {
        this._handleError('We could not load the active suggestions. Please try again.', error);
        this._updateState({ isActiveSuggestionsLoading: false });
      });
  }

  private _applyCompletedFilter(nextFilter: IFilterState): void {
    this._updateState({ isCompletedSuggestionsLoading: true, error: undefined, success: undefined });

    this._fetchCompletedSuggestions({
      page: 1,
      previousTokens: [],
      skipToken: undefined,
      filter: nextFilter
    })
      .then(() => {
        this._updateState({ isCompletedSuggestionsLoading: false });
      })
      .catch((error) => {
        this._handleError('We could not load the completed suggestions. Please try again.', error);
        this._updateState({ isCompletedSuggestionsLoading: false });
      });
  }

  private _addSuggestion = async (): Promise<void> => {
    const title: string = this.state.newTitle.trim();
    const description: string = this.state.newDescription.trim();
    const category: SuggestionCategory = this.state.newCategory;
    const selectedSubcategory: ISubcategoryDefinition | undefined = this._getSelectedSubcategoryDefinition();

    if (!title) {
      this._handleError('Please add a title before submitting your suggestion.');
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const listId: string = this._getResolvedListId();
      const payload: IGraphSuggestionItemFields = {
        Title: title,
        Details: description,
        Status: 'Active',
        Category: category
      };

      if (selectedSubcategory) {
        payload.Subcategory = selectedSubcategory.title;
      }

      await this.props.graphService.addSuggestion(listId, payload);

      const defaultCategory: SuggestionCategory = this._getDefaultCategory(this.state.categories);
      this._debouncedSimilarSuggestionsSearch.cancel();

      this._updateState({
        newTitle: '',
        newDescription: '',
        newCategory: defaultCategory,
        newSubcategoryKey: this._getValidSubcategoryKeyForCategory(
          defaultCategory,
          undefined
        ),
        similarSuggestions: {
          items: [],
          page: 1,
          currentToken: undefined,
          nextToken: undefined,
          previousTokens: []
        },
        isSimilarSuggestionsLoading: false,
        similarSuggestionsQuery: { ...EMPTY_SIMILAR_SUGGESTIONS_QUERY }
      });

      await this._loadSuggestions();

      this._updateState({ success: 'Your suggestion has been added.' });
    } catch (error) {
      this._handleError('We could not add your suggestion. Please try again.', error);
    } finally {
      this._updateState({ isLoading: false });
    }
  };

  private async _toggleVote(item: ISuggestionItem): Promise<void> {
    const normalizedUser: string | undefined = this._normalizeLoginName(this.props.userLoginName);

    if (!normalizedUser) {
      this._handleError('We could not determine the current user. Please try again later.');
      return;
    }

    const currentVote: IVoteEntry | undefined = item.voteEntries.find((vote) => vote.username === normalizedUser);
    const hasVoted: boolean = !!currentVote && currentVote.votes > 0;

    if (!hasVoted && this.state.availableVotes <= 0) {
      this._handleError('You have used all of your votes. Mark a suggestion as done or remove one of your votes to continue.');
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const voteListId: string = this._getResolvedVotesListId();

      if (hasVoted && currentVote) {
        await this.props.graphService.deleteVote(voteListId, currentVote.id);
      } else {
        await this.props.graphService.addVote(voteListId, {
          SuggestionId: item.id,
          Username: normalizedUser,
          Votes: 1
        });
      }

      await Promise.all([this._refreshActiveSuggestions(), this._loadAvailableVotes()]);

      if (this.state.selectedSimilarSuggestion?.id === item.id) {
        await this._loadSelectedSimilarSuggestion(item.id, item.status);
      }

      this._updateState({ success: hasVoted ? 'Your vote has been removed.' : 'Thanks for voting!' });
    } catch (error) {
      this._handleError('We could not update your vote. Please try again.', error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private _canCurrentUserDeleteSuggestion(item: ISuggestionItem): boolean {
    if (this.props.isCurrentUserAdmin) {
      return true;
    }

    return this._isCurrentUserSuggestionOwner(item);
  }

  private _isCurrentUserSuggestionOwner(item: ISuggestionItem): boolean {
    const ownerLoginName: string | undefined = item.createdByLoginName;
    const currentUserLoginName: string | undefined = this._normalizeLoginName(this.props.userLoginName);

    return !!ownerLoginName && !!currentUserLoginName && ownerLoginName === currentUserLoginName;
  }

  private _normalizeLoginName(value?: string): string | undefined {
    if (typeof value !== 'string') {
      return undefined;
    }

    const trimmed: string = value.trim();
    return trimmed.length > 0 ? trimmed.toLowerCase() : undefined;
  }

  private _isCommentSectionExpanded(suggestionId: number): boolean {
    return this.state.expandedCommentIds.indexOf(suggestionId) !== -1;
  }

  private _toggleCommentsSection = (suggestionId: number): void => {
    if (!this._isMounted) {
      return;
    }

    this.setState(
      (prevState) => {
        const isExpanded: boolean = prevState.expandedCommentIds.indexOf(suggestionId) !== -1;
        const nextExpanded: number[] = isExpanded
          ? prevState.expandedCommentIds.filter((id) => id !== suggestionId)
          : [...prevState.expandedCommentIds, suggestionId];

        return { expandedCommentIds: nextExpanded };
      },
      () => {
        if (this._isCommentSectionExpanded(suggestionId)) {
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          this._ensureCommentsLoaded(suggestionId);
        }
      }
    );
  };

  private _ensureCommentSectionExpanded(suggestionId: number): void {
    if (!this._isMounted) {
      return;
    }

    this.setState(
      (prevState) => {
        if (prevState.expandedCommentIds.indexOf(suggestionId) !== -1) {
          return null;
        }

        return {
          expandedCommentIds: [...prevState.expandedCommentIds, suggestionId]
        };
      },
      () => {
        if (this._isCommentSectionExpanded(suggestionId)) {
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          this._ensureCommentsLoaded(suggestionId);
        }
      }
    );
  }

  private async _ensureCommentsLoaded(suggestionId: number): Promise<void> {
    if (!this._isMounted) {
      return;
    }

    const suggestion: ISuggestionItem | undefined = this._findSuggestionById(suggestionId);

    if (!suggestion || suggestion.areCommentsLoaded) {
      return;
    }

    if (this.state.loadingCommentIds.indexOf(suggestionId) !== -1) {
      return;
    }

    this.setState((prevState) => ({
      loadingCommentIds: [...prevState.loadingCommentIds, suggestionId]
    }));

    try {
      const commentListId: string = this._getResolvedCommentsListId();
      const commentItems: IGraphCommentItem[] = await this.props.graphService.getCommentItems(commentListId, {
        suggestionIds: [suggestionId]
      });
      const commentsBySuggestion: Map<number, ISuggestionComment[]> = this._groupCommentsBySuggestion(
        commentItems
      );
      const comments: ISuggestionComment[] = commentsBySuggestion.get(suggestionId) ?? [];

      this.setState((prevState) => ({
        loadingCommentIds: prevState.loadingCommentIds.filter((id) => id !== suggestionId),
        activeSuggestions: this._updateSuggestionItem(prevState.activeSuggestions, suggestionId, {
          comments,
          commentCount: comments.length,
          areCommentsLoaded: true
        }),
        completedSuggestions: this._updateSuggestionItem(prevState.completedSuggestions, suggestionId, {
          comments,
          commentCount: comments.length,
          areCommentsLoaded: true
        }),
        selectedSimilarSuggestion:
          prevState.selectedSimilarSuggestion && prevState.selectedSimilarSuggestion.id === suggestionId
            ? {
                ...prevState.selectedSimilarSuggestion,
                comments,
                commentCount: comments.length,
                areCommentsLoaded: true
              }
            : prevState.selectedSimilarSuggestion
      }));
    } catch (error) {
      this._handleError('We could not load the comments. Please try again.', error);
      this.setState((prevState) => ({
        loadingCommentIds: prevState.loadingCommentIds.filter((id) => id !== suggestionId)
      }));
    }
  }

  private _findSuggestionById(suggestionId: number): ISuggestionItem | undefined {
    const { activeSuggestions, completedSuggestions, selectedSimilarSuggestion } = this.state;
    return (
      activeSuggestions.items.find((item) => item.id === suggestionId) ??
      completedSuggestions.items.find((item) => item.id === suggestionId) ??
      (selectedSimilarSuggestion && selectedSimilarSuggestion.id === suggestionId
        ? selectedSimilarSuggestion
        : undefined)
    );
  }

  private _updateSuggestionItem(
    source: IPaginatedSuggestionsState,
    suggestionId: number,
    updates: Partial<ISuggestionItem>
  ): IPaginatedSuggestionsState {
    const items: ISuggestionItem[] = source.items.map((item) =>
      item.id === suggestionId ? { ...item, ...updates } : item
    );

    return { ...source, items };
  }

  private async _addCommentToSuggestion(item: ISuggestionItem): Promise<void> {
    const commentInput: string | null = window.prompt('Add a comment for this suggestion.', '');

    if (commentInput === null) {
      return;
    }

    const commentText: string = commentInput.trim();

    if (commentText.length === 0) {
      this._handleError('Please enter a comment before submitting.');
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const commentListId: string = this._getResolvedCommentsListId();
      const title: string = `Suggestion #${item.id}`;

      await this.props.graphService.addCommentItem(commentListId, {
        Title: title.length > 255 ? title.slice(0, 255) : title,
        SuggestionId: item.id,
        Comment: commentText
      });

      await this._refreshActiveSuggestions();
      this._ensureCommentSectionExpanded(item.id);

      if (this.state.selectedSimilarSuggestion?.id === item.id) {
        await this._loadSelectedSimilarSuggestion(item.id, item.status);
        this._ensureCommentSectionExpanded(item.id);
      }

      this._updateState({ success: 'Your comment has been added.' });
    } catch (error) {
      this._handleError('We could not add your comment. Please try again.', error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private async _deleteCommentFromSuggestion(
    item: ISuggestionItem,
    comment: ISuggestionComment
  ): Promise<void> {
    if (!this.props.isCurrentUserAdmin) {
      this._handleError('Only administrators can delete comments.');
      return;
    }

    const confirmed: boolean = window.confirm('Are you sure you want to delete this comment?');

    if (!confirmed) {
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const commentListId: string = this._getResolvedCommentsListId();
      await this.props.graphService.deleteCommentItem(commentListId, comment.id);

      if (item.status === 'Done') {
        await this._refreshCompletedSuggestions();
      } else {
        await this._refreshActiveSuggestions();
      }

      this._ensureCommentSectionExpanded(item.id);

      if (this.state.selectedSimilarSuggestion?.id === item.id) {
        await this._loadSelectedSimilarSuggestion(item.id, item.status);
        this._ensureCommentSectionExpanded(item.id);
      }
      this._updateState({ success: 'The comment has been removed.' });
    } catch (error) {
      this._handleError('We could not remove the comment. Please try again.', error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private async _markSuggestionAsDone(item: ISuggestionItem): Promise<void> {
    if (!this.props.isCurrentUserAdmin) {
      this._handleError('Only administrators can mark suggestions as done.');
      return;
    }

    const commentInput: string | null = window.prompt(
      'Add a comment for this suggestion (optional). Leave blank to skip.',
      ''
    );

    if (commentInput === null) {
      return;
    }

    const commentText: string = commentInput.trim();

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const listId: string = this._getResolvedListId();
      const voteListId: string = this._getResolvedVotesListId();
      const commentListId: string = this._getResolvedCommentsListId();

      await this.props.graphService.updateSuggestion(listId, item.id, {
        Status: 'Done',
        Votes: item.votes,
        CompletedDateTime: new Date().toISOString()
      });

      await this.props.graphService.deleteVotesForSuggestion(voteListId, item.id);

      if (commentText.length > 0) {
        const title: string = `Suggestion #${item.id}`;
        await this.props.graphService.addCommentItem(commentListId, {
          Title: title.length > 255 ? title.slice(0, 255) : title,
          SuggestionId: item.id,
          Comment: commentText
        });
      }

      await this._loadSuggestions();

      if (this.state.selectedSimilarSuggestion?.id === item.id) {
        await this._loadSelectedSimilarSuggestion(item.id, 'Done');
      }

      this._updateState({ success: 'The suggestion has been marked as done.' });
    } catch (error) {
      this._handleError('We could not mark this suggestion as done. Please try again.', error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private async _deleteSuggestion(item: ISuggestionItem): Promise<void> {
    if (!this._canCurrentUserDeleteSuggestion(item)) {
      this._handleError('You do not have permission to remove this suggestion.');
      return;
    }

    const confirmation: boolean = window.confirm('Are you sure you want to remove this suggestion? This action cannot be undone.');

    if (!confirmation) {
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const listId: string = this._getResolvedListId();
      const voteListId: string = this._getResolvedVotesListId();
      const commentListId: string = this._getResolvedCommentsListId();

      await this.props.graphService.deleteSuggestion(listId, item.id);
      await Promise.all([
        this.props.graphService.deleteVotesForSuggestion(voteListId, item.id),
        this.props.graphService.deleteCommentsForSuggestion(commentListId, item.id)
      ]);

      if (item.status === 'Done') {
        await this._refreshCompletedSuggestions();
      } else {
        await Promise.all([this._refreshActiveSuggestions(), this._loadAvailableVotes()]);
      }

      if (this.state.selectedSimilarSuggestion?.id === item.id && this._isMounted) {
        this.setState((prevState) => ({
          selectedSimilarSuggestion: undefined,
          isSelectedSimilarSuggestionLoading: false,
          expandedCommentIds: prevState.expandedCommentIds.filter((id) => id !== item.id),
          loadingCommentIds: prevState.loadingCommentIds.filter((id) => id !== item.id)
        }));
      }

      this._updateState({ success: 'The suggestion has been removed.' });
    } catch (error) {
      this._handleError('We could not remove this suggestion. Please try again.', error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private _normalizeListTitle(value?: string): string {
    const trimmed: string = (value ?? '').trim();
    return trimmed.length > 0 ? trimmed : DEFAULT_SUGGESTIONS_LIST_TITLE;
  }

  private get _listTitle(): string {
    return this._normalizeListTitle(this.props.listTitle);
  }

  private _normalizeVoteListTitle(value?: string, listTitle?: string): string {
    const trimmed: string = (value ?? '').trim();
    const normalizedListTitle: string = this._normalizeListTitle(listTitle ?? this.props.listTitle);
    return trimmed.length > 0 ? trimmed : `${normalizedListTitle}Votes`;
  }

  private get _voteListTitle(): string {
    return this._normalizeVoteListTitle(this.props.voteListTitle, this.props.listTitle);
  }

  private _normalizeCommentListTitle(value?: string, listTitle?: string): string {
    const trimmed: string = (value ?? '').trim();
    const normalizedListTitle: string = this._normalizeListTitle(listTitle ?? this.props.listTitle);
    return trimmed.length > 0 ? trimmed : `${normalizedListTitle}Comments`;
  }

  private get _commentListTitle(): string {
    return this._normalizeCommentListTitle(this.props.commentListTitle, this.props.listTitle);
  }

  private _normalizeOptionalListTitle(value?: string): string | undefined {
    if (typeof value !== 'string') {
      return undefined;
    }

    const trimmed: string = value.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  }

  private get _subcategoryListTitle(): string | undefined {
    return this._normalizeOptionalListTitle(this.props.subcategoryListTitle);
  }

  private get _categoryListTitle(): string | undefined {
    return this._normalizeOptionalListTitle(this.props.categoryListTitle);
  }

  private _parseVotes(value: unknown): number {
    if (typeof value === 'number' && Number.isFinite(value)) {
      return value;
    }

    if (typeof value === 'string') {
      const parsed: number = parseInt(value, 10);
      if (Number.isFinite(parsed)) {
        return parsed;
      }
    }

    return 0;
  }

  private _tryNormalizeCategory(value: unknown): SuggestionCategory | undefined {
    if (typeof value !== 'string') {
      return undefined;
    }

    const normalized: string = value.trim();

    if (!normalized) {
      return undefined;
    }

    return this._findCategoryMatch(normalized, this.state.categories) ?? normalized;
  }

  private _normalizeCategory(value: unknown): SuggestionCategory {
    return this._tryNormalizeCategory(value) ?? this._getDefaultCategory(this.state.categories);
  }

  private _getResolvedListId(): string {
    if (!this._currentListId) {
      throw new Error('The suggestions list has not been initialized yet.');
    }

    return this._currentListId;
  }

  private _getResolvedVotesListId(): string {
    if (!this._currentVotesListId) {
      throw new Error('The votes list has not been initialized yet.');
    }

    return this._currentVotesListId;
  }

  private _getResolvedCommentsListId(): string {
    if (!this._currentCommentsListId) {
      throw new Error('The comments list has not been initialized yet.');
    }

    return this._currentCommentsListId;
  }

  private _getResolvedCategoryListId(): string {
    if (!this._currentCategoryListId) {
      throw new Error('The category list has not been initialized yet.');
    }

    return this._currentCategoryListId;
  }

  private _getResolvedSubcategoryListId(): string {
    if (!this._currentSubcategoryListId) {
      throw new Error('The subcategory list has not been initialized yet.');
    }

    return this._currentSubcategoryListId;
  }

  private _handleError(message: string, error?: unknown): void {
    console.error(message, error);
    this._updateState({ error: message, success: undefined });
  }

  private _toggleAddSuggestionSection = (): void => {
    if (!this._isMounted) {
      return;
    }

    this.setState((prevState) => ({
      isAddSuggestionExpanded: !prevState.isAddSuggestionExpanded
    }));
  };

  private _toggleActiveSection = (): void => {
    if (!this._isMounted) {
      return;
    }

    this.setState((prevState) => ({
      isActiveSuggestionsExpanded: !prevState.isActiveSuggestionsExpanded
    }));
  };

  private _toggleCompletedSection = (): void => {
    if (!this._isMounted) {
      return;
    }

    this.setState((prevState) => ({
      isCompletedSuggestionsExpanded: !prevState.isCompletedSuggestionsExpanded
    }));
  };

  private _flushPendingSimilarSuggestionsSearch(): void {
    if (!this._pendingSimilarSuggestionsQuery || !this._currentListId) {
      return;
    }

    const pendingQuery: ISimilarSuggestionsQuery = this._pendingSimilarSuggestionsQuery;
    this._pendingSimilarSuggestionsQuery = undefined;
    this._debouncedSimilarSuggestionsSearch(pendingQuery);
  }

  private _updateState(
    state: Partial<ISamverkansportalenState>,
    callback?: () => void
  ): void {
    if (!this._isMounted) {
      return;
    }

    this.setState(
      state as Pick<ISamverkansportalenState, keyof ISamverkansportalenState>,
      callback
    );
  }

  private _parseNumericId(value: unknown): number | undefined {
    if (typeof value === 'number' && Number.isFinite(value)) {
      return value;
    }

    if (typeof value === 'string') {
      const parsed: number = parseInt(value, 10);
      if (Number.isFinite(parsed)) {
        return parsed;
      }
    }

    return undefined;
  }
}
