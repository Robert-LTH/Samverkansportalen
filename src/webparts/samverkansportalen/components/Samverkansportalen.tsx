/* eslint-disable max-lines */
import * as React from 'react';
import {
  PrimaryButton,
  DefaultButton,
  IconButton,
  ActionButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  TextField,
  Dropdown,
  type IDropdownOption
} from '@fluentui/react';
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
  comments: ISuggestionComment[];
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
  error?: string;
  success?: string;
  isAddSuggestionExpanded: boolean;
  isActiveSuggestionsExpanded: boolean;
  isCompletedSuggestionsExpanded: boolean;
}

interface IFilterState {
  searchQuery: string;
  category?: SuggestionCategory;
  subcategory?: string;
}

interface IPaginatedSuggestionsState {
  items: ISuggestionItem[];
  page: number;
  currentToken?: string;
  nextToken?: string;
  previousTokens: (string | undefined)[];
}

const MAX_VOTES_PER_USER: number = 5;
const FALLBACK_CATEGORIES: SuggestionCategory[] = ['Change request', 'Webbinar', 'Article'];
const DEFAULT_SUGGESTION_CATEGORY: SuggestionCategory = FALLBACK_CATEGORIES[0];
const ALL_CATEGORY_FILTER_KEY: string = '__all_categories__';
const ALL_SUBCATEGORY_FILTER_KEY: string = '__all_subcategories__';
const SUGGESTIONS_PAGE_SIZE: number = 5;

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
      activeFilter: { searchQuery: '', category: undefined, subcategory: undefined },
      completedFilter: { searchQuery: '', category: undefined, subcategory: undefined },
      isAddSuggestionExpanded: true,
      isActiveSuggestionsExpanded: true,
      isCompletedSuggestionsExpanded: true
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this._initialize();
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
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
      isLoading,
      isActiveSuggestionsLoading,
      isCompletedSuggestionsLoading,
      availableVotes,
      newTitle,
      newDescription,
      newCategory,
      newSubcategoryKey,
      subcategories,
      categories,
      activeFilter,
      completedFilter,
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

    return (
      <section className={`${styles.samverkansportalen} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        <header className={styles.header}>
          <div>
            <h2 className={styles.title}>{this.props.headerTitle}</h2>
            <p className={styles.subtitle}>{this.props.headerSubtitle}</p>
          </div>
          <div className={styles.voteSummary} aria-live="polite">
            <span className={styles.voteLabel}>Votes remaining</span>
            <span className={styles.voteValue}>{availableVotes} / {MAX_VOTES_PER_USER}</span>
          </div>
        </header>

        {error && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
            onDismiss={this._dismissError}
          >
            {error}
          </MessageBar>
        )}

        {success && (
          <MessageBar
            messageBarType={MessageBarType.success}
            isMultiline={false}
            onDismiss={this._dismissSuccess}
          >
            {success}
          </MessageBar>
        )}

        <div className={styles.addSuggestion}>
          {this._renderSectionHeader(
            'Add a suggestion',
            this._sectionIds.add.title,
            this._sectionIds.add.content,
            isAddSuggestionExpanded,
            this._toggleAddSuggestionSection
          )}
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
                  label="Title"
                  required
                  value={newTitle}
                  onChange={this._onTitleChange}
                  disabled={isLoading}
                />
                <TextField
                  label="Details"
                  multiline
                  rows={3}
                  value={newDescription}
                  onChange={this._onDescriptionChange}
                  disabled={isLoading}
                />
                <Dropdown
                  label="Category"
                  options={categoryOptions}
                  selectedKey={newCategory}
                  onChange={this._onCategoryChange}
                  disabled={isLoading || categoryOptions.length === 0}
                />
                <Dropdown
                  label="Subcategory"
                  options={subcategoryOptions}
                  selectedKey={newSubcategoryKey}
                  onChange={this._onSubcategoryChange}
                  disabled={isLoading || subcategoryOptions.length === 0}
                  placeholder={
                    subcategoryOptions.length === 0
                      ? 'No subcategories available'
                      : 'Select a subcategory'
                  }
                />
                <PrimaryButton
                  text="Submit suggestion"
                  onClick={this._addSuggestion}
                  disabled={isLoading || newTitle.trim().length === 0}
                />
              </div>
            )}
          </div>
        </div>

        <div className={styles.suggestionSection}>
          {this._renderSectionHeader(
            'Active suggestions',
            this._sectionIds.active.title,
            this._sectionIds.active.content,
            isActiveSuggestionsExpanded,
            this._toggleActiveSection
          )}
          <div
            id={this._sectionIds.active.content}
            role="region"
            aria-labelledby={this._sectionIds.active.title}
            className={`${styles.sectionContent} ${
              isActiveSuggestionsExpanded ? '' : styles.sectionContentCollapsed
            }`}
            hidden={!isActiveSuggestionsExpanded}
          >
            {isActiveSuggestionsExpanded && (
              <>
                <div className={styles.filterControls}>
                  <TextField
                    label="Search"
                    value={activeFilter.searchQuery}
                    onChange={this._onActiveSearchChange}
                    disabled={isLoading || isActiveSuggestionsLoading}
                    placeholder="Search by title or details"
                    className={styles.filterSearch}
                  />
                  <Dropdown
                    label="Category"
                    options={filterCategoryOptions}
                    selectedKey={activeFilter.category ?? ALL_CATEGORY_FILTER_KEY}
                    onChange={this._onActiveFilterCategoryChange}
                    disabled={
                      isLoading ||
                      isActiveSuggestionsLoading ||
                      filterCategoryOptions.length <= 1
                    }
                    className={styles.filterDropdown}
                  />
                  <Dropdown
                    label="Subcategory"
                    options={activeFilterSubcategoryOptions}
                    selectedKey={activeFilter.subcategory ?? ALL_SUBCATEGORY_FILTER_KEY}
                    onChange={this._onActiveFilterSubcategoryChange}
                    disabled={
                      isLoading ||
                      isActiveSuggestionsLoading ||
                      activeFilterSubcategoryOptions.length <= 1
                    }
                    className={styles.filterDropdown}
                  />
                </div>
                {isLoading || isActiveSuggestionsLoading ? (
                  <Spinner label="Loading suggestions..." size={SpinnerSize.large} />
                ) : (
                  <>
                    {this._renderSuggestionList(activeSuggestions.items, false)}
                    {this._renderPaginationControls(
                      activeSuggestions.page,
                      activeSuggestions.previousTokens.length > 0,
                      !!activeSuggestions.nextToken,
                      this._goToPreviousActivePage,
                      this._goToNextActivePage
                    )}
                  </>
                )}
              </>
            )}
          </div>
        </div>

        <div className={styles.suggestionSection}>
          {this._renderSectionHeader(
            'Completed suggestions',
            this._sectionIds.completed.title,
            this._sectionIds.completed.content,
            isCompletedSuggestionsExpanded,
            this._toggleCompletedSection
          )}
          <div
            id={this._sectionIds.completed.content}
            role="region"
            aria-labelledby={this._sectionIds.completed.title}
            className={`${styles.sectionContent} ${
              isCompletedSuggestionsExpanded ? '' : styles.sectionContentCollapsed
            }`}
            hidden={!isCompletedSuggestionsExpanded}
          >
            {isCompletedSuggestionsExpanded && (
              <>
                <div className={styles.filterControls}>
                  <TextField
                    label="Search"
                    value={completedFilter.searchQuery}
                    onChange={this._onCompletedSearchChange}
                    disabled={isLoading || isCompletedSuggestionsLoading}
                    placeholder="Search by title or details"
                    className={styles.filterSearch}
                  />
                  <Dropdown
                    label="Category"
                    options={filterCategoryOptions}
                    selectedKey={completedFilter.category ?? ALL_CATEGORY_FILTER_KEY}
                    onChange={this._onCompletedFilterCategoryChange}
                    disabled={
                      isLoading ||
                      isCompletedSuggestionsLoading ||
                      filterCategoryOptions.length <= 1
                    }
                    className={styles.filterDropdown}
                  />
                  <Dropdown
                    label="Subcategory"
                    options={completedFilterSubcategoryOptions}
                    selectedKey={completedFilter.subcategory ?? ALL_SUBCATEGORY_FILTER_KEY}
                    onChange={this._onCompletedFilterSubcategoryChange}
                    disabled={
                      isLoading ||
                      isCompletedSuggestionsLoading ||
                      completedFilterSubcategoryOptions.length <= 1
                    }
                    className={styles.filterDropdown}
                  />
                </div>
                {isLoading || isCompletedSuggestionsLoading ? (
                  <Spinner label="Loading suggestions..." size={SpinnerSize.large} />
                ) : (
                  <>
                    {this._renderSuggestionList(completedSuggestions.items, true)}
                    {this._renderPaginationControls(
                      completedSuggestions.page,
                      completedSuggestions.previousTokens.length > 0,
                      !!completedSuggestions.nextToken,
                      this._goToPreviousCompletedPage,
                      this._goToNextCompletedPage
                    )}
                  </>
                )}
              </>
            )}
          </div>
        </div>
      </section>
    );
  }

  private _renderSectionHeader(
    title: string,
    titleId: string,
    contentId: string,
    isExpanded: boolean,
    onToggle: () => void
  ): React.ReactNode {
    return (
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
          {isExpanded ? 'Hide' : 'Show'}
        </ActionButton>
      </div>
    );
  }

  private _renderSuggestionList(items: ISuggestionItem[], readOnly: boolean): React.ReactNode {
    if (items.length === 0) {
      return <p className={styles.emptyState}>There are no suggestions in this section yet.</p>;
    }

    const noVotesRemaining: boolean = this.state.availableVotes <= 0;
    const normalizedUser: string | undefined = this._normalizeLoginName(this.props.userLoginName);

    return this.props.useTableLayout
      ? this._renderSuggestionTable(items, readOnly, normalizedUser, noVotesRemaining)
      : this._renderSuggestionCards(items, readOnly, normalizedUser, noVotesRemaining);
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

  private _renderSuggestionCards(
    items: ISuggestionItem[],
    readOnly: boolean,
    normalizedUser: string | undefined,
    noVotesRemaining: boolean
  ): React.ReactNode {
    return (
      <ul className={styles.suggestionList}>
        {items.map((item) => {
          const {
            hasVoted,
            disableVote,
            canAddComment,
            canMarkSuggestionAsDone,
            canDeleteSuggestion
          } = this._getInteractionState(item, readOnly, normalizedUser, noVotesRemaining);

          return (
            <li key={item.id} className={styles.suggestionCard}>
              <div className={styles.cardHeader}>
                <div className={styles.cardText}>
                  <div className={styles.cardMeta}>
                    <span className={styles.entryId} aria-label={`Entry number ${item.id}`}>
                      #{item.id}
                    </span>
                    <span className={styles.categoryBadge}>{item.category}</span>
                    {item.subcategory && (
                      <span className={styles.subcategoryBadge}>{item.subcategory}</span>
                    )}
                  </div>
                  <h4 className={styles.suggestionTitle}>{item.title}</h4>
                  {item.description && (
                    <p className={styles.suggestionDescription}>{item.description}</p>
                  )}
                  {this._renderSuggestionTimestamps(item)}
                </div>
              <div className={styles.voteBadge} aria-label={`${item.votes} ${item.votes === 1 ? 'vote' : 'votes'}`}>
                <span className={styles.voteNumber}>{item.votes}</span>
                <span className={styles.voteText}>{item.votes === 1 ? 'vote' : 'votes'}</span>
              </div>
            </div>
              {this._renderComments(item.comments)}
              {this._renderActionButtons(
                item,
                readOnly,
                hasVoted,
                disableVote,
                canAddComment,
                canMarkSuggestionAsDone,
                canDeleteSuggestion,
                styles.cardActions
              )}
            </li>
          );
        })}
      </ul>
    );
  }

  private _renderSuggestionTable(
    items: ISuggestionItem[],
    readOnly: boolean,
    normalizedUser: string | undefined,
    noVotesRemaining: boolean
  ): React.ReactNode {
    return (
      <div className={styles.tableWrapper}>
        <table className={styles.suggestionTable}>
          <thead>
            <tr>
              <th scope="col" className={styles.tableHeaderId}>#</th>
              <th scope="col" className={styles.tableHeaderSuggestion}>Suggestion</th>
              <th scope="col" className={styles.tableHeaderCategory}>Category</th>
              <th scope="col" className={styles.tableHeaderSubcategory}>Subcategory</th>
              <th scope="col" className={styles.tableHeaderVotes}>Votes</th>
              <th scope="col" className={styles.tableHeaderActions}>Actions</th>
            </tr>
          </thead>
        <tbody>
          {items.map((item) => {
            const {
              hasVoted,
              disableVote,
              canAddComment,
              canMarkSuggestionAsDone,
              canDeleteSuggestion
            } = this._getInteractionState(item, readOnly, normalizedUser, noVotesRemaining);

            return (
              <tr key={item.id}>
                <td className={styles.tableCellId} data-label="Entry">
                  <span className={styles.entryId} aria-label={`Entry number ${item.id}`}>
                    #{item.id}
                  </span>
                </td>
                <td className={styles.tableCellSuggestion} data-label="Suggestion">
                  <h4 className={styles.suggestionTitle}>{item.title}</h4>
                  {item.description && (
                    <p className={styles.suggestionDescription}>{item.description}</p>
                  )}
                  {this._renderSuggestionTimestamps(item)}
                  {this._renderComments(item.comments)}
                </td>
                <td className={styles.tableCellCategory} data-label="Category">
                  <span className={styles.categoryBadge}>{item.category}</span>
                </td>
                <td className={styles.tableCellSubcategory} data-label="Subcategory">
                  {item.subcategory ? (
                    <span className={styles.subcategoryBadge}>{item.subcategory}</span>
                  ) : (
                    <span className={styles.subcategoryPlaceholder}>—</span>
                  )}
                </td>
                <td className={styles.tableCellVotes} data-label="Votes">
                  <div className={styles.voteBadge} aria-label={`${item.votes} ${item.votes === 1 ? 'vote' : 'votes'}`}>
                    <span className={styles.voteNumber}>{item.votes}</span>
                    <span className={styles.voteText}>{item.votes === 1 ? 'vote' : 'votes'}</span>
                  </div>
                </td>
                <td className={styles.tableCellActions} data-label="Actions">
                  {this._renderActionButtons(
                    item,
                    readOnly,
                    hasVoted,
                    disableVote,
                    canAddComment,
                    canMarkSuggestionAsDone,
                    canDeleteSuggestion,
                    styles.tableActions
                  )}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
    );
  }

  private _renderActionButtons(
    item: ISuggestionItem,
    readOnly: boolean,
    hasVoted: boolean,
    disableVote: boolean,
    canAddComment: boolean,
    canMarkSuggestionAsDone: boolean,
    canDeleteSuggestion: boolean,
    containerClassName: string
  ): React.ReactNode {
    return (
      <div className={containerClassName}>
        {readOnly ? (
          <DefaultButton text="Votes closed" disabled />
        ) : (
          <PrimaryButton
            text={hasVoted ? 'Remove vote' : 'Vote'}
            onClick={() => this._toggleVote(item)}
            disabled={disableVote}
          />
        )}
        {canAddComment && (
          <DefaultButton
            text="Add comment"
            onClick={() => this._addCommentToSuggestion(item)}
            disabled={this.state.isLoading}
          />
        )}
        {canMarkSuggestionAsDone && (
          <DefaultButton
            text="Mark as done"
            onClick={() => this._markSuggestionAsDone(item)}
            disabled={this.state.isLoading}
          />
        )}
        {canDeleteSuggestion && (
          <IconButton
            iconProps={{ iconName: 'Delete' }}
            title="Remove suggestion"
            ariaLabel="Remove suggestion"
            onClick={() => this._deleteSuggestion(item)}
            disabled={this.state.isLoading}
          />
        )}
      </div>
    );
  }

  private _renderSuggestionTimestamps(item: ISuggestionItem): React.ReactNode {
    const entries: { label: string; value: string }[] = [];

    if (item.createdDateTime) {
      entries.push({ label: 'Created', value: item.createdDateTime });
    }

    if (item.lastModifiedDateTime) {
      entries.push({ label: 'Last modified', value: item.lastModifiedDateTime });
    }

    if (item.completedDateTime) {
      entries.push({ label: 'Completed', value: item.completedDateTime });
    }

    if (entries.length === 0) {
      return undefined;
    }

    return (
      <ul className={styles.timestampList}>
        {entries.map((entry) => (
          <li key={entry.label} className={styles.timestampItem}>
            <span className={styles.timestampLabel}>{entry.label}:</span>
            <span className={styles.timestampValue}>{this._formatDateTime(entry.value)}</span>
          </li>
        ))}
      </ul>
    );
  }

  private _renderComments(comments: ISuggestionComment[]): React.ReactNode {
    if (comments.length === 0) {
      return undefined;
    }

    return (
      <div className={styles.commentSection}>
        <h5 className={styles.commentHeading}>Comments</h5>
        <ul className={styles.commentList}>
          {comments.map((comment) => {
            const hasMeta: boolean = !!comment.author || !!comment.createdDateTime;

            return (
              <li key={comment.id} className={styles.commentItem}>
                {hasMeta && (
                  <div className={styles.commentMeta}>
                    {comment.author && <span className={styles.commentAuthor}>{comment.author}</span>}
                    {comment.createdDateTime && (
                      <span className={styles.commentTimestamp}>
                        {this._formatDateTime(comment.createdDateTime)}
                      </span>
                    )}
                  </div>
                )}
                <p className={styles.commentText}>{comment.text}</p>
              </li>
            );
          })}
        </ul>
      </div>
    );
  }

  private _getInteractionState(
    item: ISuggestionItem,
    readOnly: boolean,
    normalizedUser: string | undefined,
    noVotesRemaining: boolean
  ): {
    hasVoted: boolean;
    disableVote: boolean;
    canAddComment: boolean;
    canMarkSuggestionAsDone: boolean;
    canDeleteSuggestion: boolean;
  } {
    const hasVoted: boolean = !!normalizedUser && item.voters.indexOf(normalizedUser) !== -1;
    const disableVote: boolean =
      this.state.isLoading || readOnly || item.status === 'Done' || (!hasVoted && noVotesRemaining);
    const canMarkSuggestionAsDone: boolean = this.props.isCurrentUserAdmin && !readOnly && item.status !== 'Done';
    const canDeleteSuggestion: boolean = this._canCurrentUserDeleteSuggestion(item);
    const canAddComment: boolean = !readOnly && item.status !== 'Done';

    return { hasVoted, disableVote, canAddComment, canMarkSuggestionAsDone, canDeleteSuggestion };
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

  private _renderPaginationControls(
    page: number,
    hasPrevious: boolean,
    hasNext: boolean,
    onPrevious: () => void,
    onNext: () => void
  ): React.ReactNode {
    if (!hasPrevious && !hasNext && page <= 1) {
      return undefined;
    }

    return (
      <div className={styles.paginationControls}>
        <DefaultButton text="Previous" onClick={onPrevious} disabled={!hasPrevious} />
        <span className={styles.paginationInfo} aria-live="polite">
          Page {page}
        </span>
        <DefaultButton text="Next" onClick={onNext} disabled={!hasNext} />
      </div>
    );
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

    return [{ key: ALL_SUBCATEGORY_FILTER_KEY, text: 'All subcategories' }, ...options];
  }

  private _getCategoryOptions(categories: SuggestionCategory[]): IDropdownOption[] {
    return categories.map((category) => ({ key: category, text: category }));
  }

  private _getFilterCategoryOptions(categories: SuggestionCategory[]): IDropdownOption[] {
    return [{ key: ALL_CATEGORY_FILTER_KEY, text: 'All categories' }, ...this._getCategoryOptions(categories)];
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
    const { items, nextToken } = await this._getSuggestionsPage('Active', options.skipToken, filter);

    if (items.length === 0 && options.previousTokens.length > 0) {
      const tokens: (string | undefined)[] = [...options.previousTokens];
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
        page: options.page,
        currentToken: options.skipToken,
        nextToken,
        previousTokens: options.previousTokens
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
    const { items, nextToken } = await this._getSuggestionsPage('Done', options.skipToken, filter);

    if (items.length === 0 && options.previousTokens.length > 0) {
      const tokens: (string | undefined)[] = [...options.previousTokens];
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
        page: options.page,
        currentToken: options.skipToken,
        nextToken,
        previousTokens: options.previousTokens
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

    let commentsBySuggestion: Map<number, ISuggestionComment[]> = new Map();

    if (suggestionIds.length > 0 && this._currentCommentsListId) {
      const commentListId: string = this._getResolvedCommentsListId();
      const commentItems: IGraphCommentItem[] = await this.props.graphService.getCommentItems(commentListId, {
        suggestionIds
      });
      commentsBySuggestion = this._groupCommentsBySuggestion(commentItems);
    }

    const items: ISuggestionItem[] = this._mapGraphItemsToSuggestions(
      response.items,
      votesBySuggestion,
      commentsBySuggestion
    );
    return { items, nextToken: response.nextToken };
  }

  private _mapGraphItemsToSuggestions(
    graphItems: IGraphSuggestionItem[],
    votesBySuggestion: Map<number, IVoteEntry[]>,
    commentsBySuggestion: Map<number, ISuggestionComment[]>
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
          comments: commentsBySuggestion.get(suggestionId) ?? []
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

  private _onTitleChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this._updateState({ newTitle: newValue ?? '' });
  };

  private _onDescriptionChange = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this._updateState({ newDescription: newValue ?? '' });
  };

  private _onActiveSearchChange = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    const nextFilter: IFilterState = { ...this.state.activeFilter, searchQuery: newValue ?? '' };
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
      subcategory: this._normalizeFilterSubcategory(nextCategory, this.state.activeFilter.subcategory, this.state.subcategories)
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
        ? { ...this.state.activeFilter, subcategory: undefined }
        : { ...this.state.activeFilter, subcategory: key };

    this._applyActiveFilter(nextFilter);
  };

  private _onCompletedSearchChange = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    const nextFilter: IFilterState = { ...this.state.completedFilter, searchQuery: newValue ?? '' };
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
      )
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
        ? { ...this.state.completedFilter, subcategory: undefined }
        : { ...this.state.completedFilter, subcategory: key };

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

      this._updateState({
        newTitle: '',
        newDescription: '',
        newCategory: defaultCategory,
        newSubcategoryKey: this._getValidSubcategoryKeyForCategory(
          defaultCategory,
          undefined
        )
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

      this._updateState({ success: 'Your comment has been added.' });
    } catch (error) {
      this._handleError('We could not add your comment. Please try again.', error);
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

  private _updateState(state: Partial<ISamverkansportalenState>): void {
    if (!this._isMounted) {
      return;
    }

    this.setState(state as Pick<ISamverkansportalenState, keyof ISamverkansportalenState>);
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
