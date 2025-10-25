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
  SUGGESTION_CATEGORIES,
  type SuggestionCategory,
  type IGraphSuggestionItem,
  type IGraphSuggestionItemFields,
  type IGraphVoteItem,
  type IGraphSubcategoryItem
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
  voteEntries: IVoteEntry[];
}

interface IVoteEntry {
  id: number;
  username: string;
  votes: number;
}

interface ISubcategoryDefinition {
  key: string;
  title: string;
  category?: SuggestionCategory;
}

interface ISamverkansportalenState {
  suggestions: ISuggestionItem[];
  isLoading: boolean;
  newTitle: string;
  newDescription: string;
  newCategory: SuggestionCategory;
  newSubcategoryKey?: string;
  subcategories: ISubcategoryDefinition[];
  availableVotes: number;
  filterCategory?: SuggestionCategory;
  filterSubcategory?: string;
  searchQuery: string;
  error?: string;
  success?: string;
  completedPage: number;
  isAddSuggestionExpanded: boolean;
  isFilterExpanded: boolean;
  isActiveSuggestionsExpanded: boolean;
  isCompletedSuggestionsExpanded: boolean;
}

const MAX_VOTES_PER_USER: number = 5;
const DEFAULT_SUGGESTION_CATEGORY: SuggestionCategory = 'Change request';
const CATEGORY_OPTIONS: IDropdownOption[] = SUGGESTION_CATEGORIES.map((category) => ({
  key: category,
  text: category
}));
const ALL_CATEGORY_FILTER_KEY: string = '__all_categories__';
const FILTER_CATEGORY_OPTIONS: IDropdownOption[] = [
  { key: ALL_CATEGORY_FILTER_KEY, text: 'All categories' },
  ...CATEGORY_OPTIONS
];
const ALL_SUBCATEGORY_FILTER_KEY: string = '__all_subcategories__';
const COMPLETED_SUGGESTIONS_PAGE_SIZE: number = 5;

export default class Samverkansportalen extends React.Component<ISamverkansportalenProps, ISamverkansportalenState> {
  private _isMounted: boolean = false;
  private _currentListId?: string;
  private _currentVotesListId?: string;
  private _currentSubcategoryListId?: string;
  private readonly _sectionIds: {
    add: { title: string; content: string };
    filter: { title: string; content: string };
    active: { title: string; content: string };
    completed: { title: string; content: string };
  };

  public constructor(props: ISamverkansportalenProps) {
    super(props);

    const uniquePrefix: string = `samverkansportalen-${Math.random().toString(36).slice(2, 10)}`;
    this._sectionIds = {
      add: { title: `${uniquePrefix}-add-title`, content: `${uniquePrefix}-add-content` },
      filter: { title: `${uniquePrefix}-filter-title`, content: `${uniquePrefix}-filter-content` },
      active: { title: `${uniquePrefix}-active-title`, content: `${uniquePrefix}-active-content` },
      completed: {
        title: `${uniquePrefix}-completed-title`,
        content: `${uniquePrefix}-completed-content`
      }
    };

    this.state = {
      suggestions: [],
      isLoading: false,
      newTitle: '',
      newDescription: '',
      newCategory: DEFAULT_SUGGESTION_CATEGORY,
      newSubcategoryKey: undefined,
      subcategories: [],
      availableVotes: MAX_VOTES_PER_USER,
      filterCategory: undefined,
      filterSubcategory: undefined,
      searchQuery: '',
      completedPage: 1,
      isAddSuggestionExpanded: true,
      isFilterExpanded: true,
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
    const subcategoryListChanged: boolean =
      this._normalizeOptionalListTitle(prevProps.subcategoryListTitle) !== this._subcategoryListTitle;

    if (listChanged || subcategoryListChanged) {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this._initialize();
    }
  }

  public render(): React.ReactElement<ISamverkansportalenProps> {
    const {
      suggestions,
      isLoading,
      availableVotes,
      newTitle,
      newDescription,
      newCategory,
      newSubcategoryKey,
      subcategories,
      filterCategory,
      filterSubcategory,
      searchQuery,
      error,
      success,
      isAddSuggestionExpanded,
      isFilterExpanded,
      isActiveSuggestionsExpanded,
      isCompletedSuggestionsExpanded
    } = this.state;

    const subcategoryOptions: IDropdownOption[] = this._getSubcategoryOptions(newCategory, subcategories);
    const filterSubcategoryOptions: IDropdownOption[] = this._getFilterSubcategoryOptions(
      filterCategory,
      subcategories,
      suggestions
    );
    const activeSuggestions: ISuggestionItem[] = this._getFilteredSuggestions(
      suggestions.filter((item) => item.status !== 'Done'),
      filterCategory,
      filterSubcategory,
      searchQuery
    );
    const completedSuggestions: ISuggestionItem[] = this._getFilteredSuggestions(
      suggestions.filter((item) => item.status === 'Done'),
      filterCategory,
      filterSubcategory,
      ''
    );

    const totalCompletedPages: number = Math.max(
      1,
      Math.ceil(completedSuggestions.length / COMPLETED_SUGGESTIONS_PAGE_SIZE)
    );
    const completedPage: number = Math.min(this.state.completedPage, totalCompletedPages);
    const paginatedCompletedSuggestions: ISuggestionItem[] = completedSuggestions.slice(
      (completedPage - 1) * COMPLETED_SUGGESTIONS_PAGE_SIZE,
      completedPage * COMPLETED_SUGGESTIONS_PAGE_SIZE
    );

    return (
      <section className={`${styles.samverkansportalen} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        <header className={styles.header}>
          <div>
            <h2 className={styles.title}>Suggestion board</h2>
            <p className={styles.subtitle}>Share ideas, cast your votes and celebrate what has been delivered.</p>
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
                  options={CATEGORY_OPTIONS}
                  selectedKey={newCategory}
                  onChange={this._onCategoryChange}
                  disabled={isLoading}
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

        <div className={styles.filters}>
          {this._renderSectionHeader(
            'Filter suggestions',
            this._sectionIds.filter.title,
            this._sectionIds.filter.content,
            isFilterExpanded,
            this._toggleFilterSection
          )}
          <div
            id={this._sectionIds.filter.content}
            role="region"
            aria-labelledby={this._sectionIds.filter.title}
            className={`${styles.sectionContent} ${
              isFilterExpanded ? '' : styles.sectionContentCollapsed
            }`}
            hidden={!isFilterExpanded}
          >
            {isFilterExpanded && (
              <div className={styles.filterControls}>
                <TextField
                  label="Search"
                  value={searchQuery}
                  onChange={this._onSearchChange}
                  disabled={isLoading}
                  placeholder="Search by title or details"
                  className={styles.filterSearch}
                />
                <Dropdown
                  label="Category"
                  options={FILTER_CATEGORY_OPTIONS}
                  selectedKey={filterCategory ?? ALL_CATEGORY_FILTER_KEY}
                  onChange={this._onFilterCategoryChange}
                  disabled={isLoading}
                  className={styles.filterDropdown}
                />
                <Dropdown
                  label="Subcategory"
                  options={filterSubcategoryOptions}
                  selectedKey={filterSubcategory ?? ALL_SUBCATEGORY_FILTER_KEY}
                  onChange={this._onFilterSubcategoryChange}
                  disabled={isLoading || filterSubcategoryOptions.length <= 1}
                  className={styles.filterDropdown}
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
            {isActiveSuggestionsExpanded &&
              (isLoading ? (
                <Spinner label="Loading suggestions..." size={SpinnerSize.large} />
              ) : (
                this._renderSuggestionList(activeSuggestions, false)
              ))}
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
                {this._renderSuggestionList(paginatedCompletedSuggestions, true)}
                {totalCompletedPages > 1 && (
                  <div className={styles.paginationControls}>
                    <DefaultButton
                      text="Previous"
                      onClick={() => this._goToPreviousCompletedPage(completedPage)}
                      disabled={completedPage <= 1}
                    />
                    <span className={styles.paginationInfo} aria-live="polite">
                      Page {completedPage} of {totalCompletedPages}
                    </span>
                    <DefaultButton
                      text="Next"
                      onClick={() => this._goToNextCompletedPage(completedPage, totalCompletedPages)}
                      disabled={completedPage >= totalCompletedPages}
                    />
                  </div>
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

  private _setCompletedPage(page: number): void {
    const nextPage: number = Math.max(1, page);

    if (nextPage !== this.state.completedPage) {
      this._updateState({ completedPage: nextPage });
    }
  }

  private _goToPreviousCompletedPage = (currentPage: number): void => {
    if (currentPage <= 1) {
      return;
    }

    this._setCompletedPage(currentPage - 1);
  };

  private _goToNextCompletedPage = (currentPage: number, totalPages: number): void => {
    if (currentPage >= totalPages) {
      return;
    }

    this._setCompletedPage(Math.min(currentPage + 1, totalPages));
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
          const { hasVoted, disableVote, canMarkSuggestionAsDone, canDeleteSuggestion } = this._getInteractionState(
            item,
            readOnly,
            normalizedUser,
            noVotesRemaining
          );

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
                </div>
                <div className={styles.voteBadge} aria-label={`${item.votes} ${item.votes === 1 ? 'vote' : 'votes'}`}>
                  <span className={styles.voteNumber}>{item.votes}</span>
                  <span className={styles.voteText}>{item.votes === 1 ? 'vote' : 'votes'}</span>
                </div>
              </div>
              {this._renderActionButtons(
                item,
                readOnly,
                hasVoted,
                disableVote,
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
            const { hasVoted, disableVote, canMarkSuggestionAsDone, canDeleteSuggestion } = this._getInteractionState(
              item,
              readOnly,
              normalizedUser,
              noVotesRemaining
            );

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

  private _getInteractionState(
    item: ISuggestionItem,
    readOnly: boolean,
    normalizedUser: string | undefined,
    noVotesRemaining: boolean
  ): {
    hasVoted: boolean;
    disableVote: boolean;
    canMarkSuggestionAsDone: boolean;
    canDeleteSuggestion: boolean;
  } {
    const hasVoted: boolean = !!normalizedUser && item.voters.indexOf(normalizedUser) !== -1;
    const disableVote: boolean =
      this.state.isLoading || readOnly || item.status === 'Done' || (!hasVoted && noVotesRemaining);
    const canMarkSuggestionAsDone: boolean = this.props.isCurrentUserAdmin && !readOnly && item.status !== 'Done';
    const canDeleteSuggestion: boolean = this._canCurrentUserDeleteSuggestion(item);

    return { hasVoted, disableVote, canMarkSuggestionAsDone, canDeleteSuggestion };
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
    definitions: ISubcategoryDefinition[],
    suggestions: ISuggestionItem[]
  ): IDropdownOption[] {
    const availableValues: string[] = this._getAvailableFilterSubcategoryValues(
      category,
      definitions,
      suggestions
    );

    const options: IDropdownOption[] = availableValues.map((value) => ({
      key: value,
      text: value
    }));

    return [{ key: ALL_SUBCATEGORY_FILTER_KEY, text: 'All subcategories' }, ...options];
  }

  private _getSubcategoriesForCategory(
    category: SuggestionCategory,
    definitions: ISubcategoryDefinition[] = this.state.subcategories
  ): ISubcategoryDefinition[] {
    return definitions.filter((definition) => !definition.category || definition.category === category);
  }

  private _getAvailableFilterSubcategoryValues(
    category: SuggestionCategory | undefined,
    definitions: ISubcategoryDefinition[],
    suggestions: ISuggestionItem[]
  ): string[] {
    const values: string[] = [];
    const relevantDefinitions: ISubcategoryDefinition[] = category
      ? this._getSubcategoriesForCategory(category, definitions)
      : definitions;

    relevantDefinitions.forEach((definition) => {
      const trimmed: string = definition.title.trim();
      if (trimmed) {
        values.push(trimmed);
      }
    });

    suggestions.forEach((item) => {
      if (category && item.category !== category) {
        return;
      }

      if (item.subcategory) {
        values.push(item.subcategory);
      }
    });

    return Array.from(values).sort((a: string, b: string) => a.localeCompare(b));
  }

  private _getValidFilterSubcategory(
    category: SuggestionCategory | undefined,
    preferredSubcategory: string | undefined,
    definitions: ISubcategoryDefinition[],
    suggestions: ISuggestionItem[]
  ): string | undefined {
    if (!preferredSubcategory) {
      return undefined;
    }

    const availableValues: string[] = this._getAvailableFilterSubcategoryValues(
      category,
      definitions,
      suggestions
    );

    return availableValues.filter( x => x === preferredSubcategory).length > 0 ? preferredSubcategory : undefined;
  }

  private _getFilteredSuggestions(
    suggestions: ISuggestionItem[],
    category: SuggestionCategory | undefined,
    subcategory: string | undefined,
    searchQuery: string
  ): ISuggestionItem[] {
    const normalizedQuery: string = searchQuery.trim().toLowerCase();

    return suggestions.filter((item) => {
      if (category && item.category !== category) {
        return false;
      }

      if (subcategory && item.subcategory !== subcategory) {
        return false;
      }

      if (normalizedQuery.length > 0) {
        const title: string = item.title.toLowerCase();
        const description: string = item.description ? item.description.toLowerCase() : '';

        if (!title.includes(normalizedQuery) && !description.includes(normalizedQuery)) {
          return false;
        }
      }

      return true;
    });
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
    this._currentSubcategoryListId = undefined;
    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      await this._ensureLists();
      await this._ensureSubcategoryList();
      await this._loadSuggestions();
    } catch (error) {
      const message: string =
        error instanceof Error && error.message.includes('subcategory list')
          ? 'We could not load the configured subcategory list. Please verify the configuration or remove it.'
          : 'We could not load the suggestions list. Please refresh the page or contact your administrator.';
      this._handleError(message, error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private async _ensureLists(): Promise<void> {
    const listTitle: string = this._listTitle;
    const result = await this.props.graphService.ensureList(listTitle);
    this._currentListId = result.id;

    const votesResult = await this.props.graphService.ensureVoteList(listTitle);
    this._currentVotesListId = votesResult.id;
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
    const listId: string = this._getResolvedListId();
    const voteListId: string = this._getResolvedVotesListId();
    const itemsFromGraph: IGraphSuggestionItem[] = await this.props.graphService.getSuggestionItems(listId);
    const votesFromGraph: IGraphVoteItem[] = await this.props.graphService.getVoteItems(voteListId);

    const votesBySuggestion: Map<number, IVoteEntry[]> = new Map();

    votesFromGraph.forEach((entry: IGraphVoteItem) => {
      const fields = entry.fields ?? {};

      const suggestionId: number | undefined = this._parseNumericId((fields as { SuggestionId?: unknown }).SuggestionId);
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

    const baseItems = itemsFromGraph.map((entry: IGraphSuggestionItem): ISuggestionItem => {
      const fields: IGraphSuggestionItemFields = entry.fields;

      const rawId: unknown = fields.id ?? (fields as { Id?: unknown }).Id;
      const suggestionId: number = this._parseNumericId(rawId) ?? -1;
      const voteEntries: IVoteEntry[] = votesBySuggestion.get(suggestionId) ?? [];

      const storedVotes: number = this._parseVotes(fields.Votes);
      const status: 'Active' | 'Done' = fields.Status === 'Done' ? 'Done' : 'Active';
      const liveVotes: number = voteEntries.reduce((total, vote) => total + vote.votes, 0);
      const votes: number = status === 'Done' ? Math.max(liveVotes, storedVotes) : liveVotes;

      return {
        id: suggestionId,
        title: typeof fields.Title === 'string' && fields.Title.trim().length > 0 ? fields.Title : 'Untitled suggestion',
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
        voteEntries
      };
    });

    const items: ISuggestionItem[] = baseItems.map((item, index) => ({
      ...item,
      displayId: index + 1
    }));

    const normalizedUser: string | undefined = this._normalizeLoginName(this.props.userLoginName);

    const usedVotes: number = items.reduce((count, item) => {
      if (item.status === 'Done' || !normalizedUser) {
        return count;
      }

      const voteForUser: IVoteEntry | undefined = item.voteEntries.find((vote) => vote.username === normalizedUser);

      if (!voteForUser) {
        return count;
      }

      return count + voteForUser.votes;
    }, 0);

    const availableVotes: number = Math.max(MAX_VOTES_PER_USER - usedVotes, 0);

    const nextFilterSubcategory: string | undefined = this._getValidFilterSubcategory(
      this.state.filterCategory,
      this.state.filterSubcategory,
      this.state.subcategories,
      baseItems
    );

    this._updateState({
      suggestions: baseItems,
      availableVotes,
      filterSubcategory: nextFilterSubcategory,
      completedPage: 1
    });
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

    const nextFilterSubcategory: string | undefined = this._getValidFilterSubcategory(
      this.state.filterCategory,
      this.state.filterSubcategory,
      definitions,
      this.state.suggestions
    );

    this._updateState({
      subcategories: definitions,
      newSubcategoryKey: nextSubcategoryKey,
      filterSubcategory: nextFilterSubcategory
    });
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

  private _onSearchChange = (
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    this._updateState({ searchQuery: newValue ?? '' });
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
    const match: SuggestionCategory | undefined = SUGGESTION_CATEGORIES.find(
      (category) => category.toLowerCase() === normalized.toLowerCase()
    );

    const nextCategory: SuggestionCategory = match ?? DEFAULT_SUGGESTION_CATEGORY;
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

  private _onFilterCategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key).trim();

    if (key === ALL_CATEGORY_FILTER_KEY) {
      const nextFilterSubcategory: string | undefined = this._getValidFilterSubcategory(
        undefined,
        this.state.filterSubcategory,
        this.state.subcategories,
        this.state.suggestions
      );

      this._setCompletedPage(1);
      this._updateState({ filterCategory: undefined, filterSubcategory: nextFilterSubcategory });
      return;
    }

    const match: SuggestionCategory | undefined = SUGGESTION_CATEGORIES.find(
      (category) => category.toLowerCase() === key.toLowerCase()
    );

    const nextCategory: SuggestionCategory | undefined = match;
    const nextFilterSubcategory: string | undefined = this._getValidFilterSubcategory(
      nextCategory,
      this.state.filterSubcategory,
      this.state.subcategories,
      this.state.suggestions
    );

    this._setCompletedPage(1);
    this._updateState({ filterCategory: nextCategory, filterSubcategory: nextFilterSubcategory });
  };

  private _onFilterSubcategoryChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) {
      return;
    }

    const key: string = String(option.key);

    if (key === ALL_SUBCATEGORY_FILTER_KEY) {
      this._setCompletedPage(1);
      this._updateState({ filterSubcategory: undefined });
      return;
    }

    this._setCompletedPage(1);
    this._updateState({ filterSubcategory: key });
  };

  private _dismissError = (): void => {
    this._updateState({ error: undefined });
  };

  private _dismissSuccess = (): void => {
    this._updateState({ success: undefined });
  };

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

      this._updateState({
        newTitle: '',
        newDescription: '',
        newCategory: DEFAULT_SUGGESTION_CATEGORY,
        newSubcategoryKey: this._getValidSubcategoryKeyForCategory(
          DEFAULT_SUGGESTION_CATEGORY,
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

      await this._loadSuggestions();

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

  private async _markSuggestionAsDone(item: ISuggestionItem): Promise<void> {
    if (!this.props.isCurrentUserAdmin) {
      this._handleError('Only administrators can mark suggestions as done.');
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const listId: string = this._getResolvedListId();
      const voteListId: string = this._getResolvedVotesListId();

      await this.props.graphService.updateSuggestion(listId, item.id, {
        Status: 'Done',
        Votes: item.votes
      });

      await this.props.graphService.deleteVotesForSuggestion(voteListId, item.id);

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

      await this.props.graphService.deleteSuggestion(listId, item.id);

      await this._loadSuggestions();

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
    if (typeof value === 'string') {
      const normalized: string = value.trim();

      if (normalized.length > 0) {
        const match: SuggestionCategory | undefined = SUGGESTION_CATEGORIES.find(
          (category) => category.toLowerCase() === normalized.toLowerCase()
        );

        if (match) {
          return match;
        }
      }
    }

    return undefined;
  }

  private _normalizeCategory(value: unknown): SuggestionCategory {
    return this._tryNormalizeCategory(value) ?? DEFAULT_SUGGESTION_CATEGORY;
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

  private _toggleFilterSection = (): void => {
    if (!this._isMounted) {
      return;
    }

    this.setState((prevState) => ({
      isFilterExpanded: !prevState.isFilterExpanded
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
