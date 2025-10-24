import * as React from 'react';
import {
  PrimaryButton,
  DefaultButton,
  IconButton,
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
  type IGraphVoteItem
} from '../services/GraphSuggestionsService';

interface ISuggestionItem {
  id: number;
  title: string;
  description: string;
  votes: number;
  status: 'Active' | 'Done';
  voters: string[];
  category: SuggestionCategory;
  createdByLoginName?: string;
  voteEntries: IVoteEntry[];
}

interface IVoteEntry {
  id: number;
  username: string;
  votes: number;
}

interface ISamverkansportalenState {
  suggestions: ISuggestionItem[];
  isLoading: boolean;
  newTitle: string;
  newDescription: string;
  newCategory: SuggestionCategory;
  availableVotes: number;
  error?: string;
  success?: string;
}

const MAX_VOTES_PER_USER: number = 5;
const DEFAULT_SUGGESTION_CATEGORY: SuggestionCategory = 'Change request';
const CATEGORY_OPTIONS: IDropdownOption[] = SUGGESTION_CATEGORIES.map((category) => ({
  key: category,
  text: category
}));

export default class Samverkansportalen extends React.Component<ISamverkansportalenProps, ISamverkansportalenState> {
  private _isMounted: boolean = false;
  private _currentListId?: string;
  private _currentVotesListId?: string;

  public constructor(props: ISamverkansportalenProps) {
    super(props);

    this.state = {
      suggestions: [],
      isLoading: false,
      newTitle: '',
      newDescription: '',
      newCategory: DEFAULT_SUGGESTION_CATEGORY,
      availableVotes: MAX_VOTES_PER_USER
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
    if (this._normalizeListTitle(prevProps.listTitle) !== this._listTitle) {
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
      error,
      success
    } = this.state;

    const activeSuggestions: ISuggestionItem[] = suggestions.filter((item) => item.status !== 'Done');
    const completedSuggestions: ISuggestionItem[] = suggestions.filter((item) => item.status === 'Done');

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
          <h3 className={styles.sectionTitle}>Add a suggestion</h3>
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
            <PrimaryButton
              text="Submit suggestion"
              onClick={this._addSuggestion}
              disabled={isLoading || newTitle.trim().length === 0}
            />
          </div>
        </div>

        <div className={styles.suggestionSection}>
          <h3 className={styles.sectionTitle}>Active suggestions</h3>
          {isLoading ? (
            <Spinner label="Loading suggestions..." size={SpinnerSize.large} />
          ) : (
            this._renderSuggestionList(activeSuggestions, false)
          )}
        </div>

        {completedSuggestions.length > 0 && (
          <div className={styles.suggestionSection}>
            <h3 className={styles.sectionTitle}>Completed suggestions</h3>
            {this._renderSuggestionList(completedSuggestions, true)}
          </div>
        )}
      </section>
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

  private async _initialize(): Promise<void> {
    this._currentListId = undefined;
    this._currentVotesListId = undefined;
    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      await this._ensureLists();
      await this._loadSuggestions();
    } catch (error) {
      this._handleError('We could not load the suggestions list. Please refresh the page or contact your administrator.', error);
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

      return {
        id: suggestionId,
        title: typeof fields.Title === 'string' && fields.Title.trim().length > 0 ? fields.Title : 'Untitled suggestion',
        description: typeof fields.Details === 'string' ? fields.Details : '',
        votes: voteEntries.reduce((total, vote) => total + vote.votes, 0),
        status: fields.Status === 'Done' ? 'Done' : 'Active',
        category: this._normalizeCategory(fields.Category),
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

    this._updateState({
      suggestions: baseItems,
      availableVotes
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

    this._updateState({ newCategory: match ?? DEFAULT_SUGGESTION_CATEGORY });
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

    if (!title) {
      this._handleError('Please add a title before submitting your suggestion.');
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const listId: string = this._getResolvedListId();

      await this.props.graphService.addSuggestion(listId, {
        Title: title,
        Details: description,
        Status: 'Active',
        Category: category
      });

      this._updateState({
        newTitle: '',
        newDescription: '',
        newCategory: DEFAULT_SUGGESTION_CATEGORY
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

  private _normalizeCategory(value: unknown): SuggestionCategory {
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

    return DEFAULT_SUGGESTION_CATEGORY;
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

  private _handleError(message: string, error?: unknown): void {
    console.error(message, error);
    this._updateState({ error: message, success: undefined });
  }

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
