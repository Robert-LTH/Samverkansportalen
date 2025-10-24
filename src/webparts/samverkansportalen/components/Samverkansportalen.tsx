import * as React from 'react';
import {
  PrimaryButton,
  DefaultButton,
  IconButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  TextField
} from '@fluentui/react';
import styles from './Samverkansportalen.module.scss';
import { DEFAULT_SUGGESTIONS_LIST_TITLE, type ISamverkansportalenProps } from './ISamverkansportalenProps';
import {
  type IGraphSuggestionItem,
  type IGraphSuggestionItemFields
} from '../services/GraphSuggestionsService';

interface ISuggestionItem {
  id: number;
  title: string;
  description: string;
  votes: number;
  status: 'Active' | 'Done';
  voters: string[];
}

interface ISamverkansportalenState {
  suggestions: ISuggestionItem[];
  isLoading: boolean;
  newTitle: string;
  newDescription: string;
  availableVotes: number;
  error?: string;
  success?: string;
}

const MAX_VOTES_PER_USER: number = 5;

export default class Samverkansportalen extends React.Component<ISamverkansportalenProps, ISamverkansportalenState> {
  private _isMounted: boolean = false;
  private _currentListId?: string;

  public constructor(props: ISamverkansportalenProps) {
    super(props);

    this.state = {
      suggestions: [],
      isLoading: false,
      newTitle: '',
      newDescription: '',
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

    return (
      <ul className={styles.suggestionList}>
        {items.map((item) => {
          const hasVoted: boolean = item.voters.indexOf(this.props.userLoginName) !== -1;
          const disableVote: boolean = this.state.isLoading || readOnly || item.status === 'Done' || (!hasVoted && noVotesRemaining);

          return (
            <li key={item.id} className={styles.suggestionCard}>
              <div className={styles.cardHeader}>
                <div className={styles.cardText}>
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
              <div className={styles.cardActions}>
                {readOnly ? (
                  <DefaultButton text="Votes closed" disabled />
                ) : (
                  <PrimaryButton
                    text={hasVoted ? 'Remove vote' : 'Vote'}
                    onClick={() => this._toggleVote(item)}
                    disabled={disableVote}
                  />
                )}
                {!readOnly && item.status !== 'Done' && (
                  <DefaultButton
                    text="Mark as done"
                    onClick={() => this._markSuggestionAsDone(item)}
                    disabled={this.state.isLoading}
                  />
                )}
                <IconButton
                  iconProps={{ iconName: 'Delete' }}
                  title="Remove suggestion"
                  ariaLabel="Remove suggestion"
                  onClick={() => this._deleteSuggestion(item)}
                  disabled={this.state.isLoading}
                />
              </div>
            </li>
          );
        })}
      </ul>
    );
  }

  private async _initialize(): Promise<void> {
    this._currentListId = undefined;
    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      await this._ensureList();
      await this._loadSuggestions();
    } catch (error) {
      this._handleError('We could not load the suggestions list. Please refresh the page or contact your administrator.', error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private async _ensureList(): Promise<void> {
    const listTitle: string = this._listTitle;
    const result = await this.props.graphService.ensureList(listTitle);
    this._currentListId = result.id;
  }

  private async _loadSuggestions(): Promise<void> {
    const listId: string = this._getResolvedListId();
    const itemsFromGraph: IGraphSuggestionItem[] = await this.props.graphService.getSuggestionItems(listId);

    const items: ISuggestionItem[] = itemsFromGraph.map((entry: IGraphSuggestionItem) => {
      const fields: IGraphSuggestionItemFields = entry.fields;

      const rawVoters: string = typeof fields.Voters === 'string' ? fields.Voters : '[]';
      let voters: string[] = [];

      try {
        const parsed: unknown = JSON.parse(rawVoters);
        voters = Array.isArray(parsed) ? parsed.filter((value) => typeof value === 'string') : [];
      } catch (error) {
        console.warn('Failed to parse voters field for suggestion', entry.id, error);
      }

      return {
        id: entry.id,
        title: typeof fields.Title === 'string' && fields.Title.trim().length > 0 ? fields.Title : 'Untitled suggestion',
        description: typeof fields.Details === 'string' ? fields.Details : '',
        votes: this._parseVotes(fields.Votes),
        status: fields.Status === 'Done' ? 'Done' : 'Active',
        voters
      };
    });

    const usedVotes: number = items.reduce((count, item) => {
      if (item.status === 'Done') {
        return count;
      }

      return item.voters.indexOf(this.props.userLoginName) !== -1 ? count + 1 : count;
    }, 0);

    const availableVotes: number = Math.max(MAX_VOTES_PER_USER - usedVotes, 0);

    this._updateState({
      suggestions: items,
      availableVotes
    });
  }

  private _onTitleChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this._updateState({ newTitle: newValue ?? '' });
  };

  private _onDescriptionChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this._updateState({ newDescription: newValue ?? '' });
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
        Votes: 0,
        Status: 'Active',
        Voters: JSON.stringify([])
      });

      this._updateState({ newTitle: '', newDescription: '' });

      await this._loadSuggestions();

      this._updateState({ success: 'Your suggestion has been added.' });
    } catch (error) {
      this._handleError('We could not add your suggestion. Please try again.', error);
    } finally {
      this._updateState({ isLoading: false });
    }
  };

  private async _toggleVote(item: ISuggestionItem): Promise<void> {
    const hasVoted: boolean = item.voters.indexOf(this.props.userLoginName) !== -1;

    if (!hasVoted && this.state.availableVotes <= 0) {
      this._handleError('You have used all of your votes. Mark a suggestion as done or remove one of your votes to continue.');
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    const voters: string[] = hasVoted
      ? item.voters.filter((voter) => voter !== this.props.userLoginName)
      : [...item.voters, this.props.userLoginName];

    try {
      const listId: string = this._getResolvedListId();

      await this.props.graphService.updateSuggestion(listId, item.id, {
        Votes: voters.length,
        Voters: JSON.stringify(voters)
      });

      await this._loadSuggestions();

      this._updateState({ success: hasVoted ? 'Your vote has been removed.' : 'Thanks for voting!' });
    } catch (error) {
      this._handleError('We could not update your vote. Please try again.', error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private async _markSuggestionAsDone(item: ISuggestionItem): Promise<void> {
    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const listId: string = this._getResolvedListId();

      await this.props.graphService.updateSuggestion(listId, item.id, {
        Status: 'Done',
        Votes: 0,
        Voters: JSON.stringify([])
      });

      await this._loadSuggestions();

      this._updateState({ success: 'The suggestion has been marked as done.' });
    } catch (error) {
      this._handleError('We could not mark this suggestion as done. Please try again.', error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private async _deleteSuggestion(item: ISuggestionItem): Promise<void> {
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

  private _getResolvedListId(): string {
    if (!this._currentListId) {
      throw new Error('The suggestions list has not been initialized yet.');
    }

    return this._currentListId;
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
}
