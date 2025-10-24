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
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';

import styles from './Samverkansportalen.module.scss';
import { DEFAULT_SUGGESTIONS_LIST_TITLE, type ISamverkansportalenProps } from './ISamverkansportalenProps';

interface ISuggestionItem {
  id: number;
  title: string;
  description: string;
  votes: number;
  status: 'Active' | 'Done';
  voters: string[];
}

interface ISharePointSuggestionItem {
  Id: number;
  Title?: string;
  Details?: string;
  Votes?: number;
  Status?: string;
  Voters?: string;
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

    const escapedTitleForFilter: string = listTitle.replace(/'/g, "''");
    const filterQuery: string = `Title eq '${escapedTitleForFilter}'`;
    const requestUrl: string = `${this.props.siteUrl}/_api/web/lists?$select=Id&$top=1&$filter=${encodeURIComponent(filterQuery)}`;

    const response: SPHttpClientResponse = await this.props.spHttpClient.get(
      requestUrl,
      SPHttpClient.configurations.v1,
      this._createOptions()
    );

    if (!response.ok) {
      throw new Error(`Unexpected response (${response.status}) while checking for the ${listTitle} list.`);
    }

    const payload: unknown = await response.json();

    if (this._hasListMatch(payload)) {
      return;
    }

    await this._createListWithFields(listTitle);
  }

  private _hasListMatch(payload: unknown): boolean {
    const entries: Array<{ Id?: unknown }> | undefined = this._extractListEntries(payload);

    if (!entries) {
      return false;
    }

    return entries.some((entry) => !!entry && typeof entry.Id === 'number');
  }

  private _extractListEntries(payload: unknown): Array<{ Id?: unknown }> | undefined {
    if (!payload || typeof payload !== 'object') {
      return undefined;
    }

    const withValue = payload as { value?: unknown };
    if (Array.isArray(withValue.value)) {
      return withValue.value as Array<{ Id?: unknown }>;
    }

    const withVerbose = payload as { d?: { results?: unknown } };
    if (withVerbose.d && Array.isArray(withVerbose.d.results)) {
      return withVerbose.d.results as Array<{ Id?: unknown }>;
    }

    return undefined;
  }

  private async _createListWithFields(listTitle: string): Promise<void> {
    const createListResponse: SPHttpClientResponse = await this.props.spHttpClient.post(
      `${this.props.siteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      this._createOptions({
        Title: listTitle,
        Description: 'Stores user suggestions and votes from the Samverkansportalen web part.',
        BaseTemplate: 100,
        AllowContentTypes: true
      })
    );

    if (!createListResponse.ok) {
      throw new Error('Failed to create the suggestions list.');
    }

    await this._createField({
      __metadata: { type: 'SP.FieldMultiLineText' },
      Title: 'Details',
      FieldTypeKind: 3,
      RichText: false,
      NumberOfLines: 6
    });

    await this._createField({
      Title: 'Votes',
      FieldTypeKind: 9,
      MinimumValue: 0,
      DefaultValue: '0'
    });

    await this._createField({
      Title: 'Status',
      FieldTypeKind: 6,
      Choices: {
        results: ['Active', 'Done']
      },
      DefaultValue: 'Active'
    });

    await this._createField({
      __metadata: { type: 'SP.FieldMultiLineText' },
      Title: 'Voters',
      FieldTypeKind: 3,
      RichText: false,
      NumberOfLines: 6
    });
  }

  private async _createField(definition: Record<string, unknown>): Promise<void> {
    const response: SPHttpClientResponse = await this.props.spHttpClient.post(
      `${this._listEndpoint}/fields`,
      SPHttpClient.configurations.v1,
      this._createOptions(definition)
    );

    if (!response.ok) {
      throw new Error(`Failed to create field ${(definition.Title as string) || 'unknown'}.`);
    }
  }

  private async _loadSuggestions(): Promise<void> {
    const response: SPHttpClientResponse = await this.props.spHttpClient.get(
      `${this._listEndpoint}/items?$select=Id,Title,Details,Votes,Status,Voters&$orderby=Created desc`,
      SPHttpClient.configurations.v1,
      this._createOptions()
    );

    if (!response.ok) {
      throw new Error('Failed to read suggestions from the list.');
    }

    const payload: { value: ISharePointSuggestionItem[] } = await response.json() as { value: ISharePointSuggestionItem[] };

    const items: ISuggestionItem[] = payload.value.map((entry: ISharePointSuggestionItem) => {
      const rawVoters: string = entry.Voters || '[]';
      let voters: string[] = [];

      try {
        const parsed: unknown = JSON.parse(rawVoters);
        voters = Array.isArray(parsed) ? parsed.filter((value) => typeof value === 'string') : [];
      } catch (error) {
        console.warn('Failed to parse voters field for suggestion', entry.Id, error);
      }

      return {
        id: entry.Id,
        title: entry.Title || 'Untitled suggestion',
        description: entry.Details || '',
        votes: typeof entry.Votes === 'number' ? entry.Votes : 0,
        status: entry.Status === 'Done' ? 'Done' : 'Active',
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

  private async _addSuggestion(): Promise<void> {
    const title: string = this.state.newTitle.trim();
    const description: string = this.state.newDescription.trim();

    if (!title) {
      this._handleError('Please add a title before submitting your suggestion.');
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const response: SPHttpClientResponse = await this.props.spHttpClient.post(
        `${this._listEndpoint}/items`,
        SPHttpClient.configurations.v1,
        this._createOptions({
          Title: title,
          Details: description,
          Votes: 0,
          Status: 'Active',
          Voters: JSON.stringify([])
        })
      );

      if (!response.ok) {
        throw new Error('Failed to create the suggestion.');
      }

      this._updateState({ newTitle: '', newDescription: '' });

      await this._loadSuggestions();

      this._updateState({ success: 'Your suggestion has been added.' });
    } catch (error) {
      this._handleError('We could not add your suggestion. Please try again.', error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

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
      const response: SPHttpClientResponse = await this.props.spHttpClient.post(
        `${this._listEndpoint}/items(${item.id})`,
        SPHttpClient.configurations.v1,
        this._createOptions(
          {
            Votes: voters.length,
            Voters: JSON.stringify(voters)
          },
          {
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          }
        )
      );

      if (!response.ok) {
        throw new Error('Failed to update the vote.');
      }

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
      const response: SPHttpClientResponse = await this.props.spHttpClient.post(
        `${this._listEndpoint}/items(${item.id})`,
        SPHttpClient.configurations.v1,
        this._createOptions(
          {
            Status: 'Done',
            Votes: 0,
            Voters: JSON.stringify([])
          },
          {
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          }
        )
      );

      if (!response.ok) {
        throw new Error('Failed to update the suggestion status.');
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
    const confirmation: boolean = window.confirm('Are you sure you want to remove this suggestion? This action cannot be undone.');

    if (!confirmation) {
      return;
    }

    this._updateState({ isLoading: true, error: undefined, success: undefined });

    try {
      const response: SPHttpClientResponse = await this.props.spHttpClient.post(
        `${this._listEndpoint}/items(${item.id})`,
        SPHttpClient.configurations.v1,
        this._createOptions(
          undefined,
          {
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        )
      );

      if (!response.ok) {
        throw new Error('Failed to delete the suggestion.');
      }

      await this._loadSuggestions();

      this._updateState({ success: 'The suggestion has been removed.' });
    } catch (error) {
      this._handleError('We could not remove this suggestion. Please try again.', error);
    } finally {
      this._updateState({ isLoading: false });
    }
  }

  private _createOptions(body?: unknown, extraHeaders?: Record<string, string>): ISPHttpClientOptions {
    const headers: Record<string, string> = {
      'Accept': 'application/json;odata=nometadata',
      'odata-version': '3.0'
    };

    if (body !== undefined) {
      headers['Content-type'] = 'application/json;odata=nometadata';
    }

    if (extraHeaders) {
      for (const key in extraHeaders) {
        if (Object.prototype.hasOwnProperty.call(extraHeaders, key)) {
          const value: string | undefined = extraHeaders[key];
          if (typeof value === 'string') {
            headers[key] = value;
          }
        }
      }
    }

    const options: ISPHttpClientOptions = {
      headers
    };

    if (body !== undefined) {
      options.body = JSON.stringify(body);
    }

    return options;
  }

  private get _listEndpoint(): string {
    const escapedTitle: string = this._listTitle.replace(/'/g, "''");
    return `${this.props.siteUrl}/_api/web/lists/GetByTitle('${escapedTitle}')`;
  }

  private _normalizeListTitle(value?: string): string {
    const trimmed: string = (value ?? '').trim();
    return trimmed.length > 0 ? trimmed : DEFAULT_SUGGESTIONS_LIST_TITLE;
  }

  private get _listTitle(): string {
    return this._normalizeListTitle(this.props.listTitle);
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
