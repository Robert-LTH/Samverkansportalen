import { MSGraphClientFactory, type MSGraphClientV3 } from '@microsoft/sp-http';

export interface IGraphListInfo {
  id: string;
  displayName: string;
}

export const SUGGESTION_CATEGORIES = ['Change request', 'Webbinar', 'Article'] as const;

export type SuggestionCategory = (typeof SUGGESTION_CATEGORIES)[number];

export interface IGraphSuggestionItemFields {
  Title?: string;
  Details?: string;
  Votes?: number | string;
  Status?: string;
  Voters?: string;
  Category?: SuggestionCategory;
}

export interface IGraphSuggestionItem {
  id: number;
  fields: IGraphSuggestionItemFields;
  createdByUserPrincipalName?: string;
}

interface IGraphListApiModel {
  id?: unknown;
  displayName?: unknown;
  list?: {
    hidden?: unknown;
    template?: unknown;
  };
}

interface IGraphListItemApiModel {
  id?: unknown;
  fields?: unknown;
  createdBy?: {
    user?: {
      userPrincipalName?: unknown;
      email?: unknown;
      mail?: unknown;
    };
  };
}

export class GraphSuggestionsService {
  private readonly _hostname: string;
  private readonly _sitePath?: string;
  private _clientPromise?: Promise<MSGraphClientV3>;
  private _siteIdPromise?: Promise<string>;

  public constructor(
    private readonly _graphClientFactory: MSGraphClientFactory,
    siteUrl: string
  ) {
    const parsedUrl: URL = new URL(siteUrl);
    this._hostname = parsedUrl.hostname;

    const trimmedPath: string = parsedUrl.pathname.replace(/\/$/, '');
    if (trimmedPath && trimmedPath !== '/') {
      this._sitePath = trimmedPath;
    }
  }

  public async getVisibleLists(): Promise<IGraphListInfo[]> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    const response: { value?: IGraphListApiModel[] } = await client
      .api(`/sites/${siteId}/lists`)
      .version('v1.0')
      .select('id,displayName,list')
      .top(999)
      .get();

    const lists: IGraphListApiModel[] = Array.isArray(response.value) ? response.value : [];

    return lists
      .map((entry) => {
        const id: unknown = entry.id;
        const displayName: unknown = entry.displayName;
        const template: unknown = entry.list?.template;
        const hidden: unknown = entry.list?.hidden;

        if (typeof id !== 'string' || typeof displayName !== 'string') {
          return undefined;
        }

        if (hidden === true) {
          return undefined;
        }

        if (template && template !== 'genericList') {
          return undefined;
        }

        return { id, displayName } as IGraphListInfo;
      })
      .filter((item): item is IGraphListInfo => !!item);
  }

  public async ensureList(listTitle: string): Promise<{ id: string; created: boolean }> {
    const existing: IGraphListInfo | undefined = await this._getListByTitle(listTitle);

    if (existing) {
      return { id: existing.id, created: false };
    }

    const created: IGraphListInfo = await this._createListWithColumns(listTitle);
    return { id: created.id, created: true };
  }

  public async getSuggestionItems(listId: string): Promise<IGraphSuggestionItem[]> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    const response: { value?: IGraphListItemApiModel[] } = await client
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .version('v1.0')
      .select('id,createdBy')
      .expand(
        'fields($select=Title,Details,Votes,Status,Voters,Category),createdByUser($select=userPrincipalName,mail,email)'
      )
      .orderby('createdDateTime desc')
      .top(999)
      .get();

    const items: IGraphListItemApiModel[] = Array.isArray(response.value) ? response.value : [];

    return items
      .map((entry) => {
        const rawId: unknown = entry.id;
        const fields: unknown = entry.fields;

        if (typeof rawId !== 'string') {
          return undefined;
        }

        const id: number = parseInt(rawId, 10);
        if (!Number.isFinite(id)) {
          return undefined;
        }

        if (!fields || typeof fields !== 'object') {
          return undefined;
        }

        let createdByUserPrincipalName: string | undefined;
        const createdBy: unknown = entry.createdBy;

        if (createdBy && typeof createdBy === 'object') {
          const user: unknown = (createdBy as { user?: unknown }).user;

          if (user && typeof user === 'object') {
            const upn: unknown = (user as { userPrincipalName?: unknown }).userPrincipalName;

            if (typeof upn === 'string' && upn.trim().length > 0) {
              createdByUserPrincipalName = upn.trim();
            } else {
              const email: unknown = (user as { email?: unknown; mail?: unknown }).email ?? (user as { email?: unknown; mail?: unknown }).mail;

              if (typeof email === 'string' && email.trim().length > 0) {
                createdByUserPrincipalName = email.trim();
              }
            }
          }
        }

        return {
          id,
          fields: fields as IGraphSuggestionItemFields,
          createdByUserPrincipalName
        } as IGraphSuggestionItem;
      })
      .filter((item): item is IGraphSuggestionItem => !!item);
  }

  public async addSuggestion(listId: string, fields: IGraphSuggestionItemFields): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await this._executeWithVoteFallback(fields, async (payload) => {
      await client
        .api(`/sites/${siteId}/lists/${listId}/items`)
        .version('v1.0')
        .post({ fields: payload });
    });
  }

  public async updateSuggestion(listId: string, itemId: number, fields: Partial<IGraphSuggestionItemFields>): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await this._executeWithVoteFallback(fields, async (payload) => {
      await client
        .api(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`)
        .version('v1.0')
        .patch(payload);
    });
  }

  public async deleteSuggestion(listId: string, itemId: number): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await client
      .api(`/sites/${siteId}/lists/${listId}/items/${itemId}`)
      .version('v1.0')
      .delete();
  }

  private async _getListByTitle(listTitle: string): Promise<IGraphListInfo | undefined> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();
    const escapedTitle: string = listTitle.replace(/'/g, "''");

    const response: { value?: IGraphListApiModel[] } = await client
      .api(`/sites/${siteId}/lists`)
      .version('v1.0')
      .select('id,displayName')
      .filter(`displayName eq '${escapedTitle}'`)
      .top(1)
      .get();

    const lists: IGraphListApiModel[] = Array.isArray(response.value) ? response.value : [];
    const match: IGraphListApiModel | undefined = lists.find((entry) => typeof entry.displayName === 'string');

    if (!match || typeof match.id !== 'string' || typeof match.displayName !== 'string') {
      return undefined;
    }

    return {
      id: match.id,
      displayName: match.displayName
    };
  }

  private async _createListWithColumns(listTitle: string): Promise<IGraphListInfo> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    const response: IGraphListApiModel = await client
      .api(`/sites/${siteId}/lists`)
      .version('v1.0')
      .post({
        displayName: listTitle,
        description: 'Stores user suggestions and votes from the Samverkansportalen web part.',
        list: {
          template: 'genericList'
        },
        columns: [
          {
            name: 'Details',
            displayName: 'Details',
            text: {
              allowMultipleLines: true
            }
          },
          {
            name: 'Category',
            displayName: 'Category',
            choice: {
              allowTextEntry: false,
              allowMultipleSelections: false,
              choices: [...SUGGESTION_CATEGORIES]
            }
          },
          {
            name: 'Votes',
            displayName: 'Votes',
            number: {
              decimalPlaces: 0
            }
          },
          {
            name: 'Status',
            displayName: 'Status',
            text: {
              allowMultipleLines: false
            }
          },
          {
            name: 'Voters',
            displayName: 'Voters',
            text: {
              allowMultipleLines: true
            }
          }
        ]
      });

    if (typeof response.id !== 'string' || typeof response.displayName !== 'string') {
      throw new Error('Failed to create the suggestions list.');
    }

    return {
      id: response.id,
      displayName: response.displayName
    };
  }

  private _getClient(): Promise<MSGraphClientV3> {
    if (!this._clientPromise) {
      this._clientPromise = this._graphClientFactory.getClient('3');
    }

    return this._clientPromise;
  }

  private async _getSiteId(): Promise<string> {
    if (!this._siteIdPromise) {
      this._siteIdPromise = this._resolveSiteId();
    }

    return this._siteIdPromise;
  }

  private async _resolveSiteId(): Promise<string> {
    const client: MSGraphClientV3 = await this._getClient();

    const requestPath: string = this._sitePath
      ? `/sites/${this._hostname}:${encodeURI(this._sitePath)}`
      : `/sites/${this._hostname}`;

    const response: { id?: unknown } = await client
      .api(requestPath)
      .version('v1.0')
      .select('id')
      .get();

    if (!response || typeof response.id !== 'string') {
      throw new Error('Failed to resolve the site identifier from Microsoft Graph.');
    }

    return response.id;
  }

  private async _executeWithVoteFallback(
    fields: Partial<IGraphSuggestionItemFields>,
    executor: (payload: Partial<IGraphSuggestionItemFields>) => Promise<void>
  ): Promise<void> {
    try {
      await executor(fields);
    } catch (error) {
      if (!this._shouldRetryWithStringVotes(fields, error)) {
        throw error;
      }

      const fallbackPayload: Partial<IGraphSuggestionItemFields> = {
        ...fields,
        Votes: String(fields.Votes)
      };

      await executor(fallbackPayload);
    }
  }

  private _shouldRetryWithStringVotes(
    fields: Partial<IGraphSuggestionItemFields>,
    error: unknown
  ): boolean {
    const votes: unknown = fields.Votes;

    if (typeof votes !== 'number' || !Number.isFinite(votes)) {
      return false;
    }

    const message: string | undefined = this._extractErrorMessage(error);

    if (!message) {
      return false;
    }

    const normalized: string = message.toLowerCase();
    return normalized.includes('cannot convert the literal') && normalized.includes('edm.string');
  }

  private _extractErrorMessage(error: unknown): string | undefined {
    if (!error || typeof error !== 'object') {
      return undefined;
    }

    const directMessage: unknown = (error as { message?: unknown }).message;
    if (typeof directMessage === 'string') {
      return directMessage;
    }

    const body: unknown = (error as { body?: unknown }).body;
    if (body && typeof body === 'object') {
      const bodyMessage: unknown = (body as { error?: unknown }).error;
      if (bodyMessage && typeof bodyMessage === 'object') {
        const nestedMessage: unknown = (bodyMessage as { message?: unknown }).message;
        if (typeof nestedMessage === 'string') {
          return nestedMessage;
        }
      }
    }

    const nestedError: unknown = (error as { error?: unknown }).error;
    if (nestedError && typeof nestedError === 'object') {
      const nestedMessage: unknown = (nestedError as { message?: unknown }).message;
      if (typeof nestedMessage === 'string') {
        return nestedMessage;
      }
    }

    return undefined;
  }
}

export default GraphSuggestionsService;
