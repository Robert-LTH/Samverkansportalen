import { MSGraphClientFactory, type MSGraphClientV3 } from '@microsoft/sp-http';

export interface IGraphListInfo {
  id: string;
  displayName: string;
}

export const DEFAULT_CATEGORY_LIST_TITLE: string = 'Suggestion categories';
export const DEFAULT_SUBCATEGORY_LIST_TITLE: string = 'Suggestion subcategories';
export const DEFAULT_COMMENT_LIST_TITLE: string = 'Suggestion comments';

export type SuggestionCategory = string;

export interface IGraphSuggestionItemFields extends Record<string, unknown> {
  id?: number | string;
  Title?: string;
  Details?: string;
  Votes?: number | string;
  Status?: string;
  Voters?: string;
  Category?: SuggestionCategory;
  Subcategory?: string;
  CompletedDateTime?: string;
}

export interface IGraphSuggestionItem {
  fields: IGraphSuggestionItemFields;
  createdByUserPrincipalName?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
}

export interface IGraphVoteItemFields extends Record<string, unknown> {
  SuggestionId?: number | string;
  Username?: string;
  Votes?: number | string;
}

export interface IGraphVoteItem {
  id: number;
  fields: IGraphVoteItemFields;
}

export interface IGraphSubcategoryItemFields extends Record<string, unknown> {
  Title?: string;
  Category?: string;
}

export interface IGraphSubcategoryItem {
  id: number;
  fields: IGraphSubcategoryItemFields;
}

export interface IGraphCategoryItemFields extends Record<string, unknown> {
  Title?: string;
}

export interface IGraphCategoryItem {
  id: number;
  fields: IGraphCategoryItemFields;
}

export interface IGraphCommentItemFields extends Record<string, unknown> {
  SuggestionId?: number | string;
  Comment?: string;
  Title?: string;
}

export interface IGraphCommentItem {
  id: number;
  fields: IGraphCommentItemFields;
  createdByUserPrincipalName?: string;
  createdByUserDisplayName?: string;
  createdDateTime?: string;
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
      displayName?: unknown;
      userPrincipalName?: unknown;
      email?: unknown;
      mail?: unknown;
    };
  };
  createdDateTime?: unknown;
  lastModifiedDateTime?: unknown;
}

interface IGraphColumnApiModel {
  id?: unknown;
  name?: unknown;
  indexed?: unknown;
}

interface IListColumnDefinition {
  name: string;
  shouldBeIndexed?: boolean;
  createPayload: () => Record<string, unknown>;
}

const SUGGESTION_COLUMN_DEFINITIONS: IListColumnDefinition[] = [
  {
    name: 'Details',
    createPayload: () => ({
      name: 'Details',
      displayName: 'Details',
      text: {
        allowMultipleLines: true
      }
    })
  },
  {
    name: 'Votes',
    createPayload: () => ({
      name: 'Votes',
      displayName: 'Votes',
      number: {
        decimalPlaces: '0'
      }
    })
  },
  {
    name: 'Category',
    shouldBeIndexed: true,
    createPayload: () => ({
      name: 'Category',
      displayName: 'Category',
      indexed: true,
      text: {
        allowMultipleLines: false
      }
    })
  },
  {
    name: 'Subcategory',
    shouldBeIndexed: true,
    createPayload: () => ({
      name: 'Subcategory',
      displayName: 'Subcategory',
      indexed: true,
      text: {
        allowMultipleLines: false
      }
    })
  },
  {
    name: 'Status',
    shouldBeIndexed: true,
    createPayload: () => ({
      name: 'Status',
      displayName: 'Status',
      indexed: true,
      text: {
        allowMultipleLines: false
      }
    })
  },
  {
    name: 'CompletedDateTime',
    shouldBeIndexed: true,
    createPayload: () => ({
      name: 'CompletedDateTime',
      displayName: 'CompletedDateTime',
      indexed: true,
      dateTime: {
        displayAs: 'default'
      }
    })
  }
];

const VOTE_COLUMN_DEFINITIONS: IListColumnDefinition[] = [
  {
    name: 'SuggestionId',
    shouldBeIndexed: true,
    createPayload: () => ({
      name: 'SuggestionId',
      displayName: 'SuggestionId',
      indexed: true,
      number: {
        decimalPlaces: '0'
      }
    })
  },
  {
    name: 'Username',
    shouldBeIndexed: true,
    createPayload: () => ({
      name: 'Username',
      displayName: 'Username',
      indexed: true,
      text: {
        allowMultipleLines: false
      }
    })
  },
  {
    name: 'Votes',
    createPayload: () => ({
      name: 'Votes',
      displayName: 'Votes',
      number: {
        decimalPlaces: '0'
      }
    })
  }
];

const SUBCATEGORY_COLUMN_DEFINITIONS: IListColumnDefinition[] = [
  {
    name: 'Category',
    shouldBeIndexed: true,
    createPayload: () => ({
      name: 'Category',
      displayName: 'Category',
      indexed: true,
      text: {
        allowMultipleLines: false
      }
    })
  }
];

const COMMENT_COLUMN_DEFINITIONS: IListColumnDefinition[] = [
  {
    name: 'SuggestionId',
    shouldBeIndexed: true,
    createPayload: () => ({
      name: 'SuggestionId',
      displayName: 'SuggestionId',
      indexed: true,
      number: {
        decimalPlaces: '0'
      }
    })
  },
  {
    name: 'Comment',
    createPayload: () => ({
      name: 'Comment',
      displayName: 'Comment',
      text: {
        allowMultipleLines: true
      }
    })
  }
];

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
      await this._ensureColumns(existing.id, SUGGESTION_COLUMN_DEFINITIONS);
      return { id: existing.id, created: false };
    }

    const created: IGraphListInfo = await this._createListWithColumns(
      listTitle,
      'Stores user suggestions and votes from the Samverkansportalen web part.',
      SUGGESTION_COLUMN_DEFINITIONS
    );
    await this._ensureColumns(created.id, SUGGESTION_COLUMN_DEFINITIONS);
    return { id: created.id, created: true };
  }

  public async ensureVoteList(listTitle: string): Promise<{ id: string; created: boolean }> {
    const normalizedTitle: string = listTitle.trim().length > 0 ? listTitle.trim() : 'Votes';
    const existing: IGraphListInfo | undefined = await this._getListByTitle(normalizedTitle);

    if (existing) {
      await this._ensureColumns(existing.id, VOTE_COLUMN_DEFINITIONS);
      return { id: existing.id, created: false };
    }

    const created: IGraphListInfo = await this._createListWithColumns(
      normalizedTitle,
      'Stores suggestion votes for the Samverkansportalen web part.',
      VOTE_COLUMN_DEFINITIONS
    );
    await this._ensureColumns(created.id, VOTE_COLUMN_DEFINITIONS);
    return { id: created.id, created: true };
  }

  public async ensureSubcategoryList(listTitle: string): Promise<{ id: string; created: boolean }> {
    const normalizedTitle: string =
      listTitle.trim().length > 0 ? listTitle.trim() : DEFAULT_SUBCATEGORY_LIST_TITLE;
    const existing: IGraphListInfo | undefined = await this._getListByTitle(normalizedTitle);

    if (existing) {
      await this._ensureColumns(existing.id, SUBCATEGORY_COLUMN_DEFINITIONS);
      return { id: existing.id, created: false };
    }

    const created: IGraphListInfo = await this._createListWithColumns(
      normalizedTitle,
      'Defines suggestion subcategories for the Samverkansportalen web part.',
      SUBCATEGORY_COLUMN_DEFINITIONS
    );
    await this._ensureColumns(created.id, SUBCATEGORY_COLUMN_DEFINITIONS);
    return { id: created.id, created: true };
  }

  public async ensureCategoryList(listTitle: string): Promise<{ id: string; created: boolean }> {
    const normalizedTitle: string =
      listTitle.trim().length > 0 ? listTitle.trim() : DEFAULT_CATEGORY_LIST_TITLE;
    const existing: IGraphListInfo | undefined = await this._getListByTitle(normalizedTitle);

    if (existing) {
      return { id: existing.id, created: false };
    }

    const created: IGraphListInfo = await this._createListWithColumns(
      normalizedTitle,
      'Defines suggestion categories for the Samverkansportalen web part.',
      []
    );
    return { id: created.id, created: true };
  }

  public async ensureCommentList(listTitle: string): Promise<{ id: string; created: boolean }> {
    const normalizedTitle: string =
      listTitle.trim().length > 0 ? listTitle.trim() : DEFAULT_COMMENT_LIST_TITLE;
    const existing: IGraphListInfo | undefined = await this._getListByTitle(normalizedTitle);

    if (existing) {
      await this._ensureColumns(existing.id, COMMENT_COLUMN_DEFINITIONS);
      return { id: existing.id, created: false };
    }

    const created: IGraphListInfo = await this._createListWithColumns(
      normalizedTitle,
      'Stores suggestion comments for the Samverkansportalen web part.',
      COMMENT_COLUMN_DEFINITIONS
    );
    await this._ensureColumns(created.id, COMMENT_COLUMN_DEFINITIONS);
    return { id: created.id, created: true };
  }

  public async getSuggestionItems(
    listId: string,
    options: {
      status?: 'Active' | 'Done';
      top?: number;
      skipToken?: string;
      category?: SuggestionCategory;
      subcategory?: string;
      searchQuery?: string;
      orderBy?: string;
    } = {}
  ): Promise<{ items: IGraphSuggestionItem[]; nextToken?: string }> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    let request = client
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .version('v1.0')
      .select('createdBy,createdDateTime,lastModifiedDateTime')
      .expand('fields($select=Id,Title,Details,Status,Category,Subcategory,Votes,CompletedDateTime)')
      .expand('createdByUser($select=userPrincipalName,mail,email)');

    const filterParts: string[] = [];

    if (options.status) {
      const normalizedStatus: string = options.status === 'Done' ? 'Done' : 'Active';
      filterParts.push(`fields/Status eq '${normalizedStatus}'`);
    }

    if (options.category) {
      const escapedCategory: string = this._escapeFilterValue(options.category);
      filterParts.push(`fields/Category eq '${escapedCategory}'`);
    }

    if (options.subcategory) {
      const escapedSubcategory: string = this._escapeFilterValue(options.subcategory);
      filterParts.push(`fields/Subcategory eq '${escapedSubcategory}'`);
    }

    if (options.searchQuery) {
      const trimmedQuery: string = options.searchQuery.trim();

      if (trimmedQuery.length > 0) {
        const escapedQuery: string = this._escapeFilterValue(trimmedQuery);
        filterParts.push(
          `(contains(fields/Title,'${escapedQuery}') or contains(fields/Details,'${escapedQuery}'))`
        );
      }
    }

    if (filterParts.length > 0) {
      request = request.filter(filterParts.join(' and '));
    }

    if (options.orderBy) {
      request = request.orderby(options.orderBy);
    } else {
      request = request.orderby('createdDateTime desc');
    }

    if (options.top && Number.isFinite(options.top)) {
      request = request.top(options.top);
    }

    if (options.skipToken) {
      request = request.query({ $skiptoken: options.skipToken });
    }

    const response: { value?: IGraphListItemApiModel[]; '@odata.nextLink'?: unknown } = await request.get();

    const items: IGraphListItemApiModel[] = Array.isArray(response.value) ? response.value : [];

    const mappedItems: IGraphSuggestionItem[] = items
      .map((entry) => {
        const fields: unknown = entry.fields;

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
        const createdDateTime: unknown = entry.createdDateTime;
        const lastModifiedDateTime: unknown = entry.lastModifiedDateTime;

        return {
          fields: fields as IGraphSuggestionItemFields,
          createdByUserPrincipalName,
          createdDateTime: typeof createdDateTime === 'string' ? createdDateTime : undefined,
          lastModifiedDateTime: typeof lastModifiedDateTime === 'string' ? lastModifiedDateTime : undefined
        } as IGraphSuggestionItem;
      })
      .filter((item): item is IGraphSuggestionItem => !!item);

    const nextToken: string | undefined = this._extractSkipToken(response['@odata.nextLink']);

    return { items: mappedItems, nextToken };
  }

  public async getVoteItems(
    listId: string,
    options: { suggestionIds?: number[]; username?: string } = {}
  ): Promise<IGraphVoteItem[]> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    let request = client
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .version('v1.0')
      .select('id')
      .expand('fields($select=SuggestionId,Username,Votes)');

    const filterParts: string[] = [];

    if (options.username) {
      const normalizedUsername: string = options.username.trim().toLowerCase();
      if (normalizedUsername.length > 0) {
        const escapedUsername: string = this._escapeFilterValue(normalizedUsername);
        filterParts.push(`fields/Username eq '${escapedUsername}'`);
      }
    }

    const suggestionIds: number[] = (options.suggestionIds ?? [])
      .map((id) => this._normalizeIntegerId(id))
      .filter((id): id is number => typeof id === 'number');

    if (suggestionIds.length > 0) {
      const suggestionFilters: string[] = suggestionIds.map((id) => `fields/SuggestionId eq ${id}`);
      filterParts.push(`(${suggestionFilters.join(' or ')})`);
    }

    if (filterParts.length > 0) {
      request = request.filter(filterParts.join(' and '));
    }

    request = request.top(999);

    const response: { value?: IGraphListItemApiModel[] } = await request.get();

    const items: IGraphListItemApiModel[] = Array.isArray(response.value) ? response.value : [];

    return items
      .map((entry) => {
        const parsedEntry = this._extractListItemWithFields<IGraphVoteItemFields>(entry);

        if (!parsedEntry) {
          return undefined;
        }

        return parsedEntry;
      })
      .filter((item): item is IGraphVoteItem => !!item);
  }

  public async getCategoryItems(listId: string): Promise<IGraphCategoryItem[]> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    const response: { value?: IGraphListItemApiModel[] } = await client
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .version('v1.0')
      .select('id')
      .expand('fields($select=Title)')
      .top(999)
      .get();

    const items: IGraphListItemApiModel[] = Array.isArray(response.value) ? response.value : [];

    return items
      .map((entry) => {
        const parsedEntry = this._extractListItemWithFields<IGraphCategoryItemFields>(entry);

        if (!parsedEntry) {
          return undefined;
        }

        return parsedEntry;
      })
      .filter((item): item is IGraphCategoryItem => !!item);
  }

  public async getCommentItems(
    listId: string,
    options: { suggestionIds?: number[] } = {}
  ): Promise<IGraphCommentItem[]> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    let request = client
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .version('v1.0')
      .select('id,createdBy,createdDateTime')
      .expand('fields($select=Id,SuggestionId,Comment,Title)')
      .expand('createdByUser($select=userPrincipalName,mail,email,displayName)');

    const suggestionIds: number[] = (options.suggestionIds ?? [])
      .map((id) => this._normalizeIntegerId(id))
      .filter((id): id is number => typeof id === 'number');

    if (suggestionIds.length > 0) {
      const suggestionFilters: string[] = suggestionIds.map((id) => `fields/SuggestionId eq ${id}`);
      request = request.filter(`(${suggestionFilters.join(' or ')})`);
    }

    request = request.top(999);

    const response: { value?: IGraphListItemApiModel[] } = await request.get();

    const items: IGraphListItemApiModel[] = Array.isArray(response.value) ? response.value : [];

    return items
      .map((entry) => {
        const parsedEntry = this._extractListItemWithFields<IGraphCommentItemFields>(entry);

        if (!parsedEntry) {
          return undefined;
        }

        const { id, fields } = parsedEntry;

        let createdByUserPrincipalName: string | undefined;
        let createdByUserDisplayName: string | undefined;

        const createdByUser: unknown = entry.createdBy?.user;

        if (createdByUser && typeof createdByUser === 'object') {
          const {
            userPrincipalName,
            email,
            mail,
            displayName
          } = createdByUser as {
            userPrincipalName?: unknown;
            email?: unknown;
            mail?: unknown;
            displayName?: unknown;
          };

          if (typeof displayName === 'string' && displayName.trim().length > 0) {
            createdByUserDisplayName = displayName.trim();
          }

          if (typeof userPrincipalName === 'string' && userPrincipalName.trim().length > 0) {
            createdByUserPrincipalName = userPrincipalName.trim();
          } else if (typeof email === 'string' && email.trim().length > 0) {
            createdByUserPrincipalName = email.trim();
          } else if (typeof mail === 'string' && mail.trim().length > 0) {
            createdByUserPrincipalName = mail.trim();
          }
        }

        const createdDateTime: unknown = entry.createdDateTime;

        return {
          id,
          fields,
          createdByUserPrincipalName,
          createdByUserDisplayName,
          createdDateTime: typeof createdDateTime === 'string' ? createdDateTime : undefined
        } as IGraphCommentItem;
      })
      .filter((item): item is IGraphCommentItem => !!item);
  }

  public async getCommentCounts(
    listId: string,
    options: { suggestionIds?: number[] } = {}
  ): Promise<Map<number, number>> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    let request = client
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .version('v1.0')
      .select('id')
      .expand('fields($select=SuggestionId)');

    const suggestionIds: number[] = (options.suggestionIds ?? [])
      .map((id) => this._normalizeIntegerId(id))
      .filter((id): id is number => typeof id === 'number');

    if (suggestionIds.length > 0) {
      const suggestionFilters: string[] = suggestionIds.map((id) => `fields/SuggestionId eq ${id}`);
      request = request.filter(`(${suggestionFilters.join(' or ')})`);
    }

    request = request.top(999);

    const response: { value?: IGraphListItemApiModel[] } = await request.get();
    const items: IGraphListItemApiModel[] = Array.isArray(response.value) ? response.value : [];
    const counts: Map<number, number> = new Map();

    items.forEach((entry) => {
      const fields: unknown = entry.fields;

      if (!fields || typeof fields !== 'object') {
        return;
      }

      const suggestionId: number | undefined = this._normalizeIntegerId(
        (fields as { SuggestionId?: unknown }).SuggestionId
      );

      if (typeof suggestionId !== 'number') {
        return;
      }

      counts.set(suggestionId, (counts.get(suggestionId) ?? 0) + 1);
    });

    return counts;
  }

  public async addCommentItem(listId: string, fields: IGraphCommentItemFields): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    const normalizedSuggestionId: number | undefined = this._normalizeIntegerId(fields.SuggestionId);

    if (typeof normalizedSuggestionId !== 'number') {
      throw new Error('A valid SuggestionId is required when creating a comment.');
    }

    const normalizedFields: IGraphCommentItemFields = {
      ...fields,
      SuggestionId: normalizedSuggestionId
    };

    await client
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .version('v1.0')
      .post({ fields: normalizedFields });
  }

  public async deleteCommentItem(listId: string, itemId: number): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await client
      .api(`/sites/${siteId}/lists/${listId}/items/${itemId}`)
      .version('v1.0')
      .delete();
  }

  public async deleteCommentsForSuggestion(listId: string, suggestionId: number): Promise<void> {
    const commentItems: IGraphCommentItem[] = await this.getCommentItems(listId, {
      suggestionIds: [suggestionId]
    });

    if (commentItems.length === 0) {
      return;
    }

    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await Promise.all(
      commentItems.map(async (comment) => {
        await client
          .api(`/sites/${siteId}/lists/${listId}/items/${comment.id}`)
          .version('v1.0')
          .delete();
      })
    );
  }

  public async getSubcategoryItems(listId: string): Promise<IGraphSubcategoryItem[]> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    const response: { value?: IGraphListItemApiModel[] } = await client
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .version('v1.0')
      .select('id')
      .expand('fields($select=Title,Category)')
      .top(999)
      .get();

    const items: IGraphListItemApiModel[] = Array.isArray(response.value) ? response.value : [];

    return items
      .map((entry) => {
        const parsedEntry = this._extractListItemWithFields<IGraphSubcategoryItemFields>(entry);

        if (!parsedEntry) {
          return undefined;
        }

        return parsedEntry;
      })
      .filter((item): item is IGraphSubcategoryItem => !!item);
  }

  public async addCategoryItem(listId: string, fields: IGraphCategoryItemFields): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await client
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .version('v1.0')
      .post({ fields });
  }

  public async addSubcategoryItem(listId: string, fields: IGraphSubcategoryItemFields): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await client
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .version('v1.0')
      .post({ fields });
  }

  public async deleteSubcategoryItem(listId: string, itemId: number): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await client
      .api(`/sites/${siteId}/lists/${listId}/items/${itemId}`)
      .version('v1.0')
      .delete();
  }

  public async deleteCategoryItem(listId: string, itemId: number): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await client
      .api(`/sites/${siteId}/lists/${listId}/items/${itemId}`)
      .version('v1.0')
      .delete();
  }

  public async addSuggestion(listId: string, fields: IGraphSuggestionItemFields): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await this._executeWithMetadataFallback(fields, async (metadataPayload) => {
      await this._executeWithVoteFallback(metadataPayload, async (payload) => {
        await client
          .api(`/sites/${siteId}/lists/${listId}/items`)
          .version('v1.0')
          .post({ fields: payload });
      });
    });
  }

  public async updateSuggestion(listId: string, itemId: number, fields: Partial<IGraphSuggestionItemFields>): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await this._executeWithMetadataFallback(fields, async (metadataPayload) => {
      await this._executeWithVoteFallback(metadataPayload, async (payload) => {
        await client
          .api(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`)
          .version('v1.0')
          .patch(payload);
      });
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

  public async addVote(listId: string, fields: IGraphVoteItemFields): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await this._executeWithVoteFallback(fields, async (payload) => {
      await client
        .api(`/sites/${siteId}/lists/${listId}/items`)
        .version('v1.0')
        .post({ fields: payload });
    });
  }

  public async deleteVote(listId: string, itemId: number): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await client
      .api(`/sites/${siteId}/lists/${listId}/items/${itemId}`)
      .version('v1.0')
      .delete();
  }

  public async deleteVotesForSuggestion(listId: string, suggestionId: number): Promise<void> {
    const voteItems: IGraphVoteItem[] = await this.getVoteItems(listId, {
      suggestionIds: [suggestionId]
    });
    const matchingVotes: IGraphVoteItem[] = voteItems.filter((entry) => {
      const value: unknown = entry.fields?.SuggestionId;

      if (typeof value === 'number' && Number.isFinite(value)) {
        return value === suggestionId;
      }

      if (typeof value === 'string') {
        const parsed: number = parseInt(value, 10);
        return Number.isFinite(parsed) && parsed === suggestionId;
      }

      return false;
    });

    if (matchingVotes.length === 0) {
      return;
    }

    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    await Promise.all(
      matchingVotes.map(async (vote) => {
        await client
          .api(`/sites/${siteId}/lists/${listId}/items/${vote.id}`)
          .version('v1.0')
          .delete();
      })
    );
  }

  public async getListByTitle(listTitle: string): Promise<IGraphListInfo | undefined> {
    const normalized: string = listTitle.trim();

    if (!normalized) {
      return undefined;
    }

    return this._getListByTitle(normalized);
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

  private async _createListWithColumns(
    listTitle: string,
    description: string,
    definitions: IListColumnDefinition[]
  ): Promise<IGraphListInfo> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    const response: IGraphListApiModel = await client
      .api(`/sites/${siteId}/lists`)
      .version('v1.0')
      .post({
        displayName: listTitle,
        description,
        list: {
          template: 'genericList'
        },
        columns: definitions.map((definition) => definition.createPayload())
      });

    if (typeof response.id !== 'string' || typeof response.displayName !== 'string') {
      throw new Error('Failed to create the list.');
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

  private async _executeWithMetadataFallback<T extends { Category?: SuggestionCategory; Subcategory?: string }>(
    fields: Partial<T>,
    executor: (payload: Partial<T>) => Promise<void>
  ): Promise<void> {
    try {
      await executor(fields);
    } catch (error) {
      const shouldRetryWithoutCategory: boolean = this._shouldRetryWithoutCategory(fields, error);
      const shouldRetryWithoutSubcategory: boolean = this._shouldRetryWithoutSubcategory(fields, error);

      if (!shouldRetryWithoutCategory && !shouldRetryWithoutSubcategory) {
        throw error;
      }

      const fallbackPayload: Partial<T> = { ...fields };

      if (shouldRetryWithoutCategory) {
        delete (fallbackPayload as { Category?: SuggestionCategory }).Category;
      }

      if (shouldRetryWithoutSubcategory) {
        delete (fallbackPayload as { Subcategory?: string }).Subcategory;
      }

      await executor(fallbackPayload);
    }
  }

  private async _executeWithVoteFallback<T extends { Votes?: number | string }>(
    fields: Partial<T>,
    executor: (payload: Partial<T>) => Promise<void>
  ): Promise<void> {
    try {
      await executor(fields);
    } catch (error) {
      if (!this._shouldRetryWithStringVotes(fields, error)) {
        throw error;
      }

      const fallbackPayload: Partial<T> = {
        ...fields,
        Votes: String(fields.Votes)
      };

      await executor(fallbackPayload);
    }
  }

  private _shouldRetryWithStringVotes(fields: Partial<{ Votes?: number | string }>, error: unknown): boolean {
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

  private _shouldRetryWithoutCategory(
    fields: Partial<{ Category?: SuggestionCategory }>,
    error: unknown
  ): boolean {
    if (typeof fields.Category === 'undefined') {
      return false;
    }

    const message: string | undefined = this._extractErrorMessage(error);

    if (!message) {
      return false;
    }

    return message.toLowerCase().includes('field \'category\' is not recognized');
  }

  private _shouldRetryWithoutSubcategory(
    fields: Partial<{ Subcategory?: string }>,
    error: unknown
  ): boolean {
    if (typeof fields.Subcategory === 'undefined') {
      return false;
    }

    const message: string | undefined = this._extractErrorMessage(error);

    if (!message) {
      return false;
    }

    return message.toLowerCase().includes('field \'subcategory\' is not recognized');
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

  private _escapeFilterValue(value: string): string {
    return value.replace(/'/g, "''");
  }

  private _extractSkipToken(nextLink: unknown): string | undefined {
    if (typeof nextLink !== 'string' || nextLink.length === 0) {
      return undefined;
    }

    const skipTokenMatch: RegExpMatchArray | null = nextLink.match(/[?&]\$skiptoken=([^&]+)/i);

    if (!skipTokenMatch || skipTokenMatch.length < 2) {
      return undefined;
    }

    try {
      return decodeURIComponent(skipTokenMatch[1]);
    } catch (error) {
      console.warn('Failed to decode skip token from nextLink.', error);
      return undefined;
    }
  }

  private async _ensureColumns(listId: string, definitions: IListColumnDefinition[]): Promise<void> {
    const client: MSGraphClientV3 = await this._getClient();
    const siteId: string = await this._getSiteId();

    const response: { value?: IGraphColumnApiModel[] } = await client
      .api(`/sites/${siteId}/lists/${listId}/columns`)
      .version('v1.0')
      .select('id,name,indexed')
      .top(999)
      .get();

    const existingColumns: Map<string, IGraphColumnApiModel> = new Map();

    if (Array.isArray(response.value)) {
      for (const rawEntry of response.value) {
        if (!rawEntry || typeof rawEntry !== 'object') {
          continue;
        }

        const entry: IGraphColumnApiModel = rawEntry as IGraphColumnApiModel;
        const name: unknown = (entry as { name?: unknown }).name;

        if (typeof name !== 'string' || name.length === 0) {
          continue;
        }

        existingColumns.set(name, entry);
      }
    }

    for (const definition of definitions) {
      const existing = existingColumns.get(definition.name);

      if (!existing) {
        await client
          .api(`/sites/${siteId}/lists/${listId}/columns`)
          .version('v1.0')
          .post(definition.createPayload());
        continue;
      }

      if (definition.shouldBeIndexed) {
        const isIndexed: boolean = existing.indexed === true;
        const columnId: unknown = existing.id;

        if (!isIndexed && typeof columnId === 'string' && columnId.length > 0) {
          await client
            .api(`/sites/${siteId}/lists/${listId}/columns/${columnId}`)
            .version('v1.0')
            .patch({ indexed: true });
        }
      }
    }
  }

  private _extractListItemWithFields<TFields extends Record<string, unknown>>(
    entry: IGraphListItemApiModel
  ): { id: number; fields: TFields } | undefined {
    const id: number | undefined = this._normalizeIntegerId(entry.id) ?? -1;
    const fields: unknown = entry.fields;

    if (!fields && typeof fields !== 'object') {
      return undefined;
    }

    return {
      id,
      fields: fields as TFields
    };
  }

  private _normalizeIntegerId(value: unknown): number | undefined {
    if (typeof value === 'number' && Number.isFinite(value)) {
      return Math.trunc(value);
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

export default GraphSuggestionsService;
