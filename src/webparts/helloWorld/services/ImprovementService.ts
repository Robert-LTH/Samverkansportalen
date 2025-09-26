import { SPFI } from '@pnp/sp';
import { IList } from '@pnp/sp/lists';
import { IItemAddResult, IItemUpdateResult } from '@pnp/sp/items';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';

export type SuggestionStatus = 'Föreslagen' | 'Pågående' | 'Genomförd' | 'Avslutad';

export interface ISuggestionItem {
  id: number;
  title: string;
  details: string;
  status: SuggestionStatus;
  created: string;
  author: string;
  voteCount: number;
  userHasVoted: boolean;
}

export interface ISuggestionQueryResult {
  items: ISuggestionItem[];
  remainingVotes: number;
}

export interface ICurrentUserContext {
  loginName: string;
  displayName: string;
}

const SUGGESTION_LIST_TITLE = 'ImprovementSuggestions';
const VOTE_LIST_TITLE = 'SuggestionVotes';
const MAX_ACTIVE_VOTES = 5;

const SUGGESTION_FIELDS_XML = {
  details: `<Field Type="Note" DisplayName="Beskrivning" Name="Details" NumLines="6" />`,
  status: `<Field Type="Choice" DisplayName="Status" Name="Status" Format="Dropdown"><Default>Föreslagen</Default><CHOICES><CHOICE>Föreslagen</CHOICE><CHOICE>Pågående</CHOICE><CHOICE>Genomförd</CHOICE><CHOICE>Avslutad</CHOICE></CHOICES></Field>`
};

const VOTE_FIELDS_XML = {
  suggestionId: `<Field Type="Number" DisplayName="Förslags-ID" Name="SuggestionId" />`,
  voterLogin: `<Field Type="Text" DisplayName="Användarnamn" Name="VoterLoginName" />`
};

const COMPLETED_STATUSES: SuggestionStatus[] = ['Genomförd', 'Avslutad'];

const escapeForFilter = (value: string): string => value.replace(/'/g, "''");

export default class ImprovementService {
  public constructor(private readonly sp: SPFI) {}

  public async ensureInfrastructure(): Promise<void> {
    const suggestionListEnsure = await this.sp.web.lists.ensure(
      SUGGESTION_LIST_TITLE,
      'Förbättringsförslag',
      100,
      true
    );

    await this.ensureField(suggestionListEnsure.list, 'Details', SUGGESTION_FIELDS_XML.details);
    await this.ensureField(suggestionListEnsure.list, 'Status', SUGGESTION_FIELDS_XML.status);

    const votesListEnsure = await this.sp.web.lists.ensure(
      VOTE_LIST_TITLE,
      'Röster kopplade till förbättringsförslag',
      100,
      true
    );

    await this.ensureField(votesListEnsure.list, 'SuggestionId', VOTE_FIELDS_XML.suggestionId);
    await this.ensureField(votesListEnsure.list, 'VoterLoginName', VOTE_FIELDS_XML.voterLogin);
  }

  public async loadSuggestions(currentUser: ICurrentUserContext): Promise<ISuggestionQueryResult> {
    const [suggestionItems, voteItems] = await Promise.all([
      this.sp.web.lists
        .getByTitle(SUGGESTION_LIST_TITLE)
        .items.select('Id', 'Title', 'Details', 'Status', 'Author/Title', 'Author/LoginName', 'Created')
        .expand('Author')
        .orderBy('Created', false)
        .top(5000)()
        .catch(() => []),
      this.sp.web.lists
        .getByTitle(VOTE_LIST_TITLE)
        .items.select('Id', 'Title', 'SuggestionId', 'VoterLoginName')
        .top(5000)()
        .catch(() => [])
    ]);

    const votesBySuggestion = new Map<number, { id: number; voter: string }[]>();

    for (const vote of voteItems) {
      const suggestionId = vote.SuggestionId as number | undefined;
      if (!suggestionId) {
        continue;
      }
      const currentVotes = votesBySuggestion.get(suggestionId) || [];
      currentVotes.push({ id: vote.Id as number, voter: vote.VoterLoginName as string });
      votesBySuggestion.set(suggestionId, currentVotes);
    }

    const items: ISuggestionItem[] = suggestionItems.map((suggestion) => {
      const suggestionId = suggestion.Id as number;
      const votes = votesBySuggestion.get(suggestionId) || [];
      return {
        id: suggestionId,
        title: (suggestion.Title as string) ?? '',
        details: (suggestion.Details as string) ?? '',
        status: ((suggestion.Status as SuggestionStatus) ?? 'Föreslagen'),
        created: suggestion.Created as string,
        author: (suggestion.Author?.Title as string) ?? 'Okänd',
        voteCount: votes.length,
        userHasVoted: votes.some((vote) => vote.voter === currentUser.loginName)
      };
    });

    const userVotes = voteItems.filter((vote) => vote.VoterLoginName === currentUser.loginName).length;

    return {
      items,
      remainingVotes: Math.max(MAX_ACTIVE_VOTES - userVotes, 0)
    };
  }

  public async createSuggestion(title: string, details: string): Promise<IItemAddResult> {
    return this.sp.web.lists.getByTitle(SUGGESTION_LIST_TITLE).items.add({
      Title: title,
      Details: details,
      Status: 'Föreslagen'
    });
  }

  public async updateStatus(suggestionId: number, status: SuggestionStatus): Promise<IItemUpdateResult> {
    const list = this.sp.web.lists.getByTitle(SUGGESTION_LIST_TITLE);
    const result = await list.items.getById(suggestionId).update({ Status: status });

    if (COMPLETED_STATUSES.includes(status)) {
      await this.removeAllVotesForSuggestion(suggestionId);
    }

    return result;
  }

  public async castVote(suggestionId: number, user: ICurrentUserContext): Promise<IItemAddResult | void> {
    const votesList = this.sp.web.lists.getByTitle(VOTE_LIST_TITLE);
    const existingVote = await votesList.items
      .filter(`SuggestionId eq ${suggestionId} and VoterLoginName eq '${escapeForFilter(user.loginName)}'`)
      .top(1)()
      .catch(() => []);

    if (existingVote.length > 0) {
      return;
    }

    return votesList.items.add({
      Title: user.displayName,
      SuggestionId: suggestionId,
      VoterLoginName: user.loginName
    });
  }

  public async removeVote(suggestionId: number, user: ICurrentUserContext): Promise<void> {
    const votesList = this.sp.web.lists.getByTitle(VOTE_LIST_TITLE);
    const voteItems = await votesList.items
      .filter(`SuggestionId eq ${suggestionId} and VoterLoginName eq '${escapeForFilter(user.loginName)}'`)
      .top(1)()
      .catch(() => []);

    if (voteItems.length > 0) {
      await votesList.items.getById(voteItems[0].Id as number).delete();
    }
  }

  private async removeAllVotesForSuggestion(suggestionId: number): Promise<void> {
    const votesList = this.sp.web.lists.getByTitle(VOTE_LIST_TITLE);
    const votes = await votesList.items
      .filter(`SuggestionId eq ${suggestionId}`)
      .top(5000)()
      .catch(() => []);

    await Promise.all(
      votes.map((vote) => votesList.items.getById(vote.Id as number).delete())
    );
  }

  private async ensureField(list: IList, internalName: string, fieldXml: string): Promise<void> {
    try {
      await list.fields.getByInternalNameOrTitle(internalName)();
    } catch (error) {
      await list.fields.createFieldAsXml(fieldXml);
    }
  }
}
