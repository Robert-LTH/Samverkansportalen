import { SPFI } from '@pnp/sp';
import '@pnp/sp/items';
import '@pnp/sp/lists';
import '@pnp/sp/webs';
import '@pnp/sp/site-users/web';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import { ISuggestionWithVotes, SuggestionStatus, IUserInfo, activeStatuses } from '../models/ImprovementModels';

const SUGGESTION_LIST_TITLE = 'ImprovementSuggestions';
const VOTE_LIST_TITLE = 'SuggestionVotes';

export default class SuggestionService {
  private readonly _maxSuggestions = 500;

  public constructor(private readonly sp: SPFI) {}

  public async ensureSetup(): Promise<void> {
    const suggestionListId = await this.ensureSuggestionList();
    await this.ensureVotesList(suggestionListId);
  }

  public async getCurrentUser(): Promise<IUserInfo> {
    const user = await this.sp.web.currentUser.select('Id', 'Title', 'Email')();
    return {
      id: user.Id,
      title: user.Title,
      email: user.Email
    };
  }

  public async getSuggestions(searchText: string | undefined, currentUserId: number): Promise<ISuggestionWithVotes[]> {
    const suggestionList = this.sp.web.lists.getByTitle(SUGGESTION_LIST_TITLE);

    let request = suggestionList.items
      .select('Id', 'Title', 'SuggestionDescription', 'SuggestionStatus', 'Created', 'Author/Id', 'Author/Title', 'Author/EMail')
      .expand('Author')
      .orderBy('Created', false)
      .top(this._maxSuggestions);

    if (searchText && searchText.trim().length > 0) {
      const escaped = this.escapeForOData(searchText.trim());
      request = request.filter(`substringof('${escaped}',Title) or substringof('${escaped}',SuggestionDescription)`);
    }

    const suggestionItems = await request();

    if (!suggestionItems || suggestionItems.length === 0) {
      return [];
    }

    const suggestionIds = suggestionItems.map((item) => item.Id);
    const filters = suggestionIds.map((id) => `(SuggestionId eq ${id})`).join(' or ');
    const voteFilter = filters.length > 0 ? `IsWithdrawn eq 0 and (${filters})` : 'IsWithdrawn eq 0';

    const voteItems = await this.sp.web.lists
      .getByTitle(VOTE_LIST_TITLE)
      .items.select('Id', 'SuggestionId', 'Voter/Id', 'Voter/Title')
      .expand('Voter')
      .filter(voteFilter)
      .top(5000)();

    const votesBySuggestion = new Map<number, { id: number; voterId: number; voterName: string }[]>();
    for (const vote of voteItems) {
      const suggestionId: number = vote.SuggestionId;
      let entries = votesBySuggestion.get(suggestionId);
      if (!entries) {
        entries = [];
        votesBySuggestion.set(suggestionId, entries);
      }
      entries.push({
        id: vote.Id,
        voterId: vote.Voter ? vote.Voter.Id : undefined,
        voterName: vote.Voter ? vote.Voter.Title : ''
      });
    }

    return suggestionItems.map((item) => {
      const status = (item.SuggestionStatus || 'Proposed') as SuggestionStatus;
      const votes = votesBySuggestion.get(item.Id) || [];
      const userVote = votes.find((vote) => vote.voterId === currentUserId);
      const isActive = activeStatuses.indexOf(status) >= 0;

      return {
        id: item.Id,
        title: item.Title,
        description: item.SuggestionDescription || '',
        status,
        created: item.Created,
        createdBy: {
          id: item.Author ? item.Author.Id : undefined,
          title: item.Author ? item.Author.Title : 'Okänd',
          email: item.Author ? item.Author.EMail : undefined
        },
        totalVotes: votes.length,
        activeVotes: isActive ? votes.length : 0,
        userHasActiveVote: isActive && !!userVote,
        userVoteId: userVote ? userVote.id : undefined,
        userHasAnyVote: !!userVote
      };
    });
  }

  public async createSuggestion(title: string, description: string): Promise<void> {
    await this.sp.web.lists.getByTitle(SUGGESTION_LIST_TITLE).items.add({
      Title: title,
      SuggestionDescription: description,
      SuggestionStatus: 'Proposed'
    });
  }

  public async updateSuggestionStatus(id: number, status: SuggestionStatus): Promise<void> {
    await this.sp.web.lists.getByTitle(SUGGESTION_LIST_TITLE).items.getById(id).update({
      SuggestionStatus: status
    });
  }

  public async addVote(suggestionId: number, userId: number): Promise<void> {
    const existing = await this.sp.web.lists
      .getByTitle(VOTE_LIST_TITLE)
      .items.select('Id')
      .filter(`SuggestionId eq ${suggestionId} and VoterId eq ${userId} and IsWithdrawn eq 0`)
      .top(1)();

    if (existing.length > 0) {
      return;
    }

    await this.sp.web.lists.getByTitle(VOTE_LIST_TITLE).items.add({
      Title: `Vote-${suggestionId}-${userId}`,
      SuggestionId: suggestionId,
      VoterId: userId,
      IsWithdrawn: false
    });
  }

  public async withdrawVote(voteId: number): Promise<void> {
    await this.sp.web.lists
      .getByTitle(VOTE_LIST_TITLE)
      .items.getById(voteId)
      .update({ IsWithdrawn: true });
  }

  private async ensureSuggestionList(): Promise<string> {
    const ensureResult = await this.sp.web.lists.ensure(SUGGESTION_LIST_TITLE, 'Förbättringsförslag', 100);
    const suggestionList = ensureResult.list;

    if (ensureResult.created) {
      await suggestionList.fields.createFieldAsXml(
        '<Field DisplayName="Beskrivning" Type="Note" Name="SuggestionDescription" NumLines="6" RichText="FALSE" />'
      );
      await suggestionList.fields.createFieldAsXml(
        '<Field DisplayName="Status" Type="Choice" Name="SuggestionStatus" Format="Dropdown">' +
          '<CHOICES>' +
          '<CHOICE>Proposed</CHOICE>' +
          '<CHOICE>InProgress</CHOICE>' +
          '<CHOICE>Completed</CHOICE>' +
          '<CHOICE>Removed</CHOICE>' +
          '</CHOICES>' +
          '<Default>Proposed</Default>' +
          '</Field>'
      );
    }

    const info = await suggestionList.select('Id')();
    return info.Id;
  }

  private async ensureVotesList(suggestionListId: string): Promise<void> {
    const ensureResult = await this.sp.web.lists.ensure(VOTE_LIST_TITLE, 'Röster på förbättringsförslag', 100);
    const voteList = ensureResult.list;

    if (ensureResult.created) {
      await voteList.fields.createFieldAsXml(
        `<Field DisplayName="Förslag" Type="Lookup" Required="TRUE" Name="Suggestion" List="{${suggestionListId}}" ShowField="Title" />`
      );
      await voteList.fields.createFieldAsXml(
        '<Field DisplayName="Röstare" Type="User" Name="Voter" UserSelectionMode="PeopleOnly" UserSelectionScope="0" />'
      );
      await voteList.fields.createFieldAsXml(
        '<Field DisplayName="Återtagen" Type="Boolean" Name="IsWithdrawn" Default="0" />'
      );
    }
  }

  private escapeForOData(value: string): string {
    return value.replace(/'/g, "''");
  }
}
