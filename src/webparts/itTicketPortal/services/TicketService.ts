// =====================================================================
// services/TicketService.ts
// EVERY call to SharePoint happens in this one file.
// Components never talk to SharePoint directly - they call this service.
// That separation is what keeps the UI code readable and testable.
// =====================================================================

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/attachments';
import '@pnp/sp/site-users/web';
import '@pnp/sp/site-groups/web';
import {
  ITicket, INewTicket, IComment, IHistoryEntry, IAttachment,
  ITicketFilter, ICurrentUser, TicketStatus
} from '../models/ITicket';

/** Escapes single quotes so a search for O'Brien doesn't break the query. */
const esc = (value: string): string => (value || '').replace(/'/g, "''");
const formatTicketId = (id: number): string => `INC-${('00000' + id).slice(-5)}`;
const STATUS_FIELD = 'Status';

/** Fields we want back from the Tickets list. Slash = a field on a linked user. */
const TICKET_SELECT: string[] = [
  'Id', 'TicketID', 'Title', 'EmployeeName', 'Email', 'Category', 'Description',
  'Priority', STATUS_FIELD, 'Created', 'Modified',
  'AssignedTo/Title', 'AssignedTo/Id'
];
const TICKET_EXPAND: string[] = ['AssignedTo'];

export class TicketService {

  constructor(
    private sp: SPFI,
    private ticketsList: string,
    private commentsList: string,
    private historyList: string
  ) { }

  // -------------------------------------------------------------------
  // Mapping: turn a raw SharePoint item into our tidy ITicket object.
  // -------------------------------------------------------------------
  /* eslint-disable @typescript-eslint/no-explicit-any */
  private toTicket(raw: any): ITicket {
    return {
      Id: raw.Id,
      TicketID: raw.TicketID || formatTicketId(raw.Id),
      Title: raw.Title || '',
      EmployeeName: raw.EmployeeName || '',
      Email: raw.Email || '',
      Category: raw.Category || '',
      Description: raw.Description || '',
      Priority: raw.Priority || 'Low',
      TicketStatus: raw[STATUS_FIELD] || 'Open',
      AssignedToName: raw.AssignedTo ? raw.AssignedTo.Title : '',
      AssignedToId: raw.AssignedTo ? raw.AssignedTo.Id : undefined,
      Created: raw.Created,
      Modified: raw.Modified
    };
  }

  // -------------------------------------------------------------------
  // READ
  // -------------------------------------------------------------------

  /**
   * Status and Priority are filtered by SharePoint (fast, scales).
   * The free-text search is done in the browser on the returned rows,
   * which is fine up to a few hundred tickets. See the guide for how to
   * push search to the server once the list gets big.
   */
  public async getTickets(filter?: ITicketFilter): Promise<ITicket[]> {
    const conditions: string[] = [];
    if (filter?.status) { conditions.push(`${STATUS_FIELD} eq '${esc(filter.status)}'`); }
    if (filter?.priority) { conditions.push(`Priority eq '${esc(filter.priority)}'`); }

    const query = this.sp.web.lists.getByTitle(this.ticketsList).items
      .select(...TICKET_SELECT)
      .expand(...TICKET_EXPAND)
      .orderBy('Created', false)   // false = newest first
      .top(500);

    const raw: any[] = conditions.length
      ? await query.filter(conditions.join(' and '))()
      : await query();

    let tickets = raw.map(r => this.toTicket(r));

    const term = (filter?.search || '').trim().toLowerCase();
    if (term) {
      tickets = tickets.filter(t =>
        t.TicketID.toLowerCase().indexOf(term) >= 0 ||
        t.Title.toLowerCase().indexOf(term) >= 0
      );
    }
    return tickets;
  }

  public async getTicketById(id: number): Promise<ITicket> {
    const raw: any = await this.sp.web.lists.getByTitle(this.ticketsList).items
      .getById(id)
      .select(...TICKET_SELECT)
      .expand(...TICKET_EXPAND)();
    return this.toTicket(raw);
  }

  public async getComments(ticketItemId: number): Promise<IComment[]> {
    try {
      const raw: any[] = await this.sp.web.lists.getByTitle(this.commentsList).items
        .select('Id', 'TicketItemId', 'CommentText', 'Created', 'Author/Title')
        .expand('Author')
        .filter(`TicketItemId eq ${ticketItemId}`)
        .orderBy('Created', true)
        .top(200)();

      return raw.map(r => ({
        Id: r.Id,
        TicketItemId: r.TicketItemId,
        CommentText: r.CommentText || '',
        AuthorName: r.Author ? r.Author.Title : '',
        Created: r.Created
      }));
    } catch {
      return [];
    }
  }

  public async getHistory(ticketItemId: number): Promise<IHistoryEntry[]> {
    try {
      const raw: any[] = await this.sp.web.lists.getByTitle(this.historyList).items
        .select('Id', 'TicketItemId', 'FromStatus', 'ToStatus', 'Note', 'Created', 'Author/Title')
        .expand('Author')
        .filter(`TicketItemId eq ${ticketItemId}`)
        .orderBy('Created', true)
        .top(100)();

      return raw.map(r => ({
        Id: r.Id,
        TicketItemId: r.TicketItemId,
        FromStatus: r.FromStatus || '',
        ToStatus: r.ToStatus || '',
        Note: r.Note || '',
        AuthorName: r.Author ? r.Author.Title : '',
        Created: r.Created
      }));
    } catch {
      return [];
    }
  }

  public async getAttachments(ticketItemId: number): Promise<IAttachment[]> {
    const files: any[] = await this.sp.web.lists.getByTitle(this.ticketsList).items
      .getById(ticketItemId).attachmentFiles();
    return files.map(f => ({ FileName: f.FileName, ServerRelativeUrl: f.ServerRelativeUrl }));
  }

  // -------------------------------------------------------------------
  // WRITE
  // -------------------------------------------------------------------

  /**
   * Creates the ticket, then stamps a readable TicketID built from the
   * item id SharePoint just generated. Doing it in two steps guarantees
   * the number is unique even if ten people submit at the same second.
   */
  public async createTicket(data: INewTicket, files: File[]): Promise<ITicket> {
    const list = this.sp.web.lists.getByTitle(this.ticketsList);

    const added: any = await list.items.add({
      Title: data.Title,
      EmployeeName: data.EmployeeName,
      Email: data.Email,
      Category: data.Category,
      Description: data.Description,
      Priority: data.Priority,
      [STATUS_FIELD]: 'Open'
    });

    // PnPjs v3 returned { data: {...} }, v4 returns the item directly.
    // This line works with either.
    const newId: number = added?.Id ?? added?.data?.Id;

    const ticketId = formatTicketId(newId);
    await list.items.getById(newId).update({ TicketID: ticketId });

    for (const file of files || []) {
      await list.items.getById(newId).attachmentFiles.add(file.name, file);
    }

    await this.addHistory(newId, '', 'Open', 'Ticket raised');
    return this.getTicketById(newId);
  }

  public async updateStatus(
    ticketItemId: number,
    from: TicketStatus,
    to: TicketStatus,
    note: string
  ): Promise<void> {
    const changes: Record<string, unknown> = { [STATUS_FIELD]: to };

    await this.sp.web.lists.getByTitle(this.ticketsList).items
      .getById(ticketItemId).update(changes);

    await this.addHistory(ticketItemId, from, to, note);
  }

  /** Simple assignment model: an agent claims the ticket for themselves. */
  public async assignTo(ticketItemId: number, userId: number): Promise<void> {
    await this.sp.web.lists.getByTitle(this.ticketsList).items
      .getById(ticketItemId).update({ AssignedToId: userId });
  }

  public async addComment(ticketItemId: number, text: string): Promise<void> {
    await this.sp.web.lists.getByTitle(this.commentsList).items.add({
      Title: `Comment on ${ticketItemId}`,
      TicketItemId: ticketItemId,
      CommentText: text
    });
  }

  private async addHistory(
    ticketItemId: number, from: string, to: string, note: string
  ): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.historyList).items.add({
        Title: `${from || 'New'} -> ${to}`,
        TicketItemId: ticketItemId,
        FromStatus: from,
        ToStatus: to,
        Note: note
      });
    } catch {
      // History is optional; ticket operations must still succeed without it.
    }
  }

  // -------------------------------------------------------------------
  // WHO AM I
  // -------------------------------------------------------------------

  public async getCurrentUser(): Promise<ICurrentUser> {
    const user: any = await this.sp.web.currentUser();
    return { Id: user.Id, Title: user.Title, Email: user.Email };
  }

  /** True when the signed-in user belongs to the SharePoint group given. */
  public async isInGroup(groupName: string): Promise<boolean> {
    if (!groupName) { return false; }
    try {
      const groups: any[] = await this.sp.web.currentUser.groups();
      return groups.some(g => (g.Title || '').toLowerCase() === groupName.toLowerCase());
    } catch {
      return false;   // site owners sometimes can't read group membership
    }
  }
}
