// =====================================================================
// models/ITicket.ts
// All the "shapes" of data used by the portal live here.
// A TypeScript interface is just a description of an object's fields.
// It disappears at runtime - it only helps you and the compiler.
// =====================================================================

export type TicketStatus = 'Open' | 'In Progress' | 'Resolved' | 'Closed';
export type TicketPriority = 'Low' | 'Medium' | 'High' | 'Critical';

/** One row of the ITTickets SharePoint list, cleaned up for the UI. */
export interface ITicket {
  Id: number;                 // SharePoint's own item id (a number)
  TicketID: string;           // our friendly id, e.g. INC-00042
  Title: string;              // Subject (SharePoint calls this column Title)
  EmployeeName: string;
  Email: string;
  Category: string;
  Description: string;
  Priority: TicketPriority;
  TicketStatus: TicketStatus;
  AssignedToName: string;     // '' when nobody is assigned yet
  AssignedToId?: number;
  Created: string;            // ISO date string
  Modified: string;
}

/** What the Raise Ticket form hands to the service. */
export interface INewTicket {
  Title: string;
  EmployeeName: string;
  Email: string;
  Category: string;
  Description: string;
  Priority: TicketPriority;
}

export interface IComment {
  Id: number;
  TicketItemId: number;
  CommentText: string;
  AuthorName: string;
  Created: string;
}

/** One entry in the status timeline. */
export interface IHistoryEntry {
  Id: number;
  TicketItemId: number;
  FromStatus: string;
  ToStatus: string;
  Note: string;
  AuthorName: string;
  Created: string;
}

export interface IAttachment {
  FileName: string;
  ServerRelativeUrl: string;
}

export interface ITicketFilter {
  search: string;
  status: TicketStatus | '';
  priority: TicketPriority | '';
}

export interface ICurrentUser {
  Id: number;
  Title: string;
  Email: string;
}

// ---------------------------------------------------------------------
// Constants. Keep these EXACTLY in sync with the Choice values you type
// into the SharePoint list columns, or filtering will silently return 0.
// ---------------------------------------------------------------------

export const STATUSES: TicketStatus[] = ['Open', 'In Progress', 'Resolved', 'Closed'];

export const PRIORITIES: TicketPriority[] = ['Low', 'Medium', 'High', 'Critical'];

export const CATEGORIES: string[] = [
  'Hardware',
  'Software',
  'Network',
  'Access / Permissions',
  'Email',
  'Other'
];

/** Which statuses a support agent is allowed to move a ticket to. */
export const NEXT_STATUS: Record<TicketStatus, TicketStatus[]> = {
  'Open': ['In Progress', 'Resolved', 'Closed'],
  'In Progress': ['Resolved', 'Open'],
  'Resolved': ['Closed', 'In Progress'],
  'Closed': ['Open']
};
