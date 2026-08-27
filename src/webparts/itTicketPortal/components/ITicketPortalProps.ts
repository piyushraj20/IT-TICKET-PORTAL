// =====================================================================
// components/ITicketPortalProps.ts
// The contract between the web part (SharePoint world) and React.
// Everything React needs is handed over as props - the component never
// reaches out to SharePoint globals on its own.
// =====================================================================

import { SPFI } from '@pnp/sp';

export interface ITicketPortalProps {
  sp: SPFI;
  siteUrl: string;
  ticketsList: string;
  commentsList: string;
  historyList: string;
  supportGroupName: string;
  userDisplayName: string;
  userEmail: string;
}
