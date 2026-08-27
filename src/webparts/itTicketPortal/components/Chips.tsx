// =====================================================================
// components/Chips.tsx
// Two tiny presentational components used in several places.
// A "presentational" component takes props and returns markup -
// no data loading, no state.
// =====================================================================

import * as React from 'react';
import styles from './TicketPortal.module.scss';
import { TicketPriority, TicketStatus } from '../models/ITicket';

const priorityClass: Record<TicketPriority, string> = {
  'Low': styles.priorityLow,
  'Medium': styles.priorityMedium,
  'High': styles.priorityHigh,
  'Critical': styles.priorityCritical
};

const statusClass: Record<TicketStatus, string> = {
  'Open': styles.statusOpen,
  'In Progress': styles.statusInProgress,
  'Resolved': styles.statusResolved,
  'Closed': styles.statusClosed
};

export const PriorityChip: React.FC<{ value: TicketPriority }> = ({ value }) => (
  <span className={`${styles.chip} ${priorityClass[value] || ''}`}>{value}</span>
);

export const StatusChip: React.FC<{ value: TicketStatus }> = ({ value }) => (
  <span className={`${styles.chip} ${statusClass[value] || ''}`}>{value}</span>
);

/** Dates come back from SharePoint as ISO strings; show them readably. */
export const formatDate = (iso?: string): string => {
  if (!iso) { return '-'; }
  const d = new Date(iso);
  return d.toLocaleString(undefined, {
    day: '2-digit', month: 'short', year: 'numeric',
    hour: '2-digit', minute: '2-digit'
  });
};
