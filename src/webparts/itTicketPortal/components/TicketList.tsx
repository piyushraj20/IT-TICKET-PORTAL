// =====================================================================
// components/TicketList.tsx
// Search + filter bar and the results table.
// This component owns no data: the parent passes tickets in and gets
// user intent back through callbacks.
// =====================================================================

import * as React from 'react';
import styles from './TicketPortal.module.scss';
import { PriorityChip, StatusChip, formatDate } from './Chips';
import {
  ITicket, ITicketFilter, PRIORITIES, STATUSES,
  TicketPriority, TicketStatus
} from '../models/ITicket';

export interface ITicketListProps {
  tickets: ITicket[];
  filter: ITicketFilter;
  loading: boolean;
  onFilterChange: (filter: ITicketFilter) => void;
  onOpenTicket: (id: number) => void;
}

const TicketList: React.FC<ITicketListProps> = (props) => {
  const { tickets, filter, loading, onFilterChange, onOpenTicket } = props;

  // Spread the old filter, override one key. Never mutate state directly.
  const update = (patch: Partial<ITicketFilter>): void =>
    onFilterChange({ ...filter, ...patch });

  return (
    <div>
      <div className={styles.filters}>
        <div className={`${styles.field} ${styles.fieldGrow}`}>
          <label className={styles.label} htmlFor="tp-search">Search</label>
          <input
            id="tp-search"
            className={styles.input}
            placeholder="Ticket ID or subject"
            value={filter.search}
            onChange={e => update({ search: e.target.value })}
          />
        </div>

        <div className={styles.field}>
          <label className={styles.label} htmlFor="tp-fstatus">Status</label>
          <select
            id="tp-fstatus"
            className={styles.select}
            value={filter.status}
            onChange={e => update({ status: e.target.value as TicketStatus | '' })}
          >
            <option value="">All statuses</option>
            {STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>

        <div className={styles.field}>
          <label className={styles.label} htmlFor="tp-fpriority">Priority</label>
          <select
            id="tp-fpriority"
            className={styles.select}
            value={filter.priority}
            onChange={e => update({ priority: e.target.value as TicketPriority | '' })}
          >
            <option value="">All priorities</option>
            {PRIORITIES.map(p => <option key={p} value={p}>{p}</option>)}
          </select>
        </div>

        <button
          type="button"
          className={styles.button}
          onClick={() => onFilterChange({ search: '', status: '', priority: '' })}
        >
          Clear
        </button>
      </div>

      {loading && <div className={styles.loading}>Loading tickets...</div>}

      {!loading && tickets.length === 0 && (
        <div className={styles.tableWrap}>
          <div className={styles.empty}>
            No tickets match these filters. Clear the filters, or raise the first ticket.
          </div>
        </div>
      )}

      {!loading && tickets.length > 0 && (
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Ticket ID</th>
                <th>Subject</th>
                <th>Priority</th>
                <th>Status</th>
                <th>Created</th>
                <th>Assigned to</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {tickets.map(t => (
                <tr key={t.Id}>
                  <td className={styles.ticketId}>{t.TicketID}</td>
                  <td className={styles.subjectCell}>
                    {t.Title}
                    <div className={styles.dim}>{t.Category}</div>
                  </td>
                  <td><PriorityChip value={t.Priority} /></td>
                  <td><StatusChip value={t.TicketStatus} /></td>
                  <td className={styles.dim}>{formatDate(t.Created)}</td>
                  <td className={styles.dim}>{t.AssignedToName || 'Unassigned'}</td>
                  <td>
                    <button
                      type="button"
                      className={styles.linkButton}
                      onClick={() => onOpenTicket(t.Id)}
                    >
                      View
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

export default TicketList;
