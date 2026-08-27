// =====================================================================
// components/Dashboard.tsx
// Counts are derived from the tickets already loaded, so no extra
// SharePoint calls are needed. Each card is a button: clicking it
// filters the list below by that status.
// =====================================================================

import * as React from 'react';
import styles from './TicketPortal.module.scss';
import { ITicket, TicketStatus } from '../models/ITicket';

export interface IDashboardProps {
  tickets: ITicket[];
  activeStatus: TicketStatus | '';
  onSelectStatus: (status: TicketStatus | '') => void;
}

const Dashboard: React.FC<IDashboardProps> = ({ tickets, activeStatus, onSelectStatus }) => {

  const countBy = (status: TicketStatus): number =>
    tickets.filter(t => t.TicketStatus === status).length;

  const cards = [
    { key: '' as const, label: 'Total tickets', value: tickets.length, cls: styles.cardTotal },
    { key: 'Open' as const, label: 'Open', value: countBy('Open'), cls: styles.cardOpen },
    { key: 'In Progress' as const, label: 'In progress', value: countBy('In Progress'), cls: styles.cardProgress },
    { key: 'Resolved' as const, label: 'Resolved', value: countBy('Resolved'), cls: styles.cardResolved },
    { key: 'Closed' as const, label: 'Closed', value: countBy('Closed'), cls: styles.cardClosed }
  ];

  return (
    <div className={styles.cards}>
      {cards.map(card => (
        <button
          key={card.label}
          type="button"
          className={`${styles.card} ${card.cls} ${activeStatus === card.key ? styles.cardActive : ''}`}
          onClick={() => onSelectStatus(card.key)}
          aria-pressed={activeStatus === card.key}
        >
          <span className={styles.cardValue}>{card.value}</span>
          <span className={styles.cardLabel}>{card.label}</span>
        </button>
      ))}
    </div>
  );
};

export default Dashboard;
