// =====================================================================
// components/TicketPortal.tsx
// The root component. It owns the shared state (tickets, filters,
// which screen is showing) and passes slices of it to the children.
// This is the standard "smart parent, dumb children" React pattern.
// =====================================================================

import * as React from 'react';
import styles from './TicketPortal.module.scss';
import { ITicketPortalProps } from './ITicketPortalProps';
import { TicketService } from '../services/TicketService';
import Dashboard from './Dashboard';
import TicketList from './TicketList';
import TicketDetails from './TicketDetails';
import RaiseTicketForm from './RaiseTicketForm';
import {
  ICurrentUser, INewTicket, ITicket, ITicketFilter, TicketStatus
} from '../models/ITicket';

type View = 'home' | 'list' | 'new' | 'details';
type DemoRole = 'user' | 'developer';
const DEVELOPER_PASSWORD = 'admin';
const PRIORITY_ORDER: Record<string, number> = {
  Critical: 0,
  High: 1,
  Medium: 2,
  Low: 3
};

const emptyFilter: ITicketFilter = { search: '', status: '', priority: '' };

const TicketPortal: React.FC<ITicketPortalProps> = (props) => {

  // useMemo keeps the same service instance between renders instead of
  // building a new one every time React redraws.
  const service = React.useMemo(
    () => new TicketService(props.sp, props.ticketsList, props.commentsList, props.historyList),
    [props.sp, props.ticketsList, props.commentsList, props.historyList]
  );

  const [view, setView] = React.useState<View>('home');
  const [allTickets, setAllTickets] = React.useState<ITicket[]>([]);
  const [filter, setFilter] = React.useState<ITicketFilter>(emptyFilter);
  const [mineOnly, setMineOnly] = React.useState(false);
  const [selectedId, setSelectedId] = React.useState<number>(0);
  const [currentUser, setCurrentUser] = React.useState<ICurrentUser>({ Id: 0, Title: '', Email: '' });
  const [demoRole, setDemoRole] = React.useState<DemoRole>('user');
  const [developerUnlocked, setDeveloperUnlocked] = React.useState(false);
  const [showDeveloperPassword, setShowDeveloperPassword] = React.useState(false);
  const [developerPassword, setDeveloperPassword] = React.useState('');
  const [developerPasswordError, setDeveloperPasswordError] = React.useState('');
  const [ticketIdSearch, setTicketIdSearch] = React.useState('');
  const [loading, setLoading] = React.useState(true);
  const [submitting, setSubmitting] = React.useState(false);
  const [error, setError] = React.useState('');
  const [success, setSuccess] = React.useState('');

  // -------------------------------------------------------------------
  // Loading. We pull every ticket once (up to 500) and filter in the
  // browser, so the dashboard counts stay accurate while filters are on.
  // Past a few hundred tickets, switch to service.getTickets(filter),
  // which pushes Status and Priority filtering to SharePoint instead.
  // -------------------------------------------------------------------
  const loadTickets = React.useCallback(async (): Promise<void> => {
    setLoading(true);
    setError('');
    try {
      setAllTickets(await service.getTickets());
    } catch (e) {
      setError(
        'Could not read the ticket list. Check that the list names in the web part ' +
        `properties match your SharePoint lists. Details: ${(e as Error).message}`
      );
    } finally {
      setLoading(false);
    }
  }, [service]);

  React.useEffect(() => {
    const init = async (): Promise<void> => {
      try {
        const user = await service.getCurrentUser();
        setCurrentUser(user);
        setDemoRole('user');
      } catch {
        // Non-fatal: the portal still works, just without agent controls.
      }
      await service.ensureSupportLists();
      await loadTickets();
    };
    init().catch(() => undefined);
  }, [service, loadTickets]);

  const isDeveloperView = demoRole === 'developer' && developerUnlocked;

  // -------------------------------------------------------------------
  // Derived data. useMemo recalculates only when its inputs change.
  // -------------------------------------------------------------------
  const scopedTickets = React.useMemo(
    () => mineOnly ? allTickets.filter(t => t.EmployeeName === currentUser.Title) : allTickets,
    [allTickets, mineOnly, currentUser.Id]
  );

  const visibleTickets = React.useMemo(() => {
    const term = filter.search.trim().toLowerCase();
    const filtered = scopedTickets.filter(t => {
      if (filter.status && t.TicketStatus !== filter.status) { return false; }
      if (filter.priority && t.Priority !== filter.priority) { return false; }
      if (term &&
        t.TicketID.toLowerCase().indexOf(term) < 0 &&
        t.Title.toLowerCase().indexOf(term) < 0) { return false; }
      return true;
    });
    return isDeveloperView
      ? filtered.sort((left, right) => PRIORITY_ORDER[left.Priority] - PRIORITY_ORDER[right.Priority])
      : filtered;
  }, [scopedTickets, filter, isDeveloperView]);

  const selectDemoRole = (role: DemoRole): void => {
    setDemoRole(role);
    if (role === 'user') {
      setDeveloperUnlocked(false);
      setShowDeveloperPassword(false);
      setDeveloperPassword('');
      setDeveloperPasswordError('');
    }
    setMineOnly(false);
    setView(role === 'developer' ? 'list' : 'home');
  };

  const requestDeveloperView = (): void => {
    setShowDeveloperPassword(true);
    setDeveloperPasswordError('');
    setSelectedId(0);
    setSuccess('');
    setView('home');
  };

  const unlockDeveloperView = (event: React.FormEvent<HTMLFormElement>): void => {
    event.preventDefault();
    if (developerPassword !== DEVELOPER_PASSWORD) {
      setDeveloperPasswordError('Incorrect developer password.');
      return;
    }
    setDeveloperUnlocked(true);
    setDemoRole('developer');
    setShowDeveloperPassword(false);
    setDeveloperPassword('');
    setDeveloperPasswordError('');
    setMineOnly(false);
    setView('list');
  };

  // -------------------------------------------------------------------
  // Actions
  // -------------------------------------------------------------------
  const handleCreate = async (data: INewTicket, files: File[]): Promise<void> => {
    setSubmitting(true);
    setError('');
    try {
      const created = await service.createTicket(data, files);
      setSuccess(`Ticket ${created.TicketID} raised. The IT team can see it now.`);
      await loadTickets();
      setSelectedId(created.Id);
      setView('details');
    } catch (e) {
      setError(`Ticket not created. ${(e as Error).message}`);
      throw e;
    } finally {
      setSubmitting(false);
    }
  };

  const openTicket = (id: number): void => {
    setSelectedId(id);
    setSuccess('');
    setView('details');
  };

  const checkTicketStatus = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();
    if (!ticketIdSearch.trim()) {
      setError('Enter a ticket ID, for example INC-00001.');
      return;
    }
    setSubmitting(true);
    setError('');
    try {
      const ticket = await service.getTicketByTicketId(ticketIdSearch);
      setSelectedId(ticket.Id);
      setView('details');
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setSubmitting(false);
    }
  };

  const goToList = (): void => {
    setSuccess('');
    setView(isDeveloperView ? 'list' : 'home');
  };

  const handleTicketDeleted = async (): Promise<void> => {
    setSelectedId(0);
    setSuccess('Ticket deleted.');
    setView('list');
    await loadTickets();
  };

  const selectStatusCard = (status: TicketStatus | ''): void => {
    setFilter({ ...filter, status });
    setView('list');
  };

  const exportTickets = (): void => {
    const escapeCsv = (value: string): string => `"${value.replace(/"/g, '""')}"`;
    const headers = ['Ticket ID', 'Subject', 'Requester', 'Email', 'Category', 'Priority', 'Status', 'Assigned to', 'Created'];
    const rows = visibleTickets.map(ticket => [
      ticket.TicketID,
      ticket.Title,
      ticket.EmployeeName,
      ticket.Email,
      ticket.Category,
      ticket.Priority,
      ticket.TicketStatus,
      ticket.AssignedToName || 'Unassigned',
      ticket.Created
    ]);
    const csv = [headers, ...rows]
      .map(row => row.map(value => escapeCsv(value || '')).join(','))
      .join('\r\n');
    const link = document.createElement('a');
    link.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
    link.download = `it-tickets-${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
    setSuccess(`${visibleTickets.length} ticket${visibleTickets.length === 1 ? '' : 's'} exported.`);
  };

  // -------------------------------------------------------------------
  // Render
  // -------------------------------------------------------------------
  return (
    <section className={styles.portal}>

      {!isDeveloperView && <header className={styles.header}>
        <div>
          <h2 className={styles.title}>IT Service Desk</h2>
          <p className={styles.subtitle}>
            {isDeveloperView
              ? 'Developer view: assign tickets and move them through the workflow.'
              : 'Raise a ticket, follow its progress, and reply to the IT team.'}
          </p>
        </div>
        <div className={styles.actions}>
          <button
            type="button"
            className={`${styles.button} ${demoRole === 'user' ? styles.buttonPrimary : ''}`}
            onClick={() => selectDemoRole('user')}
            aria-pressed={demoRole === 'user'}
          >
            View as user
          </button>
          <button
            type="button"
            className={`${styles.button} ${demoRole === 'developer' ? styles.buttonPrimary : ''}`}
            onClick={isDeveloperView ? () => selectDemoRole('user') : requestDeveloperView}
            aria-pressed={demoRole === 'developer'}
          >
            {isDeveloperView ? 'Exit developer view' : 'Developer access'}
          </button>
        </div>
      </header>}

      {isDeveloperView && <header className={styles.header}>
        <div>
          <h2 className={styles.title}>Developer workspace</h2>
          <p className={styles.subtitle}>Manage ticket status and assignment.</p>
        </div>
        <button
          type="button"
          className={styles.button}
          onClick={() => selectDemoRole('user')}
        >
          Exit developer workspace
        </button>
      </header>}

      {!isDeveloperView && showDeveloperPassword && (
        <form className={styles.panel} onSubmit={unlockDeveloperView} style={{ marginTop: 16 }}>
          <div className={styles.field}>
            <label className={styles.label} htmlFor="tp-developer-password">Developer password</label>
            <input
              id="tp-developer-password"
              className={styles.input}
              type="password"
              value={developerPassword}
              onChange={event => setDeveloperPassword(event.target.value)}
              autoComplete="off"
              autoFocus
            />
          </div>
          {developerPasswordError && <div className={styles.errorText}>{developerPasswordError}</div>}
          <div className={styles.actions} style={{ marginTop: 12 }}>
            <button type="submit" className={`${styles.button} ${styles.buttonPrimary}`}>Unlock developer view</button>
            <button type="button" className={styles.button} onClick={() => setShowDeveloperPassword(false)}>Cancel</button>
          </div>
        </form>
      )}

      {!showDeveloperPassword && <nav className={styles.nav}>
        {isDeveloperView && <button
          type="button"
          className={`${styles.navButton} ${view === 'list' ? styles.navButtonActive : ''}`}
          onClick={goToList}
        >
          Dashboard
        </button>}
        {isDeveloperView && (
          <button
            type="button"
            className={`${styles.navButton} ${view === 'new' ? styles.navButtonActive : ''}`}
            onClick={() => setView('new')}
          >
            Raise ticket
          </button>
        )}
        {isDeveloperView && view === 'list' && (
          <button
            type="button"
            className={`${styles.navButton} ${mineOnly ? styles.navButtonActive : ''}`}
            onClick={() => setMineOnly(!mineOnly)}
            aria-pressed={mineOnly}
          >
            {mineOnly ? 'Showing my tickets' : 'Show only my tickets'}
          </button>
        )}
        {isDeveloperView && <button type="button" className={styles.navButton} onClick={loadTickets}>
          Refresh
        </button>}
        {isDeveloperView && <button type="button" className={styles.navButton} onClick={exportTickets}>
          Export CSV
        </button>}
      </nav>}

      {error && <div className={`${styles.banner} ${styles.bannerError}`}>{error}</div>}
      {success && <div className={`${styles.banner} ${styles.bannerSuccess}`}>{success}</div>}

      {!isDeveloperView && !showDeveloperPassword && view === 'home' && (
        <div className={styles.detailGrid}>
          <div className={styles.panel}>
            <h3 className={styles.panelTitle}>Raise a ticket</h3>
            <p className={styles.dim}>Tell the IT team what you need help with.</p>
            <button
              type="button"
              className={`${styles.button} ${styles.buttonPrimary}`}
              onClick={() => setView('new')}
            >
              Raise a ticket
            </button>
          </div>
          <div className={styles.panel}>
            <h3 className={styles.panelTitle}>Check ticket status</h3>
            <p className={styles.dim}>Enter your ticket ID to view its status, comments, and timeline.</p>
            <form onSubmit={checkTicketStatus}>
              <div className={styles.field}>
                <label className={styles.label} htmlFor="tp-ticket-id">Ticket ID</label>
                <input
                  id="tp-ticket-id"
                  className={styles.input}
                  value={ticketIdSearch}
                  onChange={event => setTicketIdSearch(event.target.value)}
                  placeholder="INC-00001"
                />
              </div>
              <button
                type="submit"
                className={`${styles.button} ${styles.buttonPrimary}`}
                disabled={submitting}
              >
                Check status
              </button>
            </form>
          </div>
        </div>
      )}

      {isDeveloperView && view === 'list' && (
        <>
          <Dashboard
            tickets={scopedTickets}
            activeStatus={filter.status}
            onSelectStatus={selectStatusCard}
          />
          <TicketList
            tickets={visibleTickets}
            filter={filter}
            loading={loading}
            onFilterChange={setFilter}
            onOpenTicket={openTicket}
          />
        </>
      )}

      {view === 'new' && (
        <RaiseTicketForm
          defaultName={props.userDisplayName}
          defaultEmail={props.userEmail}
          submitting={submitting}
          onSubmit={handleCreate}
          onCancel={goToList}
        />
      )}

      {view === 'details' && selectedId > 0 && (
          <TicketDetails
          service={service}
          ticketId={selectedId}
          currentUser={currentUser}
            isSupport={isDeveloperView}
          siteUrl={props.siteUrl}
          onBack={goToList}
          onChanged={loadTickets}
          onDeleted={handleTicketDeleted}
        />
      )}
    </section>
  );
};

export default TicketPortal;
