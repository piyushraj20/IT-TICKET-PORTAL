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

type View = 'list' | 'new' | 'details';
type DemoRole = 'user' | 'developer';
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

  const [view, setView] = React.useState<View>('list');
  const [allTickets, setAllTickets] = React.useState<ITicket[]>([]);
  const [filter, setFilter] = React.useState<ITicketFilter>(emptyFilter);
  const [mineOnly, setMineOnly] = React.useState(false);
  const [selectedId, setSelectedId] = React.useState<number>(0);
  const [currentUser, setCurrentUser] = React.useState<ICurrentUser>({ Id: 0, Title: '', Email: '' });
  const [isSupport, setIsSupport] = React.useState(false);
  const [demoRole, setDemoRole] = React.useState<DemoRole>('user');
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
        const [user, support] = await Promise.all([
          service.getCurrentUser(),
          service.isInGroup(props.supportGroupName)
        ]);
        setCurrentUser(user);
        setIsSupport(support);
        setDemoRole(support ? 'developer' : 'user');
      } catch {
        // Non-fatal: the portal still works, just without agent controls.
      }
      await loadTickets();
    };
    init().catch(() => undefined);
  }, [service, props.supportGroupName, loadTickets]);

  const isDeveloperView = demoRole === 'developer';

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

  const goToList = (): void => {
    setSuccess('');
    setView('list');
  };

  const selectStatusCard = (status: TicketStatus | ''): void => {
    setFilter({ ...filter, status });
    setView('list');
  };

  // -------------------------------------------------------------------
  // Render
  // -------------------------------------------------------------------
  return (
    <section className={styles.portal}>

      <header className={styles.header}>
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
            onClick={() => selectDemoRole('developer')}
            aria-pressed={demoRole === 'developer'}
          >
            View as developer
          </button>
        </div>
      </header>

      <nav className={styles.nav}>
        <button
          type="button"
          className={`${styles.navButton} ${view === 'list' ? styles.navButtonActive : ''}`}
          onClick={goToList}
        >
          Dashboard
        </button>
        {!isDeveloperView && (
          <button
            type="button"
            className={`${styles.navButton} ${view === 'new' ? styles.navButtonActive : ''}`}
            onClick={() => setView('new')}
          >
            Raise ticket
          </button>
        )}
        {view === 'list' && !isDeveloperView && (
          <button
            type="button"
            className={`${styles.navButton} ${mineOnly ? styles.navButtonActive : ''}`}
            onClick={() => setMineOnly(!mineOnly)}
            aria-pressed={mineOnly}
          >
            {mineOnly ? 'Showing my tickets' : 'Show only my tickets'}
          </button>
        )}
        <button type="button" className={styles.navButton} onClick={loadTickets}>
          Refresh
        </button>
      </nav>

      {error && <div className={`${styles.banner} ${styles.bannerError}`}>{error}</div>}
      {success && <div className={`${styles.banner} ${styles.bannerSuccess}`}>{success}</div>}

      {view === 'list' && (
        <>
          {!isDeveloperView && (
            <Dashboard
              tickets={scopedTickets}
              activeStatus={filter.status}
              onSelectStatus={selectStatusCard}
            />
          )}
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
        />
      )}
    </section>
  );
};

export default TicketPortal;
