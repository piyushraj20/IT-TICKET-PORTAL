// =====================================================================
// components/TicketDetails.tsx
// The one screen that loads its own data, because it needs three extra
// lists (comments, history, attachments) that the ticket list doesn't.
//
// useEffect(() => {...}, [ticketId]) means: run this after render, and
// run it again whenever ticketId changes. That is how you load data
// in a React function component.
// =====================================================================

import * as React from 'react';
import { Icon } from '@fluentui/react';
import styles from './TicketPortal.module.scss';
import { PriorityChip, StatusChip, formatDate } from './Chips';
import { TicketService } from '../services/TicketService';
import {
  IAttachment, IComment, ICurrentUser, IHistoryEntry, ITicket,
  NEXT_STATUS, TicketStatus
} from '../models/ITicket';

export interface ITicketDetailsProps {
  service: TicketService;
  ticketId: number;
  currentUser: ICurrentUser;
  isSupport: boolean;
  siteUrl: string;
  onBack: () => void;
  onChanged: () => void;   // tells the parent to refresh the list
  onDeleted: () => Promise<void>;
}

const TicketDetails: React.FC<ITicketDetailsProps> = (props) => {
  const { service, ticketId, currentUser, isSupport } = props;

  const [ticket, setTicket] = React.useState<ITicket | undefined>(undefined);
  const [comments, setComments] = React.useState<IComment[]>([]);
  const [history, setHistory] = React.useState<IHistoryEntry[]>([]);
  const [attachments, setAttachments] = React.useState<IAttachment[]>([]);
  const [assignableUsers, setAssignableUsers] = React.useState<ICurrentUser[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState('');
  const [newComment, setNewComment] = React.useState('');
  const [nextStatus, setNextStatus] = React.useState<TicketStatus | ''>('');
  const [selectedAgentId, setSelectedAgentId] = React.useState('');
  const [ticketIdCopied, setTicketIdCopied] = React.useState(false);

  const load = React.useCallback(async (): Promise<void> => {
    setLoading(true);
    setError('');
    try {
      const [t, c, h, a, users] = await Promise.all([
        service.getTicketById(ticketId),
        service.getComments(ticketId),
        service.getHistory(ticketId),
        service.getAttachments(ticketId),
        service.getAssignableUsers()
      ]);
      setTicket(t);
      setComments(c);
      setHistory(h.length > 0 ? h : [{
        Id: 0,
        TicketItemId: t.Id,
        FromStatus: '',
        ToStatus: t.TicketStatus,
        Note: 'Current ticket status',
        AuthorName: t.EmployeeName,
        Created: t.Created
      }]);
      setAttachments(a);
      setAssignableUsers(users);
      setSelectedAgentId(t.AssignedToId ? String(t.AssignedToId) : '');
      setNextStatus('');
    } catch (e) {
      setError(`Could not load this ticket. ${(e as Error).message}`);
    } finally {
      setLoading(false);
    }
  }, [service, ticketId]);

  React.useEffect(() => { load().catch(() => undefined); }, [load]);

  /** Requesters can close or reopen their own resolved ticket; agents can do more. */
  const allowedStatuses = (): TicketStatus[] => {
    if (!ticket) { return []; }
    if (isSupport) { return NEXT_STATUS[ticket.TicketStatus]; }
    if (ticket.EmployeeName === currentUser.Title && ticket.TicketStatus === 'Resolved') {
      return ['Closed', 'In Progress'];
    }
    return [];
  };

  const handleAddComment = async (): Promise<void> => {
    if (!newComment.trim()) { return; }
    setBusy(true);
    setError('');
    try {
      await service.addComment(ticketId, newComment.trim());
      setNewComment('');
      setComments(await service.getComments(ticketId));
    } catch (e) {
      setError(
        `Comment not saved. Make sure the ${'TicketComments'} list exists and you have permission to add items. ` +
        `${(e as Error).message}`
      );
    } finally {
      setBusy(false);
    }
  };

  const handleStatusChange = async (): Promise<void> => {
    if (!ticket || !nextStatus) { return; }
    setBusy(true);
    setError('');
    try {
      const note = newComment.trim();
      await service.updateStatus(ticket.Id, ticket.TicketStatus, nextStatus, note);
      if (note) {
        await service.addComment(ticket.Id, note);
        setNewComment('');
      }
      await load();
      props.onChanged();
    } catch (e) {
      setError(`Status not updated. ${(e as Error).message}`);
    } finally {
      setBusy(false);
    }
  };

  const handleAssign = async (): Promise<void> => {
    if (!selectedAgentId) { return; }
    setBusy(true);
    try {
      await service.assignTo(ticketId, Number(selectedAgentId));
      await load();
      props.onChanged();
    } catch (e) {
      setError(`Assignment failed. ${(e as Error).message}`);
    } finally {
      setBusy(false);
    }
  };

  const copyTicketId = async (): Promise<void> => {
    if (!ticket) { return; }
    try {
      const ticketId = ticket.TicketID;
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(ticketId);
      } else {
        const textArea = document.createElement('textarea');
        textArea.value = ticketId;
        textArea.setAttribute('readonly', '');
        textArea.style.position = 'fixed';
        textArea.style.left = '-9999px';
        document.body.appendChild(textArea);
        textArea.select();
        const copied = document.execCommand('copy');
        document.body.removeChild(textArea);
        if (!copied) { throw new Error('Copy command failed'); }
      }
      setTicketIdCopied(true);
      window.setTimeout(() => setTicketIdCopied(false), 1800);
    } catch {
      setError('Ticket ID could not be copied. Select it manually instead.');
    }
  };

  const handleDelete = async (): Promise<void> => {
    if (!ticket || !window.confirm(`Delete ticket ${ticket.TicketID}? This cannot be undone.`)) { return; }
    setBusy(true);
    setError('');
    try {
      await service.deleteTicket(ticket.Id);
      await props.onDeleted();
    } catch (e) {
      setError(`Ticket not deleted. ${(e as Error).message}`);
      setBusy(false);
    }
  };

  if (loading) { return <div className={styles.loading}>Loading ticket...</div>; }

  if (!ticket) {
    return (
      <div>
        <div className={`${styles.banner} ${styles.bannerError}`}>{error || 'Ticket not found.'}</div>
        <button type="button" className={styles.button} onClick={props.onBack}>Back to list</button>
      </div>
    );
  }

  const statusOptions = allowedStatuses();
  const assignedAgentIndex = assignableUsers.findIndex(user => user.Id === ticket.AssignedToId);
  const assignedAgentLabel = assignedAgentIndex >= 0
    ? `Agent ${assignedAgentIndex + 1}`
    : ticket.AssignedToName || 'Unassigned';

  return (
    <div>
      <div className={styles.actions} style={{ marginBottom: 14 }}>
        <button type="button" className={styles.button} onClick={props.onBack}>Back to list</button>
        {isSupport && (
          <button type="button" className={styles.button} onClick={handleDelete} disabled={busy}>
            Delete ticket
          </button>
        )}
      </div>

      {error && <div className={`${styles.banner} ${styles.bannerError}`}>{error}</div>}

      <div className={styles.detailGrid}>

        {/* ---------- Left column ---------- */}
        <div>
          <div className={styles.panel}>
            <h3 className={styles.panelTitle}>Ticket information</h3>
            <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap', marginBottom: 12 }}>
              <span className={styles.ticketId}>{ticket.TicketID}</span>
              <button
                type="button"
                className={styles.linkButton}
                onClick={copyTicketId}
                title={ticketIdCopied ? 'Ticket ID copied' : 'Copy ticket ID'}
                aria-label={ticketIdCopied ? 'Ticket ID copied' : 'Copy ticket ID'}
              >
                <Icon iconName={ticketIdCopied ? 'CheckMark' : 'Copy'} aria-hidden="true" />
              </button>
              <PriorityChip value={ticket.Priority} />
              <StatusChip value={ticket.TicketStatus} />
            </div>

            <div style={{ fontSize: 18, fontWeight: 600, marginBottom: 14 }}>{ticket.Title}</div>

            <div className={styles.metaGrid}>
              <div>
                <span className={styles.metaLabel}>Raised by</span>
                <span className={styles.metaValue}>{ticket.EmployeeName}</span>
              </div>
              <div>
                <span className={styles.metaLabel}>Email</span>
                <span className={styles.metaValue}>{ticket.Email}</span>
              </div>
              <div>
                <span className={styles.metaLabel}>Category</span>
                <span className={styles.metaValue}>{ticket.Category}</span>
              </div>
              <div>
                <span className={styles.metaLabel}>Created</span>
                <span className={styles.metaValue}>{formatDate(ticket.Created)}</span>
              </div>
              <div>
                <span className={styles.metaLabel}>Assigned to</span>
                <span className={styles.metaValue}>{assignedAgentLabel}</span>
              </div>
            </div>

            <div className={styles.description}>{ticket.Description}</div>

            {attachments.length > 0 && (
              <div style={{ marginTop: 14 }}>
                <h3 className={styles.panelTitle}>Attachments</h3>
                <ul className={styles.attachmentList}>
                  {attachments.map(a => (
                    <li key={a.FileName}>
                      <a href={a.ServerRelativeUrl} target="_blank" rel="noreferrer">{a.FileName}</a>
                    </li>
                  ))}
                </ul>
              </div>
            )}
          </div>

          <div className={styles.panel} style={{ marginTop: 16 }}>
            <h3 className={styles.panelTitle}>Comments ({comments.length})</h3>

            {comments.length === 0 && <p className={styles.dim}>No comments yet. Start the conversation.</p>}

            {comments.map(c => (
              <div key={c.Id} className={styles.comment}>
                <div className={styles.commentHead}>
                  <span className={styles.commentAuthor}>{c.AuthorName}</span>
                  <span>{formatDate(c.Created)}</span>
                </div>
                <div className={styles.commentBody}>{c.CommentText}</div>
              </div>
            ))}

            <div className={styles.field} style={{ marginTop: 14 }}>
              <label className={styles.label} htmlFor="tp-comment">Add a comment</label>
              <textarea
                id="tp-comment"
                className={styles.textarea}
                value={newComment}
                onChange={e => setNewComment(e.target.value)}
                placeholder="Share an update, a question, or the steps you tried."
              />
            </div>

            <div className={styles.actions}>
              <button
                type="button"
                className={`${styles.button} ${styles.buttonPrimary}`}
                onClick={handleAddComment}
                disabled={busy || !newComment.trim()}
              >
                Post comment
              </button>
            </div>
          </div>
        </div>

        {/* ---------- Right column ---------- */}
        <div>
          <div className={styles.panel}>
            <h3 className={styles.panelTitle}>Status timeline</h3>
            <ul className={styles.timeline}>
              {history.map(h => (
                <li key={h.Id} className={styles.timelineItem}>
                  <div className={styles.timelineTitle}>
                    {h.FromStatus ? `${h.FromStatus} to ${h.ToStatus}` : h.ToStatus}
                  </div>
                  <div className={styles.dim}>{formatDate(h.Created)} &middot; {h.AuthorName}</div>
                  {h.Note && <div className={styles.dim}>{h.Note}</div>}
                </li>
              ))}
            </ul>
          </div>

          {(statusOptions.length > 0 || isSupport) && (
            <div className={styles.panel} style={{ marginTop: 16 }}>
              <h3 className={styles.panelTitle}>Update ticket</h3>

              {statusOptions.length > 0 && (
                <>
                  <div className={styles.field}>
                    <label className={styles.label} htmlFor="tp-next">Move status to</label>
                    <select
                      id="tp-next"
                      className={styles.select}
                      value={nextStatus}
                      onChange={e => setNextStatus(e.target.value as TicketStatus | '')}
                    >
                      <option value="">Choose a status</option>
                      {statusOptions.map(s => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>
                  <p className={styles.hint}>
                    Anything typed in the comment box is saved with the status change.
                  </p>
                  <div className={styles.actions}>
                    <button
                      type="button"
                      className={`${styles.button} ${styles.buttonPrimary}`}
                      onClick={handleStatusChange}
                      disabled={busy || !nextStatus}
                    >
                      Update status
                    </button>
                  </div>
                </>
              )}

              {isSupport && (
                <>
                  <div className={styles.field}>
                    <label className={styles.label} htmlFor="tp-assigned-agent">Assigned to agent</label>
                    <select
                      id="tp-assigned-agent"
                      className={styles.select}
                      value={selectedAgentId}
                      onChange={event => setSelectedAgentId(event.target.value)}
                    >
                      <option value="">Choose an agent</option>
                      {assignableUsers.map((user, index) => (
                        <option key={user.Id} value={user.Id}>Agent {index + 1}</option>
                      ))}
                    </select>
                  </div>
                  <div className={styles.actions}>
                    <button
                      type="button"
                      className={styles.button}
                      onClick={handleAssign}
                      disabled={busy || !selectedAgentId || Number(selectedAgentId) === ticket.AssignedToId}
                    >
                      Assign ticket
                    </button>
                  </div>
                </>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default TicketDetails;
