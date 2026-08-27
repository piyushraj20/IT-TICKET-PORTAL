// =====================================================================
// components/RaiseTicketForm.tsx
// A controlled form. "Controlled" means React state holds every value
// and the input just displays it - one source of truth.
//
// Note there is no <form> element. Inside a SharePoint page a real form
// submit can reload the page, so we use a plain button + onClick.
// =====================================================================

import * as React from 'react';
import styles from './TicketPortal.module.scss';
import { CATEGORIES, INewTicket, PRIORITIES, TicketPriority } from '../models/ITicket';

export interface IRaiseTicketFormProps {
  defaultName: string;
  defaultEmail: string;
  submitting: boolean;
  onSubmit: (data: INewTicket, files: File[]) => Promise<void>;
  onCancel: () => void;
}

const MAX_FILE_MB = 10;

const RaiseTicketForm: React.FC<IRaiseTicketFormProps> = (props) => {

  const [employeeName, setEmployeeName] = React.useState(props.defaultName);
  const [email, setEmail] = React.useState(props.defaultEmail);
  const [category, setCategory] = React.useState(CATEGORIES[0]);
  const [subject, setSubject] = React.useState('');
  const [description, setDescription] = React.useState('');
  const [priority, setPriority] = React.useState<TicketPriority>('Medium');
  const [files, setFiles] = React.useState<File[]>([]);
  const [errors, setErrors] = React.useState<Record<string, string>>({});

  const validate = (): boolean => {
    const next: Record<string, string> = {};
    if (!employeeName.trim()) { next.employeeName = 'Enter your name.'; }
    if (!email.trim() || email.indexOf('@') < 0) { next.email = 'Enter a valid email address.'; }
    if (!subject.trim()) { next.subject = 'Give the ticket a short subject.'; }
    else if (subject.length > 120) { next.subject = 'Keep the subject under 120 characters.'; }
    if (description.trim().length < 10) { next.description = 'Describe the issue in at least 10 characters.'; }

    const tooBig = files.filter(f => f.size > MAX_FILE_MB * 1024 * 1024);
    if (tooBig.length) { next.files = `Each file must be under ${MAX_FILE_MB} MB.`; }

    setErrors(next);
    return Object.keys(next).length === 0;
  };

  const handleSubmit = (): void => {
    if (!validate()) { return; }
    props.onSubmit(
      {
        Title: subject.trim(),
        EmployeeName: employeeName.trim(),
        Email: email.trim(),
        Category: category,
        Description: description.trim(),
        Priority: priority
      },
      files
    ).catch(() => { /* the parent shows the error banner */ });
  };

  const onFilesPicked = (e: React.ChangeEvent<HTMLInputElement>): void => {
    setFiles(e.target.files ? Array.prototype.slice.call(e.target.files) : []);
  };

  return (
    <div className={styles.form}>
      <div className={styles.formGrid}>

        <div className={styles.field}>
          <label className={styles.label} htmlFor="tp-name">Employee name</label>
          <input
            id="tp-name"
            className={`${styles.input} ${errors.employeeName ? styles.inputError : ''}`}
            value={employeeName}
            onChange={e => setEmployeeName(e.target.value)}
          />
          {errors.employeeName && <span className={styles.errorText}>{errors.employeeName}</span>}
        </div>

        <div className={styles.field}>
          <label className={styles.label} htmlFor="tp-email">Email</label>
          <input
            id="tp-email"
            type="email"
            className={`${styles.input} ${errors.email ? styles.inputError : ''}`}
            value={email}
            onChange={e => setEmail(e.target.value)}
          />
          {errors.email && <span className={styles.errorText}>{errors.email}</span>}
        </div>

        <div className={styles.field}>
          <label className={styles.label} htmlFor="tp-category">Category</label>
          <select
            id="tp-category"
            className={styles.select}
            value={category}
            onChange={e => setCategory(e.target.value)}
          >
            {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>

        <div className={styles.field}>
          <label className={styles.label} htmlFor="tp-priority">Priority</label>
          <select
            id="tp-priority"
            className={styles.select}
            value={priority}
            onChange={e => setPriority(e.target.value as TicketPriority)}
          >
            {PRIORITIES.map(p => <option key={p} value={p}>{p}</option>)}
          </select>
          <span className={styles.hint}>Critical means work has stopped for a team.</span>
        </div>

        <div className={`${styles.field} ${styles.formRowFull}`}>
          <label className={styles.label} htmlFor="tp-subject">Subject</label>
          <input
            id="tp-subject"
            className={`${styles.input} ${errors.subject ? styles.inputError : ''}`}
            value={subject}
            placeholder="Laptop will not connect to the office Wi-Fi"
            onChange={e => setSubject(e.target.value)}
          />
          {errors.subject && <span className={styles.errorText}>{errors.subject}</span>}
        </div>

        <div className={`${styles.field} ${styles.formRowFull}`}>
          <label className={styles.label} htmlFor="tp-description">Description</label>
          <textarea
            id="tp-description"
            className={`${styles.textarea} ${errors.description ? styles.inputError : ''}`}
            value={description}
            placeholder="What happened, when it started, and what you already tried."
            onChange={e => setDescription(e.target.value)}
          />
          {errors.description && <span className={styles.errorText}>{errors.description}</span>}
        </div>

        <div className={`${styles.field} ${styles.formRowFull}`}>
          <label className={styles.label} htmlFor="tp-files">Attachments</label>
          <input id="tp-files" type="file" multiple onChange={onFilesPicked} />
          {files.length > 0 && (
            <span className={styles.hint}>{files.map(f => f.name).join(', ')}</span>
          )}
          {errors.files && <span className={styles.errorText}>{errors.files}</span>}
        </div>

      </div>

      <div className={styles.actions}>
        <button
          type="button"
          className={`${styles.button} ${styles.buttonPrimary}`}
          onClick={handleSubmit}
          disabled={props.submitting}
        >
          {props.submitting ? 'Submitting...' : 'Submit ticket'}
        </button>
        <button
          type="button"
          className={styles.button}
          onClick={props.onCancel}
          disabled={props.submitting}
        >
          Cancel
        </button>
      </div>

      <p className={styles.hint}>
        A ticket number is generated automatically once you submit.
      </p>
    </div>
  );
};

export default RaiseTicketForm;
