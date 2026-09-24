import { useEffect, useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import { Check } from '../components/Icons.jsx';
import { watchJob } from '../api.js';

// Step titles come from the job (reading a workbook or rolling forward); these hints show
// under steps not reached yet. What each step found comes from the backend as it finishes.
const STEPS = [
  { title: 'Read the trial balance', hint: 'Find the TB sheet and check debit equals credit.' },
  { title: 'Find the sheets', hint: 'Balance sheet, P&L, schedules, notes and PPE.' },
  { title: 'Trace every link back to the trial balance', hint: 'Follow each formula to its TB ledger.' },
  { title: 'Check totals', hint: 'Unused ledgers, typed amounts, year-on-year links.' },
];
const LOG_ROWS = 5; // latest formulas traced

export default function Processing({ onFinished }) {
  const { jobId } = useParams();
  const navigate = useNavigate();
  const [job, setJob] = useState(null);
  const [lost, setLost] = useState(null);

  useEffect(() => watchJob(jobId, setJob, (e) => setLost(e.message)), [jobId]);

  useEffect(() => {
    if (job?.status !== 'done') return undefined;
    onFinished();
    const t = setTimeout(() => navigate(`/client/${job.client_id}`), 700);
    return () => clearTimeout(t);
  }, [job?.status, job?.client_id, navigate, onFinished]);

  const step = job ? job.step : 0;
  const error = job?.error || lost;
  const pct = Math.round((Math.min(step, STEPS.length) / STEPS.length) * 100);
  const log = job ? job.log.slice(-LOG_ROWS) : [];

  return (
    <main className="page processing">
      <div className="page-title">
        <h1 className="h1-sm">{job?.action ?? 'Reading'} {job?.client_name ?? '…'}</h1>
        <p>{job?.filename ?? ''}</p>
      </div>

      <div className="stack-tight">
        <div className="progress"><div style={{ width: `${pct}%` }} /></div>
        <div className="progress-meta"><span>Step {Math.min(step + 1, STEPS.length)} of {STEPS.length}</span><span>{pct}%</span></div>
      </div>

      <ol className="steps">
        {STEPS.map((s, i) => {
          const state = i < step ? 'done' : i === step ? 'active' : 'todo';
          const text = state === 'done' ? job.details[i] : state === 'active' ? (error || 'Working…') : s.hint;
          return (
            <li key={s.title} className={`step ${state}`}>
              <span className="step-mark">{state === 'done' ? <Check /> : state === 'active' ? <span /> : null}</span>
              <div className="stack-tight grow">
                <span className="step-title">{job?.titles?.[i] ?? s.title}</span>
                <span className="muted" role={state === 'active' && error ? 'alert' : undefined}>{state === 'todo' && job?.action !== 'Reading' ? '' : text}</span>
                {i === 2 && state !== 'todo' && log.length > 0 && (
                  <div className="log">
                    {log.map((l) => (<div key={l[0]} className="log-row"><span>{l[0]}</span><span>{l[1]}</span><span className="muted">{l[2]}</span></div>))}
                  </div>
                )}
                {state === 'active' && error && <Link to="/" className="link-btn">Back to clients</Link>}
              </div>
            </li>
          );
        })}
      </ol>
    </main>
  );
}
