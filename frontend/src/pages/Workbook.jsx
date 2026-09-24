import { useEffect, useRef, useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import { Check, Cross, Download, Upload } from '../components/Icons.jsx';
import Spreadsheet from '../components/Spreadsheet.jsx';
import { actOnReview, downloadUrl, finaliseWorkbook, getWorkbook, saveCells, startRollForward } from '../api.js';

export default function Workbook({ clients, onChanged }) {
  const { id } = useParams();
  const navigate = useNavigate();
  const client = clients.find((c) => c.id === id);
  const [workbook, setWorkbook] = useState(null); // snapshot, cells, status, ledgers
  const [review, setReview] = useState({ items: [], checks: [] });
  const [loadError, setLoadError] = useState(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState(null);
  const sheet = useRef(null); // tab the CA is on, kept when the sheet reloads
  const tbInput = useRef(null);

  const show = (data) => {
    setWorkbook(data);
    setReview({ items: data.review, checks: data.checks });
  };

  useEffect(() => {
    setWorkbook(null);
    setLoadError(null);
    getWorkbook(id).then(show).catch((e) => setLoadError(e.message));
  }, [id]);

  // Cell edits come back with the updated review, so the panel follows the sheet.
  const save = (changes) => saveCells(id, changes).then((result) => {
    setReview({ items: result.review, checks: result.checks });
    onChanged();
    return result;
  });

  const run = async (task) => {
    setBusy(true);
    setError(null);
    try {
      show(await task());
      onChanged();
    } catch (e) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  };

  const updateTb = async (file) => {
    if (!file) return;
    setError(null);
    try {
      const { job_id: jobId } = await startRollForward(id, file);
      navigate(`/processing/${jobId}`);
    } catch (e) {
      setError(e.message);
    }
  };

  if (!client) return <main className="page"><p className="muted">Loading client…</p></main>;

  const locked = workbook?.status === 'finalised';
  const items = review.items;

  return (
    <div className="workbook">
      <div className="client-bar">
        <div className="stack-tight">
          <div className="client-bar-title"><Link to="/" className="crumb">Clients</Link><span>{client.name}</span></div>
          <div className="client-meta"><span>{client.entity}</span><span>FY {workbook?.financial_year ?? client.fy}</span><span>Trial balance updated {client.tbUpdated}</span></div>
        </div>
        <div className="page-actions">
          <a className="btn btn-outline" href={downloadUrl(id)} download><Download />Download Excel</a>
          <input ref={tbInput} type="file" accept=".xlsx,.xlsm" hidden onChange={(e) => { updateTb(e.target.files[0]); e.target.value = ''; }} />
          <button type="button" className="btn btn-primary" disabled={!workbook} onClick={() => tbInput.current.click()}>
            <Upload />Update TB for FY {workbook?.next_financial_year ?? ''}
          </button>
        </div>
      </div>

      <div className="wb-body">
        <section className="sheet-area">
          {workbook
            ? <Spreadsheet key={workbook.workbook_id} snapshot={workbook.snapshot} cells={workbook.cells} onSave={save}
                readOnly={locked} sheet={sheet.current} onSheet={(name) => { sheet.current = name; }} />
            : <p className="muted empty sheet-message">{loadError || 'Opening the workbook…'}</p>}
        </section>

        <aside className="review">
          <div className="review-head"><h2>Review</h2><span className="muted small">{locked ? 'Finalised' : `${items.length} to check`}</span></div>

          <div className="checks">
            {review.checks.map((c) => (
              <div key={c.id} className={c.ok ? 'ok' : 'bad'} title={c.detail}>{c.ok ? <Check width="16" height="16" /> : <Cross />}{c.label}</div>
            ))}
          </div>

          {error && <p className="muted small" role="alert">{error}</p>}

          <div className="review-list">
            {items.map((item) => (
              <ReviewItem key={item.id} item={item} ledgers={workbook?.ledgers ?? []} busy={busy || locked}
                onAct={(action, change) => run(() => actOnReview(id, item.id, action, change))} />
            ))}
            {workbook && items.length === 0 && !locked && <p className="muted small">Everything is checked. You can finalise this year.</p>}
          </div>

          <div className="legend">
            <span><i className="sw st-g" />Linked as last year</span>
            <span><i className="sw st-a" />AI suggestion</span>
            <span><i className="sw st-r" />Needs you</span>
            <span><i className="sw st-p" />Edited by CA</span>
          </div>

          <div className="finalise">
            <button type="button" className="btn btn-primary wide" disabled={!workbook || items.length > 0 || locked || busy}
              onClick={() => run(() => finaliseWorkbook(id))}>
              {locked ? `FY ${workbook.financial_year} finalised` : `Finalise FY ${workbook?.financial_year ?? client.fy}`}
            </button>
            {!locked && items.length > 0 && <span className="muted small center">Clear the {items.length} items above to finalise</span>}
          </div>
        </aside>
      </div>
    </div>
  );
}

// One review item: Approve-type buttons act at once; Change opens a small form.
function ReviewItem({ item, ledgers, busy, onAct }) {
  const [changing, setChanging] = useState(false);
  const [choice, setChoice] = useState('');

  const submit = () => {
    const kind = item.change.kind;
    if (!choice) return;
    if (kind === 'ledger') onAct('change', { ledger: choice });
    else if (kind === 'cell') onAct('change', { cell: choice });
    else if (choice.startsWith('=')) onAct('change', { formula: choice });
    else onAct('change', { value: Number.isNaN(Number(choice)) ? choice : Number(choice) });
  };

  return (
    <div className="review-item">
      <div className="review-item-head"><span className={`sq sq-${item.tone}`} />{item.title}<span className="ref">{item.ref}</span></div>
      <p>{item.text}</p>
      {changing ? (
        <div className="stack-tight">
          {item.change.kind === 'ledger' && (
            <select className="field" value={choice} onChange={(e) => setChoice(e.target.value)} aria-label="Ledger">
              <option value="">Choose a ledger…</option>
              {ledgers.map((l) => <option key={l.key} value={l.key}>{l.name}{l.group ? ` · ${l.group}` : ''}</option>)}
            </select>
          )}
          {item.change.kind === 'cell' && (
            <select className="field" value={choice} onChange={(e) => setChoice(e.target.value)} aria-label="Cell">
              <option value="">Add to which cell…</option>
              {item.change.options.map((o) => <option key={o.cell} value={o.cell}>{o.label}</option>)}
            </select>
          )}
          {item.change.kind === 'value' && (
            <input className="field" value={choice} onChange={(e) => setChoice(e.target.value)} placeholder="Amount or =formula" aria-label="New value" />
          )}
          <div className="row-gap">
            <button type="button" className="btn-sm btn-sm-dark" disabled={busy || !choice} onClick={submit}>Save</button>
            <button type="button" className="btn-sm" onClick={() => setChanging(false)}>Cancel</button>
          </div>
        </div>
      ) : (
        <div className="row-gap">
          {item.actions.map((a, i) => (
            <button key={a.label} type="button" className={i === 0 ? 'btn-sm btn-sm-dark' : 'btn-sm'} disabled={busy}
              onClick={() => (a.action === 'change' ? setChanging(true) : onAct(a.action))}>{a.label}</button>
          ))}
        </div>
      )}
    </div>
  );
}
