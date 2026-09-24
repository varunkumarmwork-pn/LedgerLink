import { useRef, useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import { Check, FileIcon, GridIcon, WorkbookIcon } from '../components/Icons.jsx';
import { createClient, deleteClient, startBuild, startLearning } from '../api.js';

const SHEETS = ['Balance Sheet', 'Profit & Loss', 'BS Schedules', 'P&L Notes', 'PPE', 'Accounting policies'];

export default function AddClient() {
  const navigate = useNavigate();
  const fileInput = useRef(null);
  const [mode, setMode] = useState('workbook'); // 'workbook' | 'tb'
  const [file, setFile] = useState(null);
  const [sheets, setSheets] = useState(SHEETS);
  const [name, setName] = useState('');
  const [entity, setEntity] = useState('Proprietorship');
  const [fy, setFy] = useState('2025-26');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState(null);

  const toggleSheet = (s) => setSheets((list) => (list.includes(s) ? list.filter((x) => x !== s) : [...list, s]));

  // Create the client, upload the file and follow the job (learn or build) on the Processing screen.
  const submit = async () => {
    if (!name.trim()) return setError('Enter the client name.');
    if (!file) return setError('Choose the Excel file to upload.');
    if (mode === 'tb' && sheets.length === 0) return setError('Tick at least one sheet to create.');
    setBusy(true);
    setError(null);
    let client = null;
    try {
      client = await createClient({ name: name.trim(), entity, fy });
      const { job_id: jobId } = mode === 'workbook'
        ? await startLearning(client.id, file)
        : await startBuild(client.id, file, SHEETS.filter((s) => sheets.includes(s)));
      navigate(`/processing/${jobId}`);
    } catch (e) {
      if (client) await deleteClient(client.id).catch(() => {}); // upload refused: do not leave a half-added client
      setError(e.message);
      setBusy(false);
    }
  };

  return (
    <main className="page narrow">
      <div className="page-title">
        <Link to="/" className="crumb">Clients</Link>
        <h1>Add a client</h1>
      </div>

      <div className="fields">
        <label>Client name<input className="field" value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Venus Agencies" /></label>
        <label>Entity type
          <select className="field" value={entity} onChange={(e) => setEntity(e.target.value)}><option>Proprietorship</option><option>Partnership firm</option><option>LLP</option><option>Private limited company</option></select>
        </label>
        <label>Financial year
          <select className="field" value={fy} onChange={(e) => setFy(e.target.value)}><option>2025-26</option><option>2026-27</option></select>
        </label>
      </div>

      <section className="stack">
        <h2 className="h2">What do you have for this client?</h2>
        <div className="choices">
          <button type="button" className={`choice ${mode === 'workbook' ? 'on' : ''}`} aria-pressed={mode === 'workbook'} onClick={() => setMode('workbook')}>
            {mode === 'workbook' && <span className="choice-tick"><Check /></span>}
            <WorkbookIcon />
            <span className="choice-title">Finished workbook</span>
            <span className="choice-body">Upload a completed year's Excel. LedgerLink reads every formula and learns how the trial balance flows into the schedules, balance sheet and P&amp;L.</span>
          </button>
          <button type="button" className={`choice ${mode === 'tb' ? 'on' : ''}`} aria-pressed={mode === 'tb'} onClick={() => setMode('tb')}>
            {mode === 'tb' && <span className="choice-tick"><Check /></span>}
            <GridIcon />
            <span className="choice-title">Trial balance only</span>
            <span className="choice-body">Upload the trial balance from Tally and choose which sheets to build: balance sheet, P&amp;L, schedules, PPE and notes.</span>
          </button>
        </div>
      </section>

      {mode === 'tb' && (
        <section className="stack">
          <h2 className="h2">Sheets to create</h2>
          <div className="chips">
            {SHEETS.map((s) => (
              <button key={s} type="button" className={`chip ${sheets.includes(s) ? 'on' : ''}`} aria-pressed={sheets.includes(s)} onClick={() => toggleSheet(s)}>{s}</button>
            ))}
          </div>
        </section>
      )}

      <div className="dropzone">
        <input ref={fileInput} type="file" accept=".xlsx,.xlsm" hidden onChange={(e) => setFile(e.target.files[0] || null)} />
        <div className="drop-file">
          <span className="file-badge"><FileIcon /></span>
          {file ? (
            <div className="stack-tight"><span className="strong">{file.name}</span><span className="muted small">{Math.round(file.size / 1024)} KB</span></div>
          ) : (
            <div className="stack-tight"><span className="strong">{mode === 'workbook' ? 'Drop the finished workbook here' : 'Drop the Tally trial balance here'}</span><span className="muted small">Excel file, .xlsx</span></div>
          )}
        </div>
        <button type="button" className="link-btn" onClick={() => fileInput.current.click()}>{file ? 'Choose another file' : 'Browse'}</button>
      </div>

      <div className="form-actions">
        {error && <span className="muted small" role="alert">{error}</span>}
        <Link to="/" className="link-btn">Cancel</Link>
        <button type="button" className="btn btn-primary" disabled={busy} onClick={submit}>
          {mode === 'workbook' ? 'Read workbook' : 'Build statements'}
        </button>
      </div>
    </main>
  );
}
