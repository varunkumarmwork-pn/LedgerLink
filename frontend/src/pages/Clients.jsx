import { useState } from 'react';
import { Link } from 'react-router-dom';
import { Plus, Search, Trash } from '../components/Icons.jsx';

export default function Clients({ clients, loaded, loadError, onDelete }) {
  const [query, setQuery] = useState('');
  const [toDelete, setToDelete] = useState(null);
  const [typed, setTyped] = useState('');
  const [deleting, setDeleting] = useState(false);
  const [deleteError, setDeleteError] = useState(null);

  const shown = clients.filter((c) => c.name.toLowerCase().includes(query.toLowerCase()));

  const closeDialog = () => { setToDelete(null); setTyped(''); setDeleteError(null); };
  const confirmDelete = async () => {
    setDeleting(true);
    try {
      await onDelete(toDelete.id);
      closeDialog();
    } catch (e) {
      setDeleteError(e.message);
    } finally {
      setDeleting(false);
    }
  };

  return (
    <main className="page">
      <div className="page-head">
        <div className="page-title">
          <h1>Clients</h1>
          <p>{clients.length} clients. Open one to see its workbook.</p>
        </div>
        <div className="page-actions">
          <label className="search">
            <Search />
            <input type="text" placeholder="Search clients" aria-label="Search clients" value={query} onChange={(e) => setQuery(e.target.value)} />
          </label>
          <Link to="/add" className="btn btn-primary"><Plus />Add client</Link>
        </div>
      </div>

      <table className="clients">
        <thead>
          <tr>
            <th style={{ width: '38%' }}>Client</th>
            <th>Financial year</th>
            <th>Trial balance updated</th>
            <th>Review</th>
            <th style={{ width: 56 }} />
          </tr>
        </thead>
        <tbody>
          {shown.map((c) => (
            <tr key={c.id}>
              <td>
                <Link to={`/client/${c.id}`} className="client-name">{c.name}</Link>
                <div className="muted small">{c.entity}</div>
              </td>
              <td>{c.fy}</td>
              <td>{c.tbUpdated}</td>
              <td className={c.review.tone === 'none' ? 'muted' : ''}>
                <span className="status"><span className={`dot dot-${c.review.tone}`} />{c.review.text}</span>
              </td>
              <td className="row-action">
                <button type="button" className="icon-btn danger" aria-label={`Delete ${c.name}`} onClick={() => setToDelete(c)}>
                  <Trash />
                </button>
              </td>
            </tr>
          ))}
          {shown.length === 0 && (
            <tr><td colSpan={5} className="muted empty">
              {loadError || (!loaded ? 'Loading clients…' : query ? `No clients match "${query}".` : 'No clients yet. Add one to get started.')}
            </td></tr>
          )}
        </tbody>
      </table>

      {toDelete && (
        <div className="overlay" onClick={closeDialog}>
          <div className="dialog" role="dialog" aria-modal="true" aria-labelledby="del-title" onClick={(e) => e.stopPropagation()}>
            <h2 id="del-title">Delete {toDelete.name}?</h2>
            <p className="muted">This removes the client and every workbook and trial balance saved for it. Type DELETE to confirm.</p>
            <input className="field" autoFocus value={typed} onChange={(e) => setTyped(e.target.value)} placeholder="DELETE" aria-label="Type DELETE to confirm" />
            {deleteError && <p className="muted small" role="alert">{deleteError}</p>}
            <div className="dialog-actions">
              <button type="button" className="btn btn-ghost" onClick={closeDialog}>Cancel</button>
              <button type="button" className="btn btn-danger" disabled={typed !== 'DELETE' || deleting} onClick={confirmDelete}>Delete client</button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
