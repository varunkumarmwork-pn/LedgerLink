// Calls to the Python backend. Vite forwards /api to it (see vite.config.js).

async function request(path, options) {
  let res;
  try {
    res = await fetch(`/api${path}`, options);
  } catch {
    throw new Error('Could not reach the LedgerLink server. Check that the backend is running.');
  }
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    const detail = Array.isArray(body.detail) ? body.detail.map((d) => d.msg).join('. ') : body.detail;
    throw new Error(detail || `Server error (${res.status}).`);
  }
  return res.status === 204 ? null : res.json();
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// "2026-09-12" -> "12 Sep 2026"
function formatDate(iso) {
  if (!iso) return '—';
  const [y, m, d] = iso.split('-');
  return `${d} ${MONTHS[Number(m) - 1]} ${y}`;
}

// Colour of the review dot for a status text from the backend.
function reviewTone(text) {
  if (/finalised|ready to finalise/i.test(text)) return 'green';
  if (/suggestion/i.test(text)) return 'amber';
  if (/to check|not read/i.test(text)) return 'red';
  return 'none';
}

// Backend client -> the shape the screens use.
function toView(c) {
  return {
    id: String(c.id),
    name: c.name,
    entity: c.entity_type,
    fy: c.financial_year,
    tbUpdated: formatDate(c.tb_updated_on),
    review: { tone: reviewTone(c.review_status), text: c.review_status },
  };
}

export async function listClients() {
  return (await request('/clients')).map(toView);
}

export function createClient({ name, entity, fy }) {
  return request('/clients', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, entity_type: entity, financial_year: fy }),
  });
}

export function deleteClient(id) {
  return request(`/clients/${id}`, { method: 'DELETE' });
}

// Uploads a finished workbook and starts LEARN MODE. Returns { job_id }.
export function startLearning(clientId, file) {
  const form = new FormData();
  form.append('file', file);
  return request(`/clients/${clientId}/learn`, { method: 'POST', body: form });
}

// The client's workbook for the spreadsheet: { workbook_id, financial_year, snapshot, cells }.
export function getWorkbook(clientId) {
  return request(`/clients/${clientId}/workbook`);
}

// Save cells the CA changed: [{ sheet, cell, value, formula }]. Resolves to the new status,
// trace and colours of each cell.
export function saveCells(clientId, changes) {
  return request(`/clients/${clientId}/workbook/cells`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ changes }),
  });
}

// Approve / Change a review item. Resolves to the whole workbook again (colours change).
export function actOnReview(clientId, item, action, change = {}) {
  return request(`/clients/${clientId}/workbook/review`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ item, action, ...change }),
  });
}

export function finaliseWorkbook(clientId) {
  return request(`/clients/${clientId}/workbook/finalise`, { method: 'POST' });
}

// Uploads a Tally TB and starts BUILD MODE for the ticked sheets. Returns { job_id }.
export function startBuild(clientId, file, sheets) {
  const form = new FormData();
  form.append('file', file);
  sheets.forEach((sheet) => form.append('sheets', sheet));
  return request(`/clients/${clientId}/build`, { method: 'POST', body: form });
}

// Uploads next year's Tally TB and starts ROLL-FORWARD MODE. Returns { job_id }.
export function startRollForward(clientId, file) {
  const form = new FormData();
  form.append('file', file);
  return request(`/clients/${clientId}/rollforward`, { method: 'POST', body: form });
}

export const downloadUrl = (clientId) => `/api/clients/${clientId}/workbook/download`;

// Live progress of a job: onUpdate gets the job state each time it changes.
// Returns a function that stops listening.
export function watchJob(jobId, onUpdate, onError) {
  const events = new EventSource(`/api/jobs/${jobId}/events`);
  events.onmessage = (e) => {
    const job = JSON.parse(e.data);
    onUpdate(job);
    if (job.status !== 'running') events.close();
  };
  events.onerror = () => {
    events.close();
    onError(new Error('Lost touch with the server while reading the workbook.'));
  };
  return () => events.close();
}
