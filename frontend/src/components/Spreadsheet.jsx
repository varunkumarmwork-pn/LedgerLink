import { useEffect, useRef, useSyncExternalStore } from 'react';
import { LocaleType, Univer } from '@univerjs/core';
import { FUniver } from '@univerjs/core/facade';
import { defaultTheme } from '@univerjs/themes';
import { BuiltInUIPart, SetRangeValuesMutation, UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import '@univerjs/preset-sheets-core/lib/index.css';
import { LinkIcon } from './Icons.jsx';

// Univer's accent colours are blue by default; use the plum accent instead (never blue).
const PLUM = {
  50: '#F7F2FB', 100: '#EEE6F5', 200: '#DCCBEA', 300: '#C3A6DA', 400: '#9E73C0',
  500: '#7D4BA6', 600: '#5B2C83', 700: '#4E2571', 800: '#401E5D', 900: '#33184A',
};
const THEME = { ...defaultTheme, primary: PLUM, blue: PLUM, indigo: PLUM, jiqing: PLUM };

// 0-based row/column -> "D21"
function address(row, col) {
  let letters = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return `${letters}${row + 1}`;
}

// "D21" -> { row: 20, col: 3 }
function position(cell) {
  const [, letters, digits] = cell.match(/^([A-Z]+)(\d+)$/);
  const col = [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return { row: Number(digits) - 1, col };
}

// What the trace line shows. Lives outside React state because the trace line is rendered
// by Univer (under its formula bar), not by this component.
function createTraceStore() {
  let state = { trace: null, reason: null, error: null };
  const listeners = new Set();
  return {
    get: () => state,
    set: (next) => { state = { ...state, ...next }; listeners.forEach((l) => l()); },
    subscribe: (l) => { listeners.add(l); return () => listeners.delete(l); },
  };
}

function TraceLine({ store }) {
  const { trace, reason, error } = useSyncExternalStore(store.subscribe, store.get);
  if (error) return <div className="trace-bar" role="alert"><span className="strong-ink">Not saved.</span><span>{error}</span></div>;
  if (!trace) return <div className="trace-bar"><span className="faint">Select a cell to see where its figure comes from.</span></div>;
  return (
    <div className="trace-bar">
      <LinkIcon style={{ color: 'var(--accent)' }} />
      <span>{trace[0]}</span><span className="faint">from</span><span className="strong-ink">{trace[1]}</span>
      {reason && <><span className="faint">·</span><span>{reason}</span></>}
    </div>
  );
}

// The client's workbook in Univer: Excel-like editing, formulas, formula bar, sheet tabs and
// clipboard. `cells` holds the review status, trace and reason per cell ({sheet: {D21: {...}}});
// every content change is passed to `onSave(changes)`, which resolves to the new colours.
// readOnly locks a finalised year. `sheet` is the tab to open on; `onSheet` reports tab changes.
export default function Spreadsheet({ snapshot, cells, onSave, readOnly = false, sheet, onSheet }) {
  const host = useRef(null);
  const callbacks = useRef({});
  callbacks.current = { onSave, onSheet, sheet };

  useEffect(() => {
    const details = structuredClone(cells);
    const store = createTraceStore();
    let selected = null; // { sheet, cell }

    // Each Univer gets its own element: in development React mounts effects twice, and the
    // first Univer is disposed a moment later (below) while the second is already running.
    const container = document.createElement('div');
    container.className = 'univer-root';
    host.current.appendChild(container);

    const univer = new Univer({ locale: LocaleType.EN_US, locales: { [LocaleType.EN_US]: sheetsCoreEnUS }, theme: THEME });
    const { plugins } = UniverSheetsCorePreset({ container, toolbar: false });
    plugins.forEach((p) => (Array.isArray(p) ? univer.registerPlugin(p[0], p[1]) : univer.registerPlugin(p)));
    const api = FUniver.newAPI(univer);
    const workbook = api.createWorkbook(snapshot);
    if (callbacks.current.sheet) workbook.getSheetByName(callbacks.current.sheet)?.activate();
    if (readOnly) workbook.setEditable(false);
    const disposables = [];

    // Trace line directly under Univer's formula bar.
    disposables.push(api.registerUIPart(BuiltInUIPart.HEADER, () => <TraceLine store={store} />));

    const showTrace = () => {
      const info = selected ? details[selected.sheet]?.[selected.cell] : null;
      store.set({ trace: info?.trace ?? null, reason: info?.reason ?? null });
    };

    disposables.push(api.addEvent(api.Event.SelectionChanged, ({ worksheet, selections }) => {
      const range = selections?.[0];
      if (!range || !worksheet) return;
      selected = { sheet: worksheet.getSheetName(), cell: address(range.startRow, range.startColumn) };
      callbacks.current.onSheet?.(selected.sheet);
      showTrace();
    }));

    // Colour a cell without adding a step to undo: apply the mutation directly. onlyLocal
    // marks it as ours so the listener below ignores it.
    const paint = ({ sheet, cell, bg, ink }) => {
      const worksheet = workbook.getSheetByName(sheet);
      if (!worksheet) return;
      const { row, col } = position(cell);
      api.syncExecuteCommand(SetRangeValuesMutation.id, {
        unitId: workbook.getId(),
        subUnitId: worksheet.getSheetId(),
        cellValue: { [row]: { [col]: { s: { bg: bg ? { rgb: bg } : null, cl: ink ? { rgb: ink } : null } } } },
      }, { onlyLocal: true });
    };

    // Typing, paste, fill, delete and undo all change cells through this mutation. Skipped:
    // results written by the formula engine (fromFormula / onlyLocal, e.g. on load), our own
    // colouring, and style-only changes.
    disposables.push(api.onCommandExecuted((command, options) => {
      if (command.id !== SetRangeValuesMutation.id || options?.fromFormula || options?.onlyLocal) return;
      const { unitId, subUnitId, cellValue } = command.params;
      const worksheet = unitId === workbook.getId() && workbook.getSheetBySheetId(subUnitId);
      if (!worksheet) return;
      const changes = [];
      Object.entries(cellValue).forEach(([r, columns]) => Object.entries(columns).forEach(([c, value]) => {
        if (!value || !['v', 'f', 'si', 'p'].some((k) => k in value)) return;
        const range = worksheet.getRange(Number(r), Number(c));
        const formula = range.getFormula() || null;
        changes.push({ sheet: worksheet.getSheetName(), cell: address(Number(r), Number(c)), formula, value: formula ? null : range.getValue() ?? null });
      }));
      if (!changes.length) return;

      callbacks.current.onSave(changes)
        .then(({ cells: saved }) => {
          saved.forEach((c) => {
            (details[c.sheet] ??= {})[c.cell] = { status: c.status, trace: c.trace, reason: c.reason };
            paint(c);
          });
          store.set({ error: null });
          showTrace();
        })
        .catch((e) => store.set({ error: e.message }));
    }));

    return () => {
      disposables.forEach((d) => d.dispose());
      // Dispose after React finishes this render; Univer unmounts its own React root.
      setTimeout(() => { univer.dispose(); container.remove(); }, 0);
    };
  }, [snapshot, cells, readOnly]);

  return <div ref={host} className="univer-host" />;
}
