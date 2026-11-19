<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Control de correspondencia RSO II</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap" rel="stylesheet">
<script src="https://cdn.tailwindcss.com/3.4.17"></script>
<script>
  tailwind.config = {
    theme: {
      extend: {
        fontFamily: {
          sans: ['"IBM Plex Sans"', 'system-ui', 'sans-serif'],
          mono: ['"IBM Plex Mono"', 'ui-monospace', 'monospace']
        }
      }
    }
  };
</script>
<style>
  html, body { height: 100%; }
  body { background: #f3f5f8; color: #1e293b; font-family: "IBM Plex Sans", system-ui, sans-serif; }
  button:focus-visible, select:focus-visible, input:focus-visible, textarea:focus-visible { outline: 2px solid #1f4e79; outline-offset: 2px; }
</style>
</head>
<body>
<div id="root"></div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/react/18.3.1/umd/react.production.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/react-dom/18.3.1/umd/react-dom.production.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/7.26.10/babel.min.js"></script>

<script type="text/babel" data-presets="react">
const { useState, useMemo, useEffect, useRef } = React;

/* ---------- Configuración ---------- */
const COLLECTION = 'records';
const IMPORT_FILE = 'records.json';
const DEPENDENCIAS_BASE = ['OFEX', 'CERT', 'PQRS', 'OFIN', 'COTI', 'COMU'];
const MEDIOS = ['Correo electrónico', 'Físico', 'Ventanilla Única', 'Mensajería'];
const TIPOS = [
  { v: 'Recibida', l: 'Recibida (Entrada)', cls: 'bg-sky-50 text-sky-800 ring-sky-600/20' },
  { v: 'Enviada', l: 'Enviada (Salida)', cls: 'bg-indigo-50 text-indigo-800 ring-indigo-600/20' },
  { v: 'Interna', l: 'Interna', cls: 'bg-stone-100 text-stone-700 ring-stone-500/20' }
];
const ESTADOS = [
  { v: 'Pendiente', cls: 'bg-amber-50 text-amber-800 ring-amber-600/30' },
  { v: 'En trámite', cls: 'bg-blue-50 text-blue-800 ring-blue-600/30' },
  { v: 'Respondido', cls: 'bg-emerald-50 text-emerald-800 ring-emerald-600/30' },
  { v: 'Archivado', cls: 'bg-slate-100 text-slate-600 ring-slate-500/30' }
];
const PAGE_SIZE = 20;
const inputCls = 'w-full rounded-md border border-slate-300 bg-white px-3 py-2 text-sm text-slate-900 placeholder-slate-400 focus:border-[#1f4e79] focus:outline-none focus:ring-2 focus:ring-[#1f4e79]/30';

/* ---------- Utilidades de fecha y texto ---------- */
const pad = n => String(n).padStart(2, '0');
const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const hhmm = d => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
const addDays = (iso, n) => { const [y, m, d] = iso.split('-').map(Number); const x = new Date(y, m - 1, d + n); return `${x.getFullYear()}-${pad(x.getMonth() + 1)}-${pad(x.getDate())}`; };
const diffDays = (a, b) => Math.round((new Date(a + 'T00:00') - new Date(b + 'T00:00')) / 86400000);
const fmtDate = iso => { if (!iso) return '—'; const [y, m, d] = iso.split('-'); return `${d}/${m}/${y}`; };
const estadoCls = v => (ESTADOS.find(e => e.v === v) || ESTADOS[0]).cls;
const tipoOf = v => TIPOS.find(t => t.v === v) || TIPOS[0];

// Vencida o próxima a vencer: solo si la comunicación sigue abierta
function vencimiento(r, today) {
  if (!r.fechaLimite || r.estado === 'Respondido' || r.estado === 'Archivado') return null;
  const dif = diffDays(r.fechaLimite, today);
  if (dif < 0) return 'vencida';
  if (dif <= 3) return 'proxima';
  return null;
}

// Radicado con formato RAD-AAAA-NNN, consecutivo dentro del año en curso
function nextRadicado(list) {
  const year = new Date().getFullYear();
  const re = new RegExp(`^RAD-${year}-(\\d+)$`);
  let max = 0;
  list.forEach(r => { const m = re.exec(r.radicado || ''); if (m) max = Math.max(max, Number(m[1])); });
  return `RAD-${year}-${String(max + 1).padStart(3, '0')}`;
}

function buildCSV(list) {
  const head = ['Radicado', 'Fecha de radicación', 'Hora', 'Tipo', 'Remitente / Destinatario', 'Asunto', 'Medio de recepción / envío', 'Dependencia', 'Responsable', 'Fecha límite de respuesta', 'Estado'];
  const esc = v => `"${String(v == null ? '' : v).replace(/"/g, '""')}"`;
  const rows = list.map(r => [r.radicado, fmtDate(r.fechaRad), r.hora || '', r.tipo, r.contraparte, r.asunto, r.medio, r.dependencia, r.responsable, fmtDate(r.fechaLimite), r.estado]);
  // BOM + separador ";" para que Excel en español abra las columnas bien
  return '﻿' + [head, ...rows].map(row => row.map(esc).join(';')).join('\r\n');
}

/* ---------- Componentes pequeños ---------- */
const Pill = ({ cls, children }) => (
  <span className={`inline-flex items-center whitespace-nowrap rounded-full px-2.5 py-0.5 text-xs font-medium ring-1 ring-inset ${cls}`}>{children}</span>
);

function Field({ label, required, error, full, children }) {
  return (
    <label className={`block ${full ? 'md:col-span-2' : ''}`}>
      <span className="mb-1 block text-sm font-medium text-slate-700">{label}{required && <span className="text-rose-600"> *</span>}</span>
      {children}
      {error && <span className="mt-1 block text-xs text-rose-600">{error}</span>}
    </label>
  );
}

function Modal({ title, onClose, wide, children }) {
  useEffect(() => {
    const onKey = e => { if (e.key === 'Escape') onClose(); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4"
         onMouseDown={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div role="dialog" aria-modal="true" aria-label={title}
           className={`max-h-[92vh] w-full overflow-y-auto rounded-xl bg-white shadow-2xl ${wide ? 'max-w-3xl' : 'max-w-md'}`}>
        <div className="flex items-center justify-between border-b border-slate-200 px-6 py-4">
          <h2 className="text-lg font-semibold text-slate-900">{title}</h2>
          <button type="button" onClick={onClose} aria-label="Cerrar"
                  className="rounded-md px-2 py-1 text-xl leading-none text-slate-500 hover:bg-slate-100">×</button>
        </div>
        <div className="px-6 py-5">{children}</div>
      </div>
    </div>
  );
}

/* ---------- Formulario de radicación ---------- */
function RecordForm({ record, records, deps, onSave, onCancel }) {
  const isNew = !record;
  const [now] = useState(() => new Date());
  const [preview] = useState(() => (isNew ? nextRadicado(records) : record.radicado));
  const [f, setF] = useState(() => record ? { ...record } : {
    tipo: 'Recibida', contraparte: '', asunto: '', medio: 'Correo electrónico',
    dependencia: 'OFEX', responsable: '', fechaLimite: addDays(todayISO(), 15), estado: 'Pendiente'
  });
  const [errors, setErrors] = useState({});
  const [saving, setSaving] = useState(false);

  const set = k => e => setF(p => ({ ...p, [k]: e.target.value }));

  const submit = async e => {
    e.preventDefault();
    if (saving) return;
    const err = {};
    if (!f.contraparte.trim()) err.contraparte = 'Indica la persona o entidad.';
    if (!f.asunto.trim()) err.asunto = 'Describe brevemente el contenido.';
    if (!f.dependencia.trim()) err.dependencia = 'Indica el área encargada.';
    setErrors(err);
    if (Object.keys(err).length) return;
    const clean = {
      ...f,
      contraparte: f.contraparte.trim(),
      asunto: f.asunto.trim(),
      dependencia: f.dependencia.trim().toUpperCase(),
      responsable: f.responsable.trim()
    };
    setSaving(true);
    if (isNew) {
      const rad = nextRadicado(records);
      await onSave({
        ...clean,
        id: rad,              // el id del documento es el propio radicado
        radicado: rad,
        fechaRad: todayISO(),
        hora: hhmm(new Date()),
        origen: 'Nuevo'
      }, true);
    } else {
      await onSave({ ...record, ...clean }, false);
    }
    setSaving(false);
  };

  return (
    <form onSubmit={submit} noValidate className="space-y-5">
      <div className="grid gap-3 rounded-lg bg-slate-50 p-4 text-sm sm:grid-cols-2">
        <div>
          <span className="block text-xs uppercase tracking-wide text-slate-500">Número de radicado</span>
          <span className="font-mono font-semibold text-slate-900">{preview}</span>
          {isNew && <span className="block text-xs text-slate-500">Se asigna al guardar</span>}
        </div>
        <div>
          <span className="block text-xs uppercase tracking-wide text-slate-500">Fecha y hora de radicación</span>
          <span className="font-semibold text-slate-900">
            {isNew ? `${fmtDate(todayISO())} · ${hhmm(now)}` : `${fmtDate(record.fechaRad)}${record.hora ? ' · ' + record.hora : ''}`}
          </span>
          <span className="block text-xs text-slate-500">{isNew ? 'Automática al guardar' : 'No se modifica'}</span>
        </div>
      </div>

      <div className="grid gap-4 md:grid-cols-2">
        <Field label="Tipo de comunicación" required>
          <select className={inputCls} value={f.tipo} onChange={set('tipo')}>
            {TIPOS.map(t => <option key={t.v} value={t.v}>{t.l}</option>)}
          </select>
        </Field>
        <Field label="Medio de recepción / envío">
          <select className={inputCls} value={f.medio} onChange={set('medio')}>
            {MEDIOS.map(m => <option key={m} value={m}>{m}</option>)}
          </select>
        </Field>
        <Field label="Remitente / Destinatario" required error={errors.contraparte}>
          <input className={inputCls} value={f.contraparte} onChange={set('contraparte')} placeholder="Persona o entidad" />
        </Field>
        <Field label="Dependencia" required error={errors.dependencia}>
          <input className={inputCls} list="lista-dependencias" value={f.dependencia} onChange={set('dependencia')} placeholder="Ej. OFEX" />
          <datalist id="lista-dependencias">
            {deps.map(d => <option key={d} value={d} />)}
          </datalist>
        </Field>
        <Field label="Responsable del trámite">
          <input className={inputCls} value={f.responsable} onChange={set('responsable')} placeholder="Nombre de quien lo gestiona" />
        </Field>
        <Field label="Fecha límite de respuesta">
          <input type="date" className={inputCls} value={f.fechaLimite || ''} onChange={set('fechaLimite')} />
        </Field>
        <Field label="Asunto" required full error={errors.asunto}>
          <textarea rows={3} className={inputCls} value={f.asunto} onChange={set('asunto')} placeholder="Descripción breve del contenido" />
        </Field>
        <Field label="Estado">
          <select className={inputCls} value={f.estado} onChange={set('estado')}>
            {ESTADOS.map(s => <option key={s.v} value={s.v}>{s.v}</option>)}
          </select>
        </Field>
      </div>

      <div className="flex justify-end gap-3 border-t border-slate-200 pt-4">
        <button type="button" onClick={onCancel} className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50">Cancelar</button>
        <button type="submit" disabled={saving} className="rounded-md bg-[#1f4e79] px-4 py-2 text-sm font-semibold text-white hover:bg-[#173b5c] disabled:opacity-60">
          {saving ? 'Guardando…' : isNew ? 'Radicar comunicación' : 'Guardar cambios'}
        </button>
      </div>
    </form>
  );
}

/* ---------- Aplicación ---------- */
function App() {
  const dbRef = useRef(null);
  const [records, setRecords] = useState([]);
  const [dbState, setDbState] = useState('loading'); // loading | ready | off
  const [importing, setImporting] = useState(null);  // null | {done, total}
  const [q, setQ] = useState('');
  const [fTipo, setFTipo] = useState('');
  const [fEstado, setFEstado] = useState('');
  const [fDep, setFDep] = useState('');
  const [sort, setSort] = useState({ key: 'fechaRad', dir: 'desc' });
  const [page, setPage] = useState(1);
  const [formRec, setFormRec] = useState(undefined); // undefined = cerrado · null = nuevo · objeto = editar
  const [confirm, setConfirm] = useState(null);
  const [csvData, setCsvData] = useState(null);
  const [toast, setToast] = useState(null);
  const today = todayISO();

  // Conexión a la base compartida: una sola suscripción en vivo a la colección
  useEffect(() => {
    let unsub = null, cancelled = false;
    const timer = setTimeout(() => { if (!cancelled) setDbState(s => (s === 'loading' ? 'off' : s)); }, 10000);
    (async () => {
      let db = null;
      try { db = window.claude && window.claude.use ? await window.claude.use('db') : null; } catch (e) { db = null; }
      if (cancelled) return;
      if (!db) { setDbState('off'); return; }
      dbRef.current = db;
      try {
        const off = db.collection(COLLECTION).onSnapshot(snap => {
          const items = snap && snap.docs ? snap.docs : (Array.isArray(snap) ? snap : []);
          const list = items.map(d => {
            const data = typeof d.data === 'function' ? d.data() : (d.data || {});
            return { ...data, id: d.id || data.id };
          });
          setRecords(list);
          setDbState('ready');
          clearTimeout(timer);
        });
        unsub = off;
      } catch (e) { setDbState('off'); }
    })();
    return () => { cancelled = true; clearTimeout(timer); if (typeof unsub === 'function') unsub(); };
  }, []);

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 3200);
    return () => clearTimeout(t);
  }, [toast]);
  useEffect(() => { setPage(1); }, [q, fTipo, fEstado, fDep, sort]);

  const notify = (msg, type = 'ok') => setToast({ id: Date.now(), msg, type });
  const col = () => dbRef.current.collection(COLLECTION);

  // Ejecuta una escritura compartida y avisa si no se pudo
  const runWrite = async (fn, okMsg) => {
    if (!dbRef.current) { notify('Sin acceso a los datos compartidos. Inicia sesión en claude.ai.', 'warn'); return false; }
    try { await fn(); if (okMsg) notify(okMsg); return true; }
    catch (e) { notify('No se pudo guardar. Revisa tu conexión o tu rol de edición.', 'warn'); return false; }
  };

  const deps = useMemo(() =>
    Array.from(new Set([...DEPENDENCIAS_BASE, ...records.map(r => r.dependencia)])).filter(Boolean).sort()
  , [records]);

  const stats = useMemo(() => {
    let pend = 0, tram = 0, venc = 0, prox = 0;
    records.forEach(r => {
      if (r.estado === 'Pendiente') pend++;
      if (r.estado === 'En trámite') tram++;
      const v = vencimiento(r, today);
      if (v === 'vencida') venc++;
      if (v === 'proxima') prox++;
    });
    return { total: records.length, pend, tram, venc, prox };
  }, [records, today]);

  const filtered = useMemo(() => {
    const term = q.trim().toLowerCase();
    const list = records.filter(r =>
      (!term || [r.radicado, r.asunto, r.contraparte].some(x => (x || '').toLowerCase().includes(term))) &&
      (!fTipo || r.tipo === fTipo) &&
      (!fEstado || r.estado === fEstado) &&
      (!fDep || r.dependencia === fDep)
    );
    const dir = sort.dir === 'asc' ? 1 : -1;
    return list.sort((a, b) => {
      const va = a[sort.key] || '', vb = b[sort.key] || '';
      // Sin valor siempre al final, sin importar la dirección
      if (!va || !vb) return (!va ? 1 : 0) - (!vb ? 1 : 0);
      const c = sort.key === 'radicado'
        ? va.localeCompare(vb, 'es', { numeric: true })
        : (va < vb ? -1 : va > vb ? 1 : 0);
      return c * dir;
    });
  }, [records, q, fTipo, fEstado, fDep, sort]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const current = Math.min(page, totalPages);
  const pageRows = filtered.slice((current - 1) * PAGE_SIZE, current * PAGE_SIZE);
  const hasFilters = q || fTipo || fEstado || fDep;

  const toggleSort = key => setSort(s => s.key === key ? { key, dir: s.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });

  const saveRecord = async (rec, isNew) => {
    const ok = await runWrite(
      () => col().doc(rec.id).set(rec),
      isNew ? `Radicado ${rec.radicado} registrado.` : `Radicado ${rec.radicado} actualizado.`
    );
    if (ok) setFormRec(undefined);
  };

  const changeEstado = (id, estado) =>
    runWrite(() => col().doc(id).update({ estado }), `Estado cambiado a “${estado}”.`);

  const askDelete = r => setConfirm({
    title: 'Eliminar radicado',
    message: `Vas a eliminar ${r.radicado} (${r.asunto.slice(0, 90)}${r.asunto.length > 90 ? '…' : ''}). Esto lo borra para todo el equipo y no se puede deshacer.`,
    label: 'Eliminar', danger: true,
    onOk: () => runWrite(() => col().doc(r.id).delete(), `Radicado ${r.radicado} eliminado.`)
  });

  // Importación inicial: solo aparece cuando la base compartida está vacía
  const importInitial = async () => {
    if (!dbRef.current) { notify('Sin acceso a los datos compartidos.', 'warn'); return; }
    try {
      const res = await fetch(IMPORT_FILE);
      const list = await res.json();
      for (let i = 0; i < list.length; i++) {
        await col().doc(list[i].id).set(list[i]);
        if (i % 25 === 0) setImporting({ done: i, total: list.length });
      }
      setImporting(null);
      notify(`Importados ${list.length} registros para el equipo.`);
    } catch (e) {
      setImporting(null);
      notify('No se pudo importar. Revisa tu rol de edición e intenta de nuevo.', 'warn');
    }
  };

  const exportCSV = async () => {
    const csv = buildCSV(filtered);
    let dl = null;
    try { dl = window.claude && window.claude.use ? await window.claude.use('downloads') : null; } catch (e) { dl = null; }
    if (dl) {
      try { await dl.save({ filename: `correspondencia_${today}.csv`, data: csv }); notify(`Exportadas ${filtered.length} filas.`); return; }
      catch (e) { /* si no se guarda, mostramos el respaldo */ }
    }
    setCsvData({ csv, count: filtered.length });
  };

  const copyCSV = async () => {
    try { await navigator.clipboard.writeText(csvData.csv); notify('CSV copiado. Pégalo en un archivo .csv o en Excel.'); }
    catch (e) { notify('No se pudo copiar. Selecciona el texto y cópialo manualmente.', 'warn'); }
  };

  const sortHead = (key, label, extra = '') => {
    const active = sort.key === key;
    return (
      <th scope="col" aria-sort={active ? (sort.dir === 'asc' ? 'ascending' : 'descending') : 'none'}
          className={`whitespace-nowrap px-3 py-3 text-left text-xs font-semibold uppercase tracking-wide text-slate-600 ${extra}`}>
        <button type="button" onClick={() => toggleSort(key)}
                className={`inline-flex items-center gap-1 hover:text-slate-900 ${active ? 'text-slate-900' : ''}`}>
          {label}
          <span aria-hidden="true" className="text-[10px]">{active ? (sort.dir === 'asc' ? '▲' : '▼') : '↕'}</span>
        </button>
      </th>
    );
  };
  const plainHead = (label, extra = '') => (
    <th scope="col" className={`whitespace-nowrap px-3 py-3 text-left text-xs font-semibold uppercase tracking-wide text-slate-600 ${extra}`}>{label}</th>
  );

  const kpis = [
    { label: 'Total de radicados', value: stats.total, note: 'Histórico completo', accent: 'border-l-[#1f4e79]' },
    { label: 'Pendientes', value: stats.pend, note: 'Sin iniciar trámite', accent: 'border-l-amber-500' },
    { label: 'En trámite', value: stats.tram, note: 'En gestión activa', accent: 'border-l-blue-500' },
    { label: 'Vencidas / próximas', value: stats.venc, note: `${stats.prox} vencen en los próximos 3 días`, accent: 'border-l-rose-600', alert: stats.venc > 0 }
  ];

  const emptyMsg = dbState === 'loading' ? 'Cargando registros compartidos…'
    : dbState === 'off' ? 'Sin acceso a los datos compartidos. Abre esta página desde claude.ai con tu sesión iniciada.'
    : records.length === 0 ? 'Todavía no hay comunicaciones registradas.'
    : 'No hay comunicaciones que coincidan con la búsqueda o los filtros.';

  return (
    <div className="min-h-full">
      <header className="bg-[#16324f] text-white">
        <div className="mx-auto flex max-w-7xl flex-col gap-4 px-4 py-6 sm:px-6 lg:flex-row lg:items-center lg:justify-between lg:px-8">
          <div>
            <p className="text-xs font-medium uppercase tracking-[0.18em] text-sky-200">Proyecto RSO II · Grupo 2</p>
            <h1 className="mt-1 text-2xl font-semibold tracking-tight">Control de correspondencia</h1>
            <p className="mt-1 text-sm text-slate-300">Radicación, registro y seguimiento de comunicaciones recibidas, enviadas e internas.</p>
          </div>
          <div className="flex flex-wrap gap-3">
            <button type="button" onClick={exportCSV}
                    className="rounded-md border border-white/30 px-4 py-2 text-sm font-medium text-white hover:bg-white/10">
              Exportar CSV
            </button>
            <button type="button" onClick={() => setFormRec(null)} disabled={dbState !== 'ready'}
                    className="rounded-md bg-white px-4 py-2 text-sm font-semibold text-[#16324f] shadow-sm hover:bg-sky-50 disabled:opacity-50">
              + Nueva radicación
            </button>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-7xl space-y-6 px-4 py-6 sm:px-6 lg:px-8">
        {dbState === 'off' && (
          <div role="alert" className="rounded-lg border border-amber-300 bg-amber-50 px-4 py-3 text-sm text-amber-900">
            Esta vista no tiene acceso a los datos compartidos del equipo. Abre la página desde claude.ai con tu sesión iniciada. Mientras tanto no se pueden guardar cambios.
          </div>
        )}

        {dbState === 'ready' && records.length === 0 && (
          <div className="flex flex-col gap-3 rounded-lg border border-sky-200 bg-sky-50 px-4 py-4 text-sm text-sky-900 sm:flex-row sm:items-center sm:justify-between">
            <span>La base compartida está vacía. Impórtala una sola vez con los 301 radicados del Excel.</span>
            <button type="button" onClick={importInitial} disabled={!!importing}
                    className="self-start rounded-md bg-[#1f4e79] px-4 py-2 font-semibold text-white hover:bg-[#173b5c] disabled:opacity-60 sm:self-auto">
              {importing ? `Importando ${importing.done}/${importing.total}…` : 'Importar datos del Excel'}
            </button>
          </div>
        )}

        <section className="grid grid-cols-2 gap-4 lg:grid-cols-4" aria-label="Resumen">
          {kpis.map(k => (
            <div key={k.label} className={`rounded-xl border border-slate-200 border-l-4 bg-white p-5 shadow-sm ${k.accent}`}>
              <p className="text-sm font-medium text-slate-600">{k.label}</p>
              <p className={`mt-2 text-3xl font-semibold tabular-nums ${k.alert ? 'text-rose-700' : 'text-slate-900'}`}>{k.value}</p>
              <p className="mt-1 text-xs text-slate-500">{k.note}</p>
            </div>
          ))}
        </section>

        <section className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
          <div className="flex flex-col gap-3 xl:flex-row xl:items-center">
            <input type="search" value={q} onChange={e => setQ(e.target.value)}
                   placeholder="Buscar por radicado, asunto o remitente / destinatario…"
                   aria-label="Buscar comunicaciones"
                   className={`${inputCls} xl:flex-1`} />
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-3 xl:w-auto xl:flex-none">
              <select value={fTipo} onChange={e => setFTipo(e.target.value)} aria-label="Filtrar por tipo" className={inputCls}>
                <option value="">Todos los tipos</option>
                {TIPOS.map(t => <option key={t.v} value={t.v}>{t.l}</option>)}
              </select>
              <select value={fEstado} onChange={e => setFEstado(e.target.value)} aria-label="Filtrar por estado" className={inputCls}>
                <option value="">Todos los estados</option>
                {ESTADOS.map(s => <option key={s.v} value={s.v}>{s.v}</option>)}
              </select>
              <select value={fDep} onChange={e => setFDep(e.target.value)} aria-label="Filtrar por dependencia" className={inputCls}>
                <option value="">Todas las dependencias</option>
                {deps.map(d => <option key={d} value={d}>{d}</option>)}
              </select>
            </div>
          </div>
          <div className="mt-3 flex flex-wrap items-center justify-between gap-2 text-sm text-slate-600">
            <span>Mostrando <strong className="text-slate-900">{filtered.length}</strong> de {records.length} comunicaciones</span>
            {hasFilters && (
              <button type="button" onClick={() => { setQ(''); setFTipo(''); setFEstado(''); setFDep(''); }}
                      className="font-medium text-[#1f4e79] hover:underline">Limpiar filtros</button>
            )}
          </div>
        </section>

        <section className="overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm">
          <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-slate-200 text-sm">
              <thead className="bg-slate-50">
                <tr>
                  {sortHead('radicado', 'Radicado')}
                  {sortHead('fechaRad', 'Fecha y hora')}
                  {plainHead('Tipo')}
                  {plainHead('Remitente / Destinatario', 'min-w-[200px]')}
                  {plainHead('Asunto', 'min-w-[280px]')}
                  {plainHead('Medio')}
                  {plainHead('Dependencia / Responsable')}
                  {sortHead('fechaLimite', 'Fecha límite')}
                  {plainHead('Estado')}
                  {plainHead('Acciones', 'text-right')}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {pageRows.length === 0 && (
                  <tr><td colSpan={10} className="px-4 py-12 text-center text-slate-500">{emptyMsg}</td></tr>
                )}
                {pageRows.map(r => {
                  const t = tipoOf(r.tipo);
                  const v = vencimiento(r, today);
                  return (
                    <tr key={r.id} className="align-top hover:bg-slate-50/70">
                      <td className="whitespace-nowrap px-3 py-3 font-mono text-xs font-semibold text-slate-900">{r.radicado}</td>
                      <td className="whitespace-nowrap px-3 py-3 text-slate-700">
                        <div className="tabular-nums">{fmtDate(r.fechaRad)}</div>
                        {r.hora && <div className="text-xs text-slate-500 tabular-nums">{r.hora}</div>}
                      </td>
                      <td className="px-3 py-3"><Pill cls={t.cls}>{t.v}</Pill></td>
                      <td className="px-3 py-3 font-medium text-slate-800">{r.contraparte}</td>
                      <td className="max-w-[340px] px-3 py-3">
                        <p className="line-clamp-2 text-slate-700" title={r.asunto}>{r.asunto}</p>
                      </td>
                      <td className="whitespace-nowrap px-3 py-3 text-slate-700">{r.medio}</td>
                      <td className="px-3 py-3">
                        <div className="font-medium text-slate-800">{r.dependencia}</div>
                        <div className="text-xs text-slate-500">{r.responsable || 'Sin responsable'}</div>
                      </td>
                      <td className="whitespace-nowrap px-3 py-3">
                        <div className={`tabular-nums ${v === 'vencida' ? 'font-semibold text-rose-700' : 'text-slate-700'}`}>{fmtDate(r.fechaLimite)}</div>
                        {v === 'vencida' && <span className="text-xs font-medium text-rose-700">Vencida</span>}
                        {v === 'proxima' && <span className="text-xs font-medium text-amber-700">Vence pronto</span>}
                      </td>
                      <td className="px-3 py-3">
                        <label className="sr-only" htmlFor={`estado-${r.id}`}>Estado de {r.radicado}</label>
                        <select id={`estado-${r.id}`} value={r.estado} onChange={e => changeEstado(r.id, e.target.value)}
                                className={`cursor-pointer rounded-full px-2.5 py-1 text-xs font-medium ring-1 ring-inset focus:outline-none focus:ring-2 ${estadoCls(r.estado)}`}>
                          {ESTADOS.map(s => <option key={s.v} value={s.v}>{s.v}</option>)}
                        </select>
                      </td>
                      <td className="whitespace-nowrap px-3 py-3 text-right">
                        <button type="button" onClick={() => setFormRec(r)}
                                className="rounded px-2 py-1 font-medium text-[#1f4e79] hover:bg-sky-50">Editar</button>
                        <button type="button" onClick={() => askDelete(r)}
                                className="rounded px-2 py-1 font-medium text-rose-700 hover:bg-rose-50">Eliminar</button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div className="flex flex-col gap-3 border-t border-slate-200 bg-slate-50 px-4 py-3 text-sm text-slate-600 sm:flex-row sm:items-center sm:justify-between">
            <span>Página {current} de {totalPages}</span>
            <div className="flex gap-2">
              <button type="button" disabled={current <= 1} onClick={() => setPage(current - 1)}
                      className="rounded-md border border-slate-300 bg-white px-3 py-1.5 font-medium hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-50">Anterior</button>
              <button type="button" disabled={current >= totalPages} onClick={() => setPage(current + 1)}
                      className="rounded-md border border-slate-300 bg-white px-3 py-1.5 font-medium hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-50">Siguiente</button>
            </div>
          </div>
        </section>

        <footer className="pb-6 text-xs text-slate-500">
          Los datos son compartidos: cada cambio lo ven en vivo todas las personas que tengan acceso a esta página.
        </footer>
      </main>

      {formRec !== undefined && (
        <Modal title={formRec ? `Editar ${formRec.radicado}` : 'Nueva radicación'} wide onClose={() => setFormRec(undefined)}>
          <RecordForm record={formRec} records={records} deps={deps}
                      onSave={saveRecord} onCancel={() => setFormRec(undefined)} />
        </Modal>
      )}

      {confirm && (
        <Modal title={confirm.title} onClose={() => setConfirm(null)}>
          <p className="text-sm text-slate-700">{confirm.message}</p>
          <div className="mt-6 flex justify-end gap-3">
            <button type="button" onClick={() => setConfirm(null)} className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50">Cancelar</button>
            <button type="button" onClick={() => { const fn = confirm.onOk; setConfirm(null); fn(); }}
                    className={`rounded-md px-4 py-2 text-sm font-semibold text-white ${confirm.danger ? 'bg-rose-700 hover:bg-rose-800' : 'bg-[#1f4e79] hover:bg-[#173b5c]'}`}>
              {confirm.label}
            </button>
          </div>
        </Modal>
      )}

      {csvData && (
        <Modal title="Exportar CSV" wide onClose={() => setCsvData(null)}>
          <p className="text-sm text-slate-700">
            Copia el contenido de abajo ({csvData.count} filas) y pégalo en un archivo con extensión .csv o directamente en Excel.
          </p>
          <textarea readOnly value={csvData.csv} rows={10}
                    className="mt-4 w-full rounded-md border border-slate-300 bg-slate-50 p-3 font-mono text-xs text-slate-700" />
          <div className="mt-4 flex justify-end gap-3">
            <button type="button" onClick={() => setCsvData(null)} className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50">Cerrar</button>
            <button type="button" onClick={copyCSV} className="rounded-md bg-[#1f4e79] px-4 py-2 text-sm font-semibold text-white hover:bg-[#173b5c]">Copiar CSV</button>
          </div>
        </Modal>
      )}

      {toast && (
        <div role="status" aria-live="polite"
             className={`fixed bottom-4 right-4 z-50 max-w-sm rounded-lg px-4 py-3 text-sm font-medium text-white shadow-lg ${toast.type === 'warn' ? 'bg-amber-700' : 'bg-slate-900'}`}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}

ReactDOM.createRoot(document.getElementById('root')).render(<App />);
</script>
</body>
</html>
