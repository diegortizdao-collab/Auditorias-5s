import { api } from './api.js';
import { toast, loadingHtml, emptyStateHtml, esc, safeRun } from './ui.js';
import { generateRadarSVG, generateEvolucionSVG, generateComparativoSVG } from './charts.js';
import { descargarInforme } from './xlsx-export.js';

const root = document.getElementById('view-root');
const nav = document.getElementById('app-nav');

const TIPO_LABEL = { planta: 'Planta', panol: 'Pañol', oficina: 'Oficina' };
const SCORES = [0, 1, 3, 5];

let catalogosCache = null;
async function getCatalogos() {
  if (!catalogosCache) catalogosCache = await api.catalogos();
  return catalogosCache;
}

function hoyISO() {
  return new Date().toISOString().slice(0, 10);
}

function setNav(crumbs) {
  nav.innerHTML = crumbs.map((c, i) => {
    const isLast = i === crumbs.length - 1;
    if (isLast) return `<span class="crumb current">${esc(c.label)}</span>`;
    return `<a href="#${c.href}">${esc(c.label)}</a><span class="sep">/</span>`;
  }).join('');
}

// Filtros de "mismo sector" para buscar acciones pendientes de auditorías
// anteriores — misma combinación planta+sector (Planta) o área+sector
// (Pañol/Oficina) que la auditoría dada.
function filtrosSector(audit) {
  return audit.audit_type === 'planta'
    ? { audit_type: 'planta', planta_id: audit.planta_id, uet_sector_id: audit.uet_sector_id }
    : { audit_type: audit.audit_type, area_id: audit.area_id, sector_id: audit.sector_id };
}

async function accionesPendientesDelSector(audit) {
  const filtros = filtrosSector(audit);
  if (audit.audit_type === 'planta' && !filtros.uet_sector_id) return [];
  if (audit.audit_type !== 'planta' && !filtros.sector_id) return [];
  return api.listarPlanAccion({ ...filtros, pendientes: '1', exclude_audit_id: audit.id });
}

function resolveNombres(audit, catalogos) {
  const planta = catalogos.plantas.find((p) => p.id === audit.planta_id);
  const uet = planta && planta.uets.find((u) => u.id === audit.uet_id);
  const uetSector = planta && planta.uet_sectores.find((s) => s.id === audit.uet_sector_id);
  const area = catalogos.areas.find((a) => a.id === audit.area_id);
  const sector = area && area.sectores.find((s) => s.id === audit.sector_id);
  return {
    ...audit,
    planta_nombre: planta ? planta.nombre : '',
    uet_nombre: uet ? uet.nombre : '',
    uet_sector_nombre: uetSector ? uetSector.nombre : '',
    area_nombre: area ? area.nombre : '',
    sector_nombre: sector ? sector.nombre : '',
  };
}

// ===========================================================================
// Router
// ===========================================================================
async function router() {
  const raw = window.location.hash.slice(1) || '/';
  const [hashPath, hashQuery] = raw.split('?');
  const parts = (hashPath || '/').split('/').filter(Boolean);
  const routeParams = new URLSearchParams(hashQuery || '');

  try {
    if (parts.length === 0) return viewSeleccion();
    if (parts[0] === 'dashboard') return viewDashboard();
    if (parts[0] === 'plan-accion') return viewPlanAccion(routeParams);
    if (parts[0] === 'auditoria' && parts[2] === 'revision') return viewRevision(parts[1]);
    if (parts[0] === 'auditoria' && parts[2] === 'formulario') return viewFormulario(parts[1]);
    if (parts[0] === 'auditoria' && parts[2] === 'evaluacion') return viewEvaluacion(parts[1]);
    if (parts[0] === 'auditoria' && parts[2] === 'informe') return viewInforme(parts[1]);
    return viewSeleccion();
  } catch (err) {
    console.error(err);
    root.innerHTML = emptyStateHtml('⚠️', 'No se pudo cargar', esc(err.message));
  }
}
window.addEventListener('hashchange', router);

// ===========================================================================
// Screen 1 · Selección de auditoría
// ===========================================================================
async function viewSeleccion() {
  setNav([{ label: 'Auditorías 5S', href: '/' }]);
  root.innerHTML = loadingHtml();

  const [catalogos, recientes] = await Promise.all([getCatalogos(), api.listarAuditorias({ limit: '10' })]);

  const state = { tipo: 'planta', areaId: null, sectorId: null, plantaId: null, uetId: null, uetSectorId: null };
  const plantaSuipacha = catalogos.plantas.find((p) => p.nombre === 'Suipacha') || catalogos.plantas[0];
  if (plantaSuipacha) state.plantaId = plantaSuipacha.id;

  function render() {
    root.innerHTML = `
      <div class="card">
        <div class="card-head">
          <h2>Nuevo registro</h2>
          <div class="sub">Elegí el tipo de auditoría y dónde se realiza</div>
        </div>
        <div class="card-body">
          <div class="tipo-block" style="margin-top:0;">
            <div class="tipo-block-head"><label>Tipo de auditoría</label></div>
            <div class="tipo-toggle" id="tipo-toggle">
              ${['panol', 'planta', 'oficina'].map((t) => `<button data-tipo="${t}" class="${state.tipo === t ? 'on' : 'off'}">${TIPO_LABEL[t]}</button>`).join('')}
            </div>
          </div>

          ${state.tipo === 'planta' ? `
            <div class="form-grid" style="margin-top:16px;">
              <div class="field">
                <label>Planta</label>
                <select id="sel-planta">
                  ${catalogos.plantas.map((p) => `<option value="${p.id}" ${p.id === state.plantaId ? 'selected' : ''}>${esc(p.nombre)}</option>`).join('')}
                </select>
              </div>
              <div class="field">
                <label>UET (opcional)</label>
                <select id="sel-uet">
                  <option value="">No aplica / sin definir</option>
                  ${(catalogos.plantas.find((p) => p.id === state.plantaId)?.uets || []).map((u) => `<option value="${u.id}" ${u.id === state.uetId ? 'selected' : ''}>${esc(u.nombre)}</option>`).join('')}
                </select>
              </div>
              <div class="field">
                <label>Sector</label>
                <select id="sel-uet-sector">
                  <option value="">Seleccionar</option>
                  ${(catalogos.plantas.find((p) => p.id === state.plantaId)?.uet_sectores || []).map((s) => `<option value="${s.id}" ${s.id === state.uetSectorId ? 'selected' : ''}>${esc(s.nombre)}</option>`).join('')}
                </select>
              </div>
            </div>
          ` : `
            <div class="area-grid">
              ${catalogos.areas.map((area) => `
                <div class="area-card ${state.areaId === area.id ? 'sel' : ''}" data-area="${area.id}">
                  <h3>${esc(area.nombre)}</h3>
                  <div class="sectores">
                    ${area.sectores.map((s) => `<span class="sector-chip ${state.sectorId === s.id ? 'sel' : ''}" data-area="${area.id}" data-sector="${s.id}">${esc(s.nombre)}</span>`).join('')}
                  </div>
                </div>
              `).join('')}
            </div>
          `}

          <div class="new-record-row">
            <button class="new-record-btn" id="btn-nuevo">+ Nuevo registro</button>
            <div class="hint" id="nuevo-hint">Elegí ${state.tipo === 'planta' ? 'planta y sector (la UET es opcional, ej. para Mantenimiento)' : 'un área y un sector'} para habilitar</div>
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card-head">
          <h2>Auditorías recientes</h2>
          <div class="btn-row">
            <a href="#/plan-accion" class="btn btn-secondary" style="text-decoration:none;">📋 Plan de Acción</a>
            <a href="#/dashboard" class="btn btn-secondary" style="text-decoration:none;">Ver Dashboard →</a>
          </div>
        </div>
        <div class="card-body" style="padding:0;">
          ${recientes.length ? `
            <table class="history-table">
              <thead><tr><th>Fecha</th><th>Tipo</th><th>Ubicación</th><th>Puntaje</th><th>Estado</th><th></th></tr></thead>
              <tbody>
                ${recientes.map((a) => renderHistRow(a, catalogos)).join('')}
              </tbody>
            </table>
          ` : `<div style="padding:22px;">${emptyStateHtml('📋', 'Todavía no hay auditorías cargadas', 'Creá la primera arriba.')}</div>`}
        </div>
      </div>
    `;

    document.querySelectorAll('.history-table tr[data-id]').forEach((tr) => {
      tr.addEventListener('click', () => {
        const id = tr.dataset.id;
        const estado = tr.dataset.estado;
        window.location.hash = estado === 'cerrada' ? `#/auditoria/${id}/informe` : `#/auditoria/${id}/evaluacion`;
      });
    });

    document.querySelectorAll('.btn-borrar-auditoria').forEach((btn) => {
      btn.addEventListener('click', async (e) => {
        e.stopPropagation();
        const id = Number(btn.dataset.id);
        const fecha = btn.dataset.fecha;
        if (!window.confirm(`¿Borrar la auditoría del ${fecha}? Esta acción no se puede deshacer (se pierden sus puntajes, fotos y acciones asociadas).`)) return;
        await safeRun(async () => {
          await api.borrarAuditoria(id);
          toast('Auditoría borrada');
          viewSeleccion();
        }, 'No se pudo borrar la auditoría');
      });
    });

    document.querySelectorAll('#tipo-toggle button').forEach((btn) => {
      btn.addEventListener('click', () => { state.tipo = btn.dataset.tipo; render(); });
    });

    if (state.tipo === 'planta') {
      document.getElementById('sel-planta').addEventListener('change', (e) => {
        state.plantaId = Number(e.target.value); state.uetId = null; state.uetSectorId = null; render();
      });
      document.getElementById('sel-uet').addEventListener('change', (e) => { state.uetId = Number(e.target.value) || null; });
      document.getElementById('sel-uet-sector').addEventListener('change', (e) => { state.uetSectorId = Number(e.target.value) || null; });
    } else {
      document.querySelectorAll('.area-card').forEach((card) => {
        card.addEventListener('click', (e) => {
          if (e.target.classList.contains('sector-chip')) return;
          state.areaId = Number(card.dataset.area);
          const first = card.querySelector('.sector-chip');
          if (first) state.sectorId = Number(first.dataset.sector);
          render();
        });
      });
      document.querySelectorAll('.sector-chip').forEach((chip) => {
        chip.addEventListener('click', () => {
          state.areaId = Number(chip.dataset.area);
          state.sectorId = Number(chip.dataset.sector);
          render();
        });
      });
    }

    document.getElementById('btn-nuevo').addEventListener('click', () => crearRegistro(state));
  }

  render();
}

function renderHistRow(a, catalogos) {
  let ubicacion = '—';
  if (a.audit_type === 'planta') {
    const planta = catalogos.plantas.find((p) => p.id === a.planta_id);
    const sector = planta && planta.uet_sectores.find((s) => s.id === a.uet_sector_id);
    ubicacion = [planta?.nombre, sector?.nombre].filter(Boolean).join(' · ') || '—';
  } else {
    const area = catalogos.areas.find((ar) => ar.id === a.area_id);
    const sector = area && area.sectores.find((s) => s.id === a.sector_id);
    ubicacion = [area?.nombre, sector?.nombre].filter(Boolean).join(' · ') || '—';
  }
  const estadoLbl = a.estado === 'cerrada' ? 'Cerrada' : 'En progreso';
  return `<tr data-id="${a.id}" data-estado="${a.estado}">
    <td>${a.fecha}</td>
    <td>${TIPO_LABEL[a.audit_type]}</td>
    <td>${esc(ubicacion)}</td>
    <td>${a.puntaje_total ?? '—'}${a.puntaje_objetivo ? ` / obj. ${a.puntaje_objetivo}` : ''}</td>
    <td>${estadoLbl}</td>
    <td><button class="btn-icon btn-borrar-auditoria" data-id="${a.id}" data-fecha="${a.fecha}" title="Borrar esta auditoría">🗑</button></td>
  </tr>`;
}

async function crearRegistro(state) {
  // La UET es opcional: Mantenimiento, por ejemplo, no tiene UET propia —
  // solo se divide por planta y sector (Taller / Pañol-Depósito).
  const faltaPlanta = state.tipo === 'planta' && (!state.plantaId || !state.uetSectorId);
  const faltaArea = state.tipo !== 'planta' && (!state.areaId || !state.sectorId);
  if (faltaPlanta || faltaArea) {
    toast('Completá la ubicación antes de crear el registro', 'error');
    return;
  }
  const evaluador = window.prompt('Nombre del evaluador:', '');
  if (!evaluador) return;

  await safeRun(async () => {
    const audit = await api.crearAuditoria({
      audit_type: state.tipo,
      planta_id: state.plantaId,
      uet_id: state.uetId,
      uet_sector_id: state.uetSectorId,
      area_id: state.areaId,
      sector_id: state.sectorId,
      fecha: hoyISO(),
      evaluador,
      responsable: evaluador,
    });
    toast('Registro creado');
    // Si este sector ya tuvo acciones correctivas pendientes en auditorías
    // anteriores, primero se revisan (y se puede puntuar el ítem ahí mismo)
    // antes de pasar a la auditoría en sí.
    const pendientes = await accionesPendientesDelSector(audit).catch(() => []);
    window.location.hash = pendientes.length
      ? `#/auditoria/${audit.id}/revision`
      : `#/auditoria/${audit.id}/formulario`;
  }, 'No se pudo crear el registro');
}

// ===========================================================================
// Screen 2 · Formulario
// ===========================================================================
async function viewFormulario(auditId) {
  setNav([{ label: 'Auditorías 5S', href: '/' }, { label: 'Formulario', href: `/auditoria/${auditId}/formulario` }]);
  root.innerHTML = loadingHtml();

  const [audit, catalogos] = await Promise.all([api.obtenerAuditoria(auditId), getCatalogos()]);

  root.innerHTML = `
    <div class="card">
      <div class="card-head">
        <h2>Formulario</h2>
        <div class="sub">${TIPO_LABEL[audit.audit_type]} · encabezado de la auditoría</div>
        <a href="#/auditoria/${auditId}/revision" class="btn btn-secondary" style="text-decoration:none;">📋 Revisar plan de acción anterior</a>
      </div>
      <div class="card-body">
        <div class="form-grid">
          <div class="field"><label>Fecha</label><input type="date" id="f-fecha" value="${(audit.fecha || '').slice(0, 10)}"></div>
          <div class="field"><label>Turno</label>
            <select id="f-turno">
              ${['Mañana', 'Tarde', 'Noche'].map((t) => `<option ${audit.turno === t ? 'selected' : ''}>${t}</option>`).join('')}
            </select>
          </div>
          <div class="field"><label>Evaluador</label><input type="text" id="f-evaluador" value="${esc(audit.evaluador || '')}"></div>
          <div class="field"><label>Supervisor / Líder</label><input type="text" id="f-supervisor" value="${esc(audit.supervisor_lider || '')}"></div>
          <div class="field"><label>Jefe / Gerente</label><input type="text" id="f-jefe" value="${esc(audit.jefe_gerente || '')}"></div>
          <div class="field"><label>Responsable (informe)</label><input type="text" id="f-responsable" value="${esc(audit.responsable || '')}"></div>
        </div>
        <div class="btn-row" style="margin-top:22px;">
          <button class="btn btn-primary" id="btn-guardar">Guardar</button>
          <button class="btn btn-secondary" id="btn-continuar">Continuar a Evaluación →</button>
        </div>
      </div>
    </div>
  `;

  async function guardar() {
    return safeRun(() => api.actualizarAuditoria(auditId, {
      fecha: document.getElementById('f-fecha').value,
      turno: document.getElementById('f-turno').value,
      evaluador: document.getElementById('f-evaluador').value,
      supervisor_lider: document.getElementById('f-supervisor').value,
      jefe_gerente: document.getElementById('f-jefe').value,
      responsable: document.getElementById('f-responsable').value,
    }), 'No se pudo guardar el formulario');
  }

  document.getElementById('btn-guardar').addEventListener('click', async () => { await guardar(); toast('Guardado'); });
  document.getElementById('btn-continuar').addEventListener('click', async () => {
    await guardar();
    window.location.hash = `#/auditoria/${auditId}/evaluacion`;
  });
}

// ===========================================================================
// Screen · Revisión de Plan de Acción (al abrir una auditoría nueva sobre
// un sector que ya tuvo acciones correctivas pendientes)
// ===========================================================================
async function viewRevision(auditId) {
  setNav([{ label: 'Auditorías 5S', href: '/' }, { label: 'Revisión', href: `/auditoria/${auditId}/revision` }]);
  root.innerHTML = loadingHtml();

  const audit = await api.obtenerAuditoria(auditId);
  const pendientes = await accionesPendientesDelSector(audit);

  function findItem(itemId) {
    for (const cat of audit.categorias) {
      const item = cat.items.find((i) => i.rubric_item_id === itemId);
      if (item) return item;
    }
    return null;
  }
  function catTag(itemId) {
    for (const cat of audit.categorias) {
      if (cat.items.some((i) => i.rubric_item_id === itemId)) return cat.tag;
    }
    return '';
  }

  // Agrupa las acciones pendientes por ítem de rúbrica: varias auditorías
  // anteriores pueden haber dejado más de una acción sobre el mismo ítem.
  const grupos = new Map();
  const sinItem = [];
  for (const a of pendientes) {
    if (!a.rubric_item_id) { sinItem.push(a); continue; }
    if (!grupos.has(a.rubric_item_id)) grupos.set(a.rubric_item_id, { rubric_item_id: a.rubric_item_id, item_nombre: a.item_nombre, acciones: [] });
    grupos.get(a.rubric_item_id).acciones.push(a);
  }

  function renderAccionChipRevision(a) {
    const origenLabel = a.origen === 'ia' ? '🤖 IA' : '✋ Manual';
    return `
      <div class="accion-chip ${a.estado}" data-accion="${a.id}">
        <span class="accion-origen ${a.origen}">${origenLabel}</span>
        <div class="accion-body">
          <div class="accion-desc">${esc(a.descripcion)}</div>
          <div class="accion-meta">
            <span>🗓 auditoría del ${a.audit_fecha}</span>
            ${a.responsable ? `<span>👤 ${esc(a.responsable)}</span>` : ''}
            ${a.fecha_compromiso ? `<span>📅 vencía ${a.fecha_compromiso}</span>` : ''}
          </div>
        </div>
        <div class="accion-controls">
          ${a.estado === 'cerrada'
            ? `<span class="status-pill good">✔ Cumplida</span>`
            : `<button class="btn btn-secondary btn-cumplida" data-accion="${a.id}">✓ Marcar cumplida</button>`}
        </div>
      </div>
    `;
  }

  function renderGrupo(g) {
    const item = findItem(g.rubric_item_id);
    const tag = catTag(g.rubric_item_id);
    return `
      <div class="s-block">
        <div class="s-block-head">
          ${tag ? `<span class="s-tag">${tag}</span>` : ''}
          <span class="name">${esc(g.item_nombre || item?.nombre || 'Ítem')}</span>
        </div>
        <div class="revision-body">
          <div class="acciones-block" style="border-top:none; padding-top:0;">
            ${g.acciones.map(renderAccionChipRevision).join('')}
          </div>
          ${item ? `
            <div class="revision-score">
              <span class="lbl">Puntaje de hoy para este ítem</span>
              <div class="scale">
                ${SCORES.map((s) => `<button data-item="${item.rubric_item_id}" data-score="${s}" class="${item.score === s ? 'sel' : ''}">${s}</button>`).join('')}
              </div>
            </div>
          ` : ''}
        </div>
      </div>
    `;
  }

  function render() {
    const gruposArr = Array.from(grupos.values());
    const sinAcciones = gruposArr.length === 0 && sinItem.length === 0;
    root.innerHTML = `
      <div class="card">
        <div class="card-head">
          <h2>Revisión de Plan de Acción</h2>
          <div class="sub">${TIPO_LABEL[audit.audit_type]} · antes de auditar de nuevo, revisá si se cumplieron las acciones de la vez pasada</div>
        </div>
        <div class="card-body">
          ${sinAcciones ? emptyStateHtml('✅', 'No hay acciones pendientes de auditorías anteriores en este sector', 'Podés pasar directo al formulario.') : `
            <div class="lead" style="margin-bottom:18px;">Si la acción ya se cumplió, marcala como <b>Cumplida</b> y aprovechá para puntuar el ítem acá mismo — ese puntaje queda cargado en la auditoría de hoy.</div>
            ${gruposArr.map(renderGrupo).join('')}
            ${sinItem.length ? `
              <div class="s-block">
                <div class="s-block-head"><span class="name">Otras acciones pendientes (sin ítem asociado)</span></div>
                <div class="revision-body">
                  <div class="acciones-block" style="border-top:none; padding-top:0;">
                    ${sinItem.map(renderAccionChipRevision).join('')}
                  </div>
                </div>
              </div>
            ` : ''}
          `}
          <div class="btn-row" style="margin-top:18px;">
            <button class="btn btn-secondary" id="btn-volver">← Auditorías 5S</button>
            <button class="btn btn-primary" id="btn-continuar">Continuar al Formulario →</button>
          </div>
        </div>
      </div>
    `;
    wireHandlers();
  }

  function wireHandlers() {
    document.querySelectorAll('.btn-cumplida').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const accionId = Number(btn.dataset.accion);
        await safeRun(async () => {
          const actualizada = await api.actualizarAccion(accionId, { estado: 'cerrada', fecha_cierre: hoyISO() });
          for (const g of grupos.values()) {
            const idx = g.acciones.findIndex((a) => a.id === accionId);
            if (idx >= 0) g.acciones[idx] = actualizada;
          }
          const idxSin = sinItem.findIndex((a) => a.id === accionId);
          if (idxSin >= 0) sinItem[idxSin] = actualizada;
          toast('Acción marcada como cumplida');
          render();
        }, 'No se pudo actualizar la acción');
      });
    });

    document.querySelectorAll('.revision-score .scale button').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const itemId = Number(btn.dataset.item);
        const score = Number(btn.dataset.score);
        const item = findItem(itemId);
        await safeRun(async () => {
          await api.guardarPuntaje(auditId, itemId, { score, comentario: item.comentario || null });
          item.score = score;
          toast('Puntaje guardado');
          render();
        }, 'No se pudo guardar el puntaje');
      });
    });

    document.getElementById('btn-volver').addEventListener('click', () => { window.location.hash = '/'; });
    document.getElementById('btn-continuar').addEventListener('click', () => {
      window.location.hash = `#/auditoria/${auditId}/formulario`;
    });
  }

  await safeRun(render, 'No se pudo cargar la revisión del plan de acción');
}

// ===========================================================================
// Screen 3 · Evaluación
// ===========================================================================
async function viewEvaluacion(auditId) {
  setNav([{ label: 'Auditorías 5S', href: '/' }, { label: 'Evaluación', href: `/auditoria/${auditId}/evaluacion` }]);
  root.innerHTML = loadingHtml();

  const audit = await api.obtenerAuditoria(auditId);
  // Qué ítems tienen desplegados los 4 criterios (0/1/3/5) — solo estado de
  // pantalla, no se guarda en el servidor. Por defecto colapsado para no
  // alargar la lista, pero el auditor puede abrirlo para ver qué distingue
  // a cada puntaje antes de elegir uno.
  const criteriosAbiertos = new Set();

  function subtotal(cat) {
    return cat.items.reduce((s, it) => s + (it.score ?? 0), 0);
  }
  function totalActual() {
    return audit.categorias.reduce((s, c) => s + subtotal(c), 0);
  }

  function updateAccionLocal(accion) {
    for (const cat of audit.categorias) {
      for (const item of cat.items) {
        const idx = (item.acciones || []).findIndex((a) => a.id === accion.id);
        if (idx >= 0) { item.acciones[idx] = accion; return; }
      }
    }
  }

  function render() {
    root.innerHTML = `
      <div class="card">
        <div class="card-head">
          <h2>Evaluación</h2>
          <div class="sub">${TIPO_LABEL[audit.audit_type]} · ${audit.fecha} · Puntaje actual: <b>${totalActual()}</b>${audit.puntaje_objetivo ? ` / objetivo ${audit.puntaje_objetivo}` : ''}</div>
          <a href="#/plan-accion?audit_id=${auditId}" class="btn btn-secondary" style="text-decoration:none;">📋 Plan de Acción de esta auditoría</a>
        </div>
        <div class="card-body">
          ${audit.categorias.map((cat) => `
            <div class="s-block">
              <div class="s-block-head">
                <span class="s-tag">${cat.tag}</span>
                <span class="name">${esc(cat.label)}</span>
                <span class="desc">${esc(cat.sub)}</span>
                <span class="subtotal">Sub-total <b>${subtotal(cat)}</b> / ${cat.items.length * 5}</span>
              </div>
              ${cat.items.map((item) => renderItemRow(item)).join('')}
            </div>
          `).join('')}
          <div class="btn-row" style="margin-top:10px;">
            <button class="btn btn-secondary" id="btn-volver">← Formulario</button>
            <button class="btn btn-primary" id="btn-finalizar">Finalizar y ver Informe →</button>
          </div>
        </div>
      </div>
    `;

    document.querySelectorAll('.toggle-criterios').forEach((btn) => {
      btn.addEventListener('click', () => {
        const itemId = Number(btn.dataset.item);
        if (criteriosAbiertos.has(itemId)) criteriosAbiertos.delete(itemId);
        else criteriosAbiertos.add(itemId);
        render();
      });
    });

    document.querySelectorAll('.scale button').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const itemId = Number(btn.dataset.item);
        const score = Number(btn.dataset.score);
        const item = findItem(itemId);
        await safeRun(async () => {
          const res = await api.guardarPuntaje(auditId, itemId, { score, comentario: item.comentario || null });
          item.score = score;
          item.score_id = res.score.id;
          audit.puntaje_total = res.puntaje_total;
          // Si el puntaje quedó en 0 o 1, el backend ya generó (o reutilizó)
          // la sugerencia de la IA — se sincroniza acá para mostrarla al toque.
          if (res.accion_sugerida) {
            item.acciones = item.acciones || [];
            const idx = item.acciones.findIndex((a) => a.origen === 'ia');
            if (idx >= 0) item.acciones[idx] = res.accion_sugerida;
            else item.acciones.push(res.accion_sugerida);
          }
          render();
        }, 'No se pudo guardar el puntaje');
      });
    });

    document.querySelectorAll('.accion-estado').forEach((sel) => {
      sel.addEventListener('change', async () => {
        const accionId = Number(sel.dataset.accion);
        const estado = sel.value;
        await safeRun(async () => {
          const body = { estado };
          if (estado === 'cerrada') body.fecha_cierre = hoyISO();
          const actualizada = await api.actualizarAccion(accionId, body);
          updateAccionLocal(actualizada);
          render();
        }, 'No se pudo actualizar la acción');
      });
    });

    document.querySelectorAll('.regen-accion').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const accionId = Number(btn.dataset.accion);
        await safeRun(async () => {
          const actualizada = await api.regenerarAccion(accionId);
          updateAccionLocal(actualizada);
          toast('Sugerencia regenerada');
          render();
        }, 'No se pudo regenerar la sugerencia');
      });
    });

    document.querySelectorAll('.add-accion').forEach((btn) => {
      btn.addEventListener('click', () => {
        const form = document.querySelector(`.accion-form[data-item="${btn.dataset.item}"]`);
        if (form) form.hidden = !form.hidden;
      });
    });

    document.querySelectorAll('.af-guardar').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const itemId = Number(btn.dataset.item);
        const form = document.querySelector(`.accion-form[data-item="${itemId}"]`);
        const descripcion = form.querySelector('.af-desc').value.trim();
        if (!descripcion) { toast('Escribí una descripción para la acción', 'error'); return; }
        const responsable = form.querySelector('.af-resp').value.trim();
        const fechaCompromiso = form.querySelector('.af-fecha').value || null;
        const item = findItem(itemId);
        await safeRun(async () => {
          const creada = await api.crearAccion({
            audit_id: auditId,
            rubric_item_id: itemId,
            audit_score_id: item.score_id || null,
            descripcion,
            responsable: responsable || null,
            fecha_compromiso: fechaCompromiso,
          });
          item.acciones = item.acciones || [];
          item.acciones.push(creada);
          render();
        }, 'No se pudo crear la acción');
      });
    });

    document.querySelectorAll('.item-comentario').forEach((ta) => {
      ta.addEventListener('blur', async () => {
        const itemId = Number(ta.dataset.item);
        const item = findItem(itemId);
        if (item.score === null || item.score === undefined) return; // necesita puntaje primero
        await safeRun(() => api.guardarPuntaje(auditId, itemId, { score: item.score, comentario: ta.value }), 'No se pudo guardar el comentario');
        item.comentario = ta.value;
      });
    });

    document.querySelectorAll('.photo-chip input[type="file"]').forEach((input) => {
      input.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const itemId = Number(input.dataset.item);
        const item = findItem(itemId);
        if (item.score === null || item.score === undefined) {
          toast('Elegí un puntaje antes de agregar una foto', 'error');
          return;
        }
        await safeRun(async () => {
          const foto = await api.subirFoto(auditId, itemId, file);
          item.fotos = item.fotos || [];
          item.fotos.push(foto);
          render();
        }, 'No se pudo subir la foto');
      });
    });

    document.querySelectorAll('.photo-thumb .rm').forEach((btn) => {
      btn.addEventListener('click', async (e) => {
        e.stopPropagation();
        const fotoId = Number(btn.dataset.foto);
        const itemId = Number(btn.dataset.item);
        await safeRun(async () => {
          await api.borrarFoto(fotoId);
          const item = findItem(itemId);
          item.fotos = (item.fotos || []).filter((f) => f.id !== fotoId);
          render();
        }, 'No se pudo borrar la foto');
      });
    });

    document.getElementById('btn-volver').addEventListener('click', () => { window.location.hash = `#/auditoria/${auditId}/formulario`; });
    document.getElementById('btn-finalizar').addEventListener('click', async () => {
      await safeRun(() => api.actualizarAuditoria(auditId, { estado: 'cerrada' }), 'No se pudo finalizar');
      window.location.hash = `#/auditoria/${auditId}/informe`;
    });
  }

  function findItem(itemId) {
    for (const cat of audit.categorias) {
      const item = cat.items.find((i) => i.rubric_item_id === itemId);
      if (item) return item;
    }
    return null;
  }

  function renderAccionChip(a) {
    const origenLabel = a.origen === 'ia' ? '🤖 IA' : '✋ Manual';
    return `
      <div class="accion-chip ${a.estado}" data-accion="${a.id}">
        <span class="accion-origen ${a.origen}">${origenLabel}</span>
        <div class="accion-body">
          <div class="accion-desc">${esc(a.descripcion)}</div>
          <div class="accion-meta">
            ${a.responsable ? `<span>👤 ${esc(a.responsable)}</span>` : ''}
            ${a.fecha_compromiso ? `<span>📅 vence ${a.fecha_compromiso}</span>` : ''}
            ${a.fecha_cierre ? `<span>✔ cerrada ${a.fecha_cierre}</span>` : ''}
          </div>
        </div>
        <div class="accion-controls">
          <select class="accion-estado" data-accion="${a.id}">
            <option value="abierta" ${a.estado === 'abierta' ? 'selected' : ''}>Abierta</option>
            <option value="en_proceso" ${a.estado === 'en_proceso' ? 'selected' : ''}>En proceso</option>
            <option value="cerrada" ${a.estado === 'cerrada' ? 'selected' : ''}>Cerrada</option>
          </select>
          ${a.origen === 'ia' ? `<button class="btn-icon regen-accion" data-accion="${a.id}" title="Regenerar sugerencia de la IA">↻</button>` : ''}
        </div>
      </div>
    `;
  }

  function renderAccionesBlock(item) {
    const acciones = item.acciones || [];
    const mostrar = item.score === 0 || item.score === 1 || acciones.length > 0;
    if (!mostrar) return '';
    return `
      <div class="acciones-block" data-item="${item.rubric_item_id}">
        ${acciones.map(renderAccionChip).join('')}
        <button class="btn-link add-accion" data-item="${item.rubric_item_id}">+ Agregar acción manual</button>
        <div class="accion-form" data-item="${item.rubric_item_id}" hidden>
          <div class="field"><label>Descripción</label><textarea class="af-desc" rows="1" placeholder="¿Qué hay que hacer?"></textarea></div>
          <div class="field"><label>Responsable</label><input type="text" class="af-resp"></div>
          <div class="field"><label>Vence</label><input type="date" class="af-fecha"></div>
          <button class="btn btn-primary af-guardar" data-item="${item.rubric_item_id}">Guardar</button>
        </div>
      </div>
    `;
  }

  function renderCriteriosBlock(item) {
    if (!criteriosAbiertos.has(item.rubric_item_id)) return '';
    const filas = [
      [0, item.criterio_0], [1, item.criterio_1], [3, item.criterio_3], [5, item.criterio_5],
    ];
    return `
      <div class="criterios-block" style="grid-column:1/-1;">
        ${filas.map(([s, texto]) => `
          <div class="criterio-fila ${item.score === s ? 'sel' : ''}">
            <span class="criterio-score">${s}</span>
            <span class="criterio-texto">${esc(texto || '(sin descripción)')}</span>
          </div>
        `).join('')}
      </div>
    `;
  }

  function renderItemRow(item) {
    const fotos = item.fotos || [];
    const abierto = criteriosAbiertos.has(item.rubric_item_id);
    return `
      <div class="item-row">
        <div class="item-name">
          <button class="btn-icon toggle-criterios" data-item="${item.rubric_item_id}" title="${abierto ? 'Ocultar' : 'Ver'} qué distingue a cada puntaje">${abierto ? '▾' : 'ℹ'}</button>
          ${esc(item.nombre)}
          <textarea class="item-comentario" data-item="${item.rubric_item_id}" placeholder="Comentario (opcional)" rows="1">${esc(item.comentario || '')}</textarea>
        </div>
        <div class="scale">
          ${SCORES.map((s) => `<button data-item="${item.rubric_item_id}" data-score="${s}" class="${item.score === s ? 'sel' : ''}">${s}</button>`).join('')}
        </div>
        <label class="photo-chip">📷 foto<input type="file" accept="image/*" capture="environment" data-item="${item.rubric_item_id}"></label>
        ${renderCriteriosBlock(item)}
        ${fotos.length ? `<div class="photo-thumb-row" style="grid-column:1/-1;">
          ${fotos.map((f) => `<div class="photo-thumb"><img src="${api.urlFoto(f.url)}" alt=""><button class="rm" data-foto="${f.id}" data-item="${item.rubric_item_id}">×</button></div>`).join('')}
        </div>` : ''}
        ${renderAccionesBlock(item)}
      </div>
    `;
  }

  render();
}

// ===========================================================================
// Screen 4 · Dashboard
// ===========================================================================
async function viewDashboard() {
  setNav([{ label: 'Auditorías 5S', href: '/' }, { label: 'Dashboard', href: '/dashboard' }]);
  root.innerHTML = loadingHtml();

  const catalogos = await getCatalogos();
  const state = { tipo: 'planta', plantaId: '', uetId: '', uetSectorId: '' };

  async function render() {
    const dash = await api.dashboard({
      tipo: state.tipo,
      ...(state.plantaId ? { planta_id: state.plantaId } : {}),
      ...(state.uetId ? { uet_id: state.uetId } : {}),
      ...(state.uetSectorId ? { uet_sector_id: state.uetSectorId } : {}),
    });

    const plantaSel = catalogos.plantas.find((p) => p.id === Number(state.plantaId));

    root.innerHTML = `
      <div class="card">
        <div class="card-head">
          <h2>Dashboard</h2>
          <div class="sub">Radar, evolución y comparativo por sector</div>
          <a href="#/plan-accion" class="btn btn-secondary" style="text-decoration:none;">📋 Plan de Acción</a>
        </div>
        <div class="card-body">
          <div class="dash-filters">
            <div class="field">
              <label>Tipo</label>
              <select id="d-tipo">
                ${['planta', 'panol', 'oficina'].map((t) => `<option value="${t}" ${state.tipo === t ? 'selected' : ''}>${TIPO_LABEL[t]}</option>`).join('')}
              </select>
            </div>
            ${state.tipo === 'planta' ? `
              <div class="field">
                <label>Planta</label>
                <select id="d-planta">
                  <option value="">Todas</option>
                  ${catalogos.plantas.map((p) => `<option value="${p.id}" ${String(p.id) === state.plantaId ? 'selected' : ''}>${esc(p.nombre)}</option>`).join('')}
                </select>
              </div>
              <div class="field">
                <label>UET</label>
                <select id="d-uet">
                  <option value="">Todas</option>
                  ${(plantaSel?.uets || []).map((u) => `<option value="${u.id}" ${String(u.id) === state.uetId ? 'selected' : ''}>${esc(u.nombre)}</option>`).join('')}
                </select>
              </div>
              <div class="field">
                <label>Sector</label>
                <select id="d-sector">
                  <option value="">Todos</option>
                  ${(plantaSel?.uet_sectores || []).map((s) => `<option value="${s.id}" ${String(s.id) === state.uetSectorId ? 'selected' : ''}>${esc(s.nombre)}</option>`).join('')}
                </select>
              </div>
            ` : ''}
          </div>

          ${!dash.puntaje ? `
            <div class="dash-empty">
              <div class="ico">📊</div>
              <div class="msg">Todavía no hay auditorías de tipo ${TIPO_LABEL[state.tipo]}</div>
              <div class="sub">En cuanto se cargue la primera, este panel se completa solo.</div>
            </div>
          ` : `
            <div class="dash-grid">
              <div class="dash-card">
                <h3>Radar 5S</h3>
                <div class="cap">Última auditoría (${dash.puntaje.fecha}) · % por categoría</div>
                ${generateRadarSVG(dash.radar)}
              </div>
              <div class="dash-card">
                <h3>Puntaje total</h3>
                <div class="cap">Última auditoría vs. objetivo del mes</div>
                <div style="display:flex; align-items:baseline; gap:10px; margin-top:18px;">
                  <div style="font-family:'Barlow Condensed'; font-weight:700; font-size:64px; color:var(--brand);">${dash.puntaje.actual}</div>
                  <div style="color:var(--ink-muted); font-size:15px;">/ 100</div>
                </div>
                ${dash.puntaje.objetivo_mes_actual !== null ? statusPill(dash.puntaje.actual, dash.puntaje.objetivo_mes_actual) : ''}
                <div style="margin-top:22px; display:flex; flex-direction:column; gap:10px;">
                  <div style="display:flex; justify-content:space-between; font-size:13px;"><span style="color:var(--ink-muted)">Puntaje anterior</span><b>${dash.puntaje.anterior ?? '—'}</b></div>
                  <div style="display:flex; justify-content:space-between; font-size:13px;"><span style="color:var(--ink-muted)">Variación</span><b style="color:${(dash.puntaje.variacion || 0) >= 0 ? 'var(--good)' : 'var(--critical)'}">${dash.puntaje.variacion === null ? '—' : (dash.puntaje.variacion >= 0 ? '+' : '') + dash.puntaje.variacion}</b></div>
                  <div style="display:flex; justify-content:space-between; font-size:13px;"><span style="color:var(--ink-muted)">Responsable</span><b>${esc(dash.puntaje.responsable || '—')}</b></div>
                </div>
              </div>
              <div class="dash-card full">
                <h3>Evolución — Objetivo vs. Auditado</h3>
                <div class="cap">Año ${dash.anio}</div>
                ${dash.evolucion.every((m) => m.objetivo === null) ? emptyStateHtml('📈', 'Sin objetivo mensual definido', 'Cargá los objetivos en monthly_targets para este tipo.') : generateEvolucionSVG(dash.evolucion)}
                <div class="legend">
                  <span class="sw"><span class="ln" style="background:var(--brand);"></span> Auditado</span>
                  <span class="sw"><span class="ln" style="background:var(--ink-muted); border-top:2px dashed var(--ink-muted); height:0;"></span> Objetivo del mes</span>
                </div>
              </div>
              ${state.tipo === 'planta' ? `
                <div class="dash-card full">
                  <h3>Comparativo por sector</h3>
                  <div class="cap">Última auditoría de cada sector${dash.objetivo_mes_actual !== null ? ` vs. objetivo (${dash.objetivo_mes_actual})` : ''}</div>
                  ${dash.comparativo.length ? generateComparativoSVG(dash.comparativo, dash.objetivo_mes_actual) : emptyStateHtml('📊', 'Sin datos por sector todavía', '')}
                  <div class="legend">
                    <span class="sw"><span class="dot" style="background:var(--good);"></span> Cumple objetivo</span>
                    <span class="sw"><span class="dot" style="background:var(--warning);"></span> Cerca (dentro de 10 pts)</span>
                    <span class="sw"><span class="dot" style="background:var(--critical);"></span> Atención</span>
                  </div>
                </div>
              ` : ''}
            </div>
          `}
        </div>
      </div>
    `;

    document.getElementById('d-tipo').addEventListener('change', (e) => {
      state.tipo = e.target.value; state.plantaId = ''; state.uetId = ''; state.uetSectorId = ''; render();
    });
    const dPlanta = document.getElementById('d-planta');
    if (dPlanta) dPlanta.addEventListener('change', (e) => { state.plantaId = e.target.value; state.uetId = ''; state.uetSectorId = ''; render(); });
    const dUet = document.getElementById('d-uet');
    if (dUet) dUet.addEventListener('change', (e) => { state.uetId = e.target.value; render(); });
    const dSector = document.getElementById('d-sector');
    if (dSector) dSector.addEventListener('change', (e) => { state.uetSectorId = e.target.value; render(); });
  }

  function statusPill(actual, objetivo) {
    const diff = actual - objetivo;
    const cls = diff >= 0 ? 'good' : diff >= -10 ? 'warning' : 'critical';
    const txt = diff >= 0 ? `cumple objetivo (${objetivo})` : diff >= -10 ? `cerca del objetivo (${objetivo})` : `lejos del objetivo (${objetivo})`;
    return `<span class="status-pill ${cls}" style="margin-top:6px;">● ${txt}</span>`;
  }

  await safeRun(render, 'No se pudo cargar el dashboard');
}

// ===========================================================================
// Screen · Plan de Acción (independiente de la solapa de Evaluación)
// ===========================================================================
const ESTADO_LABEL = { abierta: 'Abierta', en_proceso: 'En proceso', cerrada: 'Cerrada' };

async function viewPlanAccion(routeParams) {
  setNav([{ label: 'Auditorías 5S', href: '/' }, { label: 'Plan de Acción', href: '/plan-accion' }]);
  root.innerHTML = loadingHtml();

  const catalogos = await getCatalogos();
  const state = {
    estado: routeParams.get('estado') || '',
    tipo: routeParams.get('tipo') || '',
    plantaId: routeParams.get('planta_id') || '',
    uetSectorId: routeParams.get('uet_sector_id') || '',
    auditId: routeParams.get('audit_id') || '',
  };

  async function render() {
    const params = {
      ...(state.estado ? { estado: state.estado } : {}),
      ...(state.tipo ? { audit_type: state.tipo } : {}),
      ...(state.plantaId ? { planta_id: state.plantaId } : {}),
      ...(state.uetSectorId ? { uet_sector_id: state.uetSectorId } : {}),
      ...(state.auditId ? { audit_id: state.auditId } : {}),
    };
    const lista = await api.listarPlanAccion(params);
    const plantaSel = catalogos.plantas.find((p) => String(p.id) === state.plantaId);

    root.innerHTML = `
      <div class="card">
        <div class="card-head">
          <h2>Plan de Acción</h2>
          <div class="sub">Acciones correctivas — sugeridas por IA y cargadas a mano, de todas las auditorías</div>
        </div>
        <div class="card-body">
          ${state.auditId ? `
            <div class="lead" style="display:flex; align-items:center; justify-content:space-between; gap:10px; margin-bottom:16px;">
              <span>Mostrando solo las acciones de la auditoría <b>#${esc(state.auditId)}</b></span>
              <button class="btn-link" id="btn-quitar-filtro-audit">Ver todas →</button>
            </div>
          ` : ''}
          <div class="plan-filters">
            <div class="field">
              <label>Estado</label>
              <select id="pf-estado">
                <option value="">Todos</option>
                ${Object.entries(ESTADO_LABEL).map(([v, l]) => `<option value="${v}" ${state.estado === v ? 'selected' : ''}>${l}</option>`).join('')}
              </select>
            </div>
            <div class="field">
              <label>Tipo</label>
              <select id="pf-tipo">
                <option value="">Todos</option>
                ${['planta', 'panol', 'oficina'].map((t) => `<option value="${t}" ${state.tipo === t ? 'selected' : ''}>${TIPO_LABEL[t]}</option>`).join('')}
              </select>
            </div>
            ${state.tipo === 'planta' ? `
              <div class="field">
                <label>Planta</label>
                <select id="pf-planta">
                  <option value="">Todas</option>
                  ${catalogos.plantas.map((p) => `<option value="${p.id}" ${state.plantaId === String(p.id) ? 'selected' : ''}>${esc(p.nombre)}</option>`).join('')}
                </select>
              </div>
              <div class="field">
                <label>Sector</label>
                <select id="pf-sector">
                  <option value="">Todos</option>
                  ${(plantaSel?.uet_sectores || []).map((s) => `<option value="${s.id}" ${state.uetSectorId === String(s.id) ? 'selected' : ''}>${esc(s.nombre)}</option>`).join('')}
                </select>
              </div>
            ` : ''}
          </div>

          <div class="plan-count">${lista.length} ${lista.length === 1 ? 'acción' : 'acciones'}</div>

          ${lista.length ? `
            <div style="overflow-x:auto;">
              <table class="plan-table">
                <thead>
                  <tr>
                    <th>Auditoría</th><th>Ítem</th><th>Origen</th><th>Descripción</th>
                    <th>Responsable</th><th>Vence</th><th>Estado</th>
                  </tr>
                </thead>
                <tbody>
                  ${lista.map((a) => `
                    <tr data-accion="${a.id}">
                      <td>
                        <a href="#/auditoria/${a.audit_id}/evaluacion" style="color:var(--brand-strong); font-weight:600; text-decoration:none;">${a.audit_fecha}</a>
                        <div class="ubic">${esc([TIPO_LABEL[a.audit_type], a.planta_nombre, a.sector_nombre].filter(Boolean).join(' · ') || '—')}</div>
                      </td>
                      <td>${esc(a.item_nombre || '—')}</td>
                      <td><span class="accion-origen ${a.origen}">${a.origen === 'ia' ? '🤖 IA' : '✋ Manual'}</span></td>
                      <td class="desc-cell">${esc(a.descripcion)}</td>
                      <td>${esc(a.responsable || '—')}</td>
                      <td>${a.fecha_compromiso || '—'}</td>
                      <td>
                        <select class="accion-estado" data-accion="${a.id}">
                          ${Object.entries(ESTADO_LABEL).map(([v, l]) => `<option value="${v}" ${a.estado === v ? 'selected' : ''}>${l}</option>`).join('')}
                        </select>
                      </td>
                    </tr>
                  `).join('')}
                </tbody>
              </table>
            </div>
          ` : emptyStateHtml('✅', 'No hay acciones con estos filtros', 'Cuando una auditoría puntúe un ítem en 0 o 1, la IA va a sugerir una acción acá.')}
        </div>
      </div>
    `;

    const btnQuitar = document.getElementById('btn-quitar-filtro-audit');
    if (btnQuitar) btnQuitar.addEventListener('click', () => { state.auditId = ''; render(); });

    document.getElementById('pf-estado').addEventListener('change', (e) => { state.estado = e.target.value; render(); });
    document.getElementById('pf-tipo').addEventListener('change', (e) => {
      state.tipo = e.target.value; state.plantaId = ''; state.uetSectorId = ''; render();
    });
    const pfPlanta = document.getElementById('pf-planta');
    if (pfPlanta) pfPlanta.addEventListener('change', (e) => { state.plantaId = e.target.value; state.uetSectorId = ''; render(); });
    const pfSector = document.getElementById('pf-sector');
    if (pfSector) pfSector.addEventListener('change', (e) => { state.uetSectorId = e.target.value; render(); });

    document.querySelectorAll('.plan-table .accion-estado').forEach((sel) => {
      sel.addEventListener('change', async () => {
        const accionId = Number(sel.dataset.accion);
        const estado = sel.value;
        await safeRun(async () => {
          const body = { estado };
          if (estado === 'cerrada') body.fecha_cierre = hoyISO();
          await api.actualizarAccion(accionId, body);
          toast('Acción actualizada');
          render();
        }, 'No se pudo actualizar la acción');
      });
    });
  }

  await safeRun(render, 'No se pudo cargar el plan de acción');
}

// ===========================================================================
// Screen 5 · Informe
// ===========================================================================
async function viewInforme(auditId) {
  setNav([{ label: 'Auditorías 5S', href: '/' }, { label: 'Informe', href: `/auditoria/${auditId}/informe` }]);
  root.innerHTML = loadingHtml();

  const [auditRaw, catalogos] = await Promise.all([api.obtenerAuditoria(auditId), getCatalogos()]);
  const audit = resolveNombres(auditRaw, catalogos);

  const SCORE_COL = { 0: 0, 1: 1, 3: 2, 5: 3 };
  const TITLES = { planta: 'AUDITORÍA 5S - PLANTA', panol: 'AUDITORÍA 5S - PAÑOL', oficina: 'AUDITORÍA 5S - OFICINAS' };

  function tabla() {
    const rows = [];
    rows.push(`<tr class="xls-banner"><td colspan="7">${TITLES[audit.audit_type]}</td></tr>`);
    rows.push(`<tr><td colspan="2" class="xls-field-label">FECHA</td><td colspan="2" class="xls-field-value">${audit.fecha}</td><td class="xls-field-label">EVALUADOR</td><td colspan="2" class="xls-field-value">${esc(audit.evaluador)}</td></tr>`);
    rows.push(`<tr><td colspan="2" class="xls-field-label">PLANTA</td><td colspan="2" class="xls-field-value">${esc(audit.planta_nombre || audit.area_nombre || '—')}</td><td class="xls-field-label">SUPERVISOR - LÍDER</td><td colspan="2" class="xls-field-value">${esc(audit.supervisor_lider || '—')}</td></tr>`);
    rows.push(`<tr><td colspan="2" class="xls-field-label">SECTOR</td><td colspan="2" class="xls-field-value">${esc(audit.uet_sector_nombre || audit.sector_nombre || '—')}</td><td class="xls-field-label">JEFE - GERENTE</td><td colspan="2" class="xls-field-value">${esc(audit.jefe_gerente || '—')}</td></tr>`);
    rows.push(`<tr><td colspan="2" class="xls-field-label">TURNO</td><td colspan="2" class="xls-field-value">${esc(audit.turno || '—')}</td><td colspan="3" class="xls-field-value" style="background:#fff;border:none;"></td></tr>`);
    rows.push(`<tr class="xls-head-row"><td colspan="2">EVALUACIÓN</td><td colspan="5"></td></tr>`);
    rows.push(`<tr class="xls-head-row"><th style="width:26px;">&quot;S&quot;Nº</th><th style="min-width:220px;">ITEM EVALUADO</th><th style="min-width:190px;">0</th><th style="min-width:190px;">1</th><th style="min-width:190px;">3</th><th style="min-width:190px;">5</th><th style="width:60px;">PUNTOS</th></tr>`);

    for (const cat of audit.categorias) {
      cat.items.forEach((item, i) => {
        rows.push('<tr>');
        if (i === 0) rows.push(`<td class="xls-cat" rowspan="${cat.items.length}">${esc(cat.label)}<br>${esc(cat.sub)}</td>`);
        rows.push(`<td class="xls-item">${esc(item.nombre)}</td>`);
        const crits = [item.criterio_0, item.criterio_1, item.criterio_3, item.criterio_5];
        const selIdx = item.score !== null && item.score !== undefined ? SCORE_COL[item.score] : -1;
        crits.forEach((c, ci) => rows.push(`<td class="xls-crit${ci === selIdx ? ' sel' : ''}">${esc(c)}</td>`));
        rows.push(`<td class="xls-pts${selIdx >= 0 ? ' sel' : ''}">${item.score ?? '—'}</td>`);
        rows.push('</tr>');
      });
      const known = cat.items.every((it) => it.score !== null && it.score !== undefined);
      const subtotal = cat.items.reduce((s, it) => s + (it.score || 0), 0);
      rows.push(`<tr class="xls-subtotal"><td colspan="6">SUB-TOTAL &quot;${cat.tag}&quot;</td><td class="xls-pts">${known ? subtotal : '—'}</td></tr>`);
    }

    const hasAny = audit.categorias.some((c) => c.items.some((it) => it.score !== null && it.score !== undefined));
    rows.push(`<tr class="xls-footer">
      <td colspan="2">PUNTAJE OBJETIVO:<br><span class="fval">${audit.puntaje_objetivo ?? '—'}</span></td>
      <td colspan="2">REVISIÓN: <span class="fval">${audit.revision ?? 1}</span><br>RESPONSABLE: <span class="fval">${esc(audit.responsable || audit.evaluador)}</span></td>
      <td colspan="2" class="total-label">PUNTAJE TOTAL</td><td class="total-val">${hasAny ? audit.puntaje_total : '—'}</td>
    </tr>`);

    return `<table class="xls-sheet${hasAny ? '' : ' xls-blank'}">${rows.join('')}</table>`;
  }

  const fotos = audit.categorias.flatMap((c) => c.items.flatMap((it) => (it.fotos || []).map((f) => ({ ...f, item: it.nombre }))));
  const acciones = audit.categorias.flatMap((c) => c.items.flatMap((it) => (it.acciones || []).map((a) => ({ ...a, item_nombre: it.nombre, cat_tag: c.tag }))));

  function tablaAcciones() {
    if (!acciones.length) {
      return `<div class="imprimible-note">Esta auditoría no generó acciones correctivas — todos los ítems puntuaron 3 o 5, o todavía no se cargó ninguna acción manual.</div>`;
    }
    return `
      <div style="overflow-x:auto;">
        <table class="plan-table">
          <thead>
            <tr><th>&quot;S&quot;</th><th>Ítem</th><th>Origen</th><th>Descripción</th><th>Responsable</th><th>Vence</th><th>Estado</th></tr>
          </thead>
          <tbody>
            ${acciones.map((a) => `
              <tr data-accion="${a.id}">
                <td>${esc(a.cat_tag)}</td>
                <td>${esc(a.item_nombre)}</td>
                <td><span class="accion-origen ${a.origen}">${a.origen === 'ia' ? '🤖 IA' : '✋ Manual'}</span></td>
                <td class="desc-cell">${esc(a.descripcion)}</td>
                <td>${esc(a.responsable || '—')}</td>
                <td>${a.fecha_compromiso || '—'}</td>
                <td>
                  <select class="accion-estado" data-accion="${a.id}">
                    ${Object.entries(ESTADO_LABEL).map(([v, l]) => `<option value="${v}" ${a.estado === v ? 'selected' : ''}>${l}</option>`).join('')}
                  </select>
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  root.innerHTML = `
    <div class="card">
      <div class="card-head">
        <h2>Informe</h2>
        <div class="sub">Réplica exacta de la solapa "8. Auditoría 5S y KPI"</div>
      </div>
      <div class="card-body">
        <div class="imprimible-note"><b>Hoja 1</b> — este cuadro es el que exige la norma. Se genera 1:1 desde los datos cargados.</div>
        <div class="xls-wrap">${tabla()}</div>

        <h3 style="margin-top:24px;">Hoja 2 · Fotos</h3>
        <div class="photo-grid">
          ${fotos.length ? fotos.map((f) => `<div class="photo-tile" style="background-image:url('${api.urlFoto(f.url)}'); background-size:cover;"></div>`).join('') : `<div class="photo-tile">Sin fotos</div>`}
        </div>

        <h3 style="margin-top:24px;">Hoja 3 · Plan de Acción</h3>
        ${tablaAcciones()}

        <div class="informe-acciones">
          <button class="download-btn" id="btn-descargar">⭳ Descargar .xlsx</button>
          <button class="email-btn" id="btn-enviar-correo">✉ Enviar por correo</button>
        </div>
      </div>
    </div>
  `;

  document.querySelectorAll('.plan-table .accion-estado').forEach((sel) => {
    sel.addEventListener('change', async () => {
      const accionId = Number(sel.dataset.accion);
      const estado = sel.value;
      await safeRun(async () => {
        const body = { estado };
        if (estado === 'cerrada') body.fecha_cierre = hoyISO();
        await api.actualizarAccion(accionId, body);
        toast('Acción actualizada');
      }, 'No se pudo actualizar la acción');
    });
  });

  document.getElementById('btn-descargar').addEventListener('click', async (e) => {
    const btn = e.currentTarget;
    const original = btn.textContent;
    btn.disabled = true;
    btn.textContent = 'Generando… (incrusta las fotos, puede tardar unos segundos)';
    try {
      await safeRun(() => descargarInforme(audit), 'No se pudo generar el .xlsx');
    } finally {
      btn.disabled = false;
      btn.textContent = original;
    }
  });

  document.getElementById('btn-enviar-correo').addEventListener('click', async (e) => {
    const btn = e.currentTarget;
    const original = btn.textContent;
    btn.disabled = true;
    btn.textContent = 'Generando…';
    try {
      await safeRun(() => descargarInforme(audit), 'No se pudo generar el .xlsx');

      const sector = audit.uet_sector_nombre || audit.sector_nombre || '—';
      const hasPuntaje = audit.categorias.some((c) => c.items.some((it) => it.score !== null && it.score !== undefined));
      const puntaje = hasPuntaje ? `${audit.puntaje_total} / ${audit.puntaje_objetivo ?? '—'}` : 'sin puntaje aún';

      const asunto = `Informe 5S - ${sector} - ${audit.fecha}`;
      const cuerpo =
        `Adjunto el informe de la Auditoría 5S del sector ${sector}, ` +
        `del día ${audit.fecha}. Puntaje: ${puntaje}.\n\n` +
        `⚠ Recordá adjuntar el archivo .xlsx que se acaba de descargar — ` +
        `este mail no lo adjunta automáticamente.`;

      const mailtoUrl = `mailto:?subject=${encodeURIComponent(asunto)}&body=${encodeURIComponent(cuerpo)}`;
      window.location.href = mailtoUrl;
    } finally {
      btn.disabled = false;
      btn.textContent = original;
    }
  });
}

router();
