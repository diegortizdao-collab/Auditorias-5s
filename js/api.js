import { API_BASE } from './config.js';

async function request(path, options = {}) {
  const res = await fetch(`${API_BASE}${path}`, {
    ...options,
    headers: {
      ...(options.body && !(options.body instanceof ArrayBuffer) && !(options.body instanceof Blob)
        ? { 'Content-Type': 'application/json' } : {}),
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    let msg = `Error ${res.status}`;
    try { const data = await res.json(); msg = data.error || msg; } catch (_e) { /* noop */ }
    throw new Error(msg);
  }
  const contentType = res.headers.get('content-type') || '';
  if (contentType.includes('application/json')) return res.json();
  return res;
}

export const api = {
  catalogos: () => request('/catalogos'),
  rubrica: (tipo) => request(`/rubrica/${tipo}`),

  listarAuditorias: (params = {}) => {
    const qs = new URLSearchParams(params).toString();
    return request(`/auditorias${qs ? `?${qs}` : ''}`);
  },
  crearAuditoria: (body) => request('/auditorias', { method: 'POST', body: JSON.stringify(body) }),
  obtenerAuditoria: (id) => request(`/auditorias/${id}`),
  actualizarAuditoria: (id, body) => request(`/auditorias/${id}`, { method: 'PUT', body: JSON.stringify(body) }),
  borrarAuditoria: (id) => request(`/auditorias/${id}`, { method: 'DELETE' }),

  guardarPuntaje: (auditId, itemId, body) =>
    request(`/auditorias/${auditId}/puntajes/${itemId}`, { method: 'PUT', body: JSON.stringify(body) }),

  subirFoto: (auditId, itemId, file) =>
    request(`/auditorias/${auditId}/puntajes/${itemId}/fotos`, {
      method: 'POST',
      headers: { 'Content-Type': file.type || 'application/octet-stream' },
      body: file,
    }),
  borrarFoto: (fotoId) => request(`/fotos/${fotoId}`, { method: 'DELETE' }),

  dashboard: (params = {}) => {
    const qs = new URLSearchParams(params).toString();
    return request(`/dashboard${qs ? `?${qs}` : ''}`);
  },

  listarPlanAccion: (params = {}) => {
    const qs = new URLSearchParams(params).toString();
    return request(`/plan-accion${qs ? `?${qs}` : ''}`);
  },
  crearAccion: (body) => request('/plan-accion', { method: 'POST', body: JSON.stringify(body) }),
  actualizarAccion: (id, body) => request(`/plan-accion/${id}`, { method: 'PUT', body: JSON.stringify(body) }),
  regenerarAccion: (id) => request(`/plan-accion/${id}/regenerar`, { method: 'POST' }),

  urlFoto: (path) => `${API_BASE}${path}`,
};
