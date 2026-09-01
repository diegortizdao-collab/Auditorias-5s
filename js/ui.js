export function toast(msg, type = 'ok') {
  let wrap = document.querySelector('.toast-wrap');
  if (!wrap) {
    wrap = document.createElement('div');
    wrap.className = 'toast-wrap';
    document.body.appendChild(wrap);
  }
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.textContent = msg;
  wrap.appendChild(el);
  setTimeout(() => el.remove(), 3600);
}

export function loadingHtml(label = 'Cargando…') {
  return `<div class="loading"><div class="spinner"></div><div>${label}</div></div>`;
}

export function emptyStateHtml(icon, msg, sub) {
  return `<div class="empty-state"><div style="font-size:26px;opacity:.55;">${icon}</div><div style="font-weight:700;color:var(--ink);">${msg}</div>${sub ? `<div style="font-size:12.5px;max-width:340px;">${sub}</div>` : ''}</div>`;
}

export function esc(s) {
  if (s === null || s === undefined) return '';
  const div = document.createElement('div');
  div.textContent = String(s);
  return div.innerHTML;
}

export async function safeRun(fn, errMsg = 'Ocurrió un error') {
  try {
    return await fn();
  } catch (err) {
    console.error(err);
    toast(err.message || errMsg, 'error');
    throw err;
  }
}
