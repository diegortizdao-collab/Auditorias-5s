// Generadores de SVG para el Dashboard. Misma geometría exacta que se usó
// a mano en el preview de diseño (radar pentagonal, línea de evolución
// 0-100 con escala 1.76 px/unto, barras comparativo con escala 4.579 px/pto),
// ahora calculadas dinámicamente a partir de datos reales del API.

const MESES = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];

function polarPoint(cx, cy, r, angleDeg) {
  const rad = (angleDeg * Math.PI) / 180;
  return { x: cx + r * Math.cos(rad), y: cy + r * Math.sin(rad) };
}

const RADAR_LABEL_POS = [
  { x: 160, y: 16, anchor: 'middle' },
  { x: 296.9, y: 112, anchor: 'start' },
  { x: 244.6, y: 280, anchor: 'middle' },
  { x: 75.4, y: 280, anchor: 'middle' },
  { x: 23.1, y: 112, anchor: 'end' },
];

export function generateRadarSVG(categorias) {
  const cx = 160, cy = 160, rMax = 118;
  const n = categorias.length || 5;

  const grid = [];
  for (let k = 1; k <= 5; k++) {
    const r = (rMax * k) / 5;
    const pts = Array.from({ length: n }, (_, i) => polarPoint(cx, cy, r, -90 + i * (360 / n)));
    grid.push(`<polygon points="${pts.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ')}" fill="none" stroke="var(--border)" stroke-width="1"/>`);
  }

  const spokes = Array.from({ length: n }, (_, i) => {
    const p = polarPoint(cx, cy, rMax, -90 + i * (360 / n));
    return `<line x1="${cx}" y1="${cy}" x2="${p.x.toFixed(1)}" y2="${p.y.toFixed(1)}" stroke="var(--border)"/>`;
  });

  const valuePts = categorias.map((c, i) => {
    const pct = c.pct === null || c.pct === undefined ? 0 : c.pct;
    return polarPoint(cx, cy, (rMax * pct) / 100, -90 + i * (360 / n));
  });
  const valuePoly = `<polygon points="${valuePts.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ')}" fill="var(--brand)" fill-opacity="0.28" stroke="var(--brand)" stroke-width="2.5" stroke-linejoin="round"/>`;
  const dots = valuePts.map((p) => `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="4" fill="var(--brand)"/>`);

  const labels = categorias.map((c, i) => {
    const pos = RADAR_LABEL_POS[i] || RADAR_LABEL_POS[0];
    const pctTxt = c.pct === null || c.pct === undefined ? '—' : `${c.pct}%`;
    return `<text x="${pos.x}" y="${pos.y}" text-anchor="${pos.anchor}" font-size="13" font-weight="600" fill="var(--ink)">${c.tag} · ${pctTxt}</text>`;
  });

  return `<svg viewBox="0 0 320 300" width="100%" height="260">
    ${grid.join('\n')}
    ${spokes.join('\n')}
    ${valuePoly}
    ${dots.join('\n')}
    ${labels.join('\n')}
  </svg>`;
}

export function generateEvolucionSVG(evolucion) {
  const x0 = 44, x1 = 624, yBase = 192, yTop = 16, scale = 1.76;
  const step = (x1 - x0) / 11;
  const xFor = (idx) => x0 + step * idx;
  const yFor = (val) => yBase - val * scale;

  const grid = [0, 25, 50, 75, 100].map((v) => {
    const y = yFor(v).toFixed(1);
    return v === 0
      ? `<line x1="${x0}" y1="${y}" x2="${x1}" y2="${y}" stroke="var(--border)"/>`
      : `<line x1="${x0}" y1="${y}" x2="${x1}" y2="${y}" stroke="var(--border)" stroke-dasharray="2,4"/>`;
  });
  const yLabels = [0, 50, 100].map((v) => `<text x="38" y="${(yFor(v) + 4).toFixed(1)}" text-anchor="end" font-size="10" fill="var(--ink-muted)">${v}</text>`);

  const objPts = evolucion.filter((m) => m.objetivo !== null).map((m, i) => `${xFor(m.mes - 1).toFixed(1)},${yFor(m.objetivo).toFixed(1)}`);
  const objPath = objPts.length
    ? `<path d="M ${objPts.join(' L ')}" fill="none" stroke="var(--ink-muted)" stroke-width="2" stroke-dasharray="5,5"/>`
    : '';

  const audPts = evolucion.filter((m) => m.auditado !== null).map((m) => ({ x: xFor(m.mes - 1), y: yFor(m.auditado) }));
  const audPath = audPts.length > 1
    ? `<path d="M ${audPts.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' L ')}" fill="none" stroke="var(--brand)" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>`
    : '';
  const audDots = audPts.map((p) => `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="4.5" fill="var(--brand)"/>`);

  const mesLabels = MESES.map((m, i) => {
    const hoy = new Date().getMonth() === i;
    return `<text x="${xFor(i).toFixed(1)}" y="208" text-anchor="middle"${hoy ? ' fill="var(--brand)" font-weight="600"' : ''}>${m}</text>`;
  });

  return `<svg viewBox="0 0 640 220" width="100%" height="200">
    ${grid.join('\n')}
    ${yLabels.join('\n')}
    ${objPath}
    ${audPath}
    ${audDots.join('\n')}
    <g font-size="10.5" fill="var(--ink-muted)">${mesLabels.join('\n')}</g>
  </svg>`;
}

export function generateComparativoSVG(comparativo, objetivo) {
  const x0 = 132, scale = 4.579;
  const rowH = 54, barH = 40;
  const n = comparativo.length;
  const viewH = 56 + Math.max(n - 1, 0) * rowH;

  const objX = objetivo !== null && objetivo !== undefined ? (x0 + objetivo * scale).toFixed(1) : null;
  const objLine = objX
    ? `<line x1="${objX}" y1="4" x2="${objX}" y2="${viewH - 4}" stroke="var(--ink-muted)" stroke-width="1.5" stroke-dasharray="3,3"/>
       <text x="${objX}" y="14" text-anchor="middle" font-size="10" fill="var(--ink-muted)">Objetivo ${objetivo}</text>`
    : '';

  const bars = comparativo.map((row, i) => {
    const y = 8 + i * rowH;
    const w = Math.max(row.puntaje * scale, 0);
    const diff = objetivo !== null && objetivo !== undefined ? row.puntaje - objetivo : null;
    const color = diff === null ? 'var(--brand)' : diff >= 0 ? 'var(--good)' : diff >= -10 ? 'var(--warning)' : 'var(--critical)';
    return `<text x="0" y="${y + 16}" font-size="12.5" fill="var(--ink)">${row.sector}</text>
      <rect x="${x0}" y="${y}" width="${w.toFixed(1)}" height="${barH}" rx="7" fill="${color}" fill-opacity="0.85"/>
      <text x="${(x0 + w + 20).toFixed(1)}" y="${y + 24}" font-size="13" font-weight="700" fill="var(--ink)">${row.puntaje}</text>`;
  });

  return `<svg viewBox="0 0 640 ${viewH}" width="100%" height="${Math.min(viewH, 260)}">
    ${objLine}
    ${bars.join('\n')}
  </svg>`;
}
