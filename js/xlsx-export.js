// Genera el Informe (Hoja 1) en .xlsx, réplica exacta de la solapa
// "8. Auditoría 5S y KPI" del Excel original. Usa el ExcelJS cargado por
// CDN (ver <script> en index.html) — acá solo referenciamos window.ExcelJS.
// Prototipo verificado en Node antes de portarlo a este módulo de browser
// (mismos merges / colores / rotación de texto, confirmados con openpyxl).

const COLOR = {
  headerDark: 'FF20313E',
  brandSoft: 'FFBFE1F2',
  brand: 'FF0079BB',
  fieldLabelBg: 'FFEFEFEF',
  subtotalBg: 'FFEDEDED',
  subtotalPtsBg: 'FFDCEEFB',
  footerBg: 'FFF5F5F5',
  totalBg: 'FF3F3F3F',
  critSelBg: 'FFDCEEFB',
};

const THIN = { style: 'thin', color: { argb: 'FFB9B9B9' } };
const BORDER_ALL = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const SCORE_COL = { 0: 3, 1: 4, 3: 5, 5: 6 };

function fillColor(color) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
}

const TITLES = {
  planta: 'AUDITORÍA 5S - PLANTA',
  panol: 'AUDITORÍA 5S - PAÑOL',
  oficina: 'AUDITORÍA 5S - OFICINAS',
};

const ESTADO_LABEL = { abierta: 'Abierta', en_proceso: 'En proceso', cerrada: 'Cerrada' };
const ORIGEN_LABEL = { ia: 'IA', manual: 'Manual' };

export function buildInformeWorkbook(audit) {
  const wb = new window.ExcelJS.Workbook();
  const ws = wb.addWorksheet('Hoja 1', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 8 }, { width: 30 }, { width: 24 }, { width: 24 }, { width: 24 }, { width: 24 }, { width: 9 },
  ];

  let r = 1;
  ws.mergeCells(r, 1, r, 7);
  const banner = ws.getCell(r, 1);
  banner.value = TITLES[audit.audit_type] || 'AUDITORÍA 5S';
  banner.fill = fillColor(COLOR.headerDark);
  banner.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
  banner.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(r).height = 22;
  r++;

  const sectorNombre = audit.uet_sector_nombre || audit.sector_nombre || '';
  const plantaNombre = audit.planta_nombre || audit.area_nombre || '';

  const fieldRows = [
    ['FECHA', audit.fecha, 'EVALUADOR', audit.evaluador],
    ['PLANTA', plantaNombre, 'SUPERVISOR - LÍDER', audit.supervisor_lider || '—'],
    ['SECTOR', sectorNombre, 'JEFE - GERENTE', audit.jefe_gerente || '—'],
    ['TURNO', audit.turno || '—', '', ''],
  ];
  for (const [l1, v1, l2, v2] of fieldRows) {
    ws.mergeCells(r, 1, r, 2);
    ws.mergeCells(r, 3, r, 4);
    const c1 = ws.getCell(r, 1);
    c1.value = l1; c1.fill = fillColor(COLOR.fieldLabelBg); c1.font = { bold: true, size: 9 };
    const c2 = ws.getCell(r, 3);
    c2.value = v1; c2.font = { bold: true };
    if (l2) {
      const c3 = ws.getCell(r, 5);
      c3.value = l2; c3.fill = fillColor(COLOR.fieldLabelBg); c3.font = { bold: true, size: 9 };
      ws.mergeCells(r, 6, r, 7);
      const c4 = ws.getCell(r, 6);
      c4.value = v2; c4.font = { bold: true };
    }
    for (let c = 1; c <= 7; c++) ws.getCell(r, c).border = BORDER_ALL;
    r++;
  }

  ws.mergeCells(r, 1, r, 2);
  ws.getCell(r, 1).value = 'EVALUACIÓN';
  ws.getCell(r, 1).font = { bold: true, size: 10 };
  ws.getCell(r, 1).fill = fillColor(COLOR.fieldLabelBg);
  for (let c = 1; c <= 7; c++) ws.getCell(r, c).border = BORDER_ALL;
  r++;

  ['"S"Nº', 'ITEM EVALUADO', '0', '1', '3', '5', 'PUNTOS'].forEach((label, idx) => {
    const cell = ws.getCell(r, idx + 1);
    cell.value = label;
    cell.fill = fillColor(COLOR.fieldLabelBg);
    cell.font = { bold: true, size: 9 };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = BORDER_ALL;
  });
  r++;

  let hasAnyScore = false;

  for (const cat of audit.categorias) {
    const startRow = r;
    cat.items.forEach((item, idx) => {
      const row = r + idx;
      ws.getCell(row, 2).value = item.nombre;
      ws.getCell(row, 2).font = { bold: true, size: 10 };
      ws.getCell(row, 2).alignment = { wrapText: true, vertical: 'middle' };

      [item.criterio_0, item.criterio_1, item.criterio_3, item.criterio_5].forEach((crit, ci) => {
        const cell = ws.getCell(row, 3 + ci);
        cell.value = crit || '';
        cell.alignment = { wrapText: true, vertical: 'middle' };
        cell.font = { size: 9 };
        if (item.score !== null && item.score !== undefined && SCORE_COL[item.score] === 3 + ci) {
          cell.fill = fillColor(COLOR.critSelBg);
          cell.font = { bold: true, size: 9 };
        }
      });

      const ptsCell = ws.getCell(row, 7);
      if (item.score !== null && item.score !== undefined) {
        ptsCell.value = item.score;
        ptsCell.fill = fillColor(COLOR.brand);
        ptsCell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        hasAnyScore = true;
      } else {
        ptsCell.value = '—';
      }
      ptsCell.alignment = { horizontal: 'center', vertical: 'middle' };
      for (let c = 1; c <= 7; c++) ws.getCell(row, c).border = BORDER_ALL;
    });

    ws.mergeCells(startRow, 1, startRow + cat.items.length - 1, 1);
    const catCell = ws.getCell(startRow, 1);
    catCell.value = `${cat.label}\n${cat.sub}`;
    catCell.fill = fillColor(COLOR.brandSoft);
    catCell.font = { bold: true, size: 9 };
    catCell.alignment = { textRotation: 90, horizontal: 'center', vertical: 'middle', wrapText: true };

    r += cat.items.length;

    const subtotalKnown = cat.items.every((it) => it.score !== null && it.score !== undefined);
    const subtotal = cat.items.reduce((sum, it) => sum + (it.score || 0), 0);

    ws.mergeCells(r, 1, r, 6);
    const stLabel = ws.getCell(r, 1);
    stLabel.value = `SUB-TOTAL "${cat.tag}"`;
    stLabel.fill = fillColor(COLOR.subtotalBg);
    stLabel.font = { bold: true, size: 10 };
    stLabel.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    const stVal = ws.getCell(r, 7);
    stVal.value = subtotalKnown ? subtotal : '—';
    stVal.fill = fillColor(COLOR.subtotalPtsBg);
    stVal.font = { bold: true };
    stVal.alignment = { horizontal: 'center', vertical: 'middle' };
    for (let c = 1; c <= 7; c++) ws.getCell(r, c).border = BORDER_ALL;
    r++;
  }

  ws.mergeCells(r, 1, r, 2);
  const fObj = ws.getCell(r, 1);
  fObj.value = `PUNTAJE OBJETIVO:\n${audit.puntaje_objetivo ?? '—'}`;
  fObj.fill = fillColor(COLOR.footerBg);
  fObj.font = { bold: true, size: 9 };
  fObj.alignment = { wrapText: true, vertical: 'middle' };

  ws.mergeCells(r, 3, r, 4);
  const fRev = ws.getCell(r, 3);
  fRev.value = `REVISIÓN: ${audit.revision ?? 1}\nRESPONSABLE: ${audit.responsable || audit.evaluador}`;
  fRev.fill = fillColor(COLOR.footerBg);
  fRev.font = { bold: true, size: 9 };
  fRev.alignment = { wrapText: true, vertical: 'middle' };

  ws.mergeCells(r, 5, r, 6);
  const fLabel = ws.getCell(r, 5);
  fLabel.value = 'PUNTAJE TOTAL';
  fLabel.fill = fillColor(COLOR.totalBg);
  fLabel.font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
  fLabel.alignment = { horizontal: 'center', vertical: 'middle' };

  const fVal = ws.getCell(r, 7);
  fVal.value = hasAnyScore ? audit.puntaje_total : '—';
  fVal.fill = fillColor(COLOR.totalBg);
  fVal.font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
  fVal.alignment = { horizontal: 'center', vertical: 'middle' };
  for (let c = 1; c <= 7; c++) ws.getCell(r, c).border = BORDER_ALL;
  ws.getRow(r).height = 30;

  buildPlanAccionSheet(wb, audit);

  return wb;
}

// Hoja 2 — Plan de Acción: las acciones (sugeridas por IA + manuales) que
// dejó esta auditoría, item por item. Es la salida directa del plan de
// acción para poder imprimirla o llevarla al seguimiento.
function buildPlanAccionSheet(wb, audit) {
  const ws = wb.addWorksheet('Plan de Accion', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 6 }, { width: 26 }, { width: 10 }, { width: 44 }, { width: 18 }, { width: 13 }, { width: 13 },
  ];

  const acciones = (audit.categorias || []).flatMap((cat) =>
    cat.items.flatMap((item) => (item.acciones || []).map((a) => ({ ...a, item_nombre: item.nombre, cat_tag: cat.tag })))
  );

  let r = 1;
  ws.mergeCells(r, 1, r, 7);
  const banner = ws.getCell(r, 1);
  banner.value = 'PLAN DE ACCIÓN';
  banner.fill = fillColor(COLOR.headerDark);
  banner.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
  banner.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(r).height = 22;
  r++;

  ['"S"', 'ÍTEM', 'ORIGEN', 'DESCRIPCIÓN', 'RESPONSABLE', 'VENCE', 'ESTADO'].forEach((label, idx) => {
    const cell = ws.getCell(r, idx + 1);
    cell.value = label;
    cell.fill = fillColor(COLOR.fieldLabelBg);
    cell.font = { bold: true, size: 9 };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = BORDER_ALL;
  });
  r++;

  if (!acciones.length) {
    ws.mergeCells(r, 1, r, 7);
    const cell = ws.getCell(r, 1);
    cell.value = 'Esta auditoría no generó acciones correctivas — todos los ítems puntuaron 3 o 5, o todavía no se cargó ninguna acción manual.';
    cell.font = { italic: true, color: { argb: 'FF57707F' } };
    cell.alignment = { vertical: 'middle' };
    for (let c = 1; c <= 7; c++) ws.getCell(r, c).border = BORDER_ALL;
    return;
  }

  for (const a of acciones) {
    ws.getCell(r, 1).value = a.cat_tag || '';
    ws.getCell(r, 1).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(r, 2).value = a.item_nombre || '';
    ws.getCell(r, 2).font = { bold: true, size: 10 };
    ws.getCell(r, 2).alignment = { wrapText: true, vertical: 'middle' };
    ws.getCell(r, 3).value = ORIGEN_LABEL[a.origen] || a.origen;
    ws.getCell(r, 3).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(r, 4).value = a.descripcion || '';
    ws.getCell(r, 4).alignment = { wrapText: true, vertical: 'middle' };
    ws.getCell(r, 5).value = a.responsable || '—';
    ws.getCell(r, 5).alignment = { vertical: 'middle' };
    ws.getCell(r, 6).value = a.fecha_compromiso || '—';
    ws.getCell(r, 6).alignment = { horizontal: 'center', vertical: 'middle' };
    const estCell = ws.getCell(r, 7);
    estCell.value = ESTADO_LABEL[a.estado] || a.estado;
    estCell.alignment = { horizontal: 'center', vertical: 'middle' };
    estCell.font = { bold: true };
    if (a.estado === 'cerrada') estCell.fill = fillColor('FFE3F3EC');
    for (let c = 1; c <= 7; c++) ws.getCell(r, c).border = BORDER_ALL;
    r++;
  }
}

export async function descargarInforme(audit) {
  const wb = buildInformeWorkbook(audit);
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const fecha = (audit.fecha || '').replaceAll('-', '');
  a.href = url;
  a.download = `Auditoria5S_${audit.audit_type}_${fecha || 'sf'}_${audit.id}.xlsx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}
