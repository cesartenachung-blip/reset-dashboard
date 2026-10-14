// API del dashboard Reset: lee y actualiza los Excel de Google Drive
// GET  /api/data  -> devuelve los datos de Equipo Principal y Equipo Next
// POST /api/data  -> { records: [{ code: "C186D", mes: "OCT" }, ...] }
//                    recalcula SEP/OCT/NOV/DIC/TOTAL de los líderes presentes y guarda en Drive (LOGRO no se toca)
const { google } = require('googleapis');
const ExcelJS = require('exceljs');
const { Readable } = require('stream');

const FILES = {
  principal: process.env.FILE_ID_PRINCIPAL || '1OrJNlLTA5MMtvDtWeQeunafC0e3WWQGJ',
  next: process.env.FILE_ID_NEXT || '1b01_tVCwx5lLpkDuXvtyWUwLqAKd1foA'
};
const MESES = ['SEP', 'OCT', 'NOV', 'DIC'];
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function driveClient() {
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!raw) throw new Error('Falta la variable GOOGLE_SERVICE_ACCOUNT_JSON en Vercel');
  const credentials = JSON.parse(raw);
  const auth = new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/drive'] });
  return google.drive({ version: 'v3', auth });
}

async function loadWorkbook(drive, fileId) {
  const res = await drive.files.get({ fileId, alt: 'media', supportsAllDrives: true }, { responseType: 'arraybuffer' });
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(res.data));
  return wb;
}

function val(cell) {
  let v = cell.value;
  if (v && typeof v === 'object') {
    if ('result' in v) v = v.result;
    else if (v.richText) v = v.richText.map(t => t.text).join('');
    else if ('text' in v) v = v.text;
    else if (v instanceof Date) v = v.toISOString();
  }
  return v;
}
const num = v => (typeof v === 'number' ? v : (parseFloat(v) || 0));
const CODE_RE = /^C[12]8[0-9A-Z]{1,3}$/i;

// Ubica los bloques (Hombres / Mujeres) buscando la cabecera "SEP"
function findBlocks(wb) {
  for (const ws of wb.worksheets) {
    const blocks = [];
    ws.eachRow((row, r) => {
      row.eachCell((c, col) => {
        if (String(val(c) ?? '').trim().toUpperCase() === 'SEP') blocks.push({ ws, headerRow: r, sep: col });
      });
    });
    if (blocks.length) {
      return blocks.map(b => ({
        ws: b.ws, headerRow: b.headerRow,
        cCode: b.sep - 3, cName: b.sep - 2, cLogro: b.sep - 1,
        cMes: { SEP: b.sep, OCT: b.sep + 1, NOV: b.sep + 2, DIC: b.sep + 3 }, cTotal: b.sep + 4
      }));
    }
  }
  throw new Error('No se encontró la cabecera SEP en el Excel');
}

function readLeaders(blocks) {
  const out = [];
  for (const b of blocks) {
    for (let r = b.headerRow + 1; r <= b.ws.rowCount; r++) {
      const row = b.ws.getRow(r);
      const code = String(val(row.getCell(b.cCode)) ?? '').trim().toUpperCase();
      if (!CODE_RE.test(code)) continue;
      const item = {
        code, name: String(val(row.getCell(b.cName)) ?? '').trim(),
        sexo: code.startsWith('C18') ? 'H' : 'M',
        logro: num(val(row.getCell(b.cLogro))), row: r, block: b
      };
      MESES.forEach(m => (item[m] = num(val(row.getCell(b.cMes[m])))));
      item.total = MESES.reduce((s, m) => s + item[m], 0);
      out.push(item);
    }
  }
  return out;
}

const publicLeader = l => ({ code: l.code, name: l.name, sexo: l.sexo, logro: l.logro, SEP: l.SEP, OCT: l.OCT, NOV: l.NOV, DIC: l.DIC, total: l.total });

// Asigna cada código de inscrito al líder con el prefijo más largo (C186D -> C186, C18J3 -> C18J3 en Next / C18J en Principal)
function matchLeader(code, leaders) {
  let best = null;
  for (const l of leaders) if (code.startsWith(l.code) && (!best || l.code.length > best.code.length)) best = l;
  return best;
}

async function getData(drive) {
  const result = {};
  for (const [key, id] of Object.entries(FILES)) {
    const [wb, meta] = await Promise.all([
      loadWorkbook(drive, id),
      drive.files.get({ fileId: id, fields: 'modifiedTime,name', supportsAllDrives: true })
    ]);
    result[key] = { file: meta.data.name, modifiedTime: meta.data.modifiedTime, leaders: readLeaders(findBlocks(wb)).map(publicLeader) };
  }
  return result;
}

async function update(drive, records) {
  const summary = {};
  for (const [key, id] of Object.entries(FILES)) {
    const wb = await loadWorkbook(drive, id);
    const leaders = readLeaders(findBlocks(wb));
    const counts = new Map();
    for (const rec of records) {
      const l = matchLeader(rec.code, leaders);
      if (!l || !MESES.includes(rec.mes)) continue;
      if (!counts.has(l)) counts.set(l, { SEP: 0, OCT: 0, NOV: 0, DIC: 0 });
      counts.get(l)[rec.mes]++;
    }
    for (const [l, c] of counts) {
      const row = l.block.ws.getRow(l.row);
      MESES.forEach(m => { row.getCell(l.block.cMes[m]).value = c[m] || null; });
      const total = MESES.reduce((s, m) => s + c[m], 0);
      const tCell = row.getCell(l.block.cTotal);
      const f = tCell.value && typeof tCell.value === 'object' && (tCell.value.formula || tCell.value.sharedFormula);
      tCell.value = f ? { ...tCell.value, result: total } : total;
      row.commit();
    }
    wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
    const buffer = Buffer.from(await wb.xlsx.writeBuffer());
    await drive.files.update({ fileId: id, media: { mimeType: XLSX_MIME, body: Readable.from(buffer) }, supportsAllDrives: true });
    summary[key] = { lideresActualizados: counts.size };
  }
  return summary;
}

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  try {
    const drive = driveClient();
    if (req.method === 'GET') return res.status(200).json(await getData(drive));
    if (req.method === 'POST') {
      const key = process.env.UPLOAD_KEY;
      if (key && req.headers['x-upload-key'] !== key) return res.status(401).json({ error: 'Clave de carga incorrecta' });
      const body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
      const records = (body?.records || [])
        .map(r => ({ code: String(r.code || '').toUpperCase().trim(), mes: String(r.mes || '').toUpperCase().trim() }))
        .filter(r => CODE_RE.test(r.code) || /^C[12]8/.test(r.code));
      if (!records.length) return res.status(400).json({ error: 'No se recibieron registros válidos' });
      const summary = await update(drive, records);
      return res.status(200).json({ ok: true, summary, data: await getData(drive) });
    }
    res.setHeader('Allow', 'GET, POST');
    return res.status(405).json({ error: 'Método no permitido' });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: e.message || 'Error interno' });
  }
};
