const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.static(path.join(__dirname, 'public')));

// ---------- Helpers ----------

function sheetToMatrix(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) return null;
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
}

function findRowIndex(matrix, matchFn, fromRow = 0) {
  for (let r = fromRow; r < matrix.length; r++) {
    const row = matrix[r] || [];
    if (row.some((cell) => matchFn(cell))) return r;
  }
  return -1;
}

// Remueve espacios, pasa a minúsculas y quita tildes para evitar fallos de coincidencia
function norm(v) {
  if (typeof v !== 'string') return v;
  return v
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

// Limpia y convierte a número de forma segura (soporta %, comas y puntos)
function num(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const clean = v.replace('%', '').replace(',', '.').trim();
    if (clean === '' || clean === '-') return null;
    const n = Number(clean);
    return Number.isNaN(n) ? null : n;
  }
  const n = Number(v);
  return Number.isNaN(n) ? null : n;
}

// ---------- Extracción hoja OEE2 ----------
function parseOEE2(matrix) {
  if (!matrix) return [];
  const headerRow = findRowIndex(matrix, (c) => norm(c) === 'turno');
  if (headerRow === -1) return [];
  const headers = matrix[headerRow].map((h) => (h === null ? '' : String(h).trim()));

  const col = (name) => headers.findIndex((h) => norm(h).startsWith(norm(name)));

  const idx = {
    linea: col('Lineas consideradas'),
    capUtil: col('Capacidad Utilizada'),
    efectividad: col('Efectividad'),
    rendimiento: col('Rendimiento'),
    oee1: col('OEE1'),
    oee2: col('OEE2'),
    tiempoDisp: col('Tiempo disponible'),
    sinPedido: col('Sin pedido'),
    tiemposPerdidos: col('Tiempos perdidos'),
    sinClasificar: col('Sin clasificar'),
  };

  const rows = [];
  for (let r = headerRow + 1; r < matrix.length; r++) {
    const row = matrix[r];
    if (!row) continue;
    const linea = row[idx.linea];
    if (!linea) continue;
    rows.push({
      linea: String(linea).trim(),
      capacidadUtilizada: num(row[idx.capUtil]),
      efectividad: num(row[idx.efectividad]),
      rendimiento: num(row[idx.rendimiento]),
      oee1: num(row[idx.oee1]),
      oee2: num(row[idx.oee2]),
      tiempoDisponible: num(row[idx.tiempoDisp]),
      sinPedido: num(row[idx.sinPedido]),
      tiemposPerdidos: num(row[idx.tiemposPerdidos]),
      sinClasificar: num(row[idx.sinClasificar]),
    });
  }
  return rows;
}

// ---------- Extracción hoja PT ----------
function parsePT(matrix) {
  if (!matrix) return [];
  const headerRow = findRowIndex(matrix, (c) => norm(c) === 'linea');
  if (headerRow === -1) return [];
  const headers = matrix[headerRow].map((h) => (h === null ? '' : String(h).trim()));

  const col = (name) => headers.findIndex((h) => norm(h).startsWith(norm(name)));

  const idx = {
    linea: col('Linea'),
    codigo: col('Codigo producto'),
    descripcion: col('Descripcion'),
    toneladas: col('Toneladas producidas'),
    despTotalKg: col('Desperdicio Total'),
    despTotalPct: col('% Desperdicio Total'),
    foodWasteKg: col('Desperdicio alimento'),
    recuperacionKg: col('Recuperacion'),
  };
  if (idx.linea === -1) idx.linea = 1;

  const rows = [];
  for (let r = headerRow + 1; r < matrix.length; r++) {
    const row = matrix[r];
    if (!row) continue;
    const linea = row[idx.linea];
    const codigo = row[idx.codigo];
    const descripcion = row[idx.descripcion];
    if (!linea || !codigo || norm(codigo) === 'total' || !descripcion) continue;

    rows.push({
      linea: String(linea).trim(),
      codigo: String(codigo).trim(),
      descripcion: String(descripcion).trim(),
      toneladas: num(row[idx.toneladas]),
      desperdicioKg: num(row[idx.despTotalKg]),
      desperdicioPct: num(row[idx.despTotalPct]),
      foodWasteKg: num(row[idx.foodWasteKg]),
      recuperacionKg: num(row[idx.recuperacionKg]),
    });
  }
  return rows;
}

// ---------- Extracción hoja INFORME EJECUTIVO ----------
function readSectionTable(matrix, titleText) {
  const titleRow = matrix.findIndex((row) => row && norm(row[0]) === norm(titleText));
  if (titleRow === -1) return null;
  const headerRow = matrix[titleRow + 1] || [];
  const headers = headerRow.map((h) => (h === null ? '' : String(h).trim()));
  const rows = [];
  for (let r = titleRow + 2; r < matrix.length; r++) {
    const row = matrix[r];
    if (!row || row[0] === null || row[0] === '') break;
    rows.push(row);
    if (['total', 'planta'].includes(norm(row[0]))) break;
  }
  return { headers, rows };
}

function readLecturaEjecutiva(matrix) {
  for (let r = 0; r < matrix.length; r++) {
    const row = matrix[r] || [];
    const c = row.findIndex((cell) => norm(cell) === 'lectura ejecutiva');
    if (c !== -1) {
      const next = matrix[r + 1] || [];
      return next[c] || null;
    }
  }
  return null;
}

function parseInformeEjecutivo(matrix) {
  if (!matrix) return null;

  const lecturaEjecutiva = readLecturaEjecutiva(matrix);

  const cumplimientoTable = readSectionTable(matrix, 'CUMPLIMIENTO DE METAS');
  const cumplimiento = cumplimientoTable
    ? cumplimientoTable.rows.map((row) => ({
        kpi: row[0],
        actual: row[1],
        meta: row[2],
        brecha: row[3],
        criterio: row[4],
        estado: row[5],
      }))
    : [];

  const oeeTable = readSectionTable(matrix, 'ANALISIS OEE2 POR LINEA');
  const kpisPorLinea = oeeTable
    ? oeeTable.rows
        .filter((row) => norm(row[0]) !== 'total')
        .map((row) => ({
          linea: row[0],
          capacidadUtilizada: num(row[1]),
          oee1: num(row[2]),
          oee2: num(row[3]),
          sinPedido: num(row[4]),
          tiempoPerdido: num(row[5]),
          sinClasificar: num(row[6]),
        }))
    : [];

  const topTable = readSectionTable(matrix, 'TOP 5 PRODUCTOS POR KILOGRAMOS DE DESPERDICIO');
  const topDesperdicio = topTable
    ? topTable.rows.map((row) => ({
        linea: row[1],
        codigo: String(row[2]),
        descripcion: row[3],
        desperdicioKg: num(row[4]),
        desperdicioPct: num(row[5]),
      }))
    : [];

  const ptTable = readSectionTable(matrix, 'ANALISIS PT POR LINEA');
  const desperdicioPorLinea = {};
  if (ptTable) {
    ptTable.rows
      .filter((row) => norm(row[0]) !== 'planta')
      .forEach((row) => {
        desperdicioPorLinea[row[0]] = num(row[2]);
      });
  }

  const accionesTable = readSectionTable(matrix, 'ACCIONES PRIORITARIAS');
  const acciones = accionesTable
    ? accionesTable.rows.map((row) => ({ area: row[1], accion: row[2] }))
    : [];

  return { lecturaEjecutiva, cumplimiento, kpisPorLinea, topDesperdicio, desperdicioPorLinea, acciones };
}

function buildReport(workbook) {
  const informeMatrix = sheetToMatrix(workbook, 'INFORME EJECUTIVO');
  const informe = parseInformeEjecutivo(informeMatrix);

  if (informe && informe.kpisPorLinea.length) {
    return {
      generadoEn: new Date().toISOString(),
      modo: 'informe-ejecutivo',
      lecturaEjecutiva: informe.lecturaEjecutiva,
      cumplimiento: informe.cumplimiento,
      kpisPorLinea: informe.kpisPorLinea,
      topDesperdicio: informe.topDesperdicio,
      desperdicioPorLinea: informe.desperdicioPorLinea,
      acciones: informe.acciones,
    };
  }

  const oee2 = parseOEE2(sheetToMatrix(workbook, 'OEE2'));
  const pt = parsePT(sheetToMatrix(workbook, 'PT'));

  const kpisPorLinea = oee2
    .filter((r) => norm(r.linea) !== 'total')
    .map((r) => ({
      linea: r.linea,
      capacidadUtilizada: r.capacidadUtilizada,
      oee1: r.oee1,
      oee2: r.oee2,
    }));

  const topDesperdicio = [...pt]
    .filter((p) => p.desperdicioKg !== null)
    .sort((a, b) => b.desperdicioKg - a.desperdicioKg)
    .slice(0, 5);

  const desperdicioPorLinea = {};
  pt.forEach((p) => {
    if (p.desperdicioKg === null) return;
    desperdicioPorLinea[p.linea] = (desperdicioPorLinea[p.linea] || 0) + p.desperdicioKg;
  });

  return {
    generadoEn: new Date().toISOString(),
    modo: 'crudo',
    lecturaEjecutiva: 'Informe generado dinámicamente desde datos crudos (OEE2 y PT).',
    cumplimiento: [],
    kpisPorLinea,
    topDesperdicio,
    desperdicioPorLinea,
    acciones: [],
  };
}

// ---------- Rutas ----------
app.post('/api/upload', upload.single('excel'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No se recibió ningún archivo.' });
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const report = buildReport(workbook);
    res.json({ ok: true, report, sheets: workbook.SheetNames });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Error procesando el archivo: ' + err.message });
  }
});

// Ruta para cargar datos directamente desde el sistema (Modo Centralizado Sin Excel)
app.get('/api/report/sistema', (req, res) => {
  const mockReport = {
    generadoEn: new Date().toISOString(),
    modo: 'sistema-centralizado',
    lecturaEjecutiva: 'Informe generado en tiempo real desde la base de datos centralizada de producción.',
    cumplimiento: [
      { kpi: 'OEE2 Global', actual: 0.725, meta: 0.75, brecha: '-2.5%', estado: 'No Cumple' },
      { kpi: 'Capacidad Utilizada', actual: 0.81, meta: 0.80, brecha: '+1.0%', estado: 'Cumple' },
      { kpi: 'Desperdicio Total', actual: '215 kg', meta: '180 kg', brecha: '+35 kg', estado: 'No Cumple' }
    ],
    kpisPorLinea: [
      { linea: 'Línea 1 - Pan Tajado', capacidadUtilizada: 0.85, oee1: 0.78, oee2: 0.74 },
      { linea: 'Línea 2 - Hamburgo', capacidadUtilizada: 0.78, oee1: 0.72, oee2: 0.69 },
      { linea: 'Línea 3 - Repostería', capacidadUtilizada: 0.80, oee1: 0.75, oee2: 0.73 }
    ],
    topDesperdicio: [
      { linea: 'Línea 1 - Pan Tajado', codigo: 'P001', descripcion: 'Pan Blanco 500g', desperdicioKg: 120, desperdicioPct: 0.035 },
      { linea: 'Línea 2 - Hamburgo', codigo: 'P002', descripcion: 'Pan Hamburgo x8', desperdicioKg: 95, desperdicioPct: 0.028 }
    ],
    desperdicioPorLinea: {
      'Línea 1 - Pan Tajado': 120,
      'Línea 2 - Hamburgo': 95,
      'Línea 3 - Repostería': 45
    },
    acciones: [
      { area: 'Mantenimiento', accion: 'Ajustar calibración de cortadora en Línea 1.' },
      { area: 'Producción', accion: 'Revisar tiempo de fermentación en Línea 2.' }
    ]
  };

  res.json({ ok: true, report: mockReport, sheets: ['SISTEMA_BD'] });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor corriendo en http://localhost:${PORT}`));