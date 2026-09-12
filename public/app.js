let globalReport = null;
let chartInstanceLineas = null;
let chartInstanceDesperdicio = null;
let chartInstancePorLinea = null;

// Listeners de Eventos
document.getElementById('generarBtn').addEventListener('click', uploadFile);
document.getElementById('toggleManualBtn').addEventListener('click', toggleManualForm);
document.getElementById('procesarManualBtn').addEventListener('click', procesarDatosManuales);
document.getElementById('filtroLinea').addEventListener('change', aplicarFiltroLinea);
document.getElementById('printBtn').addEventListener('click', () => window.print());

// Buttons para agregar filas a la tabla manual
document.getElementById('addRowLineaBtn').addEventListener('click', () => addRowLinea());
document.getElementById('addRowTopBtn').addEventListener('click', () => addRowTop());
document.getElementById('addRowMetaBtn').addEventListener('click', () => addRowMeta());
document.getElementById('addRowAccionBtn').addEventListener('click', () => addRowAccion());

// ---------- Helpers de Formateo ----------
function formatPct(v) {
  if (v === null || v === undefined || v === '') return '-';
  
  let num = typeof v === 'number' ? v : parseFloat(String(v).replace('%', '').replace(',', '.').trim());
  if (isNaN(num)) return String(v);

  const val = (num > 0 && num <= 1) ? num * 100 : num;
  return val.toFixed(1) + '%';
}

function formatVal(v) {
  if (v === null || v === undefined || v === '') return '-';
  if (typeof v === 'number') return parseFloat(v.toFixed(2)).toLocaleString();
  return String(v);
}

function parseNum(v) {
  if (v === null || v === undefined) return null;
  const str = String(v).trim().replace(',', '.').replace('%', '');
  if (!str) return null;
  const n = parseFloat(str);
  return isNaN(n) ? null : n;
}

// ---------- Filtro de Limpieza de Totales y Duplicados ----------
function cleanReportData(report) {
  if (!report) return report;

  const esIgnorable = (txt) => {
    if (!txt) return false;
    const norm = String(txt).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
    const palabrasClave = ['total', 'totales', 'promedio', 'subtotal', 'gran total', 'resumen'];
    return palabrasClave.includes(norm) || norm.startsWith('total ') || norm === 'total';
  };

  // 1. Filtrar líneas OEE duplicadas o filas de totales
  const kpisPorLinea = (report.kpisPorLinea || []).filter((item, idx, self) => {
    if (!item || !item.linea || esIgnorable(item.linea)) return false;
    return self.findIndex(t => String(t.linea).trim().toLowerCase() === String(item.linea).trim().toLowerCase()) === idx;
  });

  // 2. Filtrar productos duplicados o filas de totales
  const topDesperdicio = (report.topDesperdicio || []).filter((item, idx, self) => {
    const nombre = item.descripcion || item.producto || item.codigo;
    if (!nombre || esIgnorable(nombre)) return false;
    return self.findIndex(t => (String(t.linea) + String(t.codigo) + String(nombre)).toLowerCase() === (String(item.linea) + String(item.codigo) + String(t.descripcion || t.producto)).toLowerCase()) === idx;
  });

  // 3. Recalcular el desperdicio acumulado por línea para el gráfico circular
  const desperdicioPorLinea = {};
  topDesperdicio.forEach(p => {
    const l = p.linea || 'General';
    const kg = parseNum(p.desperdicioKg) || 0;
    desperdicioPorLinea[l] = (desperdicioPorLinea[l] || 0) + kg;
  });

  return {
    ...report,
    kpisPorLinea,
    topDesperdicio,
    desperdicioPorLinea
  };
}

// ---------- 1. Mostrar / Ocultar Formulario Manual ----------
function toggleManualForm() {
  const form = document.getElementById('manualFormSection');
  
  if (form.style.display === 'none' || form.style.display === '') {
    form.style.display = 'block';

    if (document.querySelectorAll('#tableInputLineas tbody tr').length === 0) {
      addRowLinea('Pan bollería', 31.9, 27.6, 86.4);
      addRowLinea('Panadería', 24.6, 24.1, 97.7);
      addRowLinea('Tortillas', 49.6, 48.0, 96.9);
      addRowTop('Pan bollería', 'P001', 'Pan Blanco 500g', 120, 3.5);
      addRowMeta('OEE2 Global', '86.4%', '85.0%', 'Cumple');
      addRowAccion('Mantenimiento', 'Revisar calibración de máquina cortadora.');
    }
  } else {
    form.style.display = 'none';
  }
}

// ---------- 2. Constructores de Filas Editables ----------
function addRowLinea(linea = '', cap = '', oee1 = '', oee2 = '') {
  const tbody = document.querySelector('#tableInputLineas tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" class="in-linea" value="${linea}" placeholder="Ej. Línea 1" style="width: 95%;"></td>
    <td><input type="number" class="in-cap" value="${cap}" placeholder="85" style="width: 90%;"></td>
    <td><input type="number" class="in-oee1" value="${oee1}" placeholder="78" style="width: 90%;"></td>
    <td><input type="number" class="in-oee2" value="${oee2}" placeholder="74" style="width: 90%;"></td>
    <td style="text-align:center;"><button type="button" onclick="this.closest('tr').remove()" style="color:red;">❌</button></td>
  `;
  tbody.appendChild(tr);
}

function addRowTop(linea = '', codigo = '', desc = '', kg = '', pct = '') {
  const tbody = document.querySelector('#tableInputTop tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" class="in-top-linea" value="${linea}" placeholder="Ej. Línea 1" style="width: 95%;"></td>
    <td><input type="text" class="in-top-codigo" value="${codigo}" placeholder="Ej. P001" style="width: 90%;"></td>
    <td><input type="text" class="in-top-desc" value="${desc}" placeholder="Ej. Pan Blanco 500g" style="width: 95%;"></td>
    <td><input type="number" class="in-top-kg" value="${kg}" placeholder="120" style="width: 90%;"></td>
    <td><input type="number" class="in-top-pct" value="${pct}" placeholder="3.5" style="width: 90%;"></td>
    <td style="text-align:center;"><button type="button" onclick="this.closest('tr').remove()" style="color:red;">❌</button></td>
  `;
  tbody.appendChild(tr);
}

function addRowMeta(kpi = '', actual = '', meta = '', estado = '') {
  const tbody = document.querySelector('#tableInputMetas tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" class="in-meta-kpi" value="${kpi}" placeholder="Ej. OEE Global" style="width: 95%;"></td>
    <td><input type="text" class="in-meta-actual" value="${actual}" placeholder="72.5%" style="width: 90%;"></td>
    <td><input type="text" class="in-meta-meta" value="${meta}" placeholder="75%" style="width: 90%;"></td>
    <td><input type="text" class="in-meta-estado" value="${estado}" placeholder="Cumple / No Cumple" style="width: 90%;"></td>
    <td style="text-align:center;"><button type="button" onclick="this.closest('tr').remove()" style="color:red;">❌</button></td>
  `;
  tbody.appendChild(tr);
}

function addRowAccion(area = '', accion = '') {
  const tbody = document.querySelector('#tableInputAcciones tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" class="in-acc-area" value="${area}" placeholder="Ej. Mantenimiento" style="width: 95%;"></td>
    <td><input type="text" class="in-acc-accion" value="${accion}" placeholder="Ej. Ajustar sensores..." style="width: 95%;"></td>
    <td style="text-align:center;"><button type="button" onclick="this.closest('tr').remove()" style="color:red;">❌</button></td>
  `;
  tbody.appendChild(tr);
}

// ---------- 3. Recopilar Datos Tipeados por el Usuario ----------
function procesarDatosManuales() {
  const lecturaEjecutiva = document.getElementById('inputLectura').value.trim();

  const kpisPorLinea = [];
  document.querySelectorAll('#tableInputLineas tbody tr').forEach((tr) => {
    const linea = tr.querySelector('.in-linea').value.trim();
    if (!linea) return;
    kpisPorLinea.push({
      linea,
      capacidadUtilizada: parseNum(tr.querySelector('.in-cap').value),
      oee1: parseNum(tr.querySelector('.in-oee1').value),
      oee2: parseNum(tr.querySelector('.in-oee2').value),
    });
  });

  const topDesperdicio = [];
  const desperdicioPorLinea = {};
  document.querySelectorAll('#tableInputTop tbody tr').forEach((tr) => {
    const linea = tr.querySelector('.in-top-linea').value.trim();
    const codigo = tr.querySelector('.in-top-codigo').value.trim();
    const descripcion = tr.querySelector('.in-top-desc').value.trim();
    const desperdicioKg = parseNum(tr.querySelector('.in-top-kg').value);
    const desperdicioPct = parseNum(tr.querySelector('.in-top-pct').value);

    if (!linea && !descripcion) return;

    topDesperdicio.push({ linea, codigo, descripcion, desperdicioKg, desperdicioPct });

    if (linea && desperdicioKg !== null) {
      desperdicioPorLinea[linea] = (desperdicioPorLinea[linea] || 0) + desperdicioKg;
    }
  });

  const cumplimiento = [];
  document.querySelectorAll('#tableInputMetas tbody tr').forEach((tr) => {
    const kpi = tr.querySelector('.in-meta-kpi').value.trim();
    if (!kpi) return;
    cumplimiento.push({
      kpi,
      actual: tr.querySelector('.in-meta-actual').value.trim(),
      meta: tr.querySelector('.in-meta-meta').value.trim(),
      estado: tr.querySelector('.in-meta-estado').value.trim(),
    });
  });

  const acciones = [];
  document.querySelectorAll('#tableInputAcciones tbody tr').forEach((tr) => {
    const area = tr.querySelector('.in-acc-area').value.trim();
    const accion = tr.querySelector('.in-acc-accion').value.trim();
    if (!area || !accion) return;
    acciones.push({ area, accion });
  });

  globalReport = cleanReportData({
    generadoEn: new Date().toISOString(),
    modo: 'manual',
    lecturaEjecutiva: lecturaEjecutiva || 'Informe generado mediante captura manual de datos.',
    cumplimiento,
    kpisPorLinea,
    topDesperdicio,
    desperdicioPorLinea,
    acciones,
  });

  document.getElementById('status').innerText = '✅ Reporte generado exitosamente con tus datos manuales.';
  poblarFiltroLineas(globalReport.kpisPorLinea);
  renderReport(globalReport);
}

// ---------- 4. Cargar desde Archivo Excel (Lector Matricial 2D) ----------
function uploadFile() {
  const fileInput = document.getElementById('fileInput');
  const status = document.getElementById('status');
  const file = fileInput?.files[0];

  if (!file) return alert('Selecciona un archivo Excel');

  status.innerText = 'Procesando Excel...';

  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const kpisPorLinea = [];
      const topDesperdicio = [];
      const cumplimiento = [];
      const acciones = [];

      const norm = (str) => String(str || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        // Convertir la hoja a una matriz 2D de filas
        const rowsMatrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        let currentHeaders = [];

        rowsMatrix.forEach(row => {
          if (!Array.isArray(row) || row.length === 0) return;

          const rowNorm = row.map(c => norm(c));

          // Detección dinámica de nuevas filas de encabezado en hojas apiladas
          const keywordsHeader = ['linea', 'proceso', 'equipo', 'indicador', 'kpi', 'metrica', 'producto', 'descripcion', 'articulo', 'codigo', 'sku', 'accion', 'area', 'desperdicio', 'merma'];
          const esFilaEncabezado = rowNorm.some(cell => keywordsHeader.some(k => cell === k || cell.includes(k)));

          if (esFilaEncabezado) {
            currentHeaders = rowNorm;
            return;
          }

          // Helper para extraer el valor de una celda basado en los encabezados activos
          const getValByHeader = (candidatos) => {
            if (!currentHeaders.length) return '';
            const idx = currentHeaders.findIndex(h => candidatos.some(c => h.includes(c)));
            return (idx !== -1 && row[idx] !== undefined) ? row[idx] : '';
          };

          // 1. Detección de Líneas OEE
          const linea = getValByHeader(['linea', 'proceso', 'equipo', 'planta', 'area', 'seccion']);
          const cap = parseNum(getValByHeader(['capacidad', 'utilizacion', 'cap_utilizada', 'cap', 'util']));
          const oee1 = parseNum(getValByHeader(['oee1', 'oee 1', 'oee_1']));
          const oee2 = parseNum(getValByHeader(['oee2', 'oee 2', 'oee_2', 'oee']));

          if (linea && (cap !== null || oee1 !== null || oee2 !== null)) {
            kpisPorLinea.push({
              linea: String(linea).trim(),
              capacidadUtilizada: cap,
              oee1: oee1,
              oee2: oee2
            });
          }

          // 2. Detección de Top Desperdicio
          const prod = getValByHeader(['producto', 'descripcion', 'articulo', 'item', 'nombre', 'material']);
          const cod = getValByHeader(['codigo', 'sku', 'ref', 'cod', 'sap', 'id']);
          const kg = parseNum(getValByHeader(['desperdicio', 'kg', 'merma', 'scrap', 'perdida', 'mermas', 'cantidad', 'peso']));
          const pct = parseNum(getValByHeader(['pct', 'porcentaje', '%', 'porcentual']));

          if (prod && kg !== null && kg > 0) {
            topDesperdicio.push({
              linea: linea ? String(linea).trim() : 'General',
              codigo: cod ? String(cod).trim() : 'S/C',
              descripcion: String(prod).trim(),
              producto: String(prod).trim(),
              desperdicioKg: kg,
              desperdicioPct: pct
            });
          }

          // 3. Detección de Cumplimiento de Metas
          const kpiNom = getValByHeader(['indicador', 'kpi', 'metrica', 'meta', 'objetivo', 'medida', 'concepto']);
          const valActual = getValByHeader(['actual', 'real', 'ejecutado', 'resultado', 'valor']);
          const valMeta = getValByHeader(['meta', 'objetivo', 'target', 'plan', 'esperado']);
          const valEstado = getValByHeader(['estado', 'cumplimiento', 'status', 'cumple', 'condicion']);

          if (kpiNom && (valActual !== '' || valMeta !== '' || valEstado !== '')) {
            cumplimiento.push({
              kpi: String(kpiNom).trim(),
              actual: valActual !== '' ? valActual : '-',
              meta: valMeta !== '' ? valMeta : '-',
              estado: valEstado !== '' ? valEstado : 'Pendiente'
            });
          }

          // 4. Detección de Acciones Prioritarias
          const accionTexto = getValByHeader(['accion', 'tarea', 'actividad', 'recomendacion', 'plan', 'mejora', 'hallazgo']);
          const areaTexto = getValByHeader(['area', 'departamento', 'responsable', 'encargado']);

          if (accionTexto && String(accionTexto).trim() !== '') {
            acciones.push({
              area: areaTexto ? String(areaTexto).trim() : 'General',
              accion: String(accionTexto).trim()
            });
          }
        });
      });

      globalReport = cleanReportData({
        generadoEn: new Date().toISOString(),
        modo: 'excel-client',
        lecturaEjecutiva: 'Informe procesado localmente desde archivo Excel.',
        cumplimiento,
        kpisPorLinea,
        topDesperdicio,
        acciones
      });

      status.innerText = '✅ Informe generado desde Excel';
      poblarFiltroLineas(globalReport.kpisPorLinea);
      renderReport(globalReport);

    } catch (err) {
      console.error(err);
      status.innerText = '❌ Error al leer el Excel en el navegador.';
    }
  };

  reader.readAsArrayBuffer(file);
}

// ---------- 5. Renderizar Filtros y Reporte Final ----------
function poblarFiltroLineas(lineas) {
  const select = document.getElementById('filtroLinea');
  select.innerHTML = '<option value="TODAS">-- Todas las Líneas --</option>';
  if (!lineas) return;
  lineas.forEach((l) => {
    const opt = document.createElement('option');
    opt.value = l.linea;
    opt.textContent = l.linea;
    select.appendChild(opt);
  });
}

function aplicarFiltroLinea() {
  if (!globalReport) return;
  const lineaSeleccionada = document.getElementById('filtroLinea').value;

  if (lineaSeleccionada === 'TODAS') {
    renderReport(globalReport);
  } else {
    const reportFiltrado = {
      ...globalReport,
      kpisPorLinea: (globalReport.kpisPorLinea || []).filter((l) => l.linea === lineaSeleccionada),
      topDesperdicio: (globalReport.topDesperdicio || []).filter((p) => p.linea === lineaSeleccionada),
    };
    renderReport(reportFiltrado);
  }
}

function renderReport(report) {
  if (!report) return;

  const cleanReport = cleanReportData(report);

  document.getElementById('report').classList.remove('hidden');
  document.getElementById('fecha').innerText = `Fecha de Generación: ${new Date(cleanReport.generadoEn).toLocaleString()}`;

  // Lectura ejecutiva
  document.getElementById('lecturaTexto').innerText = cleanReport.lecturaEjecutiva || 'Sin lectura registrada.';

  // Tabla Cumplimiento
  const tbodyCumplimiento = document.querySelector('#tablaCumplimiento tbody');
  tbodyCumplimiento.innerHTML = '';
  (cleanReport.cumplimiento || []).forEach((row) => {
    tbodyCumplimiento.innerHTML += `
      <tr>
        <td>${formatVal(row.kpi)}</td>
        <td>${formatPct(row.actual)}</td>
        <td>${formatPct(row.meta)}</td>
        <td>${formatVal(row.estado)}</td>
      </tr>
    `;
  });

  // Tabla Líneas OEE
  const tbodyLineas = document.querySelector('#tablaLineas tbody');
  tbodyLineas.innerHTML = '';
  (cleanReport.kpisPorLinea || []).forEach((row) => {
    tbodyLineas.innerHTML += `
      <tr>
        <td>${formatVal(row.linea)}</td>
        <td>${formatPct(row.capacidadUtilizada)}</td>
        <td>${formatPct(row.oee1)}</td>
        <td>${formatPct(row.oee2)}</td>
      </tr>
    `;
  });

  // Tabla Top Desperdicio
  const tbodyTop = document.querySelector('#tablaTop tbody');
  tbodyTop.innerHTML = '';
  (cleanReport.topDesperdicio || []).forEach((row, idx) => {
    tbodyTop.innerHTML += `
      <tr>
        <td>${idx + 1}</td>
        <td>${formatVal(row.linea)}</td>
        <td>${formatVal(row.codigo)}</td>
        <td>${formatVal(row.descripcion || row.producto)}</td>
        <td>${formatVal(row.desperdicioKg)} kg</td>
        <td>${formatPct(row.desperdicioPct)}</td>
      </tr>
    `;
  });

  // Acciones prioritarias
  const listaAcciones = document.getElementById('listaAcciones');
  listaAcciones.innerHTML = '';
  (cleanReport.acciones || []).forEach((acc) => {
    listaAcciones.innerHTML += `<li><strong>${formatVal(acc.area)}:</strong> ${formatVal(acc.accion)}</li>`;
  });

  // Renderizar Gráficos con Chart.js
  renderCharts(cleanReport);
}

function renderCharts(report) {
  if (typeof Chart === 'undefined') return;

  if (chartInstanceLineas) chartInstanceLineas.destroy();
  if (chartInstanceDesperdicio) chartInstanceDesperdicio.destroy();
  if (chartInstancePorLinea) chartInstancePorLinea.destroy();

  // Gráfico 1: OEE2 y Capacidad
  const ctxLineas = document.getElementById('chartLineas');
  if (ctxLineas && report.kpisPorLinea && report.kpisPorLinea.length > 0) {
    chartInstanceLineas = new Chart(ctxLineas, {
      type: 'bar',
      data: {
        labels: report.kpisPorLinea.map((l) => l.linea),
        datasets: [
          {
            label: 'Capacidad Utilizada (%)',
            data: report.kpisPorLinea.map((l) => {
              const num = parseNum(l.capacidadUtilizada);
              if (num === null) return 0;
              return num <= 1 && num > 0 ? parseFloat((num * 100).toFixed(1)) : parseFloat(num.toFixed(1));
            }),
            backgroundColor: '#0284c7',
          },
          {
            label: 'OEE2 (%)',
            data: report.kpisPorLinea.map((l) => {
              const num = parseNum(l.oee2);
              if (num === null) return 0;
              return num <= 1 && num > 0 ? parseFloat((num * 100).toFixed(1)) : parseFloat(num.toFixed(1));
            }),
            backgroundColor: '#16a34a',
          },
        ],
      },
      options: { responsive: true, scales: { y: { beginAtZero: true, max: 100 } } },
    });
  }

  // Gráfico 2: Top Desperdicio
  const ctxDesperdicio = document.getElementById('chartDesperdicio');
  if (ctxDesperdicio && report.topDesperdicio && report.topDesperdicio.length > 0) {
    chartInstanceDesperdicio = new Chart(ctxDesperdicio, {
      type: 'bar',
      data: {
        labels: report.topDesperdicio.map((p) => p.descripcion || p.producto || p.codigo),
        datasets: [
          {
            label: 'Desperdicio (kg)',
            data: report.topDesperdicio.map((p) => parseNum(p.desperdicioKg) || 0),
            backgroundColor: '#dc2626',
          },
        ],
      },
      options: { indexAxis: 'y', responsive: true },
    });
  }

  // Gráfico 3: Desperdicio Por Línea
  const ctxPorLinea = document.getElementById('chartPorLinea');
  if (ctxPorLinea && report.desperdicioPorLinea && Object.keys(report.desperdicioPorLinea).length > 0) {
    chartInstancePorLinea = new Chart(ctxPorLinea, {
      type: 'pie',
      data: {
        labels: Object.keys(report.desperdicioPorLinea),
        datasets: [
          {
            data: Object.values(report.desperdicioPorLinea),
            backgroundColor: ['#f59e0b', '#0284c7', '#16a34a', '#8b5cf6', '#ec4899', '#ef4444', '#14b8a6'],
          },
        ],
      },
      options: { responsive: true },
    });
  }
}
