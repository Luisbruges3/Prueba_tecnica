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
  
  // Extraer el número limpio sin importar si viene como string con '%' o número decimal
  let num = typeof v === 'number' ? v : parseFloat(String(v).replace('%', '').replace(',', '.').trim());
  if (isNaN(num)) return String(v);

  // Convertir decimales puros de Excel (ej: 0.8638 -> 86.4%)
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
    if (!txt) return true;
    const norm = String(txt).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
    return ['total', 'totales', 'promedio', 'subtotal', 'linea de produccion', 'gran total', 'resumen'].includes(norm) || norm.startsWith('total');
  };

  // Filtrar líneas duplicadas o con nombre de total
  const kpisPorLinea = (report.kpisPorLinea || []).filter((item, idx, self) => {
    if (!item || !item.linea || esIgnorable(item.linea)) return false;
    return self.findIndex(t => String(t.linea).trim().toLowerCase() === String(item.linea).trim().toLowerCase()) === idx;
  });

  // Filtrar productos duplicados o inválidos
  const topDesperdicio = (report.topDesperdicio || []).filter((item, idx, self) => {
    const nombre = item.descripcion || item.producto || item.codigo;
    if (!nombre || esIgnorable(nombre) || esIgnorable(item.linea)) return false;
    return self.findIndex(t => (String(t.linea) + String(t.codigo) + String(t.descripcion)).toLowerCase() === (String(item.linea) + String(item.codigo) + String(item.descripcion)).toLowerCase()) === idx;
  });

  return {
    ...report,
    kpisPorLinea,
    topDesperdicio
  };
}

// ---------- 1. Mostrar / Ocultar Formulario Manual ----------
function toggleManualForm() {
  const form = document.getElementById('manualFormSection');
  
  if (form.style.display === 'none' || form.style.display === '') {
    form.style.display = 'block';

    if (document.querySelectorAll('#tableInputLineas tbody tr').length === 0) {
      addRowLinea('Línea 1 - Pan Tajado', 85, 78, 74);
      addRowLinea('Línea 2 - Hamburgo', 78, 72, 69);
      addRowTop('Línea 1 - Pan Tajado', 'P001', 'Pan Blanco 500g', 120, 3.5);
      addRowMeta('OEE2 Global', '72.5%', '75.0%', 'No Cumple');
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

// ---------- 4. Cargar desde Archivo Excel ----------
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
      const desperdicioPorLinea = {};

      workbook.SheetNames.forEach(sheetName => {
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });

        rows.forEach(r => {
          // Extraer claves sin importar mayúsculas/tildes
          const keys = Object.keys(r);
          const getVal = (candidatos) => {
            const key = keys.find(k => candidatos.some(c => k.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").includes(c)));
            return key ? r[key] : '';
          };

          const linea = getVal(['linea', 'proceso']);
          const cap = parseNum(getVal(['capacidad', 'utilizacion']));
          const oee1 = parseNum(getVal(['oee1', 'oee 1']));
          const oee2 = parseNum(getVal(['oee2', 'oee 2']));

          if (linea && (cap !== null || oee2 !== null)) {
            kpisPorLinea.push({ linea: String(linea).trim(), capacidadUtilizada: cap, oee1, oee2 });
          }

          const prod = getVal(['producto', 'descripcion', 'articulo']);
          const cod = getVal(['codigo', 'sku', 'ref']);
          const kg = parseNum(getVal(['desperdicio kg', 'kg desperdicio', 'desperdicio']));
          const pct = parseNum(getVal(['desperdicio %', '% desperdicio']));

          if (prod && kg !== null) {
            const lNom = linea ? String(linea).trim() : 'General';
            topDesperdicio.push({
              linea: lNom,
              codigo: cod ? String(cod).trim() : 'S/C',
              descripcion: String(prod).trim(),
              desperdicioKg: kg,
              desperdicioPct: pct
            });
            desperdicioPorLinea[lNom] = (desperdicioPorLinea[lNom] || 0) + kg;
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
        desperdicioPorLinea,
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
  if (ctxPorLinea && report.desperdicioPorLinea) {
    chartInstancePorLinea = new Chart(ctxPorLinea, {
      type: 'pie',
      data: {
        labels: Object.keys(report.desperdicioPorLinea),
        datasets: [
          {
            data: Object.values(report.desperdicioPorLinea),
            backgroundColor: ['#f59e0b', '#0284c7', '#16a34a', '#8b5cf6', '#ec4899'],
          },
        ],
      },
      options: { responsive: true },
    });
  }
}
