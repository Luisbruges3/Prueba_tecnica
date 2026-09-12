// Variable global para almacenar los datos activos del informe
let datosInformeActual = null;

// Instancias globales de gráficos para destruirlos antes de volver a renderizar
let chartLineasInstance = null;
let chartDesperdicioInstance = null;
let chartPorLineaInstance = null;

document.addEventListener('DOMContentLoaded', () => {
  // Inicializar eventos principales
  document.getElementById('generarBtn').addEventListener('click', cargarExcel);
  document.getElementById('toggleManualBtn').addEventListener('click', toggleManualForm);
  document.getElementById('procesarManualBtn').addEventListener('click', procesarFormularioManual);
  document.getElementById('printBtn').addEventListener('click', () => window.print());
  document.getElementById('filtroLinea').addEventListener('change', aplicarFiltroLinea);

  // Eventos para agregar filas a las tablas de captura manual
  document.getElementById('addRowLineaBtn').addEventListener('click', () => addRowLinea());
  document.getElementById('addRowTopBtn').addEventListener('click', () => addRowTop());
  document.getElementById('addRowMetaBtn').addEventListener('click', () => addRowMeta());
  document.getElementById('addRowAccionBtn').addEventListener('click', () => addRowAccion());
});

// ==========================================
// 1. CARGA Y PROCESAMIENTO DE EXCEL (SheetJS)
// ==========================================
function cargarExcel() {
  const fileInput = document.getElementById('fileInput');
  const statusSpan = document.getElementById('status');
  const file = fileInput.files[0];

  if (!file) {
    alert('Por favor selecciona un archivo Excel (.xlsx o .xls)');
    return;
  }

  statusSpan.textContent = ' ⏳ Leyendo archivo...';

  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      // Estructura base para el reporte
      const datosEstructurados = {
        lectura: "Informe cargado exitosamente desde archivo Excel.",
        lineas: [],
        topDesperdicio: [],
        metas: [],
        acciones: []
      };

      // Si el Excel tiene hojas específicas, las leemos; de lo contrario, leemos la primera hoja
      const sheetNames = workbook.SheetNames;

      sheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet);

        const nombreNormalizado = sheetName.toLowerCase();

        if (nombreNormalizado.includes('linea') || nombreNormalizado.includes('oee')) {
          datosEstructurados.lineas = rows.map(r => ({
            linea: r.Linea || r.Línea || r.linea || 'Línea Sin Nombre',
            capUtilizada: parseFloat(r.CapUtilizada || r['Cap. Utilizada (%)'] || r.capUtilizada || 0),
            oee1: parseFloat(r.OEE1 || r['OEE1 (%)'] || r.oee1 || 0),
            oee2: parseFloat(r.OEE2 || r['OEE2 (%)'] || r.oee2 || 0)
          }));
        } else if (nombreNormalizado.includes('desperdicio') || nombreNormalizado.includes('top')) {
          datosEstructurados.topDesperdicio = rows.map(r => ({
            linea: r.Linea || r.Línea || r.linea || 'General',
            codigo: r.Codigo || r.Código || r.codigo || 'S/C',
            producto: r.Producto || r.Descripción || r.producto || 'Producto',
            desperdicioKg: parseFloat(r.DesperdicioKg || r['Desperdicio (kg)'] || r.desperdicioKg || 0),
            desperdicioPct: parseFloat(r.DesperdicioPct || r['Desperdicio (%)'] || r.desperdicioPct || 0)
          }));
        } else if (nombreNormalizado.includes('meta') || nombreNormalizado.includes('kpi')) {
          datosEstructurados.metas = rows.map(r => ({
            indicador: r.Indicador || r.KPI || r.indicador || 'KPI',
            actual: r.Actual || r.actual || '0',
            meta: r.Meta || r.meta || '0',
            estado: r.Estado || r.estado || 'Pendiente'
          }));
        } else if (nombreNormalizado.includes('accion')) {
          datosEstructurados.acciones = rows.map(r => ({
            area: r.Area || r.Área || r.area || 'General',
            accion: r.Accion || r.Acción || r.accion || ''
          }));
        }
      });

      // Si fue una sola hoja plana, intentamos extraer los datos directamente
      if (datosEstructurados.lineas.length === 0 && sheetNames.length > 0) {
        const firstSheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNames[0]]);
        datosEstructurados.lineas = firstSheet.map(r => ({
          linea: r.Linea || r.Línea || 'Línea',
          capUtilizada: parseFloat(r.CapUtilizada || r['Cap. Utilizada (%)'] || 0),
          oee1: parseFloat(r.OEE1 || 0),
          oee2: parseFloat(r.OEE2 || 0)
        }));
      }

      statusSpan.textContent = ' ✅ ¡Archivo procesado!';
      renderizarReporte(datosEstructurados);

    } catch (err) {
      console.error(err);
      statusSpan.textContent = ' ❌ Error al leer el archivo Excel.';
      alert('Ocurrió un error al procesar el archivo Excel. Asegúrate de que tenga un formato válido.');
    }
  };

  reader.readAsArrayBuffer(file);
}

// ==========================================
// 2. LOGICA DE FORMULARIO DE CAPTURA MANUAL
// ==========================================
function toggleManualForm() {
  const form = document.getElementById('manualFormSection');
  
  if (form.style.display === 'none' || form.style.display === '') {
    form.style.display = 'block';

    // Precargar filas iniciales de guía si está vacía
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

function addRowLinea(linea = '', cap = 0, oee1 = 0, oee2 = 0) {
  const tbody = document.querySelector('#tableInputLineas tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" value="${linea}" placeholder="Nombre línea" style="width:95%"></td>
    <td><input type="number" value="${cap}" style="width:90%"></td>
    <td><input type="number" value="${oee1}" style="width:90%"></td>
    <td><input type="number" value="${oee2}" style="width:90%"></td>
    <td><button onclick="this.closest('tr').remove()" style="color:red; cursor:pointer;">X</button></td>
  `;
  tbody.appendChild(tr);
}

function addRowTop(linea = '', codigo = '', producto = '', kg = 0, pct = 0) {
  const tbody = document.querySelector('#tableInputTop tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" value="${linea}" placeholder="Línea" style="width:95%"></td>
    <td><input type="text" value="${codigo}" placeholder="Código" style="width:95%"></td>
    <td><input type="text" value="${producto}" placeholder="Nombre producto" style="width:95%"></td>
    <td><input type="number" value="${kg}" style="width:90%"></td>
    <td><input type="number" value="${pct}" step="0.1" style="width:90%"></td>
    <td><button onclick="this.closest('tr').remove()" style="color:red; cursor:pointer;">X</button></td>
  `;
  tbody.appendChild(tr);
}

function addRowMeta(kpi = '', actual = '', meta = '', estado = 'Cumple') {
  const tbody = document.querySelector('#tableInputMetas tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" value="${kpi}" placeholder="KPI" style="width:95%"></td>
    <td><input type="text" value="${actual}" placeholder="Actual" style="width:95%"></td>
    <td><input type="text" value="${meta}" placeholder="Meta" style="width:95%"></td>
    <td>
      <select style="width:95%">
        <option value="Cumple" ${estado === 'Cumple' ? 'selected' : ''}>Cumple</option>
        <option value="No Cumple" ${estado === 'No Cumple' ? 'selected' : ''}>No Cumple</option>
        <option value="En Riesgo" ${estado === 'En Riesgo' ? 'selected' : ''}>En Riesgo</option>
      </select>
    </td>
    <td><button onclick="this.closest('tr').remove()" style="color:red; cursor:pointer;">X</button></td>
  `;
  tbody.appendChild(tr);
}

function addRowAccion(area = '', accion = '') {
  const tbody = document.querySelector('#tableInputAcciones tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" value="${area}" placeholder="Área" style="width:95%"></td>
    <td><input type="text" value="${accion}" placeholder="Acción a tomar" style="width:95%"></td>
    <td><button onclick="this.closest('tr').remove()" style="color:red; cursor:pointer;">X</button></td>
  `;
  tbody.appendChild(tr);
}

function procesarFormularioManual() {
  const lectura = document.getElementById('inputLectura').value || "Resumen de jornada ingresado manualmente.";

  const lineas = Array.from(document.querySelectorAll('#tableInputLineas tbody tr')).map(tr => {
    const inputs = tr.querySelectorAll('input');
    return {
      linea: inputs[0].value || 'Línea',
      capUtilizada: parseFloat(inputs[1].value) || 0,
      oee1: parseFloat(inputs[2].value) || 0,
      oee2: parseFloat(inputs[3].value) || 0
    };
  });

  const topDesperdicio = Array.from(document.querySelectorAll('#tableInputTop tbody tr')).map(tr => {
    const inputs = tr.querySelectorAll('input');
    return {
      linea: inputs[0].value || 'General',
      codigo: inputs[1].value || 'S/C',
      producto: inputs[2].value || 'Producto',
      desperdicioKg: parseFloat(inputs[3].value) || 0,
      desperdicioPct: parseFloat(inputs[4].value) || 0
    };
  });

  const metas = Array.from(document.querySelectorAll('#tableInputMetas tbody tr')).map(tr => {
    const inputs = tr.querySelectorAll('input');
    const select = tr.querySelector('select');
    return {
      indicador: inputs[0].value || 'KPI',
      actual: inputs[1].value || '0',
      meta: inputs[2].value || '0',
      estado: select.value
    };
  });

  const acciones = Array.from(document.querySelectorAll('#tableInputAcciones tbody tr')).map(tr => {
    const inputs = tr.querySelectorAll('input');
    return {
      area: inputs[0].value || 'General',
      accion: inputs[1].value || ''
    };
  });

  renderizarReporte({ lectura, lineas, topDesperdicio, metas, acciones });
}

// ==========================================
// 3. RENDERIZADO DEL INFORME Y GRÁFICOS
// ==========================================
function renderizarReporte(datos) {
  datosInformeActual = datos;
  document.getElementById('report').classList.remove('hidden');

  // Fecha actual
  document.getElementById('fecha').textContent = `Fecha de emisión: ${new Date().toLocaleDateString('es-CO')}`;

  // Lectura Ejecutiva
  document.getElementById('lecturaTexto').textContent = datos.lectura;

  // Llenar Filtro de Líneas
  const filtroSelect = document.getElementById('filtroLinea');
  filtroSelect.innerHTML = '<option value="TODAS">-- Todas las Líneas --</option>';
  const lineasUnicas = [...new Set(datos.lineas.map(l => l.linea))];
  lineasUnicas.forEach(l => {
    const opt = document.createElement('option');
    opt.value = l;
    opt.textContent = l;
    filtroSelect.appendChild(opt);
  });

  // Renderizar secciones
  poblarTablaMetas(datos.metas);
  poblarTablaYGraficoLineas(datos.lineas);
  poblarTablaYGraficosDesperdicio(datos.topDesperdicio);
  poblarAcciones(datos.acciones);

  // Desplazar vista hacia el reporte
  document.getElementById('report').scrollIntoView({ behavior: 'smooth' });
}

function poblarTablaMetas(metas) {
  const tbody = document.querySelector('#tablaCumplimiento tbody');
  tbody.innerHTML = '';
  metas.forEach(m => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${m.indicador}</td>
      <td>${m.actual}</td>
      <td>${m.meta}</td>
      <td><strong>${m.estado}</strong></td>
    `;
    tbody.appendChild(tr);
  });
}

function poblarTablaYGraficoLineas(lineas) {
  const tbody = document.querySelector('#tablaLineas tbody');
  tbody.innerHTML = '';

  const labels = [];
  const capData = [];
  const oee2Data = [];

  lineas.forEach(l => {
    labels.push(l.linea);
    capData.push(l.capUtilizada);
    oee2Data.push(l.oee2);

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${l.linea}</td>
      <td>${l.capUtilizada}%</td>
      <td>${l.oee1}%</td>
      <td>${l.oee2}%</td>
    `;
    tbody.appendChild(tr);
  });

  // Renderizar Gráfico
  const ctx = document.getElementById('chartLineas').getContext('2d');
  if (chartLineasInstance) chartLineasInstance.destroy();

  chartLineasInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        { label: 'Capacidad Utilizada (%)', data: capData, backgroundColor: '#0284c7' },
        { label: 'OEE2 (%)', data: oee2Data, backgroundColor: '#16a34a' }
      ]
    },
    options: {
      responsive: true,
      scales: { y: { beginAtZero: true, max: 100 } }
    }
  });
}

function poblarTablaYGraficosDesperdicio(top) {
  const tbody = document.querySelector('#tablaTop tbody');
  tbody.innerHTML = '';

  const labels = [];
  const kgData = [];
  const lineaTotales = {};

  top.forEach((item, index) => {
    labels.push(item.producto);
    kgData.push(item.desperdicioKg);

    // Acumular desperdicio por línea para el gráfico circular
    lineaTotales[item.linea] = (lineaTotales[item.linea] || 0) + item.desperdicioKg;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${index + 1}</td>
      <td>${item.linea}</td>
      <td>${item.codigo}</td>
      <td>${item.producto}</td>
      <td>${item.desperdicioKg} kg</td>
      <td>${item.desperdicioPct}%</td>
    `;
    tbody.appendChild(tr);
  });

  // Gráfico 1: Top Productos (Barras Horizontales)
  const ctx1 = document.getElementById('chartDesperdicio').getContext('2d');
  if (chartDesperdicioInstance) chartDesperdicioInstance.destroy();

  chartDesperdicioInstance = new Chart(ctx1, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{ label: 'Desperdicio (kg)', data: kgData, backgroundColor: '#dc2626' }]
    },
    options: { indexAxis: 'y', responsive: true }
  });

  // Gráfico 2: Desperdicio por Línea (Pie Chart)
  const ctx2 = document.getElementById('chartPorLinea').getContext('2d');
  if (chartPorLineaInstance) chartPorLineaInstance.destroy();

  chartPorLineaInstance = new Chart(ctx2, {
    type: 'pie',
    data: {
      labels: Object.keys(lineaTotales),
      datasets: [{
        data: Object.values(lineaTotales),
        backgroundColor: ['#f87171', '#fbbf24', '#60a5fa', '#34d399', '#a78bfa']
      }]
    },
    options: { responsive: true }
  });
}

function poblarAcciones(acciones) {
  const lista = document.getElementById('listaAcciones');
  lista.innerHTML = '';
  acciones.forEach(a => {
    const li = document.createElement('li');
    li.innerHTML = `<strong>[${a.area}]</strong> ${a.accion}`;
    lista.appendChild(li);
  });
}

// ==========================================
// 4. FILTRADO DINÁMICO
// ==========================================
function aplicarFiltroLinea() {
  if (!datosInformeActual) return;
  const seleccion = document.getElementById('filtroLinea').value;

  if (seleccion === 'TODAS') {
    poblarTablaYGraficoLineas(datosInformeActual.lineas);
    poblarTablaYGraficosDesperdicio(datosInformeActual.topDesperdicio);
  } else {
    const lineasFiltradas = datosInformeActual.lineas.filter(l => l.linea === seleccion);
    const topFiltrado = datosInformeActual.topDesperdicio.filter(t => t.linea === seleccion);

    poblarTablaYGraficoLineas(lineasFiltradas);
    poblarTablaYGraficosDesperdicio(topFiltrado);
  }
}
