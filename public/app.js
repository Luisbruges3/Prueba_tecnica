// Variable global para almacenar los datos activos del informe
let datosInformeActual = null;

// Instancias globales de gráficos
let chartLineasInstance = null;
let chartDesperdicioInstance = null;
let chartPorLineaInstance = null;

document.addEventListener('DOMContentLoaded', () => {
  // Registrar eventos
  document.getElementById('generarBtn')?.addEventListener('click', cargarExcel);
  document.getElementById('toggleManualBtn')?.addEventListener('click', toggleManualForm);
  document.getElementById('procesarManualBtn')?.addEventListener('click', procesarFormularioManual);
  document.getElementById('printBtn')?.addEventListener('click', () => window.print());
  document.getElementById('filtroLinea')?.addEventListener('change', aplicarFiltroLinea);

  // Botones para agregar filas en captura manual
  document.getElementById('addRowLineaBtn')?.addEventListener('click', () => addRowLinea());
  document.getElementById('addRowTopBtn')?.addEventListener('click', () => addRowTop());
  document.getElementById('addRowMetaBtn')?.addEventListener('click', () => addRowMeta());
  document.getElementById('addRowAccionBtn')?.addEventListener('click', () => addRowAccion());
});

// =========================================================
// FUNCIONES HELPER INTELIGENTES DE LECTURA Y FORMATO
// =========================================================

// Normaliza texto eliminando acentos, espacios y caracteres especiales
function normalizarTexto(txt) {
  if (txt === null || txt === undefined) return '';
  return txt.toString().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

// Busca un valor en un objeto probando varios nombres de columna posibles (sin importar tildes/mayúsculas)
function getProp(obj, candidatos, defaultValue = '') {
  if (!obj || typeof obj !== 'object') return defaultValue;

  const mapNormalizado = {};
  Object.keys(obj).forEach(k => {
    mapNormalizado[normalizarTexto(k)] = obj[k];
  });

  for (let cand of candidatos) {
    const candNorm = normalizarTexto(cand);
    
    // Coincidencia exacta de clave limpia
    if (mapNormalizado[candNorm] !== undefined && mapNormalizado[candNorm] !== null && mapNormalizado[candNorm] !== '') {
      return mapNormalizado[candNorm];
    }

    // Coincidencia parcial (subcadena)
    for (let k in mapNormalizado) {
      if ((k.includes(candNorm) || candNorm.includes(k)) && k !== '') {
        const val = mapNormalizado[k];
        if (val !== undefined && val !== null && val !== '') {
          return val;
        }
      }
    }
  }

  return defaultValue;
}

// Convierte valores a porcentajes limpios (redondeados a 1 decimal)
function parsePorcentaje(val) {
  if (val === undefined || val === null || val === '') return 0;
  let num;
  if (typeof val === 'number') {
    num = val;
  } else {
    let str = val.toString().replace('%', '').replace(',', '.').trim();
    num = parseFloat(str);
  }
  if (isNaN(num)) return 0;

  // Si en Excel viene en formato decimal (ej: 0.85 o 0.275755)
  if (num > 0 && num <= 1) {
    num = num * 100;
  }

  return parseFloat(num.toFixed(1));
}

// Convierte valores a números flotantes limpios (redondeados a 2 decimales)
function parseNumero(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return parseFloat(val.toFixed(2));
  let str = val.toString().replace(/[^0-9.,-]/g, '').replace(',', '.').trim();
  let num = parseFloat(str);
  return isNaN(num) ? 0 : parseFloat(num.toFixed(2));
}

// ==========================================
// 1. CARGA Y PROCESAMIENTO DE EXCEL (SheetJS)
// ==========================================
function cargarExcel() {
  const fileInput = document.getElementById('fileInput');
  const statusSpan = document.getElementById('status');
  const file = fileInput?.files[0];

  if (!file) {
    alert('Por favor selecciona un archivo Excel (.xlsx o .xls)');
    return;
  }

  if (statusSpan) statusSpan.textContent = ' ⏳ Leyendo archivo...';

  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const datosEstructurados = {
        lectura: "Informe cargado exitosamente desde archivo Excel.",
        lineas: [],
        topDesperdicio: [],
        metas: [],
        acciones: []
      };

      // Recorrer todas las hojas del libro Excel
      workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        procesarFilasExcel(rows, datosEstructurados);
      });

      if (statusSpan) statusSpan.textContent = ' ✅ ¡Archivo procesado con éxito!';
      renderizarReporte(datosEstructurados);

    } catch (err) {
      console.error(err);
      if (statusSpan) statusSpan.textContent = ' ❌ Error al leer el archivo Excel.';
      alert('Ocurrió un error al procesar el archivo Excel. Revisa el formato del archivo.');
    }
  };

  reader.readAsArrayBuffer(file);
}

function procesarFilasExcel(rows, datosEstructurados) {
  rows.forEach(r => {
    // 1. Detección de Líneas OEE / Capacidad Utilizada
    const lineaNombre = getProp(r, ['linea', 'lineas', 'nombre linea', 'proceso', 'equipo', 'planta']);
    const capVal = getProp(r, ['capacidad utilizada', 'cap utilizada', 'cap. utilizada', 'capacidad', 'utilizacion', 'cap_utilizada']);
    const oee1Val = getProp(r, ['oee1', 'oee 1', 'oee_1']);
    const oee2Val = getProp(r, ['oee2', 'oee 2', 'oee_2']);

    if (oee1Val !== '' || oee2Val !== '' || capVal !== '') {
      datosEstructurados.lineas.push({
        linea: lineaNombre !== '' ? lineaNombre : 'Línea de Producción',
        capUtilizada: parsePorcentaje(capVal),
        oee1: parsePorcentaje(oee1Val),
        oee2: parsePorcentaje(oee2Val)
      });
    }

    // 2. Detección de Top Desperdicio
    const prodNombre = getProp(r, ['producto', 'descripcion', 'descripcion producto', 'articulo', 'nombre producto']);
    const despKg = getProp(r, ['desperdicio kg', 'desperdicio (kg)', 'kg desperdicio', 'desperdicio']);
    const despPct = getProp(r, ['desperdicio %', 'desperdicio (%)', 'pct desperdicio', '% desperdicio']);
    const codigoVal = getProp(r, ['codigo', 'cod', 'sku', 'ref', 'referencia']);

    if (prodNombre !== '' || (despKg !== '' && parseNumero(despKg) > 0)) {
      datosEstructurados.topDesperdicio.push({
        linea: lineaNombre !== '' ? lineaNombre : 'General',
        codigo: codigoVal !== '' ? codigoVal : 'S/C',
        producto: prodNombre !== '' ? prodNombre : 'Producto',
        desperdicioKg: parseNumero(despKg),
        desperdicioPct: parsePorcentaje(despPct)
      });
    }

    // 3. Detección de Metas
    const kpiNombre = getProp(r, ['indicador', 'kpi', 'meta indicador', 'kpis', 'metricas']);
    const actualVal = getProp(r, ['actual', 'valor actual', 'real', 'ejecutado']);
    const metaVal = getProp(r, ['meta', 'objetivo', 'target']);
    const estadoVal = getProp(r, ['estado', 'cumplimiento', 'status']);

    if (kpiNombre !== '' || (actualVal !== '' && metaVal !== '')) {
      datosEstructurados.metas.push({
        indicador: kpiNombre !== '' ? kpiNombre : 'KPI',
        actual: actualVal !== '' ? actualVal : '0',
        meta: metaVal !== '' ? metaVal : '0',
        estado: estadoVal !== '' ? estadoVal : 'Pendiente'
      });
    }

    // 4. Detección de Acciones Prioritarias
    const accionTexto = getProp(r, ['accion', 'accion a realizar', 'recomendacion', 'tarea', 'actividad', 'acciones']);
    const areaTexto = getProp(r, ['area', 'departamento', 'responsable']);

    if (accionTexto !== '') {
      datosEstructurados.acciones.push({
        area: areaTexto !== '' ? areaTexto : 'General',
        accion: accionTexto
      });
    }
  });
}

// ==========================================
// 2. LOGICA DE FORMULARIO DE CAPTURA MANUAL
// ==========================================
function toggleManualForm() {
  const form = document.getElementById('manualFormSection');
  if (!form) return;
  
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

function addRowLinea(linea = '', cap = 0, oee1 = 0, oee2 = 0) {
  const tbody = document.querySelector('#tableInputLineas tbody');
  if (!tbody) return;
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" value="${linea}" placeholder="Nombre línea" style="width:95%"></td>
    <td><input type="number" value="${cap}" style="width:90%"></td>
    <td><input type="number" value="${oee1}" style="width:90%"></td>
    <td><input type="number" value="${oee2}" style="width:90%"></td>
    <td><button type="button" onclick="this.closest('tr').remove()" style="color:red; cursor:pointer;">X</button></td>
  `;
  tbody.appendChild(tr);
}

function addRowTop(linea = '', codigo = '', producto = '', kg = 0, pct = 0) {
  const tbody = document.querySelector('#tableInputTop tbody');
  if (!tbody) return;
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" value="${linea}" placeholder="Línea" style="width:95%"></td>
    <td><input type="text" value="${codigo}" placeholder="Código" style="width:95%"></td>
    <td><input type="text" value="${producto}" placeholder="Nombre producto" style="width:95%"></td>
    <td><input type="number" value="${kg}" style="width:90%"></td>
    <td><input type="number" value="${pct}" step="0.1" style="width:90%"></td>
    <td><button type="button" onclick="this.closest('tr').remove()" style="color:red; cursor:pointer;">X</button></td>
  `;
  tbody.appendChild(tr);
}

function addRowMeta(kpi = '', actual = '', meta = '', estado = 'Cumple') {
  const tbody = document.querySelector('#tableInputMetas tbody');
  if (!tbody) return;
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
    <td><button type="button" onclick="this.closest('tr').remove()" style="color:red; cursor:pointer;">X</button></td>
  `;
  tbody.appendChild(tr);
}

function addRowAccion(area = '', accion = '') {
  const tbody = document.querySelector('#tableInputAcciones tbody');
  if (!tbody) return;
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" value="${area}" placeholder="Área" style="width:95%"></td>
    <td><input type="text" value="${accion}" placeholder="Acción a tomar" style="width:95%"></td>
    <td><button type="button" onclick="this.closest('tr').remove()" style="color:red; cursor:pointer;">X</button></td>
  `;
  tbody.appendChild(tr);
}

function procesarFormularioManual() {
  const lectura = document.getElementById('inputLectura')?.value || "Resumen de jornada ingresado manualmente.";

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
      estado: select ? select.value : 'Pendiente'
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
  const reportElem = document.getElementById('report');
  if (reportElem) reportElem.classList.remove('hidden');

  // Fecha actual
  const fechaElem = document.getElementById('fecha');
  if (fechaElem) fechaElem.textContent = `Fecha de emisión: ${new Date().toLocaleDateString('es-CO')}`;

  // Lectura Ejecutiva
  const lecturaElem = document.getElementById('lecturaTexto');
  if (lecturaElem) lecturaElem.textContent = datos.lectura;

  // Llenar Filtro de Líneas
  const filtroSelect = document.getElementById('filtroLinea');
  if (filtroSelect) {
    filtroSelect.innerHTML = '<option value="TODAS">-- Todas las Líneas --</option>';
    const lineasUnicas = [...new Set(datos.lineas.map(l => l.linea))];
    lineasUnicas.forEach(l => {
      const opt = document.createElement('option');
      opt.value = l;
      opt.textContent = l;
      filtroSelect.appendChild(opt);
    });
  }

  // Renderizar secciones
  poblarTablaMetas(datos.metas);
  poblarTablaYGraficoLineas(datos.lineas);
  poblarTablaYGraficosDesperdicio(datos.topDesperdicio);
  poblarAcciones(datos.acciones);

  if (reportElem) reportElem.scrollIntoView({ behavior: 'smooth' });
}

function poblarTablaMetas(metas) {
  const tbody = document.querySelector('#tablaCumplimiento tbody');
  if (!tbody) return;
  tbody.innerHTML = '';

  if (!metas || metas.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" style="text-align:center; color:#666;">No hay metas registradas en el archivo.</td></tr>';
    return;
  }

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
  if (!tbody) return;
  tbody.innerHTML = '';

  const labels = [];
  const capData = [];
  const oee2Data = [];

  if (!lineas || lineas.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" style="text-align:center; color:#666;">No hay líneas registradas en el archivo.</td></tr>';
  } else {
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
  }

  // Renderizar Gráfico
  const canvas = document.getElementById('chartLineas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
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
  if (!tbody) return;
  tbody.innerHTML = '';

  const labels = [];
  const kgData = [];
  const lineaTotales = {};

  if (!top || top.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; color:#666;">No hay datos de desperdicio en el archivo.</td></tr>';
  } else {
    top.forEach((item, index) => {
      labels.push(item.producto);
      kgData.push(item.desperdicioKg);

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
  }

  // Gráfico 1: Top Productos (Barras Horizontales)
  const canvas1 = document.getElementById('chartDesperdicio');
  if (canvas1) {
    const ctx1 = canvas1.getContext('2d');
    if (chartDesperdicioInstance) chartDesperdicioInstance.destroy();

    chartDesperdicioInstance = new Chart(ctx1, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [{ label: 'Desperdicio (kg)', data: kgData, backgroundColor: '#dc2626' }]
      },
      options: { indexAxis: 'y', responsive: true }
    });
  }

  // Gráfico 2: Desperdicio por Línea (Pie Chart)
  const canvas2 = document.getElementById('chartPorLinea');
  if (canvas2) {
    const ctx2 = canvas2.getContext('2d');
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
}

function poblarAcciones(acciones) {
  const lista = document.getElementById('listaAcciones');
  if (!lista) return;
  lista.innerHTML = '';

  if (!acciones || acciones.length === 0) {
    lista.innerHTML = '<li style="color:#666;">No hay acciones registradas en el archivo.</li>';
    return;
  }

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
  const seleccion = document.getElementById('filtroLinea')?.value;

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
