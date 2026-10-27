// Conciliador Bancario — todo el procesamiento ocurre en el navegador.

const state = {
  mayor: null, // { rows, headers }
  banco: null,
  resultado: [],
  filtro: 'all',
};

const CAMPOS = [
  { key: 'fecha', label: 'Fecha', pistas: ['fecha', 'fec', 'date', 'f. valor', 'fecha valor', 'fecha operacion'] },
  { key: 'desc', label: 'Descripción', pistas: ['descripcion', 'concepto', 'detalle', 'referencia', 'leyenda', 'movimiento'] },
  { key: 'monto', label: 'Monto', pistas: ['monto', 'importe', 'valor', 'total', 'amount'] },
  { key: 'debito', label: 'Débito (opcional)', pistas: ['debito', 'debe', 'egreso', 'salida'] },
  { key: 'credito', label: 'Crédito (opcional)', pistas: ['credito', 'haber', 'ingreso', 'entrada'] },
];

const $ = (id) => document.getElementById(id);
const norm = (s) => String(s ?? '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();

/* ---------- Lectura de archivos ---------- */

function leerArchivo(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
        const hoja = wb.Sheets[wb.SheetNames[0]];
        const matriz = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: '', raw: true });
        resolve(armarTabla(matriz));
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// Los extractos suelen traer filas de título antes del encabezado: buscamos la primera
// fila que parezca un encabezado (varias celdas de texto, alguna con pista conocida).
function armarTabla(matriz) {
  const todasPistas = CAMPOS.flatMap((c) => c.pistas);
  let idx = 0;
  for (let i = 0; i < Math.min(matriz.length, 30); i++) {
    const celdas = matriz[i].map(norm);
    if (celdas.filter(Boolean).length >= 2 && celdas.some((c) => todasPistas.some((p) => c.includes(p)))) {
      idx = i;
      break;
    }
  }
  const headers = matriz[idx].map((h, i) => String(h).trim() || `Columna ${i + 1}`);
  const rows = matriz.slice(idx + 1)
    .filter((r) => r.some((c) => String(c).trim() !== ''))
    .map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i]])));
  return { headers, rows };
}

function adivinarColumna(headers, pistas) {
  for (const p of pistas) {
    const h = headers.find((x) => norm(x) === p);
    if (h) return h;
  }
  for (const p of pistas) {
    const h = headers.find((x) => norm(x).includes(p));
    if (h) return h;
  }
  return '';
}

/* ---------- Normalización ---------- */

function parseMonto(v) {
  if (typeof v === 'number') return v;
  let s = String(v ?? '').trim();
  if (!s) return NaN;
  const negativo = /^\(.*\)$/.test(s) || s.includes('-');
  s = s.replace(/[^\d.,]/g, '');
  const ultComa = s.lastIndexOf(',');
  const ultPunto = s.lastIndexOf('.');
  if (ultComa > ultPunto) s = s.replace(/\./g, '').replace(',', '.'); // 1.234,56
  else s = s.replace(/,/g, '');                                        // 1,234.56
  const n = parseFloat(s);
  return isNaN(n) ? NaN : (negativo ? -Math.abs(n) : n);
}

function parseFecha(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'number') { // serial de Excel
    const d = XLSX.SSF.parse_date_code(v);
    return d ? new Date(d.y, d.m - 1, d.d) : null;
  }
  const s = String(v ?? '').trim();
  let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/); // dd/mm/aaaa
  if (m) {
    const y = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    return new Date(y, +m[2] - 1, +m[1]);
  }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); // aaaa-mm-dd
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

const fmtFecha = (d) => d ? d.toLocaleDateString('es-AR') : '—';
const fmtMonto = (n) => isNaN(n) ? '—' : n.toLocaleString('es-AR', { style: 'currency', currency: 'ARS' });

function extraerMovimientos(tabla, mapa) {
  return tabla.rows.map((r, i) => {
    let monto = mapa.monto ? parseMonto(r[mapa.monto]) : NaN;
    if (isNaN(monto) && (mapa.debito || mapa.credito)) {
      const deb = parseMonto(r[mapa.debito]) || 0;
      const cre = parseMonto(r[mapa.credito]) || 0;
      monto = cre - Math.abs(deb);
    }
    return {
      fila: i + 1,
      fecha: parseFecha(r[mapa.fecha]),
      desc: String(r[mapa.desc] ?? '').trim(),
      monto,
    };
  }).filter((m) => !isNaN(m.monto) && m.monto !== 0);
}

/* ---------- Conciliación ---------- */

// Para cada movimiento del mayor buscamos el mejor candidato del banco todavía libre:
// mismo monto (± tolerancia) y la fecha más cercana dentro de la ventana de días.
function conciliar(mayor, banco, tolMonto, tolDias) {
  const usados = new Set();
  const out = [];
  const DIA = 86400000;

  for (const m of mayor) {
    let mejor = -1, mejorDist = Infinity;
    banco.forEach((b, j) => {
      if (usados.has(j)) return;
      if (Math.abs(b.monto - m.monto) > tolMonto) return;
      const dist = m.fecha && b.fecha ? Math.abs(b.fecha - m.fecha) / DIA : 0;
      if (dist <= tolDias && dist < mejorDist) { mejor = j; mejorDist = dist; }
    });
    if (mejor >= 0) {
      usados.add(mejor);
      const b = banco[mejor];
      const nota = mejorDist > 0 ? `Banco: ${fmtFecha(b.fecha)} (${Math.round(mejorDist)} d)` : 'Coincidencia exacta';
      out.push({ ...m, origen: 'Ambos', estado: 'ok', nota });
    } else {
      out.push({ ...m, origen: 'Mayor', estado: 'mayor', nota: 'No figura en el extracto' });
    }
  }
  banco.forEach((b, j) => {
    if (!usados.has(j)) out.push({ ...b, origen: 'Banco', estado: 'banco', nota: 'No registrado en el mayor' });
  });
  return out.sort((a, b) => (a.fecha ?? 0) - (b.fecha ?? 0));
}

/* ---------- UI ---------- */

const ico = (id) => `<svg><use href="#i-${id}"/></svg>`;
const ETIQUETAS = { ok: 'Conciliado', mayor: 'Solo en mayor', banco: 'Solo en banco' };
const ICONO_ESTADO = { ok: 'check', mayor: 'alert', banco: 'x' };
const ICONO_ORIGEN = { Ambos: 'check', Mayor: 'book', Banco: 'bank' };

let toastTimer;
function aviso(msg) {
  const t = $('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 3500);
}

function marcarPaso(n) {
  document.querySelectorAll('.steps li').forEach((li) => {
    const k = +li.dataset.step;
    li.classList.toggle('active', k === n);
    li.classList.toggle('done', k < n);
  });
}

function configurarCarga(tipo, inputId, dropId, hintId) {
  const input = $(inputId), drop = $(dropId), hint = $(hintId);
  const cargar = async (file) => {
    if (!file) return;
    drop.classList.remove('loaded', 'error');
    hint.innerHTML = `${ico('file')}Leyendo…`;
    try {
      state[tipo] = await leerArchivo(file);
      if (!state[tipo].rows.length) throw new Error('vacío');
      hint.innerHTML = `${ico('check')}${escapar(file.name)} · ${state[tipo].rows.length} filas`;
      drop.classList.add('loaded');
      renderMapeo();
    } catch {
      state[tipo] = null;
      hint.innerHTML = `${ico('alert')}No pudimos leer el archivo`;
      drop.classList.add('error');
    }
  };
  input.addEventListener('change', (e) => cargar(e.target.files[0]));
  drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('over'));
  drop.addEventListener('drop', (e) => {
    e.preventDefault();
    drop.classList.remove('over');
    cargar(e.dataTransfer.files[0]);
  });
}

function renderMapeo() {
  if (!state.mayor || !state.banco) return;
  for (const [tipo, contId] of [['mayor', 'mapMayor'], ['banco', 'mapBanco']]) {
    const cont = $(contId);
    cont.querySelectorAll('label').forEach((l) => l.remove());
    const { headers } = state[tipo];
    for (const c of CAMPOS) {
      const label = document.createElement('label');
      label.textContent = c.label;
      const sel = document.createElement('select');
      sel.dataset.tipo = tipo;
      sel.dataset.campo = c.key;
      sel.innerHTML = '<option value="">— Sin asignar —</option>' +
        headers.map((h) => `<option>${escapar(h)}</option>`).join('');
      sel.value = adivinarColumna(headers, c.pistas);
      sel.classList.toggle('filled', !!sel.value);
      sel.addEventListener('change', () => sel.classList.toggle('filled', !!sel.value));
      label.appendChild(sel);
      cont.appendChild(label);
    }
  }
  $('mapping').hidden = false;
  marcarPaso(2);
  $('mapping').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function leerMapa(tipo) {
  const mapa = {};
  document.querySelectorAll(`select[data-tipo="${tipo}"]`).forEach((s) => { mapa[s.dataset.campo] = s.value; });
  return mapa;
}

function escapar(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function renderResultados() {
  const r = state.resultado;
  const cuenta = (e) => r.filter((x) => x.estado === e).length;
  const suma = (e) => r.filter((x) => x.estado === e).reduce((a, x) => a + x.monto, 0);
  const ok = cuenta('ok'), soloM = cuenta('mayor'), soloB = cuenta('banco');
  const totalMayor = ok + soloM;
  const pct = totalMayor ? Math.round((ok / totalMayor) * 100) : 0;

  $('stOk').textContent = ok;
  $('stMayor').textContent = soloM;
  $('stBanco').textContent = soloB;
  $('stDif').textContent = fmtMonto(suma('mayor') - suma('banco'));
  $('cAll').textContent = r.length;
  $('cOk').textContent = ok;
  $('cMayor').textContent = soloM;
  $('cBanco').textContent = soloB;
  $('barOk').style.width = pct + '%';
  $('resumen').textContent = soloM + soloB === 0
    ? `Todo coincide: ${ok} movimientos conciliados.`
    : `${pct}% del libro mayor está conciliado. Quedan ${soloM + soloB} movimientos por revisar.`;

  const filas = state.filtro === 'all' ? r : r.filter((x) => x.estado === state.filtro);
  $('tbody').innerHTML = filas.map((x) => `
    <tr class="st-${x.estado}">
      <td>${fmtFecha(x.fecha)}</td>
      <td>${escapar(x.desc || '—')}</td>
      <td class="r mono ${x.monto < 0 ? 'neg' : ''}">${fmtMonto(x.monto)}</td>
      <td><span class="origen">${ico(ICONO_ORIGEN[x.origen])}${x.origen}</span></td>
      <td><span class="pill ${x.estado}">${ico(ICONO_ESTADO[x.estado])}${ETIQUETAS[x.estado]}</span><small>${escapar(x.nota)}</small></td>
    </tr>`).join('') ||
    `<tr><td colspan="5" class="empty">${ico('inbox')}No hay movimientos en esta vista.</td></tr>`;
  $('results').hidden = false;
  marcarPaso(3);
}

$('btnConciliar').addEventListener('click', () => {
  const mapaM = leerMapa('mayor'), mapaB = leerMapa('banco');
  for (const [nombre, mapa] of [['libro mayor', mapaM], ['extracto', mapaB]]) {
    if (!mapa.fecha || (!mapa.monto && !mapa.debito && !mapa.credito)) {
      aviso(`Falta elegir la fecha y el monto (o débito/crédito) del ${nombre}.`);
      return;
    }
  }
  const mov = extraerMovimientos(state.mayor, mapaM);
  const ext = extraerMovimientos(state.banco, mapaB);
  if (!mov.length || !ext.length) {
    aviso('No encontramos montos válidos. Revisá que la columna de monto sea la correcta.');
    return;
  }
  const tolM = Math.max(0, parseFloat($('tolMonto').value) || 0);
  const tolD = Math.max(0, parseInt($('tolDias').value, 10) || 0);
  state.resultado = conciliar(mov, ext, tolM, tolD);
  state.filtro = 'all';
  document.querySelectorAll('.tab').forEach((x) => x.classList.toggle('active', x.dataset.f === 'all'));
  renderResultados();
  $('results').scrollIntoView({ behavior: 'smooth', block: 'start' });
});

document.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => {
  document.querySelectorAll('.tab').forEach((x) => x.classList.remove('active'));
  t.classList.add('active');
  state.filtro = t.dataset.f;
  renderResultados();
}));

$('btnExportar').addEventListener('click', () => {
  const datos = state.resultado.map((x) => ({
    Fecha: x.fecha ?? '',
    Descripción: x.desc,
    Monto: x.monto,
    Origen: x.origen,
    Estado: ETIQUETAS[x.estado],
    Observación: x.nota,
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(datos, { cellDates: true }), 'Conciliación');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(datos.filter((d) => d.Estado !== 'Conciliado'), { cellDates: true }), 'Pendientes');
  XLSX.writeFile(wb, `conciliacion_${new Date().toISOString().slice(0, 10)}.xlsx`);
  aviso('Excel descargado.');
});

$('btnReiniciar').addEventListener('click', () => {
  state.mayor = state.banco = null;
  state.resultado = [];
  for (const t of ['Mayor', 'Banco']) {
    $('file' + t).value = '';
    $('drop' + t).classList.remove('loaded', 'error');
    $('hint' + t).innerHTML = `${ico('upload')}Arrastrá o hacé clic · .xlsx .xls .csv`;
  }
  $('mapping').hidden = true;
  $('results').hidden = true;
  marcarPaso(1);
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

// Datos de ejemplo: un mes chico con casos típicos (acreditación con días de demora,
// débito/crédito separados, comisión bancaria sin registrar y un cheque no cobrado).
$('btnDemo').addEventListener('click', () => {
  state.mayor = armarTabla([
    ['Fecha', 'Concepto', 'Importe'],
    ['01/10/2026', 'Cobro factura A-0012 · Ferretería Sur', '185.400,00'],
    ['02/10/2026', 'Pago proveedor Insumos Litoral', '-92.750,50'],
    ['05/10/2026', 'Transferencia cliente Agro Federal', '310.000,00'],
    ['07/10/2026', 'Pago alquiler oficina', '-240.000,00'],
    ['10/10/2026', 'Cheque emitido N° 4471', '-58.300,00'],
    ['14/10/2026', 'Cobro factura A-0015 · Estudio Paz', '96.800,00'],
    ['20/10/2026', 'Pago servicios (luz)', '-31.245,90'],
  ]);
  state.banco = armarTabla([
    ['Banco de Ejemplo · Extracto de cuenta corriente'],
    ['Período: octubre 2026'],
    [],
    ['Fecha', 'Descripción', 'Débito', 'Crédito'],
    ['01/10/2026', 'TRANSF RECIBIDA FERRETERIA SUR', '', '185.400,00'],
    ['03/10/2026', 'TRANSF ENVIADA INSUMOS LITORAL', '92.750,50', ''],
    ['07/10/2026', 'TRANSF RECIBIDA AGRO FEDERAL', '', '310.000,00'],
    ['07/10/2026', 'DEBITO ALQUILER', '240.000,00', ''],
    ['15/10/2026', 'TRANSF RECIBIDA ESTUDIO PAZ', '', '96.800,00'],
    ['20/10/2026', 'PAGO SERVICIOS EDEER', '31.245,90', ''],
    ['31/10/2026', 'COMISION MANTENIMIENTO CUENTA', '8.500,00', ''],
    ['31/10/2026', 'IMP. LEY 25413 DEB/CRED', '5.890,20', ''],
  ]);
  for (const [t, n] of [['Mayor', 'libro_mayor_ejemplo.xlsx'], ['Banco', 'extracto_ejemplo.xlsx']]) {
    $('drop' + t).classList.remove('error');
    $('drop' + t).classList.add('loaded');
    $('hint' + t).innerHTML = `${ico('check')}${n} · ${state[t.toLowerCase()].rows.length} filas`;
  }
  renderMapeo();
  aviso('Cargamos datos de ejemplo. Revisá las columnas y tocá "Conciliar ahora".');
});

configurarCarga('mayor', 'fileMayor', 'dropMayor', 'hintMayor');
configurarCarga('banco', 'fileBanco', 'dropBanco', 'hintBanco');
