document.addEventListener('DOMContentLoaded', () => {
  // --- LOGIC SIMULATOR ELEMENTS ---
  const switchP = document.getElementById('switch-p');
  const switchQ = document.getElementById('switch-q');
  
  const decisionCard = document.getElementById('decision-card');
  const decisionIcon = document.getElementById('decision-icon');
  const decisionTitle = document.getElementById('decision-title');
  const decisionFormula = document.getElementById('decision-formula');
  const decisionMessage = document.getElementById('decision-message');
  
  const rows = [
    document.getElementById('row-0'), // V, V -> F
    document.getElementById('row-1'), // V, F -> V (Approved)
    document.getElementById('row-2'), // F, V -> F
    document.getElementById('row-3')  // F, F -> F
  ];

  // --- UNIT CONVERSION ELEMENTS ---
  const inputKm = document.getElementById('input-km');
  const outputMi = document.getElementById('output-mi');
  
  const inputLitros = document.getElementById('input-litros');
  const outputGalones = document.getElementById('output-galones');
  
  const inputKg = document.getElementById('input-kg');
  const outputLibras = document.getElementById('output-libras');
  
  const inputFob = document.getElementById('input-fob');
  const taxPercentage = document.getElementById('tax-percentage');
  const outputTaxes = document.getElementById('output-taxes');

  // --- LOGIC FUNCTIONS ---
  let isInitialLoad = true;

  function updateLogicSimulator() {
    const p = switchP.checked; // true = V, false = F
    const q = switchQ.checked; // true = V, false = F
    const notQ = !q;
    const D = p && notQ;

    // Determine active row index in Truth Table
    // Row 0: p=V, q=V (Index 0)
    // Row 1: p=V, q=F (Index 1)
    // Row 2: p=F, q=V (Index 2)
    // Row 3: p=F, q=F (Index 3)
    let activeIndex = 0;
    if (p && q) activeIndex = 0;
    else if (p && !q) activeIndex = 1;
    else if (!p && q) activeIndex = 2;
    else if (!p && !q) activeIndex = 3;

    // Remove active class from all rows and add to active row
    rows.forEach((row, index) => {
      if (index === activeIndex) {
        row.classList.add('active-row');
        // Scroll row into view only on user interaction, never on page load
        if (!isInitialLoad) {
          row.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }
      } else {
        row.classList.remove('active-row');
      }
    });

    // Update Decision Card display
    const pStr = p ? 'V' : 'F';
    const qStr = q ? 'V' : 'F';
    const notQStr = notQ ? 'V' : 'F';
    const dStr = D ? 'V' : 'F';

    decisionFormula.textContent = `D = ${pStr} ∧ ¬${qStr} = ${dStr}`;

    if (D) {
      // Approved status
      decisionCard.className = 'decision-status-card approved';
      decisionIcon.textContent = '✅';
      decisionTitle.textContent = 'Ruta Aprobada';
      decisionMessage.textContent = 'Prioridad de tiempo confirmada y presupuesto viable. Se autoriza la ruta rápida directa.';
    } else {
      // Rejected status
      decisionCard.className = 'decision-status-card rejected';
      decisionIcon.textContent = '🚫';
      decisionTitle.textContent = 'Ruta Rechazada';
      
      // Detailed feedback text per condition
      if (p && q) {
        decisionMessage.textContent = 'El envío es urgente, pero el costo de la ruta rápida excede el presupuesto disponible.';
      } else if (!p && q) {
        decisionMessage.textContent = 'El envío no es urgente y los costos estimados superan el techo financiero asignado.';
      } else {
        decisionMessage.textContent = 'No amerita ruta directa por baja urgencia. Se utilizará la ruta marítima estándar.';
      }
    }
  }

  // --- CONVERSION FUNCTIONS ---
  function formatCurrency(val) {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD'
    }).format(val);
  }

  function handleDistanceConversion() {
    const km = parseFloat(inputKm.value) || 0;
    const mi = km * 0.621371;
    outputMi.value = mi.toFixed(2);
  }

  function handleFuelConversion() {
    const L = parseFloat(inputLitros.value) || 0;
    const gal = L * 0.264172;
    outputGalones.value = gal.toFixed(2);
  }

  function handleWeightConversion() {
    const kg = parseFloat(inputKg.value) || 0;
    const lb = kg * 2.20462;
    outputLibras.value = lb.toLocaleString('en-US', { maximumFractionDigits: 2 });
  }

  function handleFinanceCalculation() {
    const fob = parseFloat(inputFob.value) || 0;
    const taxRate = parseFloat(taxPercentage.value) || 0;
    const taxes = fob * (taxRate / 100);
    outputTaxes.textContent = formatCurrency(taxes);
  }

  // --- FASE 2: 2x2 SYSTEM SOLVER (transporte x, seguro y) ---
  const inputPresupuesto = document.getElementById('input-presupuesto');
  const inputFactor = document.getElementById('input-factor');
  const eq1 = document.getElementById('eq-1');
  const eq2 = document.getElementById('eq-2');
  const solutionSteps = document.getElementById('solution-steps');
  const solutionResult = document.getElementById('solution-result');
  const verification = document.getElementById('verification');
  const systemGraph = document.getElementById('system-graph');
  const graphCaption = document.getElementById('graph-caption');

  const inputComision = document.getElementById('input-comision');
  const outputCostoTotal = document.getElementById('output-costo-total');
  const costoDetalle = document.getElementById('costo-detalle');

  function fmt(n, dec = 2) {
    return n.toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec });
  }

  function drawSystemGraph(T, k, transporte, seguro) {
    // Plot: horizontal axis = S (seguro), vertical axis = T (transporte)
    // Line 1: T = presupuesto - S   |   Line 2: T = k*S
    const W = 520, H = 340, pad = 45;
    const yMax = T;          // horizontal domain 0..presupuesto
    const xMax = T;          // vertical domain 0..presupuesto
    const sx = val => pad + (val / yMax) * (W - 2 * pad);
    const sy = val => (H - pad) - (val / xMax) * (H - 2 * pad);

    let svg = '';
    // Axes
    svg += `<line x1="${pad}" y1="${H - pad}" x2="${W - pad}" y2="${H - pad}" stroke="#9C7F55" stroke-width="1.5"/>`;
    svg += `<line x1="${pad}" y1="${H - pad}" x2="${pad}" y2="${pad - 10}" stroke="#9C7F55" stroke-width="1.5"/>`;
    // Axis labels
    svg += `<text x="${W - pad}" y="${H - pad + 30}" text-anchor="end" font-size="12" fill="#6B5744">S (seguro $)</text>`;
    svg += `<text x="${pad - 35}" y="${pad - 18}" font-size="12" fill="#6B5744">T (transporte $)</text>`;
    // Ticks: 0, T/2, T on each axis
    [0, T / 2, T].forEach(v => {
      svg += `<text x="${sx(v)}" y="${H - pad + 16}" text-anchor="middle" font-size="10" fill="#8A7358">${fmt(v, 0)}</text>`;
      svg += `<text x="${pad - 6}" y="${sy(v) + 3}" text-anchor="end" font-size="10" fill="#8A7358">${fmt(v, 0)}</text>`;
    });
    // Line 1: x = T - y (from (0, T) to (T, 0))
    svg += `<line x1="${sx(0)}" y1="${sy(T)}" x2="${sx(T)}" y2="${sy(0)}" stroke="#2E1B0B" stroke-width="2.5"/>`;
    svg += `<text x="${sx(T * 0.62)}" y="${sy(T * 0.38) - 10}" font-size="11" fill="#2E1B0B" font-weight="bold">T + S = ${fmt(T, 0)}</text>`;
    // Line 2: T = k*S (from (0,0) until it leaves the box)
    const yEnd = Math.min(T, T / k); // where k*S reaches the top of the box
    svg += `<line x1="${sx(0)}" y1="${sy(0)}" x2="${sx(yEnd)}" y2="${sy(k * yEnd)}" stroke="#B87A0F" stroke-width="2.5"/>`;
    svg += `<text x="${sx(yEnd * 0.55) + 8}" y="${sy(k * yEnd * 0.55)}" font-size="11" fill="#B87A0F" font-weight="bold">T = ${k}S</text>`;
    // Intersection point
    svg += `<circle cx="${sx(seguro)}" cy="${sy(transporte)}" r="6" fill="#8C3A22" stroke="#ffffff" stroke-width="2"/>`;
    svg += `<text x="${sx(seguro) + 10}" y="${sy(transporte) - 10}" font-size="11.5" fill="#8C3A22" font-weight="bold">(${fmt(seguro, 0)}, ${fmt(transporte, 0)})</text>`;

    systemGraph.innerHTML = svg;
  }

  function solveSystem() {
    const T = parseFloat(inputPresupuesto.value) || 0;
    const k = parseFloat(inputFactor.value) || 1;
    if (T <= 0 || k <= 0) return;

    // T + S = presupuesto ; T = k*S  ->  S = presupuesto/(k+1), T = k*S
    const S = T / (k + 1);
    const Tr = k * S;

    eq1.textContent = `T + S = ${fmt(T, 0)}`;
    eq2.textContent = `T = ${k}S`;

    solutionSteps.innerHTML =
      `<li>Sustituir T = ${k}S en la 1ª ecuación: ${k}S + S = ${fmt(T, 0)}</li>` +
      `<li>${k + 1}S = ${fmt(T, 0)} → S = ${fmt(S)}</li>` +
      `<li>T = ${k}(${fmt(S)}) = ${fmt(Tr)}</li>`;

    solutionResult.innerHTML = `T = $${fmt(Tr)} (transporte) &nbsp;•&nbsp; S = $${fmt(S)} (seguro)`;
    verification.innerHTML = `✔ Comprobación: ${fmt(Tr)} + ${fmt(S)} = ${fmt(Tr + S)} ✓ &nbsp;|&nbsp; ${fmt(Tr)} = ${k} × ${fmt(S)} ✓`;
    graphCaption.textContent = `La solución del sistema es el punto donde se cruzan ambas rectas: (S, T) = (${fmt(S, 0)}, ${fmt(Tr, 0)}).`;

    drawSystemGraph(T, k, Tr, S);
    calcCostoTotal();
  }

  function calcCostoTotal() {
    const T = parseFloat(inputPresupuesto.value) || 0;
    const k = parseFloat(inputFactor.value) || 1;
    const a = parseFloat(inputComision.value) || 0;
    const fob = parseFloat(inputFob.value) || 0;
    const S = T / (k + 1);
    const Tr = k * S;
    const comision = fob * (a / 100);
    const C = Tr + S + comision;
    outputCostoTotal.textContent = formatCurrency(C);
    costoDetalle.textContent = `C = ${fmt(Tr)} + ${fmt(S)} + (${a}% de ${fmt(fob)}) = ${fmt(Tr)} + ${fmt(S)} + ${fmt(comision)}`;
  }

  // --- AVANCE PRODUCTO 2/3: UTILITY FUNCTION U(x) = -2500x^2 + 23550x - 20000 ---
  const A_U = -2500, B_U = 23550, C_U = -20000;
  const inputContenedores = document.getElementById('input-contenedores');
  const xValue = document.getElementById('x-value');
  const outputUtilidad = document.getElementById('output-utilidad');
  const utilidadDetalle = document.getElementById('utilidad-detalle');
  const utilityGraph = document.getElementById('utility-graph');

  const U = x => A_U * x * x + B_U * x + C_U;

  function drawUtilityGraph(selected) {
    const W = 560, H = 380, padL = 66, padR = 22, padT = 28, padB = 46;
    const xMin = 0, xMax = 9.4;
    const yMin = -25000, yMax = 42000;
    const sx = v => padL + ((v - xMin) / (xMax - xMin)) * (W - padL - padR);
    const sy = v => (H - padB) - ((v - yMin) / (yMax - yMin)) * (H - padT - padB);

    const xv = -B_U / (2 * A_U);              // 4.71
    const disc = B_U * B_U - 4 * A_U * C_U;
    const r1 = (-B_U + Math.sqrt(disc)) / (2 * A_U);
    const r2 = (-B_U - Math.sqrt(disc)) / (2 * A_U);

    let svg = '';

    // Positive-profit shaded region
    let area = `${sx(r1)},${sy(0)} `;
    for (let x = r1; x <= r2; x += 0.05) area += `${sx(x)},${sy(U(x))} `;
    area += `${sx(r2)},${sy(0)}`;
    svg += `<polygon points="${area}" fill="#B87A0F" opacity="0.10"/>`;

    // Axes
    svg += `<line x1="${padL}" y1="${sy(0)}" x2="${W - padR}" y2="${sy(0)}" stroke="#9C7F55" stroke-width="1.4"/>`;
    svg += `<line x1="${sx(0)}" y1="${padT}" x2="${sx(0)}" y2="${H - padB}" stroke="#9C7F55" stroke-width="1.4"/>`;
    svg += `<text x="${W - padR}" y="${H - padB + 34}" text-anchor="end" font-size="11" fill="#6B5744">x (contenedores/mes)</text>`;
    svg += `<text x="${padL - 58}" y="${padT - 10}" font-size="11" fill="#6B5744">U(x) en USD</text>`;

    // Ticks
    for (let x = 0; x <= 9; x++) {
      svg += `<line x1="${sx(x)}" y1="${sy(0)}" x2="${sx(x)}" y2="${sy(0) + 5}" stroke="#9C7F55" stroke-width="1"/>`;
      svg += `<text x="${sx(x)}" y="${sy(0) + 18}" text-anchor="middle" font-size="10" fill="#8A7358">${x}</text>`;
    }
    [-20000, 0, 20000, 40000].forEach(v => {
      svg += `<line x1="${padL - 5}" y1="${sy(v)}" x2="${padL}" y2="${sy(v)}" stroke="#9C7F55" stroke-width="1"/>`;
      svg += `<text x="${padL - 9}" y="${sy(v) + 4}" text-anchor="end" font-size="9.5" fill="#8A7358">${v === 0 ? '0' : (v / 1000) + 'k'}</text>`;
    });

    // Parabola
    let path = '';
    for (let x = xMin; x <= xMax; x += 0.04) {
      const y = U(x);
      if (y < yMin - 5000) continue;
      path += (path ? ' L ' : 'M ') + sx(x) + ',' + sy(y);
    }
    svg += `<path d="${path}" fill="none" stroke="#2E1B0B" stroke-width="2.6"/>`;

    // Roots
    [r1, r2].forEach(r => {
      svg += `<circle cx="${sx(r)}" cy="${sy(0)}" r="5.5" fill="#8C3A22" stroke="#fff" stroke-width="1.5"/>`;
    });
    svg += `<text x="${sx(r1)}" y="${sy(0) + 34}" text-anchor="middle" font-size="10" fill="#8C3A22" font-weight="bold">x ≈ ${r1.toFixed(2)}</text>`;
    svg += `<text x="${sx(r2)}" y="${sy(0) + 34}" text-anchor="middle" font-size="10" fill="#8C3A22" font-weight="bold">x ≈ ${r2.toFixed(2)}</text>`;

    // Integer domain points
    for (let x = 1; x <= 8; x++) {
      svg += `<rect x="${sx(x) - 4}" y="${sy(U(x)) - 4}" width="8" height="8" fill="#5E2A46"/>`;
    }

    // Vertex
    svg += `<line x1="${sx(xv)}" y1="${sy(0)}" x2="${sx(xv)}" y2="${sy(U(xv))}" stroke="#B87A0F" stroke-dasharray="3,3" stroke-width="1"/>`;
    svg += `<circle cx="${sx(xv)}" cy="${sy(U(xv))}" r="6" fill="#B87A0F" stroke="#fff" stroke-width="2"/>`;
    svg += `<text x="${sx(xv)}" y="${sy(U(xv)) - 14}" text-anchor="middle" font-size="10.5" fill="#B87A0F" font-weight="bold">V(4.71 ; $35,460)</text>`;

    // Monotonicity labels
    svg += `<text x="${sx(2.4)}" y="${sy(39000)}" text-anchor="middle" font-size="10.5" fill="#6B5744">crece</text>`;
    svg += `<text x="${sx(7.1)}" y="${sy(39000)}" text-anchor="middle" font-size="10.5" fill="#6B5744">decrece</text>`;

    // Selected point
    if (selected >= 0 && selected <= 9) {
      svg += `<rect x="${sx(selected) - 6.5}" y="${sy(U(selected)) - 6.5}" width="13" height="13" fill="#2F5D3F" stroke="#fff" stroke-width="1.6"/>`;
    }

    utilityGraph.innerHTML = svg;
  }

  function updateUtility() {
    const x = parseInt(inputContenedores.value, 10);
    const u = U(x);
    xValue.textContent = x;
    outputUtilidad.textContent = formatCurrency(u);
    const estado = u > 0 ? 'operación con ganancia' : 'operación en pérdida';
    utilidadDetalle.textContent =
      `U(${x}) = −2,500(${x})² + 23,550(${x}) − 20,000 = ${fmt(u)}  →  ${estado}.`;
    drawUtilityGraph(x);
  }

  // ===================== PHASE TAB NAVIGATION =====================
  const navButtons = document.querySelectorAll('.nav-btn');
  const tabPanels = document.querySelectorAll('.tab-panel');

  navButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.panel;
      navButtons.forEach(b => b.classList.toggle('is-active', b === btn));
      tabPanels.forEach(p => p.classList.toggle('is-active', p.id === target));
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  });

  // ===================== EXPORT CALCULATOR =====================
  // Distances are approximate sea routes from the Port of Guayaquil via the Panama Canal.
  // km = ruta marítima aprox. vía Canal de Panamá; kmAire = ruta aérea aprox.
  // arancelPct: tarifa efectiva para cacao en grano (partida 1801).
  const DESTINOS = {
    nl: { cod: 'RTM', nombre: 'Rotterdam, Países Bajos', km: 10700, kmAire: 9700, arancelPct: 0, arancelBase: 'Acuerdo Multipartes Ecuador–UE', regional: false },
    de: { cod: 'HAM', nombre: 'Hamburgo, Alemania', km: 11200, kmAire: 9900, arancelPct: 0, arancelBase: 'Acuerdo Multipartes Ecuador–UE', regional: false },
    be: { cod: 'ANR', nombre: 'Amberes, Bélgica', km: 10750, kmAire: 9650, arancelPct: 0, arancelBase: 'Acuerdo Multipartes Ecuador–UE', regional: false },
    es: { cod: 'BCN', nombre: 'Barcelona, España', km: 10300, kmAire: 9100, arancelPct: 0, arancelBase: 'Acuerdo Multipartes Ecuador–UE', regional: false },
    us: { cod: 'NYC', nombre: 'Nueva York, Estados Unidos', km: 5850, kmAire: 4300, arancelPct: 0, arancelBase: 'Acuerdo Comercial Recíproco (mar-2026), aplica al grano', regional: false },
    mx: { cod: 'ZLO', nombre: 'Manzanillo, México', km: 3900, kmAire: 3100, arancelPct: 20, arancelBase: 'tarifa NMF; verificar preferencia ACE-29', regional: false },
    my: { cod: 'PKG', nombre: 'Port Klang, Malasia', km: 17800, kmAire: 19000, arancelPct: 0, arancelBase: 'grano libre de arancel (país procesador)', regional: false },
    co: { cod: 'BUN', nombre: 'Buenaventura, Colombia', km: 640, kmAire: 600, arancelPct: 0, arancelBase: 'Comunidad Andina', regional: true },
    pe: { cod: 'CLL', nombre: 'Callao, Perú', km: 1450, kmAire: 1350, arancelPct: 0, arancelBase: 'Comunidad Andina', regional: true }
  };

  // Provincias cacaoteras: distancia por carretera al Puerto de Guayaquil
  // y al aeropuerto de carga más cercano (GYE Olmedo o UIO Mariscal Sucre).
  const ORIGENES = {
    gys: { prov: 'Guayas', ciudad: 'Guayaquil', kmPuerto: 0, aero: 'GYE', aeroNombre: 'Aeropuerto José Joaquín de Olmedo, Guayaquil', kmAero: 15 },
    man: { prov: 'Manabí', ciudad: 'Portoviejo', kmPuerto: 190, aero: 'GYE', aeroNombre: 'Aeropuerto José Joaquín de Olmedo, Guayaquil', kmAero: 190 },
    esm: { prov: 'Esmeraldas', ciudad: 'Esmeraldas', kmPuerto: 470, aero: 'UIO', aeroNombre: 'Aeropuerto Mariscal Sucre, Quito', kmAero: 320 },
    lrs: { prov: 'Los Ríos', ciudad: 'Babahoyo', kmPuerto: 60, aero: 'GYE', aeroNombre: 'Aeropuerto José Joaquín de Olmedo, Guayaquil', kmAero: 70 },
    sdt: { prov: 'Santo Domingo', ciudad: 'Santo Domingo', kmPuerto: 290, aero: 'UIO', aeroNombre: 'Aeropuerto Mariscal Sucre, Quito', kmAero: 150 },
    oro: { prov: 'El Oro', ciudad: 'Machala', kmPuerto: 180, aero: 'GYE', aeroNombre: 'Aeropuerto José Joaquín de Olmedo, Guayaquil', kmAero: 190 },
    suc: { prov: 'Sucumbíos', ciudad: 'Lago Agrio', kmPuerto: 600, aero: 'UIO', aeroNombre: 'Aeropuerto Mariscal Sucre, Quito', kmAero: 250 },
    orl: { prov: 'Orellana', ciudad: 'El Coca', kmPuerto: 560, aero: 'UIO', aeroNombre: 'Aeropuerto Mariscal Sucre, Quito', kmAero: 300 }
  };

  // Conversion factor of each unit to kilograms.
  const UNIDADES = {
    qq: { kg: 45.36, nombre: 'quintales' },
    saco60: { kg: 60, nombre: 'sacos de 60 kg' },
    kg: { kg: 1, nombre: 'kilogramos' },
    t: { kg: 1000, nombre: 'toneladas' },
    lb: { kg: 0.45359237, nombre: 'libras' },
    tonus: { kg: 907.18474, nombre: 'toneladas cortas' }
  };

  // Rates calibrated against the project data: maritime reproduces T = $2,800 for
  // 25 t to Rotterdam, and land reproduces the Esmeraldas-Guayaquil leg.
  const MODOS = {
    mar: { nombre: 'marítimo', tarifa: 0.0000105, kmDia: 650, diasFijos: 5, soloRegional: false },
    aire: { nombre: 'aéreo', tarifa: 0.00033, kmDia: 8000, diasFijos: 1, soloRegional: false },
    tierra: { nombre: 'terrestre', tarifa: 0.00006, kmDia: 500, diasFijos: 1, soloRegional: true }
  };

  // Incoterms 2020: FOB solo es válido para transporte marítimo y fluvial.
  // Para aéreo y terrestre el término equivalente es FCA (Free Carrier).
  const INCOTERM = {
    mar: { sigla: 'FOB', nombre: 'Free On Board', punto: 'puesta a bordo del buque' },
    aire: { sigla: 'FCA', nombre: 'Free Carrier', punto: 'entregada a la aerolínea de carga' },
    tierra: { sigla: 'FCA', nombre: 'Free Carrier', punto: 'entregada al transportista' }
  };

  const PRECIO_TON = 5900;   // USD per tonne (Anecacao, july 2026)
  const CAP_CONT = 25000;    // kg per 40-foot container
  const TASA_SEGURO = 0.0047;
  const TASA_COMISION = 0.02;

  let modoActivo = 'mar';
  let ultimaCotizacion = {};

  const calcOrigen = document.getElementById('calc-origen');
  const calcCantidad = document.getElementById('calc-cantidad');
  const calcUnidad = document.getElementById('calc-unidad');
  const calcDestino = document.getElementById('calc-destino');
  const calcUrgente = document.getElementById('calc-urgente');
  const calcPresupuesto = document.getElementById('calc-presupuesto');
  const modoBtns = document.querySelectorAll('.modo-btn');

  function calcular() {
    const dest = DESTINOS[calcDestino.value];
    const uni = UNIDADES[calcUnidad.value];
    const cantidad = parseFloat(calcCantidad.value) || 0;
    const pesoKg = cantidad * uni.kg;

    // Land transport only reaches neighbouring countries.
    modoBtns.forEach(b => {
      const m = MODOS[b.dataset.modo];
      const bloqueado = m.soloRegional && !dest.regional;
      b.classList.toggle('is-disabled', bloqueado);
      b.disabled = bloqueado;
    });
    if (MODOS[modoActivo].soloRegional && !dest.regional) {
      modoActivo = 'mar';
      modoBtns.forEach(b => b.classList.toggle('is-active', b.dataset.modo === 'mar'));
    }

    const modo = MODOS[modoActivo];
    const orig = ORIGENES[calcOrigen.value];
    const fob = (pesoKg / 1000) * PRECIO_TON;

    // El viaje son dos tramos: acarreo interno hasta el nodo de salida
    // (puerto o aeropuerto) y luego el tramo internacional.
    let kmInterno, kmIntl, codSalida, nodoSalida;
    if (modoActivo === 'tierra') {
      // Todo el trayecto es por carretera: no hay nodo intermedio.
      kmInterno = 0;
      kmIntl = orig.kmPuerto + dest.km;
      codSalida = orig.ciudad.slice(0, 3).toUpperCase();
      nodoSalida = `${orig.ciudad}, Ecuador`;
    } else if (modoActivo === 'aire') {
      kmInterno = orig.kmAero;
      kmIntl = dest.kmAire;
      codSalida = orig.aero;
      nodoSalida = orig.aeroNombre;
    } else {
      kmInterno = orig.kmPuerto;
      kmIntl = dest.km;
      codSalida = 'GYE';
      nodoSalida = 'Puerto de Guayaquil, Ecuador';
    }

    const costoInterno = pesoKg * kmInterno * MODOS.tierra.tarifa;
    const transporte = pesoKg * kmIntl * modo.tarifa;
    const seguro = fob * TASA_SEGURO;
    const comisiones = fob * TASA_COMISION;
    const total = costoInterno + transporte + seguro + comisiones;

    const diasInterno = kmInterno / MODOS.tierra.kmDia;
    const dias = diasInterno + kmIntl / modo.kmDia + modo.diasFijos;
    const contenedores = Math.ceil(pesoKg / CAP_CONT);
    const arancel = fob * (dest.arancelPct / 100);

    document.getElementById('hint-peso').textContent =
      `Equivale a ${fmt(pesoKg / 45.36, 1)} quintales · ${fmt(pesoKg, 0)} kg · ` +
      `${fmt(pesoKg * 2.20462, 0)} lb · ${fmt(pesoKg / 1000, 2)} toneladas`;

    document.getElementById('hint-origen').textContent = kmInterno > 0
      ? `Desde ${orig.ciudad} hay ${fmt(kmInterno, 0)} km por carretera hasta ${nodoSalida}.`
      : `Sales directamente desde ${nodoSalida}: no hay acarreo interno.`;

    document.getElementById('hint-distancia').textContent =
      `Tramo internacional desde ${nodoSalida} hasta ${dest.nombre}: ` +
      `${fmt(kmIntl, 0)} km (${fmt(kmIntl * 0.621371, 0)} millas)`;

    document.getElementById('hint-modo').textContent = dest.regional
      ? 'Este destino es un país vecino, así que el transporte terrestre sí es posible.'
      : 'El transporte terrestre está deshabilitado porque no existe ruta por carretera hasta este destino.';

    // Marca de embarque (elemento firma)
    document.getElementById('cm-origen').textContent = codSalida;
    document.getElementById('cm-origen-full').textContent = nodoSalida;
    document.getElementById('cm-destino').textContent = dest.cod;
    document.getElementById('cm-destino-full').textContent = dest.nombre;
    document.getElementById('cm-peso').textContent = `${fmt(pesoKg, 0)} kg`;
    document.getElementById('cm-modo').textContent = modo.nombre;

    document.getElementById('res-total').textContent = formatCurrency(total);
    document.getElementById('res-total-2').textContent = formatCurrency(total);
    document.getElementById('res-porcentaje').textContent =
      fob > 0 ? `${(total / fob * 100).toFixed(1)}% del valor de la carga` : '';
    document.getElementById('res-dias').textContent = `${dias.toFixed(1)} días`;
    document.getElementById('res-contenedores').textContent =
      `${contenedores} × 25 t`;
    document.getElementById('res-fob').textContent = formatCurrency(fob);

    // El Incoterm depende de la vía: FOB es exclusivo del transporte marítimo.
    const inco = INCOTERM[modoActivo];
    document.getElementById('bd-valor-lbl').firstChild.textContent =
      `Valor ${inco.sigla} de la carga `;
    document.getElementById('bd-incoterm-det').textContent =
      `${inco.nombre} · ${inco.punto}`;
    document.getElementById('bd-seguro-det').textContent =
      `0.47% del valor ${inco.sigla}`;
    document.getElementById('bd-comision-det').textContent =
      `2% del valor ${inco.sigla}`;

    document.getElementById('res-interno').textContent = formatCurrency(costoInterno);
    document.getElementById('res-transporte').textContent = formatCurrency(transporte);
    document.getElementById('res-seguro').textContent = formatCurrency(seguro);
    document.getElementById('res-comisiones').textContent = formatCurrency(comisiones);
    document.getElementById('bd-modo-nombre').textContent = modo.nombre;

    // Acarreo interno: se oculta cuando no aplica
    const filaInterno = document.getElementById('bd-interno-row');
    filaInterno.style.display = kmInterno > 0 ? '' : 'none';
    document.getElementById('bd-interno-det').textContent =
      `${orig.ciudad} → ${codSalida}, ${fmt(kmInterno, 0)} km`;

    // Arancel: tarifa efectiva y monto, no un "consultar"
    document.getElementById('res-arancel').textContent = dest.arancelPct > 0
      ? `${dest.arancelPct}% · ${formatCurrency(arancel)}`
      : '0% · exento';
    document.getElementById('bd-arancel-base').textContent = dest.arancelBase;

    // Datos que viajan en el mensaje de WhatsApp
    ultimaCotizacion = {
      origen: `${orig.prov} (${orig.ciudad})`,
      destino: dest.nombre,
      carga: `${fmt(pesoKg / 45.36, 1)} quintales / ${fmt(pesoKg, 0)} kg`,
      via: `${modo.nombre} desde ${nodoSalida}`,
      total: formatCurrency(total),
      dias: `${dias.toFixed(1)} días`
    };
    refrescarContacto();

    // Decision rule D = p AND NOT q, reusing the logic from Fase 1.
    const presupuesto = parseFloat(calcPresupuesto.value) || 0;
    const p = calcUrgente.checked;
    const q = total > presupuesto;
    const D = p && !q;

    const card = document.getElementById('calc-decision');
    const icon = document.getElementById('calc-dec-icon');
    const title = document.getElementById('calc-dec-title');
    const msg = document.getElementById('calc-dec-msg');
    document.getElementById('calc-dec-formula').textContent =
      `D = ${p ? 'V' : 'F'} ∧ ¬${q ? 'V' : 'F'} = ${D ? 'V' : 'F'}`;

    if (D) {
      card.className = 'decision-status-card approved';
      icon.textContent = '✅';
      title.textContent = 'Ruta rápida aprobada';
      msg.textContent = `El envío es urgente y su costo (${formatCurrency(total)}) cabe en el presupuesto de ${formatCurrency(presupuesto)}.`;
    } else {
      card.className = 'decision-status-card rejected';
      icon.textContent = '🚫';
      title.textContent = 'Ruta estándar';
      if (p && q) {
        msg.textContent = `Es urgente, pero el costo (${formatCurrency(total)}) supera el presupuesto de ${formatCurrency(presupuesto)}.`;
      } else if (!p && q) {
        msg.textContent = 'No es urgente y además el costo supera el presupuesto disponible.';
      } else {
        msg.textContent = 'El envío no es urgente, así que conviene la ruta marítima estándar.';
      }
    }
  }

  // ===================== CONTACTO POR WHATSAPP =====================
  // >>> CAMBIA ESTE NÚMERO POR EL TUYO <<<
  // Formato: código de país + número, sin "+", sin espacios ni guiones.
  // Ecuador es 593 y se quita el 0 inicial del celular.
  // Ejemplo: el celular 0987654321 se escribe 593987654321
  const WA_NUMERO = '593999999999';
  const CORREO = 'cacao@ejemplo.ec';

  const waFab = document.getElementById('wa-fab');
  const cdWa = document.getElementById('cd-wa');
  const formContacto = document.getElementById('form-contacto');
  const ctResumenTxt = document.getElementById('ct-resumen-txt');

  function textoCotizacion() {
    const c = ultimaCotizacion;
    if (!c.destino) return '';
    return `Origen: ${c.origen}\n` +
      `Destino: ${c.destino}\n` +
      `Carga: ${c.carga}\n` +
      `Vía: ${c.via}\n` +
      `Costo logístico estimado: ${c.total}\n` +
      `Tránsito estimado: ${c.dias}`;
  }

  function urlWhatsApp(extra) {
    const cuerpo = 'Hola, hice una cotización en Pepa de Oro y quiero un precio en firme.\n\n' +
      textoCotizacion() + (extra ? `\n\n${extra}` : '');
    return `https://wa.me/${WA_NUMERO}?text=${encodeURIComponent(cuerpo)}`;
  }

  function refrescarContacto() {
    ctResumenTxt.textContent = textoCotizacion().replace(/\n/g, ' · ');
    waFab.href = urlWhatsApp('');
    cdWa.href = urlWhatsApp('');
    cdWa.textContent = '+' + WA_NUMERO.replace(/^(\d{3})(\d{2})(\d{3})(\d{4})$/, '$1 $2 $3 $4');
    document.getElementById('cd-mail').textContent = CORREO;
    document.getElementById('cd-mail').href =
      `mailto:${CORREO}?subject=${encodeURIComponent('Cotización de exportación de cacao')}` +
      `&body=${encodeURIComponent(textoCotizacion())}`;
  }

  formContacto.addEventListener('submit', e => {
    e.preventDefault();
    const nombre = document.getElementById('ct-nombre').value.trim();
    const contacto = document.getElementById('ct-contacto').value.trim();
    const mensaje = document.getElementById('ct-mensaje').value.trim();
    let extra = '';
    if (nombre) extra += `Mi nombre: ${nombre}\n`;
    if (contacto) extra += `Mi contacto: ${contacto}\n`;
    if (mensaje) extra += `\n${mensaje}`;
    window.open(urlWhatsApp(extra.trim()), '_blank', 'noopener');
  });

  modoBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      if (btn.disabled) return;
      modoActivo = btn.dataset.modo;
      modoBtns.forEach(b => b.classList.toggle('is-active', b === btn));
      calcular();
    });
  });

  [calcOrigen, calcCantidad, calcUnidad, calcDestino, calcUrgente, calcPresupuesto]
    .forEach(el => el.addEventListener('input', calcular));
  calcOrigen.addEventListener('change', calcular);
  calcDestino.addEventListener('change', calcular);
  calcUnidad.addEventListener('change', calcular);

  // --- INITIALIZE EVENT LISTENERS ---
  switchP.addEventListener('change', updateLogicSimulator);
  switchQ.addEventListener('change', updateLogicSimulator);

  inputKm.addEventListener('input', handleDistanceConversion);
  inputLitros.addEventListener('input', handleFuelConversion);
  inputKg.addEventListener('input', handleWeightConversion);
  inputFob.addEventListener('input', () => { handleFinanceCalculation(); calcCostoTotal(); });
  taxPercentage.addEventListener('input', handleFinanceCalculation);

  inputPresupuesto.addEventListener('input', solveSystem);
  inputFactor.addEventListener('input', solveSystem);
  inputComision.addEventListener('input', calcCostoTotal);
  inputContenedores.addEventListener('input', updateUtility);

  // --- RUN INITIAL CALCULATIONS ---
  updateLogicSimulator();
  isInitialLoad = false;
  handleDistanceConversion();
  handleFuelConversion();
  handleWeightConversion();
  handleFinanceCalculation();
  solveSystem();
  updateUtility();
  calcular();
});
