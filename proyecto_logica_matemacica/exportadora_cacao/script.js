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
        // Scroll row into view slightly if on mobile
        row.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
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
    return new Intl.NumberFormat('es-EC', {
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
    outputLibras.value = lb.toLocaleString('es-EC', { maximumFractionDigits: 2 });
  }

  function handleFinanceCalculation() {
    const fob = parseFloat(inputFob.value) || 0;
    const taxRate = parseFloat(taxPercentage.value) || 0;
    const taxes = fob * (taxRate / 100);
    outputTaxes.textContent = formatCurrency(taxes);
  }

  // --- INITIALIZE EVENT LISTENERS ---
  switchP.addEventListener('change', updateLogicSimulator);
  switchQ.addEventListener('change', updateLogicSimulator);

  inputKm.addEventListener('input', handleDistanceConversion);
  inputLitros.addEventListener('input', handleFuelConversion);
  inputKg.addEventListener('input', handleWeightConversion);
  inputFob.addEventListener('input', handleFinanceCalculation);
  taxPercentage.addEventListener('input', handleFinanceCalculation);

  // --- RUN INITIAL CALCULATIONS ---
  updateLogicSimulator();
  handleDistanceConversion();
  handleFuelConversion();
  handleWeightConversion();
  handleFinanceCalculation();
});
