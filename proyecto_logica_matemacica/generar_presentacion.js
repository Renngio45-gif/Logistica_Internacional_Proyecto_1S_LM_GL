// Presentación "Pepa de Oro" — Proyecto Integrador, PUCE Sede Esmeraldas
// 4 diapositivas: portada + lógica + álgebra + herramienta y conclusión.
// La portada es independiente: se puede borrar si el docente exige solo 3.
const pptxgen = require('pptxgenjs');
const path = require('path');

const TINTA = '241505';   // grano fermentado — fondo dominante
const ORO = 'B87A0F';     // mazorca madura — acento
const ORO_VIVO = 'D89A22';// mazorca al sol — realce sobre oscuro
const PAPEL = 'E8E6DC';   // papel de manifiesto
const TENDAL = 'F5F3EB';  // superficie clara
const YUTE = '9C7F55';    // fibra del saco
const VERDE = '2F5D3F';   // hoja de cacao
const GRIS = '6B5744';    // texto atenuado sobre claro

const F_TIT = 'Arial';
const F_TXT = 'Calibri';
const F_MONO = 'Courier New';

const AUTOR = 'Geovanny Farías Estupiñán';
const IMG = path.join(__dirname, 'graficas_clase9', 'parte_D.png');

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';           // 13.33 x 7.5 pulgadas
pres.author = AUTOR;
pres.title = 'Pepa de Oro — Proyecto Integrador';

function rotulo(slide, texto, x, y, color, ancho) {
  slide.addText(texto, {
    x, y, w: ancho || 6, h: 0.25,
    fontFace: F_MONO, fontSize: 10, bold: true,
    color: color || ORO, charSpacing: 2, margin: 0
  });
}

/* ================================ 1 · PORTADA ================================ */
const s0 = pres.addSlide();
s0.background = { color: TINTA };

rotulo(s0, 'PONTIFICIA UNIVERSIDAD CATÓLICA DEL ECUADOR · SEDE ESMERALDAS', 0.9, 0.75, YUTE, 11);
rotulo(s0, 'PUCETEC · TECNOLOGÍA EN DESARROLLO DE SOFTWARE', 0.9, 1.02, YUTE, 11);

s0.addShape(pres.ShapeType.line, {
  x: 0.9, y: 1.5, w: 4.2, h: 0, line: { color: ORO, width: 2 }
});

s0.addText('PROYECTO INTEGRADOR', {
  x: 0.9, y: 1.75, w: 11.5, h: 0.35,
  fontFace: F_MONO, fontSize: 13, bold: true, color: ORO, charSpacing: 3, margin: 0
});

s0.addText('TEMA 7', {
  x: 0.9, y: 2.2, w: 11.5, h: 0.45,
  fontFace: F_TIT, fontSize: 24, bold: true, color: PAPEL, charSpacing: 1, margin: 0
});

s0.addText('Presupuesto y logística de\nun envío internacional', {
  x: 0.9, y: 2.7, w: 8.6, h: 1.5,
  fontFace: F_TIT, fontSize: 40, bold: true, color: ORO_VIVO, lineSpacing: 44, margin: 0
});

s0.addText('Caso aplicado: exportación de cacao fino de aroma desde Esmeraldas', {
  x: 0.9, y: 4.3, w: 8.6, h: 0.4,
  fontFace: F_TXT, fontSize: 17, color: PAPEL, italic: true, margin: 0
});

// Sello de la herramienta, a la derecha
s0.addShape(pres.ShapeType.rect, {
  x: 9.9, y: 2.7, w: 2.55, h: 1.9,
  fill: { color: TINTA }, line: { color: ORO, width: 2 }
});
s0.addText('LA HERRAMIENTA', {
  x: 10.1, y: 2.9, w: 2.15, h: 0.25,
  fontFace: F_MONO, fontSize: 8, bold: true, color: ORO, charSpacing: 1.5, margin: 0
});
s0.addText('PEPA\nDE ORO', {
  x: 10.1, y: 3.18, w: 2.15, h: 0.95,
  fontFace: F_TIT, fontSize: 26, bold: true, color: ORO_VIVO, lineSpacing: 28, margin: 0
});
s0.addText('aplicación web', {
  x: 10.1, y: 4.15, w: 2.15, h: 0.3,
  fontFace: F_MONO, fontSize: 9, color: PAPEL, margin: 0
});

// Datos del estudiante
s0.addShape(pres.ShapeType.line, {
  x: 0.9, y: 5.35, w: 11.55, h: 0, line: { color: '4A3A22', width: 1 }
});

const meta = [
  ['ESTUDIANTE', AUTOR],
  ['ASIGNATURA', 'Habilidades Lógico-Matemáticas'],
  ['DOCENTE', 'Msc. Adrián Vargas'],
  ['PERÍODO', 'Primer Nivel · 2026-1']
];
meta.forEach((m, i) => {
  const x = 0.9 + i * 2.95;
  s0.addText(m[0], {
    x, y: 5.6, w: 2.8, h: 0.25,
    fontFace: F_MONO, fontSize: 8, bold: true, color: ORO, charSpacing: 1.5, margin: 0
  });
  s0.addText(m[1], {
    x, y: 5.87, w: 2.8, h: 0.55,
    fontFace: F_TXT, fontSize: 13, color: PAPEL, lineSpacing: 16, margin: 0
  });
});

s0.addText('Esmeraldas, Ecuador · 2026', {
  x: 0.9, y: 6.75, w: 11.55, h: 0.3,
  fontFace: F_MONO, fontSize: 9, color: YUTE, margin: 0
});

s0.addNotes(
  'Buenos días. Soy Geovanny Farías. Mi proyecto es el tema 7, presupuesto y logística de un envío internacional, ' +
  'y lo apliqué a un caso de mi provincia: la exportación de cacao fino de aroma desde Esmeraldas. ' +
  'Construí una herramienta web que se llama Pepa de Oro, que es como se le dice al cacao ecuatoriano. ' +
  'Voy a explicar primero la lógica, después las ecuaciones, y al final les muestro la herramienta funcionando.'
);

/* ============================ 2 · PROBLEMA Y LÓGICA ============================ */
const s1 = pres.addSlide();
s1.background = { color: TINTA };

rotulo(s1, 'PARTE 1 · LÓGICA DEL SISTEMA', 0.6, 0.45);

s1.addText('¿CUÁNDO CONVIENE LA RUTA RÁPIDA?', {
  x: 0.6, y: 0.75, w: 12.1, h: 0.6,
  fontFace: F_TIT, fontSize: 34, bold: true, color: PAPEL, charSpacing: 1, margin: 0
});

s1.addText('EL PROBLEMA REAL', {
  x: 0.6, y: 1.65, w: 5.6, h: 0.3,
  fontFace: F_MONO, fontSize: 10, bold: true, color: ORO, charSpacing: 2, margin: 0
});

s1.addText(
  'Las asociaciones cacaoteras de Esmeraldas eligen sus rutas de exportación sin un criterio claro: pagan una ruta urgente que no necesitaban, o mueven la carga tan lento que el grano pierde calidad.',
  { x: 0.6, y: 2.0, w: 5.6, h: 1.05, fontFace: F_TXT, fontSize: 15, color: PAPEL, lineSpacing: 22, margin: 0 }
);

s1.addText('LAS TRES VARIABLES DEL SISTEMA', {
  x: 0.6, y: 3.25, w: 5.6, h: 0.3,
  fontFace: F_MONO, fontSize: 10, bold: true, color: ORO, charSpacing: 2, margin: 0
});

const vars = [
  ['p', 'El envío es urgente', 'contrato con multa o grano que se puede dañar'],
  ['q', 'El presupuesto está excedido', 'el costo pasa el techo asignado al lote'],
  ['r', 'Peso total de la carga', 'define el valor y cuántos contenedores hacen falta']
];
vars.forEach((v, i) => {
  const y = 3.65 + i * 0.92;
  s1.addShape(pres.ShapeType.rect, {
    x: 0.6, y, w: 0.55, h: 0.72, fill: { color: ORO }, line: { color: ORO }
  });
  s1.addText(v[0], {
    x: 0.6, y, w: 0.55, h: 0.72,
    fontFace: F_MONO, fontSize: 20, bold: true, color: TINTA, align: 'center', valign: 'middle', margin: 0
  });
  s1.addText(v[1], {
    x: 1.3, y: y + 0.03, w: 4.9, h: 0.32,
    fontFace: F_TXT, fontSize: 15, bold: true, color: PAPEL, margin: 0
  });
  s1.addText(v[2], {
    x: 1.3, y: y + 0.35, w: 4.9, h: 0.32,
    fontFace: F_TXT, fontSize: 12, color: YUTE, margin: 0
  });
});

// Bloque estarcido con la regla
s1.addShape(pres.ShapeType.rect, {
  x: 6.9, y: 1.65, w: 5.8, h: 4.95,
  fill: { color: TINTA }, line: { color: ORO, width: 2 }
});

s1.addText('LA REGLA DE DECISIÓN', {
  x: 7.2, y: 1.9, w: 5.2, h: 0.3,
  fontFace: F_MONO, fontSize: 10, bold: true, color: ORO_VIVO, charSpacing: 2, margin: 0
});

s1.addText('D = p ∧ ¬q', {
  x: 7.2, y: 2.2, w: 5.2, h: 0.6,
  fontFace: F_MONO, fontSize: 30, bold: true, color: ORO_VIVO, margin: 0
});

s1.addText('Se lee: se aprueba la ruta rápida si el envío es urgente Y el presupuesto NO está excedido.', {
  x: 7.2, y: 2.82, w: 5.2, h: 0.55,
  fontFace: F_TXT, fontSize: 13, color: PAPEL, italic: true, lineSpacing: 18, margin: 0
});

const filas = [
  ['p', 'q', '¬q', 'D'],
  ['V', 'V', 'F', 'F'],
  ['V', 'F', 'V', 'V'],
  ['F', 'V', 'F', 'F'],
  ['F', 'F', 'V', 'F']
];
const tabla = filas.map((f, i) => {
  const esCab = i === 0;
  const aprueba = i === 2;
  return f.map(c => ({
    text: c,
    options: {
      fontFace: F_MONO,
      fontSize: esCab ? 12 : 15,
      bold: true,
      color: esCab ? YUTE : (aprueba ? TINTA : PAPEL),
      fill: { color: esCab ? TINTA : (aprueba ? ORO_VIVO : TINTA) },
      align: 'center',
      valign: 'middle'
    }
  }));
});
s1.addTable(tabla, {
  x: 7.2, y: 3.55, w: 3.4, colW: [0.85, 0.85, 0.85, 0.85], rowH: 0.36,
  border: { type: 'solid', color: YUTE, pt: 1 }
});

s1.addText('← de las 4 combinaciones\n   posibles, solo esta\n   aprueba la ruta rápida', {
  x: 10.75, y: 4.05, w: 1.75, h: 0.75,
  fontFace: F_MONO, fontSize: 9, color: ORO_VIVO, lineSpacing: 12, valign: 'middle', margin: 0
});

s1.addText('Inferencia válida · Modus Ponens', {
  x: 7.2, y: 5.6, w: 5.2, h: 0.28,
  fontFace: F_MONO, fontSize: 10, bold: true, color: ORO, charSpacing: 1.5, margin: 0
});
s1.addText('Si (p ∧ ¬q) → D    y ocurre    p ∧ ¬q    entonces    ∴ D', {
  x: 7.2, y: 5.9, w: 5.2, h: 0.3,
  fontFace: F_MONO, fontSize: 12, color: PAPEL, margin: 0
});
s1.addText('Su forma [(A→B) ∧ A] → B es una tautología: siempre verdadera.', {
  x: 7.2, y: 6.2, w: 5.2, h: 0.3,
  fontFace: F_TXT, fontSize: 12, color: YUTE, italic: true, margin: 0
});

s1.addNotes(
  'El problema real: las asociaciones deciden a ojo. Definí tres variables. ' +
  'p y q son proposiciones, o sea que solo pueden ser verdaderas o falsas, y por eso son las que entran en la tabla de verdad. ' +
  'r es numérica: alimenta el valor de la carga y el número de contenedores. ' +
  'La regla dice que la ruta rápida se aprueba solo si el envío es urgente Y el presupuesto no está excedido. ' +
  'Con dos proposiciones hay cuatro combinaciones posibles, y como ven en la tabla, una sola aprueba: cuando p es verdadera y q es falsa. ' +
  'Para validar la regla usé Modus Ponens, cuya forma es una tautología, así que si las premisas son verdaderas la conclusión no puede ser falsa.'
);

/* =========================== 3 · MODELO ALGEBRAICO =========================== */
const s2 = pres.addSlide();
s2.background = { color: PAPEL };

rotulo(s2, 'PARTE 2 · MODELO ALGEBRAICO', 0.6, 0.45, ORO);

s2.addText('¿CÓMO REPARTO EL DINERO Y CUÁNTO EXPORTO?', {
  x: 0.6, y: 0.75, w: 12.1, h: 0.6,
  fontFace: F_TIT, fontSize: 32, bold: true, color: TINTA, charSpacing: 1, margin: 0
});

// --- Sistema 2x2 ---
s2.addShape(pres.ShapeType.rect, {
  x: 0.6, y: 1.62, w: 5.9, h: 2.35,
  fill: { color: TENDAL }, line: { color: YUTE, width: 1 }
});
s2.addText('SISTEMA 2×2 · REPARTIR EL PRESUPUESTO', {
  x: 0.85, y: 1.8, w: 5.4, h: 0.28,
  fontFace: F_MONO, fontSize: 10, bold: true, color: ORO, charSpacing: 1.5, margin: 0
});
s2.addText('T + S = 3 500        T = transporte\nT = 4S               S = seguro', {
  x: 0.85, y: 2.1, w: 5.4, h: 0.55,
  fontFace: F_MONO, fontSize: 13, bold: true, color: TINTA, lineSpacing: 19, margin: 0
});
s2.addText('Por sustitución:  4S + S = 3 500  →  5S = 3 500\nS = 700          T = 4(700) = 2 800', {
  x: 0.85, y: 2.7, w: 5.4, h: 0.55,
  fontFace: F_MONO, fontSize: 11, color: GRIS, lineSpacing: 17, margin: 0
});
s2.addText('✓ 2 800 + 700 = 3 500        ✓ 2 800 = 4 × 700', {
  x: 0.85, y: 3.32, w: 5.4, h: 0.32,
  fontFace: F_MONO, fontSize: 12, bold: true, color: VERDE, margin: 0
});
s2.addText('Cada envío lleva $2 800 de transporte y $700 de seguro.', {
  x: 0.85, y: 3.62, w: 5.4, h: 0.3,
  fontFace: F_TXT, fontSize: 12, italic: true, color: GRIS, margin: 0
});

// --- Función de utilidad ---
s2.addShape(pres.ShapeType.rect, {
  x: 0.6, y: 4.15, w: 5.9, h: 2.7,
  fill: { color: TENDAL }, line: { color: YUTE, width: 1 }
});
s2.addText('FUNCIÓN DE UTILIDAD · CUÁNTO EXPORTAR AL MES', {
  x: 0.85, y: 4.33, w: 5.4, h: 0.28,
  fontFace: F_MONO, fontSize: 10, bold: true, color: ORO, charSpacing: 1.5, margin: 0
});
s2.addText('Supuesto: comprar más cacao encarece el precio,\nporque hay que ir a fincas más lejanas.', {
  x: 0.85, y: 4.62, w: 5.4, h: 0.5,
  fontFace: F_TXT, fontSize: 12, color: GRIS, lineSpacing: 16, margin: 0
});
s2.addText('P(x) = 4 700 + 100x    (dólares por tonelada)', {
  x: 0.85, y: 5.12, w: 5.4, h: 0.28,
  fontFace: F_MONO, fontSize: 11, color: TINTA, margin: 0
});
s2.addText('Compra = x · 25 · P(x) = 117 500x + 2 500x²', {
  x: 0.85, y: 5.4, w: 5.4, h: 0.28,
  fontFace: F_MONO, fontSize: 11, color: TINTA, margin: 0
});
s2.addText('U(x) = −2 500x² + 23 550x − 20 000', {
  x: 0.85, y: 5.75, w: 5.4, h: 0.42,
  fontFace: F_MONO, fontSize: 15, bold: true, color: TINTA, margin: 0
});
s2.addText('El x² no lo inventé: sale de multiplicar contenedores por un precio que depende de los contenedores.', {
  x: 0.85, y: 6.2, w: 5.4, h: 0.5,
  fontFace: F_TXT, fontSize: 11.5, italic: true, color: GRIS, lineSpacing: 15, margin: 0
});

// --- Gráfica y cifras ---
s2.addImage({ path: IMG, x: 6.9, y: 1.62, w: 5.8, h: 3.55 });

const kpis = [
  ['VÉRTICE', '(4.71 ; $35 460)', 'el punto más alto'],
  ['RAÍCES', '0.94  y  8.48', 'donde no gano ni pierdo'],
  ['ÓPTIMO REAL', '5 contenedores', '$35 250 al mes']
];
kpis.forEach((k, i) => {
  const x = 6.9 + i * 1.97;
  const dest = i === 2;
  s2.addShape(pres.ShapeType.rect, {
    x, y: 5.35, w: 1.85, h: 1.5,
    fill: { color: dest ? TINTA : TENDAL }, line: { color: dest ? ORO : YUTE, width: dest ? 2 : 1 }
  });
  s2.addText(k[0], {
    x: x + 0.15, y: 5.5, w: 1.55, h: 0.25,
    fontFace: F_MONO, fontSize: 8.5, bold: true, color: dest ? ORO_VIVO : ORO, charSpacing: 1, margin: 0
  });
  s2.addText(k[1], {
    x: x + 0.15, y: 5.76, w: 1.55, h: 0.5,
    fontFace: F_MONO, fontSize: 12.5, bold: true, color: dest ? PAPEL : TINTA, margin: 0
  });
  s2.addText(k[2], {
    x: x + 0.15, y: 6.3, w: 1.55, h: 0.42,
    fontFace: F_TXT, fontSize: 10, color: dest ? PAPEL : GRIS, lineSpacing: 12, margin: 0
  });
});

s2.addNotes(
  'Dos cosas resuelvo con álgebra. Primero, cómo reparto el presupuesto de cada envío: un sistema de dos ecuaciones con dos incógnitas, transporte y seguro. ' +
  'Lo resolví por sustitución y lo comprobé en las dos ecuaciones: 2 800 de transporte y 700 de seguro. ' +
  'Segundo, cuántos contenedores conviene exportar al mes. Ahí el supuesto es que comprar más cacao lo encarece, porque toca ir a fincas más lejanas: el precio sube 100 dólares por tonelada por cada contenedor. ' +
  'Al multiplicar los contenedores por ese precio aparece el x cuadrado solo, no lo inventé. ' +
  'La parábola abre hacia abajo, o sea que tiene un máximo. El vértice cae en 4.71, pero contenedores fraccionarios no existen, así que evalué 4 y 5: gana 5, con 35 250 dólares al mes.'
);

/* =================== 4 · HERRAMIENTA, MEDIDAS Y CONCLUSIÓN =================== */
const s3 = pres.addSlide();
s3.background = { color: TINTA };

rotulo(s3, 'PARTE 3 · LA HERRAMIENTA Y EL RESULTADO', 0.6, 0.45);

s3.addText('TODO ESTO FUNCIONANDO EN UNA PÁGINA WEB', {
  x: 0.6, y: 0.75, w: 12.1, h: 0.6,
  fontFace: F_TIT, fontSize: 32, bold: true, color: PAPEL, charSpacing: 1, margin: 0
});

// --- Maqueta de la aplicación ---
s3.addShape(pres.ShapeType.roundRect, {
  x: 0.6, y: 1.6, w: 6.1, h: 3.55, rectRadius: 0.06,
  fill: { color: PAPEL }, line: { color: YUTE, width: 1 }
});
s3.addShape(pres.ShapeType.rect, {
  x: 0.6, y: 1.6, w: 6.1, h: 0.34, fill: { color: 'D6D0BE' }, line: { color: YUTE, width: 1 }
});
s3.addText('pepa-de-oro · calculadora de exportación', {
  x: 0.8, y: 1.63, w: 5.7, h: 0.28,
  fontFace: F_MONO, fontSize: 9, color: GRIS, valign: 'middle', margin: 0
});

// Pasos de la calculadora
const pasos = ['1  Provincia de origen', '2  Cuánto cacao', '3  País de destino', '4  Aire, mar o tierra'];
pasos.forEach((p, i) => {
  s3.addShape(pres.ShapeType.rect, {
    x: 0.85, y: 2.12 + i * 0.44, w: 2.5, h: 0.36,
    fill: { color: TENDAL }, line: { color: 'D6D0BE', width: 1 }
  });
  s3.addText(p, {
    x: 0.95, y: 2.12 + i * 0.44, w: 2.3, h: 0.36,
    fontFace: F_TXT, fontSize: 11, color: TINTA, valign: 'middle', margin: 0
  });
});

// La marca de embarque, elemento firma de la aplicación
s3.addShape(pres.ShapeType.rect, {
  x: 3.6, y: 2.12, w: 2.85, h: 2.75,
  fill: { color: TINTA }, line: { color: ORO, width: 2 }
});
s3.addText('MARCA DE EMBARQUE', {
  x: 3.78, y: 2.26, w: 2.5, h: 0.22,
  fontFace: F_MONO, fontSize: 7, bold: true, color: ORO_VIVO, charSpacing: 1, margin: 0
});
s3.addText('GYE  →  RTM', {
  x: 3.78, y: 2.5, w: 2.5, h: 0.42,
  fontFace: F_TIT, fontSize: 21, bold: true, color: ORO_VIVO, margin: 0
});
s3.addText('Guayaquil → Rotterdam', {
  x: 3.78, y: 2.92, w: 2.5, h: 0.22,
  fontFace: F_MONO, fontSize: 7.5, color: YUTE, margin: 0
});
s3.addText(
  'PESO NETO      25 000 kg\nBULTOS         1 × 25 t\nVÍA            marítimo\nTRÁNSITO       21.5 días',
  { x: 3.78, y: 3.2, w: 2.5, h: 0.95, fontFace: F_MONO, fontSize: 8, color: PAPEL, lineSpacing: 12, margin: 0 }
);
s3.addText('COSTO TOTAL', {
  x: 3.78, y: 4.2, w: 2.5, h: 0.2,
  fontFace: F_MONO, fontSize: 7, bold: true, color: YUTE, charSpacing: 1, margin: 0
});
s3.addText('$6 452.00', {
  x: 3.78, y: 4.4, w: 2.5, h: 0.4,
  fontFace: F_TIT, fontSize: 20, bold: true, color: ORO_VIVO, margin: 0
});

s3.addText('La página resuelve el sistema, evalúa la función y aplica la regla lógica en vivo.', {
  x: 0.6, y: 5.25, w: 6.1, h: 0.32,
  fontFace: F_TXT, fontSize: 12, italic: true, color: YUTE, margin: 0
});

// --- Medidas ---
s3.addText('MEDIDAS Y CONVERSIONES DEL LOTE', {
  x: 7.1, y: 1.6, w: 5.6, h: 0.28,
  fontFace: F_MONO, fontSize: 10, bold: true, color: ORO, charSpacing: 1.5, margin: 0
});
s3.addShape(pres.ShapeType.rect, {
  x: 7.1, y: 1.92, w: 5.6, h: 1.72,
  fill: { color: TINTA }, line: { color: ORO, width: 2 }
});
s3.addText(
  '10 700 km × 0.621371 = 6 648.67 millas\n' +
  '25 000 kg × 2.20462  = 55 115.50 libras\n' +
  '25 000 kg ÷ 45.36    = 551 quintales\n' +
  'Seguro: 700 / 147 500 = 0.47 % del FOB',
  { x: 7.32, y: 2.12, w: 5.2, h: 1.35, fontFace: F_MONO, fontSize: 12, color: PAPEL, lineSpacing: 20, margin: 0 }
);
s3.addText('El quintal (100 lb = 45.36 kg) es la unidad real con la que se comercia el cacao en Ecuador.', {
  x: 7.1, y: 3.72, w: 5.6, h: 0.45,
  fontFace: F_TXT, fontSize: 11.5, italic: true, color: YUTE, lineSpacing: 15, margin: 0
});

// --- Datos con fuente ---
s3.addText('DATOS REALES', {
  x: 7.1, y: 4.3, w: 5.6, h: 0.28,
  fontFace: F_MONO, fontSize: 10, bold: true, color: ORO, charSpacing: 1.5, margin: 0
});
s3.addText(
  'Cacao $5 900/t (Anecacao) · Esmeraldas 58 965 t/año (MAG)\nFlete 40\' ≈ $2 286 (Drewry) · Arancel UE 0 % (Ac. Multipartes)',
  { x: 7.1, y: 4.6, w: 5.6, h: 0.6, fontFace: F_TXT, fontSize: 12, color: PAPEL, lineSpacing: 17, margin: 0 }
);

// --- Conclusión ---
s3.addShape(pres.ShapeType.rect, {
  x: 0.6, y: 5.7, w: 12.1, h: 1.15,
  fill: { color: ORO }, line: { color: ORO_VIVO, width: 2 }
});
s3.addText('CONCLUSIÓN', {
  x: 0.9, y: 5.85, w: 3, h: 0.25,
  fontFace: F_MONO, fontSize: 10, bold: true, color: TINTA, charSpacing: 2.5, margin: 0
});
s3.addText('Conviene exportar 5 contenedores al mes (125 toneladas)', {
  x: 0.9, y: 6.08, w: 11.5, h: 0.46,
  fontFace: F_TIT, fontSize: 23, bold: true, color: TINTA, margin: 0
});
s3.addText('La lógica dice CUÁNDO usar la ruta rápida · el sistema dice CÓMO repartir el dinero · la función dice CUÁNTO exportar', {
  x: 0.9, y: 6.53, w: 11.5, h: 0.28,
  fontFace: F_TXT, fontSize: 12.5, color: '3A2410', margin: 0
});

s3.addNotes(
  'Todo esto no se queda en el papel: lo programé en una página web. ' +
  'El usuario elige su provincia, cuánto cacao tiene, a qué país lo manda y cómo lo transporta, y la página le devuelve esta marca de embarque, que es el formato que se usa en el oficio: puerto de salida, puerto de llegada, peso, bultos, vía, días y costo. ' +
  'Ahí también están las conversiones entre sistema internacional e imperial, y usé el quintal porque es la unidad con la que un productor realmente piensa. ' +
  'Todos los datos tienen fuente: Anecacao, el Ministerio de Agricultura, el índice Drewry. ' +
  'La conclusión es que conviene exportar 5 contenedores al mes. Ahora se los muestro funcionando.'
);

const salida = path.join(__dirname, 'Presentacion_Pepa_de_Oro.pptx');
pres.writeFile({ fileName: salida }).then(() => console.log('Creado:', salida));
