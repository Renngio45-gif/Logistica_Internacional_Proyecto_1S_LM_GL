# -*- coding: utf-8 -*-
"""Paper académico del proyecto Pepa de Oro.
Estructura tomada del documento de referencia SIAM; citas y referencias en APA 7.
"""
import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE = os.path.dirname(os.path.abspath(__file__))
GRAF = os.path.join(BASE, 'graficas_clase9')

NEGRO = RGBColor(0x00, 0x00, 0x00)
GRIS = RGBColor(0x44, 0x44, 0x44)
HEX_CAB = 'D9D9D9'
HEX_ALT = 'F2F2F2'


def cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def cell_pad(cell, v=90, h=110):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement('w:tcMar')
    for n, val in [('top', v), ('bottom', v), ('left', h), ('right', h)]:
        e = OxmlElement(f'w:{n}')
        e.set(qn('w:w'), str(val))
        e.set(qn('w:type'), 'dxa')
        mar.append(e)
    tcPr.append(mar)


def borders(table, color='999999'):
    tblPr = table._tbl.tblPr
    b = OxmlElement('w:tblBorders')
    for edge in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        e = OxmlElement(f'w:{edge}')
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), '4')
        e.set(qn('w:color'), color)
        b.append(e)
    tblPr.append(b)


doc = Document()
for s in doc.sections:
    s.top_margin = Inches(1)
    s.bottom_margin = Inches(1)
    s.left_margin = Inches(1)
    s.right_margin = Inches(1)

st = doc.styles['Normal']
st.font.name = 'Times New Roman'
st.font.size = Pt(11)
st.font.color.rgb = NEGRO
st.paragraph_format.line_spacing = 1.5          # APA 7: doble o 1.5
st.paragraph_format.space_after = Pt(6)


def p(text='', size=11, bold=False, italic=False, color=None,
      align=WD_ALIGN_PARAGRAPH.JUSTIFY, space=6, indent=0, first=0, ls=1.5):
    par = doc.add_paragraph()
    par.alignment = align
    par.paragraph_format.space_after = Pt(space)
    par.paragraph_format.line_spacing = ls
    if indent:
        par.paragraph_format.left_indent = Inches(indent)
    if first:
        par.paragraph_format.first_line_indent = Inches(first)
    if text:
        r = par.add_run(text)
        r.bold = bold
        r.italic = italic
        r.font.size = Pt(size)
        if color:
            r.font.color.rgb = color
    return par


def h1(text):
    par = p(text, size=12, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space=6, ls=1.5)
    par.paragraph_format.space_before = Pt(14)
    return par


def h2(text):
    par = p(text, size=11, bold=True, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT, space=5, ls=1.5)
    par.paragraph_format.space_before = Pt(10)
    return par


def cuerpo(text, first=0.5):
    return p(text, first=first)


def vinieta(text, indent=0.35):
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    par.paragraph_format.left_indent = Inches(indent)
    par.paragraph_format.space_after = Pt(4)
    par.paragraph_format.line_spacing = 1.5
    r = par.add_run('•  ' + text)
    r.font.size = Pt(11)
    return par


def formula(text, size=12):
    par = p(text, size=size, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space=10, ls=1.0)
    par.paragraph_format.space_before = Pt(8)
    return par


def tabla(titulo, cabeceras, filas, anchos=None, fuente=10):
    cap = p(titulo, size=10, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space=4, ls=1.0)
    cap.paragraph_format.space_before = Pt(10)
    t = doc.add_table(rows=1, cols=len(cabeceras))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    borders(t)
    hdr = t.rows[0].cells
    for i, c in enumerate(cabeceras):
        hdr[i].text = ''
        cell_bg(hdr[i], HEX_CAB)
        cell_pad(hdr[i])
        par = hdr[i].paragraphs[0]
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.line_spacing = 1.0
        par.paragraph_format.space_after = Pt(0)
        r = par.add_run(c)
        r.bold = True
        r.font.size = Pt(fuente)
    for ri, fila in enumerate(filas):
        row = t.add_row().cells
        bg = 'FFFFFF' if ri % 2 == 0 else HEX_ALT
        for ci, c in enumerate(fila):
            row[ci].text = ''
            cell_bg(row[ci], bg)
            cell_pad(row[ci])
            par = row[ci].paragraphs[0]
            par.alignment = WD_ALIGN_PARAGRAPH.LEFT if ci == 0 else WD_ALIGN_PARAGRAPH.CENTER
            par.paragraph_format.line_spacing = 1.0
            par.paragraph_format.space_after = Pt(0)
            r = par.add_run(c)
            r.font.size = Pt(fuente)
    if anchos:
        for i, w in enumerate(anchos):
            t.columns[i].width = Inches(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def figura(archivo, ancho, epigrafe):
    ruta = os.path.join(GRAF, archivo)
    if not os.path.exists(ruta):
        p(f'[Falta la figura {archivo}]', italic=True)
        return
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_after = Pt(4)
    par.paragraph_format.space_before = Pt(8)
    par.add_run().add_picture(ruta, width=Inches(ancho))
    cap = p(epigrafe, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER, space=12, ls=1.0)
    cap.runs[0].font.color.rgb = GRIS


# ============================== ENCABEZADO ==============================
p('Pepa de Oro: modelo lógico-matemático y herramienta web para la planificación '
  'de exportaciones de cacao fino de aroma desde Esmeraldas, Ecuador',
  size=15, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space=10, ls=1.2)

p('Geovanny Daniel Farías Estupiñán ¹', size=11.5, bold=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, space=3, ls=1.2)

p('¹ Pontificia Universidad Católica del Ecuador, Sede Esmeraldas (PUCESE); '
  'gdfarias@pucese.edu.ec', size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER, space=4, ls=1.2)

p('Carrera de Tecnología Superior en Desarrollo de Software (PUCETEC). '
  'Asignatura: Habilidades Lógico-Matemáticas. Docente: Msc. Adrián Vargas. '
  'Período académico 2026-1.', size=9.5, italic=True,
  align=WD_ALIGN_PARAGRAPH.CENTER, space=14, ls=1.2)

# ------------------------------ RESUMEN ------------------------------
h1('Resumen')
p('Las asociaciones cacaoteras de Esmeraldas deciden sus envíos al exterior sin criterios '
  'formalizados. El resultado son rutas urgentes que nadie necesitaba y grano que pierde '
  'calidad en tránsitos demasiado largos. Este trabajo formaliza esa decisión en un modelo '
  'lógico-matemático y lo implementa en una herramienta web llamada Pepa de Oro. El modelo '
  'tiene tres piezas. La primera es una regla de decisión en lógica proposicional, '
  'D = p ∧ ¬q, con su tabla de verdad y validada por Modus Ponens. La segunda es un sistema '
  'de ecuaciones lineales 2×2 que reparte el presupuesto logístico entre transporte y '
  'seguro. La tercera es una función cuadrática de utilidad, U(x) = −2 500x² + 23 550x − '
  '20 000, cuyo vértice fija el volumen mensual que conviene exportar. Se trabajó con datos '
  'de mercado de 2026 y con conversiones entre el Sistema Internacional y el Imperial, '
  'incluido el quintal, que es la unidad con la que el sector cacaotero ecuatoriano comercia '
  'realmente. El óptimo hallado son cinco contenedores mensuales —125 toneladas— con una '
  'utilidad estimada de USD 35 250. La herramienta se programó con tecnologías web estándar '
  'y se despliega como sitio estático; admite variar provincia de origen, destino, modalidad '
  'de transporte y volumen. Un análisis de sensibilidad muestra que la forma de la '
  'conclusión no depende del único parámetro que no pudo verificarse.',
  space=8)

par = p('', space=14)
r = par.add_run('Palabras clave: ')
r.bold = True
r.font.size = Pt(11)
r2 = par.add_run('lógica proposicional; sistemas de ecuaciones; función cuadrática; '
                 'optimización; logística internacional; exportación de cacao; '
                 'conversión de unidades; Esmeraldas.')
r2.italic = True
r2.font.size = Pt(11)

# ============================== 1. INTRODUCCIÓN ==============================
h1('1. Introducción')

cuerpo('El cacao es uno de los principales rubros de exportación no petrolera del Ecuador. '
       'En 2025 el país despachó cerca de 600 000 toneladas de grano y derivados, por un '
       'valor aproximado de USD 4 668 millones (Primicias, 2025). Esmeraldas produce 58 965 '
       'toneladas anuales en 90 248 hectáreas y ocupa el tercer lugar nacional (Ministerio '
       'de Agricultura y Ganadería [MAG], 2026). Su cacao fino de aroma tiene demanda en la '
       'Unión Europea y en Estados Unidos (Cámara Marítima del Ecuador [CAMAE], 2025).')

cuerpo('El peso económico del sector no se refleja en su planificación. En el eslabón '
       'asociativo las decisiones logísticas se siguen tomando por experiencia acumulada. '
       'Tres de ellas se repiten en cada envío: cuándo conviene pagar una ruta rápida más '
       'cara, cómo repartir el presupuesto entre transporte y seguro, y qué volumen mensual '
       'deja la mayor utilidad. Rara vez se resuelven juntas. Sin un instrumento que las '
       'integre, la asociación termina pagando urgencias que no necesitaba o perdiendo '
       'calidad de grano en tránsitos largos.')

cuerpo('Cada una de esas decisiones cae en un dominio matemático distinto, y ahí está el '
       'interés del caso. La primera se reduce a la conjunción de dos condiciones binarias, '
       'de modo que admite una formalización en lógica proposicional. La segunda es un '
       'sistema de ecuaciones lineales con dos incógnitas. La tercera exige optimizar, '
       'porque la relación entre volumen y utilidad no crece sin límite: pasado cierto '
       'punto, exportar más deteriora el margen.')

cuerpo('Este trabajo propone Pepa de Oro: un modelo integrado y su implementación como '
       'herramienta web de acceso público. El nombre retoma la denominación tradicional del '
       'cacao ecuatoriano. Las técnicas empleadas son elementales y corresponden al primer '
       'nivel de formación. El aporte está en otra parte: en articularlas sobre un caso '
       'real, con datos verificables y con los supuestos declarados donde el lector pueda '
       'verlos.')

# ============================== 2. OBJETIVOS ==============================
h1('2. Objetivos')

h2('2.1. Objetivo general')
cuerpo('Diseñar e implementar un modelo lógico-matemático integrado, materializado en una '
       'herramienta web, que permita a las asociaciones cacaoteras de Esmeraldas determinar '
       'la modalidad de ruta, la distribución del presupuesto logístico y el volumen mensual '
       'óptimo de exportación de cacao fino de aroma, con base en datos de mercado '
       'verificables.')

h2('2.2. Objetivos específicos')
vinieta('Formalizar mediante lógica proposicional la regla de decisión que determina la '
        'aprobación de una ruta rápida, construyendo su tabla de verdad y validándola con '
        'una regla de inferencia.')
vinieta('Plantear y resolver un sistema de ecuaciones lineales 2×2 que distribuya el '
        'presupuesto logístico de cada envío entre los rubros de transporte y seguro, '
        'verificando la solución en las ecuaciones originales.')
vinieta('Derivar una función de utilidad a partir de la estructura real de ingresos y costos '
        'de la operación, y determinar su punto óptimo considerando un dominio con sentido '
        'físico.')
vinieta('Establecer las conversiones entre el Sistema Internacional y el Sistema Imperial '
        'requeridas por la operación exportadora, incorporando la unidad tradicional del '
        'sector cacaotero ecuatoriano.')
vinieta('Implementar el modelo en una herramienta web parametrizable, desplegable como sitio '
        'estático y accesible desde dispositivos de escritorio y móviles.')
vinieta('Evaluar la robustez de la recomendación mediante análisis de sensibilidad sobre el '
        'supuesto de modelado no verificado empíricamente.')

# ============================== 3. ALCANCE ==============================
h1('3. Alcance')

h2('3.1. Elementos incluidos')
vinieta('Formalización lógica de la decisión de ruta con dos proposiciones atómicas.')
vinieta('Sistema de ecuaciones lineales 2×2 resuelto por el método de sustitución.')
vinieta('Función cuadrática de utilidad con análisis de vértice, raíces, dominio, rango y '
        'monotonía.')
vinieta('Conversiones de distancia, masa y volumen entre sistemas de medida.')
vinieta('Herramienta web con nueve destinos internacionales, ocho provincias de origen y '
        'tres modalidades de transporte.')
vinieta('Análisis de sensibilidad sobre el parámetro de encarecimiento del acopio.')

h2('3.2. Elementos excluidos')
vinieta('Integración con sistemas aduaneros o plataformas de comercio exterior en tiempo real.')
vinieta('Cotizaciones en firme de navieras o aerolíneas; las tarifas empleadas son promedios '
        'de mercado con carácter referencial.')
vinieta('Modelado de la variabilidad estacional del precio internacional del cacao.')
vinieta('Optimización multiobjetivo o programación lineal con restricciones múltiples.')

h2('3.3. Supuestos')
cuerpo('El modelo da por sentado que la asociación tiene capacidad de acopio para el volumen '
       'analizado y que el precio de venta no varía dentro del mes. También asume que la '
       'producción provincial no limita la operación; este último supuesto sí puede '
       'comprobarse, porque el óptimo hallado representa apenas el 2,5 % de lo que Esmeraldas '
       'produce al mes.')

# ============================== 4. METODOLOGÍA ==============================
h1('4. Fundamento teórico y metodología')

cuerpo('El trabajo siguió un enfoque de modelado matemático aplicado: traducir una situación '
       'real a un lenguaje formal, resolverla con las herramientas de ese lenguaje y '
       'devolver el resultado al problema original. La metodología se organizó en cuatro '
       'fases.')

cuerpo('Primero se caracterizó el sistema y se identificaron sus variables. Después se '
       'formalizó la regla de decisión y se construyó el modelo algebraico. La tercera fase '
       'derivó la función de utilidad y localizó su óptimo. La última implementó la '
       'herramienta web y contrastó sus resultados numéricos contra los cálculos '
       'analíticos.')

cuerpo('Los datos de mercado provienen de fuentes institucionales y de prensa económica '
       'especializada, con preferencia por publicaciones de 2026. Los parámetros que no pudo '
       'respaldarse documentalmente se declararon como supuestos y se sometieron a análisis '
       'de sensibilidad, que mide cuánto depende la conclusión del valor asumido.')

# ============================== 5. MODELO LÓGICO ==============================
h1('5. Modelo lógico del sistema')

h2('5.1. Variables del sistema')
cuerpo('El sistema se describe con tres variables. Dos son proposiciones atómicas: solo '
       'admiten verdadero o falso. La tercera es numérica y alimenta los cálculos '
       'posteriores.')

vinieta('p: el envío es calificado como urgente, sea por riesgo de deterioro del grano o por '
        'la existencia de una cláusula contractual de penalización por retraso.')
vinieta('q: el presupuesto asignado al lote se encuentra excedido por los costos calculados.')
vinieta('r: peso total de la carga, expresado en kilogramos. Determina el valor de la '
        'mercancía, el número de contenedores requeridos y las conversiones de unidades.')

h2('5.2. Regla de decisión')
cuerpo('La aprobación de la ruta rápida y directa se formaliza mediante la conjunción de la '
       'urgencia con la negación del exceso presupuestario:')

formula('D = p ∧ ¬q', 14)

cuerpo('La expresión dice que la ruta rápida se aprueba si y solo si el envío es urgente y, '
       'al mismo tiempo, el presupuesto no está excedido. Con dos proposiciones atómicas el '
       'espacio de casos es 2² = 4 combinaciones. La Tabla 1 las recoge todas.')

tabla('Tabla 1. Tabla de verdad de la regla de decisión D = p ∧ ¬q.',
      ['p (urgente)', 'q (presupuesto excedido)', '¬q', 'D', 'Interpretación operativa'],
      [
          ['V', 'V', 'F', 'F', 'Urgente, pero el costo excede el techo'],
          ['V', 'F', 'V', 'V', 'Se aprueba la ruta rápida'],
          ['F', 'V', 'F', 'F', 'Ni urgente ni viable financieramente'],
          ['F', 'F', 'V', 'F', 'No amerita ruta rápida; se usa la estándar'],
      ],
      anchos=[0.9, 1.5, 0.6, 0.6, 2.9])

cuerpo('Solo la segunda fila da verdadero. El resultado coincide con la lógica del negocio: '
       'el sobrecosto de la ruta rápida se justifica cuando hay una necesidad temporal real y '
       'además dinero para cubrirla. Si falta cualquiera de las dos, conviene la ruta '
       'estándar.')

h2('5.3. Validación mediante inferencia')
cuerpo('Para verificar que la regla deriva conclusiones válidas se aplicó Modus Ponens:')

formula('Premisa 1:  (p ∧ ¬q) → D\nPremisa 2:  p ∧ ¬q\nConclusión:  ∴ D', 11.5)

cuerpo('La inferencia es válida porque su forma, [(A → B) ∧ A] → B, es una tautología: sale '
       'verdadera para cualquier asignación de sus variables. Eso vuelve imposible que las '
       'premisas sean verdaderas y la conclusión falsa. El procedimiento de decisión que '
       'ejecuta la herramienta queda así formalmente respaldado.')

# ============================== 6. MODELO ALGEBRAICO ==============================
h1('6. Modelo algebraico')

h2('6.1. Distribución del presupuesto logístico')
cuerpo('El presupuesto de cada envío se reparte entre dos rubros: transporte internacional '
       '(T) y seguro de la carga (S), ambos en dólares. La notación no es casual. Se '
       'evitaron las letras x e y porque la sección 7 usa x para contar contenedores, que es '
       'una magnitud distinta, y arrastrar la misma letra habría confundido las dos '
       'lecturas.')

cuerpo('Dos condiciones operativas definen el sistema. El presupuesto conjunto del lote es de '
       'USD 3 500. Y el transporte cuadruplica al seguro, proporción que se desprende de los '
       'valores de mercado: el flete de un contenedor de cuarenta pies promedia USD 2 286 '
       '(Drewry, 2026), mientras que el seguro de carga oscila entre el 0,3 % y el 1,5 % del '
       'valor CIF (Logintec, 2026; Transnatur, 2024). El sistema queda así:')

formula('T + S = 3 500\nT = 4S', 13)

cuerpo('Por sustitución, la segunda ecuación entra en la primera y deja una sola incógnita:')

formula('4S + S = 3 500  →  5S = 3 500  →  S = 700\nT = 4(700) = 2 800', 11.5)

cuerpo('La solución se comprueba en las dos ecuaciones originales: 2 800 + 700 = 3 500 y '
       '2 800 = 4 × 700. Cada envío necesita entonces USD 2 800 de transporte y USD 700 de '
       'seguro. Ese seguro equivale al 0,47 % del valor FOB del lote, dentro del rango de '
       'mercado citado arriba.')

h2('6.2. Ecuación del costo logístico total')
cuerpo('El costo logístico total de un envío se modela como la suma de los rubros anteriores '
       'y de las comisiones logísticas, aplicadas estas últimas como un porcentaje del valor '
       'de la mercancía:')

formula('C = T + S + (a / 100) · V', 13)

cuerpo('donde a es el porcentaje de comisiones logísticas y V el valor de la mercancía. El '
       'nombre de V no es indiferente: depende de la vía contratada. Los Incoterms 2020 '
       'reservan el término FOB para el transporte marítimo y por vías navegables '
       'interiores; en las modalidades aérea y terrestre corresponde FCA (Cámara de Comercio '
       'Internacional [ICC], 2019). La herramienta cambia la denominación sola según lo que '
       'elija el usuario. Con a = 2 %, el costo logístico del lote modelado llega a USD 6 450 '
       'por contenedor.')

# ============================== 7. MODELO FUNCIONAL ==============================
h1('7. Modelo funcional y determinación del óptimo')

h2('7.1. Supuesto de modelado')
cuerpo('Hallar el volumen óptimo obliga a modelar cómo se relacionan la cantidad exportada y '
       'la utilidad. Si el precio de compra al productor fuera constante, la utilidad '
       'crecería sin techo, lo que no tiene sentido económico. En la práctica ocurre otra '
       'cosa: subir el volumen obliga a acopiar en fincas cada vez más distantes y a pagar '
       'una prima para asegurar el abastecimiento. Es el fenómeno de los rendimientos '
       'decrecientes.')

cuerpo('Ese comportamiento se modeló con una función lineal del precio de compra por tonelada '
       'respecto del número de contenedores x:')

formula('P(x) = 4 700 + 100x     (USD por tonelada)', 13)

cuerpo('El precio base de USD 4 700 por tonelada asume que el productor percibe alrededor del '
       '80 % del precio de exportación. Esa proporción se contrastó con los valores '
       'registrados durante 2026, en la Tabla 2.')

tabla('Tabla 2. Validación de la proporción entre precio al productor y precio internacional.',
      ['Período', 'Precio internacional', 'Pagado al productor', 'Proporción'],
      [
          ['Enero 2026', 'USD 4 500 / t', 'USD 185 / qq = USD 4 078 / t', '90,6 %'],
          ['Mayo 2026', 'USD 4 169 / t', 'USD 151,70 / qq = USD 3 344 / t', '80,2 %'],
          ['Supuesto del modelo', 'USD 5 900 / t', 'USD 4 700 / t', '79,7 %'],
      ],
      anchos=[1.5, 1.5, 2.0, 1.0])

cuerpo('La proporción asumida, 79,7 %, coincide con el 80,2 % observado en mayo de 2026 y cae '
       'en el extremo conservador del rango del año (Expreso, 2026a; Extra, 2026). La brecha '
       'entre ambos precios tiene explicación documentada: los compradores internacionales '
       'aplican al cacao ecuatoriano un diferencial de hasta USD 700 por tonelada, y la '
       'exportación suma entre USD 12 y USD 15 por quintal en gastos (Expreso, 2026b). El '
       'precio base no salió de una estimación libre, sino de una cadena de descuentos '
       'registrada.')

cuerpo('El otro parámetro corre distinta suerte. El incremento de USD 100 por tonelada por '
       'cada contenedor adicional es una estimación propia, sin fuente publicada que la '
       'respalde. La limitación se declara aquí y se aborda en el análisis de sensibilidad de '
       'la sección 7.4.')

h2('7.2. Derivación de la función de utilidad')
cuerpo('La utilidad neta mensual es la diferencia entre el ingreso por ventas y todos los '
       'costos. Cada contenedor transporta 25 toneladas y el precio de exportación es de USD '
       '5 900 por tonelada (Anecacao, 2026). La Tabla 3 desglosa los componentes.')

tabla('Tabla 3. Componentes de la función de utilidad.',
      ['Componente', 'Expresión', 'Aporte a U(x)'],
      [
          ['Ingreso por ventas', 'x · 25 · 5 900', '+147 500x'],
          ['Costo de compra del grano', 'x · 25 · P(x)', '−117 500x − 2 500x²'],
          ['Costo logístico', 'x · 6 450', '−6 450x'],
          ['Costos fijos mensuales', '20 000', '−20 000'],
      ],
      anchos=[2.3, 1.7, 2.0])

cuerpo('Conviene detenerse en el término cuadrático, porque no es un supuesto añadido sino '
       'una consecuencia algebraica de la estructura del modelo. Al multiplicar el volumen '
       'por un precio que ya depende del volumen, el producto genera por fuerza un término de '
       'segundo grado: x · 25 · (4 700 + 100x) = 117 500x + 2 500x². El coeficiente 2 500 sale '
       'de multiplicar las 25 toneladas del contenedor por el incremento de USD 100 por '
       'tonelada, y tiene unidades de dólares por contenedor al cuadrado.')

cuerpo('Agrupando los términos lineales, 147 500 − 117 500 − 6 450 = 23 550, la función de '
       'utilidad queda expresada como:')

formula('U(x) = −2 500x² + 23 550x − 20 000', 14)

h2('7.3. Análisis de la función')
cuerpo('El coeficiente cuadrático es negativo (a = −2 500), así que la parábola abre hacia '
       'abajo y su vértice es un máximo. La abscisa se obtiene con:')

formula('x_v = −b / 2a = −23 550 / (2 · (−2 500)) = 4,71', 12)

cuerpo('con U(4,71) = USD 35 460,25. Las raíces, calculadas por la fórmula general, caen en '
       'x ≈ 0,94 y x ≈ 8,48 y marcan el intervalo rentable. La función crece en (0; 4,71) y '
       'decrece a partir de ahí.')

cuerpo('Hay un matiz que decide el resultado. El dominio matemático de la función es '
       'continuo, pero el dominio con sentido real solo admite enteros: nadie exporta '
       'fracciones de contenedor. Se evaluaron entonces los enteros vecinos al vértice, con '
       'U(4) = USD 34 200 y U(5) = USD 35 250. El óptimo operativo son cinco contenedores '
       'mensuales, es decir 125 toneladas.')

figura('parte_D.png', 5.6,
       'Figura 1. Función de utilidad U(x). Los marcadores cuadrados representan el dominio '
       'con sentido real (contenedores enteros); el vértice teórico se ubica en x = 4,71 y el '
       'óptimo aplicable en x = 5.')

h2('7.4. Análisis de sensibilidad')
cuerpo('Para medir cuánto depende la conclusión del parámetro no verificado, se denota m al '
       'incremento del precio por tonelada por cada contenedor adicional. El coeficiente '
       'cuadrático pasa a ser 25m y la abscisa del vértice, x_v = 23 550 / (50m). La Tabla 4 '
       'recoge los resultados para cuatro valores de m.')

tabla('Tabla 4. Sensibilidad del óptimo respecto del parámetro de encarecimiento del acopio.',
      ['m (USD/t por contenedor)', 'Función U(x)', 'Vértice', 'Óptimo entero', 'Utilidad'],
      [
          ['50', '−1 250x² + 23 550x − 20 000', '9,42', '9 contenedores', 'USD 90 700'],
          ['100 (caso base)', '−2 500x² + 23 550x − 20 000', '4,71', '5 contenedores', 'USD 35 250'],
          ['150', '−3 750x² + 23 550x − 20 000', '3,14', '3 contenedores', 'USD 16 900'],
          ['200', '−5 000x² + 23 550x − 20 000', '2,35', '2 contenedores', 'USD 7 100'],
      ],
      anchos=[1.4, 2.1, 0.7, 1.2, 1.1], fuente=9)

cuerpo('La estructura de la conclusión aguanta. En los cuatro escenarios existe un volumen '
       'óptimo finito y el exceso de volumen siempre deteriora la utilidad. Lo que cambia es '
       'la magnitud del óptimo, que varía en sentido inverso al costo marginal del acopio. '
       'La recomendación de cinco contenedores vale, por tanto, bajo el supuesto de USD 100 '
       'por tonelada. Antes de comprometer volúmenes mayores, la asociación debería medir ese '
       'parámetro en campo.')

# ============================== 8. MEDIDAS ==============================
h1('8. Sistema de medidas y datos de mercado')

cuerpo('Exportar obliga a manejar dos sistemas de unidades a la vez, porque los mercados de '
       'destino no usan las mismas convenciones. La Tabla 5 reúne las conversiones aplicadas '
       'al lote modelado: 25 toneladas con destino a Rotterdam.')

tabla('Tabla 5. Conversiones de unidades aplicadas al lote modelado.',
      ['Magnitud', 'Sistema Internacional', 'Operación', 'Sistema Imperial'],
      [
          ['Distancia de ruta', '10 700 km', '× 0,621371', '6 648,67 mi'],
          ['Masa de la carga', '25 000 kg', '× 2,20462', '55 115,50 lb'],
          ['Combustible terrestre', '188 L', '× 0,264172', '49,66 gal'],
          ['Masa (unidad sectorial)', '25 000 kg', '÷ 45,36', '551 quintales'],
      ],
      anchos=[1.7, 1.6, 1.3, 1.4])

cuerpo('El quintal se incorporó a propósito. Equivale a 100 libras o 45,36 kilogramos y es la '
       'referencia tradicional del cacao en Ecuador: los precios al productor se publican en '
       'dólares por quintal, no por tonelada. Una herramienta pensada para productores '
       'debería hablar en las unidades en que ellos razonan.')

cuerpo('El modelo usa proporcionalidad y porcentajes en varios puntos. El combustible del '
       'tramo terrestre se calculó por proporcionalidad directa desde el rendimiento del '
       'vehículo. El seguro y las comisiones logísticas se expresaron como porcentaje del '
       'valor de la mercancía.')

cuerpo('Los aranceles varían mucho según el destino. El cacao ecuatoriano en grano —partida '
       '1801— entra a la Unión Europea con arancel cero por el Acuerdo Comercial Multipartes '
       '(MAG, 2026). Hacia Estados Unidos, el Acuerdo Comercial Recíproco firmado en marzo de '
       '2026 eliminó la sobretasa, si bien el beneficio cubre solo el grano y deja fuera a '
       'los elaborados (El Universo, 2026; Infobae, 2026). México, en cambio, mantiene una '
       'tarifa de nación más favorecida del 20 % sobre la misma partida (Camtom, 2026). La '
       'herramienta trata cada caso por separado.')

# ============================== 9. IMPLEMENTACIÓN ==============================
h1('9. Implementación de la herramienta web')

h2('9.1. Arquitectura')
cuerpo('La herramienta es una aplicación web de página única, escrita con tecnologías '
       'estándar del lado del cliente: HTML5 para la estructura, CSS3 para la presentación y '
       'JavaScript para el cálculo. Se descartaron las dependencias externas y los '
       'componentes del lado del servidor. La decisión tiene tres razones. Los cálculos del '
       'modelo son deterministas y corren sin problema en el navegador. El despliegue como '
       'sitio estático no cuesta infraestructura. Y un proyecto sin dependencias sigue '
       'funcionando dentro de varios años, cosa que no puede darse por sentada cuando se '
       'arrastra media docena de librerías.')

h2('9.2. Organización de la interfaz')
cuerpo('La interfaz tiene cuatro secciones navegables. La primera es la calculadora; las '
       'otras tres documentan los modelos lógico, algebraico y funcional. Así la herramienta '
       'sirve para dos cosas distintas: calcular, si quien entra es un productor, y verificar '
       'el sustento matemático, si quien entra es un evaluador.')

h2('9.3. Parametrización')
cuerpo('La calculadora recibe cuatro parámetros: provincia de origen, entre ocho provincias '
       'cacaoteras; cantidad de carga, en seis unidades distintas; país de destino, entre '
       'nueve mercados; y modalidad de transporte. El trayecto se calcula en dos tramos, el '
       'acarreo interno hasta el nodo de salida y el tramo internacional. La modalidad '
       'terrestre se deshabilita sola cuando el destino no tiene conexión vial, que es lógica '
       'condicional aplicada al comportamiento de la propia interfaz.')

h2('9.4. Validación de resultados')
cuerpo('Los valores de la herramienta se contrastaron contra los cálculos analíticos. Para el '
       'caso base —25 toneladas de Guayaquil a Rotterdam por vía marítima— la herramienta '
       'devuelve un costo logístico de USD 6 452 y un tránsito de 21,5 días. El primero '
       'coincide con el resultado analítico de USD 6 450; la diferencia viene del redondeo de '
       'la tarifa unitaria. El segundo coincide con el tiempo de tránsito documentado para '
       'esa ruta.')

# ============================== 10. RESULTADOS ==============================
h1('10. Resultados')

cuerpo('El modelo integrado produce una recomendación en tres niveles de decisión, que la '
       'Tabla 6 sintetiza.')

tabla('Tabla 6. Síntesis de resultados del modelo integrado.',
      ['Decisión', 'Herramienta matemática', 'Resultado'],
      [
          ['¿Cuándo usar ruta rápida?', 'Lógica proposicional', 'Solo si D = p ∧ ¬q es verdadera'],
          ['¿Cómo repartir el presupuesto?', 'Sistema de ecuaciones 2×2', 'T = USD 2 800; S = USD 700'],
          ['¿Cuánto exportar al mes?', 'Función cuadrática', '5 contenedores; USD 35 250'],
      ],
      anchos=[1.9, 1.9, 2.2])

cuerpo('Las raíces de la función delimitan el intervalo rentable: de uno a ocho contenedores '
       'al mes. Por debajo no se cubren los costos fijos. Por encima, el encarecimiento del '
       'acopio se come el margen. El óptimo de cinco contenedores equivale a 125 toneladas '
       'mensuales, un 2,5 % de lo que Esmeraldas produce al mes. La materia prima no limita '
       'el modelo.')

# ============================== 11. CONCLUSIONES ==============================
h1('11. Conclusiones y trabajo futuro')

cuerpo('Tres áreas de la matemática de primer nivel —lógica proposicional, sistemas de '
       'ecuaciones lineales y funciones cuadráticas— pueden articularse sobre un mismo caso '
       'real y producir una recomendación operativa coherente. Ahí está la contribución del '
       'trabajo, y no en la complejidad de las técnicas, que son elementales. Cada '
       'herramienta responde una pregunta distinta y las tres respuestas se encadenan en una '
       'sola decisión de negocio.')

cuerpo('La formalización lógica redujo una decisión que suele tomarse por olfato a una regla '
       'verificable con tabla de verdad y validada por inferencia. El sistema de ecuaciones '
       'dio un reparto presupuestario que se comprueba analíticamente. Y la función de '
       'utilidad mostró algo que va contra la intuición del sector: exportar más no siempre '
       'conviene.')

cuerpo('Vale la pena separar el dominio matemático del dominio con sentido real. El vértice '
       'teórico cae en 4,71 contenedores y no significa nada en la práctica; evaluar los '
       'enteros vecinos es el procedimiento que devuelve una recomendación aplicable.')

cuerpo('La limitación principal es conocida: el parámetro de encarecimiento del acopio no se '
       'verificó empíricamente. El análisis de sensibilidad muestra que la estructura de la '
       'conclusión resiste variaciones de ese parámetro, aunque la magnitud del óptimo sí '
       'depende de su valor.')

h2('11.1. Trabajo futuro')
vinieta('Medición empírica del costo marginal de acopio mediante levantamiento de información '
        'con asociaciones cacaoteras de la provincia, a fin de sustituir el supuesto por un '
        'valor documentado.')
vinieta('Incorporación de la estacionalidad del precio internacional del cacao, modelando el '
        'precio de venta como variable dependiente del período en lugar de constante.')
vinieta('Extensión del modelo lógico mediante la inclusión de proposiciones adicionales '
        'relativas a disponibilidad de cupo naviero y certificaciones de calidad.')
vinieta('Integración de la variable r —peso de la carga— en una regla de decisión propia, '
        'considerando umbrales de consolidación de carga.')
vinieta('Validación de la herramienta con usuarios reales del sector, evaluando su usabilidad '
        'y la pertinencia de las unidades y terminología empleadas.')

# ============================== REFERENCIAS ==============================
doc.add_page_break()
h1('Referencias')

referencias = [
    'Anecacao. (2026). Cacao del Ecuador: exportadores, calidad y comercio internacional. '
    'Asociación Nacional de Exportadores de Cacao del Ecuador. https://anecacao.com/',

    'Cámara de Comercio Internacional. (2019). Incoterms 2020: reglas de la ICC para el uso '
    'de términos comerciales nacionales e internacionales. ICC Publications.',

    'Cámara Marítima del Ecuador. (2025). El cacao ecuatoriano, líder en exportaciones 2025. '
    'CAMAE. https://www.camae.org/sin-categoria/el-cacao-ecuatoriano-lider-en-exportaciones-2025/',

    'Camtom. (2026). Arancel de importación de granos de cacao por país de origen. '
    'https://www.camtomx.com/en/aranceles/granos-cacao',

    'Drewry. (2026). World Container Index. Drewry Shipping Consultants.',

    'El Universo. (2026, febrero). Eliminación por decreto del arancel del 15 % está vigente; '
    'en el cacao solo aplica para el grano. '
    'https://www.eluniverso.com/noticias/economia/arancel-sobretasa-15-cacao-comercio-exterior-'
    'estados-unidos-ecuador-nota/',

    'Expreso. (2026a, mayo). El cacao vuelve a superar los $4.000: en Ecuador se pagará hasta '
    '$151 por quintal este miércoles. '
    'https://www.expreso.ec/economia-y-negocios/cacao-vuelve-superar-4-000-ecuador-pagara-151-'
    'quintal-miercoles-283434.html',

    'Expreso. (2026b, mayo). Cacao a precios altos, pero ¿cuánto de esos $151 por quintal '
    'llega realmente al campesino? '
    'https://www.expreso.ec/economia-y-negocios/cacao-precios-altos-cuanto-151-quintal-llega-'
    'realmente-campesino-283652.html',

    'Extra. (2026, enero). Precio del cacao cae en 2026: tonelada baja a $4.500 y preocupa al '
    'sector exportador en Ecuador. '
    'https://www.extra.ec/noticia/economia/precio-cacao-cae-2026-tonelada-baja-4-500-preocupa-'
    'sector-exportador-ecuador-145525.html',

    'Infobae. (2026, febrero 17). Flores, banano, cacao, café y otros productos ecuatorianos '
    'podrán entrar a Estados Unidos sin la tasa del 15 %. '
    'https://www.infobae.com/america/america-latina/2026/02/17/flores-banano-cacao-cafe-y-otros-'
    'productos-ecuatorianos-podran-entrar-a-estados-unidos-sin-la-tasa-del-15/',

    'Logintec. (2026). Seguro de carga internacional: guía 2026. '
    'https://logintec.co/blog/seguro-carga-internacional',

    'Ministerio de Agricultura y Ganadería. (2026). Acuerdo comercial Ecuador – Unión Europea. '
    'Gobierno del Ecuador. https://www.agricultura.gob.ec/acuerdo-comercial-ecuador-union-europea-2/',

    'Primicias. (2025). Cacao de Ecuador: exportaciones crecen 82 % y el sector proyecta '
    '600.000 toneladas en 2025. '
    'https://www.primicias.ec/economia/cacao-ecuador-exportaciones-banco-central-amecacao-feria-'
    'chokao2025-103767/',

    'SeaRates. (2026). Distance and transit time calculator. '
    'https://www.searates.com/distance-time/',

    'Transnatur. (2024). Cómo calcular el seguro de transporte internacional. '
    'https://www.transnatur.com/actualidad/como-calcular-seguro-transporte-internacional/',
]

for ref in sorted(referencias):
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.LEFT
    par.paragraph_format.left_indent = Inches(0.5)
    par.paragraph_format.first_line_indent = Inches(-0.5)   # sangría francesa APA 7
    par.paragraph_format.space_after = Pt(10)
    par.paragraph_format.line_spacing = 1.5
    r = par.add_run(ref)
    r.font.size = Pt(10.5)


# ============================== ANEXOS ==============================
doc.add_page_break()
h1('Anexos')

cuerpo('Se incluyen a continuación los datos de referencia que sustentan los cálculos '
       'presentados en el cuerpo del documento. Su finalidad es permitir la verificación '
       'independiente de los resultados y la reproducción del modelo por parte de terceros.',
       first=0)

# ---------------------------------------------------------------- Anexo A
h2('Anexo A. Evaluación completa de la función de utilidad')

cuerpo('La Tabla A1 presenta el valor de U(x) para cada número entero de contenedores dentro '
       'del rango analizado. Los valores confirman que el óptimo operativo se ubica en cinco '
       'contenedores y que la operación deja de ser rentable a partir del noveno.', first=0)

tabla('Tabla A1. Utilidad mensual según el número de contenedores exportados.',
      ['x (contenedores)', 'Toneladas', 'U(x) en USD', 'Situación'],
      [
          ['0', '0', '−20 000', 'Pérdida: solo costos fijos'],
          ['1', '25', '1 050', 'Rentable'],
          ['2', '50', '17 100', 'Rentable'],
          ['3', '75', '28 150', 'Rentable'],
          ['4', '100', '34 200', 'Rentable'],
          ['5', '125', '35 250', 'Óptimo operativo'],
          ['6', '150', '31 300', 'Rentable'],
          ['7', '175', '22 350', 'Rentable'],
          ['8', '200', '8 400', 'Rentable'],
          ['9', '225', '−10 550', 'Pérdida'],
      ],
      anchos=[1.5, 1.2, 1.5, 2.3], fuente=9.5)

# ---------------------------------------------------------------- Anexo B
h2('Anexo B. Precio de compra según volumen')

cuerpo('La función P(x) = 4 700 + 100x determina el precio pagado al productor. La Tabla B1 '
       'muestra cómo se estrecha el margen unitario a medida que crece el volumen. El margen '
       'se anula en x = 12, punto en el cual el precio de compra iguala al de venta.', first=0)

tabla('Tabla B1. Precio de compra y margen unitario según volumen.',
      ['x (contenedores)', 'P(x) en USD/t', 'Precio de venta', 'Margen por tonelada'],
      [
          ['0', '4 700', '5 900', '1 200'],
          ['1', '4 800', '5 900', '1 100'],
          ['2', '4 900', '5 900', '1 000'],
          ['5', '5 200', '5 900', '700'],
          ['8', '5 500', '5 900', '400'],
          ['10', '5 700', '5 900', '200'],
          ['12', '5 900', '5 900', '0'],
      ],
      anchos=[1.5, 1.5, 1.5, 2.0], fuente=9.5)

# ---------------------------------------------------------------- Anexo C
h2('Anexo C. Parámetros de la herramienta')

cuerpo('Las tablas C1 a C4 reúnen los parámetros con los que opera la calculadora. Las '
       'distancias marítimas corresponden a rutas aproximadas vía Canal de Panamá; las aéreas '
       'siguen trayectoria directa; y las terrestres emplean la red vial estatal ecuatoriana.',
       first=0)

tabla('Tabla C1. Destinos internacionales.',
      ['Código', 'Puerto y país', 'Ruta marítima', 'Ruta aérea', 'Arancel'],
      [
          ['RTM', 'Rotterdam, Países Bajos', '10 700 km', '9 700 km', '0 %'],
          ['HAM', 'Hamburgo, Alemania', '11 200 km', '9 900 km', '0 %'],
          ['ANR', 'Amberes, Bélgica', '10 750 km', '9 650 km', '0 %'],
          ['BCN', 'Barcelona, España', '10 300 km', '9 100 km', '0 %'],
          ['NYC', 'Nueva York, Estados Unidos', '5 850 km', '4 300 km', '0 %'],
          ['ZLO', 'Manzanillo, México', '3 900 km', '3 100 km', '20 %'],
          ['PKG', 'Port Klang, Malasia', '17 800 km', '19 000 km', '0 %'],
          ['BUN', 'Buenaventura, Colombia', '640 km', '600 km', '0 %'],
          ['CLL', 'Callao, Perú', '1 450 km', '1 350 km', '0 %'],
      ],
      anchos=[0.8, 2.1, 1.2, 1.1, 0.9], fuente=9)

tabla('Tabla C2. Provincias de origen y distancias por carretera.',
      ['Provincia', 'Ciudad', 'Al Puerto de Guayaquil', 'Aeropuerto', 'Al aeropuerto'],
      [
          ['Guayas', 'Guayaquil', '0 km', 'GYE', '15 km'],
          ['Manabí', 'Portoviejo', '190 km', 'GYE', '190 km'],
          ['Esmeraldas', 'Esmeraldas', '470 km', 'UIO', '320 km'],
          ['Los Ríos', 'Babahoyo', '60 km', 'GYE', '70 km'],
          ['Santo Domingo', 'Santo Domingo', '290 km', 'UIO', '150 km'],
          ['El Oro', 'Machala', '180 km', 'GYE', '190 km'],
          ['Sucumbíos', 'Lago Agrio', '600 km', 'UIO', '250 km'],
          ['Orellana', 'El Coca', '560 km', 'UIO', '300 km'],
      ],
      anchos=[1.3, 1.2, 1.7, 1.0, 1.0], fuente=9)

tabla('Tabla C3. Modalidades de transporte.',
      ['Modalidad', 'Tarifa (USD por kg-km)', 'Velocidad', 'Días fijos'],
      [
          ['Marítima', '0,0000105', '650 km/día', '5'],
          ['Aérea', '0,00033', '8 000 km/día', '1'],
          ['Terrestre', '0,00006', '500 km/día', '1'],
      ],
      anchos=[1.4, 2.0, 1.4, 1.2], fuente=9.5)

cuerpo('Las tarifas se calibraron contra los valores de mercado citados en el cuerpo del '
       'documento. La tarifa marítima reproduce el flete de USD 2 800 para 25 toneladas con '
       'destino a Rotterdam, y la terrestre reproduce el costo del tramo Esmeraldas–Guayaquil.',
       first=0)

tabla('Tabla C4. Unidades admitidas y factores de conversión.',
      ['Unidad', 'Equivalencia en kilogramos'],
      [
          ['Quintal (100 libras)', '45,36'],
          ['Saco', '60'],
          ['Kilogramo', '1'],
          ['Tonelada métrica', '1 000'],
          ['Libra', '0,45359237'],
          ['Tonelada corta', '907,18474'],
      ],
      anchos=[2.6, 2.6], fuente=9.5)

# ---------------------------------------------------------------- Anexo D
h2('Anexo D. Constantes del modelo')

tabla('Tabla D1. Constantes empleadas en los cálculos.',
      ['Constante', 'Valor', 'Origen'],
      [
          ['Precio de exportación', 'USD 5 900 / t', 'Anecacao (2026)'],
          ['Precio base en finca', 'USD 4 700 / t', '80 % del precio de exportación'],
          ['Incremento por contenedor', 'USD 100 / t', 'Supuesto de modelado'],
          ['Capacidad del contenedor', '25 000 kg', 'Contenedor estándar de 40 pies'],
          ['Tasa de seguro', '0,47 % del valor', 'Resultado del sistema 2×2'],
          ['Tasa de comisiones', '2 % del valor', 'Promedio de mercado'],
          ['Costos fijos mensuales', 'USD 20 000', 'Administración y certificación'],
      ],
      anchos=[2.1, 1.5, 2.4], fuente=9.5)

cuerpo('De los siete valores anteriores, únicamente el incremento por contenedor carece de '
       'respaldo documental. Su efecto sobre la conclusión se examina en la sección 7.4.',
       first=0)

# ---------------------------------------------------------------- Anexo E
h2('Anexo E. Acceso a la herramienta')

cuerpo('La aplicación se distribuye como sitio estático y su código fuente está disponible '
       'para revisión. No requiere instalación ni servicios externos: basta abrir el archivo '
       'index.html en cualquier navegador moderno.', first=0)

tabla('Tabla E1. Componentes del proyecto.',
      ['Archivo', 'Contenido'],
      [
          ['index.html', 'Estructura de las cuatro secciones y de la calculadora'],
          ['style.css', 'Sistema visual: paleta, tipografía y diseño responsivo'],
          ['script.js', 'Lógica de cálculo, parámetros y generación de gráficas'],
      ],
      anchos=[1.6, 3.9], fuente=9.5)

cuerpo('Repositorio del proyecto: https://github.com/Renngio45-gif/'
       'Logistica_Internacional_Proyecto_1S_LM_GL', first=0)

cuerpo('Ruta de la aplicación dentro del repositorio: proyecto_logica_matemacica/sitio_web/',
       first=0)


salida = os.path.join(BASE, 'Paper_Pepa_de_Oro_PUCESE.docx')
doc.save(salida)
print('Creado:', salida)
