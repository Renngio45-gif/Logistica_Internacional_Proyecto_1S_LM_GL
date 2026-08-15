# -*- coding: utf-8 -*-
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_background(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tc_pr.append(shd)

def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tc_mar.append(node)
    tc_pr.append(tc_mar)

def set_table_borders(table, color="CCCCCC", sz="4", val="single"):
    tbl_pr = table._tbl.tblPr
    tbl_borders = OxmlElement('w:tblBorders')
    for b_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{b_name}')
        border.set(qn('w:val'), val)
        border.set(qn('w:sz'), sz)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tbl_borders.append(border)
    tbl_pr.append(tbl_borders)

def main():
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    COLOR_PRIMARY = RGBColor(27, 54, 93)
    COLOR_SECONDARY = RGBColor(70, 130, 180)
    COLOR_TEXT = RGBColor(51, 51, 51)
    COLOR_GREEN = RGBColor(34, 139, 34)
    HEX_PRIMARY = '1B365D'
    HEX_LIGHT_GREY = 'F4F6F9'

    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Arial'
    style_normal.font.size = Pt(11)
    style_normal.font.color.rgb = COLOR_TEXT
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(6)

    def add_p(text="", space_after=6, line_spacing=1.15, bold=False, italic=False, color=None, align=WD_ALIGN_PARAGRAPH.LEFT):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        p.alignment = align
        if text:
            run = p.add_run(text)
            run.bold = bold
            run.italic = italic
            if color:
                run.font.color.rgb = color
        return p

    def add_heading(num_title, size=12.5):
        p = add_p(num_title, bold=True, color=COLOR_PRIMARY, space_after=10)
        p.runs[0].font.size = Pt(size)
        return p

    def add_bullet(label, text, indent=0.25, space_after=6):
        p = add_p(space_after=space_after)
        p.paragraph_format.left_indent = Inches(indent)
        r = p.add_run("•  " + label)
        r.bold = True
        p.add_run(text)
        return p

    def styled_table(headers, rows_data, widths=None):
        table = doc.add_table(rows=1, cols=len(headers))
        table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(table)
        hdr_cells = table.rows[0].cells
        for i, title in enumerate(headers):
            hdr_cells[i].text = ""
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(2)
            run = p.add_run(title)
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(10)
            set_cell_background(hdr_cells[i], HEX_PRIMARY)
            set_cell_margins(hdr_cells[i], top=160, bottom=160, left=120, right=120)
        for row_idx, data in enumerate(rows_data):
            row = table.add_row()
            cells = row.cells
            bg_color = HEX_LIGHT_GREY if row_idx % 2 == 1 else "FFFFFF"
            for col_idx, text in enumerate(data):
                cells[col_idx].text = ""
                p = cells[col_idx].paragraphs[0]
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.space_before = Pt(2)
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx == 0 or col_idx == len(data) - 1 else WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(text)
                run.font.size = Pt(10)
                set_cell_background(cells[col_idx], bg_color)
                set_cell_margins(cells[col_idx], top=120, bottom=120, left=120, right=120)
        return table

    # ================= PORTADA =================
    p_uni = add_p("PONTIFICIA UNIVERSIDAD CATÓLICA DEL ECUADOR", bold=True, color=COLOR_PRIMARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    p_uni.runs[0].font.size = Pt(16)
    p_uni.paragraph_format.space_before = Pt(36)

    p_sede = add_p("SEDE ESMERALDAS", bold=True, color=COLOR_PRIMARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    p_sede.runs[0].font.size = Pt(14)

    p_pucetec = add_p("PUCETEC", bold=True, color=COLOR_SECONDARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    p_pucetec.runs[0].font.size = Pt(11)
    p_tech = add_p("TECNOLOGÍA EN DESARROLLO DE SOFTWARE", bold=True, color=COLOR_SECONDARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=48)
    p_tech.runs[0].font.size = Pt(11)

    p_title_cov = add_p("PROYECTO INTEGRADOR: FASE 2 (MODELO ALGEBRAICO)", bold=True, color=COLOR_PRIMARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    p_title_cov.runs[0].font.size = Pt(16)
    p_title_cov.paragraph_format.space_before = Pt(48)

    p_topic_cov = add_p("TEMA:\nPresupuesto y logística de un envío internacional\n(Caso: Exportación de Cacao desde Esmeraldas)\nIncluye: Producto 2 — Parte C (Sistema de ecuaciones 2×2)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=48)
    p_topic_cov.runs[0].font.size = Pt(12)

    p_meta = add_p(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    p_meta.paragraph_format.space_before = Pt(48)
    p_meta.paragraph_format.line_spacing = 1.3
    r_int = p_meta.add_run("Integrantes:\n"); r_int.bold = True
    p_meta.add_run("Alan Cárdenas Morán\nGeovanny Farías Estupiñán\n\n")
    r_asig = p_meta.add_run("Asignatura:\n"); r_asig.bold = True
    p_meta.add_run("Habilidades Lógico-Matemáticas\n\n")
    r_doc = p_meta.add_run("Docente:\n"); r_doc.bold = True
    p_meta.add_run("Msc. Adrián Vargas\n\n")
    r_per = p_meta.add_run("Nivel y Período:\n"); r_per.bold = True
    p_meta.add_run("Primer Nivel — Período 2026-1")

    p_date = add_p("Esmeraldas - Ecuador\n2026", bold=True, color=COLOR_SECONDARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    p_date.runs[0].font.size = Pt(11)
    p_date.paragraph_format.space_before = Pt(60)

    doc.add_page_break()

    # ================= TITULO =================
    p_title = add_p("REPORTE DE PROYECTO INTEGRADOR: FASE 2 (MODELO ALGEBRAICO)", bold=True, color=COLOR_PRIMARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)
    p_title.runs[0].font.size = Pt(14)

    # ---------- 1. DATOS GENERALES ----------
    add_heading("1. DATOS GENERALES")
    add_bullet("Grupo: ", "2")
    add_bullet("Integrantes: ", "Alan Cárdenas Morán, Geovanny Farías Estupiñán")
    add_bullet("Idea de proyecto elegida: ", "7. Presupuesto y logística de un envío internacional (Caso: Exportación de Cacao)", space_after=20)

    # ---------- 2. INFERENCIA LOGICA ----------
    add_heading("2. INFERENCIA LÓGICA VÁLIDA (COMPLEMENTO DE LA FASE 1)")
    add_p("Para validar que la regla de decisión de la Fase 1 permite concluir correctamente la aprobación de una ruta, se aplica la regla de inferencia Modus Ponens: si se afirma un condicional y se afirma su antecedente, se concluye necesariamente su consecuente.", space_after=10)

    styled_table(
        ["Elemento", "Fórmula", "Lectura en el sistema"],
        [
            ["Premisa 1 (condicional)", "(p ∧ ¬q) → D", "Si el envío es urgente y el presupuesto no está excedido, entonces se aprueba la ruta rápida directa."],
            ["Premisa 2 (antecedente)", "p ∧ ¬q", "El lote de cacao es urgente (contrato con penalización) y su costo logístico no supera el techo presupuestario."],
            ["Conclusión (∴)", "D", "Se aprueba la ruta rápida directa para el envío del cacao."],
        ]
    )
    add_p("", space_after=8)
    add_p("La inferencia es válida porque su forma lógica [(A → B) ∧ A] → B es una tautología: en toda fila de su tabla de verdad el resultado es V, por lo que la conclusión nunca puede ser falsa si las premisas son verdaderas.", space_after=20)

    # ---------- 3. DATOS REALES ----------
    add_heading("3. DATOS REALES UTILIZADOS EN EL MODELO")
    add_p("El modelo se construye con datos reales del mercado cacaotero ecuatoriano (julio 2026) y de la logística marítima internacional:", space_after=10)

    styled_table(
        ["Dato", "Valor real", "Fuente"],
        [
            ["Precio internacional del cacao (jul. 2026)", "≈ $5,900 / t", "Anecacao / Ecuavisa (10-jul-2026)"],
            ["Producción de cacao en Esmeraldas", "58,965 t en 90,248 ha", "Ministerio de Agricultura y Ganadería"],
            ["Exportación nacional de cacao (2025)", "≈ 600,000 t ($4,668 millones)", "Banco Central del Ecuador / Anecacao"],
            ["Flete marítimo contenedor 40 pies (promedio global)", "≈ $2,286", "Drewry World Container Index (2026)"],
            ["Seguro de carga internacional", "0.3% – 1.5% del valor CIF", "Logintec / TransNatur"],
            ["Distancia marítima Guayaquil → Rotterdam (vía Canal de Panamá)", "≈ 10,700 km", "SeaRates / sea-distances.org"],
            ["Arancel del cacao ecuatoriano en la Unión Europea", "0% (Acuerdo Multipartes)", "Ministerio de Agricultura / EEAS-UE"],
        ]
    )
    add_p("", space_after=8)
    p_lote = add_p(space_after=20)
    r = p_lote.add_run("Lote modelado: ")
    r.bold = True
    p_lote.add_run("un contenedor de 40 pies con 25,000 kg (25 t) de cacao fino de aroma. Valor FOB = 25 t × $5,900/t = $147,500.")

    # ---------- 4. FORMULAS DEL SISTEMA ----------
    add_heading("4. FÓRMULAS DEL SISTEMA")
    add_bullet("Costo logístico total: ", "C = T + S + (a/100)·FOB, donde T = costo del transporte, S = costo del seguro y a = comisión logística (%) aplicada al valor FOB.")
    add_bullet("Aranceles / tasas: ", "A = (t/100)·FOB, donde t es el porcentaje de impuestos o tasas del mercado de destino (UE: t = 0%).")
    add_bullet("Conversión de distancia: ", "millas = kilómetros × 0.621371")
    add_bullet("Conversión de volumen: ", "galones = litros × 0.264172")
    add_bullet("Conversión de masa: ", "libras = kilogramos × 2.20462", space_after=20)

    # ---------- 5. SISTEMA 2x2 (PARTE C) ----------
    add_heading("5. SISTEMA DE ECUACIONES 2×2 — PRODUCTO 2, PARTE C")

    p_v = add_p(space_after=4)
    r = p_v.add_run("Variables del sistema:"); r.bold = True
    add_bullet("T: ", "costo del transporte marítimo del lote (dólares).")
    add_bullet("S: ", "costo del seguro de la carga (dólares).", space_after=6)
    p_nom = add_p(space_after=10)
    rn = p_nom.add_run("Nota de nomenclatura: ")
    rn.bold = True
    p_nom.add_run("las incógnitas se nombran T y S, y no x e y, para no confundirlas "
                  "con la variable x del análisis de funciones (Producto 2/3), que "
                  "cuenta contenedores. Son magnitudes distintas y por eso llevan "
                  "nombres distintos.")

    p_c = add_p(space_after=4)
    r = p_c.add_run("Condiciones reales de la operación:"); r.bold = True
    add_bullet("Condición 1: ", "el presupuesto logístico conjunto asignado al lote (transporte + seguro) es de $3,500.")
    add_bullet("Condición 2: ", "según los datos de mercado (flete ≈ $2,286–2,900 frente a un seguro de ≈ 0.5% del FOB), el transporte cuesta el cuádruple del seguro.", space_after=10)

    p_sys = add_p(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    r_eq = p_sys.add_run("Ecuación 1:   T + S = 3500\nEcuación 2:   T = 4S")
    r_eq.bold = True
    r_eq.font.size = Pt(12.5)
    r_eq.font.color.rgb = COLOR_PRIMARY

    p_met = add_p(space_after=4)
    r = p_met.add_run("Resolución por el método de sustitución:"); r.bold = True
    add_bullet("Paso 1: ", "sustituir la Ecuación 2 en la Ecuación 1:  4S + S = 3500")
    add_bullet("Paso 2: ", "reducir términos semejantes:  5S = 3500")
    add_bullet("Paso 3: ", "despejar S:  S = 3500 / 5 = 700")
    add_bullet("Paso 4: ", "reemplazar en la Ecuación 2:  T = 4(700) = 2800", space_after=10)

    p_sol = add_p(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    r_sol = p_sol.add_run("Solución:   T = $2,800 (transporte)   ;   S = $700 (seguro)")
    r_sol.bold = True
    r_sol.font.size = Pt(12)
    r_sol.font.color.rgb = COLOR_GREEN

    p_comp = add_p(space_after=4)
    r = p_comp.add_run("Comprobación (sustituyendo en ambas ecuaciones originales):"); r.bold = True
    add_bullet("Ecuación 1: ", "2800 + 700 = 3500 ✓")
    add_bullet("Ecuación 2: ", "2800 = 4 × 700 = 2800 ✓", space_after=10)

    p_val = add_p(space_after=10)
    r = p_val.add_run("Validación con datos reales: ")
    r.bold = True
    p_val.add_run("T = $2,800 es coherente con el flete promedio de un contenedor de 40 pies (Drewry: ≈ $2,286 promedio global; las rutas específicas suelen superar el promedio), y S = $700 equivale al 0.47% del valor FOB ($147,500), dentro del rango real de 0.3%–1.5% del seguro de carga internacional.")

    p_dec = add_p(space_after=8)
    r = p_dec.add_run("Decisión técnica que permite tomar este resultado (una línea): ")
    r.bold = True
    p_dec.add_run("como el costo logístico total ($3,500) no excede el techo presupuestario del lote, ¬q es verdadera; siendo el envío urgente (p), por Modus Ponens se concluye D: se aprueba la ruta rápida directa.")

    p_evid = add_p(space_after=20)
    r = p_evid.add_run("Evidencia gráfica: ")
    r.bold = True
    p_evid.add_run("al graficar ambas rectas (T + S = 3500 y T = 4S) en GeoGebra/Desmos o en el simulador web del proyecto, las rectas se intersecan exactamente en el punto (S, T) = (700, 2800), que es la solución del sistema.")

    # ---------- 6. CONVERSIONES ----------
    add_heading("6. CONVERSIONES DE UNIDADES DEL LOTE (SI ↔ IMPERIAL)")
    add_p("Aplicadas a los datos reales del envío modelado:", space_after=10)

    styled_table(
        ["Magnitud", "Dato (SI)", "Procedimiento", "Resultado (Imperial)"],
        [
            ["Distancia de la ruta", "10,700 km", "10,700 × 0.621371", "6,648.67 mi"],
            ["Peso de la carga", "25,000 kg", "25,000 × 2.20462", "55,115.50 lb"],
            ["Combustible terrestre", "188 L", "188 × 0.264172", "49.66 gal"],
        ]
    )
    add_p("", space_after=8)
    p_prop = add_p(space_after=4)
    r = p_prop.add_run("Proporcionalidad: ")
    r.bold = True
    p_prop.add_run("el combustible del tramo terrestre Esmeraldas → Puerto de Guayaquil (≈ 470 km) se calcula por proporcionalidad directa con el consumo del camión (40 L por cada 100 km): 470 × 40/100 = 188 L.")

    p_porc = add_p(space_after=20)
    r = p_porc.add_run("Porcentajes: ")
    r.bold = True
    p_porc.add_run("el seguro representa 700/147,500 = 0.47% del valor FOB; la comisión logística del 2% del FOB equivale a $2,950; con ellos el costo logístico total es C = 2,800 + 700 + 2,950 = $6,450. En la UE el arancel es 0% gracias al Acuerdo Multipartes.")

    # ---------- 7. HERRAMIENTAS ----------
    add_heading("7. HERRAMIENTAS UTILIZADAS")
    add_bullet("Simulador web del proyecto: ", "resuelve el sistema 2×2 en vivo, grafica las dos rectas con su punto de intersección y calcula el costo total y las conversiones.")
    add_bullet("GeoGebra / Desmos: ", "verificación gráfica del sistema (las rectas T + S = 3500 y T = 4S se cruzan en (700, 2800)).")
    add_bullet("Hoja de cálculo: ", "comprobación numérica de conversiones y porcentajes.", space_after=20)

    # ---------- 8. CONCLUSIONES ----------
    add_heading("8. CONCLUSIONES")
    add_bullet("", "El sistema 2×2 permitió repartir el presupuesto logístico real del lote entre transporte ($2,800) y seguro ($700), con una solución comprobada algebraica y gráficamente.")
    add_bullet("", "Los valores obtenidos son consistentes con los datos reales del mercado (flete Drewry y rango del seguro de carga), lo que valida el modelo.")
    add_bullet("", "La regla lógica D = p ∧ ¬q y la inferencia Modus Ponens conectan el modelo algebraico con la decisión técnica final: aprobar la ruta rápida directa del envío de cacao.")

    filename = "Reporte_Fase_2_Modelo_Completo.docx"
    doc.save(filename)
    print(f"Document successfully created: {filename}")

if __name__ == "__main__":
    main()
