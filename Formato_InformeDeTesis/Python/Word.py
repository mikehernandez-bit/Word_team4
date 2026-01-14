from docx import Document
from docx.shared import Cm, Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import platform
import subprocess

def configurar_formato_unac(doc):
    """Configura A4, Márgenes UNAC (3.5 izq) y estilo Arial 12"""
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(3.5)
        section.right_margin = Cm(2.5)
        section.top_margin = Cm(3.0)
        section.bottom_margin = Cm(3.0)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    # Forzar Arial en Word
    rFonts = style.element.rPr.rFonts
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')
    
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

def agregar_bloque(doc, texto, negrita=False, tamano=12, antes=0, despues=0, cursiva=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE # Carátula con espacio simple para control total
    p.paragraph_format.space_before = Pt(antes)
    p.paragraph_format.space_after = Pt(despues)
    run = p.add_run(texto)
    run.bold = negrita
    run.italic = cursiva
    run.font.size = Pt(tamano)
    return p

def crear_caratula_elegante(doc):
    # 1. ENCABEZADO (Impacto Institucional)
    agregar_bloque(doc, "UNIVERSIDAD NACIONAL DEL CALLAO", negrita=True, tamano=18, despues=4)
    agregar_bloque(doc, "FACULTAD DE [NOMBRE DE LA FACULTAD]", negrita=True, tamano=14, despues=4)
    agregar_bloque(doc, "ESCUELA PROFESIONAL DE [NOMBRE DE LA ESCUELA]", negrita=True, tamano=14, despues=25)

    # 2. LOGO (GRANDE Y CENTRADO)
    ruta_script = os.path.dirname(__file__)
    ruta_logo = os.path.join(ruta_script, "..", "Imagenes", "LogoUNAC.png")
    
    if os.path.exists(ruta_logo):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = p_logo.add_run()
        # Aumentado a 3.2 pulgadas para máxima elegancia
        run_logo.add_picture(ruta_logo, width=Inches(3.2))
    else:
        agregar_bloque(doc, "[LOGO INSTITUCIONAL]", tamano=10, antes=40, despues=40)

    # 3. TÍTULO DEL DOCUMENTO
    agregar_bloque(doc, "INFORME DE TESIS", negrita=True, tamano=16, antes=30)
    
    # El título en mayúsculas, negrita y con espacio generoso
    titulo_placeholder = '"[ESCRIBA AQUÍ EL TÍTULO DE LA TESIS EN MAYÚSCULAS Y ENTRE COMILLAS]"'
    agregar_bloque(doc, titulo_placeholder, negrita=True, tamano=14, antes=30, despues=30)

    # 4. GRADO ACADÉMICO
    agregar_bloque(doc, "PARA OPTAR EL TÍTULO PROFESIONAL DE:", tamano=12, antes=10)
    agregar_bloque(doc, "[INGENIERO DE ...]", negrita=True, tamano=13, despues=35)

    # 5. BLOQUE DE AUTORES (Presentación limpia)
    agregar_bloque(doc, "AUTOR: [NOMBRES Y APELLIDOS]", negrita=True, tamano=12, antes=5)
    agregar_bloque(doc, "ASESOR: [NOMBRES Y APELLIDOS]", negrita=True, tamano=12, antes=5, despues=20)
    
    agregar_bloque(doc, "LÍNEA DE INVESTIGACIÓN: [NOMBRE DE LA LÍNEA]", tamano=11, cursiva=True, despues=40)

    # 6. PIE DE PÁGINA
    agregar_bloque(doc, "Callao, 2026", tamano=12)
    agregar_bloque(doc, "PERÚ", negrita=True, tamano=12)

def agregar_contenido_preliminar(doc):
    """
    Estructura las secciones preliminares y unifica los índices en una sola hoja
    según la normativa UNAC.
    """
    # --- 1. HOJA DE RESPETO (Blanca) ---
    # ELIMINAMOS el doc.add_page_break() que estaba aquí arriba
    doc.add_paragraph() 
    doc.add_page_break() # Este salto de página separa la hoja de respeto de la Dedicatoria

    # Función interna para Títulos Formales (Arial 14, Negrita, Negro)
    def agregar_titulo_formal(texto, espaciado_antes=0):
        h = doc.add_heading(level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = h.add_run(texto)
        run.font.name = 'Arial'
        run.font.size = Pt(14)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        h.paragraph_format.space_before = Pt(espaciado_antes)
        h.paragraph_format.space_after = Pt(12)

    # --- 2. DEDICATORIA / AGRADECIMIENTO ---
    agregar_titulo_formal("DEDICATORIA / AGRADECIMIENTO")
    doc.add_paragraph("[Escriba aquí su dedicatoria o agradecimientos...]").alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.add_page_break()

    # --- 3. RESUMEN / ABSTRACT ---
    agregar_titulo_formal("RESUMEN / ABSTRACT")
    
    # Nota técnica según directiva
    p_nota = doc.add_paragraph()
    p_nota.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_nota = p_nota.add_run("Nota: Síntesis de objetivos, métodos y resultados principales.")
    run_nota.italic = True
    run_nota.font.size = Pt(11)
    
    doc.add_paragraph("\n[Escriba aquí el cuerpo del resumen...]").alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.add_page_break()

    # --- 4. HOJA DE ÍNDICES UNIFICADA ---
    # Título principal de la sección de índices (opcional)
    
    # Índice de Contenido
    agregar_titulo_formal("ÍNDICE DE CONTENIDO")
    p_gen_cont = doc.add_paragraph("(Generarlo)")
    p_gen_cont.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Índice de Tablas (Misma hoja)
    agregar_titulo_formal("ÍNDICE DE TABLAS", espaciado_antes=30)
    p_gen_tab = doc.add_paragraph("(Generarlo)")
    p_gen_tab.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Índice de Figuras (Misma hoja)
    agregar_titulo_formal("ÍNDICE DE FIGURAS", espaciado_antes=30)
    p_gen_fig = doc.add_paragraph("(Generarlo)")
    p_gen_fig.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Índice de Abreviaturas (Misma hoja)
    agregar_titulo_formal("ÍNDICE DE ABREVIATURAS", espaciado_antes=30)
    p_gen_abr = doc.add_paragraph("(Generarlo)")
    p_gen_abr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_page_break()
    
    # --- 5. INTRODUCCIÓN (En hoja nueva) ---
    agregar_titulo_formal("INTRODUCCIÓN")
    
    p_intro = doc.add_paragraph()
    p_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_intro.add_run("[Escriba aquí la introducción de su tesis. La introducción debe presentar de manera general el tema, el propósito de la investigación y la estructura del trabajo documental.]")
    
    doc.add_page_break()

def agregar_cuerpo_informe(doc):
    """
    Agrega los capítulos del I al VI con subtítulos oficiales 
    y notas guía elegantes.
    """
    
    def agregar_titulo_capitulo(texto):
        h = doc.add_heading(level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = h.add_run(texto)
        run.font.name = 'Arial'
        run.font.size = Pt(14)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        h.paragraph_format.space_before = Pt(24)
        h.paragraph_format.space_after = Pt(18)

    def agregar_subtitulo(texto):
        p = doc.add_paragraph()
        run = p.add_run(texto)
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        run.bold = True
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)

    def agregar_nota_guia(texto):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = p.add_run(f"Nota: {texto}")
        run.font.name = 'Arial'
        run.font.size = Pt(10)
        run.italic = True
        # Gris oscuro para elegancia
        run.font.color.rgb = RGBColor(89, 89, 89) 
        p.paragraph_format.space_after = Pt(12)

    # --- CAPÍTULO I: PLANTEAMIENTO DEL PROBLEMA ---
    agregar_titulo_capitulo("I. PLANTEAMIENTO DEL PROBLEMA")
    
    agregar_subtitulo("1.1 Descripción de la realidad problemática")
    agregar_nota_guia("Describa la situación actual del problema a nivel macro, meso y micro.")
    
    agregar_subtitulo("1.2 Formulación del problema")
    agregar_subtitulo("1.3 Objetivos (General y específicos)")
    agregar_subtitulo("1.4 Justificación")
    agregar_subtitulo("1.5 Delimitantes de la investigación")
    doc.add_page_break()

    # --- CAPÍTULO II: MARCO TEÓRICO ---
    agregar_titulo_capitulo("II. MARCO TEÓRICO")
    
    agregar_subtitulo("2.1 Antecedentes (Internacional y nacional)")
    agregar_nota_guia("Incluir tesis y artículos científicos relacionados (últimos 5 años).")
    
    agregar_subtitulo("2.2 Bases teóricas")
    agregar_subtitulo("2.3 Marco conceptual")
    agregar_subtitulo("2.4 Definición de términos básicos")
    doc.add_page_break()

    # --- CAPÍTULO III: METODOLOGÍA ---
    agregar_titulo_capitulo("III. METODOLOGÍA")
    agregar_subtitulo("3.1 Tipo y diseño de investigación")
    agregar_subtitulo("3.2 Método de investigación")
    agregar_subtitulo("3.3 Población y muestra")
    agregar_subtitulo("3.4 Lugar de estudio y periodo")
    agregar_subtitulo("3.5 Técnicas e instrumentos de recolección")
    agregar_subtitulo("3.6 Análisis y procesamiento de datos")
    doc.add_page_break()

    # --- CAPÍTULO IV: RESULTADOS Y DISCUSIÓN ---
    agregar_titulo_capitulo("IV. RESULTADOS Y DISCUSIÓN")
    
    agregar_subtitulo("4.1 Presentación de resultados")
    agregar_nota_guia("Contrastación con estadística descriptiva e inferencial.")
    
    agregar_subtitulo("4.2 Contrastación de hipótesis")
    
    agregar_subtitulo("4.3 Discusión de resultados")
    agregar_nota_guia("Comparación de hallazgos con antecedentes y bases teóricas.")
    doc.add_page_break()

    # --- CAPÍTULO V: CONCLUSIONES ---
    agregar_titulo_capitulo("V. CONCLUSIONES")
    agregar_nota_guia("Mínimo una conclusión por cada objetivo específico.")
    doc.add_page_break()

    # --- CAPÍTULO VI: RECOMENDACIONES ---
    agregar_titulo_capitulo("VI. RECOMENDACIONES")
    agregar_nota_guia("Sugerencias metodológicas, académicas y prácticas.")

def agregar_referencias_y_anexos(doc):
    """
    Agrega las secciones finales del informe: Referencias y Anexos.
    """
    
    def agregar_titulo_final(texto):
        h = doc.add_heading(level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = h.add_run(texto)
        run.font.name = 'Arial'
        run.font.size = Pt(14)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        h.paragraph_format.space_before = Pt(24)
        h.paragraph_format.space_after = Pt(18)

    def agregar_nota_guia(texto):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = p.add_run(f"Nota: {texto}")
        run.font.name = 'Arial'
        run.font.size = Pt(10)
        run.italic = True
        run.font.color.rgb = RGBColor(89, 89, 89)
        p.paragraph_format.space_after = Pt(12)

    # --- VII. REFERENCIAS BIBLIOGRÁFICAS ---
    agregar_titulo_final("VII. REFERENCIAS BIBLIOGRÁFICAS")
    agregar_nota_guia("Utilice gestores como Mendeley o Zotero. Para Ingeniería se recomienda IEEE, para otras facultades APA 7ma edición.")
    doc.add_paragraph("1.\tAPELLIDO, Nombre. \"Título del artículo\". Editorial, Año.\n"
                      "2.\tAPELLIDO, Nombre. \"Título del libro\". Ciudad: Editorial, Año.")
    doc.add_page_break()

    # --- VIII. ANEXOS ---
    agregar_titulo_final("VIII. ANEXOS")
    
    # Anexo 1: Matriz de Consistencia
    p_anexo1 = doc.add_paragraph()
    run_a1 = p_anexo1.add_run("Anexo 1: Matriz de Consistencia")
    run_a1.bold = True
    agregar_nota_guia("La matriz debe resumir todo el proyecto. Columnas: Problemas, Objetivos, Hipótesis, Variables, Metodología.")
    
    # Crear una tabla base para la Matriz de Consistencia
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, txt in enumerate(['Problemas', 'Objetivos', 'Hipótesis', 'Variables', 'Metodología']):
        hdr_cells[i].text = txt
        hdr_cells[i].paragraphs[0].runs[0].bold = True
    
    doc.add_paragraph() # Espacio

    # Anexo 2: Instrumentos
    p_anexo2 = doc.add_paragraph()
    run_a2 = p_anexo2.add_run("Anexo 2: Instrumento de recolección de datos")
    run_a2.bold = True
    agregar_nota_guia("Adjunte aquí el cuestionario, guía de entrevista o ficha técnica de los equipos/sensores utilizados.")
    
    doc.add_paragraph() # Espacio

    # Anexo 3: Validación
    p_anexo3 = doc.add_paragraph()
    run_a3 = p_anexo3.add_run("Anexo 3: Validación de instrumento (Certificado de expertos)")
    run_a3.bold = True
    agregar_nota_guia("Incluya las fichas firmadas por los 3 expertos que validaron su instrumento antes de la aplicación.")

    doc.add_page_break()
    
def agregar_numeracion_paginas(doc):
    """
    Agrega numeración de páginas centrada en el pie de página.
    Nota: La numeración romana vs arábiga avanzada requiere secciones 
    manuales en Word, por lo que aplicaremos la estándar oficial (arábiga).
    """
    for section in doc.sections:
        footer = section.footer
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Inicio del campo de numeración
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = "PAGE"
        
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        
        run = p.add_run()
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
        run.font.name = 'Arial'
        run.font.size = Pt(10) 

def generar_tesis_pro():
    try:
        doc = Document()
        # 1. Configuración de márgenes (3.5cm Izq) y fuente Arial
        configurar_formato_unac(doc)
        
        # 2. Carátula Profesional (Diseño UNAC)
        crear_caratula_elegante(doc)
        
        # 3. Secciones Preliminares e Introducción
        agregar_contenido_preliminar(doc)
        
        # 4. Cuerpo del Informe (Capítulos I al VI)
        agregar_cuerpo_informe(doc)
        
        # 5. Referencias y Anexos (Puntos VII y VIII)
        agregar_referencias_y_anexos(doc)

        # --- NUEVO: AGREGAR NUMERACIÓN DE PÁGINAS ---
        # Se llama antes de guardar para que se aplique a todas las secciones
        agregar_numeracion_paginas(doc)
        
        # --- PROCESO DE GUARDADO SEGURO ---
        nombre = "Estructura_Informe_de_Tesis_Pregrado.docx"
        ruta_final = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", nombre))
        
        try:
            doc.save(ruta_final)
        except PermissionError:
            print("❌ ERROR: El archivo Word ya está abierto. Ciérralo y vuelve a ejecutar.")
            return

        print(f"✅ ¡Estructura de Tesis Generada con éxito!")
        print(f"📍 Ubicación: {ruta_final}")
        
        # --- APERTURA AUTOMÁTICA ---
        if platform.system() == 'Windows':
            os.startfile(ruta_final)
        else:
            cmd = 'open' if platform.system() == 'Darwin' else 'xdg-open'
            subprocess.call((cmd, ruta_final))
            
    except Exception as e:
        print(f"❌ Error crítico en el flujo principal: {e}")

if __name__ == "__main__":
    generar_tesis_pro()