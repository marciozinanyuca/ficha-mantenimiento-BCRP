import uvicorn
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import pandas as pd
import unicodedata
import shutil
import os
import re
import docx
import difflib
import warnings
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

app = FastAPI(title="API Fichas de Mantenimiento BCRP")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = "temp_files"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# --- 0. FUNCIONES AUXILIARES Y DE INTELIGENCIA ---
def normalizar_texto(texto):
    if pd.isna(texto): return ""
    texto = str(texto).strip().upper()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^\w\s\(\)]', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto)
    return texto.strip()

def obtener_maquinas_excel():
    """Lee el Excel y devuelve la lista oficial de máquinas."""
    file_path = os.path.join(UPLOAD_DIR, "matriz_actual.xlsx")
    if not os.path.exists(file_path): return []
    try: df = pd.read_excel(file_path, sheet_name='PLAN.MANT', header=None)
    except: return []

    idx_maquina, fila_inicio = 1, 0
    for r_idx, row in df.head(25).iterrows():
        for c_idx, cell in enumerate(row):
            if 'MAQUINA' in normalizar_texto(cell):
                idx_maquina = c_idx; fila_inicio = r_idx + 1
                break
        if fila_inicio > 0: break

    maquinas = []
    for r_idx in range(fila_inicio, len(df)):
        celda = df.iloc[r_idx, idx_maquina]
        if pd.notna(celda):
            nombre_raw = str(celda).strip()
            nombre_norm = normalizar_texto(nombre_raw)
            if nombre_norm and nombre_norm != "MAQUINA" and not any(p in nombre_norm for p in ["MANTENIMIENTO", "ELABORADO", "FECHA", "AREA"]):
                maquinas.append(nombre_raw)
    return list(dict.fromkeys(maquinas))

def coincidencia_inteligente(hint, lista_maquinas):
    """Cruza un texto cualquiera con el catálogo del Excel usando IA difusa avanzada"""
    hint_norm = normalizar_texto(hint)
    if not hint_norm: return None

    mejor_maq = None
    max_score = 0.0
    
    # Limpiamos los paréntesis para comparar palabras sueltas puras
    hint_words = re.sub(r'[\(\)]', ' ', hint_norm).split()
    if not hint_words: return None
    
    for maq in lista_maquinas:
        maq_norm = normalizar_texto(maq)
        maq_words = re.sub(r'[\(\)]', ' ', maq_norm).split()
        
        if not maq_words: continue
            
        # 1. Coincidencia Exacta (Premio mayor inmediato)
        if hint_norm == maq_norm:
            return maq
            
        score = 0.0
        
        # 2. Si uno está contenido exactamente dentro del otro
        if hint_norm in maq_norm or maq_norm in hint_norm:
            min_len = min(len(hint_norm), len(maq_norm))
            max_len = max(len(hint_norm), len(maq_norm))
            score = max(score, 0.8 + 0.2 * (min_len / max_len))
            
        # 3. Intersección de Palabras Inteligente
        matches = 0
        for hw in hint_words:
            if hw in maq_words:
                matches += 1
            elif not hw.isnumeric(): # No adivinar números o códigos, deben ser exactos
                close = difflib.get_close_matches(hw, maq_words, n=1, cutoff=0.8)
                if close:
                    matches += 1
                    
        # Calculamos qué tanto del texto buscado está en el Excel
        score_hint = matches / len(hint_words)
        # Calculamos qué tanto del Excel fue cubierto
        score_maq = matches / len(maq_words)
        
        # Peso: 75% importa que las palabras que escribiste estén, 25% importa la longitud extra
        word_score = (score_hint * 0.75) + (score_maq * 0.25)
        score = max(score, word_score)
        
        # 4. Similitud de Secuencia (Para errores de tipeo generales)
        seq_ratio = difflib.SequenceMatcher(None, hint_norm, maq_norm).ratio()
        score = max(score, seq_ratio)
        
        # Guardar el puntaje más alto encontrado en toda la lista
        if score > max_score:
            max_score = score
            mejor_maq = maq
            
    # Retornar el ganador solo si la confianza de la IA es mayor al 60%
    if max_score >= 0.60:
        return mejor_maq
        
    return hint # Si falla todo, devuelve el texto original

def buscar_datos_excel(codigo_buscado):
    codigo_norm = normalizar_texto(codigo_buscado)
    file_path = os.path.join(UPLOAD_DIR, "matriz_actual.xlsx")
    if not os.path.exists(file_path): return None
    try: df = pd.read_excel(file_path, sheet_name='PLAN.MANT', header=None)
    except: return None

    idx_maquina, idx_horas, idx_fecha = -1, -1, -1
    idx_t1, idx_t2, idx_t3, idx_t4 = -1, -1, -1, -1
    idx_tipo_atencion = -1
    fila_enc = 0
    columnas_fecha = [] # NUEVO: Historial de todas las fechas
    
    for r, row in df.head(25).iterrows():
        for c, cell in enumerate(row):
            c_norm = normalizar_texto(cell)
            if 'MAQUINA' in c_norm and idx_maquina == -1: idx_maquina = c; fila_enc = r
            elif 'TECNICO 1' in c_norm and idx_t1 == -1: idx_t1 = c
            elif 'TECNICO 2' in c_norm and idx_t2 == -1: idx_t2 = c
            elif 'TECNICO 3' in c_norm and idx_t3 == -1: idx_t3 = c
            elif 'TECNICO 4' in c_norm and idx_t4 == -1: idx_t4 = c
            elif ('HORAS' in c_norm or 'REALES' in c_norm or 'H H' in c_norm) and idx_horas == -1: idx_horas = c
            elif 'ATENCION' in c_norm and idx_tipo_atencion == -1: idx_tipo_atencion = c
            # ---> NUEVA LÓGICA DE FECHAS <---
            elif 'TERMINO' in c_norm: 
                idx_fecha = c
            elif 'FECHA' in c_norm:
                columnas_fecha.append(c)

    # Si no encontró la palabra explícita 'TERMINO', asume que es la última 'FECHA' del Excel
    if idx_fecha == -1 and columnas_fecha:
        idx_fecha = columnas_fecha[-1]

    if idx_maquina == -1: idx_maquina = 1
    if idx_fecha == -1: idx_fecha = 17

    fila_encontrada = None
    for r in range(fila_enc + 1, len(df)):
        celda_raw = df.iloc[r, idx_maquina]
        if pd.notna(celda_raw):
            c_raw_str = str(celda_raw).strip()
            if codigo_buscado == c_raw_str or codigo_norm == normalizar_texto(c_raw_str):
                fila_encontrada = df.iloc[r]
                break

    if fila_encontrada is None: return None

    orden_trabajo = ""
    for cell in fila_encontrada:
        if pd.notna(cell):
            val = str(cell).strip().upper()
            match = re.search(r'OSS\d+', val)
            if match:
                orden_trabajo = match.group(0)
                break
                
    horas_str = ""
    if idx_horas != -1 and pd.notna(fila_encontrada.iloc[idx_horas]):
        val = str(fila_encontrada.iloc[idx_horas]).strip().upper().replace('HORAS', '').strip()
        if val and "OSS" not in val and "REALIZADO" not in val: horas_str = val
            
    fecha_str = ""
    if idx_fecha != -1 and pd.notna(fila_encontrada.iloc[idx_fecha]):
        fecha_val = fila_encontrada.iloc[idx_fecha]
        if hasattr(fecha_val, 'month'): 
            meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
            try: fecha_str = f"{meses[fecha_val.month - 1]} - {fecha_val.year}"
            except: fecha_str = str(fecha_val).strip().upper()
        else: fecha_str = str(fecha_val).strip().upper()

    def limpiar_t(idx):
        if idx == -1 or pd.isna(fila_encontrada.iloc[idx]): return ""
        val = str(fila_encontrada.iloc[idx]).strip()
        if "OSS" in val.upper() or "REALIZADO" in val.upper(): return ""
        return val
        
    t1, t2, t3, t4 = limpiar_t(idx_t1), limpiar_t(idx_t2), limpiar_t(idx_t3), limpiar_t(idx_t4)
    datos_vacios = not orden_trabajo and not horas_str and not t1

    tipo_atencion_str = ""
    if idx_tipo_atencion != -1 and pd.notna(fila_encontrada.iloc[idx_tipo_atencion]):
        val_atencion = str(fila_encontrada.iloc[idx_tipo_atencion]).strip().upper()
        if "INTERNO" in val_atencion: tipo_atencion_str = "INTERNO"
        elif "EXTERNO" in val_atencion: tipo_atencion_str = "EXTERNO"

    return {
        "orden_trabajo": orden_trabajo, "horas": horas_str, "fecha": fecha_str,
        "tecnico1": t1, "tecnico2": t2, "tecnico3": t3, "tecnico4": t4,
        "tipo_mantenimiento": tipo_atencion_str,  # <--- SE AGREGA AQUÍ
        "alerta_vacio": datos_vacios
    }


# --- RUTAS DE LA API ---
@app.post("/api/cargar-excel")
async def cargar_excel(excel: UploadFile = File(...)):
    file_path = os.path.join(UPLOAD_DIR, "matriz_actual.xlsx")
    with open(file_path, "wb") as buffer: shutil.copyfileobj(excel.file, buffer)
    maquinas = obtener_maquinas_excel()
    if not maquinas: raise HTTPException(status_code=400, detail="Error leyendo Excel.")
    return {"maquinas": maquinas}


# --- 2. RUTA: PROCESAR WORD (AUTOMÁTICO) ---
@app.post("/api/procesar-word-auto")
async def procesar_word_auto(word: UploadFile = File(...)):
    file_path = os.path.join(UPLOAD_DIR, "plantilla_temp.docx")
    with open(file_path, "wb") as buffer: shutil.copyfileobj(word.file, buffer)
        
    lista_maquinas = obtener_maquinas_excel()
    maquina_oficial = None
    area = ""
    tipo_mantenimiento = ""
    diagnostico_encontrado = ""
    nota_encontrada = ""

    # PRIORIDAD 1: Buscar usando el Nombre del Archivo (Indestructible)
    filename_limpio = word.filename.upper().replace(".DOCX", "").replace(".DOC", "")
    maquina_oficial = coincidencia_inteligente(filename_limpio, lista_maquinas)
    
    if maquina_oficial not in lista_maquinas:
        maquina_oficial = None

    codigo_interno = ""
    texto_completo = []

    try: 
        doc = docx.Document(file_path)
        
        # A) Escaneo de Tablas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells: texto_completo.append(cell.text)
                if len(row.cells) >= 2:
                    col0_norm = normalizar_texto(row.cells[0].text)
                    col1_text = row.cells[1].text.strip()
                    
                    if ("MAQUINA O EQUIPO" in col0_norm or "CODIGO" in col0_norm) and not codigo_interno:
                        codigo_interno = col1_text
                    elif "AREA" in col0_norm and not area: 
                        area = col1_text
                    elif "DIAGNOSTICO" in col0_norm and not diagnostico_encontrado:
                        texto_celda = row.cells[1].text.strip() if len(row.cells) >= 2 else ""
                        if texto_celda:
                            # Limpiar prefijos si la celda arrastró la etiqueta
                            texto_limpio = re.sub(r'^(?:DIAGN[ÓO]STICO(?:\s+DE\s+FALLA)?|DE\s+FALLA)\s*[:\-]*\s*', '', texto_celda, flags=re.IGNORECASE)
                            diagnostico_encontrado = texto_limpio.strip().upper()

        # B) Escaneo de Párrafos
        for p in doc.paragraphs:
            texto_completo.append(p.text)
            p_norm = normalizar_texto(p.text)
            p_text = p.text.strip()
            
            if not codigo_interno:
                if "MAQUINA O EQUIPO" in p_norm or "CODIGO DE LA MAQUINA" in p_norm:
                    if ":" in p_text: codigo_interno = p_text.split(":", 1)[1].strip()
                    else: codigo_interno = p_norm.split("EQUIPO")[-1].strip()
            if not area:
                if "AREA" in p_norm and ":" in p_text:
                    area = p_text.split(":", 1)[1].strip()
            if not diagnostico_encontrado and "DIAGNOSTICO" in p_norm:
                if ":" in p_text:
                    val = p_text.split(":", 1)[1].strip()
                else:
                    val = p_text
                # Limpiar cualquier residuo de la etiqueta al inicio
                texto_limpio = re.sub(r'^(?:DIAGN[ÓO]STICO(?:\s+DE\s+FALLA)?|DE\s+FALLA)\s*[:\-]*\s*', '', val, flags=re.IGNORECASE)
                diagnostico_encontrado = texto_limpio.strip().upper()
            if "NOTA" in p_norm:
                if ":" in p_text:
                    val = p_text.split(":", 1)[1].strip().upper()
                else:
                    val = p_text.upper()
                # Limpiamos la palabra NOTA al inicio
                texto_limpio = re.sub(r'^NOTA\s*[:\-]*\s*', '', val, flags=re.IGNORECASE).strip()
                if texto_limpio:
                    if nota_encontrada:
                        nota_encontrada += "\n" + texto_limpio
                    else:
                        nota_encontrada = texto_limpio

        # ====================================================================
        # C) INTELIGENCIA EXTREMA PARA "TIPO DE MANTENIMIENTO" (SUPER CAZADOR)
        # ====================================================================
        

        r'''
        # --- 1. CAZADOR DE TABLAS DIBUJADAS (EL SECRETO REVELADO) ---
        # Detecta si las casillas son en realidad celdas de una pequeña tabla
        try:
            tablas_casilla = []
            for table in doc.tables:
                if len(table.rows) >= 2:
                    for i in range(len(table.rows) - 1):
                        f1_txt = "".join([c.text.upper() for c in table.rows[i].cells])
                        f2_txt = "".join([c.text.upper() for c in table.rows[i+1].cells])
                        
                        f1 = re.sub(r'[^A-Z0-9]', '', f1_txt)
                        f2 = re.sub(r'[^A-Z0-9]', '', f2_txt)
                        
                        # Caso A: Tablita de 2 filas solo con la X y un vacío
                        if f1 in ["X", ""] and f2 in ["X", ""] and (f1 or f2):
                            if f2 == "X": tipo_mantenimiento = "EXTERNO"
                            elif f1 == "X": tipo_mantenimiento = "INTERNO"
                            
                        # Caso B: Tabla que incluye los textos TIPO I y TIPO II
                        elif "TIPOI" in f1 and "TIPOII" in f2:
                            if "X" in f2: tipo_mantenimiento = "EXTERNO"
                            elif "X" in f1: tipo_mantenimiento = "INTERNO"
                
                # Caso C: Tablitas sueltas de 1x1
                texto_total = re.sub(r'[^A-Z0-9]', '', "".join([c.text.upper() for row in table.rows for c in row.cells]))
                if len(texto_total) <= 1:
                    tablas_casilla.append(texto_total)
            
            # Si encontró dos tablitas sueltas (una para cada opción)
            if not tipo_mantenimiento and len(tablas_casilla) >= 2:
                if tablas_casilla[1] == "X" and tablas_casilla[0] == "": tipo_mantenimiento = "EXTERNO"
                elif tablas_casilla[0] == "X" and tablas_casilla[1] == "": tipo_mantenimiento = "INTERNO"
        except Exception:
            pass

        # --- 2. CAZADOR XML (CHECKBOXES NATIVOS Y WINGDINGS) ---
        if not tipo_mantenimiento:
            import zipfile
            try:
                with zipfile.ZipFile(file_path, 'r') as docx_zip:
                    xml_content = docx_zip.read('word/document.xml').decode('utf-8').upper()
                    
                    checkboxes = []
                    for m in re.finditer(r'<W:CHECKBOX(.*?)</W:CHECKBOX>', xml_content, flags=re.DOTALL):
                        is_checked = '<W:CHECKED' in m.group(1) and '<W:VAL="0"' not in m.group(1) and '<W:VAL="FALSE"' not in m.group(1)
                        checkboxes.append((m.start(), is_checked))
                        
                    for m in re.finditer(r'<W14:CHECKBOX(.*?)</W14:CHECKBOX>', xml_content, flags=re.DOTALL):
                        is_checked = '<W14:CHECKED' in m.group(1) and '<W14:VAL="0"' not in m.group(1)
                        checkboxes.append((m.start(), is_checked))
                        
                    for m in re.finditer(r'<W:SYM W:FONT="WINGDINGS" W:CHAR="([A-F0-9]+)"/>', xml_content):
                        c_code = m.group(1)
                        if c_code in ["FE", "FC", "FD", "FB", "58"]: checkboxes.append((m.start(), True))
                        elif c_code in ["A8", "6F", "9F", "F0"]: checkboxes.append((m.start(), False))
                            
                    checkboxes.sort(key=lambda x: x[0])
                    
                    if len(checkboxes) >= 2:
                        if checkboxes[1][1] and not checkboxes[0][1]: tipo_mantenimiento = "EXTERNO"
                        elif checkboxes[0][1] and not checkboxes[1][1]: tipo_mantenimiento = "INTERNO"
            except:
                pass

        # --- 3. RADAR DE TEXTO LINEAL (RED DE SEGURIDAD FINAL) ---
        if not tipo_mantenimiento:
            texto_lineal = []
            for element in doc.element.body:
                if element.tag.endswith('p'):
                    texto_lineal.append(docx.text.paragraph.Paragraph(element, doc).text.strip())
                elif element.tag.endswith('tbl'):
                    table = docx.table.Table(element, doc)
                    for row in table.rows:
                        for cell in row.cells: texto_lineal.append(cell.text.strip())
                            
            texto_colapsado = re.sub(r'\s+', ' ', " | ".join([t for t in texto_lineal if t]).upper())
            
            if any(x in texto_colapsado.replace(" ", "") for x in ["[X]TIPOII", "☒TIPOII", "TIPOII-EXTERNO[X]", "TIPOIIX"]):
                tipo_mantenimiento = "EXTERNO"
            elif any(x in texto_colapsado.replace(" ", "") for x in ["[X]TIPOI", "☒TIPOI", "TIPOI-INTERNO[X]", "TIPOIX"]):
                tipo_mantenimiento = "INTERNO"
            else:
                match_t1 = re.search(r'TIPO\s*I\b(?!I)', texto_colapsado)
                match_t2 = re.search(r'TIPO\s*I{2}\b', texto_colapsado)
                
                pt1_s = match_t1.start() if match_t1 else -1
                pt1_e = match_t1.end() if match_t1 else -1
                pt2_s = match_t2.start() if match_t2 else -1
                pt2_e = match_t2.end() if match_t2 else -1
                
                matches_x = list(re.finditer(r'(?:\[\s*X\s*\]|\(\s*X\s*\)|\bX\b|[☒☑✔✗×])', texto_colapsado))
                mejor_distancia = 300 
                
                for m in matches_x:
                    x_s, x_e = m.start(), m.end()
                    d1 = min(abs(x_s - pt1_e), abs(x_e - pt1_s)) if pt1_s != -1 else 9999
                    d2 = min(abs(x_s - pt2_e), abs(x_e - pt2_s)) if pt2_s != -1 else 9999
                    
                    if min(d1, d2) < mejor_distancia:
                        mejor_distancia = min(d1, d2)
                        tipo_mantenimiento = "INTERNO" if d1 < d2 else "EXTERNO"
         '''

    except Exception:
        pass

    # D) Cruzar datos finales
    if not maquina_oficial and codigo_interno:
        codigo_interno = re.sub(r'\s+X$', '', codigo_interno, flags=re.IGNORECASE).strip()
        maquina_oficial = coincidencia_inteligente(codigo_interno, lista_maquinas)

    datos_finales = {}
    if maquina_oficial:
        datos_excel = buscar_datos_excel(maquina_oficial)
        if datos_excel: 
            datos_finales = datos_excel
            # ---> NUEVO: Usar el tipo de mantenimiento que viene directo del Excel
            if datos_excel.get("tipo_mantenimiento"):
                tipo_mantenimiento = datos_excel["tipo_mantenimiento"]
            
    if area: datos_finales["area"] = area.upper()
    if tipo_mantenimiento: datos_finales["tipo_mantenimiento"] = tipo_mantenimiento
    
    # Siempre enviamos diagnóstico y nota para forzar la limpieza en el frontend
    datos_finales["diagnostico"] = diagnostico_encontrado
    datos_finales["nota"] = nota_encontrada

    if maquina_oficial:
        return {"autodetectado": True, "codigoDetectado": maquina_oficial, "data": datos_finales}
    else:
        return {"autodetectado": False, "codigoDetectado": "", "data": datos_finales}


@app.post("/api/consultar-maquina")
async def consultar_maquina(data: dict):
    codigo_buscado = data.get("codigo", "").strip()
    datos = buscar_datos_excel(codigo_buscado)
    if datos: return {"data": datos}
    else: raise HTTPException(status_code=404, detail="Máquina no encontrada.")


# --- 4. RUTA FINAL: GENERACIÓN DE DOCUMENTO DESDE CERO ---
@app.post("/api/generar-ficha")
async def generar_ficha(
    file: UploadFile = File(...),
    codigo: str = Form(""),
    orden_trabajo: str = Form(""),
    tecnico1: str = Form(""),
    tecnico2: str = Form(""),
    tecnico3: str = Form(""),
    tecnico4: str = Form(""),
    fecha: str = Form(""),
    horas: str = Form(""),
    diagnostico: str = Form(""),
    tipo_mantenimiento: str = Form(""),
    nota: str = Form(""),
    area: str = Form("")
):
    file_path = os.path.join(UPLOAD_DIR, "generacion_temp.docx")
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # =================================================================
    # FASE A: EXTRAER TAREAS POR CATEGORÍA, DIAGNÓSTICO Y NOTA
    # =================================================================
    doc_origen = docx.Document(file_path)
    tareas_mecanicas = []
    tareas_electricas = []
    tareas_mixtas = []
    tareas_electronicas = []
    tareas_neumaticas = []
    repuestos_lista = []
    categoria_actual = "MECANICO"
    orden_categorias = ["MECANICO"]
    diagnostico_extraido = ""
    nota_extraida = ""
    
    # Lista ampliada con todas las etiquetas del nuevo formato automatizado
    titulos_basura = [
        "OPERACION", "EJECUCION", "MECANICO", "ELECTRICO", "ELECTRONICO", "NEUMATICO", 
        "ELECTRICO MECANICO", "REPUESTOS",
        "CODIGO DE LA MAQUINA O EQUIPO", "ORDEN DE TRABAJO OT", "AREA",
        "TIPO DE MANTENIMIENTO", "TIPO DE EJECUCION", "SERVICIO",
        "FECHA DE EJECUCION", "HORAS REALES UTILIZADAS", "TECNICO RESPONSABLE",
        "DIAGNOSTICO DE FALLA", 
        "OPERACION MECANICO", "OPERACION ELECTRICO", "OPERACION ELECTRONICO", 
        "OPERACION NEUMATICO", "OPERACION ELECTRICO MECANICO"
    ]
    
    for element in doc_origen.element.body:
        # --- ESCANEO DE PÁRRAFOS ---
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, doc_origen)
            
            for linea_raw in p.text.split('\n'):
                p_text = linea_raw.strip()
                p_norm = normalizar_texto(p_text)
                
                if not p_text: continue
                
                # 1. Detectar cambio de categoría aislando la palabra clave
                nueva_cat = None
                cat_base = p_norm.replace("OPERACION", "").replace("EJECUCION", "").strip()
                
                if cat_base in ["REPUESTOS", "REPUESTO"]: nueva_cat = "REPUESTOS"
                elif cat_base in ["ELECTRONICO", "ELECTRONICA"]: nueva_cat = "ELECTRONICO"
                elif cat_base in ["NEUMATICO", "NEUMATICA"]: nueva_cat = "NEUMATICO"
                elif cat_base in ["ELECTRICO", "ELECTRICA"]: nueva_cat = "ELECTRICO"
                elif cat_base in ["MECANICO", "MECANICA"]: nueva_cat = "MECANICO"
                elif cat_base in ["ELECTRICO MECANICO", "MECANICO ELECTRICO", "MIXTO"]: nueva_cat = "MIXTO"
                
                if nueva_cat:
                    categoria_actual = nueva_cat
                    if categoria_actual not in orden_categorias:
                        orden_categorias.append(categoria_actual)
                    continue
                    
                # 2. Atrapar repuestos si están escritos como texto normal
                if categoria_actual == "REPUESTOS" and "NOTA" not in p_norm and "DIAGNOSTICO" not in p_norm:
                    if p_text.upper() not in repuestos_lista and p_norm not in titulos_basura:
                        repuestos_lista.append(p_text.upper())
                    
                # 3. Extraer Diagnóstico
                if "DIAGNOSTICO" in p_norm and not diagnostico_extraido:
                    val = p_text.split(":", 1)[1].strip() if ":" in p_text else p_text
                    texto_limpio = re.sub(r'^(?:DIAGN[ÓO]STICO(?:\s+DE\s+FALLA)?|DE\s+FALLA)\s*[:\-]*\s*', '', val, flags=re.IGNORECASE)
                    if texto_limpio: diagnostico_extraido = texto_limpio.strip().upper()
                
                # 4. Acumular Notas
                if "NOTA" in p_norm and "NOTAS" not in p_norm:
                    val = p_text.split(":", 1)[1].strip() if ":" in p_text else p_text
                    texto_limpio = re.sub(r'^NOTA\s*[:\-]*\s*', '', val, flags=re.IGNORECASE).strip().upper()
                    if texto_limpio:
                        if nota_extraida: nota_extraida += "\n" + texto_limpio
                        else: nota_extraida = texto_limpio
                        
        # --- ESCANEO DE TABLAS ---
        elif element.tag.endswith('tbl'):
            table = docx.table.Table(element, doc_origen)
            
            for row in table.rows:
                if len(row.cells) >= 2:
                    c0_norm = normalizar_texto(row.cells[0].text)
                    c1_text = row.cells[1].text.strip()
                    if "DIAGNOSTICO" in c0_norm and c1_text and not diagnostico_extraido:
                        texto_limpio = re.sub(r'^(?:DIAGN[ÓO]STICO(?:\s+DE\s+FALLA)?|DE\s+FALLA)\s*[:\-]*\s*', '', c1_text, flags=re.IGNORECASE)
                        diagnostico_extraido = texto_limpio.strip().upper()

            # Evitar tabla principal de datos generales
            text_tabla = "".join(cell.text.upper() for row in table.rows for cell in row.cells)
            if "ORDEN DE TRABAJO" in text_tabla or "CÓDIGO DE LA MÁQUINA" in text_tabla or "CODIGO DE LA MAQUINA" in text_tabla:
                continue
                
            for row in table.rows:
                if not row.cells: continue
                
                c0_text = row.cells[0].text.strip()
                c0_norm = normalizar_texto(c0_text)
                if not c0_text: continue
                
                # Detectar cambio de categoría DENTRO de la tabla aislando la palabra clave
                nueva_cat_tbl = None
                cat_base = c0_norm.replace("OPERACION", "").replace("EJECUCION", "").strip()
                
                if cat_base in ["REPUESTOS", "REPUESTO"]: nueva_cat_tbl = "REPUESTOS"
                elif cat_base in ["ELECTRONICO", "ELECTRONICA"]: nueva_cat_tbl = "ELECTRONICO"
                elif cat_base in ["NEUMATICO", "NEUMATICA"]: nueva_cat_tbl = "NEUMATICO"
                elif cat_base in ["ELECTRICO", "ELECTRICA"]: nueva_cat_tbl = "ELECTRICO"
                elif cat_base in ["MECANICO", "MECANICA"]: nueva_cat_tbl = "MECANICO"
                elif cat_base in ["ELECTRICO MECANICO", "MECANICO ELECTRICO", "MIXTO"]: nueva_cat_tbl = "MIXTO"

                if nueva_cat_tbl:
                    categoria_actual = nueva_cat_tbl
                    if categoria_actual not in orden_categorias:
                        orden_categorias.append(categoria_actual)
                    continue

                if categoria_actual == "REPUESTOS":
                    # Si es un repuesto, lo agregamos y lo filtramos para que no sea la palabra "REPUESTOS"
                    if c0_text.upper() not in repuestos_lista and c0_norm not in titulos_basura and "DIAGNOSTICO" not in c0_norm and "NOTA" not in c0_norm:
                        for item in c0_text.split('\n'):
                            i_txt = item.strip().upper() # <--- Fuerza repuestos de tabla a mayúscula
                            if i_txt and i_txt not in repuestos_lista:
                                repuestos_lista.append(i_txt)
                else:
                    if len(row.cells) >= 2:
                        t_estado = row.cells[-1].text.strip()
                        
                        if not t_estado or "<" in t_estado or "TD COLSPAN" in t_estado.upper():
                            t_estado = "OK"
                        elif "OK" in t_estado.upper():
                            t_estado = "OK"
                            
                        # <-- EL FILTRO MÁGICO Y CONVERSIÓN A MAYÚSCULA
                        if len(c0_text) > 3 and c0_norm not in titulos_basura and "DIAGNOSTICO" not in c0_norm and "NOTA" not in c0_norm:
                            tarea_mayuscula = c0_text.upper() # <--- Transforma la tarea (ej: "Revisión...") a mayúscula absoluta
                            
                            if categoria_actual == "MIXTO": tareas_mixtas.append((tarea_mayuscula, t_estado))
                            elif categoria_actual == "ELECTRICO": tareas_electricas.append((tarea_mayuscula, t_estado))
                            elif categoria_actual == "ELECTRONICO": tareas_electronicas.append((tarea_mayuscula, t_estado))
                            elif categoria_actual == "NEUMATICO": tareas_neumaticas.append((tarea_mayuscula, t_estado))
                            elif categoria_actual == "MECANICO": tareas_mecanicas.append((tarea_mayuscula, t_estado))

    # El formulario web ahora es la autoridad absoluta
    diagnostico_final = diagnostico.strip().upper()
    nota_final = nota.strip().upper()

    # =================================================================
    # FASE B: CONSTRUCCIÓN MILIMÉTRICA
    # =================================================================
    doc = docx.Document()
    
    # 1. Configurar Márgenes Exactos
    for section in doc.sections:
        section.top_margin = Cm(1.2)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.17)
        section.right_margin = Cm(3.17)
        # Ancho utilizable de página = 21cm - 3.17cm - 3.17cm = 14.66cm

    # 2. TÍTULOS CENTRALES
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Primer texto: Institución (Tamaño 18)
    run_bcrp = p_titulo.add_run("BANCO CENTRAL DE RESERVA DEL PERÚ\n")
    run_bcrp.font.name = 'Arial'
    run_bcrp.font.size = Pt(18)
    run_bcrp.bold = False
    
    # El espacio intermedio: Salto de línea (Tamaño 11)
    run_espacio = p_titulo.add_run("\n") 
    run_espacio.font.size = Pt(11)
    
    # Segundo texto: Subtítulo (Tamaño 14)
    run_ficha = p_titulo.add_run("FICHA DE MANTENIMIENTO")
    run_ficha.font.name = 'Arial'
    run_ficha.font.size = Pt(14)
    run_ficha.bold = True
    
    doc.add_paragraph() # Espacio separador antes de la primera tabla

    # --- FUNCIÓN MAESTRA PARA TABLAS (Calibri 11) ---
    def crear_bloque_tabla(datos):
        table = doc.add_table(rows=len(datos), cols=2)
        table.style = 'Table Grid'
        for i, (label, valor) in enumerate(datos):
            cells = table.rows[i].cells
            
            # Celda Izquierda (Etiqueta) - Negrita, Izquierda
            cells[0].text = ""
            p0 = cells[0].paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r0 = p0.add_run(label)
            r0.font.name = 'Calibri'
            r0.font.size = Pt(11)
            r0.bold = True
            
            # PINTAR EL FONDO DE GRIS (Columna 1)
            tcPr0 = cells[0]._element.get_or_add_tcPr()
            shd0 = OxmlElement('w:shd')
            shd0.set(qn('w:val'), 'clear')
            shd0.set(qn('w:color'), 'auto')
            shd0.set(qn('w:fill'), 'F2F2F2')
            tcPr0.append(shd0)
            
            # Celda Derecha (Valor) - Con soporte para poner la X en negrita
            cells[1].text = ""
            p1 = cells[1].paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            # Separar por líneas (por si hay saltos de línea como en Tipo de Ejecución)
            lineas = str(valor).split('\n')
            for idx_l, linea in enumerate(lineas):
                if idx_l > 0:
                    p1.add_run("\n")
                
                # Si la línea contiene la marca [ X ], aplicamos negrita solo a la X
                if "[ X ]" in linea:
                    partes = linea.split("[ X ]")
                    
                    r_antes = p1.add_run(partes[0] + "[ ")
                    r_antes.font.name = 'Calibri'
                    r_antes.font.size = Pt(11)
                    
                    r_x = p1.add_run("X")
                    r_x.font.name = 'Calibri'
                    r_x.font.size = Pt(11)
                    r_x.bold = True  # <--- Aquí va la X en negrita
                    
                    r_despues = p1.add_run(" ]" + partes[1])
                    r_despues.font.name = 'Calibri'
                    r_despues.font.size = Pt(11)
                else:
                    # Texto normal si no lleva X marcada
                    r_norm = p1.add_run(linea)
                    r_norm.font.name = 'Calibri'
                    r_norm.font.size = Pt(11)
            
            # Ajuste perfecto con el "ancho preferido" de tu plantilla
            cells[0].width = Cm(6.24)
            cells[1].width = Cm(9.45)
            
            # Centrado vertical de la celda
            for c in cells:
                tcVAlign = OxmlElement('w:vAlign')
                tcVAlign.set(qn('w:val'), 'center')
                c._element.tcPr.append(tcVAlign)
        
        # Espacio real de tamaño 4 (sin márgenes ocultos)
        p_espacio = doc.add_paragraph()
        p_espacio.paragraph_format.space_after = Pt(0)
        p_espacio.add_run().font.size = Pt(4)
        
        return table

    # 3. CONSTRUIR LOS 4 BLOQUES DE CABECERA
    crear_bloque_tabla([
        ("CÓDIGO DE LA MÁQUINA O EQUIPO", codigo.upper()),
        ("ORDEN DE TRABAJO (OT):", orden_trabajo.upper()),
        ("ÁREA:", area.upper())
    ])

    check_interno = "[ X ]" if tipo_mantenimiento == "INTERNO" else "[   ]"
    check_externo = "[ X ]" if tipo_mantenimiento == "EXTERNO" else "[   ]"
    texto_ejecucion = f"{check_interno} TIPO I - INTERNO\n{check_externo} TIPO II - EXTERNO"
    
    crear_bloque_tabla([
        ("TIPO DE MANTENIMIENTO:", "PREVENTIVO PROGRAMADO"),
        ("TIPO DE EJECUCIÓN:", texto_ejecucion),
        ("SERVICIO:", "MECÁNICO") 
    ])

    tecnicos_lista = [t.upper() for t in [tecnico1, tecnico2, tecnico3, tecnico4] if t]
    tecnicos_unidos = "\n".join(tecnicos_lista)
    
    crear_bloque_tabla([
        ("FECHA DE EJECUCIÓN:", fecha.upper()),
        ("HORAS REALES UTILIZADAS:", str(horas)),
        ("TÉCNICO RESPONSABLE:", tecnicos_unidos)
    ])

    crear_bloque_tabla([
        ("DIAGNÓSTICO DE FALLA:", diagnostico_final)
    ])

    # 4. TABLAS DE OPERACIONES (Mecánico y Eléctrico)

    def construir_tabla_categoria(titulo_cat, lista_tareas):
        if not lista_tareas:
            return
            
        # (Eliminamos la creación del subtítulo suelto para ponerlo en la tabla)

        table_ops = doc.add_table(rows=1, cols=2)
        table_ops.style = 'Table Grid'
        
        table_ops.rows[0].height = Cm(0.34)
        table_ops.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        
        hdr_cells = table_ops.rows[0].cells
        
        # ---> NUEVO: Combinamos "OPERACIÓN" con la categoría (Ej: OPERACIÓN - MECÁNICO) <---
        titulos_encabezado = [f'OPERACIÓN - {titulo_cat}', 'EJECUCIÓN']
        
        for i, cell in enumerate(hdr_cells):
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
            r = p.add_run(titulos_encabezado[i])
            r.font.name = 'Calibri'
            r.font.size = Pt(11)
            r.bold = True
            
            tcPr = cell._element.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), 'D9D9D9')
            tcPr.append(shd)
            
            tcVAlign = OxmlElement('w:vAlign')
            tcVAlign.set(qn('w:val'), 'center')
            tcPr.append(tcVAlign)
            
            cell.width = Cm(13.74) if i == 0 else Cm(0.92)

        for tarea, estado in lista_tareas:
            nueva_fila = table_ops.add_row()
            nueva_fila.height = Cm(0.62)
            nueva_fila.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            
            row_cells = nueva_fila.cells
            
            row_cells[0].text = ""
            p0 = row_cells[0].paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r0 = p0.add_run(tarea)
            r0.font.name = 'Calibri'
            r0.font.size = Pt(11)
            row_cells[0].width = Cm(13.74)
            
            tcPr0 = row_cells[0]._element.get_or_add_tcPr()
            tcVAlign0 = OxmlElement('w:vAlign')
            tcVAlign0.set(qn('w:val'), 'center')
            tcPr0.append(tcVAlign0)
            
            row_cells[1].text = ""
            p1 = row_cells[1].paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1 = p1.add_run(estado)
            r1.font.name = 'Calibri'
            r1.font.size = Pt(11)
            row_cells[1].width = Cm(0.92)
            
            tcPr1 = row_cells[1]._element.get_or_add_tcPr()
            tcVAlign1 = OxmlElement('w:vAlign')
            tcVAlign1.set(qn('w:val'), 'center')
            tcPr1.append(tcVAlign1)

        p_espacio_ops = doc.add_paragraph()
        p_espacio_ops.paragraph_format.space_after = Pt(0)
        p_espacio_ops.add_run().font.size = Pt(4)

   # Construir las tablas en el orden exacto en el que aparecieron en el documento original
    for cat in orden_categorias:
        if cat == "MECANICO":
            construir_tabla_categoria("MECÁNICO", tareas_mecanicas)
        elif cat == "ELECTRICO":
            construir_tabla_categoria("ELÉCTRICO", tareas_electricas)
        elif cat == "MIXTO":
            construir_tabla_categoria("ELÉCTRICO - MECÁNICO", tareas_mixtas)
        elif cat == "ELECTRONICO":
            construir_tabla_categoria("ELECTRÓNICO", tareas_electronicas)
        elif cat == "NEUMATICO":
            construir_tabla_categoria("NEUMÁTICO", tareas_neumaticas)
        elif cat == "REPUESTOS" and repuestos_lista:
            p_rep = doc.add_paragraph()
            r_rep = p_rep.add_run("REPUESTOS:")
            r_rep.font.name = 'Calibri'
            r_rep.font.size = Pt(11)
            r_rep.bold = True
            
            table_rep = doc.add_table(rows=len(repuestos_lista), cols=1)
            table_rep.style = 'Table Grid'
            
            for i, rep in enumerate(repuestos_lista):
                cell = table_rep.rows[i].cells[0]
                cell.text = rep
                
                p_cell = cell.paragraphs[0]
                p_cell.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in p_cell.runs:
                    run.font.name = 'Calibri'
                    run.font.size = Pt(11)
                    
                tcPr = cell._element.get_or_add_tcPr()
                tcVAlign = OxmlElement('w:vAlign')
                tcVAlign.set(qn('w:val'), 'center')
                tcPr.append(tcVAlign)
                
            p_espacio_rep = doc.add_paragraph()
            p_espacio_rep.paragraph_format.space_after = Pt(0)
            p_espacio_rep.add_run().font.size = Pt(4)

    # 5. NOTA FINAL (Calibri 11) - Solo se dibuja si hay contenido
    if nota_final:
        p_nota = doc.add_paragraph()
        p_nota.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_nota_label = p_nota.add_run("NOTA: ")
        run_nota_label.font.name = 'Calibri'
        run_nota_label.font.size = Pt(11)
        run_nota_label.bold = True
        
        run_nota_texto = p_nota.add_run(nota_final)
        run_nota_texto.font.name = 'Calibri'
        run_nota_texto.font.size = Pt(11)
    # =================================================================
    # FASE C: EXPORTACIÓN
    # =================================================================
    nombre_archivo = f"{orden_trabajo}_{codigo}.docx".replace(" ", "_").replace("/", "-")
    output_path = os.path.join(UPLOAD_DIR, nombre_archivo)
    doc.save(output_path)
    
    return FileResponse(
        output_path, 
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=nombre_archivo
    )

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True) 