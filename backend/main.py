import uvicorn
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import unicodedata
import shutil
import os
import re
import docx

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
    """Cruza un texto cualquiera (ej. nombre de archivo) con el catálogo del Excel"""
    hint_norm = normalizar_texto(hint)
    if not hint_norm: return None
    
    # 1. Búsqueda exacta
    for maq in lista_maquinas:
        if normalizar_texto(maq) == hint_norm: return maq
        
    # 2. Si el nombre de la máquina del Excel ESTÁ DENTRO del hint (Ej: Nombre de archivo)
    for maq in lista_maquinas:
        maq_norm = normalizar_texto(maq)
        if maq_norm and maq_norm in hint_norm: return maq
        
    # 3. Si el hint ESTÁ DENTRO de la máquina del Excel
    for maq in lista_maquinas:
        maq_norm = normalizar_texto(maq)
        if maq_norm and hint_norm in maq_norm: return maq
        
    return hint # Si falla todo, devuelve el texto original

def buscar_datos_excel(codigo_buscado):
    codigo_norm = normalizar_texto(codigo_buscado)
    file_path = os.path.join(UPLOAD_DIR, "matriz_actual.xlsx")
    if not os.path.exists(file_path): return None
    try: df = pd.read_excel(file_path, sheet_name='PLAN.MANT', header=None)
    except: return None

    idx_maquina, idx_horas, idx_fecha = -1, -1, -1
    idx_t1, idx_t2, idx_t3, idx_t4 = -1, -1, -1, -1
    fila_enc = 0
    
    for r, row in df.head(25).iterrows():
        for c, cell in enumerate(row):
            c_norm = normalizar_texto(cell)
            if 'MAQUINA' in c_norm and idx_maquina == -1: idx_maquina = c; fila_enc = r
            elif 'TECNICO 1' in c_norm and idx_t1 == -1: idx_t1 = c
            elif 'TECNICO 2' in c_norm and idx_t2 == -1: idx_t2 = c
            elif 'TECNICO 3' in c_norm and idx_t3 == -1: idx_t3 = c
            elif 'TECNICO 4' in c_norm and idx_t4 == -1: idx_t4 = c
            elif ('HORAS' in c_norm or 'REALES' in c_norm or 'H H' in c_norm) and idx_horas == -1: idx_horas = c
            elif ('TERMINO' in c_norm or 'FECHA' in c_norm) and idx_fecha == -1: idx_fecha = c

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
            meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
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

    return {
        "orden_trabajo": orden_trabajo, "horas": horas_str, "fecha": fecha_str,
        "tecnico1": t1, "tecnico2": t2, "tecnico3": t3, "tecnico4": t4,
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
                    col1_norm = normalizar_texto(row.cells[0].text)
                    if ("MAQUINA O EQUIPO" in col1_norm or "CODIGO" in col1_norm) and not codigo_interno:
                        codigo_interno = str(row.cells[1].text).strip()
                    elif "AREA" in col1_norm and not area: 
                        area = str(row.cells[1].text).strip()

        # B) Escaneo de Párrafos
        for p in doc.paragraphs:
            texto_completo.append(p.text)
            p_norm = normalizar_texto(p.text)
            
            if not codigo_interno:
                if "MAQUINA O EQUIPO" in p_norm or "CODIGO DE LA MAQUINA" in p_norm:
                    if ":" in p.text: codigo_interno = p.text.split(":", 1)[1].strip()
                    else: codigo_interno = p_norm.split("EQUIPO")[-1].strip()
            if not area:
                if "AREA" in p_norm and ":" in p.text:
                    area = p.text.split(":", 1)[1].strip()

        # C) INTELIGENCIA EXTREMA PARA "TIPO DE MANTENIMIENTO"
        texto_unido = "\n".join(texto_completo).upper()
        texto_sin_espacios = re.sub(r'\s+', '', texto_unido)
        
        # 1. Búsqueda exacta (Formato Nuevo)
        if "[X]TIPOII" in texto_sin_espacios or "☒TIPOII" in texto_sin_espacios or "☑TIPOII" in texto_sin_espacios:
            tipo_mantenimiento = "EXTERNO"
        elif "[X]TIPOI" in texto_sin_espacios or "☒TIPOI" in texto_sin_espacios or "☑TIPOI" in texto_sin_espacios:
            tipo_mantenimiento = "INTERNO"
        else:
            # 2. Algoritmo de "Radar" mejorado
            match_t1 = re.search(r'TIPO\s*I\b(?!I)', texto_unido)
            match_t2 = re.search(r'TIPO\s*I{2}\b', texto_unido)
            
            pos_t1 = match_t1.start() if match_t1 else -1
            pos_t2 = match_t2.start() if match_t2 else -1
            
            matches_x = list(re.finditer(r'(?:\bX\b|[☒☑✔])', texto_unido))
            mejor_distancia = 200 
            
            for m in matches_x:
                pos_x = m.start()
                if pos_t1 != -1:
                    dist1 = abs(pos_x - pos_t1)
                    if dist1 < mejor_distancia:
                        mejor_distancia = dist1; tipo_mantenimiento = "INTERNO"
                if pos_t2 != -1:
                    dist2 = abs(pos_x - pos_t2)
                    if dist2 < mejor_distancia:
                        mejor_distancia = dist2; tipo_mantenimiento = "EXTERNO"

    except Exception:
        # Si el Word está dañado o tiene formato viejo, ignoramos la lectura interna
        # El sistema no se detendrá y usará el Nombre del Archivo de la Prioridad 1
        pass

    # ====================================================================
    # D) LECTURA XML DIRECTA (Cazador de Cuadros de Texto Invisibles)
    # Si python-docx no vio la X porque está flotando, la buscamos a la fuerza.
    # ====================================================================
    if not tipo_mantenimiento:
        import zipfile
        try:
            with zipfile.ZipFile(file_path, 'r') as docx_zip:
                xml_content = docx_zip.read('word/document.xml').decode('utf-8').upper()
                raw_text = re.sub(r'<[^>]+>', '', xml_content) # Borra todo el código XML
                raw_text_clean = re.sub(r'\s+', '', raw_text) # Borra los espacios
                
                if "XTIPOII" in raw_text_clean or "X-TIPOII" in raw_text_clean: tipo_mantenimiento = "EXTERNO"
                elif "XTIPOI" in raw_text_clean or "X-TIPOI" in raw_text_clean: tipo_mantenimiento = "INTERNO"
        except:
            pass

    # E) Cruzar datos finales
    if not maquina_oficial and codigo_interno:
        codigo_interno = re.sub(r'\s+X$', '', codigo_interno, flags=re.IGNORECASE).strip()
        maquina_oficial = coincidencia_inteligente(codigo_interno, lista_maquinas)

    datos_finales = {}
    if maquina_oficial:
        datos_excel = buscar_datos_excel(maquina_oficial)
        if datos_excel: datos_finales = datos_excel
            
    if area: datos_finales["area"] = area.upper()
    if tipo_mantenimiento: datos_finales["tipo_mantenimiento"] = tipo_mantenimiento

    if maquina_oficial and datos_finales.get("orden_trabajo"):
        return {"autodetectado": True, "codigoDetectado": maquina_oficial, "data": datos_finales}
    else:
        return {"autodetectado": False, "codigoDetectado": maquina_oficial or "", "data": datos_finales}

        

@app.post("/api/consultar-maquina")
async def consultar_maquina(data: dict):
    codigo_buscado = data.get("codigo", "").strip()
    datos = buscar_datos_excel(codigo_buscado)
    if datos: return {"data": datos}
    else: raise HTTPException(status_code=404, detail="Máquina no encontrada.")


@app.post("/api/generar-ficha")
async def generar_ficha():
    pass

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)