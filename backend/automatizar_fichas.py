import os
import re
import io
import sys
import base64
import datetime
import subprocess
from pathlib import Path

import docx
import pandas as pd
import requests

# ==========================================
# CONFIGURACIÓN GENERAL
# ==========================================
# Si tienes el enlace público de compartir de OneDrive, colócalo aquí.
# Si está vacío (None), el script usará el archivo de Excel local.
ONEDRIVE_SHARE_LINK = None  

EXCEL_LOCAL_PATH = "PLAN ANUAL DE MANTENIMIENTO 2025 2.xlsx"
CARPETA_SALIDA = "Fichas"


# ==========================================
# 1. LECTURA Y CONEXIÓN CON EL EXCEL
# ==========================================
def obtener_dataframe_excel(onedrive_link=None, local_path=None):
    """
    Carga la hoja 'PLAN.MANT' desde una URL pública de OneDrive o un archivo local.
    """
    if onedrive_link:
        try:
            print("🌐 Descargando Excel actualizado desde OneDrive...")
            # Convertir URL de compartición a formato compatible con la API de OneDrive
            base64_bytes = base64.b64encode(onedrive_link.encode("utf-8"))
            base64_str = base64_bytes.decode("utf-8").strip("=").replace('/', '_').replace('+', '-')
            direct_url = f"https://api.onedrive.com/v1.0/shares/u!{base64_str}/root/content"
            
            resp = requests.get(direct_url, timeout=15)
            if resp.status_code == 200:
                excel_bytes = io.BytesIO(resp.content)
                df = pd.read_excel(excel_bytes, sheet_name='PLAN.MANT', skiprows=9)
                print(" Excel descargado exitosamente desde la nube.")
                return df
            else:
                print(f"⚠️ No se pudo descargar desde OneDrive (Código HTTP {resp.status_code}). Usando respaldo local...")
        except Exception as e:
            print(f"⚠️ Error al conectar con OneDrive ({e}). Usando respaldo local...")

    # Carga desde archivo local
    if local_path and os.path.exists(local_path):
        print(f"📂 Cargando Excel local: {local_path}")
        return pd.read_excel(local_path, sheet_name='PLAN.MANT', skiprows=9)
    else:
        raise FileNotFoundError("No se encontró el archivo Excel local ni se pudo acceder a OneDrive.")


def extraer_datos_mantenimiento(df, codigo_maquina):
    """
    Busca la fila correspondiente al código de la máquina y consolida la información.
    """
    fila_encontrada = None
    
    for idx, row in df.iterrows():
        nombre_maquina = str(row.iloc[1])
        if str(codigo_maquina) in nombre_maquina:
            fila_encontrada = row
            break
            
    if fila_encontrada is None:
        raise ValueError(f"No se encontró la máquina con código '{codigo_maquina}' en la hoja de Excel.")
    
    # A. Orden de Trabajo (Extraída desde la columna Observaciones, ej: 'Realizado con OSS28025')
    obs = str(fila_encontrada.iloc[11]) if pd.notna(fila_encontrada.iloc[11]) else ""
    ot_match = re.search(r'OSS\d+', obs)
    orden_trabajo = ot_match.group(0) if ot_match else "S/N"
    
    # B. Horas Utilizadas (Horas-Hombre Reales)
    horas = fila_encontrada.iloc[10]
    horas_str = f"{horas} HORAS" if pd.notna(horas) else "0 HORAS"
    
    # C. Fecha de Ejecución
    fecha_val = fila_encontrada.iloc[13]
    if isinstance(fecha_val, (datetime.datetime, pd.Timestamp)):
        meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", 
                 "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
        fecha_str = f"{meses[fecha_val.month - 1]} - {fecha_val.year}"
    else:
        fecha_str = "DICIEMBRE - 2025"
        
    # D. Técnicos Responsables (Evaluación de las columnas Tecnico 1 a Tecnico 4)
    tecnicos = []
    cols_tecnicos = ['Tecnico 1', 'Tecnico 2', 'Tecnico 3', 'Tecnico 4']
    
    for col in cols_tecnicos:
        if col in df.columns:
            val = fila_encontrada[col]
            if pd.notna(val) and str(val).strip():
                tecnicos.append(str(val).strip())
                
    tecnico_str = " / ".join(tecnicos) if tecnicos else "MIGUEL GUERRERO"
    
    return {
        "orden_trabajo": orden_trabajo,
        "tecnico": tecnico_str,
        "fecha": fecha_str,
        "horas": horas_str
    }


# ==========================================
# 2. EDICIÓN DEL DOCUMENTO WORD
# ==========================================
def actualizar_plantilla_word(docx_input, output_docx_path, datos):
    """
    Edita únicamente las 4 secciones variables en la plantilla Word manteniendo el formato intacto.
    Soporta docx_input como ruta de archivo (str) o como flujo de bytes (io.BytesIO).
    """
    doc = docx.Document(docx_input)
    
    for paragraph in doc.paragraphs:
        texto = paragraph.text
        
        if "ORDEN DE TRABAJO:" in texto:
            paragraph.text = f"ORDEN DE TRABAJO: {datos['orden_trabajo']}"
        elif "TECNICO RESPONSABLE" in texto:
            paragraph.text = f"TECNICO RESPONSABLE\t{datos['tecnico']}"
        elif "FECHA DE EJECUCION" in texto:
            paragraph.text = f"FECHA DE EJECUCION\t   {datos['fecha']}"
        elif "HORAS UTILIZADAS" in texto:
            paragraph.text = f"HORAS UTILIZADAS\t  {datos['horas']}"
            
    doc.save(output_docx_path)
    print(f" Documento Word generado: {output_docx_path}")


# ==========================================
# 3. CONVERSIÓN MULTIPLATAFORMA A PDF
# ==========================================
def convertir_a_pdf(docx_path, folder_salida):
    """
    Convierte el archivo .docx a .pdf adaptándose al sistema operativo (Linux vs Windows).
    """
    docx_path = Path(docx_path)
    folder_salida = Path(folder_salida)
    
    # Entorno Linux (GitHub Codespaces / Servidores Cloud)
    if sys.platform.startswith('linux'):
        try:
            subprocess.run([
                'libreoffice', '--headless', '--convert-to', 'pdf',
                str(docx_path), '--outdir', str(folder_salida)
            ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            pdf_path = folder_salida / f"{docx_path.stem}.pdf"
            print(f" PDF generado con LibreOffice: {pdf_path}")
        except Exception as e:
            print(f"⚠️ Error al generar PDF con LibreOffice en Linux: {e}")
            print("Asegúrate de haber instalado LibreOffice (`sudo apt-get install libreoffice`).")
            
    # Entorno Windows (Equipo local)
    elif sys.platform.startswith('win'):
        try:
            from docx2pdf import convert
            pdf_path = folder_salida / f"{docx_path.stem}.pdf"
            convert(str(docx_path), str(pdf_path))
            print(f" PDF generado con Word/docx2pdf: {pdf_path}")
        except Exception as e:
            print(f"⚠️ Error al convertir a PDF en Windows: {e}")


# ==========================================
# 4. FLUJO PRINCIPAL DE EJECUCIÓN
# ==========================================
def procesar_ficha(docx_plantilla, onedrive_link=None, excel_local=EXCEL_LOCAL_PATH, carpeta_destino=CARPETA_SALIDA):
    """
    Función orquestadora principal.
    """
    # Crear carpeta de destino
    Path(carpeta_destino).mkdir(exist_ok=True, parents=True)
    
    # 1. Identificar el Código de la Máquina desde la plantilla subida/leída
    doc = docx.Document(docx_plantilla)
    codigo_maquina = None
    
    for p in doc.paragraphs:
        if "CODIGO DE LA MAQUINA" in p.text:
            match = re.search(r'\((\d+)\)', p.text)
            if match:
                codigo_maquina = match.group(1)
                break
                
    if not codigo_maquina:
        # Intento de rescate desde el nombre del archivo
        nombre_archivo = Path(docx_plantilla).name if isinstance(docx_plantilla, (str, Path)) else ""
        match_file = re.search(r'\(.*?(\d{5}).*?\)', nombre_archivo)
        codigo_maquina = match_file.group(1) if match_file else "60010"
        
    print(f"\n🔍 Procesando plantilla para máquina Código: {codigo_maquina}")
    
    # 2. Leer la matriz Excel
    df_excel = obtener_dataframe_excel(onedrive_link, excel_local)
    
    # 3. Extraer nuevos datos operativos
    datos_actualizados = extraer_datos_mantenimiento(df_excel, codigo_maquina)
    print(f"📋 Datos obtenidos: OT={datos_actualizados['orden_trabajo']} | Fecha={datos_actualizados['fecha']} | Horas={datos_actualizados['horas']} | Técnicos={datos_actualizados['tecnico']}")
    
    # 4. Generar nuevo archivo Word en la carpeta "Fichas"
    nombre_salida_base = f"{datos_actualizados['orden_trabajo']} ({codigo_maquina})_ACTUALIZADO"
    ruta_docx_salida = Path(carpeta_destino) / f"{nombre_salida_base}.docx"
    
    actualizar_plantilla_word(docx_plantilla, ruta_docx_salida, datos_actualizados)
    
    # 5. Exportar a PDF
    convertir_a_pdf(ruta_docx_salida, carpeta_destino)


# ==========================================
# PUNTO DE ENTRADA (PRUEBA DE EJECUCIÓN)
# ==========================================
if __name__ == "__main__":
    # Ruta de la plantilla Word a procesar (puedes cambiarla por cualquier ficha subida)
    PLANTILLA_EJEMPLO = "plantillas/OSS28025 (APILADOR ELÉCTRICO FN 1225 60010).docx"
    
    # Si la plantilla está en la raíz directamente, ajustar la ruta:
    if not os.path.exists(PLANTILLA_EJEMPLO) and os.path.exists("OSS28025 (APILADOR ELÉCTRICO FN 1225 60010).docx"):
        PLANTILLA_EJEMPLO = "OSS28025 (APILADOR ELÉCTRICO FN 1225 60010).docx"
        
    procesar_ficha(
        docx_plantilla=PLANTILLA_EJEMPLO,
        onedrive_link=ONEDRIVE_SHARE_LINK,
        excel_local=EXCEL_LOCAL_PATH,
        carpeta_destino=CARPETA_SALIDA
    )