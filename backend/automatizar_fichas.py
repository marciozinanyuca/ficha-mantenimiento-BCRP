import os
import re
import io
import sys
import datetime
import subprocess
from pathlib import Path

import docx
import pandas as pd
import requests

# ==========================================
# CONFIGURACIÓN
# ==========================================
ONEDRIVE_SHARE_LINK = "https://bcrpgob-my.sharepoint.com/:x:/g/personal/practicante_cnm2_bcrp_gob_pe/IQD6DWHDQYkuRZ_RbKo8gaTeATB0Fcw4koy2WIOTfCqdDlE?e=EoW35n"
EXCEL_LOCAL_PATH = "PLAN ANUAL DE MANTENIMIENTO 2025 2.xlsx"
CARPETA_SALIDA = "Fichas"


def obtener_dataframe_excel(onedrive_link=None, local_path=None):
    """
    Carga la hoja 'PLAN.MANT' desde SharePoint/OneDrive o un respaldo local.
    """
    if onedrive_link:
        try:
            print("🌐 Intentando descargar Excel desde SharePoint/OneDrive...")
            # Forzar parámetro de descarga directa para SharePoint
            direct_url = onedrive_link.replace("?e=", "&download=1?e=") if "?e=" in onedrive_link else onedrive_link + "&download=1"
            
            resp = requests.get(direct_url, timeout=15)
            if resp.status_code == 200:
                excel_bytes = io.BytesIO(resp.content)
                df = pd.read_excel(excel_bytes, sheet_name='PLAN.MANT', skiprows=9)
                print(" Excel cargado exitosamente desde la nube.")
                return df
            else:
                print(f"⚠️ No se pudo acceder vía URL directa (Status {resp.status_code}). Recurriendo al archivo local...")
        except Exception as e:
            print(f"⚠️ Error al conectar con SharePoint ({e}). Recurriendo al archivo local...")

    if local_path and os.path.exists(local_path):
        print(f"📂 Cargando archivo local: {local_path}")
        return pd.read_excel(local_path, sheet_name='PLAN.MANT', skiprows=9)
    else:
        raise FileNotFoundError("No se pudo obtener el Excel desde la nube ni se encontró el archivo local.")


def extraer_datos_mantenimiento(df, identificador_maquina):
    """
    Búsqueda híbrida: busca por código entre paréntesis o por coincidencia de texto en el nombre.
    """
    fila_encontrada = None
    identificador_clean = str(identificador_maquina).strip().upper()
    
    for idx, row in df.iterrows():
        nombre_excel = str(row.iloc[1]).strip().upper()
        if identificador_clean in nombre_excel:
            fila_encontrada = row
            break
            
    if fila_encontrada is None:
        raise ValueError(f"No se encontró la máquina '{identificador_maquina}' en la hoja de Excel.")
    
    # 1. Orden de Trabajo
    obs = str(fila_encontrada.iloc[11]) if pd.notna(fila_encontrada.iloc[11]) else ""
    ot_match = re.search(r'OSS\d+', obs)
    orden_trabajo = ot_match.group(0) if ot_match else "S/N"
    
    # 2. Horas Utilizadas
    horas = fila_encontrada.iloc[10]
    horas_str = f"{horas} HORAS" if pd.notna(horas) else "0 HORAS"
    
    # 3. Fecha de Ejecución
    fecha_val = fila_encontrada.iloc[13]
    if isinstance(fecha_val, (datetime.datetime, pd.Timestamp)):
        meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", 
                 "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
        fecha_str = f"{meses[fecha_val.month - 1]} - {fecha_val.year}"
    else:
        fecha_str = "DICIEMBRE - 2025"
        
    # 4. Técnicos Responsables (Tecnico 1 al 4)
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


def actualizar_plantilla_word(docx_input, output_docx_path, datos):
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


def convertir_a_pdf(docx_path, folder_salida):
    docx_path = Path(docx_path)
    folder_salida = Path(folder_salida)
    
    if sys.platform.startswith('linux'):
        try:
            subprocess.run([
                'libreoffice', '--headless', '--convert-to', 'pdf',
                str(docx_path), '--outdir', str(folder_salida)
            ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            print(f"⚠️ Error LibreOffice: {e}")
    elif sys.platform.startswith('win'):
        try:
            from docx2pdf import convert
            pdf_path = folder_salida / f"{docx_path.stem}.pdf"
            convert(str(docx_path), str(pdf_path))
        except Exception as e:
            print(f"⚠️ Error docx2pdf: {e}")


def procesar_ficha(docx_plantilla, onedrive_link=ONEDRIVE_SHARE_LINK, excel_local=EXCEL_LOCAL_PATH, carpeta_destino=CARPETA_SALIDA):
    Path(carpeta_destino).mkdir(exist_ok=True, parents=True)
    
    doc = docx.Document(docx_plantilla)
    identificador = None
    
    # Intento 1: Extraer código entre paréntesis (ej. 60010)
    for p in doc.paragraphs:
        if "CODIGO DE LA MAQUINA" in p.text:
            match = re.search(r'\((\d+)\)', p.text)
            if match:
                identificador = match.group(1)
                break
            else:
                # Intento 2: Extraer texto después de los dos puntos
                partes = p.text.split(":")
                if len(partes) > 1 and partes[1].strip():
                    identificador = partes[1].strip()
                    break
                    
    if not identificador:
        identificador = "60010"
        
    print(f"\n🔍 Buscando información para: {identificador}")
    
    df_excel = obtener_dataframe_excel(onedrive_link, excel_local)
    datos = extraer_datos_mantenimiento(df_excel, identificador)
    
    nombre_salida = f"{datos['orden_trabajo']} ({identificador})_ACTUALIZADO"
    ruta_docx = Path(carpeta_destino) / f"{nombre_salida}.docx"
    
    actualizar_plantilla_word(docx_plantilla, ruta_docx, datos)
    convertir_a_pdf(ruta_docx, carpeta_destino)
    
    print(f" Procesamiento finalizado para {identificador}.")
    return ruta_docx, Path(carpeta_destino) / f"{nombre_salida}.pdf"


if __name__ == "__main__":
    PLANTILLA = "OSS28025 (APILADOR ELÉCTRICO FN 1225 60010).docx"
    if os.path.exists(PLANTILLA):
        procesar_ficha(PLANTILLA)