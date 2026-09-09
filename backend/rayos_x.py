import zipfile
import re

# Pon aquí el nombre exacto de tu archivo Word de prueba
ruta_archivo = "OSS18026 (APILADOR ELECTRICO REDDOT ES20-WA ACUÑACIÓN (60023)).docx" 

try:
    with zipfile.ZipFile(ruta_archivo, 'r') as docx_zip:
        xml = docx_zip.read('word/document.xml').decode('utf-8')
        
        # Extraemos un bloque de código alrededor de la palabra TIPO II
        match = re.search(r'.{0,350}TIPO II.{0,350}', xml, re.IGNORECASE)
        
        print("\n" + "="*60)
        print("CÓDIGO XML DETECTADO (Copia y pega esto en el chat):")
        print("="*60)
        if match:
            print(match.group(0))
        else:
            print("No se encontró la palabra TIPO II.")
        print("="*60 + "\n")
except Exception as e:
    print("Error al leer el archivo:", e)