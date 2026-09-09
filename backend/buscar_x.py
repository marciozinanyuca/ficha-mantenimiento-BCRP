import zipfile
import re

# Asegúrate de colocar el nombre exacto de tu archivo Word
ruta_archivo = "OSS18026 (APILADOR ELECTRICO REDDOT ES20-WA ACUÑACIÓN (60023)).docx"

try:
    with zipfile.ZipFile(ruta_archivo, 'r') as docx_zip:
        xml = docx_zip.read('word/document.xml').decode('utf-8')
        
        # Buscamos la X flotante y capturamos su entorno XML
        matches = re.finditer(r'.{0,200}<w:t>X</w:t>.{0,200}', xml, re.IGNORECASE)
        
        print("\n" + "="*60)
        print("CÓDIGO XML DE LA 'X' FLOTANTE:")
        print("="*60)
        for m in matches:
            print(m.group(0))
            print("-" * 60)
        print("="*60 + "\n")
except Exception as e:
    print("Error:", e)