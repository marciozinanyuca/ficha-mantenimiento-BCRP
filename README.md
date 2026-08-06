# ficha-mantenimiento-cnm

-- Instalar LibreOffice en el Codespac
sudo apt-get update && sudo apt-get install -y libreoffice

pip install -r requirements.txt


ficha-mantenimiento-cnm/
├── backend/                       # API en Python (FastAPI o Flask)
│   ├── main.py                    # Endpoint POST /api/generar-ficha
│   ├── automatizar_fichas.py      # Lógica de procesamiento
│   └── requirements.txt
├── frontend/                      # Aplicación Next.js
│   ├── src/pages/ (o app/)        # Interfaz de usuario para subir Word/Excel y descargar PDF
│   └── package.json
└── Fichas/                        # Almacenamiento temporal de archivos