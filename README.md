📋 Sistema de Automatización de Fichas de Mantenimiento

Aplicación Full-Stack diseñada para optimizar, automatizar y cruzar información entre las matrices anuales de mantenimiento (Excel) y las plantillas oficiales de fichas técnicas (.docx).

🛠️ Tecnologías Utilizadas
Frontend
Next.js (App Router / React)
TypeScript
Tailwind CSS (Diseño responsivo y moderno)
Lucide React (Iconografía)
Backend
Node.js & Express
Multer (Gestión de archivos en memoria)
XLSX (Lectura y procesamiento de matrices Excel)
Mammoth (Extracción de texto y auto-detección en archivos Word)
📂 Estructura del Proyecto

El repositorio está organizado en dos módulos principales:

/
├── backend/                       # Servidor API REST en Node.js / Express
│   ├── server.js                  # Configuración de endpoints y rutas
│   ├── automatizar_fichas.js      # Lógica de negocio y manipulación de documentos
│   └── package.json
│
└── frontend/                      # Interfaz de usuario en Next.js
    ├── app/
    │   └── page.tsx               # Panel de control principal (Dashboard)
    ├── package.json
    └── .env.local                 # Variables de entorno (Desarrollo)

🚀 Guía de Instalación y Ejecución Local
Prerrequisitos

Antes de comenzar, asegúrate de tener instalado:

Node.js versión 18 o superior (recomendada).
Un editor de código. Se recomienda Visual Studio Code.
1. Configurar y arrancar el Backend

Abre una terminal y navega a la carpeta del backend:

cd backend
npm install
node server.js


El servidor correrá por defecto en:

http://localhost:8000

2. Configurar y arrancar el Frontend

Abre otra pestaña en la terminal y navega a la carpeta del frontend:

cd frontend
npm install


Crea un archivo llamado .env.local en la raíz de la carpeta frontend y define la URL de tu API local:

NEXT_PUBLIC_API_URL=http://localhost:8000


Inicia el entorno de desarrollo:

npm run dev


Una vez iniciado, podrás acceder al frontend desde la URL proporcionada por Next.js, normalmente:

http://localhost:3000

⚙️ Flujo de Funcionamiento
1. 📊 Carga de Excel

El usuario sube la matriz del Plan Anual de Mantenimiento (.xlsx).

El sistema procesa la hoja PLAN.MANT para extraer el catálogo de máquinas y la información asociada.

2. 📄 Carga de Word

El usuario sube la plantilla de la ficha del año anterior (.docx).

El sistema analiza automáticamente su contenido, detecta la máquina correspondiente y busca sus registros asociados en la matriz de Excel.

3. ✅ Validación

El sistema autocompleta los principales campos operativos:

Orden de Trabajo
Fecha
Horas Reales
Técnicos Responsables
Diagnóstico de Falla

El usuario puede revisar y modificar la información antes de generar el documento. También puede:

Cambiar la máquina mediante un selector desplegable.
Editar los datos detectados automáticamente.
Ajustar manualmente cualquier información necesaria.
4. 📥 Generación

Al hacer clic en Descargar, el backend procesa los cambios en memoria RAM y genera la nueva ficha de mantenimiento en formato .docx.

☁️ Despliegue en Producción
Frontend

El frontend puede ser alojado en plataformas como Vercel, conectado directamente al repositorio de GitHub.

Durante el despliegue, se debe configurar la variable de entorno:

NEXT_PUBLIC_API_URL=https://URL-PUBLICA-DE-TU-BACKEND


Esta variable debe apuntar a la URL pública donde se encuentre desplegada la API del backend.

Backend

El backend puede desplegarse en servicios de contenedores o plataformas como:

Render
Railway
Otros servicios compatibles con aplicaciones Node.js/Express.

Una vez desplegado el backend, se debe actualizar NEXT_PUBLIC_API_URL en el frontend para que apunte a la API de producción.

📌 Resumen

El sistema permite automatizar el proceso de generación de fichas de mantenimiento, reduciendo el trabajo manual y facilitando la integración de información proveniente de matrices Excel y plantillas Word.

El flujo general es:

Matriz Excel (.xlsx)
        │
        ▼
   Procesamiento
        │
        ├──────────────┐
        ▼              │
Catálogo de       Plantilla Word
 máquinas             (.docx)
        │              │
        └──────┬───────┘
               ▼
          Cruce de datos
               │
               ▼
           Validación
               │
               ▼
       Generación de ficha
               │
               ▼
          Archivo .docx
