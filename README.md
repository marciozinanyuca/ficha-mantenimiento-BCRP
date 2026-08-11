# 📋 Sistema de Automatización de Fichas de Mantenimiento
Aplicación Full-Stack diseñada para optimizar, automatizar y cruzar información entre las matrices anuales de mantenimiento (Excel) y las plantillas oficiales de fichas técnicas (`.docx`).

## 🛠️ Tecnologías Utilizadas

### Frontend
*   **Next.js** (App Router / React)
*   **TypeScript**
*   **Tailwind CSS** (Diseño responsivo y moderno)
*   **Lucide React** (Iconografía)

### Backend
*   **Node.js & Express**
*   **Multer** (Gestión de archivos en memoria)
*   **XLSX** (Lectura y procesamiento de matrices Excel)
*   **Mammoth** (Extracción de texto y auto-detección en archivos Word)

---

## 📂 Estructura del Proyecto

El repositorio está organizado en dos módulos principales:

```text
/
├── backend/            # Servidor API REST en Node.js / Express
│   ├── server.js       # Configuración de endpoints y rutas
│   ├── automatizar_fichas.js # Lógica de negocio y manipulación de documentos
│   └── package.json
│
└── frontend/           # Interfaz de usuario en Next.js
    ├── app/
    │   └── page.tsx    # Panel de control principal (Dashboard)
    ├── package.json
    └── .env.local      # Variables de entorno (Desarrollo)

```

## 🚀 Guía de Instalación y Ejecución Local

Prerrequisitos
### 1. Tener instalado Node.js (versión 18 o superior recomendada).
### 2. Un editor de código (se recomienda VS Code).

1. Configurar y arrancar el Backend
Abre una terminal y navega a la carpeta del backend:
cd backend
npm install
node server.js

(El servidor correrá por defecto en el puerto 8000.)


2. Configurar y arrancar el Frontend
Abre otra pestaña en la terminal y navega a la carpeta del frontend:
cd frontend
npm install


Crea un archivo llamado .env.local en la raíz de la carpeta frontend y define la URL de tu API local:
NEXT_PUBLIC_API_URL=http://localhost:8000

Inicia el entorno de desarrollo:
npm run dev


## ⚙️ Flujo de Funcionamiento
### 1. Carga de Excel: El usuario sube la matriz del Plan Anual de Mantenimiento (.xlsx). El sistema procesa la hoja PLAN.MANT para extraer el catálogo de máquinas.
### 2. Carga de Word: El usuario sube la plantilla de la ficha del año anterior (.docx). El sistema analiza el contenido, detecta la máquina y busca sus registros asociados en el Excel.
### 3. Validación: El sistema autocompleta los campos operativos (Orden de Trabajo, Fecha, Horas Reales, Técnicos Responsables y Diagnóstico de Falla). El usuario puede editar, cambiar la máquina (selector desplegable) o ajustar datos manualmente.
### 4. Generación: Al hacer clic en descargar, el backend procesa los cambios en memoria RAM y genera la nueva ficha en formato .docx.


##☁️ Despliegue en Producción
###Frontend: Ideal para ser alojado en Vercel conectado a tu repositorio de GitHub. Recuerda configurar la variable de entorno NEXT_PUBLIC_API_URL apuntando a la URL pública de tu backend.
###Backend: Puede ser desplegado en servicios de contenedores o plataformas como Render o Railway.


