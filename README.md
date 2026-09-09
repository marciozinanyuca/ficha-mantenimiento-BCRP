# 📋 Sistema de Automatización de Fichas de Mantenimiento

Aplicación **Full-Stack** diseñada para optimizar, automatizar y cruzar información entre las matrices anuales de mantenimiento (Excel) y las plantillas oficiales de fichas técnicas (`.docx`). 

---

## 🛠️ Tecnologías Utilizadas

### Frontend
- **Next.js** (App Router / React)
- **TypeScript**
- **Tailwind CSS** (Diseño responsivo y moderno)
- **Lucide React** (Iconografía)

### Backend
- **Python 3.9+** (Lenguaje principal)
- **FastAPI** (Framework web ultrarrápido)
- **Uvicorn** (Servidor ASGI para producción)
- **Pandas & Openpyxl** (Lectura y cruce de datos en matrices Excel)
- **Python-docx** (Generación milimétrica y lectura de archivos Word)
- **Difflib & Regex** (Inteligencia difusa para coincidencia de máquinas y escaneo de XML)

---

## 📂 Estructura del Proyecto

El repositorio está organizado en dos módulos principales:

```text
/
├── backend/                       # Servidor API REST en Python / FastAPI
│   ├── main.py                    # Lógica principal, rutas y motor de extracción (Súper Cazador)
│   └── requirements.txt           # Dependencias de Python
│
└── frontend/                      # Interfaz de usuario en Next.js
    ├── app/
    │   └── page.tsx               # Panel de control principal (Dashboard)
    ├── package.json
    └── .env.local                 # Variables de entorno (Desarrollo)

```

## 🔄 Flujo General del Sistema

```text
┌─────────────────────────┐
│     Matriz Excel        │
│        (.xlsx)          │
└────────────┬────────────┘
             │
             ▼
┌─────────────────────────┐
│  Procesamiento de Excel │
│       Hoja PLAN.MANT    │
└────────────┬────────────┘
             │
             │
             ▼
┌─────────────────────────┐
│   Catálogo de máquinas  │
└────────────┬────────────┘
             │
             │
             │       ┌─────────────────────────┐
             │       │    Plantilla Word       │
             │       │        (.docx)          │
             │       └────────────┬────────────┘
             │                    │
             │                    ▼
             │       ┌─────────────────────────┐
             │       │ Extracción de contenido │
             │       └────────────┬────────────┘
             │                    │
             └──────────┬─────────┘
                        ▼
             ┌──────────────────────┐
             │     Cruce de datos   │
             └──────────┬───────────┘
                        │
                        ▼
             ┌──────────────────────┐
             │      Validación      │
             │                      │
             │ - Orden de Trabajo   │
             │ - Fecha              │
             │ - Horas Reales       │
             │ - Técnicos           │
             │ - Diagnóstico        │
             └──────────┬───────────┘
                        │
                        ▼
             ┌──────────────────────┐
             │ Generación de ficha  │
             │       (.docx)        │
             └──────────────────────┘
```

---

## 🚀 Guía de Instalación y Ejecución Local

### Prerrequisitos

Antes de comenzar, asegúrate de tener instalado:

* **Node.js** versión 18 o superior (para el frontend).
* **Python** versión 3.9 o superior.
* Un editor de código (ej. **Visual Studio Code**).

### 1. Configurar y arrancar el Backend

Abre una terminal y navega a la carpeta del backend:

```bash
cd backend
pip install -r requirements.txt
python main.py

```

*(Alternativamente, puedes usar `uvicorn main:app --reload`)*. El servidor correrá por defecto en el puerto `8000`.

### 2. Configurar y arrancar el Frontend

Abre otra pestaña en la terminal y navega a la carpeta del frontend:

```bash
cd frontend
npm install

```

Crea un archivo llamado `.env.local` en la raíz de la carpeta `frontend` y define la URL de tu API local:

```env
NEXT_PUBLIC_API_URL=http://localhost:8000

```

Inicia el entorno de desarrollo:

```bash
npm run dev

```

---

## ☁️ Despliegue en Producción

El proyecto está diseñado para desplegarse fácilmente en entornos modernos en la nube.

### Backend (Recomendado en Render)

1. Crea un nuevo **Web Service** conectado a tu rama `main`.
2. Configuración del entorno:
* **Language:** Python 3
* **Build Command:** `pip install -r requirements.txt`
* **Start Command:** `uvicorn main:app --host 0.0.0.0 --port $PORT`


3. Copia la URL pública generada (ej. `https://tu-api.onrender.com`).

### Frontend (Recomendado en Vercel)

1. Importa tu repositorio en Vercel.
2. En la sección **Environment Variables**, añade:
* `NEXT_PUBLIC_API_URL` = `https://tu-api.onrender.com` (La URL pública de Render).


3. Haz clic en **Deploy**.

---

## ⚙️ Flujo de Funcionamiento

### 1. 📊 Carga de Excel

El usuario sube la matriz del **Plan Anual de Mantenimiento** (`.xlsx`). El sistema procesa la hoja `PLAN.MANT` para extraer el catálogo oficial de máquinas.

### 2. 📄 Carga de Word (Extracción Inteligente)

El usuario sube la plantilla de la ficha del año anterior (`.docx`). El backend ejecuta un escaneo de 4 fases (Súper Cazador):

* **Cazador de Tablas:** Detecta si las casillas de verificación fueron dibujadas como celdas.
* **Cazador XML:** Lee el código fuente del documento para identificar *Checkboxes* nativos de Word y símbolos *Wingdings*.
* **Radar Matemático Lineal:** Mide la distancia física de caracteres sueltos (como la "X") ignorando ilusiones ópticas de cuadros flotantes.
* **Fuzzy Matching (IA Difusa):** Identifica a qué máquina del Excel pertenece el archivo analizando intersección de palabras, ignorando códigos extra o errores de tipeo en el nombre.

### 3. ✅ Validación y Edición

El sistema autocompleta:

* **Orden de Trabajo (OT)**, **Fecha**, **Horas Reales**, **Técnicos Responsables** (desde el Excel).
* **Tipo de Mantenimiento**, **Diagnóstico de Falla** y **Nota** (desde el Word).

El usuario puede editar libremente cualquier campo en la interfaz antes de la generación final.

### 4. 📥 Generación

Al hacer clic en **Descargar**, el sistema construye un nuevo documento `.docx` desde cero respetando márgenes milimétricos, separando dinámicamente las tablas operativas (Mecánico / Eléctrico) y resaltando los valores clave.

---

## 📌 Características Principales

* 🐍 **Motor Backend en Python:** Alto rendimiento para manipulación de documentos y procesamiento de datos.
* 🎯 **Fuzzy Matching:** Algoritmo difuso (`difflib`) para cruce inteligente de nombres de máquinas inexactos.
* 🕵️‍♂️ **Extracción Deep XML:** Capacidad de leer estados de *Form Fields*, *Checkboxes* modernos y *Wingdings* que otras librerías ignoran.
* 📐 **Renderizado Dinámico de Tablas:** Creación inteligente de secciones (MECÁNICO, ELÉCTRICO, o MIXTO) según las tareas extraídas.
* 💾 **Gestión en RAM:** Procesamiento de archivos temporales ultrarrápido y seguro.
* 🌐 **Arquitectura Desacoplada:** Frontend en Next.js independiente del backend RESTful en FastAPI.

```

```