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

- **Node.js & Express**
- **Multer** (Gestión de archivos en memoria)
- **XLSX** (Lectura y procesamiento de matrices Excel)
- **Mammoth** (Extracción de texto y auto-detección en archivos Word)

---

## 📂 Estructura del Proyecto

El repositorio está organizado en dos módulos principales:

```text
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
```

---

## 🚀 Guía de Instalación y Ejecución Local

### Prerrequisitos

Antes de comenzar, asegúrate de tener instalado:

- **Node.js** versión 18 o superior recomendada.
- Un editor de código. Se recomienda **Visual Studio Code**.

### 1. Configurar y arrancar el Backend

Abre una terminal y navega a la carpeta del backend:

```bash
cd backend
npm install
node server.js
```

El servidor correrá por defecto en el puerto `8000`.

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

## ☁️ Ejecución en GitHub Codespaces (Opcional)

Si estás ejecutando este proyecto en un GitHub Codespaces, sigue estos pasos para garantizar que el frontend se comunique correctamente con el backend sin bloqueos de red:

### 1. Configurar los Puertos como Públicos

Ve a la pestaña Puertos en la parte inferior de tu panel de VS Code (junto a la terminal).
Verás los puertos 3000 (Frontend) y 8000 (Backend).
Haz clic derecho en cada uno de ellos y cambia su visibilidad de Privado a Público.

### 2. Arrancar el Backend

Abre una terminal, navega a la carpeta del backend, instala dependencias e inicia el servidor:

```bash
cd backend
npm install
node server.js
```

El servidor correrá en el puerto 8000.

### 3. Configurar y Arrancar el Frontend

Abre otra pestaña en la terminal y navega a la carpeta del frontend:

```bash
cd frontend
npm install
```

Crea un archivo .env.local en la raíz de la carpeta frontend apuntando a la URL pública de tu backend en Codespaces (o directamente a http://localhost:8000 si estás en local tradicional):

```env
NEXT_PUBLIC_API_URL=https://<tu-codigo-codespace>-8000.app.github.dev
```

Para evitar problemas de congelamiento por WebSockets (HMR) típicos de Next.js en contenedores en la nube, se recomienda compilar y ejecutar en Modo Producción:

```bash
npm run build
npx next start -H 0.0.0.0
```

El frontend correrá en el puerto 3000 y estará accesible mediante su URL pública correspondiente.

---

## ⚙️ Flujo de Funcionamiento

### 1. 📊 Carga de Excel

El usuario sube la matriz del **Plan Anual de Mantenimiento** (`.xlsx`).

El sistema procesa la hoja `PLAN.MANT` para extraer el catálogo de máquinas.

### 2. 📄 Carga de Word

El usuario sube la plantilla de la ficha del año anterior (`.docx`).

El sistema analiza el contenido, detecta la máquina y busca sus registros asociados en el Excel.

### 3. ✅ Validación

El sistema autocompleta los siguientes campos operativos:

- **Orden de Trabajo**
- **Fecha**
- **Horas Reales**
- **Técnicos Responsables**
- **Diagnóstico de Falla**

El usuario puede:

- Editar los datos autocompletados.
- Cambiar la máquina mediante un selector desplegable.
- Ajustar manualmente cualquier dato antes de generar la ficha.

### 4. 📥 Generación

Al hacer clic en **Descargar**, el backend procesa los cambios en memoria RAM y genera la nueva ficha en formato `.docx`.

---

## ☁️ Despliegue en Producción

### Frontend

Ideal para ser alojado en **Vercel**, conectado al repositorio de GitHub.

Recuerda configurar la variable de entorno:

```env
NEXT_PUBLIC_API_URL=https://URL-PUBLICA-DE-TU-BACKEND
```

Esta variable debe apuntar a la URL pública de tu backend.

### Backend

Puede ser desplegado en servicios de contenedores o plataformas como:

- **Render**
- **Railway**
- Otros servicios compatibles con aplicaciones **Node.js / Express**

Una vez desplegado el backend, actualiza la variable `NEXT_PUBLIC_API_URL` del frontend para que apunte a la URL pública correspondiente.

---

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
             │       │       con Mammoth       │
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

## 🏗️ Arquitectura del Proyecto

El sistema utiliza una arquitectura **Full-Stack** separada en dos aplicaciones principales:

### Frontend

Construido con **Next.js**, encargado de:

- Interfaz de usuario.
- Carga de archivos.
- Selección de máquinas.
- Visualización y edición de información.
- Comunicación con la API.
- Descarga de las fichas generadas.

### Backend

Construido con **Node.js y Express**, encargado de:

- Recibir archivos mediante la API.
- Procesar matrices Excel.
- Analizar documentos Word.
- Detectar máquinas automáticamente.
- Cruzar información entre Excel y Word.
- Procesar los datos modificados por el usuario.
- Generar los documentos `.docx`.

---

## 📁 Archivos Principales

| Archivo | Descripción |
|---|---|
| `backend/server.js` | Configuración del servidor y endpoints de la API REST. |
| `backend/automatizar_fichas.js` | Lógica principal para el procesamiento y generación de fichas. |
| `frontend/app/page.tsx` | Dashboard principal de la aplicación. |
| `frontend/.env.local` | Variables de entorno utilizadas durante el desarrollo. |
| `backend/package.json` | Dependencias y configuración del backend. |
| `frontend/package.json` | Dependencias y configuración del frontend. |

---

## 🔧 Variables de Entorno

### Desarrollo

El frontend utiliza la siguiente variable de entorno:

```env
NEXT_PUBLIC_API_URL=http://localhost:8000
```

### Producción

En producción debe utilizarse la URL pública del backend:

```env
NEXT_PUBLIC_API_URL=https://URL-PUBLICA-DE-TU-BACKEND
```

> **Nota:** No se deben incluir archivos `.env.local` con información sensible dentro del repositorio.

---

## 🚀 Comandos Rápidos

### Backend

```bash
cd backend
npm install
node server.js
```

### Frontend

```bash
cd frontend
npm install
npm run dev
```

---

## 📌 Características Principales

- 📊 Procesamiento automático de matrices Excel.
- 📄 Lectura y análisis de plantillas Word.
- 🔍 Detección automática de máquinas.
- 🔄 Cruce de información entre Excel y Word.
- ✏️ Edición manual de los datos detectados.
- 📋 Selección de máquinas mediante selector desplegable.
- ⚙️ Autocompletado de información operativa.
- 📥 Generación automática de fichas `.docx`.
- 💾 Procesamiento de archivos directamente en memoria RAM.
- 🌐 Arquitectura separada de frontend y backend.
- ☁️ Preparado para despliegue en plataformas cloud.
