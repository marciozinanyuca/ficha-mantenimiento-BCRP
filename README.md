# 📋 Sistema de Automatización de Fichas de Mantenimiento

**Banco Central de Reserva del Perú (BCRP)**  
*Gerencia de Gestión del Circulante — Casa Nacional de Moneda (CNM)*

---

## 📖 Descripción
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
