# Usamos una imagen oficial ligera de Python
FROM python:3.11-slim

# Directorio de trabajo dentro del contenedor
WORKDIR /app

# Instalamos dependencias del sistema necesarias si Pandas o python-docx lo requieren
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copiamos e instalamos los requerimientos de Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiamos el resto del código del backend
COPY . .

# Azure Container Apps inyecta la variable de entorno PORT (por defecto usaremos 8000)
ENV PORT=8000
EXPOSE 8000

# Comando para arrancar FastAPI con Uvicorn
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT}"]