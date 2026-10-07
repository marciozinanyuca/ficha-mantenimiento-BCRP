# Usar una imagen ligera oficial de Python
FROM python:3.11-slim

# Crear directorio de trabajo
WORKDIR /app

# Copiar primero solo los requerimientos para aprovechar el caché de Docker
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el resto del código
COPY . .

# Exponer el puerto donde corre FastAPI
EXPOSE 8000

# Comando de arranque inmutable
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]