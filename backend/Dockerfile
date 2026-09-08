# Usamos una versión ligera de Node.js
FROM node:18-alpine

# Creamos el directorio de trabajo
WORKDIR /app

# Copiamos los archivos de dependencias e instalamos
COPY package*.json ./
RUN npm install

# Copiamos todo el resto del código
COPY . .

# Exponemos el puerto que usa tu servidor (asumo que es el 8000)
EXPOSE 8000

# Comando para iniciar la aplicación
CMD ["node", "server.js"]