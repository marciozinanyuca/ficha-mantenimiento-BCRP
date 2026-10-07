# ETAPA 1: Construcción
FROM node:24-alpine AS builder
WORKDIR /app
COPY package*.json ./
RUN npm install
COPY . .

# Permitir que GitHub Actions inyecte la URL de Render o Azure durante la compilación
ARG NEXT_PUBLIC_API_URL
ENV NEXT_PUBLIC_API_URL=$NEXT_PUBLIC_API_URL

RUN npm run build

# ETAPA 2: Producción (Servidor limpio)
FROM node:24-alpine AS runner
WORKDIR /app
ENV NODE_ENV production

# Copiar solo la carpeta standalone y los estáticos necesarios
COPY --from=builder /app/public ./public
COPY --from=builder /app/.next/standalone ./
COPY --from=builder /app/.next/static ./.next/static

EXPOSE 3000
ENV PORT 3000

# Arrancar el servidor empaquetado
CMD ["node", "server.js"]