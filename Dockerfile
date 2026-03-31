# ---- Stage 1: Build ----
FROM node:18-slim AS builder

WORKDIR /app

COPY package.json package-lock.json* ./
RUN npm ci

COPY . .
RUN npm run build

# ---- Stage 2: Production ----
FROM node:18-slim AS runner

# Install LibreOffice + fonts for XLSX → PDF conversion
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice \
        fonts-dejavu \
        fonts-liberation \
        fonts-crosextra-carlito \
        fonts-noto-core \
        fontconfig && \
    rm -rf /var/lib/apt/lists/* && \
    fc-cache -f -v

WORKDIR /app

ENV NODE_ENV=production
ENV PORT=8080

# Copy standalone Next.js output
COPY --from=builder /app/.next/standalone ./
COPY --from=builder /app/.next/static ./.next/static
COPY --from=builder /app/public ./public

EXPOSE 8080

CMD ["node", "server.js"]
