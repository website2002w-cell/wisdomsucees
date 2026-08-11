FROM node:20-alpine AS builder
WORKDIR /app

# Copy package files first for better caching
COPY package.json package-lock.json* ./
COPY tsconfig.json vite.config.ts ./
COPY . .

RUN npm ci && npm run build

FROM node:20-alpine AS runtime
WORKDIR /app
ENV NODE_ENV=production

# Copy built assets and package manifest
COPY --from=builder /app/dist ./dist
COPY --from=builder /app/package.json ./package.json

# Install only production dependencies
RUN npm ci --production

EXPOSE 3000
CMD ["node", "dist/server.mjs"]
