FROM node:20-alpine

RUN apk add --no-cache python3 make g++

RUN addgroup -S appgroup && adduser -S appuser -G appgroup

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev

COPY . .

RUN mkdir -p /app/data && chown -R appuser:appgroup /app/data

ENV NODE_ENV=production

USER appuser

HEALTHCHECK --interval=30s --timeout=10s --start-period=30s --retries=5 \
    CMD node -e "const port = process.env.PORT || 8080; require('http').get('http://localhost:' + port + '/health', (r) => { if (r.statusCode !== 200) throw new Error(); r.resume(); }).on('error', () => { process.exit(1); })"

CMD ["npm", "start"]
