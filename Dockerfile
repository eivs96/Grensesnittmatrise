FROM node:20-alpine

RUN addgroup -S appgroup && adduser -S appuser -G appgroup

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev

COPY . .

RUN mkdir -p /app/data && chown -R appuser:appgroup /app/data

ENV NODE_ENV=production
ENV PORT=3000

USER appuser

EXPOSE 3000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD node -e "require('http').get('http://localhost:3000/health', (r) => { if (r.statusCode !== 200) throw new Error(); r.resume(); }).on('error', () => { process.exit(1); })"

CMD ["npm", "start"]
