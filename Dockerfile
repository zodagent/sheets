# Zod Sheets over HTTP: docker build -t zod-sheets . && docker run -p 8080:8080 -v sheets-data:/data -e MCP_TOKEN=long-secret zod-sheets
# In Zod: Connectors > Add custom > Add from a GitHub link or address, paste https://your-host/mcp and the token.
FROM oven/bun:1-alpine
WORKDIR /app
COPY package.json server.js ./
COPY src ./src
ENV PORT=8080 SHEETS_DIR=/data NODE_ENV=production
VOLUME /data
EXPOSE 8080
HEALTHCHECK CMD wget -qO- http://127.0.0.1:8080/health || exit 1
CMD ["bun", "server.js"]
