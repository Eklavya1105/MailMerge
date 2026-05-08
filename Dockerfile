FROM node:20-slim

RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice-writer && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json ./
RUN npm install --production

COPY . .

RUN mkdir -p uploads output

EXPOSE 3001

CMD ["node", "server.js"]
