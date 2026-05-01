FROM node:18-bullseye

RUN apt-get update \
  && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    fonts-dejavu \
    fontconfig \
  && apt-get clean \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm install --omit=dev

COPY . .

ENV NODE_ENV=production
ENV PORT=10000
ENV LIBREOFFICE_PATH=soffice

EXPOSE 10000

CMD ["node", "server.js"]