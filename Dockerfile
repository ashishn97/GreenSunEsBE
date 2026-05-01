FROM node:22-bookworm-slim

RUN apt-get update \
  && apt-get install -y --no-install-recommends libreoffice fonts-dejavu fontconfig \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev

COPY . .

ENV NODE_ENV=production
ENV LIBREOFFICE_PATH=soffice
ENV OUTPUT_DIR=/tmp/gse-output

EXPOSE 10000

CMD ["npm", "start"]
