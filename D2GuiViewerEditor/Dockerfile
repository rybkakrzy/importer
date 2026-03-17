# ── Stage 1: build ──────────────────────────────────────────────────────────
FROM docker.io/library/node:22-alpine AS build
WORKDIR /app

# Zainstaluj zależności (cache warstw — kopiuj tylko manifesty)
COPY package.json package-lock.json ./
RUN npm ci --prefer-offline

# Kopiuj resztę kodu i zbuduj produkcyjnie
COPY . .
RUN npm run build -- --configuration production

# ── Stage 2: runtime (nginx) ─────────────────────────────────────────────────
FROM docker.io/library/nginx:1.27-alpine AS runtime

# Usuń domyślną konfigurację nginx
RUN rm -rf /usr/share/nginx/html/* /etc/nginx/conf.d/default.conf

COPY nginx.conf /etc/nginx/conf.d/default.conf

# Angular 17+ generuje pliki w dist/<nazwa>/browser
COPY --from=build /app/dist/d2-gui-viewereditor/browser /usr/share/nginx/html

EXPOSE 80

CMD ["nginx", "-g", "daemon off;"]
