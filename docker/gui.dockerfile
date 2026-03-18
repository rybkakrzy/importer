FROM docker.io/library/node:22-alpine AS build
WORKDIR /app

RUN apk upgrade --no-cache openssl libavif expat libpng curl libxml2

COPY D2GuiViewerEditor/package.json D2GuiViewerEditor/package-lock.json ./
RUN npm ci --prefer-offline

COPY D2GuiViewerEditor/ .
RUN npm run build -- --configuration production

FROM docker.io/library/nginx:1.27-alpine AS runtime

RUN apk upgrade --no-cache openssl libavif expat libpng curl libxml2

RUN rm -rf /usr/share/nginx/html/* /etc/nginx/conf.d/default.conf

COPY D2GuiViewerEditor/nginx.conf /etc/nginx/conf.d/default.conf

COPY --from=build /app/dist/d2-gui-viewereditor/browser /usr/share/nginx/html

EXPOSE 80

CMD ["nginx", "-g", "daemon off;"]
