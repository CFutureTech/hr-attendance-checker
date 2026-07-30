# syntax=docker/dockerfile:1

FROM alpine:3.20 AS static-assets
WORKDIR /site

COPY index.html ./

FROM nginx:1.27-alpine AS production

COPY nginx.conf.template /etc/nginx/templates/default.conf.template
COPY --from=static-assets /site/ /usr/share/nginx/html/

EXPOSE 80

HEALTHCHECK --interval=30s --timeout=3s --start-period=10s --retries=3 \
  CMD wget -qO- http://127.0.0.1/ >/dev/null || exit 1
