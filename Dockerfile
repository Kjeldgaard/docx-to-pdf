FROM alpine:3.14

RUN apt -qq update && \
    apt -qq install -y --no-install-recommends libreoffice

COPY docker-entrypoint.sh /usr/local/bin/

ENTRYPOINT ["docker-entrypoint.sh"]
