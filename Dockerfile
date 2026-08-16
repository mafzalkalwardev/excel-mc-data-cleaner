FROM alpine:3.24
WORKDIR /src
COPY . .
LABEL org.opencontainers.image.source="https://github.com/mafzalkalwardev/excel-mc-data-cleaner"
CMD ["sh", "-c", "echo 'excel-mc-data-cleaner source package' && ls -1"]
