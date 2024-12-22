FROM ubuntu:latest
LABEL authors="edemw"

ENTRYPOINT ["top", "-b"]