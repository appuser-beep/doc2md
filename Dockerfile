FROM python:3.12-slim-bookworm

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    ffmpeg \
    exiftool \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY converter.py cleanup.py excel_convert.py ipynb_convert.py zip_convert.py \
    llm_settings.py azure_settings.py advanced_settings.py cli.py ./

ENV EXIFTOOL_PATH=/usr/bin/exiftool

WORKDIR /data
VOLUME ["/data"]

ENTRYPOINT ["python", "/app/cli.py"]
CMD ["--help"]
