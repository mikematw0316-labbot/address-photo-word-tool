FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8765 \
    DATA_DIR=/tmp/address-photo-tool \
    OCR_LANG=chi_tra+eng

RUN apt-get update && apt-get install -y --no-install-recommends \
      tesseract-ocr \
      tesseract-ocr-chi-tra \
      tesseract-ocr-eng \
      fonts-noto-cjk \
      libheif1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py ./
COPY static ./static
COPY templates ./templates

RUN useradd --create-home --uid 10001 appuser \
    && mkdir -p /tmp/address-photo-tool \
    && chown -R appuser:appuser /tmp/address-photo-tool /app

USER appuser
EXPOSE 8765
CMD ["python", "app.py"]
