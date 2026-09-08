# صورة Docker جاهزة لتشغيل الموقع كاملاً (واجهة + خادم + محرك التحويل)
FROM python:3.11-slim

# تثبيت LibreOffice (محرك التحويل) و Tesseract (للتعرف الضوئي OCR) وأدوات PDF
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-core \
    tesseract-ocr \
    tesseract-ocr-ara \
    tesseract-ocr-fra \
    tesseract-ocr-spa \
    tesseract-ocr-deu \
    poppler-utils \
    fonts-dejavu \
    fonts-noto \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY backend/requirements.txt /app/backend/requirements.txt
RUN pip install --no-cache-dir -r /app/backend/requirements.txt

COPY backend /app/backend
COPY frontend /app/frontend

WORKDIR /app/backend

ENV PORT=8000
EXPOSE 8000

# gunicorn: عدة عمّال، ومهلة كافية لأن LibreOffice قد يستغرق بضع ثوانٍ للملفات الكبيرة
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT} --workers 3 --timeout 180 app:app"]
