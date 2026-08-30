```dockerfile
FROM python:3.11-slim

# Install LibreOffice and useful fonts
RUN apt-get update && \
    apt-get install -y \
    libreoffice \
    libreoffice-writer \
    libreoffice-draw \
    fonts-liberation \
    fonts-dejavu \
    fonts-noto-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .

EXPOSE 10000

CMD ["gunicorn", "--bind", "0.0.0.0:10000", "--workers", "1", "--timeout", "180", "app:app"]
```
