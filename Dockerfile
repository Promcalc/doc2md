FROM python:3.10-slim

# Устанавливаем LibreOffice для конвертации .doc
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app/ ./app/

# Создаём пользователя для безопасности (опционально)
RUN useradd -m -u 1000 appuser && chown -R appuser /app
USER appuser

# Порт по умолчанию
ENV PORT=8000

EXPOSE $PORT

# Запускаем сервер через uvicorn
CMD uvicorn app.main:app --host 0.0.0.0 --port $PORT