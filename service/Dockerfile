FROM python:3.12-slim
RUN apt-get update && apt-get install -y --no-install-recommends poppler-utils libreoffice-calc && rm -rf /var/lib/apt/lists/*
WORKDIR /app/service
COPY service/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY service/*.py ./
RUN useradd -m -u 1000 app && mkdir -p /app/runs && chown app /app/runs
USER app
ENV APP_BASE=/app TEMPLATE=/app/template/template.xlsx
EXPOSE 8000
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
