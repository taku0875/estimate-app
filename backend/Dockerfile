FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
# root では動かさない
RUN useradd --create-home --uid 10001 app && chown -R app /app
USER app
# フロントは別ホスティングにする前提(API だけを配信)
ENV SERVE_FRONTEND=0 ENABLE_CATALOG_WRITE=0
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}"]
