FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# COPY preserva permissões restritivas do repo; o appuser precisa de ler o
# template e os estáticos → a+rX garante leitura sem dar escrita.
RUN useradd --create-home appuser \
    && chmod -R a+rX /app
USER appuser

CMD ["sh", "-c", "gunicorn app:app --bind 0.0.0.0:${PORT:-8000} --timeout 120 --workers 2"]
