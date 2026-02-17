FROM python:3.11.7-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 10000

CMD ["gunicorn", "ga4_api:app", "--bind", "0.0.0.0:10000", "--workers", "1", "--timeout", "120"]
