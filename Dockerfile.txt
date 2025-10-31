# Usa una base leggera
FROM python:3.10-slim

WORKDIR /app
COPY . .

RUN pip install --no-cache-dir -r requirements.txt

# Streamlit usa la porta 8501
EXPOSE 8501

# Avvio
CMD ["streamlit", "run", "studio_isa_web.py", "--server.port=8501", "--server.address=0.0.0.0"]
