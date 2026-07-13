FROM python:3.13-slim

WORKDIR /app

# Instala as dependencias primeiro (melhor cache de build)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia o codigo e os dados do projeto
COPY app.py .
COPY Produtos_comercializavel_S.xlsx .
COPY imagens/ ./imagens/

EXPOSE 5000

# HOST/PORT podem ser sobrescritos com -e na hora de rodar
ENV HOST=0.0.0.0
ENV PORT=5000

CMD ["python", "app.py"]
