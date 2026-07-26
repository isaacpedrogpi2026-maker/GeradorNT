FROM python:3.11-slim

# Evita criação de arquivos .pyc e força stdout/stderr sem buffer
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

WORKDIR /app

# Instala pacotes do sistema necessários para compilação e dependências nativas
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    curl \
    libpq-dev \
    && rm -rf /var/lib/apt/lists/*

# Copia e instala as dependências Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia os arquivos do projeto para o container
COPY . .

# Expõe a porta padrão do Streamlit
EXPOSE 8501

# Healthcheck para monitorar se a aplicação Streamlit está saudável
HEALTHCHECK --interval=30s --timeout=10s --start-period=10s --retries=3 \
  CMD curl --fail http://localhost:8501/_stcore/health || exit 1

# Comando de inicialização do Streamlit
ENTRYPOINT ["streamlit", "run", "GeradorNT.py", "--server.port=8501", "--server.address=0.0.0.0"]
