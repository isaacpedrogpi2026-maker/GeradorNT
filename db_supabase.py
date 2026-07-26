import os
import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text
from supabase import create_client, Client


def get_supabase_credentials():
    """
    Recupera as credenciais do Supabase das configurações do Streamlit (secrets.toml)
    ou de variáveis de ambiente.
    """
    supabase_url = None
    supabase_key = None
    db_url = None

    # Tenta ler do st.secrets (local ou Streamlit Cloud)
    if hasattr(st, "secrets"):
        supabase_url = st.secrets.get("SUPABASE_URL", None)
        supabase_key = st.secrets.get("SUPABASE_KEY", None)
        db_url = st.secrets.get("SUPABASE_DB_URL", None)

    # Fallback para variáveis de ambiente (Docker env)
    if not supabase_url:
        supabase_url = os.environ.get("SUPABASE_URL", "")
    if not supabase_key:
        supabase_key = os.environ.get("SUPABASE_KEY", "")
    if not db_url:
        db_url = os.environ.get("SUPABASE_DB_URL", "")

    return {
        "url": supabase_url,
        "key": supabase_key,
        "db_url": db_url
    }


@st.cache_resource
def get_db_engine():
    """
    Cria e armazena em cache uma conexão SQLAlchemy com o banco PostgreSQL do Supabase.
    """
    creds = get_supabase_credentials()
    db_url = creds.get("db_url")

    if not db_url:
        return None

    try:
        engine = create_engine(db_url, pool_pre_ping=True, pool_size=5, max_overflow=10)
        return engine
    except Exception as e:
        st.error(f"Erro ao criar conexão com o banco de dados Supabase: {e}")
        return None


def get_supabase_client() -> Client | None:
    """
    Cria um cliente da API do Supabase (para operações diretas REST/Storage).
    """
    creds = get_supabase_credentials()
    url = creds.get("url")
    key = creds.get("key")

    if url and key:
        try:
            return create_client(url, key)
        except Exception as e:
            st.error(f"Erro ao inicializar o cliente Supabase API: {e}")
            return None
    return None


def test_supabase_connection():
    """
    Testa se a conexão com o Supabase PostgreSQL está funcionando corretamente.
    """
    engine = get_db_engine()
    if engine is None:
        return False, "URL do banco de dados (SUPABASE_DB_URL) não configurada."

    try:
        with engine.connect() as conn:
            result = conn.execute(text("SELECT 1;"))
            row = result.fetchone()
            if row and row[0] == 1:
                return True, "Conexão com o Supabase estabelecida com sucesso!"
            return False, "Resposta inesperada do banco de dados."
    except Exception as e:
        return False, f"Falha na conexão com Supabase: {e}"


def load_table_from_supabase(table_name: str) -> pd.DataFrame:
    """
    Carrega uma tabela do Supabase diretamente para um DataFrame pandas.
    """
    engine = get_db_engine()
    if engine is None:
        raise ValueError("Conexão com o Supabase não configurada.")

    try:
        query = f'SELECT * FROM "{table_name}"'
        df = pd.read_sql(query, engine)
        return df
    except Exception as e:
        st.error(f"Erro ao ler a tabela '{table_name}' do Supabase: {e}")
        return pd.DataFrame()


def save_dataframe_to_supabase(df: pd.DataFrame, table_name: str, if_exists: str = "replace") -> bool:
    """
    Salva um DataFrame pandas como uma tabela no Supabase.
    `if_exists`: 'fail', 'replace' ou 'append'.
    """
    engine = get_db_engine()
    if engine is None:
        st.error("Conexão com o Supabase não está configurada.")
        return False

    try:
        # Normaliza nomes de colunas para minúsculas sem espaços para convenção SQL
        df_to_save = df.copy()
        df_to_save.to_sql(table_name, engine, if_exists=if_exists, index=False, method="multi")
        st.success(f"Tabela '{table_name}' salva no Supabase com sucesso!")
        return True
    except Exception as e:
        st.error(f"Erro ao salvar a tabela '{table_name}' no Supabase: {e}")
        return False


def upload_excel_dict_to_supabase(excel_data_dict: dict) -> bool:
    """
    Envia todas as abas de um dicionário de DataFrames Excel para tabelas separadas no Supabase.
    """
    engine = get_db_engine()
    if engine is None:
        return False

    success = True
    for sheet_name, df in excel_data_dict.items():
        # Limpa o nome da aba para ser compatível com tabela SQL
        clean_table_name = sheet_name.strip().lower().replace(" ", "_")
        clean_table_name = "".join(c for c in clean_table_name if c.isalnum() or c == "_")
        
        ok = save_dataframe_to_supabase(df, clean_table_name, if_exists="replace")
        if not ok:
            success = False

    return success


@st.cache_data(ttl=300)
def load_all_sheets_from_supabase() -> dict:
    """
    Carrega todas as tabelas salvas no Supabase e as converte no dicionário de abas esperado pelo GeradorNT.
    """
    engine = get_db_engine()
    if engine is None:
        return {}

    table_mapping = {
        "segreg_por_município": "Segreg por município",
        "segreg_por_municipio": "Segreg por município",
        "volumes_por_município": "Volumes por município",
        "volumes_por_municipio": "Volumes por município",
        "abastecimento": "Abastecimento",
        "população_por_município": "População por município",
        "populacao_por_municipio": "População por município",
        "tv_tspe": "TV+TSPE",
        "tvtspe": "TV+TSPE",
        "investimentos": "Investimentos"
    }

    loaded_dict = {}
    try:
        from sqlalchemy import inspect
        inspector = inspect(engine)
        existing_tables = inspector.get_table_names()

        for raw_table in existing_tables:
            clean_name = raw_table.lower().strip()
            sheet_name = table_mapping.get(clean_name, raw_table)
            df = load_table_from_supabase(raw_table)
            if not df.empty:
                loaded_dict[sheet_name] = df
        return loaded_dict
    except Exception as e:
        st.error(f"Erro ao recuperar tabelas do Supabase: {e}")
        return {}

