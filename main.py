import os
import re
import unicodedata
import io
import pandas as pd
from sqlalchemy import create_engine, text, inspect
from sqlalchemy.dialects.postgresql import UUID
from dotenv import load_dotenv
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware

# ── Carrega .env e Conexão ────────────────────────────────────────
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
DATABASE_URL = os.getenv("DATABASE_URL")

if not DATABASE_URL:
    raise RuntimeError("❌ ERRO: DATABASE_URL não encontrada no .env")

engine = create_engine(DATABASE_URL)

# ── Setup da FastAPI ──────────────────────────────────────────────
app = FastAPI(title="API Importação Iniflex")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Mapeamento exato: coluna Excel (após limpeza) → coluna do banco
#
# Colunas do Excel que NÃO aparecem aqui são descartadas silenciosamente.

COLUNAS_PROJETOS = {
    'cliente':          'cliente',
    'portifolio':       'portifolio',
    'projeto':          'projeto',
    'periodo_inicio':   'periodo_inicio',
    'periodo_fim':      'periodo_fim',
    'lider':            'lider',
    'horas_contrato':   'horas_contrato',
    'horas_utilizada':  'horas_utilizada',
}

COLUNAS_FASES = {
    'atividades':   'atividades',
    'escopo':       'escopo',
    'data':         'data',
    'datafim':      'datafim',
    'recurso':      'recurso',
    'concluido':    'concluido',
    'situacao':     'situacao',
    'comentario':   'comentario',
    'area':         'area',
    'fase':         'fase',
}

COLUNAS_AREAS = {
    'area':           'area',
    'status':         'status',
    'prazo':          'prazo',
    'escopo':         'escopo',
    'acao_requerida': 'acao_requerida',
    'fase':           'fase',
}

COLUNAS_PONTOS_ATENCAO = {
    'indicado_por_area': 'indicado_por_area',
    'descricao':         'descricao',
    'situacao':          'situacao',
    'probalidade':       'probabilidade',   # typo no Excel → nome correto no banco
    'impacto':           'impacto',
}

COLUNAS_RISCOS = {
    'indicado_por_area': 'indicado_por_area',
    'detalhes':          'detalhes',
    'fase':              'fase',
    'probalidade':       'probabilidade',   # typo no Excel → nome correto no banco
    'impacto':           'impacto',
}

ABAS_MAPEAMENTO = {
    #  sheet_name        tabela_banco        mapeamento_colunas
    'ATIVIDADE':      ('fases',           COLUNAS_FASES),
    'AREAS':          ('areas',           COLUNAS_AREAS),
    'PONTOS_ATENCAO': ('pontos_atencao',  COLUNAS_PONTOS_ATENCAO),
    'RISCOS':         ('riscos',          COLUNAS_RISCOS),
}

# ── Helpers ───────────────────────────────────────────────────────
def limpar_nome_coluna(nome: str) -> str:
    nome = str(nome).strip()
    nome = unicodedata.normalize('NFD', nome)
    nome = ''.join(c for c in nome if unicodedata.category(c) != 'Mn')
    nome = nome.lower()
    nome = re.sub(r'[\s\.\-\/]+', '_', nome)
    nome = re.sub(r'[^\w]', '', nome)
    nome = re.sub(r'_+', '_', nome)
    nome = nome.strip('_')
    if nome and nome[0].isdigit():
        nome = f'col_{nome}'
    return nome


def normalizar_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    """Corrige incompatibilidade pandas 2.0 StringDtype → SQLAlchemy."""
    for col in df.columns:
        if isinstance(df[col].dtype, pd.StringDtype):
            df[col] = df[col].astype(object)
        elif pd.api.types.is_datetime64_any_dtype(df[col].dtype):
            df[col] = pd.to_datetime(df[col]).dt.tz_localize(None)
    return df


def selecionar_e_renomear(df: pd.DataFrame, mapeamento: dict) -> pd.DataFrame:
    """
    Mantém só as colunas presentes no mapeamento e as renomeia para os
    nomes exatos do banco. Colunas ausentes no Excel são ignoradas.
    """
    presentes = {k: v for k, v in mapeamento.items() if k in df.columns}
    return df[list(presentes.keys())].rename(columns=presentes)


def ler_aba(file_buffer: io.BytesIO, sheet_name: str, **kwargs) -> pd.DataFrame | None:
    """Lê uma aba do Excel; retorna None se não existir ou estiver vazia."""
    try:
        file_buffer.seek(0)
        df = pd.read_excel(file_buffer, sheet_name=sheet_name, **kwargs)
        return df if not df.empty else None
    except Exception:
        return None


# ── Endpoint de Importação ────────────────────────────────────────
@app.post("/api/importar")
async def importar_dados(
    arquivo: UploadFile = File(...),
    owner_id: str = Form(...),
):
    if not arquivo.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="O arquivo deve ser Excel (.xlsx ou .xls)")

    try:
        contents = await arquivo.read()
        file_buffer = io.BytesIO(contents)

        tabelas_existentes = set(inspect(engine).get_table_names())

        with engine.begin() as conn:

            # ── 1. NOME DO PROJETO ────────────────────────────────
            df_projeto_raw = ler_aba(file_buffer, 'PROJETO', header=None, usecols="A:B")
            if df_projeto_raw is None:
                raise HTTPException(status_code=400, detail="Aba 'PROJETO' não encontrada ou vazia.")

            df_projeto = df_projeto_raw.set_index(0).T
            df_projeto.columns = [limpar_nome_coluna(c) for c in df_projeto.columns]
            df_projeto = df_projeto.reset_index(drop=True)

            if 'projeto' not in df_projeto.columns:
                raise HTTPException(status_code=400, detail="Coluna 'projeto' não encontrada na aba PROJETO.")

            nome_do_projeto = str(df_projeto['projeto'].iloc[0]).strip()

            # ── 2. LIMPEZA SEGURA ─────────────────────────────────
            for tabela in ['fases', 'areas', 'pontos_atencao', 'riscos', 'objetivos']:
                if tabela in tabelas_existentes:
                    conn.execute(
                        text(f"DELETE FROM {tabela} WHERE projeto_vinculo = :projeto AND owner_id = :owner"),
                        {"projeto": nome_do_projeto, "owner": owner_id},
                    )

            if 'projetos' in tabelas_existentes:
                conn.execute(
                    text("DELETE FROM projetos WHERE projeto = :projeto AND owner_id = :owner"),
                    {"projeto": nome_do_projeto, "owner": owner_id},
                )

            # ── 3. INSERIR PROJETO ────────────────────────────────
            df_projeto = selecionar_e_renomear(df_projeto, COLUNAS_PROJETOS)
            df_projeto = normalizar_dtypes(df_projeto)
            df_projeto['owner_id'] = owner_id

            df_projeto.to_sql(
                'projetos', conn,
                if_exists='append', index=False,
                dtype={'owner_id': UUID},
            )

            # ── 4. INSERIR ABAS FILHAS ────────────────────────────
            for sheet, (tabela, mapeamento) in ABAS_MAPEAMENTO.items():
                df_aba = ler_aba(file_buffer, sheet)
                if df_aba is None:
                    continue

                df_aba.columns = [limpar_nome_coluna(c) for c in df_aba.columns]
                df_aba = selecionar_e_renomear(df_aba, mapeamento)
                df_aba = normalizar_dtypes(df_aba)
                df_aba['projeto_vinculo'] = nome_do_projeto
                df_aba['owner_id'] = owner_id

                df_aba.to_sql(
                    tabela, conn,
                    if_exists='append', index=False,
                    dtype={'owner_id': UUID},
                )

        return {
            "status": "sucesso",
            "mensagem": f"Projeto '{nome_do_projeto}' importado com sucesso!",
            "projeto": nome_do_projeto,
        }

    except HTTPException:
        raise
    except Exception as e:
        print(f"❌ Erro interno: {e}")
        raise HTTPException(status_code=500, detail=f"Erro ao processar planilha: {e}")
