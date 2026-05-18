import os
import re
import unicodedata
import io
import gc
import pandas as pd
from sqlalchemy import create_engine, text, inspect
from sqlalchemy.dialects.postgresql import UUID
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from mangum import Mangum

# ── Conexão ───────────────────────────────────────────────────────
DATABASE_URL = os.getenv("DATABASE_URL")

if not DATABASE_URL:
    raise RuntimeError("❌ ERRO: DATABASE_URL não encontrada nas variáveis de ambiente.")

engine = create_engine(DATABASE_URL, pool_size=1, max_overflow=0)

# ── Setup da FastAPI ──────────────────────────────────────────────
app = FastAPI(title="API Importação Iniflex")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Mapeamentos ───────────────────────────────────────────────────
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
    'probabilidade':     'probabilidade',
    'impacto':           'impacto',
}

COLUNAS_RISCOS = {
    'indicado_por_area': 'indicado_por_area',
    'detalhes':          'detalhes',
    'fase':              'fase',
    'probabilidade':     'probabilidade',
    'impacto':           'impacto',
}
# ─────────────────────────────────────────────────────────────────
# ADICIONAR ao dicionário ABAS_MAPEAMENTO (junto com as outras abas)
# ─────────────────────────────────────────────────────────────────

COLUNAS_OBJETIVOS = {
    'objetivo':  'objetivo',
    'descricao': 'descricao',
    'status':    'status',
    'icone':     'icone',
}

ABAS_MAPEAMENTO = {
    'ATIVIDADE':      ('fases',           COLUNAS_FASES),
    'AREAS':          ('areas',           COLUNAS_AREAS),
    'PONTOS_ATENCAO': ('pontos_atencao',  COLUNAS_PONTOS_ATENCAO),
    'RISCOS':         ('riscos',          COLUNAS_RISCOS),
    'OBJETIVOS':      ('objetivos',       COLUNAS_OBJETIVOS),   
}

 ─────────────────────────────────────────────────────────────────

# ── Helpers ───────────────────────────────────────────────────────
def limpar_nome_coluna(nome) -> str | None:
    """
    Normaliza um nome de coluna para snake_case sem acentos.
    Retorna None se o valor for vazio ou nulo (coluna sem cabeçalho).
    """
    if nome is None:
        return None
    nome = str(nome).strip()
    if not nome:
        return None
    nome = unicodedata.normalize('NFD', nome)
    nome = ''.join(c for c in nome if unicodedata.category(c) != 'Mn')
    nome = nome.lower()
    nome = re.sub(r'[\s\.\-\/]+', '_', nome)
    nome = re.sub(r'[^\w]', '', nome)
    nome = re.sub(r'_+', '_', nome)
    nome = nome.strip('_')
    if nome and nome[0].isdigit():
        nome = f'col_{nome}'
    return nome or None


def normalizar_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        if isinstance(df[col].dtype, pd.StringDtype):
            df[col] = df[col].astype(object)
        elif pd.api.types.is_datetime64_any_dtype(df[col].dtype):
            df[col] = pd.to_datetime(df[col]).dt.tz_localize(None)
    return df


def selecionar_e_renomear(df: pd.DataFrame, mapeamento: dict, sheet: str = '') -> pd.DataFrame:
    """
    Seleciona e renomeia colunas conforme o mapeamento.
    Loga um aviso para colunas esperadas que não foram encontradas.
    """
    presentes = {k: v for k, v in mapeamento.items() if k in df.columns}
    faltando = [k for k in mapeamento if k not in df.columns]

    if faltando:
        print(f"⚠️  [{sheet}] Colunas do mapeamento não encontradas no Excel (serão ignoradas): {faltando}")

    return df[list(presentes.keys())].rename(columns=presentes)


def remover_colunas_sem_cabecalho(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove colunas cujo nome é None após a limpeza (cabeçalhos vazios/nulos no Excel).
    Evita que colunas extras vazias do Excel poluam o DataFrame silenciosamente.
    """
    colunas_validas = [c for c in df.columns if c is not None]
    removidas = len(df.columns) - len(colunas_validas)
    if removidas:
        print(f"⚠️  {removidas} coluna(s) sem cabeçalho removida(s).")
    return df[colunas_validas]


def ler_aba(file_buffer: io.BytesIO, sheet_name: str, **kwargs):
    try:
        file_buffer.seek(0)
        df = pd.read_excel(file_buffer, sheet_name=sheet_name, **kwargs)
        return df if not df.empty else None
    except Exception:
        return None


# ── Endpoint ──────────────────────────────────────────────────────
@app.post("/api/importar")
async def importar_dados(
    arquivo: UploadFile = File(...),
    owner_id: str = Form(...),
):
    if not arquivo.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="O arquivo deve ser Excel (.xlsx ou .xls)")

    contents = None
    file_buffer = None

    try:
        contents = await arquivo.read()
        file_buffer = io.BytesIO(contents)

        tabelas_existentes = set(inspect(engine).get_table_names())

        with engine.begin() as conn:

            # 1. NOME DO PROJETO
            df_projeto_raw = ler_aba(file_buffer, 'PROJETO', header=None, usecols="A:B")
            if df_projeto_raw is None:
                raise HTTPException(status_code=400, detail="Aba 'PROJETO' não encontrada ou vazia.")

            df_projeto = df_projeto_raw.set_index(0).T
            df_projeto.columns = [limpar_nome_coluna(c) for c in df_projeto.columns]
            df_projeto = remover_colunas_sem_cabecalho(df_projeto)
            df_projeto = df_projeto.reset_index(drop=True)

            if 'projeto' not in df_projeto.columns:
                raise HTTPException(status_code=400, detail="Coluna 'projeto' não encontrada na aba PROJETO.")

            nome_do_projeto = str(df_projeto['projeto'].iloc[0]).strip()

            # 2. LIMPEZA
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

            # 3. INSERIR PROJETO
            df_projeto = selecionar_e_renomear(df_projeto, COLUNAS_PROJETOS, sheet='PROJETO')
            df_projeto = normalizar_dtypes(df_projeto)
            df_projeto['owner_id'] = owner_id
            df_projeto.to_sql('projetos', conn, if_exists='append', index=False, dtype={'owner_id': UUID})

            # 4. INSERIR ABAS FILHAS
            for sheet, (tabela, mapeamento) in ABAS_MAPEAMENTO.items():
                df_aba = ler_aba(file_buffer, sheet)
                if df_aba is None:
                    print(f"⚠️  Aba '{sheet}' não encontrada ou vazia — ignorada.")
                    continue

                # Limpa nomes de colunas e remove as que não têm cabeçalho
                df_aba.columns = [limpar_nome_coluna(c) for c in df_aba.columns]
                df_aba = remover_colunas_sem_cabecalho(df_aba)

                df_aba = selecionar_e_renomear(df_aba, mapeamento, sheet=sheet)
                df_aba = normalizar_dtypes(df_aba)
                df_aba['projeto_vinculo'] = nome_do_projeto
                df_aba['owner_id'] = owner_id
                df_aba.to_sql(tabela, conn, if_exists='append', index=False, dtype={'owner_id': UUID})

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
    finally:
        if file_buffer:
            file_buffer.close()
        del contents, file_buffer
        gc.collect()


# ── Handler Vercel ────────────────────────────────────────────────
handler = Mangum(app)