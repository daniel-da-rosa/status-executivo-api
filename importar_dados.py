import os
import re
import unicodedata
import pandas as pd
from sqlalchemy import create_engine, text
from dotenv import load_dotenv

# ── Carrega .env ──────────────────────────────────────────────────
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
DATABASE_URL = os.getenv("DATABASE_URL")

if not DATABASE_URL:
    print("❌ ERRO: DATABASE_URL não encontrada no .env")
    exit()

engine = create_engine(DATABASE_URL)

# ── Helpers ───────────────────────────────────────────────────────
def limpar_nome_coluna(nome):
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


COLUNAS_EXCLUIR_PROJETOS = {
    'col_1_levantamento', 'col_2_cadastros',
    'col_3_etapa_i',      'col_4_etapa_ii',
    'col_5_etapa_iii',    'col_6_etapa_iv',
    'encerramento',       'nao_planejado',
}

def importar_dados():
    arquivo = 'escopo.xlsx'

    if not os.path.exists(arquivo):
        print(f"❌ Arquivo '{arquivo}' não encontrado.")
        return

    try:
        with engine.begin() as conn:

            # ── 1. DESCOBRIR NOME DO PROJETO (ABA PROJETO) ────────
            df_projeto_raw = pd.read_excel(
                arquivo, sheet_name='PROJETO', header=None, usecols="A:B"
            )
            df_projeto = df_projeto_raw.set_index(0).T
            df_projeto.columns = [limpar_nome_coluna(c) for c in df_projeto.columns]
            df_projeto = df_projeto.reset_index(drop=True)

            if 'projeto' not in df_projeto.columns:
                print("❌ Coluna 'projeto' não encontrada na aba PROJETO.")
                return

            nome_do_projeto = str(df_projeto['projeto'].iloc[0]).strip()
            print(f"\n📁 Projeto identificado: '{nome_do_projeto}'")

            # ── 2. LIMPEZA FOCADA (Apenas dados deste projeto) ────
            print("\n🧹 Limpando dados antigos deste projeto...")
            tabelas_filhas = ['fases', 'areas', 'pontos_atencao', 'riscos']
            for tabela in tabelas_filhas:
                # Tenta deletar (ignora erro caso a tabela ainda não exista no banco)
                try:
                    query = text(f"DELETE FROM {tabela} WHERE projeto_vinculo = :projeto")
                    conn.execute(query, {"projeto": nome_do_projeto})
                    print(f"   🗑️  Registros limpos na tabela '{tabela}'.")
                except Exception:
                    pass

            # Limpa o projeto na tabela principal
            try:
                query = text("DELETE FROM projetos WHERE projeto = :projeto")
                conn.execute(query, {"projeto": nome_do_projeto})
                print("   🗑️  Registro limpo na tabela 'projetos'.\n")
            except Exception:
                pass


            # ── 3. INSERIR DADOS DA ABA PROJETO ───────────────────
            colunas_remover = [c for c in df_projeto.columns if c in COLUNAS_EXCLUIR_PROJETOS]
            if colunas_remover:
                df_projeto = df_projeto.drop(columns=colunas_remover)

            df_projeto.to_sql('projetos', conn, if_exists='append', index=False)
            print(f"✅ Aba PROJETO salva.")

            # ── 4. INSERIR ABAS FILHAS ────────────────────────────
            abas = {
                'ATIVIDADE':      'fases',
                'AREAS':          'areas',
                'PONTOS_ATENCAO': 'pontos_atencao',
                'RISCOS':         'riscos',
            }

            for aba, tabela in abas.items():
                try:
                    df = pd.read_excel(arquivo, sheet_name=aba)
                except Exception:
                    print(f"⚠️  Aba '{aba}' não encontrada — pulando.")
                    continue

                if df.empty:
                    continue

                df.columns = [limpar_nome_coluna(c) for c in df.columns]
                df['projeto_vinculo'] = nome_do_projeto

                df.to_sql(tabela, conn, if_exists='append', index=False)
                print(f"✅ Aba {aba} → tabela '{tabela}' ({len(df)} linhas)")

        print(f"\n🚀 Importação de '{nome_do_projeto}' finalizada com sucesso!\n")

    except Exception as e:
        print(f"\n❌ Erro durante a importação: {e}")
        raise

if __name__ == "__main__":
    importar_dados()