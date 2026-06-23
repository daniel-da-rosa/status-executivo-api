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
    re_limpo = re.sub(r'[\s\.\-\/]+', '_', nome)
    re_limpo = re.sub(r'[^\w]', '', re_limpo)
    re_limpo = re.sub(r'_+', '_', re_limpo)
    re_limpo = re.sub(r'_+', '_', re_limpo)
    re_limpo = re_limpo.strip('_')
    if re_limpo and re_limpo[0].isdigit():
        re_limpo = f'col_{re_limpo}'
    return re_limpo


def importar_dados():
    arquivo = 'escopo.xlsx'

    if not os.path.exists(arquivo):
        print(f"❌ Arquivo '{arquivo}' não encontrado.")
        return

    try:
        # ── 1. DESCOBRIR NOME DO PROJETO (Apenas linhas 1-8 da aba PROJETO) ──
        # nrows=8 garante que não vamos puxar as linhas do cronograma para a tabela de projetos
        df_projeto_raw = pd.read_excel(
            arquivo, sheet_name='PROJETO', header=None, usecols="A:B", nrows=8
        )
        df_projeto = df_projeto_raw.set_index(0).T
        df_projeto.columns = [limpar_nome_coluna(c) for c in df_projeto.columns]
        df_projeto = df_projeto.reset_index(drop=True)

        if 'projeto' not in df_projeto.columns:
            print("❌ Coluna 'projeto' não encontrada na aba PROJETO.")
            return

        nome_do_projeto = str(df_projeto['projeto'].iloc[0]).strip()
        print(f"\n📁 Projeto identificado: '{nome_do_projeto}'")

        # ── 2. LIMPEZA FOCADA (Incluindo a nova tabela cronograma) ──────────
        print("\n🧹 Limpando dados antigos deste projeto...")
        tabelas_filhas = ['cronograma', 'fases', 'areas', 'pontos_atencao', 'riscos']
        
        for tabela in tabelas_filhas:
            try:
                with engine.begin() as conn_limpeza:
                    query = text(f"DELETE FROM {tabela} WHERE projeto_vinculo = :projeto")
                    conn_limpeza.execute(query, {"projeto": nome_do_projeto})
                print(f"   🗑️  Registros limpos na tabela '{tabela}'.")
            except Exception:
                pass

        try:
            with engine.begin() as conn_limpeza:
                query = text("DELETE FROM projetos WHERE projeto = :projeto")
                conn_limpeza.execute(query, {"projeto": nome_do_projeto})
            print("   🗑️  Registro limpo na tabela 'projetos'.\n")
        except Exception:
            pass


        # ── 3. INSERIR DADOS DA ABA PROJETO (METADADOS) ───────────────────
        df_projeto.to_sql('projetos', engine, if_exists='append', index=False, method='multi')
        print(f"✅ Dados mestres do PROJETO salvos.")


        # ── 3.5 EXTRAIR E INSERIR NOVA TABELA CRONOGRAMA (Linhas 9 a 15) ────
        print(f"⏳ Extraindo cronograma (linhas 9-15)...")
        # skiprows=8 pula as primeiras 8 linhas de metadados, nrows=7 lê as 7 etapas (A, B, C)
        df_crono = pd.read_excel(
            arquivo, sheet_name='PROJETO', skiprows=8, nrows=7, header=None, usecols="A:C"
        )
        
        # Nomeia e limpa as colunas do cronograma
        df_crono.columns = ['etapa', 'data_inicio', 'data_fim']
        df_crono['projeto_vinculo'] = nome_do_projeto
        
        # Garante a formatação correta de datas para o banco
        df_crono['data_inicio'] = pd.to_datetime(df_crono['data_inicio'], errors='coerce')
        df_crono['data_fim'] = pd.to_datetime(df_crono['data_fim'], errors='coerce')

        df_crono.to_sql('cronograma', engine, if_exists='append', index=False, method='multi')
        print(f"✅ Tabela 'cronograma' salva com sucesso ({len(df_crono)} linhas).")


        # ── 4. INSERIR ABAS FILHAS DEMAIS ────────────────────────────────────
        abas = {
            'ATIVIDADE':      'fases',
            'AREAS':          'areas',
            'PONTOS_ATENCAO': 'pontos_atencao',
            'RISCOS':         'riscos',
        }

        for aba, tabela in abas.items():
            try:
                df = pd.read_excel(arquivo, sheet_name=aba)
                df = df.dropna(how='all')
            except Exception:
                print(f"⚠️  Aba '{aba}' não encontrada — pulando.")
                continue

            if df.empty:
                continue

            df.columns = [limpar_nome_coluna(c) for c in df.columns]
            df['projeto_vinculo'] = nome_do_projeto

            df.to_sql(tabela, engine, if_exists='append', index=False, method='multi')
            print(f"✅ Aba {aba} → tabela '{tabela}' ({len(df)} linhas)")

        print(f"\n🚀 Importação de '{nome_do_projeto}' finalizada com sucesso!\n")

    except Exception as e:
        print(f"\n❌ Erro durante a importação: {e}")
        raise

if __name__ == "__main__":
    importar_dados()