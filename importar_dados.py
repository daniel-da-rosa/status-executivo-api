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
    re_limpo = re_limpo.strip('_')
    if re_limpo and re_limpo[0].isdigit():
        re_limpo = f'col_{re_limpo}'
    return re_limpo


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

        # ── 2. LIMPEZA FOCADA ─────────────────────────────────────────────
        print("\n🧹 Limpando dados antigos deste projeto...")
        tabelas_filhas = ['fases', 'areas', 'pontos_atencao', 'riscos']
        
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


        # ── 3. INSERIR DADOS DA ABA PROJETO ───────────────────
        colunas_remover = [c for c in df_projeto.columns if c in COLUNAS_EXCLUIR_PROJETOS]
        if colunas_remover:
            df_projeto = df_projeto.drop(columns=colunas_remover)

        df_projeto.to_sql('projetos', engine, if_exists='append', index=False, method='multi')
        print(f"✅ Aba PROJETO salva.")

        # ── 4. INSERIR ABAS FILHAS (COM RASTREAMENTO CORRIGIDO) ────────────────
        abas = {
            'ATIVIDADE':      'fases',
            'AREAS':          'areas',
            'PONTOS_ATENCAO': 'pontos_atencao',
            'RISCOS':         'riscos',
        }

        for aba, tabela in abas.items():
            try:
                # nrows=1000 força o Pandas a parar de ler o Excel se houverem linhas fantasmas infinitas
                print(f"⏳ [1/2] Lendo a aba '{aba}' do Excel...")
                df = pd.read_excel(arquivo, sheet_name=aba, nrows=1000)
                
                # Remove linhas completamente vazias que possam ter vindo do Excel
                df = df.dropna(how='all')
            except Exception as e:
                print(f"⚠️  Erro ao ler aba '{aba}': {e} — pulando.")
                continue

            if df.empty:
                print(f"ℹ️  Aba '{aba}' está vazia no arquivo.")
                continue

            df.columns = [limpar_nome_coluna(c) for c in df.columns]
            df['projeto_vinculo'] = nome_do_projeto

            print(f"🚀 [2/2] Gravando {len(df)} linhas na tabela '{tabela}' do Supabase...")
            df.to_sql(tabela, engine, if_exists='append', index=False, method='multi')
            print(f"✅ Aba {aba} → salva com sucesso!\n")

        print(f"\n🚀 Importação de '{nome_do_projeto}' finalizada com sucesso!\n")

    except Exception as e:
        print(f"\n❌ Erro durante a importação: {e}")
        raise

if __name__ == "__main__":
    importar_dados()