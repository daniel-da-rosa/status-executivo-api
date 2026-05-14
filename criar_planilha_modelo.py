import pandas as pd

def criar_planilha_validacao():
    file_name = 'dashboard_data.xlsx'
    
    # Aba 1: RESUMO (Para o Gauge e Títulos)
    df_resumo = pd.DataFrame({
        'cliente': ['MB EMBALAGENS'],
        'projeto': ['IMPLANTAÇÃO INIFLEX'],
        'lider': ['DANIEL DA ROSA'],
        'porcentagem_total': [35],
        'horas_contrato': [1135],
        'horas_utilizadas': [68.41],
        'status_geral': ['NO PRAZO']
    })

    # Aba 2: TIMELINE (Para o componente de Cronograma)
    df_timeline = pd.DataFrame({
        'fase_nome': ['1-LEVANTAMENTO', '2-CADASTROS', '3-ETAPA I', '4-ETAPA II', '5-ETAPA III', '6-ETAPA IV', 'ENCERRAMENTO'],
        'data_inicio': ['2026-03-23', '2026-04-06', '2026-05-25', '2026-06-15', '2026-07-06', '2026-08-03', '2026-08-24'],
        'data_fina': ['2026-03-27', '2026-04-10', '2026-06-06', '2026-06-26', '2026-07-17', '2026-08-14', '2026-09-04'],
        'progresso': [1.0, 0.4, 0.0, 0.0, 0.0, 0.0, 0.0],
        'status': ['CONCLUÍDO', 'EM ANDAMENTO', 'PENDENTE', 'PENDENTE', 'PENDENTE', 'PENDENTE', 'PENDENTE']
    })

    # Aba 3: RISCOS (Para a Matriz 3x3)
    # prob/imp: 0=Baixo, 1=Médio, 2=Alto
    df_riscos = pd.DataFrame({
        'descricao_risco': ['Infraestrutura de Rede', 'Adesão da Equipe', 'Qualidade dos Dados'],
        'probabilidade': [0, 1, 2], 
        'impacto': [1, 2, 2],
        'cor': ['#f1c40f', '#e67e22', '#e74c3c'] 
    })

    # Aba 4: STATUS_AREAS (Para a tabela lateral)
    df_areas = pd.DataFrame({
        'area': ['Cadastros', 'Compras', 'Financeiro', 'Produção', 'Vendas', 'Qualidade'],
        'status': ['OK', 'OK', 'PENDENTE', 'ALERTA', 'OK', 'PENDENTE'],
        'prazo': ['OK', 'OK', 'ATRASO', 'OK', 'OK', 'OK'],
        'escopo': ['OK', 'OK', 'OK', 'AJUSTAR', 'OK', 'OK']
    })

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, sheet_name='RESUMO', index=False)
        df_timeline.to_excel(writer, sheet_name='TIMELINE', index=False)
        df_riscos.to_excel(writer, sheet_name='RISCOS', index=False)
        df_areas.to_excel(writer, sheet_name='STATUS_AREAS', index=False)

    print(f"Planilha '{file_name}' criada com sucesso para validação!")

if __name__ == "__main__":
    criar_planilha_validacao()
