import streamlit as st
import pandas as pd

st.title("🎓 Sistema de Gestão de Avaliações")

# 1. Upload dos Ficheiros
file_alunos = st.file_uploader("Carregar Ficheiro de Alunos (Excel)", type=["xlsx"])
file_criterios = st.file_uploader("Carregar Ficheiro de Critérios (Excel)", type=["xlsx"])

if file_alunos and file_criterios:
    df_alunos = pd.read_excel(file_alunos)
    df_crit = pd.read_excel(file_criterios)
    
    st.subheader("Lançamento de Notas")
    
    notas_teoricas = []
    # Criar um formulário para entrada manual
    with st.form("form_notas"):
        for index, row in df_alunos.iterrows():
            col1, col2 = st.columns([3, 1])
            col1.write(f"**Aluno:** {row['Nome']}")
            nota = col2.number_input(f"Nota Teórica", key=f"nota_{index}", min_value=0.0, max_value=20.0)
            notas_teoricas.append(nota)
        
        submetido = st.form_submit_button("Processar Avaliações")

    if submetido:
        # 2. Cálculos e Preenchimento
        df_alunos['Nota Teórica'] = notas_teoricas
        
        # Exemplo de cálculo de Nota Final (Média Simples)
        # Assumindo que a nota prática vem do outro ficheiro
        nota_pratica_media = df_crit.iloc[:, 1:].mean(axis=1) # Média das colunas de critérios
        df_alunos['Nota Final'] = (df_alunos['Nota Teórica'] + nota_pratica_media) / 2

        # 3. Disponibilizar para Download
        st.success("Avaliações processadas com sucesso!")
        
        # Botão para baixar o Excel de Alunos Atualizado
        st.download_button(
            label="Baixar Pauta Final",
            data=df_alunos.to_csv(index=False).encode('utf-8'),
            file_name="pauta_final.csv",
            mime="text/csv"
        )
