import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io
import zipfile

st.title("Gerador de Avaliações Práticas")

# 1. Upload dos ficheiros necessários
file_importacao = st.file_uploader("1. Carregar ficheiro 'Importação' (Nomes)", type=["xlsx"])
file_modelo = st.file_uploader("2. Carregar Ficheiro Modelo (.xlsx)", type=["xlsx"])

if file_importacao and file_modelo:
    # Ler os nomes do ficheiro de importação
    df_nomes = pd.read_excel(file_importacao)
    
    # Tentar encontrar a coluna de nomes (ajuste conforme o seu Excel)
    coluna_nome = st.selectbox("Selecione a coluna que contém os nomes:", df_nomes.columns)
    nomes = df_nomes[coluna_nome].dropna().tolist()

    # Campo para nota ou avaliação (opcional, se quiser aplicar a mesma a todos)
    avaliacao_padrao = st.text_input("Avaliação Prática (texto ou nota para preencher):", "Apto")

    if st.button(f"Gerar {len(nomes)} ficheiros"):
        # Criar um buffer na memória para guardar o ficheiro ZIP
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for nome in nomes:
                # Carregar o modelo original para cada formando
                # Reiniciar o ponteiro do modelo para cada iteração
                file_modelo.seek(0)
                wb = load_workbook(file_modelo)
                ws = wb.active # Ou use wb["Nome_da_Folha"]

                # --- LÓGICA DE PREENCHIMENTO ---
                # Aqui você deve indicar a célula exata (ex: 'B5') 
                # onde o nome deve ser inserido. 
                # Se o campo se chama 'Nome_Formando', vamos supor que é a célula B2.
                ws['B2'] = nome 
                # Exemplo: preencher a avaliação na célula C10
                ws['C10'] = avaliacao_padrao

                # Guardar o ficheiro modificado num buffer temporário
                temp_file_buffer = io.BytesIO()
                wb.save(temp_file_buffer)
                
                # Adicionar ao ficheiro ZIP
                file_name = f"Avaliacao_{nome.replace(' ', '_')}.xlsx"
                zip_file.writestr(file_name, temp_file_buffer.getvalue())

        st.success("Todos os ficheiros foram gerados!")

        # Botão para descarregar todos os ficheiros de uma vez num ZIP
        st.download_button(
            label="Descarregar ficheiros (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="avaliacoes_formandos.zip",
            mime="application/zip"
        )
