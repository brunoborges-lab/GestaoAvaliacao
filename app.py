import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Editor de Grelhas UFCD", layout="wide")

st.title("📝 Editor e Consolidador de Avaliações")
st.info("Pode editar as notas diretamente na tabela abaixo antes de exportar o ficheiro final.")

# 1. Upload
uploaded_files = st.file_uploader("Carregue os ficheiros Excel", type=["xlsx", "xls"], accept_multiple_files=True)

def processar_grelha(file):
    # Lógica de extração (ajustada aos seus ficheiros)
    df = pd.read_excel(file, skiprows=12)
    df = df.dropna(subset=[df.columns[2]])
    
    colunas_uteis = {
        df.columns[0]: "Nº",
        df.columns[2]: "Nome do Formando",
        df.columns[18]: "Nota Teórica",
        df.columns[58]: "Média Final",
        df.columns[67]: "Situação"
    }
    df = df.rename(columns=colunas_uteis)
    return df[["Nº", "Nome do Formando", "Nota Teórica", "Média Final", "Situação"]]

if uploaded_files:
    lista_dfs = []
    for file in uploaded_files:
        try:
            dados = processar_grelha(file)
            dados['Origem'] = file.name
            lista_dfs.append(dados)
        except:
            st.error(f"Erro no ficheiro {file.name}")

    if lista_dfs:
        df_base = pd.concat(lista_dfs, ignore_index=True)

        # --- A MÁGICA ACONTECE AQUI ---
        st.subheader("Tabela Interativa (Clique numa célula para editar)")
        
        # O data_editor permite alterar valores, adicionar ou remover linhas
        df_editado = st.data_editor(
            df_base, 
            num_rows="dynamic", # Permite adicionar/remover linhas se quiser
            use_container_width=True,
            key="editor_avaliacoes"
        )

        # 2. Download do que foi editado
        st.divider()
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Salvamos o df_editado e não o original!
            df_editado.to_excel(writer, index=False, sheet_name='Notas_Editadas')
        
        st.download_button(
            label="💾 Guardar Alterações e Descarregar Excel",
            data=output.getvalue(),
            file_name="Avaliacao_Final_Corrigida.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
