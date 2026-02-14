import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador de Grelhas UFCD", layout="wide")

st.title("📊 Recolha e União de Dados UFCD")
st.markdown("""
Esta aplicação extrai os dados dos formandos das grelhas de avaliação (UFCD 9889) 
e junta tudo num único ficheiro consolidado.
""")

# 1. Upload dos Ficheiros
uploaded_files = st.file_uploader("Selecione os ficheiros Excel (.xlsx ou .xls)", type=["xlsx", "xls"], accept_multiple_files=True)

def processar_grelha(file):
    # Ler o ficheiro ignorando as linhas de cabeçalho decorativas
    # Ajustamos para começar a ler onde os nomes dos formandos costumam estar
    df = pd.read_excel(file, skiprows=12) # Salta os logos e títulos
    
    # Limpeza básica: remover colunas totalmente vazias e linhas sem nome
    df = df.dropna(subset=[df.columns[2]]) # Assume que o nome está na 3ª coluna
    
    # Renomear colunas para algo legível (ajustado à sua estrutura)
    colunas_uteis = {
        df.columns[0]: "Nº",
        df.columns[2]: "Nome do Formando",
        df.columns[18]: "Nota Teórica",
        df.columns[28]: "Ferramentas (0.6)",
        df.columns[38]: "Equipamentos (0.2)",
        df.columns[48]: "Estabilização (0.2)",
        df.columns[58]: "Média Final",
        df.columns[67]: "Situação"
    }
    df = df.rename(columns=colunas_uteis)
    
    # Manter apenas as colunas que nos interessam
    return df[["Nº", "Nome do Formando", "Nota Teórica", "Ferramentas (0.6)", "Equipamentos (0.2)", "Estabilização (0.2)", "Média Final", "Situação"]]

if uploaded_files:
    lista_dfs = []
    
    for file in uploaded_files:
        try:
            dados = processar_grelha(file)
            dados['Origem'] = file.name # Para saber de que ficheiro veio a nota
            lista_dfs.append(dados)
        except Exception as e:
            st.error(f"Erro ao processar {file.name}: {e}")

    if lista_dfs:
        df_final = pd.concat(lista_dfs, ignore_index=True)
        
        st.subheader("Visualização dos Dados Consolidados")
        st.dataframe(df_final)

        # 2. Botão para Download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Consolidado')
        
        st.download_button(
            label="📥 Descarregar Excel Consolidado",
            data=output.getvalue(),
            file_name="Avaliacao_Total_UFCD.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
