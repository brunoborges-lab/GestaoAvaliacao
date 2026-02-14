import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Gestor UFCD 9889", layout="wide")

st.title("🚀 Consolidador Inteligente UFCD")

# --- SIDEBAR: Configurações de Importação ---
st.sidebar.header("1. Lista de Formandos")
import_file = st.sidebar.file_uploader("Ficheiro de Importação (Nomes)", type=["xlsx", "xls"])

st.sidebar.header("2. Ficheiros de Avaliação")
eval_files = st.sidebar.file_uploader("Grelhas de Avaliação (Notas)", type=["xlsx", "xls"], accept_multiple_files=True)

# Função para extrair nomes do ficheiro de Importação
def obter_lista_nomes(file):
    # Ajuste o 'skiprows' ou 'usecols' conforme a estrutura real do seu ficheiro de importação
    df_imp = pd.read_excel(file)
    # Procuramos uma coluna que contenha 'Nome'
    coluna_nome = [col for col in df_imp.columns if 'Nome' in str(col)][0]
    return df_imp[coluna_nome].dropna().unique().tolist()

# Função para processar as notas das grelhas
def processar_notas(file):
    df = pd.read_excel(file, skiprows=12)
    # Selecionamos colunas de interesse (ajustado à Grelha UFCD 9889)
    # Coluna 2 costuma ser o Nome, Coluna 58 a Média, Coluna 67 a Situação
    cols = {df.columns[2]: "Nome do Formando", df.columns[58]: "Média Final", df.columns[67]: "Situação"}
    df = df.rename(columns=cols)
    return df[["Nome do Formando", "Média Final", "Situação"]].dropna(subset=["Nome do Formando"])

# --- LÓGICA PRINCIPAL ---
nomes_mestre = []
if import_file:
    nomes_mestre = obter_lista_nomes(import_file)
    st.success(f"Foram encontrados {len(nomes_mestre)} formandos no ficheiro de importação.")

if eval_files:
    dfs_notas = []
    for f in eval_files:
        dfs_notas.append(processar_notas(f))
    
    df_consolidado = pd.concat(dfs_notas, ignore_index=True)

    # Se tivermos a lista de nomes, garantimos que todos aparecem (mesmo sem nota)
    if nomes_mestre:
        df_nomes = pd.DataFrame({"Nome do Formando": nomes_mestre})
        # Unimos a lista de nomes com as notas encontradas (Left Join)
        df_final = pd.merge(df_nomes, df_consolidado, on="Nome do Formando", how="left")
    else:
        df_final = df_consolidado

    st.subheader("Edição de Dados e Notas")
    # Ativação da edição
    df_editado = st.data_editor(df_final, use_container_width=True, num_rows="dynamic")

    # Botão de Exportação
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_editado.to_excel(writer, index=False, sheet_name='Pauta_Final')
    
    st.download_button(
        label="📥 Descarregar Pauta Consolidada",
        data=output.getvalue(),
        file_name="Pauta_UFCD_9889.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Aguardando o upload das grelhas de avaliação...")
