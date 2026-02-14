import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Gestor de Avaliação UFCD", layout="wide")

st.title("📊 Sistema de Gestão de Notas - UFCD 9889")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("📂 Upload de Ficheiros")
    arquivo_importacao = st.file_uploader("1. Importação (Nomes em K13)", type=["xlsx", "xls"])
    arquivos_grelha = st.file_uploader("2. Grelhas de Avaliação", type=["xlsx", "xls"], accept_multiple_files=True)

# --- FUNÇÕES DE PROCESSAMENTO ---

def extrair_nomes_mestre(file):
    try:
        # Lê a coluna K a partir da linha 13
        df = pd.read_excel(file, skiprows=12, usecols="K")
        df.columns = ["Nome"]
        return df.dropna(subset=["Nome"]).drop_duplicates()
    except:
        return pd.DataFrame(columns=["Nome"])

def extrair_detalhes_pratica(file):
    try:
        # Nas suas grelhas, os dados costumam estar nestas posições:
        # Nome: Coluna C (Index 2)
        # Ferramentas: Coluna AC (Index 28)
        # Equipamentos: Coluna AM (Index 38)
        # Estabilização: Coluna AW (Index 48)
        df = pd.read_excel(file, skiprows=12)
        
        colunas_pratica = {
            df.columns[2]: "Nome",
            df.columns[28]: "Ferramentas (60%)",
            df.columns[38]: "Equipamentos (20%)",
            df.columns[48]: "Estabilização (20%)",
            df.columns[58]: "Média Prática"
        }
        df_resumo = df.rename(columns=colunas_pratica)
        return df_resumo[["Nome", "Ferramentas (60%)", "Equipamentos (20%)", "Estabilização (20%)", "Média Prática"]].dropna(subset=["Nome"])
    except:
        return pd.DataFrame()

# --- LÓGICA DE INTERFACE ---

if arquivo_importacao:
    df_mestre = extrair_nomes_mestre(arquivo_importacao)
    df_mestre["Nome"] = df_mestre["Nome"].astype(str).str.strip()
    
    # Processar grelhas se existirem
    df_pratica_total = pd.DataFrame()
    if arquivos_grelha:
        lista_pratica = [extrair_detalhes_pratica(f) for f in arquivos_grelha]
        df_pratica_total = pd.concat(lista_pratica, ignore_index=True).drop_duplicates(subset=["Nome"])

    # Criar os Separadores (Tabs)
    tab1, tab2 = st.tabs(["📋 Lista Consolidada", "🛠️ Detalhe Avaliação Prática"])

    with tab1:
        st.subheader("Pauta Geral de Avaliação")
        # União simples para a pauta geral
        df_geral = pd.merge(df_mestre, df_pratica_total[["Nome", "Média Prática"]], on="Nome", how="left")
        st.data_editor
