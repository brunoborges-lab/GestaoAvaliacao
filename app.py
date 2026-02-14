import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Gestor UFCD - Critérios Práticos", layout="wide")

st.title("🛠️ Sistema de Avaliação com Critérios Técnicos")

# --- SIDEBAR ---
with st.sidebar:
    st.header("📂 Upload de Configuração")
    arquivo_importacao = st.file_uploader("1. Importação (Nomes em K13)", type=["xlsx", "xls"])
    arquivo_criterios = st.file_uploader("2. Ficheiro de Avaliação Prática (Critérios)", type=["xlsx", "xls"])
    arquivos_grelha = st.file_uploader("3. Grelhas de Notas (Formadores)", type=["xlsx", "xls"], accept_multiple_files=True)

# --- FUNÇÕES ---

def obter_nomes(file):
    df = pd.read_excel(file, skiprows=12, usecols="K")
    df.columns = ["Nome"]
    return df.dropna(subset=["Nome"]).drop_duplicates()

def extrair_lista_criterios(file):
    """Lê o ficheiro de avaliação para extrair as frases dos critérios"""
    try:
        # Lendo a folha específica de observação
        df = pd.read_excel(file, sheet_name=None)
        # Procuramos a folha que contém "Grelha" ou "Observação"
        nome_folha = [s for s in df.keys() if 'Grelha' in s or 'Observação' in s][0]
        df_criterios = df[nome_folha]
        
        # Extrair textos da coluna onde estão as descrições (ajustado para a coluna G/H)
        # Aqui fazemos uma busca por palavras-chave para identificar as linhas certas
        criterios = df_criterios.iloc[:, 6].dropna().tolist() # Coluna 6 costuma ter as descrições
        return [c for c in criterios if len(str(c)) > 10] # Filtra apenas frases longas
    except:
        return ["Transporta as ferramentas...", "Opera perpendicular ao objetivo...", "Estabiliza o veículo..."]

def processar_notas_detalhadas(file):
    """Extrai as notas parciais das grelhas dos formadores"""
    df = pd.read_excel(file, skiprows=12)
    # Índices baseados na estrutura padrão da UFCD 9889
    return {
        "Nome": df.iloc[:, 2], 
        "Ferramentas": df.iloc[:, 28],
        "Equipamentos": df.iloc[:, 38],
        "Estabilização": df.iloc[:, 48]
    }

# --- INTERFACE ---

if arquivo_importacao:
    df_mestre = obter_nomes(arquivo_importacao)
    df_mestre["Nome"] = df_mestre["Nome"].astype(str).str.strip()

    # Se carregou o ficheiro de critérios, mostra-os como ajuda visual
    if arquivo_criterios:
        with st.expander("🔍 Ver Critérios de Preenchimento Extraídos"):
            crit_list = extrair_lista_criterios(arquivo_criterios)
            for c in crit_list:
                st.write(f"- {c}")

    tab1, tab2 = st.tabs(["📋 Pauta Final", "📝 Avaliação Prática Detalhada"])

    with tab2:
        st.subheader("Preenchimento de Critérios")
        
        # Criar colunas base para edição
        if arquivos_grelha:
            # Lógica para consolidar notas se já existirem ficheiros
            lista_aux = []
            for f in arquivos_grelha:
                data = processar_notas_detalhadas(f)
                lista_aux.append(pd.DataFrame(data))
            df_notas = pd.concat(lista_aux).drop_duplicates(subset=["Nome"])
            df_pratica = pd.merge(df_mestre, df_notas, on="Nome", how="left")
        else:
            # Criar colunas vazias para preenchimento manual
            df_pratica = df_mestre.copy()
            for col in ["Ferramentas", "Equipamentos", "Estabilização"]:
                df_pratica[col] = 0.0

        # EDITOR INTERATIVO
        df_editado = st.data_editor(
            df_pratica,
            use_container
