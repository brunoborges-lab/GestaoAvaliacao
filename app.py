import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador UFCD - K13", layout="wide")

st.title("📋 Consolidação (Importação K13)")
st.info("Configurado para ler nomes a partir da célula K13 do ficheiro de importação.")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("1. Ficheiro Mestre (Importação)")
    arquivo_importacao = st.file_uploader("Carregue a lista (Nomes em K13)", type=["xlsx", "xls"], key="mestre")
    
    st.header("2. Ficheiros de Notas (Grelhas)")
    arquivos_grelha = st.file_uploader("Carregue as grelhas de avaliação", type=["xlsx", "xls"], accept_multiple_files=True, key="grelhas")

# --- FUNÇÕES ---

def obter_nomes_k13(file):
    """
    Lê o ficheiro focando apenas na coluna K e assumindo cabeçalho na linha 13.
    """
    try:
        # skiprows=12 -> A linha 13 torna-se o cabeçalho
        # usecols="K" -> Carrega apenas a coluna K
        df = pd.read_excel(file, skiprows=12, usecols="K")
        
        # Independentemente do nome que estiver na célula K13, vamos chamar-lhe "Nome"
        # para o código funcionar com o resto da lógica.
        df.columns = ["Nome"]
        
        # Limpar linhas vazias
        return df.dropna().drop_duplicates()
    except Exception as e:
        st.error(f"Erro ao ler a coluna K13: {e}")
        return None

def processar_grelha_notas(file):
    """Extrai notas da grelha de avaliação (Lógica da UFCD 9889)"""
    try:
        # NOTA: Mantive a lógica original para as grelhas.
        # Se as grelhas também tiverem mudado de sítio, avise-me!
        df = pd.read_excel(file, skiprows=12)
        
        # Mapeamento baseado nos ficheiros anteriores:
        # Coluna C (Index 2) -> Nome na Grelha
        # Coluna BG (Index 58) -> Média
        # Coluna BP (Index 67) -> Situação
        colunas_map = {
            df.columns[2]: "Nome",
            df.columns[58]: "Média Final",
            df.columns[67]: "Situação"
        }
        
        df = df.rename(columns=colunas_map)
        
        # Normalização: Retirar espaços extra dos nomes para bater certo com a lista mestre
        df["Nome"] = df["Nome"].astype(str).str.strip()
        
        return df[["Nome", "Média Final", "Situação"]].dropna(subset=["Nome"])
        
    except Exception as e:
        st.warning(f"Não foi possível ler as notas de {file.name}. Verifique se é uma grelha válida.")
        return pd.DataFrame()

# --- LÓGICA PRINCIPAL ---

if arquivo_importacao:
    # 1. Carregar Nomes da K13
    df_mestre = obter_nomes_k13(arquivo_importacao)
    
    if df_mestre is not None and not df_mestre.empty:
        st.success(f"✅ Lista carregada: {len(df_mestre)} nomes encontrados (Coluna K).")
        
        # Garantir que os nomes do mestre não têm espaços "invisíveis"
        df_mestre["Nome"] = df_mestre["Nome"].astype(str).str.strip()
        
        df_final = df_mestre.copy()

        # 2. Processar Grelhas (se existirem)
        if arquivos_grelha:
            lista_notas = []
            for arquivo in arquivos_grelha:
                notas = processar_grelha_notas(arquivo)
                lista_notas.append(notas)
            
            if lista_notas:
                df_todas_notas = pd.concat(lista_notas, ignore_index=True)
                
                # Remover duplicados nas notas (caso tenha carregado o mesmo ficheiro 2x)
                df_todas_notas = df_todas_notas.drop_duplicates(subset=["Nome"])

                # CRUZAMENTO DE DADOS (VLOOKUP)
                df_final = pd.merge(df_mestre, df_todas_notas, on="Nome", how="left")
        
        # 3. Mostrar Editor
        st.write("### 📝 Pauta Final")
        
        df_editado = st.data_editor(
            df_final,
            use_container_width=True,
            num_rows="dynamic",
            height=600
        )

        # 4. Download
        st.divider()
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_editado.to_excel(writer, index=False, sheet_name='Pauta_K13')
            
        st.download_button(
            label="💾 Descarregar Ficheiro Consolidado",
            data=buffer.getvalue(),
            file_name="Pauta_Consolidada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.error("⚠️ A coluna K parece estar vazia a partir da linha 13.")

else:
    st.info("👈 Carregue o ficheiro de Importação (onde os nomes estão na K13).")
