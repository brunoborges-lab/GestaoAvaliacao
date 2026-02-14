import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador UFCD - Lista Mestre", layout="wide")

st.title("📋 Consolidação por Lista de Importação")
st.markdown("Esta aplicação usa a coluna **'Nome'** do ficheiro de importação como referência principal.")

# --- BARRA LATERAL PARA UPLOADS ---
with st.sidebar:
    st.header("1. Ficheiro Mestre (Importação)")
    arquivo_importacao = st.file_uploader("Carregue a lista de formandos", type=["xlsx", "xls"], key="mestre")
    
    st.header("2. Ficheiros de Notas (Grelhas)")
    arquivos_grelha = st.file_uploader("Carregue as grelhas preenchidas", type=["xlsx", "xls"], accept_multiple_files=True, key="grelhas")

# --- FUNÇÕES DE PROCESSAMENTO ---

def obter_nomes_mestre(file):
    """Lê o ficheiro de importação e procura examente a coluna 'Nome'"""
    try:
        df = pd.read_excel(file)
        # Limpar espaços nos nomes das colunas (ex: "Nome " vira "Nome")
        df.columns = df.columns.str.strip()
        
        if "Nome" in df.columns:
            # Retorna um DataFrame apenas com a coluna Nome, removendo vazios
            return df[["Nome"]].dropna().drop_duplicates()
        else:
            st.error("❌ ERRO: Não encontrei uma coluna chamada 'Nome' no ficheiro de importação.")
            return None
    except Exception as e:
        st.error(f"Erro ao ler ficheiro de importação: {e}")
        return None

def processar_grelha_notas(file):
    """Extrai notas da grelha de avaliação"""
    try:
        # Pula o cabeçalho decorativo (ajuste o skiprows se necessário)
        df = pd.read_excel(file, skiprows=12)
        
        # Mapeamento das colunas da Grelha UFCD 9889
        # Coluna C (índice 2) costuma ser o Nome
        # Coluna BG (índice 58) costuma ser a Média Final
        # Coluna BP (índice 67) costuma ser a Situação
        
        colunas_map = {
            df.columns[2]: "Nome",  # Renomeamos para "Nome" para bater certo com o Mestre
            df.columns[58]: "Média Final",
            df.columns[67]: "Situação"
        }
        
        df = df.rename(columns=colunas_map)
        
        # Filtra apenas o que interessa e remove linhas sem nome
        df_limpo = df[["Nome", "Média Final", "Situação"]].dropna(subset=["Nome"])
        return df_limpo
        
    except Exception as e:
        st.warning(f"Não foi possível processar o ficheiro {file.name}. Verifique o formato.")
        return pd.DataFrame()

# --- LÓGICA PRINCIPAL ---

if arquivo_importacao:
    # 1. Carregar a Lista Mestre
    df_mestre = obter_nomes_mestre(arquivo_importacao)
    
    if df_mestre is not None:
        st.info(f"✅ Lista Mestre carregada com {len(df_mestre)} formandos.")
        
        df_final = df_mestre.copy()

        # 2. Se houver grelhas, processar e juntar
        if arquivos_grelha:
            lista_notas = []
            for arquivo in arquivos_grelha:
                notas = processar_grelha_notas(arquivo)
                lista_notas.append(notas)
            
            if lista_notas:
                df_todas_notas = pd.concat(lista_notas, ignore_index=True)
                
                # --- O CRUZAMENTO (VLOOKUP AUTOMÁTICO) ---
                # "Left Join": Mantém todos os nomes do Mestre e tenta encontrar a nota correspondente
                df_final = pd.merge(df_mestre, df_todas_notas, on="Nome", how="left")
                
        else:
            st.warning("A aguardar ficheiros de notas... (Mostrando apenas a lista de nomes)")

        # 3. Tabela Editável
        st.write("### 📝 Verificar e Editar Dados")
        st.write("Se algum nome não tiver nota, a célula aparecerá vazia. Pode preencher manualmente.")
        
        df_editado = st.data_editor(
            df_final,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True
        )

        # 4. Botão de Download
        st.divider()
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_editado.to_excel(writer, index=False, sheet_name='Pauta_Final')
            
        st.download_button(
            label="💾 Descarregar Ficheiro Final",
            data=buffer.getvalue(),
            file_name="Pauta_Consolidada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

else:
    st.info("👈 Por favor, carregue primeiro o Ficheiro de Importação (com a coluna 'Nome') na barra lateral.")
