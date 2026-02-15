import streamlit as st
import pandas as pd
import io

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Excel Online | Consolidador", layout="wide", page_icon="📊")

# --- ESTILO OFFICE 365 (CSS CUSTOMIZADO) ---
st.markdown("""
    <style>
    /* Fonte e Background Geral */
    @import url('https://fonts.cdnfonts.com/css/segoe-ui-4');
    html, body, [class*="css"] {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .main {
        background-color: #f3f2f1; /* Cinza claro do Office */
    }
    
    /* Header estilo Ribbon */
    .header-ribbon {
        background-color: #0078d4; /* Azul Microsoft */
        padding: 1rem;
        color: white;
        border-radius: 0px 0px 5px 5px;
        margin-bottom: 2rem;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    
    /* Estilo dos Cards de Upload */
    .stFileUploader {
        background-color: white;
        padding: 20px;
        border-radius: 2px;
        border-bottom: 3px solid #0078d4;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    /* Botões estilo Office */
    .stButton>button {
        background-color: #0078d4;
        color: white;
        border-radius: 2px;
        border: none;
        padding: 0.5rem 2rem;
        font-weight: 600;
    }
    .stButton>button:hover {
        background-color: #005a9e;
        color: white;
    }
    </style>
    
    <div class="header-ribbon">
        <h1 style='margin:0; font-size: 24px;'>📊 Microsoft 365 | Consolidador de Dados</h1>
        <p style='margin:0; font-size: 14px; opacity: 0.9;'>Junção inteligente de ficheiros Excel</p>
    </div>
""", unsafe_allow_html=True)

# --- LÓGICA DA APLICAÇÃO ---

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📁 Ficheiro Principal")
    main_file = st.file_uploader("Carregar documento(7).xlsx", type=["xlsx"], key="main")

with col2:
    st.subheader("📎 Ficheiros de Dados Extra")
    extra_files = st.file_uploader("Arraste os outros 2 ficheiros para aqui", type=["xlsx"], accept_multiple_files=True, key="extra")

if main_file and extra_files:
    # Lendo o ficheiro base
    df_main = pd.read_excel(main_file)
    
    st.divider()
    
    # Interface de configuração
    st.info("💡 **Configuração de Colunas**: Escolha a chave de ligação (ex: ID, Email ou Nome) que existe em todos os documentos.")
    common_col = st.selectbox("Coluna de Referência:", df_main.columns)

    if st.button("🔄 Processar e Unir Dados"):
        try:
            with st.spinner('A processar no motor Excel...'):
                df_final = df_main.copy()
                
                for extra in extra_files:
                    df_temp = pd.read_excel(extra)
                    
                    # Evitar duplicados de colunas, mantendo apenas a chave e as novas colunas
                    cols_to_use = [c for c in df_temp.columns if c not in df_final.columns or c == common_col]
                    df_final = pd.merge(df_final, df_temp[cols_to_use], on=common_col, how="left")
                
                st.success("Concluído! Os dados foram mesclados com sucesso.")
                
                # Visualização prévia em tabela estilo Excel
                st.write("### Pré-visualização do Resultado")
                st.dataframe(df_final.head(10), use_container_width=True)

                # Preparar ficheiro para download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Sheet1')
                
                st.download_button(
                    label="📥 Descarregar Ficheiro Final (XLSX)",
                    data=output.getvalue(),
                    file_name="documento_consolidado_365.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Erro na junção: Certifique-se de que a coluna '{common_col}' existe em todos os ficheiros.")

else:
    st.warning("Aguardando upload dos ficheiros para iniciar...")

# --- RODAPÉ ---
st.markdown("---")
st.caption("Ficheiros processados localmente na memória da aplicação. Nenhum dado é guardado permanentemente.")
                    
