import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador Automático", layout="wide")

st.title("📊 Consolidador de Dados Inteligente")
st.markdown("Misture o **documento(7).xlsx** com qualquer outro ficheiro de forma automática.")

# 1. Upload dos ficheiros
col1, col2 = st.columns(2)

with col1:
    main_file = st.file_uploader("📂 Ficheiro Principal (Base)", type=["xlsx"])
with col2:
    extra_files = st.file_uploader("📂 Ficheiros de Complemento (um ou mais)", type=["xlsx"], accept_multiple_files=True)

if main_file and extra_files:
    df_main = pd.read_excel(main_file)
    
    st.subheader("Configuração da Junção")
    # Deixar o utilizador escolher a coluna comum (ID, Código, Email, etc.)
    common_col = st.selectbox("Selecione a coluna que existe em TODOS os ficheiros para servir de ligação:", df_main.columns)

    if st.button("Executar Fusão de Dados"):
        try:
            df_final = df_main.copy()
            
            for extra in extra_files:
                df_temp = pd.read_excel(extra)
                
                # Remover colunas duplicadas que não sejam a chave para evitar sufixos _x, _y
                cols_to_use = [col for col in df_temp.columns if col not in df_final.columns or col == common_col]
                
                # Realizar a junção
                df_final = pd.merge(df_final, df_temp[cols_to_use], on=common_col, how="left")
            
            st.success(f"✅ Sucesso! O ficheiro final tem agora {len(df_final.columns)} colunas.")
            
            # Preview do resultado
            st.dataframe(df_final.head(10))

            # 2. Preparar download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Consolidado')
            
            st.download_button(
                label="📥 Descarregar Ficheiro Preenchido",
                data=output.getvalue(),
                file_name="resultado_consolidado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Erro ao processar: {e}. Verifique se a coluna '{common_col}' existe em todos os ficheiros.")
