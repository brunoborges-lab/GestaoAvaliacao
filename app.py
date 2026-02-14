import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Portal de Avaliação UFCD", layout="wide")

# Inicializar base de dados na memória para guardar o que for preenchido
if 'db_notas' not in st.session_state:
    st.session_state.db_notas = {}

st.title("📝 Formulário Individual de Avaliação")

# --- SIDEBAR: Carregamento de Estrutura ---
with st.sidebar:
    st.header("⚙️ Configuração")
    f_import = st.file_uploader("1. Ficheiro Importação (Nomes K13)", type=["xlsx", "xls"])
    f_criterios = st.file_uploader("2. Ficha de Avaliação Prática (Critérios)", type=["xlsx", "xls"])

# --- PROCESSAMENTO INICIAL ---
if f_import and f_criterios:
    # Obter Nomes
    df_nomes = pd.read_excel(f_import, skiprows=12, usecols="K").dropna()
    df_nomes.columns = ["Nome"]
    lista_formandos = df_nomes["Nome"].tolist()

    # Seleção do Formando
    formando_selecionado = st.selectbox("🎯 Selecione o Formando para avaliar:", lista_formandos)

    st.divider()

    # --- FORMULÁRIO DE AVALIAÇÃO ---
    with st.form("form_avaliacao"):
        st.subheader(f"Avaliação: {formando_selecionado}")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📘 Avaliação Teórica")
            nota_teorica = st.number_input("Nota do Teste (0-20)", min_value=0.0, max_value=20.0, step=0.1, key="teorica")

        with col2:
            st.markdown("### 🛠️ Avaliação Prática")
            st.caption("Ponderação: Ferramentas (60%), Equipamentos (20%), Estabilização (20%)")
            nota_ferr = st.slider("Operação com Ferramentas", 0, 20, 10)
            nota_equip = st.slider("Manuseamento de Equipamentos", 0, 20, 10)
            nota_estab = st.slider("Estabilização e Segurança", 0, 20, 10)

        # Cálculo da Média Prática e Final
        media_pratica = (nota_ferr * 0.6) + (nota_equip * 0.2) + (nota_estab * 0.2)
        nota_final = (nota_teorica * 0.5) + (media_pratica * 0.5)
        
        situacao = "APROVADO" if nota_final >= 9.5 else "NÃO APROVADO"

        st.info(f"**Resumo Atual:** Média Prática: {media_pratica:.2f} | **Nota Final: {nota_final:.2f}** ({situacao})")

        submetido = st.form_submit_button("✅ Guardar Avaliação")
        
        if submetido:
            # Guarda os dados no estado da sessão
            st.session_state.db_notas[formando_selecionado] = {
                "Nome": formando_selecionado,
                "Teórica": nota_teorica,
                "Prática_Ferramentas": nota_ferr,
                "Prática_Equipamentos": nota_equip,
                "Prática_Estabilização": nota_estab,
                "Média_Prática": media_pratica,
                "Nota_Final": nota_final,
                "Situação": situacao
            }
            st.success(f"Dados de {formando_selecionado} guardados com sucesso!")

    # --- TABELA DE RESUMO E EXPORTAÇÃO ---
    if st.session_state.db_notas:
        st.divider()
        st.subheader("📋 Registos Efetuados")
        df_final = pd.DataFrame.from_dict(st.session_state.db_notas, orient='index')
        st
