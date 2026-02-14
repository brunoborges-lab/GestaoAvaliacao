import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Avaliação Detalhada UFCD", layout="wide")

if 'registos' not in st.session_state:
    st.session_state.registos = {}

st.title("📑 Formulário de Avaliação por Subcategorias")

# --- DEFINIÇÃO DOS CRITÉRIOS (Extraídos do seu ficheiro) ---
CRITERIOS_DETALHADOS = {
    "Operação com Ferramentas (60%)": [
        "Transporta as ferramentas e procede a abertura e fecho em segurança",
        "Opera com a ferramenta perpendicular ao objetivo de trabalho",
        "Coloca-se do lado certo da ferramenta",
        "Efetua comunicação sobre abertura ou corte de estruturas",
        "Protege a(s) vítima(s) e o(s) socorrista(s) com proteção rígida"
    ],
    "Manuseamento de Equipamento (20%)": [
        "Escolhe equipamento adequado à função",
        "Transporta e opera os equipamentos em segurança",
        "Opera corretamente com o grupo energético",
        "Opera corretamente com equipamento de estabilização",
        "Opera corretamente equipamento pneumático"
    ],
    "Estabilização e Segurança (20%)": [
        "Sinaliza e delimita zonas de trabalho e zela pela segurança",
        "Estabiliza o(s) veículo(s) acidentado(s) de forma adequada",
        "Controla estabilização inicial e efetua estabilização progressiva",
        "Efetua limpeza da zona de trabalho",
        "Aplica as proteções nos pontos agressivos"
    ]
}

# --- SIDEBAR ---
with st.sidebar:
    f_import = st.file_uploader("Carregue Ficheiro Importação (K13)", type=["xlsx", "xls"])

if f_import:
    df_nomes = pd.read_excel(f_import, skiprows=12, usecols="K").dropna()
    df_nomes.columns = ["Nome"]
    formando = st.selectbox("Seleccione o Formando:", df_nomes["Nome"].unique())

    st.divider()

    with st.form("ficha_detalhada"):
        st.subheader(f"Avaliação de: {formando}")
        
        # --- AVALIAÇÃO TEÓRICA ---
        nota_teorica = st.number_input("Nota Avaliação Teórica (0-20)", 0.0, 20.0, 10.0)
        
        st.divider()
        st.markdown("### 🛠️ Avaliação Prática (Subcategorias)")
        
        notas_ferramentas = []
        notas_equipamento = []
        notas_estabilizacao = []

        # Criar a interface para cada subcategoria
        cols = st.columns(3)
        
        with cols[0]:
            st.info("Operação com Ferramentas")
            for item in CRITERIOS_DETALHADOS["Operação com Ferramentas (60%)"]:
                n = st.select_slider(f"{item}", options=[1, 3, 5], value=3, key=f"ferr_{item}")
                notas_ferramentas.append(n)
        
        with cols[1]:
            st.warning("Manuseamento de Equipamento")
            for item in CRITERIOS_DETALHADOS["Manuseamento de Equipamento (20%)"]:
                n = st.select_slider(f"{item}", options=[1, 3, 5], value=3, key=f"equip_{item}")
                notas_equipamento.append(n)
        
        with cols[2]:
            st.success("Estabilização e Segurança")
            for item in CRITERIOS_DETALHADOS["Estabilização e Segurança (20%)"]:
                n = st.select_slider(f"{item}", options=[1, 3, 5], value=3, key=f"estab_{item}")
                notas_estabilizacao.append(n)

        # CÁLCULOS (Convertendo a escala 1-5 para 0-20 se necessário, ou mantendo a média)
        # Média de cada bloco (escala 1 a 5) convertida para 0-20: (soma / (n*5)) * 20
        med_ferr = (sum(notas_ferramentas) / (len(notas_ferramentas) * 5)) * 20
        med_equip = (sum(notas_equipamento) / (len(notas_equipamento) * 5)) * 20
        med_estab = (sum(notas_estabilizacao) / (len(notas_estabilizacao) * 5)) * 20
        
        media_pratica = (med_ferr * 0.6) + (med_equip * 0.2) + (med_estab * 0.2)
        nota_final = (nota_teorica * 0.5) + (media_pratica * 0.5)

        btn_guardar = st.form_submit_button("💾 Guardar Avaliação Completa")

        if btn_guardar:
            st.session_state.registos[formando] = {
                "Nome": formando,
                "Teórica": nota_teorica,
                "Média Ferramentas": round(med_ferr, 2),
                "Média Equipamento": round(med_equip, 2),
                "Média Estabilização": round(med_estab, 2),
                "Média Prática": round(media_pratica, 2),
                "Nota Final": round(nota_final, 2),
                "Resultado": "APROVADO" if nota_final >= 9.5 else "NÃO APROVADO"
            }
            st.balloons()

    # --- TABELA DE RESULTADOS ---
    if st.session_state.registos:
        st.subheader("📋 Pauta Consolidada")
        df_resumo = pd.DataFrame.from_dict(st.session_state.registos, orient='index')
        st.dataframe(df_resumo, use_container_width=True)

        # Exportação para Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_resumo.to_excel(writer, index=False, sheet_name='Resultados_UFCD9889')
        
        st.download_button("📥 Descarregar Pauta Final", output.getvalue(), "Pauta_UFCD9889.xlsx")

else:
    st.info("Por favor, carregue o ficheiro de importação na barra lateral.")
