import streamlit as st
import pandas as pd
import io
from fpdf import FPDF

st.set_page_config(page_title="Gerador de Pautas PDF", layout="wide")

# Inicialização do estado
if 'base_dados' not in st.session_state:
    st.session_state.base_dados = {}

# --- CRITÉRIOS ---
CRITERIOS = {
    "Operação com Ferramentas (60%)": [
        "Transporta as ferramentas e procede a abertura e fecho em segurança",
        "Opera com a ferramenta perpendicular ao objetivo",
        "Coloca-se do lado certo da ferramenta",
        "Efetua comunicação sobre abertura/corte",
        "Protege a(s) vítima(s) e socorrista(s)"
    ],
    "Manuseamento Equipamento (20%)": [
        "Escolhe equipamento adequado à função",
        "Transporta e opera equipamentos em segurança",
        "Opera corretamente grupo energético",
        "Opera corretamente equip. estabilização",
        "Opera corretamente equip. pneumático"
    ],
    "Estabilização e Segurança (20%)": [
        "Sinaliza e delimita zonas de trabalho",
        "Estabiliza o(s) veículo(s) adequadamente",
        "Controla estabilização inicial e progressiva",
        "Efetua limpeza da zona de trabalho",
        "Aplica proteções nos pontos agressivos"
    ]
}

# --- FUNÇÃO PARA GERAR PDF ---
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'FICHA DE AVALIAÇÃO PRÁTICA - UFCD 9889', 0, 1, 'C')
        self.ln(5)

def gerar_pdf_final(dados_todos):
    pdf = PDF()
    for nome, dados in dados_todos.items():
        pdf.add_page()
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(0, 10, f"Formando: {nome}", 0, 1)
        pdf.set_font('Arial', '', 10)
        
        # Notas
        pdf.cell(0, 8, f"Nota Teórica: {dados['Teórica']}", 0, 1)
        pdf.cell(0, 8, f"Média Prática: {dados['Média Prática']}", 0, 1)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(0, 10, f"CLASSIFICAÇÃO FINAL: {dados['Nota Final']} - {dados['Situação']}", 0, 1)
        
        pdf.ln(5)
        pdf.set_font('Arial', 'I', 8)
        pdf.cell(0, 5, "-" * 100, 0, 1)
        
    return pdf.output(dest='S').encode('latin-1')

# --- INTERFACE ---
st.title("🎓 Gerador de Avaliações PDF (UFCD 9889)")

with st.sidebar:
    f_import = st.file_uploader("Ficheiro Importação (K13)", type=["xlsx", "xls"])

if f_import:
    df_nomes = pd.read_excel(f_import, skiprows=12, usecols="K").dropna()
    df_nomes.columns = ["Nome"]
    nomes = df_nomes["Nome"].unique()
    
    formando = st.selectbox("Escolha o formando para avaliar:", nomes)

    with st.form("avaliacao_pdf"):
        nota_t = st.number_input("Nota Teórica", 0.0, 20.0, 10.0)
        
        cols = st.columns(3)
        res_pratica = {}
        
        for i, (cat, subcats) in enumerate(CRITERIOS.items()):
            with cols[i]:
                st.markdown(f"**{cat}**")
                soma = 0
                for s in subcats:
                    val = st.radio(f"{s[:30]}...", [1, 3, 5], index=1, key=f"{formando}_{s}")
                    soma += val
                res_pratica[cat] = (soma / (len(subcats) * 5)) * 20

        # Cálculos
        m_pratica = (res_pratica["Operação com Ferramentas (60%)"] * 0.6) + \
                    (res_pratica["Manuseamento Equipamento (20%)"] * 0.2) + \
                    (res_pratica["Estabilização e Segurança (20%)"] * 0.2)
        
        n_final = (nota_t * 0.5) + (m_pratica * 0.5)
        
        if st.form_submit_button("Guardar e Adicionar ao PDF"):
            st.session_state.base_dados[formando] = {
                "Teórica": nota_t,
                "Média Prática": round(m_pratica, 2),
                "Nota Final": round(n_final, 2),
                "Situação": "APROVADO" if n_final >= 9.5 else "REPROVADO"
            }
            st.success(f"Avaliação de {formando} guardada!")

    # --- EXPORTAÇÃO ---
    if st.session_state.base_dados:
        st.divider()
        st.subheader("Gerar Documento Final")
        st.write(f"Total de formandos avaliados: {len(st.session_state.base_dados)}")
        
        if st.button("🚀 Unir tudo num PDF Final"):
            pdf_bytes = gerar_pdf_final(st.session_state.base_dados)
            st.download_button(
                label="📥 Descarregar PDF Único",
                data=pdf_bytes,
                file_name="Avaliacoes_Completas.pdf",
                mime="application/pdf"
            )
