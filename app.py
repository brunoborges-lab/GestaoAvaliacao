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
                    val =
