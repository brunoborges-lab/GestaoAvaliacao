import streamlit as st
import pandas as pd
import io
from fpdf import FPDF

# Configuração da Página
st.set_page_config(page_title="Gerador de Fichas UFCD", layout="wide")

# Inicializar a base de dados na sessão para não perder dados ao mudar de formando
if 'avaliacoes_completas' not in st.session_state:
    st.session_state.avaliacoes_completas = {}

# --- ESTRUTURA DE CRITÉRIOS (Conforme a sua Grelha M500) ---
CRITERIOS = {
    "OPERAÇÃO COM FERRAMENTAS (60%)": [
        "Transporta as ferramentas e procede a abertura e fecho em segurança",
        "Opera com a ferramenta perpendicular ao objetivo de trabalho",
        "Coloca-se do lado certo da ferramenta",
        "Efetua comunicação sobre abertura ou corte de estruturas",
        "Protege a(s) vítima(s) e o(s) socorrista(s) com proteção rígida"
    ],
    "MANUSEAMENTO DE EQUIPAMENTO (20%)": [
        "Escolhe equipamento adequado à função",
        "Transporta e opera os equipamentos em segurança",
        "Opera corretamente com o grupo energético",
        "Opera corretamente com equipamento de estabilização",
        "Opera corretamente equipamento pneumático"
    ],
    "ESTABILIZAÇÃO E SEGURANÇA (20%)": [
        "Sinaliza e delimita zonas de trabalho e zela pela segurança",
        "Estabiliza o(s) veículo(s) acidentado(s) de forma adequada",
        "Controla estabilização inicial e efetua estabilização progressiva",
        "Efetua limpeza da zona de trabalho",
        "Aplica as proteções nos pontos agressivos"
    ]
}

# --- CLASSE PARA GERAR O DOCUMENTO PDF ---
class PDF(FPDF):
    def header(self):
        # Título do Documento
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'FICHA DE AVALIAÇÃO PRÁTICA', 0, 1, 'C')
        self.set_font('Arial', '', 10)
        self.cell(0, 5, 'UFCD 9889 - SALVAMENTO RODOVIÁRIO - INICIAÇÃO', 0, 1, 'C')
        self.ln(10)

    def ficha_formando(self, nome, dados):
        self.add_page()
        # Cabeçalho do Formando
        self.set_fill_color(230, 230, 230)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, f" FORMANDO: {nome.upper()}", 1, 1, 'L', True)
        self.ln(5)

        # Detalhe das Notas por Categoria
        for cat, nota in dados['medias_parciais'].items():
            self.set_font('Arial', 'B', 10)
            self.cell(150, 8, cat, 1, 0)
            self.cell(40, 8, f"{nota:.2f} / 20", 1, 1, 'C')

        self.ln(10)
        # Resultados Finais
        self.set_font('Arial', 'B', 12)
        self.cell(95, 12, f"MÉDIA PRÁTICA: {dados['media_pratica']:.2f}", 1, 0, 'C')
        self.cell(95, 12, f"NOTA TEÓRICA: {dados['nota_teorica']:.2f}", 1, 1, 'C')
        
        self.set_fill_color(200, 255, 200) if dados['nota_final'] >= 9.5 else self.set_fill_color(255, 200, 200)
        self.cell(0, 15, f"CLASSIFICAÇÃO FINAL: {dados['nota_final']:.2f} - {dados['situacao']}", 1, 1, 'C', True)
        
        # Espaço para Assinaturas
        self.ln(20)
        self.set_font('Arial', 'I', 8)
        self.cell(95, 10, "__________________________________", 0, 0, 'C')
        self.cell(95, 10, "__________________________________", 0, 1, 'C')
        self.cell(95, 5, "O Formador", 0, 0, 'C')
        self.cell(95, 5, "O Formando", 0, 1, 'C')

# --- INTERFACE STREAMLIT ---
st.title("🚀 Sistema de Emissão de Fichas PDF")

with st.sidebar:
    st.header("Configuração Base")
    f_import = st.file_uploader("Ficheiro de Importação (K13)", type=["xlsx"])

if f_import:
    df_nomes = pd.read_excel(f_import, skiprows=12, usecols="K").dropna()
    lista_nomes = df_nomes.iloc[:, 0].unique().tolist()
    
    formando = st.selectbox("Escolha o Formando para avaliar:", lista_nomes)

    with st.form("form_pdf"):
        col_t, col_p = st.columns([1, 2])
        
        with col_t:
            st.subheader("Teórica")
            nota_t = st.number_input("Nota Teste", 0.0, 20.0, 10.0)

        with col_p:
            st.subheader("Prática - Itens de Observação")
            notas_input = {}
            for cat, subcats in CRITERIOS.items():
                st.markdown(f"**{cat}**")
                soma_cat = 0
                for sub in subcats:
                    # Escala 1, 3, 5 conforme o seu ficheiro
                    valor = st.select_slider(f"{sub}", options=[1, 3, 5], value=3, key=f"{formando}_{sub}")
                    soma_cat += valor
                # Converter escala 1-5 para 0-20
                notas_input[cat] = (soma_cat / (len(subcats) * 5)) * 20
        
        if st.form_submit_button("✅ Guardar Avaliação"):
            # Cálculos Finais
            m_pratica = (notas_input["OPERAÇÃO COM FERRAMENTAS (60%)"] * 0.6) + \
                        (notas_input["MANUSEAMENTO DE EQUIPAMENTO (20%)"] * 0.2) + \
                        (notas_input["ESTABILIZAÇÃO E SEGURANÇA (20%)"] * 0.2)
            
            n_final = (nota_t * 0.5) + (m_pratica * 0.5)
            
            st.session_state.avaliacoes_completas[formando] = {
                "nota_teorica": nota_t,
                "medias_parciais": notas_input,
                "media_pratica": m_pratica,
                "nota_final": n_final,
                "situacao": "APROVADO" if n_final >= 9.5 else "NÃO APROVADO"
            }
            st.success(f"Avaliação de {formando} registada!")

    # --- ZONA DE EXPORTAÇÃO ---
    if st.session_state.avaliacoes_completas:
        st.divider()
        st.subheader(f"📦 Finalização ({len(st.session_state.avaliacoes_completas)} formandos prontos)")
        
        if st.button("🛠️ Gerar PDF Único com todas as Fichas"):
            pdf = PDF()
            for nome, dados in st.session_state.avaliacoes_completas.items():
                pdf.ficha_formando(nome, dados)
            
            pdf_output = pdf.output(dest='S').encode('latin-1')
            st.download_button(
                label="📥 Descarregar Dossier de Avaliação (PDF)",
                data=pdf_output,
                file_name="Fichas_Avaliacao_UFCD9889.pdf",
                mime="application/pdf"
            )
