import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import load_workbook

st.set_page_config(page_title="Automação UFCD 9889", layout="wide")

# --- BASE DE DADOS EM MEMÓRIA ---
if 'historico_avaliacoes' not in st.session_state:
    st.session_state.historico_avaliacoes = {}

# --- CRITÉRIOS TÉCNICOS (Devem ser iguais ao texto no Excel) ---
SUB_CATEGORIAS = {
    "Ferramentas": [
        "Transporta as ferramentas e procede a abertura e fecho das mesmas em segurança",
        "Opera com a ferramenta prependicular ao obetivo de trabalho",
        "Coloca-se do lado certo da ferramenta",
        "Efectua cominucação sobre abertura ou corte de estruturas do veiculo",
        "Protege a(s) vítima(s) e o(s) socorrista(s) com proteção rigida"
    ],
    "Equipamentos": [
        "Escolhe  equipamento adequado à função",
        "Transporta  e opera os equipamentos em segurança",
        "Opera corretamente com o grupo energetico",
        "Opera corretamente com equipamento de estabilização",
        "Opera corretamente equipamento pneumático"
    ],
    "Estabilização": [
        "Sinaliza e delimita zonas de trabalho e zela pela segurança",
        "Estabiliza o(s) veículo(s) acidentado(s) de forma adequada",
        "Controla estabilização inicial e efetua estabilização progressiva",
        "Efetua limpeza da zona de trabalho",
        "Aplica as proteções nos pontos agressivos"
    ]
}

def preencher_grelha_com_cruzes(template_bytes, nome, dados):
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active # Assume que a grelha de observação é a folha ativa
    
    # 1. Inserir Nome do Formando (Procura célula que contém "Nome")
    for row in ws.iter_rows(max_row=15):
        for cell in row:
            if cell.value and "Nome" in str(cell.value):
                ws.cell(row=cell.row, column=cell.column + 1).value = nome
                break

    # 2. Definir Colunas das Cruzes (Baseado no 1.0, 3.0, 5.0 do seu ficheiro)
    # Estas colunas costumam ser AH (34), AI (35), AJ (36)
    col_map = {1: 34, 3: 35, 5: 36} 

    # 3. Colocar os "X" em cada subcategoria
    for cat_nome, lista_subs in SUB_CATEGORIAS.items():
        for i, texto_sub in enumerate(lista_subs):
            valor_x = dados['detalhe_pratica'][f"{cat_nome}_{i}"]
            col_alvo = col_map.get(valor_x, 35)
            
            # Procura a linha que contém o texto da subcategoria
            for row in ws.iter_rows(min_row=1, max_row=100, min_col=1, max_col=15):
                for cell in row:
                    if cell.value and texto_sub[:30] in str(cell.value):
                        ws.cell(row=cell.row, column=col_alvo).value = "X"
                        break
    
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# --- INTERFACE ---
st.title("📋 Avaliação Prática Detalhada com Geração de Excel")

with st.sidebar:
    f_modelo = st.file_uploader("1. Modelo de Ficha Prática (.xlsx)", type=["xlsx"])
    f_import = st.file_uploader("2. Ficheiro de Importação (K13)", type=["xlsx"])

if f_modelo and f_import:
    df_nomes = pd.read_excel(f_import, skiprows=12, usecols="K").dropna()
    nomes = df_nomes.iloc[:, 0].astype(str).str.strip().tolist()
    
    formando = st.selectbox("Seleccione o Formando:", nomes)

    with st.form("form_detalhado"):
        col_t, col_p = st.columns([1, 3])
        
        with col_t:
            st.subheader("Teórica")
            nota_t = st.number_input("Nota (0-20)", 0.0, 20.0, 10.0)
        
        with col_p:
            st.subheader("Prática (Subcategorias)")
            c1, c2, c3 = st.columns(3)
            detalhe_notas = {}
            
            for i, (cat, lista) in enumerate(SUB_CATEGORIAS.items()):
                target_col = [c1, c2, c3][i]
                with target_col:
                    st.markdown(f"**{cat}**")
                    for idx, sub in enumerate(lista):
                        # Escala de cruzes: 1, 3 ou 5
                        detalhe_notas[f"{cat}_{idx}"] = st.radio(f"{sub[:45]}...", [1, 3, 5], index=1, horizontal=True, key=f"{formando}_{cat}_{idx}")

        if st.form_submit_button("💾 Guardar e Gerar Ficheiros"):
            # Cálculos de médias para o resumo
            # (Simplificado: média aritmética dos X convertida para 0-20)
            st.session_state.historico_avaliacoes[formando] = {
                "nota_teorica": nota_t,
                "detalhe_pratica": detalhe_notas,
                "nota_final": 0.0 # Calculado na exportação
            }
            st.success(f"Dados de {formando} guardados!")

    # --- EXPORTAÇÃO ZIP ---
    if st.session_state.historico_avaliacoes:
        st.divider()
        if st.button("🚀 Gerar Dossier Final (ZIP com todos os Excels preenchidos)"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                
                f_modelo.seek(0)
                modelo_bytes = f_modelo.read()
                
                for nome, dados in st.session_state.historico_avaliacoes.items():
                    # Gerar cada Excel individual com as cruzes
                    excel_aluno = preencher_grelha_com_cruzes(modelo_bytes, nome, dados)
                    zf.writestr(f"Ficha_Pratica_{nome}.xlsx", excel_aluno)
                
            st.download_button("📥 Descarregar ZIP das Avaliações", zip_buffer.getvalue(), "Dossier_UFCD9889.zip", type="primary")

else:
    st.info("Por favor, carregue os dois ficheiros na barra lateral.")
