import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import load_workbook

st.set_page_config(page_title="Gestor Avaliação", layout="wide")

# --- DEFINIÇÃO DOS TEXTOS PARA BUSCA (Devem ser idênticos ao Excel) ---
CRITERIOS_EXCEL = {
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

def processar_modelo_macro(template_bytes, nome_aluno, notas_individuais):
    # keep_vba=True é CRÍTICO para não apagar as macros do ficheiro
    wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    ws = wb.active 

    # 1. Colocar o Nome (Procura campo NomeFormando ou similar)
    ws['C6'] = nome_aluno # Ajustar para a célula exata do seu modelo

    # 2. Mapeamento de Colunas AH=34 (1.0), AI=35 (3.0), AJ=36 (5.0)
    col_map = {1: 34, 3: 35, 5: 36}

    # 3. Marcar Cruzes e Somar Totais
    for cat, lista in CRITERIOS_EXCEL.items():
        soma_cat = 0
        for i, texto in enumerate(lista):
            nota = notas_individuais[f"{cat}_{i}"]
            soma_cat += nota
            col_x = col_map[nota]
            
            # Busca a linha do critério (procura nas colunas A a G)
            for row in ws.iter_rows(min_row=1, max_row=100, min_col=1, max_col=10):
                for cell in row:
                    if cell.value and texto[:25] in str(cell.value):
                        ws.cell(row=cell.row, column=col_x).value = "X"
                        break
        
        # 4. Escrever Média da Categoria (Ex: 0-20)
        media_final_cat = (soma_cat / (len(lista) * 5)) * 20
        # Procura a célula "Classificação no parâmetro"
        for row in ws.iter_rows(min_row=1, max_row=100):
            for cell in row:
                if cell.value and "Classificação no parâmetro" in str(cell.value):
                    # Escreve o valor 4 colunas à frente do texto (ajustar se necessário)
                    ws.cell(row=cell.row, column=cell.column + 4).value = round(media_final_cat, 2)

    # Guardar como Bytes
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- INTERFACE STREAMLIT ---
st.title("📂 Gerador de Avaliação")

with st.sidebar:
    f_xlsm = st.file_uploader("Modelo Original (.xlsm)", type=["xlsm"])
    f_nomes = st.file_uploader("Ficheiro de Importação (K13)", type=["xlsx"])

if f_xlsm and f_nomes:
    df = pd.read_excel(f_nomes, skiprows=12, usecols="K").dropna()
    formando = st.selectbox("Escolha o Formando:", df.iloc[:, 0].tolist())

    with st.form("avaliacao_tecnica"):
        st.subheader(f"Avaliar: {formando}")
        notas_form = {}
        c1, c2, c3 = st.columns(3)
        
        for i, (cat, itens) in enumerate(CRITERIOS_EXCEL.items()):
            with [c1, c2, c3][i]:
                st.
