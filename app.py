import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import load_workbook
from datetime import datetime

st.set_page_config(page_title="Sistema Integrado de Avaliação UFCD", layout="wide")

# --- CRITÉRIOS DE AVALIAÇÃO (Devem coincidir com o texto no Excel) ---
CRITERIOS = {
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

def preencher_ficha_individual(template_bytes, nome, dados):
    # keep_vba=True para manter as macros do modelo .xlsm
    wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    ws = wb.active 

    # 1. Nome e Data
    ws['C6'] = nome 
    ws['C8'] = datetime.now().strftime("%d/%m/%Y")

    # 2. Mapeamento das colunas de "X" (AH=34, AI=35, AJ=36)
    col_map = {1: 34, 3: 35, 5: 36}

    # 3. Marcar Cruzes e Totais por Categoria
    for cat_nome, lista_subs in CRITERIOS.items():
        soma_pontos = 0
        ultima_linha = 10
        for i, texto_sub in enumerate(lista_subs):
            valor = dados['pratica'][f"{cat_nome}_{i}"]
            soma_pontos += valor
            col_x = col_map[valor]
            
            # Localizar linha do critério
            for row in ws.iter_rows(min_row=10, max_row=80):
                if row[6].value and texto_sub[:30] in str(row[6].value): # Coluna G
                    ws.cell(row=row[6].row, column=col_x).value = "X"
                    ultima_linha = row[6].row
                    break
        
        # 4. Escrever Média (0-20) no campo "Classificação no parâmetro"
        media_parcial = (soma_pontos / (len(lista_subs) * 5)) * 20
        for row in ws.iter_rows(min_row=ultima_linha, max_row=ultima_linha+5):
            for cell in row:
                if cell.value and "Classificação no parâmetro" in str(cell.value):
                    ws.cell(row=cell.row, column=cell.column + 4).value = round(media_parcial, 2)
                    break

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# --- INTERFACE STREAMLIT ---
st.title("🚀 Portal de Avaliação UFCD 9889")

with st.sidebar:
    st.header("Upload de Modelos")
    f_import = st.file_uploader("1. Ficheiro Importação (Nomes Coluna K)", type=["xlsx"])
    f_modelo_macro = st.file_uploader("2. Modelo Ficha Prática (.xlsm)", type=["xlsm"])
    f_pauta_final = st.file_uploader("3. Pauta Final (.xlsx)", type=["xlsx"])

if f_import and f_modelo_macro and f_pauta_final:
    # Extrair nomes da Coluna K do ficheiro de importação (K13 em diante)
    df_nomes = pd.read_excel(f_import, skiprows=12, usecols="K").dropna()
    lista_nomes = df_nomes.iloc[:, 0].astype(str).tolist()
    
    formando = st.selectbox("Selecione o Formando para avaliar:", lista_nomes)

    if 'db' not in st.session_state: st.session_state.db = {}

    with st.form("avaliacao_completa"):
        st.subheader(f"Avaliação: {formando}")
        nota_teorica = st.number_input("Avaliação Teórica (0-20)", 0.0, 20.0, 10.0)
        
        st.divider()
        st.markdown("### Avaliação Prática (Cruzes)")
        c1, c2, c3 = st.columns(3)
        notas_p = {}
        
        for i, (cat, itens) in enumerate(CRITERIOS.items()):
            with [c1, c2, c3][i]:
                st.markdown(f"**{cat}**")
                for idx, item in enumerate(itens):
                    notas_p[f"{cat}_{idx}"] = st.radio(f"{item[:45]}...", [1, 3, 5], index=1, key=f"{formando}_{cat}_{idx}")

        if st.form_submit_button("💾 Guardar Avaliação"):
            # Cálculos automáticos para exportação posterior
            m_ferr = (sum([notas_p[f"Ferramentas_{i}"] for i in range(5)])/25)*20
            m_equip = (sum([notas_p[f"Equipamentos_{i}"] for i in range(5)])/25)*20
            m_estab = (sum([notas_p[f"Estabilização_{i}"] for i in range(5)])/25)*20
            
            st.session_state.db[formando] = {
                "teorica": nota_teorica,
                "pratica": notas_p,
                "m_ferr": m_ferr, "m_equip": m_equip, "m_estab": m_estab
            }
            st.success(f"Dados de {formando} guardados com sucesso!")

    # --- EXPORTAÇÃO FINAL ---
    if st.session_state.db:
        st.divider()
        if st.button("🚀 Gerar Dossier Completo (ZIP com Macros)"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                
                f_modelo_macro.seek(0)
                template_data = f_modelo_macro.read()
                
                for nome, dados in st.session_state.db.items():
                    # Gerar cada ficha .xlsm mantendo as macros
                    ficheiro_individual = preencher_ficha_individual(template_data, nome, dados)
                    zf.writestr(f"Ficha_Pratica_{nome.replace(' ', '_')}.xlsm", ficheiro_individual)
            
            st.download_button(
                label="📥 Descarregar ZIP das Avaliações",
                data=zip_buffer.getvalue(),
                file_name="Dossier_UFCD9889.zip",
                mime="application/zip"
            )
else:
    st.info("Aguardando carregamento dos 3 ficheiros necessários na barra lateral.")
