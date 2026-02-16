import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import load_workbook
from datetime import datetime

st.set_page_config(page_title="Sistema Integrado UFCD 9889", layout="wide")

# --- ESTADO DA SESSÃO ---
if 'avaliacoes' not in st.session_state:
    st.session_state.avaliacoes = {}

# --- CRITÉRIOS TÉCNICOS EXTRAÍDOS DO MODELO ---
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

def processar_ficheiro_macro(template_bytes, nome, dados):
    # keep_vba=True preserva as macros do seu modelo .xlsm
    wb = load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    ws = wb.active # Folha "10-Grelha Observação M500"

    # 1. Identificação do Formando e Data
    # Ajuste as coordenadas das células conforme o seu modelo real
    ws['C6'] = nome 
    ws['C8'] = datetime.now().strftime("%d/%m/%Y")

    # 2. Mapeamento das colunas de "X" (AH=34, AI=35, AJ=36)
    col_map = {1: 34, 3: 35, 5: 36}

    # 3. Preenchimento de Subcategorias (Cruzes)
    for cat_nome, lista_subs in CRITERIOS.items():
        soma_pontos = 0
        ultima_linha = 0
        for i, texto_sub in enumerate(lista_subs):
            valor = dados['pratica'][f"{cat_nome}_{i}"]
            soma_pontos += valor
            col_x = col_map[valor]
            
            # Localizar linha do critério por texto
            for row in ws.iter_rows(min_row=10, max_row=80, min_col=1, max_col=10):
                if row[6].value and texto_sub[:30] in str(row[6].value): # Coluna G
                    ws.cell(row=row[6].row, column=col_x).value = "X"
                    ultima_linha = row[6].row
                    break
        
        # 4. Escrever Média do Parâmetro (0-20)
        media_parcial = (soma_pontos / (len(lista_subs) * 5)) * 20
        # Procura a célula de classificação logo abaixo do bloco
        for row in ws.iter_rows(min_row=ultima_linha, max_row=ultima_linha+5):
            for cell in row:
                if cell.value and "Classificação no parâmetro" in str(cell.value):
                    ws.cell(row=cell.row, column=cell.column + 4).value = round(media_parcial, 2)
                    break

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def atualizar_avaliacao_final(final_bytes, dados_todos):
    wb = load_workbook(io.BytesIO(final_bytes))
    ws = wb.active # Grelha de apoio à classificação
    
    # Mapeamento de colunas no ficheiro final (conforme o snippet fornecido)
    # Avaliação Teórica: Coluna U (21), Ferramentas: AF (32), Equip: AP (42), Estab: AZ (52)
    for row in ws.iter_rows(min_row=12, max_row=100):
        nome_celula = row[2].value # Coluna C
        if nome_celula and str(nome_celula).strip() in dados_todos:
            d = dados_todos[str(nome_celula).strip()]
            
            # Exportar Notas
            ws.cell(row=row[0].row, column=21).value = d['teorica']
            ws.cell(row=row[0].row, column=32).value = d['media_ferr']
            ws.cell(row=row[0].row, column=42).value = d['media_equip']
            ws.cell(row=row[0].row, column=52).value = d['media_estab']

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# --- INTERFACE ---
st.title("⚙️ Gestor de Avaliações UFCD 9889")

with st.sidebar:
    st.header("Modelos")
    f_ficha_macro = st.file_uploader("Ficha Prática (.xlsm)", type=["xlsm"])
    f_pauta_final = st.file_uploader("Pauta de Avaliação Final (.xlsx)", type=["xlsx"])

if f_ficha_macro and f_pauta_final:
    # Ler nomes da pauta final (Coluna C)
    df_nomes = pd.read_excel(f_pauta_final, skiprows=10, usecols="C").dropna()
    lista_nomes = df_nomes.iloc[:, 0].tolist()
    
    formando = st.selectbox("Escolha o Formando:", lista_nomes)

    with st.form("avaliacao_total"):
        st.subheader(f"Registo: {formando}")
        
        # CAMPO AVALIAÇÃO TEÓRICA
        nota_teorica = st.number_input("Avaliação Teórica (0-20)", 0.0, 20.0, 10.0, step=0.1)
        
        st.divider()
        st.markdown("### Avaliação Prática")
        c1, c2, c3 = st.columns(3)
        notas_p = {}
        
        for i, (cat, itens) in enumerate(CRITERIOS.items()):
            with [c1, c2, c3][i]:
                st.write(f"**{cat}**")
                for idx, item in enumerate(itens):
                    notas_p[f"{cat}_{idx}"] = st.radio(f"{item[:40]}...", [1, 3, 5], index=1, key=f"{formando}_{cat}_{idx}")

        if st.form_submit_button("💾 Guardar e Calcular"):
            # Cálculos de Médias Parciais para Exportação
            m_ferr = (sum([notas_p[f"Ferramentas_{i}"] for i in range(5)])/25)*20
            m_equip = (sum([notas_p[f"Equipamentos_{i}"] for i in range(5)])/25)*20
            m_estab = (sum([notas_p[f"Estabilização_{i}"] for i in range(5)])/25)*20
            
            st.session_state.avaliacoes[formando] = {
                "teorica": nota_teorica,
                "pratica": notas_p,
                "media_ferr": m_ferr,
                "media_equip": m_equip,
                "media_estab": m_estab
            }
            st.success(f"Avaliação de {formando} concluída!")

    # --- EXPORTAÇÃO ---
    if st.session_state.avaliacoes:
        st.divider()
        if st.button("🚀 Gerar Dossier e Atualizar Avaliação Final"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                
                # 1. Gerar as Fichas Individuais (.xlsm)
                f_ficha_macro.seek(0)
                template_macro = f_ficha_macro.read()
                for nome, dados in st.session_state.avaliacoes.items():
                    xlsm_preenchido = processar_ficheiro_macro(template_macro, nome, dados)
                    zf.writestr(f"Ficha_Pratica_{nome}.xlsm", xlsm_preenchido)
                
                # 2. Atualizar o Ficheiro de Avaliação Final
                f_pauta_final.seek(0)
                final_preenchido = atualizar_avaliacao_final(f_pauta_final.read(), st.session_state.avaliacoes)
                zf.writestr("Avaliacao_Final_Consolidada.xlsx", final_preenchido)
            
            st.download_button("📥 Descarregar Dossier Completo (ZIP)", zip_buffer.getvalue(), "Dossier_UFCD9889.zip")
