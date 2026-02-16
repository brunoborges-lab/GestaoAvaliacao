import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
import zipfile

st.set_page_config(page_title="Preenchimento Automático Excel", layout="wide")

st.title("✍️ Preenchimento Automático de Grelhas")
st.markdown("Este sistema pega no seu Excel original e desenha os 'X' nas colunas corretas.")

# --- 1. CONFIGURAÇÃO ---
with st.sidebar:
    st.header("Uploads")
    f_template = st.file_uploader("1. Modelo Vazio (.xlsx)", type=["xlsx"])
    f_import = st.file_uploader("2. Lista de Nomes (K13)", type=["xlsx", "xls"])

# --- CRITÉRIOS (Texto Exato para procura) ---
# DICA: O texto aqui tem de ser IGUAL ao que está no Excel para o robot o encontrar.
CRITERIOS_TEXTO = {
    "Ferramentas": [
        "Transporta as ferramentas e procede a abertura e fecho das mesmas em segurança", # Ajuste se necessário
        "Opera com a ferramenta prependicular ao obetivo de trabalho", # Notei "prependicular" no seu snippet
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

# --- FUNÇÃO MÁGICA DE PREENCHIMENTO ---
def preencher_excel(template_bytes, nome_formando, avaliacao_dict):
    # Carregar o Excel mantendo estilos
    wb = load_workbook(io.BytesIO(template_bytes))
    
    # Tentar encontrar a folha correta (Grelha Observação)
    sheet_name = None
    for sheet in wb.sheetnames:
        if "Grelha" in sheet or "M500" in sheet:
            sheet_name = sheet
            break
    
    if not sheet_name:
        return None
        
    ws = wb[sheet_name]

    # 1. Escrever o Nome (Procura pela célula que diz "NomeFormando" ou escreve numa posição fixa)
    # Estratégia de Busca: Procura a célula que contém "Nome do Formando" e escreve na seguinte
    nome_escrito = False
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and ("Nome" in cell.value or "Formando" in cell.value):
                # Escreve na célula à direita (offset column=1)
                ws.cell(row=cell.row, column=cell.column + 1).value = nome_formando
                nome_escrito = True
                break
        if nome_escrito: break
    
    # Se não encontrou lugar para o nome, tenta escrever numa célula comum (ajuste C3 se souber a exata)
    if not nome_escrito:
        ws['C3'] = nome_formando 

    # 2. Marcar os X
    # Definir colunas dos scores (Baseado no seu snippet, parecem estar à direita)
    # VOU ASSUMIR COLUNAS FIXAS baseadas na estrutura visual comum.
    # Se os X aparecerem no sítio errado, altere estes números:
    COL_1 = 34  # Ajustar conforme o Excel (Coluna AH?)
    COL_3 = 35  # Coluna AI?
    COL_5 = 36  # Coluna AJ?
    
    # Vamos tentar detetar as colunas dinamicamente procurando "1.0", "3.0", "5.0"
    header_row = 0
    for row in ws.iter_rows(min_row=1, max_row=10):
        for cell in row:
            if cell.value == 1.0: COL_1 = cell.column; header_row = cell.row
            if cell.value == 3.0: COL_3 = cell.column
            if cell.value == 5.0: COL_5 = cell.column

    # Preenchimento
    for grupo, criterios in CRITERIOS_TEXTO.items():
        for i, texto_criterio in enumerate(criterios):
            nota = avaliacao_dict.get(f"{grupo}_{i}", 3) # Default é 3
            
            # Escolher coluna alvo
            col_alvo = COL_3
            if nota == 1: col_alvo = COL_1
            if nota == 5: col_alvo = COL_5
            
            # Procurar a linha que tem este texto
            for row in ws.iter_rows(min_row=header_row, max_row=60, min_col=1, max_col=10):
                for cell in row:
                    # Uso de "in" para ser flexível caso haja espaços extras
                    if cell.value and isinstance(cell.value, str) and texto_criterio[:20] in cell.value:
                        ws.cell(row=cell.row, column=col_alvo).value = "X"
                        # Limpar as outras (opcional, caso o template já tenha X)
                        break

    # Guardar em memória
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- INTERFACE ---
if f_template and f_import:
    # Ler Nomes
    df_nomes = pd.read_excel(f_import, skiprows=12, usecols="K").dropna()
    lista_nomes = df_nomes.iloc[:, 0].unique().tolist()
    
    # Estado da Sessão para guardar avaliações
    if 'dados_excel' not in st.session_state:
        st.session_state.dados_excel = {}

    col_esq, col_dir = st.columns([1, 2])

    with col_esq:
        st.subheader("Selecione o Formando")
        formando = st.selectbox("Formando Atual:", lista_nomes)
        
        st.info("Preencha a grelha ao lado e clique em 'Guardar'. No final, poderá descarregar todos os Excel num ficheiro ZIP.")

    with col_dir:
        with st.form("form_x"):
            st.subheader(f"Avaliação de: {formando}")
            
            notas_temp = {}
            
            # Gerar formulário dinâmico
            for grupo, lista in CRITERIOS_TEXTO.items():
                st.markdown(f"**{grupo}**")
                for i, texto in enumerate(lista):
                    # Slider de 1, 3, 5
                    val = st.select_slider(f"...{texto[:40]}", options=[1, 3, 5], value=3, key=f"{formando}_{grupo}_{i}")
                    notas_temp[f"{grupo}_{i}"] = val
            
            if st.form_submit_button("💾 Guardar e Processar Excel"):
                # Gerar o Excel Individual IMEDIATAMENTE em memória
                f_template.seek(0) # Reiniciar leitura do template
                excel_preenchido = preencher_excel(f_template.read(), formando, notas_temp)
                
                if excel_preenchido:
                    st.session_state.dados_excel[formando] = excel_preenchido
                    st.success(f"Ficheiro de {formando} gerado com sucesso!")
                else:
                    st.error("Não encontrei a folha 'Grelha Observação' no modelo.")

    # --- ÁREA DE DOWNLOAD (ZIP) ---
    if st.session_state.dados_excel:
        st.divider()
        st.write(f"📂 Tem {len(st.session_state.dados_excel)} ficheiros prontos.")
        
        # Criar ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for nome, dados in st.session_state.dados_excel.items():
                # Nome do ficheiro limpo
                nome_limpo = "".join([c for c in nome if c.isalnum() or c in (' ', '_')]).strip()
                zf.writestr(f"Avaliacao_{nome_limpo}.xlsx", dados)
        
        st.download_button(
            label="📦 Descarregar Todos os Ficheiros (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Todas_Avaliacoes_Preenchidas.zip",
            mime="application/zip",
            type="primary"
        )
