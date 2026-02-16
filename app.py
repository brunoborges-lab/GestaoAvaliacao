import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import zipfile

st.set_page_config(page_title="Preenchimento Automático Grelhas", layout="wide")

st.title("🤖 Preenchimento Automático de Pautas")
st.markdown("Preencha as notas na aplicação e o sistema escreve diretamente nos seus ficheiros Excel originais.")

# --- 1. GESTÃO DE ESTADO (MEMÓRIA) ---
if 'db_notas' not in st.session_state:
    st.session_state.db_notas = {}

# --- 2. UPLOAD DOS FICHEIROS ---
with st.sidebar:
    st.header("Carregar Ficheiros Originais")
    file_grelha = st.file_uploader("1. Grelha de Avaliação (Modelo)", type=["xlsx"])
    file_import = st.file_uploader("2. Ficheiro de Importação (Nomes)", type=["xlsx"])

# --- 3. FUNÇÕES DE PROCESSAMENTO EXCEL ---
def encontrar_coluna(ws, texto_procura, linha_max=15):
    """Procura em que coluna está um determinado cabeçalho"""
    for row in ws.iter_rows(min_row=1, max_row=linha_max):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                if texto_procura.lower() in cell.value.lower():
                    return cell.column
    return None

def processar_ficheiros(bytes_grelha, bytes_import, dados_alunos):
    # --- A. PREENCHER A GRELHA DE AVALIAÇÃO ---
    wb_grelha = load_workbook(io.BytesIO(bytes_grelha))
    
    # Tenta encontrar a folha ativa ou a primeira
    ws_grelha = wb_grelha.active
    
    # Mapeamento Inteligente de Colunas (Procura pelos cabeçalhos)
    col_teorica = encontrar_coluna(ws_grelha, "teórica") or 22  # Default aprox. V
    col_ferramentas = encontrar_coluna(ws_grelha, "ferramentas") or 32 # Default aprox. AF
    col_equipamentos = encontrar_coluna(ws_grelha, "equipamento") or 42 # Default aprox. AP
    col_estabilizacao = encontrar_coluna(ws_grelha, "estabilização") or 52 # Default aprox. AZ
    col_nome_grelha = 3 # Assumindo Coluna C para nomes na Grelha

    # Escrever na Grelha
    # Itera sobre as linhas da grelha para encontrar o aluno correspondente
    for row in ws_grelha.iter_rows(min_row=10, max_row=100):
        celula_nome = row[col_nome_grelha - 1] # Ajuste de índice 0
        nome_grelha = str(celula_nome.value).strip() if celula_nome.value else ""
        
        # Se encontrarmos o nome na nossa base de dados
        if nome_grelha in dados_alunos:
            aluno = dados_alunos[nome_grelha]
            
            # Escrever valores nas colunas detetadas
            ws_grelha.cell(row=celula_nome.row, column=col_teorica).value = aluno['teorica']
            ws_grelha.cell(row=celula_nome.row, column=col_ferramentas).value = aluno['ferramentas']
            ws_grelha.cell(row=celula_nome.row, column=col_equipamentos).value = aluno['equipamentos']
            ws_grelha.cell(row=celula_nome.row, column=col_estabilizacao).value = aluno['estabilizacao']

    # --- B. PREENCHER O FICHEIRO DE IMPORTAÇÃO (NOTA FINAL) ---
    wb_import = load_workbook(io.BytesIO(bytes_import))
    ws_import = wb_import.active # Ou procurar folha específica
    
    # Identificar coluna de nomes (K = 11) e onde pôr a nota
    col_k_idx = 11
    col_destino_nota = encontrar_coluna(ws_import, "final") or 12 # Tenta achar "Nota Final" ou usa L (12)
    
    for row in ws_import.iter_rows(min_row=13, max_row=200): # Começa na linha 13
        celula_nome = row[col_k_idx - 1] # Coluna K
        nome_imp = str(celula_nome.value).strip() if celula_nome.value else ""
        
        if nome_imp in dados_alunos:
            nota_final = dados_alunos[nome_imp]['final']
            # Escreve a nota final na coluna definida
            cell_nota = ws_import.cell(row=celula_nome.row, column=col_destino_nota)
            cell_nota.value = nota_final
            
            # Formatação opcional (Verde se passou, Vermelho se reprovou)
            if nota_final < 9.5:
                cell_nota.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

    # --- SALVAR E RETORNAR ---
    out_grelha = io.BytesIO()
    wb_grelha.save(out_grelha)
    
    out_import = io.BytesIO()
    wb_import.save(out_import)
    
    return out_grelha.getvalue(), out_import.getvalue()

# --- 4. INTERFACE DE ENTRADA DE DADOS ---
if file_grelha and file_import:
    # Carregar lista de formandos
    df_nomes = pd.read_excel(file_import, skiprows=12, usecols="K").dropna()
    lista_nomes = df_nomes.iloc[:, 0].astype(str).str.strip().tolist()

    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("1. Selecione o Formando")
        aluno_sel = st.selectbox("Formando:", lista_nomes)
        
        # Mostrar progresso
        total = len(lista_nomes)
        preenchidos = len(st.session_state.db_notas)
        st.progress(preenchidos / total)
        st.caption(f"{preenchidos} de {total} avaliados")

    with col2:
        st.subheader("2. Lance as Notas")
        with st.form("form_notas"):
            nt = st.number_input("Avaliação Teórica (0-20)", 0.0, 20.0, step=0.1)
            
            st.markdown("**Avaliação Prática (0-20)**")
            c1, c2, c3 = st.columns(3)
            with c1: nf = st.number_input("Ferramentas (60%)", 0.0, 20.0)
            with c2: ne = st.number_input("Equipamentos (20%)", 0.0, 20.0)
            with c3: ns = st.number_input("Estabilização (20%)", 0.0, 20.0)
            
            # Cálculo Prévio
            media_pratica = (nf * 0.6) + (ne * 0.2) + (ns * 0.2)
            nota_final = (nt * 0.5) + (media_pratica * 0.5)
            
            if st.form_submit_button("💾 Registar Notas"):
                st.session_state.db_notas[aluno_sel] = {
                    "teorica": nt,
                    "ferramentas": nf,
                    "equipamentos": ne,
                    "estabilizacao": ns,
                    "final": round(nota_final, 2)
                }
                st.success(f"Nota de {aluno_sel} registada: {nota_final:.2f}")

    # --- 5. BOTÃO FINAL DE PROCESSAMENTO ---
    st.divider()
    if st.session_state.db_notas:
        st.subheader("3. Exportação Final")
        st.write("Quando terminar de lançar as notas de todos os alunos, clique abaixo.")
        
        if st.button("🚀 Preencher Ficheiros Excel Originais"):
            file_grelha.seek(0); file_import.seek(0) # Reset ponteiros
            
            try:
                novo_grelha, novo_import = processar_ficheiros(
                    file_grelha.read(), 
                    file_import.read(), 
                    st.session_state.db_notas
                )
                
                # Criar ZIP com os dois ficheiros
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    zf.writestr("Grelha_Avaliacao_Preenchida.xlsx", novo_grelha)
                    zf.writestr("Importacao_Com_Notas_Finais.xlsx", novo_import)
                
                st.download_button(
                    label="📦 Descarregar Ficheiros Preenchidos (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Ficheiros_Finais_UFCD.zip",
                    mime="application/zip",
                    type="primary"
                )
            except Exception as e:
                st.error(f"Erro ao processar ficheiros: {e}")
                st.info("Verifique se os ficheiros originais não estão protegidos por palavra-passe ou corrompidos.")

else:
    st.info("Por favor, carregue a Grelha Modelo e o Ficheiro de Importação na barra lateral.")
