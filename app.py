import streamlit as st
import pandas as pd
import io
import xlsxwriter

# Configuração da Página
st.set_page_config(page_title="Gestor Avaliação UFCD 9889", layout="wide")

st.title("🚒 Gestor de Avaliação - Salvamento Rodoviário (UFCD 9889)")
st.markdown("### Preencha os dados uma única vez e gere todos os relatórios.")

# --- DADOS DOS FORMANDOS (Extraídos do seu ficheiro 1) ---
# Lista limpa de nomes baseada no seu upload
nomes_formandos = [
    "Cátia Filipa Vilar Regufe",
    "Diogo Ferreira Soares",
    "Erlisson Ribeiro de Oliveira Rocha",
    "Fátima Milena Rocha Abreu de Oliveira",
    "Fernando Bravo Figueroa",
    "Francisco Manuel da Fonseca Ferreira",
    "Jorge Miguel Ferreira Maranha Soares Miranda",
    "José Duarte da Costa Machado Brás",
    "Paulo Jorge Oliveira Maio",
    "Pedro Alexandre Rodrigues Costa Ferreira"
]

# Criar DataFrame inicial
df = pd.DataFrame(nomes_formandos, columns=["Nome do Formando"])

# --- BARRA LATERAL (CONFIGURAÇÕES) ---
st.sidebar.header("⚙️ Configuração de Ponderações")
peso_teorica = st.sidebar.slider("Peso Avaliação Teórica (%)", 0, 100, 50) / 100
peso_pratica = st.sidebar.slider("Peso Avaliação Prática (%)", 0, 100, 50) / 100

st.sidebar.subheader("Sub-parâmetros Prática (File 3)")
peso_ferramentas = st.sidebar.number_input("Peso Operação Ferramentas", 0.0, 1.0, 0.6)
peso_equipamentos = st.sidebar.number_input("Peso Manuseamento Equip.", 0.0, 1.0, 0.2)
peso_estabilizacao = st.sidebar.number_input("Peso Estabilização/Seg.", 0.0, 1.0, 0.2)

# --- ÁREA DE INSERÇÃO DE DADOS ---
st.info("👇 Insira as notas abaixo (0 a 20). A Nota Final e a Situação são calculadas automaticamente.")

# Adicionar colunas para entrada de dados
# Inicializamos com valores padrão para facilitar a edição
df["Nota Teórica"] = 0.0
df["Prática: Ferramentas (0-20)"] = 0.0
df["Prática: Equipamentos (0-20)"] = 0.0
df["Prática: Estabilização (0-20)"] = 0.0

# Editor de dados interativo
edited_df = st.data_editor(
    df,
    column_config={
        "Nota Teórica": st.column_config.NumberColumn(min_value=0, max_value=20, step=0.1, format="%.1f"),
        "Prática: Ferramentas (0-20)": st.column_config.NumberColumn(min_value=0, max_value=20, step=0.1, help="Baseado na grelha de observação"),
        "Prática: Equipamentos (0-20)": st.column_config.NumberColumn(min_value=0, max_value=20, step=0.1),
        "Prática: Estabilização (0-20)": st.column_config.NumberColumn(min_value=0, max_value=20, step=0.1),
    },
    hide_index=True,
    num_rows="dynamic",
    use_container_width=True
)

# --- CÁLCULOS ---
# 1. Calcular Nota Prática Ponderada
edited_df["Nota Prática Final"] = (
    (edited_df["Prática: Ferramentas (0-20)"] * peso_ferramentas) +
    (edited_df["Prática: Equipamentos (0-20)"] * peso_equipamentos) +
    (edited_df["Prática: Estabilização (0-20)"] * peso_estabilizacao)
)

# 2. Calcular Nota Final do Curso
edited_df["Classificação Final"] = (
    (edited_df["Nota Teórica"] * peso_teorica) +
    (edited_df["Nota Prática Final"] * peso_pratica)
)

# 3. Definir Situação e Menção Qualitativa
def get_situacao(nota):
    return "APROVADO" if nota >= 9.5 else "NÃO APROVADO"

def get_qualitativa(nota):
    if nota < 9.5: return "Insuficiente"
    if nota < 13.5: return "Suficiente"
    if nota < 17.5: return "Bom"
    return "Muito Bom"

edited_df["Situação"] = edited_df["Classificação Final"].apply(get_situacao)
edited_df["Menção"] = edited_df["Classificação Final"].apply(get_qualitativa)

# Mostrar tabela de resultados
st.markdown("### 📊 Pré-visualização dos Resultados")
st.dataframe(edited_df.style.format({
    "Nota Teórica": "{:.2f}",
    "Nota Prática Final": "{:.2f}",
    "Classificação Final": "{:.2f}"
}).applymap(lambda x: 'background-color: #ffcccc' if x == 'NÃO APROVADO' else 'background-color: #ccffcc', subset=['Situação']), use_container_width=True)

# --- EXPORTAÇÃO EXCEL ---
st.markdown("---")
st.header("📥 Download dos Ficheiros")

def to_excel(df):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Formatos
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    cell_fmt = workbook.add_format({'border': 1})
    num_fmt = workbook.add_format({'border': 1, 'num_format': '0.00'})
    
    # --- FOLHA 1: Grelha Avaliação Final (Documento 7) ---
    ws1 = workbook.add_worksheet("Grelha Avaliação Final")
    ws1.write(0, 0, "GRELHA DE AVALIAÇÃO FINAL - UFCD 9889", header_fmt)
    
    headers1 = ["Nome do Formando", "Avaliação Qualitativa", "Avaliação Quantitativa", "Resultado Final"]
    for col, h in enumerate(headers1):
        ws1.write(2, col, h, header_fmt)
        
    for row, data in enumerate(df.itertuples(), start=3):
        ws1.write(row, 0, data._1, cell_fmt) # Nome
        ws1.write(row, 1, data.Menção, cell_fmt) # Qualitativa
        ws1.write(row, 2, data._8, num_fmt) # Quantitativa (Classificação Final)
        ws1.write(row, 3, data.Situação, cell_fmt) # Resultado
        
    # --- FOLHA 2: Grelha de Apoio (Baseado no Ficheiro 2) ---
    ws2 = workbook.add_worksheet("Cálculo Detalhado")
    headers2 = ["Nome", "Teórica (50%)", "Ferramentas (60%)", "Equipamentos (20%)", "Estabilização (20%)", "Prática Total (50%)", "Final"]
    
    for col, h in enumerate(headers2):
        ws2.write(0, col, h, header_fmt)
        
    for row, data in enumerate(df.itertuples(), start=1):
        ws2.write(row, 0, data._1, cell_fmt)
        ws2.write(row, 1, data._2, num_fmt)
        ws2.write(row, 2, data._3, num_fmt)
        ws2.write(row, 3, data._4, num_fmt)
        ws2.write(row, 4, data._5, num_fmt)
        ws2.write(row, 5, data._7, num_fmt) # Pratica Final
        ws2.write(row, 6, data._8, num_fmt) # Final

    # --- FOLHA 3: Ficha Prática (Resumo Ficheiro 3) ---
    ws3 = workbook.add_worksheet("Parâmetros Práticos")
    ws3.write(0, 0, "Resumo dos Parâmetros de Avaliação Prática (Checklist)", header_fmt)
    ws3.write(1, 0, "Nota: Esta folha consolida os resultados da observação direta.", cell_fmt)
    
    # Escrever dados
    for row, data in enumerate(df.itertuples(), start=3):
        ws3.write(row, 0, f"Formando: {data._1}", header_fmt)
        ws3.write(row, 1, f"Nota Ferramentas: {data._3}", cell_fmt)
        ws3.write(row, 2, f"Nota Equipamentos: {data._4}", cell_fmt)
        ws3.write(row, 3, f"Nota Estabilização: {data._5}", cell_fmt)

    workbook.close()
    return output.getvalue()

excel_data = to_excel(edited_df)

st.download_button(
    label="📥 Descarregar Excel Consolidado (3 em 1)",
    data=excel_data,
    file_name='Avaliacao_UFCD9889_Consolidada.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)
