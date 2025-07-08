
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

st.set_page_config(page_title="Contagem", layout="wide")
st.title("Exportar Contagens")

# Estilo customizado
st.markdown("""
    <style>
    .custom-label {
        font-size: 20px;
        font-weight: bold;
        color: #2A76D2;
        margin-top: 20px;
    }
    .stDownloadButton > button {
        background-color: #2A76D2;
        color: white;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# 1. Escolha do separador
sep = st.selectbox(
    "Seleciona o separador do teu CSV:",
    options=[(";", "Ponto e Vírgula (;)"), (",", "Vírgula (,)")],
    format_func=lambda x: x[1]
)[0]

uploaded_file = st.file_uploader(
    "Carregar ficheiro CSV (primeira coluna = Data ou DataHora)",
    type=["csv"]
)

def determinar_ano_letivo(data):
    if data.month >= 9:
        return f"{data.year}/{data.year + 1}"
    else:
        return f"{data.year - 1}/{data.year}"

if uploaded_file:
    try:
        df = pd.read_csv(uploaded_file, sep=sep, encoding='latin1')
        df.columns = df.columns.str.strip()
    except Exception as e:
        st.error(f"Erro ao ler o CSV: {e}")
        st.stop()

    if len(df.columns) == 1:
        st.error("O ficheiro CSV parece não estar separado corretamente.")
        st.stop()

    if pd.Series(df.columns).isnull().any() or any(c.strip() == "" for c in df.columns):
        st.error("Todos os cabeçalhos devem estar preenchidos.")
        st.stop()

    primeira_coluna = df.columns[0]
    try:
        df[primeira_coluna] = pd.to_datetime(df[primeira_coluna], errors='raise')
    except Exception:
        st.error(f"A primeira coluna «{primeira_coluna}» não contém datas válidas.")
        st.stop()

    colunas_vazias = df.columns[df.isnull().all()]
    colunas_com_nulos = df.columns[df.isnull().any()]

    if len(colunas_vazias) > 0:
        st.error("Colunas totalmente vazias:")
        st.error(colunas_vazias.tolist())

    if len(colunas_com_nulos) > 0:
        st.warning("Colunas com pelo menos um valor nulo:")
        st.warning(colunas_com_nulos.tolist())

    df['AnoLetivo'] = df[primeira_coluna].apply(determinar_ano_letivo)

    # NOVO: Escolha do ano letivo com label destacada
    st.markdown('<div class="custom-label">Seleciona o(s) Ano(s) Letivo(s) a incluir:</div>', unsafe_allow_html=True)
    anos_disponiveis = sorted(df['AnoLetivo'].unique())
    anos_escolhidos = st.multiselect(
        label="",
        options=anos_disponiveis,
        default=anos_disponiveis
    )

    if not anos_escolhidos:
        st.warning("Seleciona pelo menos um ano letivo.")
        st.stop()

    # Filtrar pelo(s) ano(s) letivo(s)
    df = df[df['AnoLetivo'].isin(anos_escolhidos)]

    restantes_colunas = df.columns[1:-1]  # Exclui data e AnoLetivo

    st.write("### Pré-visualização dos dados")
    st.dataframe(df)

    # NOVO: Label destacada e comportamento de fechar ao selecionar
    st.markdown('<div class="custom-label">Seleciona as colunas para a contagem:</div>', unsafe_allow_html=True)
    colunas_selecionadas = st.multiselect(
        label="",
        options=list(restantes_colunas) + ['AnoLetivo'],
        default=list(restantes_colunas),
        key="multiselect_colunas"
    )

    if not colunas_selecionadas:
        st.warning("Seleciona pelo menos uma coluna para contagem.")
        st.stop()

    tabela = df.groupby(colunas_selecionadas).size().reset_index(name="Contagem")
    descricao = "Por " + " + ".join(colunas_selecionadas)

    st.subheader(f"Resultado da Contagem ({descricao})")
    st.dataframe(tabela)

    # NOVO: Input para nome do ficheiro
    nome_ficheiro = st.text_input("Nome do ficheiro Excel a exportar (sem extensão):", value="contagem_inteligente")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DadosTratados")
        tabela.to_excel(writer, index=False, sheet_name="Resumo")
    output.seek(0)

    st.download_button(
        label="Descarregar Excel",
        data=output.read(),
        file_name=nome_ficheiro + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Seleciona o separador e carrega um ficheiro CSV para começar.")
