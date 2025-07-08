import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

st.set_page_config(page_title="Contagem Inteligente", layout="wide")
st.title("Exportar Contagens")

# Estilo customizado
st.markdown("""
    <style>
    .custom-label {
        font-size: 20px;
        font-weight: bold;
        color: white;
        margin-top: 20px;
        padding: 10px;
        background-color: #1E88E5;
        border-radius: 5px;
    }
    .stDownloadButton > button {
        background-color: #2A76D2;
        color: white;
        font-weight: bold;
    }
    .preview-section, .count-section {
        background-color: #BBDEFB;
        padding: 20px;
        border-radius: 8px;
        margin-top: 20px;
    }
    </style>
""", unsafe_allow_html=True)

st.info("""**Notas importantes**
1 - Ficheiro a carregar:
    a) Só pode carregar um ficheiro CSV (guarda o ficheiro excel como CSV)
    b) Primeira coluna (informação que está antes do primeiro separador, isto é, antes do primeiro ";" ou "," ou tabulação)
       tem que conter a informação da data ou Datahora. Esta informação irá referenciar o Ano letivo)
    c) Colocar na primeira linha uma linha nova com os cabeçalhos de cada coluna. Por exemplo:

    Ano Letivo;Aluno;Atividade;Ciclo;Ano Turma;Turma;Disciplina
    25/09/2023;Aluno;Ler;2º Ciclo;6º Ano;H;Português
    25/09/2023;Aluno;Pesquisar na Internet;3º Ciclo;7º Ano;C;Atividade da Biblioteca
    25/09/2023;Aluno;Trabalhar em grupo;2º Ciclo;6º Ano;A;Tempo Livre""")

st.info("""**Selecione**:  
1 - **primeiro** o tipo de separador que tem dentro do ficheiro que vai carregar (podes escolher ";" ou "," ou tabulação).  
2 - **Arraste ou clique** no botão para carregar o ficheiro CSV .""")

# 1. Escolha do separador
        #sep = st.selectbox(
        #    "**Selecione o separador do teu CSV:**",
        #    options=[(";", "Ponto e Vírgula (;)"), (",", "Vírgula (,)"), ("\t", "Tabulação")],
        #    format_func=lambda x: x[1]
        #)[0]

        # Texto com HTML para tamanho maior e negrito
        st.markdown("<p style='font-size:20px; font-weight:bold;'>Selecione o separador do teu CSV:</p>", unsafe_allow_html=True)
        
        # Selectbox sem label (label já foi tratada acima)
        sep = st.selectbox(
            "",
            options=[(";", "Ponto e Vírgula (;)"), (",", "Vírgula (,)"), ("\t", "Tabulação")],
            format_func=lambda x: x[1]
        )[0]

st.markdown("<p style='font-size:20px; font-weight:bold;'>Carregar ficheiro CSV:</p>", unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "",
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
    df['Mês'] = df[primeira_coluna].dt.month_name()

    st.markdown('<div class="custom-label">Selecione o(s) Ano(s) Letivo(s) que queres incluir:</div>', unsafe_allow_html=True)
    anos_disponiveis = sorted(df['AnoLetivo'].unique())
    anos_escolhidos = st.multiselect("", options=anos_disponiveis, default=anos_disponiveis)

    if not anos_escolhidos:
        st.warning("Seleciona pelo menos um ano letivo.")
        st.stop()

    df = df[df['AnoLetivo'].isin(anos_escolhidos)]

    # Filtro por mês
    st.markdown('<div class="custom-label">Seleciona o(s) Mês(es):</div>', unsafe_allow_html=True)
    meses_disponiveis = df['Mês'].unique()
    meses_escolhidos = st.multiselect("", options=meses_disponiveis, default=meses_disponiveis)
    df = df[df['Mês'].isin(meses_escolhidos)]

    restantes_colunas = df.columns[1:-2]  # Exclui data, AnoLetivo e Mês

    st.markdown('<div class="preview-section">', unsafe_allow_html=True)
    st.write("### Pré-visualização dos dados")
    df_preview = df.copy()
    numeric_cols = df_preview.select_dtypes(include='number').columns
    total_row = pd.Series(["" for _ in df_preview.columns], index=df_preview.columns)
    total_row[numeric_cols] = df_preview[numeric_cols].sum(numeric_only=True)
    df_preview.loc['Total'] = total_row
    st.dataframe(df_preview)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="custom-label">Seleciona as colunas para fazer a contagem:</div>', unsafe_allow_html=True)
    colunas_selecionadas = st.multiselect(
        "",
        options=list(restantes_colunas) + ['AnoLetivo', 'Mês'],
        default=list(restantes_colunas),
        key="multiselect_colunas"
    )

    if not colunas_selecionadas:
        st.warning("Seleciona pelo menos uma coluna para contagem.")
        st.stop()

    tabela = df.groupby(colunas_selecionadas).size().reset_index(name="Contagem")

    descricao = "Por " + " + ".join(colunas_selecionadas)

    st.markdown('<div class="count-section">', unsafe_allow_html=True)
    st.subheader(f"Resultado da Contagem ({descricao})")
    st.dataframe(tabela)
    st.markdown(f'<div style="text-align:right; font-weight:bold;">Total geral: {tabela["Contagem"].sum()}</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    nome_ficheiro = st.text_input("Nome do ficheiro Excel a exportar (sem extensão):", value=f"contagem_{anos_escolhidos}")

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
#else:
#    st.info("Seleciona o separador e carrega um ficheiro CSV para começar.")
