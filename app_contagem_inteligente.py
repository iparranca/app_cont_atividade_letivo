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

# Notas
#st.info("""**Notas importantes**

#Ficheiro a carregar:  
#   a) Só pode carregar um ficheiro CSV (guarde o ficheiro Excel como CSV);  
#   b) A primeira coluna (informação que está antes do primeiro separador, isto é, antes do primeiro ";" ou "," ou tabulação) deve conter a informação da data ou Datahora. Esta informação irá referenciar o Ano Letivo;  
#   c) Colocar na primeira linha uma linha nova com os cabeçalhos de cada coluna.  

#   **Exemplo:**  
#   Ano Letivo;Aluno;Atividade;Ciclo;Ano Turma;Turma;Disciplina  
#  25/09/2023;Aluno;Ler;2º Ciclo;6º Ano;H;Português  
#   25/09/2023;Aluno;Pesquisar na Internet;3º Ciclo;7º Ano;C;Atividade da Biblioteca  
#   25/09/2023;Aluno;Trabalhar em grupo;2º Ciclo;6º Ano;A;Tempo Livre
#""")

st.markdown("""
<div style='border-left: 6px solid #2196F3; background-color: #f0f8ff; padding: 16px; border-radius: 6px;'>
  <p style='font-size:16px;'><strong>Notas importantes</strong></p>
  <p>Ficheiro a carregar:</p>
  <ul>
    <li>a) Só pode carregar um ficheiro <span style='color:#007BFF; font-weight:bold;'>CSV</span> (guarde o ficheiro Excel como CSV);</li>
    <li>b) A primeira coluna (informação que está antes do primeiro separador, isto é, antes do primeiro <span style='color:#007BFF;'>";"</span> ou <span style='color:#007BFF;'>","</span> ou <span style='color:#007BFF;'>tabulação</span>) deve conter a informação da <span style='color:#007BFF;'>data</span> ou <span style='color:#007BFF;'>Datahora</span>. Esta informação irá referenciar o <span style='color:#007BFF;'>Ano Letivo</span>;</li>
    <li>c) Colocar na primeira linha uma linha nova com os <span style='color:#007BFF;'>cabeçalhos</span> de cada coluna.</li>
  </ul>

  <p><strong>Exemplo:</strong></p>
  <pre style='background-color:#e8f4fd; padding:4px; border-radius:4px;font-weight:bold;'>Ano Letivo;Aluno;Atividade;Ciclo;Ano Turma;Turma;Disciplina</pre>
  <pre style='background-color:#e8f4fd; padding:4px; border-radius:4px;'>25/09/2023;Aluno;Ler;2º Ciclo;6º Ano;H;Português</pre>
  <pre style='background-color:#e8f4fd; padding:4px; border-radius:4px;'>25/09/2023;Aluno;Pesquisar na Internet;3º Ciclo;7º Ano;C;Atividade da Biblioteca</pre>
  <pre style='background-color:#e8f4fd; padding:4px; border-radius:4px;margin-bottom: 200px'>25/09/2023;Aluno;Trabalhar em grupo;2º Ciclo;6º Ano;A;Tempo Livre</pre>
</div>
""", unsafe_allow_html=True)

st.info("""**Selecione**:  
1 - **primeiro** o tipo de separador que tem dentro do ficheiro (podes escolher ";" ou "," ou tabulação).  
2 - **Arraste ou clique** no botão para carregar o ficheiro CSV .""")

# Separador
st.markdown("<p style='font-size:20px; font-weight:bold;margin-top: 280px;margin-bottom: 0,2px;color:#0056b3;'>Selecione22 o separador do teu CSV:</p>", unsafe_allow_html=True)
sep = st.selectbox(
    "",
    options=[(";", "Ponto e Vírgula (;)"), (",", "Vírgula (,)"), ("\t", "Tabulação")],
    format_func=lambda x: x[1],
    key="select_sep"
)[0]  # <--- AQUI! Isso extrai apenas o separador (string)

# Upload do ficheiro
st.markdown("<p style='font-size:20px; font-weight:bold;margin-top: 60px;margin-bottom: 0,2px;color:#0056b3;'>Carregar ficheiro CSV:</p>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("", type=["csv"])

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
    df['Dia'] = df[primeira_coluna].dt.date
    df['Trimestre'] = df[primeira_coluna].dt.quarter
    df['Semestre'] = df[primeira_coluna].dt.month.map(lambda m: 1 if m <= 6 else 2)

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

    restantes_colunas = df.columns[1:-6]


    
    #Isabel - Inicio
    # Pré-visualização
    #st.markdown('<div class="preview-section">', unsafe_allow_html=True)
    #st.write("### Pré-visualização dos dados")
    #df_preview = df.copy()
    #numeric_cols = df_preview.select_dtypes(include='number').columns
    #total_row = pd.Series(["" for _ in df_preview.columns], index=df_preview.columns)
    #total_row[numeric_cols] = df_preview[numeric_cols].sum(numeric_only=True)
    #df_preview.loc['Total'] = total_row
    #st.dataframe(df_preview)
    # Pré-visualização dos dados
    st.markdown('<div class="preview-section">', unsafe_allow_html=True)
    st.write("### Pré-visualização dos dados")

    df_preview = df.copy()

# Renomeia colunas
    col_data = df.columns[0]  # Primeira coluna original (de data)
    df_preview = df_preview.rename(columns={
        col_data: "Data",
        "AnoLetivo": "Ano Letivo"
    })

# Reordena colunas: Ano Letivo primeiro
    cols = df_preview.columns.tolist()
    cols.insert(0, cols.pop(cols.index("Ano Letivo")))  # Move "Ano Letivo" para o início

# Oculta colunas técnicas
    colunas_a_ocultar = {"Dia", "Mês", "Trimestre", "Semestre"}
    cols = [c for c in cols if c not in colunas_a_ocultar]
    df_preview = df_preview[cols]

# Adiciona linha de totais (apenas para colunas numéricas)
    numeric_cols = df_preview.select_dtypes(include='number').columns
    total_row = pd.Series(["" for _ in df_preview.columns], index=df_preview.columns)
    total_row[numeric_cols] = df_preview[numeric_cols].sum(numeric_only=True)
    df_preview.loc['Total'] = total_row

# Exibe
    st.dataframe(df_preview)
    st.markdown('</div>', unsafe_allow_html=True)
#Isabel - Fim

    st.markdown(f"<div style='text-align:right; font-weight:bold;'>Total de registos: {len(df)}</div>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Escolha das colunas e agregação
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

    st.markdown('<div class="custom-label">Seleciona como queres calcular a média:</div>', unsafe_allow_html=True)
    agregacao = st.selectbox(
        "",
        options=["Nenhuma", "Por Dia", "Por Mês", "Por Trimestre", "Por Semestre"]
    )

    base_tabela = df.copy()

    if agregacao == "Por Dia":
        total_periodos = base_tabela['Dia'].nunique()
    elif agregacao == "Por Mês":
        total_periodos = base_tabela['Mês'].nunique()
    elif agregacao == "Por Trimestre":
        total_periodos = base_tabela['Trimestre'].nunique()
    elif agregacao == "Por Semestre":
        total_periodos = base_tabela['Semestre'].nunique()
    else:
        total_periodos = None

    tabela = base_tabela.groupby(colunas_selecionadas).size().reset_index(name="Contagem")

    if total_periodos and total_periodos > 0:
        tabela["Média"] = tabela["Contagem"] / total_periodos
        tabela["Média"] = tabela["Média"].round(2)

    descricao = "Por " + " + ".join(colunas_selecionadas)
    st.markdown('<div class="count-section">', unsafe_allow_html=True)
    st.subheader(f"Resultado da Contagem ({descricao})" + (f" — Média {agregacao.lower()}" if agregacao != "Nenhuma" else ""))
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
