import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

st.set_page_config(page_title="Contagem Inteligente", layout="wide")
st.title("Exportar Contagens")
# 1. Escolha do separador
sep = st.selectbox(
    "Seleciona o separador do teu CSV:",
    options=[(";", "Ponto e Vírgula (;)"), (",", "Vírgula (,)") ],
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
        # Tentar ler com o separador escolhido
        df = pd.read_csv(uploaded_file, sep=sep, encoding='latin1')
        df.columns = df.columns.str.strip()
        
        '''
        Esse CSV vira um DataFrame assim:

        nome	idade	cidade
        Ana	    30	    São Paulo
        João	25	    Recife
        Maria	40	    Salvador
        
        df.columns retorna:
        Index(['nome', 'idade', 'cidade'], dtype='object')
        
        Ou seja, os nomes das colunas são:
 
        'nome', 'idade', 'cidade'
        
        len(df.columns) retorna:
        3
        Porque há 3 colunas.
       
        '''
    
    except Exception as e:
        st.error(f"Erro ao ler o CSV: {e}")
        st.stop()

    # Verificar se a primeira linha tem apenas uma coluna (se não separaram com , ou ; assume um texto sendo uma coluna)
    if len(df.columns) == 1:
        st.error("O ficheiro CSV parece não estar separado corretamente. Verifica se escolheste o separador correto (vírgula ou ponto e vírgula).")
        st.stop()

    # Verificar se há cabeçalhos vazios
    if df.columns.isnull().any() or any(c.strip() == "" or c.strip().isnull().any() for c in df.columns):
        st.error("Todos os cabeçalhos devem estar preenchidos.")
        st.stop()

    primeira_coluna = df.columns[0]
    try:
        df[primeira_coluna] = pd.to_datetime(df[primeira_coluna], errors='raise')
    except Exception:
        st.error(f"A primeira coluna «{primeira_coluna}» não contém datas válidas. Verifique se o ficheiro tem cabeçalhos.")
        st.stop()

    df['AnoLetivo'] = df[primeira_coluna].apply(determinar_ano_letivo)

    restantes_colunas = df.columns[1:-1]  # Exclui data e AnoLetivo

    st.write("### Pré-visualização dos dados")
    st.dataframe(df)

    colunas_selecionadas = st.multiselect(
        "Seleciona as colunas para fazer a contagem:",
        list(restantes_colunas) + ['AnoLetivo'],
        default=list(restantes_colunas)
    )

    if not colunas_selecionadas:
        st.warning("Seleciona pelo menos uma coluna para contagem.")
        st.stop()

    tabela = df.groupby(colunas_selecionadas).size().reset_index(name="Contagem")
    descricao = "Por " + " + ".join(colunas_selecionadas)

    st.subheader(f"Resultado da Contagem ({descricao})")
    st.dataframe(tabela)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DadosTratados")
        tabela.to_excel(writer, index=False, sheet_name="Resumo")
    output.seek(0)

    st.download_button(
        label="📥 Descarregar Excel",
        data=output.read(),
        file_name="contagem_inteligente.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👆 Seleciona o separador e carrega um ficheiro CSV para começar.")
