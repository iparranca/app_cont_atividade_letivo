import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

st.set_page_config(page_title="Contagem Dinâmica", layout="wide")
st.title("📊 Contagem Dinâmica de Atividades")

# 1. Escolha do separador
sep = st.selectbox(
    "Seleciona o separador do teu CSV:",
    options=[(",", "Vírgula (,)"), (";", "Ponto e Vírgula (;)")],
    format_func=lambda x: x[1]
)[0]

uploaded_file = st.file_uploader(
    "Carregar ficheiro CSV (primeira coluna = DataHora)",
    type=["csv"]
)

def determinar_ano_letivo(data):
    if data.month >= 9:
        return f"{data.year}/{data.year + 1}"
    else:
        return f"{data.year - 1}/{data.year}"

if uploaded_file:
    # 2. Leitura do CSV com o separador escolhido
    try:
        df = pd.read_csv(
            uploaded_file,
            sep=sep,
            encoding='latin1'
        )
    except Exception as e:
        st.error(f"Erro ao ler o CSV: {e}")
        st.stop()

    # 3. Validar primeira coluna como datetime
    cols = list(df.columns)
    datahora_col = cols[0]
    try:
        df[datahora_col] = pd.to_datetime(df[datahora_col], errors='raise')
    except Exception:
        st.error(f"A primeira coluna «{datahora_col}» não contém datas válidas.")
        st.stop()

    # 4. Mapear colunas dinamicamente
    restantes = cols[1:]
    if len(restantes) < 3:
        st.error("O ficheiro deve ter pelo menos 4 colunas (DataHora + 3 colunas para Atividade, Turma, Disciplina).")
        st.stop()

    st.subheader("🔧 Mapeamento de colunas")
    atividade_col = st.selectbox("Coluna de Atividade", restantes)
    turma_col     = st.selectbox("Coluna de Turma", [c for c in restantes if c != atividade_col])
    disciplina_col= st.selectbox(
        "Coluna de Disciplina",
        [c for c in restantes if c not in (atividade_col, turma_col)]
    )

    # 5. Criar coluna AnoLetivo
    df['AnoLetivo'] = df[datahora_col].apply(determinar_ano_letivo)

    st.write("### 👁️ Dados carregados e mapeados")
    st.dataframe(df[[datahora_col, atividade_col, turma_col, disciplina_col, 'AnoLetivo']])

    # 6. Seleção do tipo de contagem
    tipos = {
        "Por Atividade":            [atividade_col],
        "Por Turma":                [turma_col],
        "Por Disciplina":           [disciplina_col],
        "Por Ano Letivo":           ['AnoLetivo'],
        "Por Atividade e Turma":    [atividade_col, turma_col],
        "Por Atividade e Ano Letivo":[atividade_col, 'AnoLetivo'],
        "Por Disciplina e Turma":   [disciplina_col, turma_col]
    }
    tipo_contagem = st.selectbox("Selecionar tipo de contagem", list(tipos.keys()))

    # 7. Agrupar e contar
    group_cols = tipos[tipo_contagem]
    tabela = df.groupby(group_cols).size().reset_index(name="Contagem")

    st.subheader("📋 Resultado da Contagem")
    st.dataframe(tabela)

    # 8. Exportar para Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DadosTratados")
        tabela.to_excel(writer, index=False, sheet_name="Resumo")
    output.seek(0)

    st.download_button(
        label="📥 Descarregar Excel",
        data=output.read(),
        file_name="contagem_dinamica.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👆 Seleciona o separador e carrega um ficheiro CSV para começar.")
