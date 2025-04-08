import streamlit as st
import pandas as pd

# ============================ CONFIGURAÇÕES GERAIS ============================
# Defina o caminho para o arquivo Excel (ajuste conforme necessário)
EXCEL_PATH = r"C:\Users\andre.suhai\Desktop\PYTHON\XP.xlsx"

def carregar_dados():
    """Carrega os dados da planilha Excel."""
    try:
        return pd.read_excel(EXCEL_PATH)
    except Exception as e:
        st.warning("Arquivo Excel não encontrado ou erro ao ler. Um novo será criado se necessário.")
        return pd.DataFrame()

def salvar_dados(df):
    """Salva os dados no arquivo Excel."""
    try:
        df.to_excel(EXCEL_PATH, index=False)
        st.success("Dados atualizados com sucesso!")
    except Exception as e:
        st.error(f"Erro ao salvar o arquivo Excel: {e}")

# ============================ CONTEÚDO DO APP ============================
st.title("Aplicativo de Teste")

# Carrega e exibe os dados da planilha Excel
df = carregar_dados()
st.subheader("Visualização dos Dados")
st.dataframe(df)

st.markdown("---")
st.subheader("Editar Registro")

if df.empty:
    st.warning("Nenhum dado disponível para editar.")
else:
    # Formata as opções de seleção de linha usando a coluna "Nome" se existir
    def format_row(i):
        if "Nome" in df.columns:
            return f"Linha {i} - {df.loc[i, 'Nome']}"
        else:
            return f"Linha {i}"

    linha = st.selectbox(
        "Selecione a linha para editar",
        df.index,
        format_func=format_row
    )
    registro = df.loc[linha].to_dict()

    # Formulário para editar os valores do registro
    with st.form("editar_registro"):
        novos_valores = {}
        for coluna, valor in registro.items():
            novos_valores[coluna] = st.text_input(f"{coluna}", value=str(valor))
        editar = st.form_submit_button("Salvar Alterações")

    if editar:
        for coluna, novo in novos_valores.items():
            try:
                valor_atual = registro[coluna]
                if isinstance(valor_atual, int):
                    df.at[linha, coluna] = int(novo)
                elif isinstance(valor_atual, float):
                    df.at[linha, coluna] = float(novo)
                else:
                    df.at[linha, coluna] = novo
            except Exception:
                df.at[linha, coluna] = novo
        salvar_dados(df)
        st.experimental_rerun()  # Recarrega o app para refletir as alterações
