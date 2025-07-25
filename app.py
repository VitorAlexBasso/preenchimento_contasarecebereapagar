import pandas as pd
import streamlit as st
from io import BytesIO
import time
import numpy as np  # Compatibilidade

# Configuração inicial
def config_app():
    st.set_page_config(
        page_title="⚡ Preenchimento Turbo 3.13",
        page_icon="📊",
        layout="centered",
        menu_items={
            'Get Help': 'https://github.com/seu-usuario/preenchimento-contas',
            'About': "Versão otimizada para Python 3.13 | pandas 2.2.1 | numpy 2.0.0"
        }
    )

# Leitura segura com fallback
def load_data(file):
    try:
        return pd.read_excel(file, engine='openpyxl')
    except ImportError:
        st.error("⚠️ Falta a dependência 'openpyxl'. Inclua no requirements.txt: openpyxl>=3.1.2")
        st.stop()
    except Exception as e:
        st.error(f"Erro na leitura: {str(e)}")
        st.stop()

# Função para buscar coluna com variações
def encontrar_coluna(colunas, nomes_possiveis):
    colunas_normalizadas = {c.strip().lower(): c for c in colunas}
    for nome in nomes_possiveis:
        nome_normalizado = nome.strip().lower()
        if nome_normalizado in colunas_normalizadas:
            return colunas_normalizadas[nome_normalizado]
    return None

def main():
    config_app()

    st.title("🚀 Preenchimento Automático Turbo")
    st.caption("Versão 4.0 - Otimizada para Python 3.13+")

    col1, col2 = st.columns(2)
    with col1:
        st.header("Banco de Referência")
        db_file = st.file_uploader("Carregue aqui", type=["xlsx"], key="db",
                                   help="Deve conter 'Razão Social' e 'CPF/CNPJ'")
    with col2:
        st.header("Planilha a Preencher")
        input_file = st.file_uploader("Carregue aqui", type=["xlsx"], key="input",
                                      help="Deve conter 'Nome da Pessoa' e 'CPF'")

    if db_file and input_file:
        start_time = time.perf_counter()

        with st.spinner("🔍 Processando..."):
            try:
                df_banco = load_data(db_file)
                df_input = load_data(input_file)

                df_banco.columns = df_banco.columns.str.strip()
                df_input.columns = df_input.columns.str.strip()

                # Encontra as colunas corretamente, mesmo com variações
                col_razao_social = encontrar_coluna(df_banco.columns, ["Razao Social", "Razão Social"])
                col_cpf_cnpj = encontrar_coluna(df_banco.columns, ["CPF/CNPJ", "Cpf/Cnpj", "Documento"])
                col_nome_pessoa = encontrar_coluna(df_input.columns, ["Nome da Pessoa"])
                col_cpf = encontrar_coluna(df_input.columns, ["CPF"])

                # Validação
                missing = []
                if not col_razao_social: missing.append("Razao Social")
                if not col_cpf_cnpj: missing.append("CPF/CNPJ")
                if not col_nome_pessoa: missing.append("Nome da Pessoa")
                if not col_cpf: missing.append("CPF")

                if missing:
                    st.error(f"🚨 Faltam colunas: {', '.join(missing)}")
                    return

                df_final = df_input.merge(
                    df_banco[[col_razao_social, col_cpf_cnpj]],
                    left_on=col_nome_pessoa,
                    right_on=col_razao_social,
                    how='left'
                )

                df_final.drop(columns=col_razao_social, inplace=True)
                df_final[col_cpf] = df_final[col_cpf_cnpj].fillna('')
                df_final.drop(columns=col_cpf_cnpj, inplace=True)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False)
                output.seek(0)

                elapsed = time.perf_counter() - start_time
                st.success(f"✅ Concluído em {elapsed:.2f} segundos!")

                st.download_button(
                    label="⬇️ Baixar Planilha Processada",
                    data=output,
                    file_name=f"resultado_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

                with st.expander("🔍 Visualizar Resultado"):
                    st.dataframe(df_final.head(), use_container_width=True)

            except Exception as e:
                st.error(f"❌ Falha: {str(e)}")
                st.stop()

if __name__ == "__main__":
    main()
