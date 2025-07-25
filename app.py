import pandas as pd
import streamlit as st
from io import BytesIO
import time
import numpy as np
import unicodedata
import re
from openpyxl.styles import numbers

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

def load_data(file):
    try:
        # Usar dtype=str para manter a formatação original como texto
        return pd.read_excel(file, engine='openpyxl', dtype=str)
    except ImportError:
        st.error("⚠️ Falta a dependência 'openpyxl'. Inclua no requirements.txt: openpyxl>=3.1.2")
        st.stop()
    except Exception as e:
        st.error(f"Erro na leitura: {str(e)}")
        st.stop()

def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto)
    texto = unicodedata.normalize('NFKD', texto)
    texto = texto.encode('ASCII', 'ignore').decode('utf-8')
    texto = re.sub(r'[^a-zA-Z0-9]', '', texto)
    return texto.lower()

def encontrar_coluna(colunas, nomes_possiveis):
    colunas_dict = {normalizar_texto(c): c for c in colunas}
    for nome in nomes_possiveis:
        nome_normalizado = normalizar_texto(nome)
        if nome_normalizado in colunas_dict:
            return colunas_dict[nome_normalizado]
    return None

def formatar_planilha(worksheet):
    # Aplicar formatação de texto para todas as células
    for row in worksheet.iter_rows():
        for cell in row:
            cell.number_format = numbers.FORMAT_TEXT

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
                # Carregar dados mantendo o formato de texto
                df_banco = load_data(db_file)
                df_input = load_data(input_file)

                df_banco.columns = df_banco.columns.str.strip()
                df_input.columns = df_input.columns.str.strip()

                col_razao_social = encontrar_coluna(df_banco.columns, ["Razao Social", "Razão Social"])
                col_cpf_cnpj = encontrar_coluna(df_banco.columns, ["CPF/CNPJ", "Cpf/Cnpj", "Documento"])
                col_nome_pessoa = encontrar_coluna(df_input.columns, ["Nome da Pessoa"])
                col_cpf = encontrar_coluna(df_input.columns, ["CPF"])

                if not col_razao_social or not col_cpf_cnpj:
                    st.error(f"🚨 Banco: Faltam colunas: {', '.join([c for c, v in zip(['Razao Social', 'CPF/CNPJ'], [col_razao_social, col_cpf_cnpj]) if not v])}")
                    return

                if not col_nome_pessoa or not col_cpf:
                    st.error(f"🚨 Input: Faltam colunas: {', '.join([c for c, v in zip(['Nome da Pessoa', 'CPF'], [col_nome_pessoa, col_cpf]) if not v])}")
                    return

                # Garantir que as colunas de junção sejam strings
                df_banco[col_razao_social] = df_banco[col_razao_social].astype(str)
                df_banco[col_cpf_cnpj] = df_banco[col_cpf_cnpj].astype(str)
                df_input[col_nome_pessoa] = df_input[col_nome_pessoa].astype(str)
                df_input[col_cpf] = df_input[col_cpf].astype(str)

                df_final = df_input.merge(
                    df_banco[[col_razao_social, col_cpf_cnpj]],
                    left_on=col_nome_pessoa,
                    right_on=col_razao_social,
                    how='left'
                )

                df_final.drop(columns=col_razao_social, inplace=True)
                # Preencher valores vazios com string vazia
                df_final[col_cpf] = df_final[col_cpf_cnpj].fillna('').astype(str)
                df_final.drop(columns=col_cpf_cnpj, inplace=True)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False)
                    # Acessar a worksheet para aplicar formatação
                    worksheet = writer.sheets['Sheet1']
                    formatar_planilha(worksheet)
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
                    # Exibir mantendo os zeros à esquerda
                    st.dataframe(df_final.head().style.set_properties(**{
                        'text-align': 'left',
                        'white-space': 'pre'
                    }), use_container_width=True)

            except Exception as e:
                st.error(f"❌ Falha: {str(e)}")
                st.stop()

if __name__ == "__main__":
    main()
