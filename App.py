import pandas as pd
import streamlit as st
from io import BytesIO
import time
import numpy as np  # Reforço de compatibilidade

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

# Leitura segura, sem cache (upload muda o objeto toda vez)
def load_data(file):
    try:
        df = pd.read_excel(file)
        # Normaliza os nomes das colunas para evitar erros
        df.columns = (
            df.columns.str.strip()
                      .str.normalize('NFKD')
                      .str.encode('ascii', errors='ignore')
                      .str.decode('utf-8')
        )
        return df
    except Exception as e:
        st.error(f"Erro na leitura: {str(e)}")
        st.stop()

def main():
    config_app()

    st.title("🚀 Preenchimento Automático Turbo")
    st.caption("Versão 4.0 - Otimizada para Python 3.13+")

    # Uploads
    col1, col2 = st.columns(2)
    with col1:
        st.header("Banco de Referência")
        db_file = st.file_uploader("Carregue aqui", 
                                   type=["xlsx"], 
                                   key="db",
                                   help="Deve conter 'Razão Social' e 'CPF/CNPJ'")
    with col2:
        st.header("Planilha a Preencher")
        input_file = st.file_uploader("Carregue aqui", 
                                      type=["xlsx"], 
                                      key="input",
                                      help="Deve conter 'Nome da Pessoa' e 'CPF'")

    if db_file and input_file:
        start_time = time.perf_counter()

        with st.spinner("🔍 Processando..."):
            try:
                df_banco = load_data(db_file)
                df_input = load_data(input_file)

                # Padroniza colunas para facilitar a validação
                df_banco.columns = df_banco.columns.str.strip()
                df_input.columns = df_input.columns.str.strip()

                # Validação
                required = {
                    'Banco': ['Razao Social', 'CPF/CNPJ'],  # Sem acento
                    'Input': ['Nome da Pessoa', 'CPF']
                }

                for df_name, df, cols in zip(required.keys(), [df_banco, df_input], required.values()):
                    missing = [col for col in cols if col not in df.columns]
                    if missing:
                        st.error(f"🚨 {df_name}: Faltam colunas: {', '.join(missing)}")
                        return

                # Merge ao invés de map (mais robusto e rápido)
                df_final = df_input.merge(
                    df_banco[['Razao Social', 'CPF/CNPJ']],
                    left_on='Nome da Pessoa',
                    right_on='Razao Social',
                    how='left'
                )

                df_final.drop(columns='Razao Social', inplace=True)
                df_final['CPF'] = df_final['CPF/CNPJ'].fillna('')
                df_final.drop(columns='CPF/CNPJ', inplace=True)

                # Exporta Excel
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
