import pandas as pd
import streamlit as st
from io import BytesIO
import time

# Configuração à prova de erros
@st.cache_resource
def config_app():
    st.set_page_config(
        page_title="⚡ Preenchimento Turbo",
        page_icon="📊",
        layout="centered"
    )

# Cache de dados com timeout
@st.cache_data(ttl=3600, show_spinner=False)
def load_data(file):
    try:
        return pd.read_excel(file, engine='openpyxl')
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {str(e)}")
        st.stop()

def main():
    config_app()
    
    st.title("🚀 Preenchimento Automático Turbo")
    st.caption("Versão 3.0 - Otimizada para Python 3.11")
    
    # Uploads paralelos
    col1, col2 = st.columns(2)
    with col1:
        st.header("Banco de Referência")
        db_file = st.file_uploader("Carregue aqui", type=["xlsx"], key="db")
    with col2:
        st.header("Planilha a Preencher")
        input_file = st.file_uploader("Carregue aqui", type=["xlsx"], key="input")

    if db_file and input_file:
        start_time = time.time()
        
        with st.spinner("🔍 Processando..."):
            try:
                # Carregamento acelerado
                df_banco = load_data(db_file)
                df_input = load_data(input_file)
                
                # Verificação relâmpago
                required_cols = {
                    'Banco': ['Razão Social', 'CPF/CNPJ'],
                    'Input': ['Nome da Pessoa', 'CPF']
                }
                
                for df, cols in zip([df_banco, df_input], required_cols.values()):
                    if not all(col in df.columns for col in cols):
                        missing = [col for col in cols if col not in df.columns]
                        st.error(f"🚨 Colunas faltando: {', '.join(missing)}")
                        return
                
                # Processamento turbo
                mapping = df_banco.set_index('Razão Social')['CPF/CNPJ'].to_dict()
                df_input['CPF'] = df_input['Nome da Pessoa'].map(mapping)
                
                # Saída direta
                output = BytesIO()
                df_input.to_excel(output, index=False, engine='openpyxl')
                output.seek(0)
                
                st.success(f"✅ Concluído em {time.time() - start_time:.2f} segundos!")
                
                st.download_button(
                    label="⬇️ Baixar Planilha Processada",
                    data=output,
                    file_name="resultado_turbo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"❌ Erro crítico: {str(e)}")
                st.stop()

if __name__ == "__main__":
    main()
