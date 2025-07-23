import pandas as pd
import streamlit as st
from io import BytesIO
import time
import numpy as np  # Adicionado para garantir compatibilidade

# Configuração otimizada para Python 3.13
@st.cache_resource
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

# Cache com tratamento para numpy 2.0
@st.cache_data(ttl=3600, show_spinner=False)
def load_data(file):
    try:
        # Engine padrão para Excel (openpyxl já incluso no requirements)
        return pd.read_excel(file)
    except Exception as e:
        st.error(f"Erro na leitura: {str(e)}")
        st.stop()

def main():
    config_app()
    
    st.title("🚀 Preenchimento Automático Turbo")
    st.caption("Versão 4.0 - Otimizada para Python 3.13+")
    
    # Uploads com validação
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
        start_time = time.perf_counter()  # Mais preciso para Python 3.13
        
        with st.spinner("🔍 Processando..."):
            try:
                # Carregamento seguro
                df_banco = load_data(db_file)
                df_input = load_data(input_file)
                
                # Verificação robusta de colunas
                required = {
                    'Banco': ['Razão Social', 'CPF/CNPJ'],
                    'Input': ['Nome da Pessoa', 'CPF']
                }
                
                for df, cols in zip([df_banco, df_input], required.values()):
                    missing = [col for col in cols if col not in df.columns]
                    if missing:
                        st.error(f"🚨 Faltam colunas: {', '.join(missing)}")
                        return
                
                # Processamento otimizado
                mapping = df_banco.set_index('Razão Social')['CPF/CNPJ'].astype(str).to_dict()
                df_input['CPF'] = df_input['Nome da Pessoa'].map(mapping).fillna('')
                
                # Saída eficiente
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_input.to_excel(writer, index=False)
                output.seek(0)
                
                # Feedback de performance
                elapsed = time.perf_counter() - start_time
                st.success(f"✅ Concluído em {elapsed:.2f} segundos!")
                
                # Download
                st.download_button(
                    label="⬇️ Baixar Planilha Processada",
                    data=output,
                    file_name=f"resultado_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                # Visualização rápida
                with st.expander("🔍 Visualizar Resultado"):
                    st.dataframe(df_input.head(), use_container_width=True)
                    
            except Exception as e:
                st.error(f"❌ Falha: {str(e)}")
                st.stop()

if __name__ == "__main__":
    main()
