import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime

@st.cache_data(show_spinner=False)  # Cache para otimização
def load_data(uploaded_file):
    """Carrega dados com tratamento de erros"""
    try:
        return pd.read_excel(uploaded_file, engine='openpyxl')
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {str(e)}")
        return None

def preencher_documentos(df_original, df_banco):
    """Versão otimizada do preenchimento"""
    # Otimização: Converter para dict uma vez
    mapeamento = df_banco.set_index('Razão Social')['CPF/CNPJ'].to_dict()
    
    # Preenchimento vetorizado (mais rápido que .map())
    df_original['CPF'] = df_original['Nome da Pessoa'].apply(
        lambda x: mapeamento.get(x, '')
    )
    return df_original

def main():
    # Configuração otimizada
    st.set_page_config(
        page_title="⚡ Preenchimento Rápido de Documentos",
        layout="centered"
    )
    
    st.title("🔄 Preenchimento Automático de Documentos (Otimizado)")
    st.caption("Versão acelerada com tratamento de erros melhorado")

    # Uploads otimizados
    with st.expander("📁 Carregar Arquivos", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            st.header("1. Banco de Referência")
            db_file = st.file_uploader(
                "Planilha de referência",
                type=["xlsx"],
                key="db_file"
            )
        with col2:
            st.header("2. Planilha a Preencher")
            input_file = st.file_uploader(
                "Planilha de trabalho",
                type=["xlsx"],
                key="input_file"
            )

    if db_file and input_file:
        start_time = datetime.now()
        
        with st.spinner("⏳ Carregando e verificando arquivos..."):
            df_banco = load_data(db_file)
            df_input = load_data(input_file)
            
            if df_banco is None or df_input is None:
                return

            # Verificação rápida de colunas
            required_cols = {
                'Banco': ['Razão Social', 'CPF/CNPJ'],
                'Input': ['Nome da Pessoa', 'CPF']
            }
            
            for df, cols in zip([df_banco, df_input], required_cols.values()):
                missing = [col for col in cols if col not in df.columns]
                if missing:
                    st.error(f"Colunas faltantes: {', '.join(missing)}")
                    return

        with st.spinner("🔍 Processando dados..."):
            df_processado = preencher_documentos(df_input, df_banco)
            
            # Gerar arquivo em memória
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_processado.to_excel(writer, index=False)
            output.seek(0)

        st.success(f"✅ Concluído em {(datetime.now() - start_time).total_seconds():.2f} segundos!")
        
        # Visualização e download
        st.subheader("Resultado Final")
        st.dataframe(df_processado.head(), use_container_width=True)
        
        st.download_button(
            label="⬇️ Baixar Planilha Processada",
            data=output,
            file_name=f"preenchido_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
