import pandas as pd
import streamlit as st
from io import BytesIO

def preencher_documentos(df_original, df_banco):
    """
    Preenche CPF/CNPJ mantendo a estrutura original
    :param df_original: DataFrame da planilha a ser preenchida
    :param df_banco: DataFrame com o banco de dados de referência
    :return: DataFrame processado
    """
    # Criar dicionário de mapeamento (razão social -> documento)
    mapeamento = dict(zip(df_banco['Razão Social'], df_banco['CPF/CNPJ']))
    
    # Preencher os documentos mantendo o formato original
    df_original['CPF'] = df_original['Nome da Pessoa'].map(mapeamento).fillna('nan')
    
    return df_original

def main():
    st.title("🔄 Preenchimento Automático de Documentos")
    st.caption("Preenche CPF/CNPJ em planilhas mantendo o formato original")
    
    # Upload do banco de dados
    st.header("1. Banco de Dados de Referência")
    st.info("Deve conter colunas 'Razão Social' e 'CPF/CNPJ'")
    db_file = st.file_uploader("Carregue a planilha de referência", type=["xlsx", "xls"])
    
    # Upload da planilha a processar
    st.header("2. Planilha a Ser Preenchida")
    st.info("Deve conter coluna 'Nome da Pessoa' e 'CPF'")
    input_file = st.file_uploader("Carregue a planilha a ser preenchida", type=["xlsx", "xls"])
    
    if db_file and input_file:
        try:
            with st.spinner("Processando..."):
                # Carregar os dados
                df_banco = pd.read_excel(db_file)
                df_input = pd.read_excel(input_file)
                
                # Verificar colunas necessárias
                colunas_necessarias = ['Nome da Pessoa', 'CPF']
                if not all(col in df_input.columns for col in colunas_necessarias):
                    st.error(f"A planilha a ser preenchida deve conter as colunas: {', '.join(colunas_necessarias)}")
                    return
                
                if 'Razão Social' not in df_banco.columns or 'CPF/CNPJ' not in df_banco.columns:
                    st.error("A planilha de referência deve conter as colunas 'Razão Social' e 'CPF/CNPJ'")
                    return
                
                # Processar a planilha
                df_processado = preencher_documentos(df_input, df_banco)
                
                # Gerar arquivo para download
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_processado.to_excel(writer, index=False, sheet_name='Dados Processados')
                
                st.success("Processamento concluído!")
                
                # Mostrar prévia
                st.subheader("Prévia do Resultado")
                st.dataframe(df_processado.head())
                
                # Botão de download
                st.download_button(
                    label="⬇️ Baixar Planilha Processada",
                    data=output.getvalue(),
                    file_name="planilha_processada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"Erro no processamento: {str(e)}")
            st.error("Verifique se os formatos das planilhas estão corretos")

if __name__ == "__main__":
    main()
