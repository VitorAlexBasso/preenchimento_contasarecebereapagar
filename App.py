import pandas as pd
import streamlit as st
from io import BytesIO

def main():
    st.title("Preenchimento de CPF/CNPJ em Contas a Receber")
    
    # Upload da planilha de banco de dados (fonte de pesquisa)
    st.header("1. Banco de Dados para Pesquisa")
    db_file = st.file_uploader("Carregue a planilha com os dados de CPF/CNPJ (Banco de Dados)", type=["xlsx", "xls"])
    
    # Upload da planilha a ser preenchida
    st.header("2. Planilha a Ser Preenchida")
    receber_file = st.file_uploader("Carregue a planilha de Contas a Receber", type=["xlsx", "xls"])
    
    if db_file and receber_file:
        try:
            # Carregar os dados
            db_df = pd.read_excel(db_file)
            receber_df = pd.read_excel(receber_file)
            
            # Verificar colunas necessárias
            if 'Nome da Pessoa' not in receber_df.columns or 'CPF' not in receber_df.columns:
                st.error("A planilha de Contas a Receber deve conter as colunas 'Nome da Pessoa' e 'CPF'")
                return
                
            if 'Nome' not in db_df.columns or 'CPF_CNPJ' not in db_df.columns:
                st.error("A planilha de Banco de Dados deve conter as colunas 'Nome' e 'CPF_CNPJ'")
                return
            
            # Criar dicionário de mapeamento (nome -> cpf/cnpj)
            mapeamento = dict(zip(db_df['Nome'], db_df['CPF_CNPJ']))
            
            # Preencher os CPFs/CNPJs
            receber_df['CPF'] = receber_df['Nome da Pessoa'].map(mapeamento).fillna('nan')
            
            # Preparar para download
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                receber_df.to_excel(writer, index=False)
            
            st.success("Planilha processada com sucesso!")
            
            # Botão de download
            st.download_button(
                label="Baixar Planilha Preenchida",
                data=output.getvalue(),
                file_name="contas_a_receber_preenchida.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Ocorreu um erro: {str(e)}")

if __name__ == "__main__":
    main()
