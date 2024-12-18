import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile
from io import BytesIO

def main():
    st.title('Gerador de tabela para cadastro na Evolucional')
    st.write("Este aplicativo gera uma tabela formatada para o cadastro na plataforma Evolucional a partir de dados extraídos do seu sistema acadêmico.")

    # Opção de seleção do sistema acadêmico
    sistema = st.selectbox(
        "Selecione o sistema acadêmico utilizado pela sua escola:",
        ["Lyceum", "Prime", "GVDasa"]
    )
    
    st.write("Por favor, envie o arquivo Excel gerado pelo seu sistema acadêmico. O arquivo não será enviado para nenhum servidor, todo o processamento acontece localmente.")
    
    # Upload do arquivo Excel
    uploaded_file = st.file_uploader("Selecione o arquivo Excel", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # Leitura do arquivo Excel
        try:
            df_original = pd.read_excel(uploaded_file)
            
            # Exibição do DataFrame original (opcional, para conferência)
            st.write("Visualização do arquivo original:")
            st.dataframe(df_original)
            
            # Aqui faríamos as validações necessárias no df_original
            # Exemplo de verificação: garantir que colunas obrigatórias existam
            colunas_obrigatorias = ["Nome", "Sobrenome", "Matricula", "Turma"]
            colunas_faltando = [col for col in colunas_obrigatorias if col not in df_original.columns]
            
            if len(colunas_faltando) > 0:
                st.error(f"As seguintes colunas obrigatórias estão faltando no arquivo: {colunas_faltando}")
                st.stop()
            
            # Dependendo do sistema selecionado, ajustamos o DataFrame
            if sistema == "Sistema A":
                # Exemplo de formatação específica para o Sistema A
                # Vamos supor que o Sistema A precisa da coluna "Nome Completo" em vez de duas colunas separadas
                df_original["Nome Completo"] = df_original["Nome"] + " " + df_original["Sobrenome"]
                # Ajuste de algumas colunas, renomear, filtrar, etc.
                df_final = df_original[["Nome Completo", "Matricula", "Turma"]].copy()
                df_final["Plataforma"] = "Evolucional"
                
            elif sistema == "Sistema B":
                # Exemplo de formatação específica para o Sistema B
                # Digamos que o Sistema B já vem com uma coluna "Nome Completo" mas precisamos separar nome e sobrenome
                # (Exemplo fictício, só para demonstrar)
                df_original[["Nome", "Sobrenome"]] = df_original["Nome Completo"].str.split(" ", 1, expand=True)
                df_final = df_original[["Nome", "Sobrenome", "Matricula", "Turma"]].copy()
                df_final["Plataforma"] = "Evolucional"
                
            else:  # Sistema C
                # Exemplo de formatação específica para o Sistema C
                # Vamos supor que precisamos filtrar apenas alunos ativos
                df_final = df_original[df_original["Status"] == "Ativo"].copy()
                # Renomear colunas para adequar ao padrão da Evolucional
                df_final.rename(columns={"Matricula": "Matrícula"}, inplace=True)
                df_final["Plataforma"] = "Evolucional"
            
            # Exibição do DataFrame final
            st.write("Visualização do DataFrame final formatado:")
            st.dataframe(df_final)
            
            # Opção de baixar o resultado em Excel
            # Criar um objeto em memória
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Cadastro')
            
            # Botão de download 
            st.download_button(
                label="Baixar tabela formatada",
                data=output.getvalue(),
                file_name="tabela_formatada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error("Ocorreu um erro ao processar o arquivo. Por favor, verifique se o arquivo está correto.")
            st.error(str(e))

if __name__ == "__main__":
    main()