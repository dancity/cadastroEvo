import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import re

def limpar_nome_arquivo(nome):
    """
    Remove caracteres não seguros de nomes de arquivo.
    
    :param nome: O nome do arquivo a ser limpo.
    :return: Uma string com o nome do arquivo limpo.
    """
    nome_limpo = re.sub(r'[^\w\s\-_]', '', str(nome))
    nome_limpo = re.sub(r'\s+', '_', nome_limpo)
    return nome_limpo

#Lyceum
def preparar_df_sistema_a(df_original, senha_padrao):
    colunas_obrig = ["UNIDADE", "ALUNO", "NOME_COMPL", "TURMA", "CURSO", "CODSERIE"]
    faltando = [c for c in colunas_obrig if c not in df_original.columns]
    if faltando:
        raise ValueError(f"O seu relatório do Lyceum não possui as seguintes colunas: {faltando}")

    if "Status" in df_original.columns:
        df_filtrado = df_original[df_original["Status"] == "Ativo"].copy()
    else:
        df_filtrado = df_original.copy()

    # Aplicar filtro para excluir linhas com "SEM TURMA" ou turmas que contenham "P"
    df_filtrado = df_filtrado[~df_filtrado["TURMA"].str.contains("SEM TURMA", case=False, na=False)]
    df_filtrado = df_filtrado[~df_filtrado["TURMA"].str.contains("P", case=False, na=False)]

    # Aplicar filtro para manter apenas "ENSINO FUNDAMENTAL" e "ENSINO MÉDIO" na coluna "CURSO"
    df_filtrado = df_filtrado[df_filtrado["CURSO"].isin(["ENSINO FUNDAMENTAL", "ENSINO MÉDIO"])]

    df_filtrado["Ano/Série"] = df_filtrado.apply(
        lambda row: f"{row['CODSERIE']}º ano do Ensino Fundamental" if row['CURSO'] == "ENSINO FUNDAMENTAL" 
        else f"{row['CODSERIE']}ª série do Ensino Médio", axis=1
    )

    df_final = pd.DataFrame()
    df_final["RA"] = df_filtrado["ALUNO"]
    df_final["Nome"] = df_filtrado["NOME_COMPL"].astype(str)
    
    # Alterar o campo "E-mail"
    df_final["E-mail"] = df_filtrado["ALUNO"].astype(str) + "@maristabrasil.g12.br"
    
    df_final["Senha"] = senha_padrao if senha_padrao else ""
    df_final["Ano/Série"] = df_filtrado["Ano/Série"]
    df_final["Turma"] = df_filtrado["TURMA"]
    
    return df_final, df_filtrado

def gerar_tabela_turmas(df_filtrado):
    if "Ano/Série" not in df_filtrado.columns:
        raise ValueError("A coluna 'Ano/Série' não foi encontrada no DataFrame filtrado.")
    turmas_df = df_filtrado.groupby(["UNIDADE", "Ano/Série", "TURMA"], as_index=False).size()
    turmas_df = turmas_df.drop(columns="size")  # Remover a coluna desnecessária gerada por groupby
    return turmas_df

def preparar_df_sistema_b(df_original, senha_padrao):
    st.warning("A funcionalidade para o sistema Prime está desativada no momento.")
    return pd.DataFrame()

def preparar_df_sistema_c(df_original, senha_padrao):
    st.warning("A funcionalidade para o sistema GVDasa está desativada no momento.")
    return pd.DataFrame()

def main():
    st.title('Gerador de tabela para cadastro na Evolucional')
    st.write("Este aplicativo gera uma tabela formatada para o cadastro na plataforma Evolucional a partir de dados extraídos do seu sistema acadêmico.")
    
    sistema = st.selectbox(
        "Selecione o sistema acadêmico utilizado pela sua escola:",
        ["Lyceum", "Prime", "GVDasa"]
    )
    
    senha_padrao = st.text_input("Informe uma senha padrão para todos os alunos:", type="password")
    if not senha_padrao:
        st.error("A senha padrão é obrigatória.")
        return

    st.write("Por favor, envie o arquivo Excel ou CSV gerado pelo seu sistema acadêmico. O arquivo não será enviado para nenhum servidor, todo o processamento acontece localmente.")
    
    uploaded_file = st.file_uploader("Selecione o arquivo Excel ou CSV", type=["xlsx", "xls", "csv"])
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                df_original = pd.read_csv(uploaded_file, sep=';', decimal=',')
            else:
                df_original = pd.read_excel(uploaded_file)
            
            st.write("Prévia do arquivo original:")
            st.dataframe(df_original.head())

            if sistema == "Lyceum":
                df_final, df_filtrado = preparar_df_sistema_a(df_original, senha_padrao)
                turmas_df = gerar_tabela_turmas(df_filtrado)
            elif sistema == "Prime":
                df_final = preparar_df_sistema_b(df_original, senha_padrao)
                turmas_df = pd.DataFrame()
            else:
                df_final = preparar_df_sistema_c(df_original, senha_padrao)
                turmas_df = pd.DataFrame()
            
            if df_final.empty:
                st.stop()

            st.write("Prévia do arquivo formatado:")
            st.dataframe(df_final.head())
            
            if "UNIDADE" in df_original.columns:
                unidades = df_original["UNIDADE"].unique()
            elif "Unidade" in df_original.columns:
                unidades = df_original["Unidade"].unique()
            else:
                unidades = ["Todas"]
            
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for unidade in unidades:
                    if unidade == "Todas":
                        df_unidade = df_final.copy()
                        turmas_unidade = turmas_df.copy()
                        nome_cadastro = "Cadastro_Todas.xlsx"
                        nome_turmas = "Turmas_Todas.xlsx"
                    else:
                        if "UNIDADE" in df_original.columns:
                            mask = df_original["UNIDADE"] == unidade
                        else:
                            mask = df_original["Unidade"] == unidade

                        df_unidade = df_final[mask].copy()
                        turmas_unidade = turmas_df[turmas_df["UNIDADE"] == unidade].copy()
                        nome_cadastro = f"Cadastro_{limpar_nome_arquivo(unidade)}.xlsx"
                        nome_turmas = f"Turmas_{limpar_nome_arquivo(unidade)}.xlsx"

                    if len(df_unidade) > 0:
                        unidade_buffer = BytesIO()
                        with pd.ExcelWriter(unidade_buffer, engine='xlsxwriter') as writer:
                            df_unidade.to_excel(writer, index=False, sheet_name='Cadastro')
                        zip_file.writestr(nome_cadastro, unidade_buffer.getvalue())

                        turmas_buffer = BytesIO()
                        with pd.ExcelWriter(turmas_buffer, engine='xlsxwriter') as writer:
                            turmas_unidade.to_excel(writer, index=False, sheet_name='Turmas')
                        zip_file.writestr(nome_turmas, turmas_buffer.getvalue())

            if len(unidades) > 1:
                st.download_button(
                    label="Baixar arquivos formatados (.zip)",
                    data=zip_buffer.getvalue(),
                    file_name="cadastros_unidades.zip",
                    mime="application/zip"
                )
            else:
                with pd.ExcelWriter("Cadastro_Unidade.xlsx", engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Cadastro')
                st.download_button(
                    label="Baixar arquivo formatado (.xlsx)",
                    data=unidade_buffer.getvalue(),
                    file_name="Cadastro_Unidade.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                with pd.ExcelWriter("Turmas_Unidade.xlsx", engine='xlsxwriter') as writer:
                    turmas_df.to_excel(writer, index=False, sheet_name='Turmas')
                st.download_button(
                    label="Baixar arquivo de turmas (.xlsx)",
                    data=turmas_buffer.getvalue(),
                    file_name="Turmas_Unidade.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error("Ocorreu um erro ao processar o arquivo. Por favor, verifique se o arquivo está correto.")
            st.error(str(e))

if __name__ == "__main__":
    main()
