import streamlit as st
import tabula
import pandas as pd
import os

# Título da aplicação
st.title("Conversor de PDF para Excel")
st.markdown("Faça o upload de um arquivo PDF e converta as tabelas para um arquivo Excel.")

# Widget para upload do arquivo
uploaded_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")

if uploaded_file is not None:
    # Cria uma pasta temporária para salvar o PDF
    os.makedirs("temp", exist_ok=True)
    temp_pdf_path = os.path.join("temp", uploaded_file.name)
    
    # Salva o arquivo enviado para processamento
    with open(temp_pdf_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success(f"Arquivo '{uploaded_file.name}' carregado com sucesso!")
    
    # Botão para iniciar a conversão
    if st.button("Converter para Excel"):
        try:
            # Lê todas as tabelas do PDF
            st.info("Lendo tabelas do PDF... Isso pode levar alguns segundos.")
            tabelas = tabula.read_pdf(temp_pdf_path, pages='all', multiple_tables=True)

            if not tabelas:
                st.warning("Nenhuma tabela encontrada no PDF.")
            else:
                # Cria um arquivo Excel temporário
                excel_path = os.path.join("temp", uploaded_file.name.replace(".pdf", ".xlsx"))
                
                with pd.ExcelWriter(excel_path) as writer:
                    for i, tabela in enumerate(tabelas):
                        tabela.to_excel(writer, sheet_name=f"Tabela_{i+1}", index=False)
                
                # Botão para download do arquivo Excel
                with open(excel_path, "rb") as f:
                    st.download_button(
                        label="Clique para baixar o arquivo Excel",
                        data=f,
                        file_name=os.path.basename(excel_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                st.balloons()  # Efeito de balões para comemorar!
                st.success("Conversão concluída com sucesso!")

        except Exception as e:
            st.error(f"Ocorreu um erro durante a conversão: {e}")
        finally:
            # Limpa os arquivos temporários
            os.remove(temp_pdf_path)
            if 'excel_path' in locals() and os.path.exists(excel_path):
                os.remove(excel_path)