import streamlit as st
import os
import tempfile
from extract_xml_jinja_V6 import processar_todos_xmls, PALAVRAS_CHAVE_MARACANAU, PALAVRAS_CHAVE_PACATUBA
import zipfile
import io

def main():
    st.title("Gerador de Relatórios NFs ASP")
    st.write("Faça o upload dos arquivos XML e escolha a cidade para processá-los.")

    uploaded_files = st.file_uploader("Faça o upload dos arquivos XML", accept_multiple_files=True, type="xml")

    col1, col2 = st.columns(2)
    with col1:
        maracanau_button = st.button("MARACANAU")
    with col2:
        pacatuba_button = st.button("PACATUBA")

    if uploaded_files:
        st.write("Arquivos carregados:")
        for file in uploaded_files:
            st.write(file.name)

    if maracanau_button and uploaded_files:
        processar_xmls("MARACANAU", uploaded_files, PALAVRAS_CHAVE_MARACANAU)
    elif pacatuba_button and uploaded_files:
        processar_xmls("PACATUBA", uploaded_files, PALAVRAS_CHAVE_PACATUBA)

def criar_zip_documentos(arquivos_gerados):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for arquivo in arquivos_gerados:
            if os.path.exists(arquivo):
                # Adiciona arquivo ao ZIP mantendo apenas o nome do arquivo, sem o caminho completo
                zip_file.write(arquivo, os.path.basename(arquivo))
    return zip_buffer

def processar_xmls(cidade, arquivos, palavras_chave):
    st.write(f"Processando arquivos para {cidade}...")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        # Salvar arquivos no diretório temporário
        for uploaded_file in arquivos:
            file_path = os.path.join(temp_dir, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
        
        templates_base_dir = os.path.abspath(f"{cidade.upper()}")
        
        try:
            arquivos_gerados = processar_todos_xmls(temp_dir, templates_base_dir, palavras_chave)
            
            if arquivos_gerados:
                st.success(f"Processamento concluído para {len(arquivos_gerados)} arquivo(s).")
                
                # Criar arquivo ZIP com todos os documentos
                zip_buffer = criar_zip_documentos(arquivos_gerados)
                
                # Botão para download do ZIP com todos os arquivos
                st.download_button(
                    label="⬇️ Baixar todos os documentos (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"documentos_{cidade.lower()}.zip",
                    mime="application/zip",
                    key="download_todos"
                )
                
                # Exibir lista de arquivos gerados
                st.write("Arquivos gerados:")
                for arquivo in arquivos_gerados:
                    st.write(os.path.basename(arquivo))
                    
            else:
                st.warning("Nenhum arquivo foi processado.")
        except Exception as e:
            st.error(f"Erro ao processar XMLs: {str(e)}")

if __name__ == "__main__":
    main()
