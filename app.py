import streamlit as st
import os
import tempfile
import xml.etree.ElementTree as ET
from docxtpl import DocxTemplate
from datetime import datetime
from decimal import Decimal
from num2words import num2words
import zipfile
import io
import re
import base64

# Variáveis globais para palavras-chave
PALAVRAS_CHAVE_MARACANAU = {
    "MARACANAU_SEFIN": ["FINANÇAS", "PAPEL"],
    "MARACANAU_EDUCACAO": ["EDUCAÇÃO", "PAPEL"],
    "MARACANAU_SAUDE": ["SAÚDE", "PAPEL"],
    "MARACANAU_SASC": ["CIDADANIA", "PAPEL"],
    "MARACANAU_ARQUIVO": ["ARQUIVÍSTICO", "LAPSO"]
}

PALAVRAS_CHAVE_PACATUBA = {
    "PACATUBA_ADM": ["PACATUBA", "ADMINISTRAÇÃO"],
    "PACATUBA_EDUCACAO": ["PACATUBA", "EDUCAÇÃO"],
    "PACATUBA_INFRA": ["PACATUBA", "INFRAESTRUTURA"],
    "PACATUBA_IPMP": ["PACATUBA", "SERVIDORES"],
    "PACATUBA_FMAS": ["PACATUBA", "HUMANOS"],
    "PACATUBA_SAUDE": ["PACATUBA", "SAÚDE"]
}

def formatar_moeda_brasileira(valor):
    valor_str = f'{valor:,.2f}'
    return f'R$ {valor_str.replace(".", ",").replace(",", ".", 1)}'

def decimal_para_extenso(valor):
    parte_inteira = int(valor)
    centavos = int(round((valor % 1) * 100))
    extenso = num2words(parte_inteira, lang='pt_BR') + ' reais'
    if centavos > 0:
        extenso += ' e ' + num2words(centavos, lang='pt_BR') + ' centavos'
    return extenso.capitalize()

def formatar_competencia(competencia):
    data_obj = datetime.strptime(competencia, '%Y-%m-%d')
    meses_pt = {
        1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
        5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
        9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
    }
    mes_nome = meses_pt[data_obj.month]
    ano = data_obj.year
    return f"{mes_nome} {ano}"

def extrair_informacoes_xml(xml_file):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        infNfse = root.find('.//InfNfse')
        if infNfse is None:
            print(f"Aviso: Tag 'InfNfse' não encontrada no XML {xml_file}.")
            return None

        numeroNF = infNfse.find('.//Numero')
        data = infNfse.find('.//DataEmissao')
        valor = infNfse.find('.//ValorServicos')
        competencia = infNfse.find('.//Competencia')
        discriminacao = infNfse.find('.//Discriminacao')

        numeroNF = numeroNF.text if numeroNF is not None else "N/A"
        data = data.text if data is not None else None
        valor = Decimal(valor.text) if valor is not None else None
        competencia = competencia.text if competencia is not None else None
        discriminacao_text = discriminacao.text if discriminacao is not None else ""

        match_quant = re.search(r'(\d+)\s+R\$', discriminacao_text)
        quant = match_quant.group(1) if match_quant else "N/A"

        data_formatada = "N/A"
        competencia_formatada = "N/A"
        if data:
            try:
                data_obj = datetime.strptime(data, '%Y-%m-%d')
                meses_pt = {
                    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
                    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
                    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
                }
                dia = data_obj.day
                mes_nome = meses_pt[data_obj.month]
                ano = data_obj.year
                data_formatada = f"{dia} de {mes_nome} de {ano}"
            except ValueError:
                print(f"Aviso: Data inválida no XML {xml_file}.")

        if competencia:
            try:
                competencia_formatada = formatar_competencia(competencia)
            except ValueError:
                print(f"Aviso: Competência inválida no XML {xml_file}.")

        dados = {
            "numeroNF": numeroNF,
            "data": data_formatada,
            "valor": formatar_moeda_brasileira(valor) if valor else "N/A",
            "valor_extenso": decimal_para_extenso(valor) if valor else "N/A",
            "competencia": competencia_formatada,
            "discriminacao": discriminacao_text,
            "quant": quant
        }
        return dados
    except Exception as e:
        print(f"Erro ao processar o XML {xml_file}: {str(e)}")
        return None

def identificar_template(dados, palavras_chave):
    discriminacao = dados["discriminacao"].upper()
    for template, palavras in palavras_chave.items():
        if all(palavra.upper() in discriminacao for palavra in palavras):
            return template
    return None

def gerar_documentos(template_dir, dados, xml_file_name):
    arquivos_gerados = []
    templates = ["Planilha.docx", "Relatorio.docx"]
    for template_name in templates:
        template_path = os.path.join(template_dir, template_name)
        if not os.path.exists(template_path):
            print(f"Template '{template_path}' não encontrado. Pulando...")
            continue

        # Nome do arquivo de saída
        output_filename = f"{os.path.splitext(xml_file_name)[0]}_{template_name}"

        # Cria um arquivo temporário para salvar o documento gerado
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
            destino_path = temp_file.name

        doc = DocxTemplate(template_path)
        doc.render(dados)
        doc.save(destino_path)
        print(f"Documento gerado: {destino_path}")
        arquivos_gerados.append(destino_path)
    return arquivos_gerados

def processar_todos_xmls(xml_dir, templates_base_dir, palavras_chave):
    resultados = []
    for xml_file in os.listdir(xml_dir):
        if xml_file.endswith('.xml'):
            xml_path = os.path.join(xml_dir, xml_file)
            dados = extrair_informacoes_xml(xml_path)
            if dados is None:
                print(f"Pulando o arquivo {xml_file} devido a erros.")
                continue

            template_folder = identificar_template(dados, palavras_chave)
            if template_folder is None:
                print(f"Nenhum template correspondente encontrado para {xml_file}. Pulando...")
                continue

            template_dir = os.path.join(templates_base_dir, template_folder)
            arquivos_gerados = gerar_documentos(template_dir, dados, xml_file)
            resultados.extend(arquivos_gerados)
    return resultados


def gerar_documentos_em_memoria(template_dir, dados, xml_file_name):
    arquivos_gerados = []
    templates = ["Planilha.docx", "Relatorio.docx"]
    st.text(f"Templates disponíveis para {xml_file_name}: {templates}")

    for template_name in templates:
        template_path = os.path.join(template_dir, template_name)
        if not os.path.exists(template_path):
            st.warning(f"Template '{template_path}' não encontrado. Pulando...")
            continue

        try:
            output_filename = f"{os.path.splitext(xml_file_name)[0]}_{template_name}"
            st.text(f"Gerando documento: {output_filename}")

            with io.BytesIO() as temp_file:
                doc = DocxTemplate(template_path)
                doc.render(dados)
                doc.save(temp_file)
                temp_file.seek(0)
                arquivos_gerados.append((output_filename, temp_file.read()))
                st.text(f"Documento '{output_filename}' gerado com sucesso.")
        except Exception as e:
            st.error(f"Erro ao gerar documento '{template_name}': {e}")
            continue

    return arquivos_gerados

def criar_zip_em_memoria(arquivos_gerados):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for nome_arquivo, conteudo in arquivos_gerados:
            zip_file.writestr(nome_arquivo, conteudo)
    zip_buffer.seek(0)  # Redefine o ponteiro para o início
    return zip_buffer

def processar_xmls(cidade, arquivos, palavras_chave):
    st.write(f"Processando arquivos para {cidade}...")
    st.text(f"Iniciando processamento de {len(arquivos)} arquivo(s) XML para {cidade}.")

    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            # Log do diretório temporário
            st.text(f"Diretório temporário criado: {temp_dir}")
            for uploaded_file in arquivos:
                file_path = os.path.join(temp_dir, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                st.text(f"Arquivo salvo no diretório temporário: {file_path}")

            templates_base_dir = os.path.abspath(f"{cidade.upper()}")
            st.text(f"Diretório base dos templates: {templates_base_dir}")

            arquivos_gerados = []
            for xml_file in os.listdir(temp_dir):
                if xml_file.endswith(".xml"):
                    xml_path = os.path.join(temp_dir, xml_file)
                    st.text(f"Processando arquivo XML: {xml_path}")

                    dados = extrair_informacoes_xml(xml_path)
                    if dados is None:
                        st.warning(f"Pulando o arquivo {xml_file} devido a erros.")
                        continue
                    st.text(f"Dados extraídos do XML: {dados}")

                    template_folder = identificar_template(dados, palavras_chave)
                    if template_folder is None:
                        st.warning(f"Nenhum template correspondente encontrado para {xml_file}. Pulando...")
                        continue
                    st.text(f"Template identificado: {template_folder}")

                    template_dir = os.path.join(templates_base_dir, template_folder)
                    st.text(f"Diretório do template: {template_dir}")

                    arquivos = gerar_documentos_em_memoria(template_dir, dados, xml_file)
                    st.text(f"Documentos gerados para {xml_file}: {len(arquivos)} arquivo(s).")
                    arquivos_gerados.extend(arquivos)

            if arquivos_gerados:
                st.success(f"Processamento concluído para {len(arquivos_gerados)} arquivo(s).")

                # Cria o ZIP em memória
                zip_buffer = criar_zip_em_memoria(arquivos_gerados)
                st.text("Arquivo ZIP gerado em memória.")

                # Download do ZIP
                st.download_button(
                    label="⬇️ Baixar todos os documentos (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"documentos_{cidade.lower()}.zip",
                    mime="application/zip",
                )
            else:
                st.warning("Nenhum arquivo foi processado.")
        except Exception as e:
            st.error(f"Erro ao processar XMLs: {str(e)}")
            st.text(f"Erro detalhado: {e}")

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

if __name__ == "__main__":
    main()
