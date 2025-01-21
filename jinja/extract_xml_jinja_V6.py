import os
import xml.etree.ElementTree as ET
from docxtpl import DocxTemplate
from datetime import datetime
from decimal import Decimal
from num2words import num2words
import re
import tempfile

# LER O README

# inserir palavras que contém dentro do XML que sirva como chave/gancho para poder identificar o template.
# por ex: "MARACANAU_ARQUIVO": ["ARQUIVÍSTICO", "LAPSO"] - a nota fiscal do arquivo é a unica que contém as palavras ARQUIVÍSTICO e LAPSO. Sendo assim ele vai acusar prontamente o templete MARACANAU_ARQUIVO.

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

# formatar o valor no xml que esta como dolar
def formatar_moeda_brasileira(valor):
    valor_str = f'{valor:,.2f}'
    return f'R$ {valor_str.replace(".", ",").replace(",", ".", 1)}'

# transformar o texto por extenso
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

        output_filename = f"{os.path.splitext(xml_file_name)[0]}_{template_name}"
        destino_path = os.path.join(template_dir, output_filename)

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


# if __name__ == "__main__":
#     print("Escolha qual conjunto de XMLs processar:")
#     print("1. Maracanaú")
#     print("2. Pacatuba")
#     escolha = input("Digite o número da sua escolha: ")

#     if escolha == '1':
#         xml_dir = "jinja/xml_maracanau"
#         templates_base_dir = "jinja/MARACANAU"
#         palavras_chave = PALAVRAS_CHAVE_MARACANAU
#         print("Processando XMLs de Maracanaú...")
#     elif escolha == '2':
#         xml_dir = "jinja/xml_pacatuba"
#         templates_base_dir = "jinja/PACATUBA"
#         palavras_chave = PALAVRAS_CHAVE_PACATUBA
#         print("Processando XMLs de Pacatuba...")
#     else:
#         print("Escolha inválida. Por favor, execute o script novamente e escolha 1 ou 2.")
#         exit()

#     processar_todos_xmls(xml_dir, templates_base_dir, palavras_chave)