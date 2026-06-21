import base64
import io
import os
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import streamlit as st
from docxtpl import DocxTemplate
from num2words import num2words


BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"
LOGO_PATH = ASSETS_DIR / "logo.svg"

PALAVRAS_CHAVE_MARACANAU = {
    "MARACANAU_SEFIN": ["FINANÇAS", "PAPEL"],
    "MARACANAU_EDUCACAO": ["EDUCAÇÃO", "PAPEL"],
    "MARACANAU_SAUDE": ["SAÚDE", "PAPEL"],
    "MARACANAU_SASC": ["ALIMENTAR", "SEGURANÇA"],
    "MARACANAU_ARQUIVO": ["ARQUIVÍSTICO", "LAPSO"],
}

PALAVRAS_CHAVE_PACATUBA = {
    "PACATUBA_ADM": ["PACATUBA", "ADMINISTRAÇÃO"],
    "PACATUBA_EDUCACAO": ["PACATUBA", "EDUCAÇÃO"],
    "PACATUBA_INFRA": ["PACATUBA", "INFRAESTRUTURA"],
    "PACATUBA_IPMP": ["PACATUBA", "SERVIDORES"],
    "PACATUBA_FMAS": ["PACATUBA", "HUMANOS"],
    "PACATUBA_SAUDE": ["PACATUBA", "SAÚDE"],
}


def carregar_svg_base64(path):
    if not path.exists():
        return ""
    return base64.b64encode(path.read_bytes()).decode("utf-8")


def aplicar_estilos():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;900&family=Space+Grotesk:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap');

        :root {{
            --navy: #060D1A;
            --navy2: #0C1A33;
            --gold: #C9A84C;
            --gold2: #E8C96A;
            --cyan: #4FD1E0;
            --gray: #8A96AC;
            --line: rgba(201, 168, 76, 0.18);
        }}

        .stApp {{
            background:
                radial-gradient(circle at 8% 0%, rgba(201, 168, 76, 0.16), transparent 32rem),
                radial-gradient(circle at 90% 40%, rgba(79, 209, 224, 0.11), transparent 28rem),
                linear-gradient(rgba(201, 168, 76, 0.04) 1px, transparent 1px),
                linear-gradient(90deg, rgba(201, 168, 76, 0.04) 1px, transparent 1px),
                var(--navy);
            background-size: auto, auto, 58px 58px, 58px 58px, auto;
            color: #FFFFFF;
            font-family: 'Space Grotesk', sans-serif;
        }}

        .stApp::before {{
            content: "";
            position: fixed;
            inset: 0;
            pointer-events: none;
            background-image: url("data:image/svg+xml,%3Csvg viewBox='0 0 200 200' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)' opacity='0.035'/%3E%3C/svg%3E");
            opacity: 0.7;
            z-index: 0;
        }}

        [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"], #MainMenu, footer {{
            visibility: hidden;
            height: 0;
        }}

        [data-testid="stAppViewContainer"] > .main {{
            position: relative;
            z-index: 1;
        }}

        .block-container {{
            max-width: 1180px;
            padding: 2.2rem 2rem 4rem;
        }}

        h1, h2, h3, p, label, span, div {{
            font-family: 'Space Grotesk', sans-serif;
        }}

        .brand-shell {{
            display: grid;
            grid-template-columns: minmax(0, 1.15fr) minmax(320px, 0.85fr);
            gap: 2.2rem;
            align-items: stretch;
            margin-bottom: 1.8rem;
        }}

        .hero-panel, .control-panel, .status-panel {{
            position: relative;
            overflow: hidden;
            border: 1px solid var(--line);
            border-radius: 14px;
            background: linear-gradient(160deg, rgba(12, 26, 51, 0.88), rgba(6, 13, 26, 0.72));
            box-shadow: 0 24px 70px rgba(0, 0, 0, 0.35), inset 0 0 45px rgba(201, 168, 76, 0.04);
        }}

        .hero-panel {{
            min-height: 430px;
            padding: 2.4rem;
        }}

        .hero-panel::after, .control-panel::after, .status-panel::after {{
            content: "";
            position: absolute;
            top: 16px;
            right: 16px;
            width: 22px;
            height: 22px;
            border-top: 1px solid var(--gold);
            border-right: 1px solid var(--gold);
            opacity: 0.45;
        }}

        .brand-logo {{
            width: 188px;
            max-width: 70%;
            margin-bottom: 2.4rem;
        }}

        .eyebrow {{
            display: inline-flex;
            align-items: center;
            gap: 0.6rem;
            color: var(--gold);
            border: 1px solid var(--line);
            border-radius: 999px;
            background: rgba(201, 168, 76, 0.08);
            padding: 0.45rem 0.85rem;
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.72rem;
            letter-spacing: 0.12rem;
            text-transform: uppercase;
        }}

        .eyebrow-dot {{
            width: 7px;
            height: 7px;
            border-radius: 999px;
            background: var(--cyan);
            box-shadow: 0 0 12px rgba(79, 209, 224, 0.9);
        }}

        .hero-title {{
            margin: 1.4rem 0 1rem;
            max-width: 720px;
            font-family: 'Playfair Display', serif;
            font-size: clamp(2.4rem, 5vw, 4.6rem);
            font-weight: 900;
            line-height: 1.03;
            letter-spacing: 0;
            color: #FFFFFF;
        }}

        .hero-title span {{
            font-family: 'Playfair Display', serif;
            color: var(--gold2);
            text-shadow: 0 0 26px rgba(201, 168, 76, 0.32);
        }}

        .hero-sub {{
            max-width: 620px;
            color: var(--gray);
            font-size: 1.04rem;
            line-height: 1.7;
        }}

        .hero-metrics {{
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 0.9rem;
            margin-top: 2.2rem;
        }}

        .metric-card {{
            border: 1px solid rgba(201, 168, 76, 0.12);
            border-radius: 10px;
            background: rgba(255, 255, 255, 0.025);
            padding: 1rem;
        }}

        .metric-card strong {{
            display: block;
            color: var(--gold);
            font-family: 'JetBrains Mono', monospace;
            font-size: 1.25rem;
            margin-bottom: 0.25rem;
        }}

        .metric-card span {{
            color: var(--gray);
            font-size: 0.78rem;
        }}

        .control-panel {{
            padding: 1.6rem;
        }}

        .panel-title {{
            color: #FFFFFF;
            font-size: 1.05rem;
            font-weight: 700;
            margin: 0 0 0.3rem;
        }}

        .panel-copy {{
            color: var(--gray);
            font-size: 0.88rem;
            line-height: 1.55;
            margin-bottom: 1.2rem;
        }}

        .scan-card {{
            margin-top: 1.2rem;
            border: 1px solid rgba(79, 209, 224, 0.2);
            border-radius: 12px;
            padding: 1rem;
            background: rgba(79, 209, 224, 0.045);
        }}

        .scan-card code {{
            color: var(--cyan);
            background: transparent;
            font-family: 'JetBrains Mono', monospace;
        }}

        .file-chip {{
            display: inline-flex;
            align-items: center;
            margin: 0.2rem 0.25rem 0.2rem 0;
            padding: 0.42rem 0.72rem;
            border-radius: 999px;
            border: 1px solid rgba(201, 168, 76, 0.18);
            color: #F5F0E8;
            background: rgba(201, 168, 76, 0.07);
            font-size: 0.78rem;
            font-family: 'JetBrains Mono', monospace;
        }}

        .status-panel {{
            padding: 1.4rem;
            margin-top: 1.4rem;
        }}

        .stRadio [role="radiogroup"] {{
            gap: 0.6rem;
        }}

        .stFileUploader {{
            border: 1px dashed rgba(201, 168, 76, 0.32);
            border-radius: 12px;
            padding: 0.5rem 0.7rem 0.1rem;
            background: rgba(255, 255, 255, 0.025);
        }}

        .stFileUploader label, .stRadio label {{
            color: #F5F0E8 !important;
            font-weight: 600;
        }}

        .stButton > button, .stDownloadButton > button {{
            width: 100%;
            min-height: 3rem;
            border: 1px solid rgba(201, 168, 76, 0.28);
            border-radius: 7px;
            background: linear-gradient(135deg, var(--gold), var(--gold2));
            color: var(--navy);
            font-weight: 800;
            box-shadow: 0 0 26px rgba(201, 168, 76, 0.32);
        }}

        .stButton > button:hover, .stDownloadButton > button:hover {{
            border-color: var(--gold2);
            color: var(--navy);
            box-shadow: 0 0 34px rgba(201, 168, 76, 0.45);
            transform: translateY(-1px);
        }}

        .stAlert {{
            border-radius: 10px;
            border: 1px solid rgba(201, 168, 76, 0.18);
            background: rgba(12, 26, 51, 0.75);
        }}

        div[data-testid="stStatusWidget"] {{
            visibility: hidden;
            height: 0;
        }}

        @media (max-width: 880px) {{
            .block-container {{
                padding: 1.2rem 1rem 3rem;
            }}
            .brand-shell {{
                grid-template-columns: 1fr;
            }}
            .hero-panel {{
                min-height: auto;
                padding: 1.5rem;
            }}
            .hero-metrics {{
                grid-template-columns: 1fr;
            }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def renderizar_hero():
    logo_b64 = carregar_svg_base64(LOGO_PATH)
    logo_html = ""
    if logo_b64:
        logo_html = f'<img src="data:image/svg+xml;base64,{logo_b64}" alt="ASPdoc" class="brand-logo">'

    st.markdown(
        f"""
        <section class="hero-panel">
            {logo_html}
            <div class="eyebrow"><span class="eyebrow-dot"></span>GED ASP // automação documental</div>
            <h1 class="hero-title">Gerador de <span>relatórios NFS-e</span> com padrão ASP.</h1>
            <p class="hero-sub">
                Envie os XMLs, selecione a cidade e gere automaticamente as planilhas e relatórios
                em DOCX. O fluxo continua simples, agora com uma interface mais clara e alinhada ao
                site institucional.
            </p>
            <div class="hero-metrics">
                <div class="metric-card"><strong>XML</strong><span>Leitura automática dos dados da nota.</span></div>
                <div class="metric-card"><strong>DOCX</strong><span>Modelos oficiais por secretaria.</span></div>
                <div class="metric-card"><strong>ZIP</strong><span>Entrega consolidada para download.</span></div>
            </div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def renderizar_cabecalho_controles():
    st.markdown(
        """
        <section class="control-panel">
            <p class="panel-title">Central de processamento</p>
            <p class="panel-copy">
                Escolha o município, carregue um ou mais XMLs e clique em processar.
                Os documentos gerados ficam disponíveis em um único arquivo ZIP.
            </p>
        </section>
        """,
        unsafe_allow_html=True,
    )


def formatar_moeda_brasileira(valor):
    valor_str = f"{valor:,.2f}"
    return f"R$ {valor_str.replace('.', ',').replace(',', '.', 1)}"


def decimal_para_extenso(valor):
    parte_inteira = int(valor)
    centavos = int(round((valor % 1) * 100))
    extenso = num2words(parte_inteira, lang="pt_BR") + " reais"
    if centavos > 0:
        extenso += " e " + num2words(centavos, lang="pt_BR") + " centavos"
    return extenso.capitalize()


def formatar_competencia(competencia):
    data_obj = datetime.strptime(competencia, "%Y-%m-%d")
    meses_pt = {
        1: "Janeiro",
        2: "Fevereiro",
        3: "Março",
        4: "Abril",
        5: "Maio",
        6: "Junho",
        7: "Julho",
        8: "Agosto",
        9: "Setembro",
        10: "Outubro",
        11: "Novembro",
        12: "Dezembro",
    }
    return f"{meses_pt[data_obj.month]} {data_obj.year}"


def extrair_informacoes_xml(xml_file):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        infNfse = root.find(".//InfNfse")
        if infNfse is None:
            print(f"Aviso: tag 'InfNfse' não encontrada no XML {xml_file}.")
            return None

        numero_nf = infNfse.find(".//Numero")
        data = infNfse.find(".//DataEmissao")
        valor = infNfse.find(".//ValorServicos")
        competencia = infNfse.find(".//Competencia")
        discriminacao = infNfse.find(".//Discriminacao")

        numero_nf = numero_nf.text if numero_nf is not None else "N/A"
        data = data.text if data is not None else None
        valor = Decimal(valor.text) if valor is not None else None
        competencia = competencia.text if competencia is not None else None
        discriminacao_text = discriminacao.text if discriminacao is not None else ""

        match_quant = re.search(r"(\d+)\s+R\$", discriminacao_text)
        quant = match_quant.group(1) if match_quant else "N/A"

        data_formatada = "N/A"
        competencia_formatada = "N/A"
        if data:
            try:
                data_obj = datetime.strptime(data, "%Y-%m-%d")
                meses_pt = {
                    1: "Janeiro",
                    2: "Fevereiro",
                    3: "Março",
                    4: "Abril",
                    5: "Maio",
                    6: "Junho",
                    7: "Julho",
                    8: "Agosto",
                    9: "Setembro",
                    10: "Outubro",
                    11: "Novembro",
                    12: "Dezembro",
                }
                data_formatada = f"{data_obj.day} de {meses_pt[data_obj.month]} de {data_obj.year}"
            except ValueError:
                print(f"Aviso: data inválida no XML {xml_file}.")

        if competencia:
            try:
                competencia_formatada = formatar_competencia(competencia)
            except ValueError:
                print(f"Aviso: competência inválida no XML {xml_file}.")

        return {
            "numeroNF": numero_nf,
            "data": data_formatada,
            "valor": formatar_moeda_brasileira(valor) if valor else "N/A",
            "valor_extenso": decimal_para_extenso(valor) if valor else "N/A",
            "competencia": competencia_formatada,
            "discriminacao": discriminacao_text,
            "quant": quant,
        }
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
        if xml_file.endswith(".xml"):
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

    for template_name in templates:
        template_path = os.path.join(template_dir, template_name)
        if not os.path.exists(template_path):
            st.warning(f"Template '{template_path}' não encontrado. Pulando...")
            continue

        try:
            output_filename = f"{os.path.splitext(xml_file_name)[0]}_{template_name}"
            with io.BytesIO() as temp_file:
                doc = DocxTemplate(template_path)
                doc.render(dados)
                doc.save(temp_file)
                temp_file.seek(0)
                arquivos_gerados.append((output_filename, temp_file.read()))
        except Exception as e:
            st.error(f"Erro ao gerar documento '{template_name}': {e}")
            continue

    return arquivos_gerados


def criar_zip_em_memoria(arquivos_gerados):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for nome_arquivo, conteudo in arquivos_gerados:
            zip_file.writestr(nome_arquivo, conteudo)
    zip_buffer.seek(0)
    return zip_buffer


def renderizar_arquivos_carregados(uploaded_files):
    if not uploaded_files:
        st.markdown(
            '<div class="scan-card"><code>aguardando_upload</code><br>Selecione os XMLs para iniciar.</div>',
            unsafe_allow_html=True,
        )
        return

    chips = "".join(f'<span class="file-chip">{file.name}</span>' for file in uploaded_files)
    st.markdown(
        f"""
        <div class="scan-card">
            <code>{len(uploaded_files)} arquivo(s) pronto(s)</code><br>
            {chips}
        </div>
        """,
        unsafe_allow_html=True,
    )


def processar_xmls(cidade, arquivos, palavras_chave):
    st.markdown('<section class="status-panel">', unsafe_allow_html=True)
    st.markdown(f'<p class="panel-title">Processamento: {cidade.title()}</p>', unsafe_allow_html=True)

    progress = st.progress(0, text="Preparando arquivos...")
    detalhes = st.empty()

    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            for uploaded_file in arquivos:
                file_path = os.path.join(temp_dir, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

            templates_base_dir = os.path.abspath(cidade.upper())
            xmls = [xml_file for xml_file in os.listdir(temp_dir) if xml_file.endswith(".xml")]
            arquivos_gerados = []
            identificados = []

            for index, xml_file in enumerate(xmls, start=1):
                progress.progress(index / max(len(xmls), 1), text=f"Lendo {xml_file}...")
                xml_path = os.path.join(temp_dir, xml_file)
                dados = extrair_informacoes_xml(xml_path)
                if dados is None:
                    st.warning(f"Pulando {xml_file}: não foi possível ler a NFS-e.")
                    continue

                template_folder = identificar_template(dados, palavras_chave)
                if template_folder is None:
                    st.warning(f"Nenhum template correspondente encontrado para {xml_file}.")
                    continue

                identificados.append(f"{xml_file} -> {template_folder}")
                detalhes.markdown(
                    "<br>".join(f"<code>{item}</code>" for item in identificados),
                    unsafe_allow_html=True,
                )

                template_dir = os.path.join(templates_base_dir, template_folder)
                arquivos_gerados.extend(gerar_documentos_em_memoria(template_dir, dados, xml_file))

            if arquivos_gerados:
                progress.progress(1.0, text="ZIP pronto para download.")
                zip_buffer = criar_zip_em_memoria(arquivos_gerados)
                st.success(f"{len(arquivos_gerados)} documento(s) gerado(s) com sucesso.")
                st.download_button(
                    label="Baixar documentos em ZIP",
                    data=zip_buffer.getvalue(),
                    file_name=f"documentos_{cidade.lower()}.zip",
                    mime="application/zip",
                    type="primary",
                )
            else:
                progress.empty()
                st.warning("Nenhum arquivo foi processado. Verifique os XMLs e as palavras-chave.")
        except Exception as e:
            progress.empty()
            st.error(f"Erro ao processar XMLs: {str(e)}")

    st.markdown("</section>", unsafe_allow_html=True)


def main():
    st.set_page_config(
        page_title="ASPdoc | Gerador de Relatórios",
        page_icon="A",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    aplicar_estilos()

    hero_col, control_col = st.columns([1.15, 0.85], gap="large")

    with hero_col:
        renderizar_hero()

    with control_col:
        renderizar_cabecalho_controles()
        cidade = st.radio(
            "Município",
            ["MARACANAU", "PACATUBA"],
            horizontal=True,
            captions=["Templates Maracanaú", "Templates Pacatuba"],
        )
        uploaded_files = st.file_uploader(
            "Upload dos XMLs da NFS-e",
            accept_multiple_files=True,
            type="xml",
            help="Você pode selecionar vários XMLs de uma vez.",
        )
        renderizar_arquivos_carregados(uploaded_files)
        processar = st.button("Processar documentos", type="primary", use_container_width=True)

    st.markdown(
        """
        <section class="status-panel">
            <p class="panel-title">Como o fluxo decide o modelo</p>
            <p class="panel-copy">
                O sistema lê a discriminação da NFS-e e compara palavras-chave para escolher a pasta
                correta de templates. Depois renderiza a planilha e o relatório em DOCX.
            </p>
        </section>
        """,
        unsafe_allow_html=True,
    )

    if processar:
        if not uploaded_files:
            st.warning("Envie pelo menos um XML antes de processar.")
            return

        palavras = PALAVRAS_CHAVE_MARACANAU if cidade == "MARACANAU" else PALAVRAS_CHAVE_PACATUBA
        processar_xmls(cidade, uploaded_files, palavras)


if __name__ == "__main__":
    main()
