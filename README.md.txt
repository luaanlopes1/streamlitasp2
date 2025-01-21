# Gerador de Relatórios NFs ASP

## Como Adicionar Novas Cidades ou Templates

### 1. Adicionar Nova Cidade

Para adicionar uma nova cidade, você precisa:

1. Adicionar novo dicionário de palavras-chave em `extract_xml_jinja_V6.py`:

PALAVRAS_CHAVE_NOVA_CIDADE = {
"NOVA_CIDADE_SETOR1": ["NOVA_CIDADE", "PALAVRA_CHAVE1"],
"NOVA_CIDADE_SETOR2": ["NOVA_CIDADE", "PALAVRA_CHAVE2"],
# Adicione mais setores conforme necessário
}



2. Criar estrutura de diretórios:

PROJECT_extractXML/
├── NOVA_CIDADE/
│ ├── NOVA_CIDADE_SETOR1/
│ │ ├── Planilha.docx
│ │ └── Relatorio.docx
│ └── NOVA_CIDADE_SETOR2/
│ ├── Planilha.docx
│ └── Relatorio.docx
└── xml_nova_cidade/


3. Modificar interface no Streamlit (`app.py`):

Importar nova palavra-chave
from extract_xml_jinja_V6 import PALAVRAS_CHAVE_NOVA_CIDADE
Adicionar novo botão
with col3: # Criar nova coluna se necessário
nova_cidade_button = st.button("NOVA_CIDADE")
Adicionar nova condição
if nova_cidade_button and uploaded_files:
processar_xmls("NOVA_CIDADE", uploaded_files, PALAVRAS_CHAVE_NOVA_CIDADE)




### 2. Adicionar Novos Templates

Para adicionar novos templates para uma cidade existente:

1. Adicionar nova entrada no dicionário de palavras-chave correspondente:

PALAVRAS_CHAVE_CIDADE = {
# Templates existentes...
"CIDADE_NOVO_SETOR": ["CIDADE", "NOVA_PALAVRA_CHAVE"],
}


2. Criar nova pasta de template:

CIDADE/
└── CIDADE_NOVO_SETOR/
├── Planilha.docx
└── Relatorio.docx


### 3. Estrutura de Arquivos Necessária

- Cada template deve ter dois arquivos:
  - `Planilha.docx`
  - `Relatorio.docx`
- Os nomes dos arquivos são fixos e case-sensitive
- Os templates devem estar na pasta correspondente ao seu setor

### 4. Regras para Palavras-chave

- Use sempre letras maiúsculas nas chaves do dicionário
- Primeira palavra-chave geralmente é o nome da cidade
- Segunda palavra-chave é específica do setor
- Todas as palavras-chave devem estar presentes no XML para match

### 5. Estrutura de Diretórios

PROJECT_extractXML/
├── CIDADE1/
│ ├── CIDADE1_SETOR1/
│ └── CIDADE1_SETOR2/
├── CIDADE2/
│ ├── CIDADE2_SETOR1/
│ └── CIDADE2_SETOR2/
├── xml_cidade1/
├── xml_cidade2/
└── jinja/
├── app.py
└── extract_xml_jinja_V6.py


### 6. Observações Importantes

- Mantenha consistência na nomenclatura
- Verifique permissões de pasta
- Faça backup antes de adicionar novos templates
- Teste com XMLs de exemplo
- Verifique se as palavras-chave são únicas para cada setor



