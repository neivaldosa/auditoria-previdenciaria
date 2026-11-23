<p align="center">
  <img src="Logo.png" alt="Inteligência Previdenciária & Auditoria" width="180">
</p>

# Inteligência Previdenciária & Auditoria

[![Python](https://img.shields.io/badge/Python-3.11-3776AB?logo=python&logoColor=white)](https://www.python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-App-FF4B4B?logo=streamlit&logoColor=white)](https://streamlit.io)
[![License](https://img.shields.io/badge/license-not%20specified-lightgrey)](https://github.com/neivaldosa/auditoria-previdenciaria)

Aplicação web em Python para leitura e auditoria do Extrato de Contribuições do CNIS, com apuração de tempo contributivo, identificação de pendências, simulação de regras especiais e geração de relatórios técnicos em PDF e Excel.

> ⚠️ **Aviso importante:** esta ferramenta fornece estimativas automatizadas e não substitui análise jurídica ou previdenciária individual. Regras, tetos, coeficientes e procedimentos podem mudar ao longo do tempo. Documentos oficiais emitidos pelo INSS devem ser considerados como fonte primária.

---

## Visão geral

A plataforma permite que o usuário envie o PDF completo do CNIS baixado pelo Meu INSS e configure seu perfil previdenciário. Com base nas informações extraídas, a aplicação:

- identifica vínculos empregatícios e recolhimentos;
- calcula o tempo bruto e líquido de contribuição;
- desconta períodos concomitantes;
- sinaliza vínculos sem data de saída, recolhimentos bloqueados e outros indicadores;
- simula conversão de períodos especiais de acordo com o ambiente informado;
- compara cenários da regra comum e da aposentadoria da pessoa com deficiência — PCD, conforme LC 142/2013;
- atualiza remunerações históricas para estimar média salarial e RMI; e
- gera parecer executivo, planilha detalhada e relatório técnico em PDF.

## Funcionalidades

### Auditoria do CNIS

- Extração de nome, data de nascimento e vínculos do PDF.
- Organização de empresas, CNPJ/códigos, tipos de filiação e datas.
- Identificação de períodos concomitantes.
- Detecção de indicadores como recolhimentos bloqueados, remunerações pendentes e baixa de vínculo.
- Geração de plano de ação para inconsistências encontradas.

### Cálculos previdenciários

- Soma de períodos brutos.
- União de intervalos sobrepostos para obtenção do tempo líquido.
- Simulação de aposentadoria especial com base em ambiente operacional e fator de conversão.
- Comparação de requisitos PCD para graus grave, moderado e leve.
- Atualização aproximada das remunerações desde julho de 1994.
- Estimativa da média salarial, coeficiente e renda mensal inicial.

### Documentos gerados

- **PDF executivo:** parecer estruturado e anexos técnicos.
- **Planilha Excel:** seis abas coloridas, com vínculos, tempo líquido, especial, PCD, valores e parecer.
- **Relatório visual:** interface responsiva com destaque dos principais indicadores.

## Arquitetura do projeto

```text
.
├── app.py                    # Interface Streamlit e fluxo principal da aplicação
├── extrator_cnis.py          # Extração e processamento dos dados do CNIS
├── gerador_pdf.py            # Geração do relatório executivo em PDF
├── obrigado                  # Página de confirmação de acesso/pagamento
├── Logo.png                  # Identidade visual utilizada na interface e nos documentos
├── requirements.txt          # Dependências Python
├── config.toml               # Configuração temática geral
├── .streamlit/config.toml    # Tema utilizado pelo Streamlit
└── .devcontainer/
    └── devcontainer.json     # Ambiente de desenvolvimento para GitHub Codespaces/Dev Containers
```

## Stack tecnológica

- **Python 3.11**
- **Streamlit** — interface web interativa
- **Pandas** — organização e manipulação dos dados
- **pdfplumber** — leitura textual dos PDFs do CNIS
- **openpyxl** — geração e personalização de planilhas Excel
- **FPDF2** — criação dos relatórios PDF
- **Pillow** — tratamento da identidade visual

## Instalação e execução

### Pré-requisitos

- Python 3.11 ou versão compatível
- `pip`
- Ambiente virtual recomendado

### 1. Clone o repositório

```bash
git clone https://github.com/neivaldosa/auditoria-previdenciaria.git
cd auditoria-previdenciaria
```

### 2. Crie e ative um ambiente virtual

**Linux/macOS:**

```bash
python3 -m venv .venv
source .venv/bin/activate
```

**Windows — PowerShell:**

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

### 3. Instale as dependências

```bash
pip install --user -r requirements.txt
```

### 4. Execute a aplicação

```bash
streamlit run app.py
```

O aplicativo estará disponível normalmente em:

```text
http://localhost:8501
```

## Uso

1. Acesse a aplicação.
2. Informe o sexo utilizado nas regras simuladas.
3. Selecione o ambiente de trabalho predominante.
4. Indique se existe enquadramento PCD/TEA.
5. Baixe o CNIS completo pelo Meu INSS em formato PDF.
6. Faça o upload do documento na área de análise.
7. Gere o diagnóstico gratuito ou utilize um código de acesso liberado.
8. Consulte os módulos de vínculos, tempo, especialidade, PCD e memória financeira.
9. Baixe o relatório executivo em PDF e/ou a planilha detalhada em Excel.

O CNIS deve conter vínculos, contribuições e remunerações no formato completo utilizado pelo extrator.

## Configuração

### Códigos de acesso

O aplicativo possui uma camada de códigos de acesso configurada diretamente em `app.py`. Os códigos representam modos como prévia gratuita, dossiê completo e passe ilimitado. Essa implementação é adequada para demonstração, mas não substitui um sistema seguro de autenticação ou licenciamento.

Para produção, recomenda-se substituir essa lógica por autenticação servidor-side, verificação de pagamento e armazenamento seguro de assinaturas.

### Integrações opcionais

A aplicação possui referências a:

- Google Tag Manager e Google Ads;
- links de pagamento;
- WhatsApp;
- página de confirmação de pagamento; e
- suporte por e-mail.

Os links e identificadores comerciais devem ser revisados antes do deploy em ambiente próprio.

### Secrets do Streamlit

Se utilizar integrações adicionais ou recursos que dependam de credenciais, configure-as por meio de `.streamlit/secrets.toml` ou pelos secrets da plataforma de hospedagem.

## Proteção de dados e privacidade

- O PDF enviado é salvo temporariamente no diretório da aplicação durante a análise.
- A geração dos resultados ocorre em memória ou em arquivos temporários.
- Os arquivos temporários são removidos ao término do processamento.
- O projeto não inclui base de dados própria de usuários ou CNIS.
- Links e scripts de marketing podem gerar eventos ou dados nos respectivos provedores.

Antes do uso em produção, deve-se realizar avaliação jurídica de LGPD, política de privacidade, termos de uso, segurança da infraestrutura e retenção de logs.

## Desenvolvimento em Codespaces

O repositório inclui configuração de dev container baseada em Python 3.11 com:

- instalação automática das dependências;
- execução automática do Streamlit;
- encaminhamento da porta `8501`; e
- extensões recomendadas para Python e Pylance.

No GitHub Codespaces, abra o repositório e escolha a opção de criação do ambiente de desenvolvimento.

## Estrutura técnica

### `extrator_cnis.py`

Responsável por:

- processar páginas do PDF;
- identificar vínculos e remunerações;
- calcular períodos brutos e líquidos;
- gerar indicadores de auditoria;
- simular aposentadoria especial;
- montar cenários PCD;
- produzir a planilha de resultados;
- retornar o parecer textual.

### `gerador_pdf.py`

Responsável por:

- criar cabeçalho e rodapé profissionais;
- formatar resumo executivo;
- incluir parecer técnico;
- adicionar anexos detalhados;
- inserir logo e informações de contato;
- gerar a saída em PDF A4.

### `app.py`

Responsável por:

- montar a interface Streamlit;
- receber o PDF;
- controlar os perfis e planos;
- apresentar os resultados;
- disponibilizar downloads;
- criar a planilha personalizada; e
- integrar suporte e confirmação de pagamento.

## Limitações conhecidas

- Os cálculos são aproximados e dependem da qualidade do texto extraído do PDF.
- CNIS em layout alterado pode exigir ajustes no extrator.
- Os fatores históricos e índices de atualização podem precisar de revisão periódica.
- A versão atual não possui testes automatizados.
- Não há licença definida no repositório.
- A validação de pagamento/código não deve ser considerada um sistema de segurança robusto sem backend adequado.
- Os resultados não devem ser interpretados como decisão definitiva do INSS.

## Contribuição

Contribuições são bem-vindas.

1. Faça um fork do projeto.
2. Crie uma branch para sua alteração:

```bash
git checkout -b feature/melhoria-auditoria
```

3. Realize commits claros e objetivos.
4. Abra um Pull Request descrevendo o problema, a solução e os testes realizados.

## Licença

Este repositório não possui uma licença definida atualmente. Até que uma licença seja adicionada, os termos de uso, distribuição e modificação não estão formalmente concedidos.

## Contato

- **WhatsApp:** +55 21 96880-7965
- **E-mail:** inteligenciaprevidenciariaaudi@gmail.com
- **Repositório:** [neivaldosa/auditoria-previdenciaria](https://github.com/neivaldosa/auditoria-previdenciaria)

---

<p align="center">
  <strong>Inteligência Previdenciária & Auditoria</strong>
  <br>
  Diagnóstico técnico apoiado por automação — não substitui parecer profissional.
</p>
