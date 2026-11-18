# Auditoria Previdenciária

Aplicação web em Streamlit para analisar o CNIS oficial do INSS, auditar vínculos e contribuições e estimar diferentes cenários de aposentadoria no Brasil.

A ferramenta processa um PDF de **“Vínculos, Contribuições e Remunerações”** e gera:

- Auditoria dos vínculos extraídos
- Cálculo do tempo líquido de contribuição
- Identificação de possíveis pendências
- Análise de aposentadoria especial
- Comparação com a aposentadoria PCD pela Lei Complementar nº 142/2013
- Estimativa da Renda Mensal Inicial
- Relatórios em PDF e Excel

## Funcionalidades

- Upload do extrato CNIS em PDF
- Extração de dados pessoais, vínculos e competências
- Remoção de períodos concomitantes
- Auditoria de vínculos ausentes ou pendências
- Identificação de indícios de atividade especial
- Análise de enquadramento PCD/TEA
- Comparação entre regras comuns e PCD
- Projeção financeira baseada em tetos e fatores do INSS
- Exportação dos resultados em Excel
- Geração de dossiê executivo em PDF
- Interface desenvolvida para uso em computadores e dispositivos móveis

## Tecnologias

- Python
- Streamlit
- Pandas
- pdfplumber
- openpyxl
- FPDF2

## Estrutura do projeto

```text
app.py                 Interface Streamlit e fluxo da aplicação
extrator_cnis.py       Extração e análise dos dados do PDF
gerador_pdf.py         Geração do relatório executivo em PDF
requirements.txt       Dependências Python
config.toml            Tema da aplicação
.streamlit/config.toml Configuração específica do Streamlit
```

## Instalação

### Pré-requisitos

- Python 3.10 ou superior
- Gerenciador de pacotes `pip`

### Executar localmente

```bash
git clone https://github.com/pedrohd259/auditoria-previdenciaria.git
cd auditoria-previdenciaria

python -m venv .venv
source .venv/bin/activate
# No Windows:
# .venv\Scripts\activate

pip install -r requirements.txt
streamlit run app.py
```

A aplicação estará disponível em:

```text
http://localhost:8501
```

## Como utilizar

1. Acesse o portal ou aplicativo **Meu INSS**.
2. Entre em **Extratos e Comprovantes**.
3. Selecione **Extrato de Contribuições (CNIS)**.
4. Baixe o documento na opção:
   - **Vínculos, Contribuições e Remunerações**
5. Abra a aplicação.
6. Informe o sexo e o ambiente de trabalho.
7. Indique quando aplicável o enquadramento PCD/TEA.
8. Envie o arquivo PDF.
9. Execute a auditoria.
10. Consulte os resultados ou faça o download dos relatórios.

## Arquivos gerados

A aplicação pode gerar:

- `Dossie_Auditoria_INSS.xlsx`
- `Relatorio_Executivo_Previdenciario.pdf`

Os arquivos temporários enviados e gerados durante a análise são removidos automaticamente ao concluir o processamento.

## Privacidade

O processamento é realizado em arquivos temporários na própria execução da aplicação. Os arquivos enviados não devem ser mantidos permanentemente após a auditoria.

Evite compartilhar arquivos públicos que contenham:

- CPF
- nome completo
- endereço
- dados bancários
- informações médicas
- números completos de documentos

## Limitações importantes

Os cálculos são estimativas baseadas nas informações disponíveis no CNIS e nas regras implementadas no código. A aplicação não substitui:

- consulta a advogado previdenciário;
- análise técnica de contador especializado;
- decisão administrativa ou judicial do INSS;
- diagnóstico médico para fins de enquadramento PCD.

Especialmente, a conversão de tempo especial depende da comprovação por PPP, LTCAT e demais documentos exigidos. Contribuições em períodos concomitantes não devem ser somadas automaticamente.

## Licença

Nenhuma licença foi declarada neste repositório. Entre em contato com o autor antes de redistribuir ou utilizar comercialmente o projeto.

## Autor

**pedrohd259**  
GitHub: https://github.com/pedrohd259
