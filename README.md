# Inteligência Previdenciária & Auditoria

Aplicação web em Python para leitura, auditoria e simulação previdenciária a partir do Extrato de Contribuições (CNIS), com apuração de tempo contributivo, diagnóstico de pendências, simulação de regras especiais e geração de relatórios técnicos em PDF e Excel.

> ⚠️ **Aviso importante:** Esta ferramenta fornece estimativas automatizadas para apoio técnico e não substitui análise jurídica ou previdenciária individualizada. Regras, tetos, coeficientes e procedimentos do INSS sofrem alterações periódicas. Documentos oficiais emitidos pelo INSS devem ser sempre considerados como fonte primária.

---

## Visão Geral

A plataforma processa o PDF completo do CNIS emitido pelo Meu INSS, cruza os dados com o perfil do segurado e entrega um diagnóstico estruturado:

- **Mapeamento de vínculos:** extração de empregadores, CNPJs, tipos de filiação, datas de início/fim e remunerações.
- **Apuração de tempo:** cálculo de tempo bruto e tempo líquido com união automática de períodos concomitantes.
- **Auditoria de pendências:** detecção de vínculos abertos (sem data de saída), recolhimentos bloqueados, indicadores de extemporaneidade e lacunas contributivas.
- **Simulação de regras:** conversão de tempo especial conforme o ambiente operacional e comparativo de requisitos para aposentadoria da Pessoa com Deficiência (PCD — LC 142/2013).
- **Cálculo financeiro:** atualização monetária das remunerações desde julho/1994, estimativa de média salarial, coeficiente e Renda Mensal Inicial (RMI).
- **Exportação:** geração instantânea de parecer executivo em PDF e planilha analítica em Excel.

---

## Entregáveis Gerados

1. **Relatório Visual Interativo:** painel na própria interface com indicadores-chave, alertas e planos de ação.
2. **Dossiê Executivo (PDF):** documento formatado em A4 com resumo técnico, parecer textual e anexos detalhados.
3. **Planilha Analítica (Excel):** arquivo `.xlsx` organizado em seis abas temáticas (Vínculos, Tempo Líquido, Especial, PCD, Memória de Cálculo e Parecer).

---

## Stack Tecnológica

- **Linguagem:** Python 3.11
- **Interface Web:** Streamlit
- **Manipulação de Dados:** Pandas
- **Extração de PDF:** pdfplumber
- **Geração de Planilhas:** openpyxl
- **Geração de PDF:** FPDF2
- **Processamento de Imagem:** Pillow

---

## Arquitetura do Projeto

```text
.
├── app.py                    # Interface Streamlit, controle de fluxo e visualização
├── extrator_cnis.py          # Parser do PDF, regras previdenciárias e motor de cálculo
├── gerador_pdf.py            # Renderização e diagramação do relatório executivo em PDF
├── obrigado                  # Página de confirmação pós-checkout
├── Logo.png                  # Identidade visual da aplicação e dos relatórios
├── requirements.txt          # Dependências do projeto
├── .gitignore                # Exclusão de arquivos locais e sensíveis do controle de versão
├── .streamlit/
│   └── config.toml           # Configurações de tema e servidor do Streamlit
└── .devcontainer/
    └── devcontainer.json     # Configuração de ambiente para GitHub Codespaces
