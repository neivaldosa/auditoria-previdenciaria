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
```

---

## Instalação e Execução Local

### Pré-requisitos
- Python 3.11 ou superior
- `pip` e ambiente virtual (`venv`)

### Passo a passo

1. **Clone o repositório:**
   Baixe uma cópia completa dos arquivos do projeto do GitHub para o seu computador e entre na pasta raiz criada.
```bash
git clone https://github.com/neivaldosa/auditoria-previdenciaria.git
cd auditoria-previdenciaria
```

2. **Crie e ative o ambiente virtual:**
   Isole as bibliotecas deste projeto para evitar conflitos com outros programas Python instalados na sua máquina. O terminal exibirá `(.venv)` antes do caminho quando estiver ativo.

   *Linux / macOS:*
```bash
python3 -m venv .venv
source .venv/bin/activate
```
   *Windows (PowerShell):*
```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

3. **Instale as dependências:**
   Leia o arquivo `requirements.txt` e baixe automaticamente todas as bibliotecas necessárias (Streamlit, Pandas, pdfplumber, openpyxl, FPDF2 e Pillow).
```bash
pip install -r requirements.txt
```

4. **Configure as variáveis de ambiente e segredos:**
   Crie o arquivo `.streamlit/secrets.toml` localmente (ou configure na aba **Settings > Secrets** do Streamlit Cloud) com os códigos de acesso e chaves de API antes de iniciar a aplicação:
```toml
[codigos_acesso]
dossie_avulso = ["SEU_CODIGO_DOSSIE"]
passe_ilimitado = ["SEU_CODIGO_VIP"]
desconto_50 = ["SEU_CUPOM_50"]
brinde_gratis = ["SEU_CUPOM_CORTESIA"]
```

5. **Execute a aplicação:**
   Inicie o servidor local do Streamlit; o sistema abrirá automaticamente uma aba no seu navegador padrão no endereço `http://localhost:8501`.
```bash
streamlit run app.py
```

---

## Fluxo de Uso

1. Acesse a plataforma e preencha os parâmetros do segurado (sexo, ambiente de trabalho predominante e eventual enquadramento PCD/TEA).
2. Faça o upload do **Extrato CNIS completo** (formato PDF oficial do Meu INSS contendo vínculos e remunerações).
3. Visualize a prévia diagnóstica ou insira sua chave de acesso para liberar a auditoria completa.
4. Navegue pelas abas de vínculos, tempo de contribuição, especialidade, PCD e memória financeira.
5. Baixe o parecer técnico em **PDF** e a planilha detalhada em **Excel**.

---

## Privacidade e Tratamento de Dados (LGPD)

- **Processamento 100% em memória:** o PDF do CNIS enviado e a planilha gerada são processados diretamente em memória (`io.BytesIO`), sem gravação de documentos dos segurados em disco.
- **Sem armazenamento permanente:** a aplicação não mantém banco de dados próprio de segurados ou históricos de CNIS.
- **Deploy em produção:** recomenda-se hospedagem com HTTPS ativo, gestão de segredos via `secrets.toml` e adequação aos Termos de Uso e Política de Privacidade vigentes.

---

## Licença

**Proprietário — Todos os direitos reservados.**  
O código-fonte, a lógica de extração e os modelos de relatórios deste repositório são de uso exclusivo do autor. Não é permitida a cópia, redistribuição ou uso comercial sem autorização expressa.

---

## Contato e Suporte

- **E-mail:** inteligenciaprevidenciariaaudi@gmail.com
- **WhatsApp:** +55 21 96880-7955
