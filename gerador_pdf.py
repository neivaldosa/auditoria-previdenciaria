import os
from datetime import datetime
from fpdf import FPDF


class ParecerPrevidenciarioPDF(FPDF):
    def __init__(self, logo_path="logo.png"):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.logo_path = logo_path
        self.set_auto_page_break(auto=True, margin=26)
        self.azul_petroleo = (15, 41, 58)
        self.dourado = (197, 155, 39)
        self.cinza_texto = (50, 58, 69)
        self.cinza_claro = (244, 246, 249)
        self.borda_suave = (218, 223, 230)

    def header(self):
        self.set_fill_color(*self.azul_petroleo)
        self.rect(0, 0, 210, 32, style="F")

        self.set_fill_color(*self.dourado)
        self.rect(0, 32, 210, 1.5, style="F")

        self.set_xy(15, 8)
        self.set_font("Helvetica", "B", 14)
        self.set_text_color(255, 255, 255)
        self.cell(150, 7, "INTELIGÊNCIA PREVIDENCIÁRIA & AUDITORIA", ln=True)

        self.set_x(15)
        self.set_font("Helvetica", "", 8.5)
        self.set_text_color(*self.dourado)
        self.cell(
            150,
            5,
            "DOSSIÊ TÉCNICO DE CNIS, ENQUADRAMENTO PCD (LC 142) E PROJEÇÃO DE RENDA",
            ln=True,
        )

        if os.path.exists(self.logo_path):
            self.image(self.logo_path, x=173, y=4.5, w=22, h=22)

        self.set_y(40)

    def footer(self):
        self.set_y(-22)
        self.set_draw_color(*self.dourado)
        self.set_line_width(0.4)
        self.line(15, self.get_y(), 195, self.get_y())

        y_inicio = self.get_y() + 2
        x_texto = 15

        if os.path.exists(self.logo_path):
            self.image(self.logo_path, x=15, y=y_inicio, w=12, h=12)
            x_texto = 29

        self.set_xy(x_texto, y_inicio + 0.5)
        self.set_font("Helvetica", "B", 8)
        self.set_text_color(*self.azul_petroleo)
        self.cell(0, 4, "Inteligência Previdenciária & Auditoria", ln=True)

        self.set_x(x_texto)
        self.set_font("Helvetica", "", 7.5)
        self.set_text_color(90, 98, 110)
        self.cell(
            0,
            4,
            "WhatsApp: +55 21 96880-7965  |  E-mail: inteligenciaprevidenciariaaudi@gmail.com",
            ln=True,
        )

        self.set_x(x_texto)
        self.set_font("Helvetica", "I", 7)
        self.set_text_color(130, 138, 150)
        data_geracao = datetime.now().strftime("%d/%m/%Y %H:%M")
        self.cell(
            125,
            4,
            f"Gerado em {data_geracao}  |  Documento Confidencial - Protegido pela LGPD",
        )
        self.cell(0, 4, f"Página {self.page_no()} de {{nb}}", align="R")

    def titulo_secao(self, texto):
        self.ln(3)
        self.set_font("Helvetica", "B", 10.5)
        self.set_text_color(*self.azul_petroleo)
        self.cell(0, 7, texto, ln=True)
        self.set_draw_color(*self.dourado)
        self.set_line_width(0.3)
        self.line(15, self.get_y(), 195, self.get_y())
        self.ln(3)

    def linha_tabela(self, rotulo, valor, destaque=False, cor_valor=None):
        altura = 7.5
        largura_rotulo = 105
        largura_valor = 75

        if destaque:
            self.set_fill_color(235, 241, 246)
            self.set_font("Helvetica", "B", 8.8)
        else:
            self.set_fill_color(*self.cinza_claro)
            self.set_font("Helvetica", "B", 8.5)

        self.set_draw_color(*self.borda_suave)
        self.set_line_width(0.2)

        self.set_text_color(*self.cinza_texto)
        self.cell(largura_rotulo, altura, f"  {rotulo}", border=1, fill=True)

        if cor_valor:
            self.set_text_color(*cor_valor)
            self.set_font("Helvetica", "B", 8.8)
        elif destaque:
            self.set_text_color(*self.azul_petroleo)
            self.set_font("Helvetica", "B", 8.8)
        else:
            self.set_text_color(*self.cinza_texto)
            self.set_font("Helvetica", "", 8.5)

        self.cell(largura_valor, altura, f"  {valor}", border=1, fill=True, ln=True)

    def card_parecer(self, badge, titulo, corpo_texto, cor_destaque, cor_fundo):
        largura_total = 180
        x_inicio = 15
        y_inicio = self.get_y()

        self.set_font("Helvetica", "", 8.5)
        linhas = self.multi_cell(largura_total - 10, 4.8, corpo_texto, split_only=True)
        altura_card = 12 + (len(linhas) * 4.8)

        if y_inicio + altura_card > 265:
            self.add_page()
            y_inicio = self.get_y()

        self.set_fill_color(*cor_fundo)
        self.set_draw_color(*self.borda_suave)
        self.set_line_width(0.2)
        self.rect(x_inicio, y_inicio, largura_total, altura_card, style="DF")

        self.set_fill_color(*cor_destaque)
        self.rect(x_inicio, y_inicio, 3.5, altura_card, style="F")

        self.set_xy(x_inicio + 7, y_inicio + 2.5)
        self.set_font("Helvetica", "B", 8.5)
        self.set_text_color(*cor_destaque)
        self.cell(0, 5, f"[{badge}]  {titulo}", ln=True)

        self.set_xy(x_inicio + 7, y_inicio + 8)
        self.set_font("Helvetica", "", 8.5)
        self.set_text_color(*self.cinza_texto)
        self.multi_cell(largura_total - 11, 4.8, corpo_texto)

        self.set_y(y_inicio + altura_card + 3.5)


def gerar_relatorio_pdf(
    dados, caminho_saida="Dossie_Auditoria_Previdenciaria.pdf", logo_path="logo.png"
):
    pdf = ParecerPrevidenciarioPDF(logo_path=logo_path)
    pdf.alias_nb_pages()
    pdf.set_margins(15, 15, 15)
    pdf.add_page()

    pdf.titulo_secao("1. RESUMO EXECUTIVO DO SEGURADO")

    pdf.linha_tabela("Segurado:", dados.get("nome", "NÃO INFORMADO"))
    pdf.linha_tabela(
        "Data de Nascimento / Idade Atual:", dados.get("nascimento_idade", "-")
    )
    pdf.linha_tabela(
        "Ambiente de Trabalho Informado:",
        dados.get("ambiente", "Administrativo / Escritório"),
    )
    pdf.linha_tabela("Soma Bruta de Todos os Vínculos:", dados.get("soma_bruta", "-"))
    pdf.linha_tabela(
        "Períodos Concomitantes Descontados:", dados.get("concomitantes", "0 dias")
    )
    pdf.linha_tabela(
        "TEMPO LÍQUIDO DE CONTRIBUIÇÃO (COMUM):",
        dados.get("tempo_liquido", "-"),
        destaque=True,
    )

    pdf.titulo_secao("2. PROJEÇÃO FINANCEIRA E COMPARATIVO DE REGRAS (BASE 2026)")

    pdf.linha_tabela(
        "Competências Extraídas (Desde 07/1994):", dados.get("competencias", "-")
    )
    pdf.linha_tabela(
        "Teto Oficial INSS Vigente (2026):", dados.get("teto_inss", "R$ 8.475,55")
    )
    pdf.linha_tabela(
        "Média de 100% dos Salários (Regra Geral):", dados.get("media_100", "-")
    )
    pdf.linha_tabela(
        "Média dos 80% Maiores Salários (Regra PCD LC 142):",
        dados.get("media_80_pcd", "-"),
    )
    pdf.linha_tabela(
        "Estimativa RMI Regra Comum (Coeficiente 60%):", dados.get("rmi_comum", "-")
    )
    pdf.linha_tabela(
        "Estimativa RMI Aposentadoria PCD (100% da Média):",
        dados.get("rmi_pcd", "-"),
        destaque=True,
    )
    pdf.linha_tabela(
        "Diferença Mensal a Favor da Regra PCD:",
        dados.get("diferenca_mensal", "-"),
        destaque=True,
        cor_valor=(22, 130, 76),
    )

    pdf.titulo_secao("3. PARECER TÉCNICO E PLANO DE AÇÃO NO INSS")

    pdf.card_parecer(
        badge="TEMPO COMPROVADO",
        titulo="1. TEMPO LÍQUIDO DE CONTRIBUIÇÃO APURADO",
        corpo_texto=dados.get(
            "texto_tempo",
            "Tempo líquido comprovado apurado conforme vínculos válidos no CNIS, já descontados os períodos concomitantes.",
        ),
        cor_destaque=(15, 41, 58),
        cor_fundo=(243, 246, 250),
    )

    pdf.card_parecer(
        badge="AÇÃO NECESSÁRIA",
        titulo="2. PENDÊNCIAS E INDICADORES NO MEU INSS",
        corpo_texto=dados.get(
            "texto_pendencias",
            "Nenhuma pendência crítica impeditiva localizada nos vínculos extraídos.",
        ),
        cor_destaque=(185, 45, 45),
        cor_fundo=(253, 245, 245),
    )

    pdf.card_parecer(
        badge="ENQUADRAMENTO ESTRATÉGICO",
        titulo="3. ANÁLISE ESPECIAL E REGRA PCD (LC 142/2013)",
        corpo_texto=dados.get(
            "texto_especial_pcd",
            "Análise comparativa entre conversão especial (PPP) e redução de tempo/cálculo integral pela LC 142/2013.",
        ),
        cor_destaque=(175, 132, 20),
        cor_fundo=(253, 250, 240),
    )

    pdf.output(caminho_saida)
    return caminho_saida


if __name__ == "__main__":
    dados_exemplo = {
        "nome": "NEIVALDO BARBOSA DE SA FILHO",
        "nascimento_idade": "03/05/1985 (41 anos)",
        "ambiente": "Administrativo / Escritório",
        "soma_bruta": "7400 dias (20 anos, 3 meses e 4 dias)",
        "concomitantes": "8 dias",
        "tempo_liquido": "7392 dias (20 anos, 2 meses e 26 dias)",
        "competencias": "246 meses",
        "teto_inss": "R$ 8.475,55",
        "media_100": "R$ 6.294,66",
        "media_80_pcd": "R$ 7.386,93",
        "rmi_comum": "R$ 3.776,80",
        "rmi_pcd": "R$ 7.386,93",
        "diferenca_mensal": "+ R$ 3.610,13 / mês",
        "texto_tempo": "Tempo líquido comprovado hoje: 20 anos, 2 meses e 26 dias (7392 dias), já descontados 8 dias concomitantes.",
        "texto_pendencias": "Vínculo 5 (RHBRASIL SERVICOS TEMPORARIOS LTDA): ERRO CRÍTICO: Sem Data de Saída. Ação: Solicitar Acerto de Vínculos e Remunerações no Meu INSS e anexar baixa na CTPS ou TRCT.",
        "texto_especial_pcd": (
            "- Conversão Especial (PPP): Perfil Administrativo/Escritório selecionado (sem acréscimo automático de insalubridade/periculosidade).\n"
            "- Regra PCD (LC 142/2013): Direito à redução de tempo (25, 29 ou 33 anos H / 20, 24 ou 28 anos M) sem idade mínima e com média dos 80% maiores salários (100%), elevando a estimativa de R$ 3.776,80 para R$ 7.386,93 (+ R$ 3.610,13/mês)."
        ),
    }

    arquivo = gerar_relatorio_pdf(dados_exemplo)
    print(f"PDF gerado com sucesso: {arquivo}")
