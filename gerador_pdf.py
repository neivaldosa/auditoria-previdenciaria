from datetime import datetime
from fpdf import FPDF


def limpar_texto_pdf(texto):
    return str(texto).encode("latin-1", "ignore").decode("latin-1")


class RelatorioPrevidenciarioPDF(FPDF):
    def header(self):
        self.set_fill_color(15, 41, 58)
        self.rect(0, 0, 210, 32, "F")
        self.set_fill_color(197, 155, 39)
        self.rect(0, 32, 210, 2, "F")
        self.set_y(8)
        self.set_font("Helvetica", "B", 15)
        self.set_text_color(255, 255, 255)
        self.cell(0, 7, "INTELIGÊNCIA PREVIDENCIÁRIA & AUDITORIA", ln=True)
        self.set_font("Helvetica", "", 9)
        self.set_text_color(243, 208, 107)
        self.cell(0, 5, "DOSSIÊ TÉCNICO DE CNIS, ENQUADRAMENTO PCD (LC 142) E PROJEÇÃO DE RENDA", ln=True)
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(120, 130, 140)
        data_emissao = datetime.now().strftime("%d/%m/%Y %H:%M")
        self.cell(0, 10, f"Gerado em {data_emissao} | Confidencial - Protegido pela LGPD | Página {self.page_no()}", align="C")


def criar_pdf_executivo(caminho_pdf_saida, df_resumo, df_financeiro, parecer_texto):
    pdf = RelatorioPrevidenciarioPDF()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()
    pdf.set_y(40)

    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(15, 41, 58)
    pdf.cell(0, 8, "1. RESUMO EXECUTIVO DO SEGURADO", ln=True)
    pdf.set_draw_color(197, 155, 39)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)

    pdf.set_text_color(30, 41, 59)
    for _, row in df_resumo.iterrows():
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.cell(75, 6.5, f"{limpar_texto_pdf(row['Indicador'])}:")
        pdf.set_font("Helvetica", "", 9.5)
        pdf.cell(0, 6.5, limpar_texto_pdf(row["Valor"]), ln=True)

    pdf.ln(5)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(15, 41, 58)
    pdf.cell(0, 8, "2. PROJEÇÃO FINANCEIRA E COMPARATIVO DE REGRAS (BASE 2026)", ln=True)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(3)

    pdf.set_fill_color(248, 250, 252)
    for _, row in df_financeiro.iterrows():
        met = limpar_texto_pdf(row["Metrica_Financeira"])
        val = limpar_texto_pdf(row["Valor_Estimado"])
        destaque = "Diferença Mensal" in met or "Aposentadoria PCD" in met
        pdf.set_font("Helvetica", "B" if destaque else "", 9.5)
        pdf.set_text_color(15, 41, 58) if destaque else pdf.set_text_color(51, 65, 85)
        pdf.cell(125, 7, f"  {met}", border=1, fill=destaque)
        pdf.cell(65, 7, f"  {val}", border=1, ln=True, fill=destaque)

    pdf.ln(6)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(15, 41, 58)
    pdf.cell(0, 8, "3. PARECER TÉCNICO E PLANO DE AÇÃO NO INSS", ln=True)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(4)

    pdf.set_font("Helvetica", "", 9.5)
    pdf.set_text_color(30, 41, 59)
    pdf.multi_cell(0, 5.5, limpar_texto_pdf(parecer_texto))

    pdf.output(caminho_pdf_saida)

