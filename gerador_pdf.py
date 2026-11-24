import os
import re
import tempfile
from datetime import datetime
from fpdf import FPDF
from fpdf.fonts import FontFace

_CACHE_LOGO_TRANSP = None


def _achar_logo():
    base = os.path.dirname(os.path.abspath(__file__))
    for nome in ["Logo.png", "logo.png", "LOGO.PNG", "Logo.jpg", "logo.jpg"]:
        if os.path.exists(nome):
            return nome
        p = os.path.join(base, nome)
        if os.path.exists(p):
            return p
    return None


def _obter_logo_transparente_cache(logo_path):
    global _CACHE_LOGO_TRANSP
    if _CACHE_LOGO_TRANSP and os.path.exists(_CACHE_LOGO_TRANSP):
        return _CACHE_LOGO_TRANSP
    if not logo_path:
        return None
    try:
        from PIL import Image

        with Image.open(logo_path) as img:
            rgba = img.convert("RGBA")
            r0, g0, b0, _ = rgba.getpixel((2, 2))
            novos_pixels = [
                (r, g, b, 0) if (abs(r - r0) < 25 and abs(g - g0) < 25 and abs(b - b0) < 25) else (r, g, b, a)
                for r, g, b, a in rgba.getdata()
            ]
            rgba.putdata(novos_pixels)
            tmp = os.path.join(tempfile.gettempdir(), "logo_header_transp_cache.png")
            rgba.save(tmp, "PNG")
            _CACHE_LOGO_TRANSP = tmp
            return tmp
    except Exception:
        return logo_path


class ParecerPrevidenciarioPDF(FPDF):
    def __init__(self):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.logo_path = _achar_logo()
        self.logo_transparente_path = _obter_logo_transparente_cache(self.logo_path)
        self.set_auto_page_break(auto=True, margin=22)
        self.azul = (15, 41, 58)
        self.dourado = (197, 155, 39)
        self.texto = (50, 58, 69)
        self.cinza = (244, 246, 249)
        self.borda = (218, 223, 230)


    def _logo_vetorial(self, x, y, tam, fundo=False):
        if fundo:
            self.set_fill_color(*self.azul)
            self.rect(x, y, tam, tam, style="F")
        cx, cy, s = x + tam / 2.0, y + tam / 2.0, tam / 24.0
        self.set_draw_color(*self.dourado)
        self.set_line_width(max(0.35, tam * 0.045))
        pts = [
            (cx - 7.5 * s, cy - 8.2 * s),
            (cx + 7.5 * s, cy - 8.2 * s),
            (cx + 7.5 * s, cy - 0.5 * s),
            (cx + 5.2 * s, cy + 4.8 * s),
            (cx, cy + 8.5 * s),
            (cx - 5.2 * s, cy + 4.8 * s),
            (cx - 7.5 * s, cy - 0.5 * s),
        ]
        for i in range(len(pts)):
            a, b = pts[i], pts[(i + 1) % len(pts)]
            self.line(a, a, b, b)
        self.set_line_width(max(0.3, tam * 0.035))
        self.line(cx, cy - 5.0 * s, cx, cy + 1.5 * s)
        self.line(cx - 4.4 * s, cy - 3.8 * s, cx + 4.4 * s, cy - 3.8 * s)
        self.line(cx - 2.4 * s, cy + 1.5 * s, cx + 2.4 * s, cy + 1.5 * s)
        self.line(cx - 4.2 * s, cy + 4.2 * s, cx - 1.2 * s, cy + 2.6 * s)
        self.line(cx - 1.2 * s, cy + 2.6 * s, cx + 1.2 * s, cy + 3.4 * s)
        self.line(cx + 1.2 * s, cy + 3.4 * s, cx + 4.6 * s, cy + 0.8 * s)

    def header(self):
        self.set_fill_color(*self.azul)
        self.rect(0, 0, 210, 29, style="F")
        self.set_fill_color(*self.dourado)
        self.rect(0, 29, 210, 1.4, style="F")

        self.set_xy(15, 7)
        self.set_font("Helvetica", "B", 13.5)
        self.set_text_color(255, 255, 255)
        self.cell(150, 6.5, "INTELIGÊNCIA PREVIDENCIÁRIA & AUDITORIA", ln=True)

        self.set_x(15)
        self.set_font("Helvetica", "", 8.2)
        self.set_text_color(*self.dourado)
        self.cell(
            150,
            5,
            "DOSSIÊ TÉCNICO DE CNIS, ENQUADRAMENTO PCD (LC 142) E PROJEÇÃO DE RENDA",
            ln=True,
        )

        if self.logo_transparente_path and os.path.exists(self.logo_transparente_path):
            self.image(self.logo_transparente_path, x=174, y=3.5, w=21, h=21)
        else:
            self._logo_vetorial(174, 3.5, 21, fundo=False)
        self.set_y(35)

    def footer(self):
        self.set_y(-20)
        self.set_draw_color(*self.dourado)
        self.set_line_width(0.4)
        self.line(15, self.get_y(), 195, self.get_y())

        y0 = self.get_y() + 2
        if self.logo_path:
            self.image(self.logo_path, x=15, y=y0, w=11.5, h=11.5)
        else:
            self._logo_vetorial(15, y0, 11.5, fundo=True)

        self.set_xy(29, y0 + 0.3)
        self.set_font("Helvetica", "B", 8)
        self.set_text_color(*self.azul)
        self.cell(0, 3.8, "Inteligência Previdenciária & Auditoria", ln=True)

        self.set_x(29)
        self.set_font("Helvetica", "", 7.3)
        self.set_text_color(90, 98, 110)
        self.cell(0, 3.8, "WhatsApp: +55 21 96880-7955  |  E-mail: inteligenciaprevidenciariaaudi@gmail.com", ln=True)
        
        self.set_x(29)
        self.set_font("Helvetica", "I", 6.8)
        self.set_text_color(130, 138, 150)
        dt = datetime.now().strftime("%d/%m/%Y %H:%M")
        self.cell(125, 3.8, f"Gerado em {dt}  |  Documento Confidencial - Protegido pela LGPD")
        self.cell(0, 3.8, f"Página {self.page_no()} de {{nb}}", align="R")

    def secao(self, titulo):
        self.ln(2)
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(*self.azul)
        self.cell(0, 6, titulo, ln=True)
        self.set_draw_color(*self.dourado)
        self.set_line_width(0.3)
        self.line(15, self.get_y(), 195, self.get_y())
        self.ln(2)

    def linha(self, rotulo, valor, destaque=False, cor_val=None):
        self.set_fill_color(*(235, 241, 246) if destaque else self.cinza)
        self.set_draw_color(*self.borda)
        self.set_line_width(0.2)
        self.set_font("Helvetica", "B", 8.5 if destaque else 8.2)
        self.set_text_color(*self.texto)
        self.cell(105, 6.5, f"  {rotulo}", border=1, fill=True)

        self.set_text_color(*(cor_val if cor_val else (self.azul if destaque else self.texto)))
        self.set_font("Helvetica", "B" if (destaque or cor_val) else "", 8.5 if destaque else 8.2)
        self.cell(75, 6.5, f"  {valor}", border=1, fill=True, ln=True)

    def card(self, badge, titulo, corpo, cor_borda, cor_fundo):
        x0, y0, larg = 15, self.get_y(), 180
        self.set_font("Helvetica", "", 8.0)
        linhas = self.multi_cell(larg - 10, 4.2, corpo, split_only=True)
        alt = 10 + len(linhas) * 4.2
        if y0 + alt > 272:
            self.add_page()
            y0 = self.get_y()
        self.set_fill_color(*cor_fundo)
        self.set_draw_color(*self.borda)
        self.rect(x0, y0, larg, alt, style="DF")
        self.set_fill_color(*cor_borda)
        self.rect(x0, y0, 3.2, alt, style="F")

        self.set_xy(x0 + 6.5, y0 + 2.0)
        self.set_font("Helvetica", "B", 8.0)
        self.set_text_color(*cor_borda)
        self.cell(0, 4.5, f"[{badge}]  {titulo}", ln=True)

        self.set_xy(x0 + 6.5, y0 + 6.8)
        self.set_font("Helvetica", "", 8.0)
        self.set_text_color(*self.texto)
        self.multi_cell(larg - 10.5, 4.2, corpo)
        self.set_y(y0 + alt + 2.5)
    def adicionar_tabela_detalhada(self, titulo_secao, df):
        if df is None or df.empty:
            return
        self.add_page(orientation="L")
        self.secao(titulo_secao)

        df_limpo = df.fillna("-").astype(str)
        colunas = [str(c).replace("_", " ") for c in df_limpo.columns]
        num_cols = len(colunas)
        larg_util = self.w - 30

        pesos = []
        for idx_c in range(num_cols):
            tam_cab = len(colunas[idx_c])
            tam_dados = int(df_limpo.iloc[:, idx_c].str.len().quantile(0.85)) if len(df_limpo) > 0 else 10
            pesos.append(min(max(tam_cab, tam_dados, 8), 55))
        soma_pesos = sum(pesos) or 1
        larguras = tuple((p / soma_pesos) * larg_util for p in pesos)

        estilo_cab = FontFace(emphasis="BOLD", color=(243, 208, 107), fill_color=self.azul, size_pt=7.5)
        estilo_zebra = FontFace(fill_color=(250, 251, 253))

        self.set_font("Helvetica", "", 7.2)
        self.set_text_color(*self.texto)
        self.set_draw_color(*self.borda)

        with self.table(
            borders_layout="ALL",
            cell_fill_color=estilo_zebra,
            cell_fill_mode="ROWS",
            col_widths=larguras,
            headings_style=estilo_cab,
            line_height=4.0,
            padding=1.5,
            text_align="LEFT" if num_cols >= 5 else "CENTER",
            width=larg_util,
        ) as tabela:
            row_h = tabela.row()
            for nome_c in colunas:
                row_h.cell(nome_c)

            for _, linha_df in df_limpo.iterrows():
                row_d = tabela.row()
                for val in linha_df:
                    row_d.cell(val)

def _parse_texto(txt, dados):
    if not txt:
        return
    m1 = re.search(r"Segurado\(a\):\s*([^|\n]+)", txt)
    if m1 and not dados.get("nome"):
        dados["nome"] = m1.group(1).strip()
    m2 = re.search(r"Idade Atual:\s*([^\n]+)", txt)
    if m2 and not dados.get("nascimento_idade"):
        dados["nascimento_idade"] = m2.group(1).strip()

    b1 = re.search(r"1\.\s*TEMPO LÍQUIDO DE CONTRIBUIÇÃO\s*\n(.*?)(?=\n\s*2\.|\Z)", txt, re.S)
    b2 = re.search(r"2\.\s*PENDÊNCIAS NO MEU INSS\s*\n(.*?)(?=\n\s*3\.|\Z)", txt, re.S)
    b3 = re.search(r"3\.\s*ANÁLISE ESPECIAL E PCD.*?\n(.*)", txt, re.S)

    if b1:
        dados["texto_tempo"] = b1.group(1).strip()
        ml = re.search(r"hoje:\s*([^\.]+)", dados["texto_tempo"])
        if ml and not dados.get("tempo_liquido"):
            limpo = ml.group(1).split(", já descontados")[0].strip()
            dados["tempo_liquido"] = limpo
        mc = re.search(r"descontados\s+(\d+\s*dias)", dados["texto_tempo"])
        if mc and not dados.get("concomitantes"):
            dados["concomitantes"] = mc.group(1).strip()
    if b2:
        dados["texto_pendencias"] = b2.group(1).strip()
    if b3:
        dados["texto_especial_pcd"] = b3.group(1).strip()
        mv = re.search(r"estimativa de\s*(R\$\s*[\d\.,]+)\s*para\s*(R\$\s*[\d\.,]+)\s*\(([^)]+)\)", dados["texto_especial_pcd"])
        if mv:
            dados.setdefault("rmi_comum", mv.group(1).strip())
            dados.setdefault("rmi_pcd", mv.group(2).strip())
            dados.setdefault("diferenca_mensal", mv.group(3).strip())


def gerar_relatorio_pdf(dados, caminho_saida="Dossie_Auditoria_Previdenciaria.pdf", abas_completas=None):
    pdf = ParecerPrevidenciarioPDF()
    pdf.alias_nb_pages()
    pdf.set_margins(15, 15, 15)
    pdf.add_page()

    pdf.secao("1. RESUMO EXECUTIVO DO SEGURADO")
    pdf.linha("Segurado:", str(dados.get("nome", "NÃO INFORMADO")))
    pdf.linha("Data de Nascimento / Idade Atual:", str(dados.get("nascimento_idade", "-")))
    pdf.linha("Ambiente de Trabalho Informado:", str(dados.get("ambiente", "Administrativo / Escritório")))
    pdf.linha("Soma Bruta de Todos os Vínculos:", str(dados.get("soma_bruta", "-")))
    pdf.linha("Períodos Concomitantes Descontados:", str(dados.get("concomitantes", "0 dias")))
    pdf.linha("TEMPO LÍQUIDO DE CONTRIBUIÇÃO (COMUM):", str(dados.get("tempo_liquido", "-")), destaque=True)

    pdf.secao("2. PROJEÇÃO FINANCEIRA E COMPARATIVO DE REGRAS (BASE 2026)")
    pdf.linha("Competências Extraídas (Desde 07/1994):", str(dados.get("competencias", "-")))
    pdf.linha("Teto Oficial INSS Vigente (2026):", str(dados.get("teto_inss", "R$ 8.475,55")))
    pdf.linha("Média de 100% dos Salários (Regra Geral):", str(dados.get("media_100", "-")))
    pdf.linha("Média dos 80% Maiores Salários (Regra PCD LC 142):", str(dados.get("media_80_pcd", "-")))
    pdf.linha("Estimativa RMI Regra Comum (Coeficiente 60%):", str(dados.get("rmi_comum", "-")))
    pdf.linha("Estimativa RMI Aposentadoria PCD (100% da Média):", str(dados.get("rmi_pcd", "-")), destaque=True)
    pdf.linha("Diferença Mensal a Favor da Regra PCD:", str(dados.get("diferenca_mensal", "-")), destaque=True, cor_val=(22, 130, 76))

    pdf.secao("3. PARECER TÉCNICO E PLANO DE AÇÃO NO INSS")
    pdf.card("TEMPO COMPROVADO", "1. TEMPO LÍQUIDO DE CONTRIBUIÇÃO APURADO", str(dados.get("texto_tempo", "-")), (15, 41, 58), (243, 246, 250))
    pdf.card("AÇÃO NECESSÁRIA", "2. PENDÊNCIAS E INDICADORES NO MEU INSS", str(dados.get("texto_pendencias", "Nenhuma pendência crítica.")), (185, 45, 45), (253, 245, 245))
    pdf.card("ENQUADRAMENTO ESTRATÉGICO", "3. ANÁLISE ESPECIAL E REGRA PCD (LC 142/2013)", str(dados.get("texto_especial_pcd", "-")), (175, 132, 20), (253, 250, 240))

    if abas_completas and isinstance(abas_completas, dict):
        mapa_secoes = [
            ("1. 📋 CNIS EXTRAÍDO", "ANEXO I — AUDITORIA DETALHADA DE VÍNCULOS E PENDÊNCIAS DO CNIS"),
            ("2. ⏱️ TEMPO LÍQUIDO", "ANEXO II — APURAÇÃO DE TEMPO LÍQUIDO DE CONTRIBUIÇÃO"),
            ("3. ⚠️ APOSENT. ESPECIAL", "ANEXO III — ANÁLISE DE APOSENTADORIA ESPECIAL E GILRAT/PPP"),
            ("4. ♿ REGRA PCD (LC 142)", "ANEXO IV — SIMULAÇÃO DE ENQUADRAMENTO PCD (LC 142/2013)"),
            ("5. 💰 VALOR ESTIMADO (RMI)", "ANEXO V — MEMÓRIA DE CÁLCULO FINANCEIRO (RMI DESDE 07/1994)"),
        ]
        for chave_aba, titulo_anexo in mapa_secoes:
            if chave_aba in abas_completas:
                pdf.adicionar_tabela_detalhada(titulo_anexo, abas_completas[chave_aba])

    pasta = os.path.dirname(caminho_saida)
    if pasta and not os.path.exists(pasta):
        os.makedirs(pasta, exist_ok=True)
    pdf.output(caminho_saida)
    return caminho_saida

def criar_pdf_executivo(*args, **kwargs):
    dados, caminho_saida, textos = {}, None, []
    abas_completas = kwargs.pop("abas_completas", None)

    for arg in args:
        if isinstance(arg, dict):
            if any(hasattr(v, "iterrows") for v in arg.values()):
                abas_completas = arg
            else:
                dados.update(arg)
        elif isinstance(arg, str):
            if arg.lower().endswith(".pdf"):
                caminho_saida = arg
            else:
                textos.append(arg)
        elif hasattr(arg, "iterrows") and hasattr(arg, "columns"):
            cols = list(arg.columns)
            if len(cols) >= 2:
                for _, row in arg.iterrows():
                    chave = str(row[cols[0]]).strip()
                    valor = str(row[cols[1]]).strip()
                    if "Data de Nascimento" in chave or "Idade" in chave:
                        dados["nascimento_idade"] = valor
                    elif "Ambiente" in chave:
                        dados["ambiente"] = valor
                    elif "Soma Bruta" in chave:
                        dados["soma_bruta"] = valor
                    elif "Concomitante" in chave:
                        dados["concomitantes"] = valor
                    elif "TEMPO LÍQUIDO" in chave or "Tempo Líquido" in chave:
                        dados["tempo_liquido"] = valor
                    elif "Competências" in chave or "Competencias" in chave:
                        dados["competencias"] = valor
                    elif "Teto" in chave:
                        dados["teto_inss"] = valor
                    elif "100%" in chave and "Média" in chave or "100%" in chave and "Media" in chave:
                        dados["media_100"] = valor
                    elif "80%" in chave:
                        dados["media_80_pcd"] = valor
                    elif "Regra Comum" in chave:
                        dados["rmi_comum"] = valor
                    elif "Aposentadoria PCD" in chave:
                        dados["rmi_pcd"] = valor
                    elif "Diferença Mensal" in chave or "Diferenca Mensal" in chave:
                        dados["diferenca_mensal"] = valor

    for k, v in kwargs.items():
        if isinstance(v, dict):
            dados.update(v)
        elif isinstance(v, str) and v.lower().endswith(".pdf") and not caminho_saida:
            caminho_saida = v
        elif isinstance(v, str) and k != "caminho_saida":
            dados.setdefault(k, v)

    if textos:
        _parse_texto("\n".join(textos), dados)

    return gerar_relatorio_pdf(
        dados,
        caminho_saida or "Dossie_Auditoria_Previdenciaria.pdf",
        abas_completas=abas_completas,
    )
