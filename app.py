import io
import os
import re
import base64
import tempfile
import urllib.parse
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from extrator_cnis import extrair_cnis_v5
from gerador_pdf import criar_pdf_executivo
from auth_supabase import (
    renderizar_painel_acesso,
    gerar_link_checkout_personalizado,
    registrar_uso_credito_avulso,
    enviar_email_recuperacao,
    atualizar_nova_senha,
    get_supabase,
)


def _achar_logo_oficial():
    for nome in ["Logo.png", "logo.png", "LOGO.PNG", "Logo.jpg", "logo.jpg"]:
        if os.path.exists(nome):
            return nome
    return None


@st.cache_data
def _gerar_bytes_logo_transparente():
    logo_path = _achar_logo_oficial()
    if not logo_path:
        return None
    try:
        from PIL import Image
        with Image.open(logo_path) as img:
            rgba = img.convert("RGBA")
            r0, g0, b0, _ = rgba.getpixel((2, 2))
            novos = [
                (r, g, b, 0) if (abs(r - r0) < 25 and abs(g - g0) < 25 and abs(b - b0) < 25) else (r, g, b, a)
                for r, g, b, a in rgba.getdata()
            ]
            rgba.putdata(novos)
            buf = io.BytesIO()
            rgba.save(buf, format="PNG")
            return buf.getvalue()
    except Exception:
        with open(logo_path, "rb") as f:
            return f.read()


def aplicar_logo_todas_abas_excel(buffer_excel):
    buffer_excel.seek(0)
    wb = load_workbook(buffer_excel)
    bytes_logo_transp = _gerar_bytes_logo_transparente()

    fill_azul = PatternFill(start_color="0F293A", end_color="0F293A", fill_type="solid")
    borda_dourada = Border(bottom=Side(style="medium", color="C59B27"))
    font_titulo = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    font_subtitulo = Font(name="Calibri", size=9.5, bold=True, color="C59B27")
    font_contato = Font(name="Calibri", size=8.5, italic=True, color="E2E8F0")

    mapa_acentuacao = {
        "Parecer_Tecnico_Completo": "Parecer Técnico Completo",
        "Metrica_Financeira": "Métrica Financeira",
        "Valor_Estimado": "Valor Estimado",
        "Nome_Empresa": "Nome da Empresa",
        "CNPJ_Codigo": "CNPJ / Código",
        "Tipo_Filiacao": "Tipo de Filiação",
        "Data_Inicio": "Data de Início",
        "Data_Fim": "Data de Saída",
        "Status_Auditoria": "Status da Auditoria",
        "Plano_Acao_INSS": "Plano de Ação no INSS",
        "Ultima_Remuneracao": "Última Remuneração",
        "Dias_Brutos": "Dias Brutos",
        "Dias_Liquidos": "Dias Líquidos",
        "Periodo_Concomitante": "Período Concomitante",
        "Criterio_Legal": "Critério Legal",
        "Situacao_Atual": "Situação Atual",
        "Tempo_Exigido": "Tempo Exigido",
        "Tempo_Faltante": "Tempo Faltante",
        "Observacao_Tecnica": "Observação Técnica",
    }

    for idx_aba, ws in enumerate(wb.worksheets):
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                if isinstance(cell.value, str):
                    val = cell.value.strip()
                    if val in mapa_acentuacao:
                        cell.value = mapa_acentuacao[val]
                    elif "_" in val and not val.startswith("http"):
                        cell.value = val.replace("_", " ")

        ws.insert_rows(1, amount=3)
        ws.row_dimensions[1].height = 20
        ws.row_dimensions[2].height = 17
        ws.row_dimensions[3].height = 16
        if idx_aba == 0:
            ws.row_dimensions[4].height = 34
        if (ws.column_dimensions["A"].width or 0) < 12:
            ws.column_dimensions["A"].width = 12

        max_col = max(ws.max_column, 8)
        for r in range(1, 4):
            for c in range(1, max_col + 1):
                cel = ws.cell(row=r, column=c)
                cel.fill = fill_azul
                if r == 3:
                    cel.border = borda_dourada

        ws["B1"] = "INTELIGÊNCIA PREVIDENCIÁRIA & AUDITORIA"
        ws["B1"].font = font_titulo
        ws["B1"].alignment = Alignment(vertical="center", horizontal="left")

        ws["B2"] = f"DOSSIÊ TÉCNICO DE AUDITORIA — ABA: {ws.title.upper()}"
        ws["B2"].font = font_subtitulo
        ws["B2"].alignment = Alignment(vertical="center", horizontal="left")

        ws["B3"] = "WhatsApp: +55 21 96880-7955  |  E-mail: inteligenciaprevidenciariaaudi@gmail.com"
        ws["B3"].font = font_contato
        ws["B3"].alignment = Alignment(vertical="center", horizontal="left")

        if bytes_logo_transp:
            img = XLImage(io.BytesIO(bytes_logo_transp))
            img.width = 58
            img.height = 58
            ws.add_image(img, "A1")

        ws.freeze_panes = "A5"

        if idx_aba == 0:
            linha_ini = ws.max_row + 3
            linha_fim = linha_ini + 2
            ws.merge_cells(start_row=linha_ini, start_column=3, end_row=linha_fim, end_column=6)

            cel_dica = ws.cell(row=linha_ini, column=3)
            cel_dica.value = (
                "💡 DICA DE NAVEGAÇÃO: Esta planilha possui 6 ABAS COLORIDAS na barra inferior "
                "(Tempo Líquido, Especial, PCD, Valor RMI e Parecer Final). Clique nas guias lá embaixo para ver seu dossiê completo!"
            )
            cel_dica.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
            cel_dica.font = Font(name="Calibri", size=11, bold=True, color="92400E")
            cel_dica.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            borda_ouro = Side(style="medium", color="C59B27")
            for r_dica in range(linha_ini, linha_fim + 1):
                ws.row_dimensions[r_dica].height = 20
                for c_dica in range(3, 7):
                    ws.cell(row=r_dica, column=c_dica).border = Border(
                        left=borda_ouro if c_dica == 3 else None,
                        right=borda_ouro if c_dica == 6 else None,
                        top
