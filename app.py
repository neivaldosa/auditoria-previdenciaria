import os
import uuid
import urllib.parse
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from extrator_cnis import extrair_cnis_v5
from gerador_pdf import criar_pdf_executivo

def _achar_logo_oficial():
    for nome in ["Logo.png", "logo.png", "LOGO.PNG", "Logo.jpg", "logo.jpg"]:
        if os.path.exists(nome):
            return nome
    return None

def aplicar_logo_todas_abas_excel(caminho_excel):
    if not os.path.exists(caminho_excel):
        return
    wb = load_workbook(caminho_excel)
    logo_path = _achar_logo_oficial()
    fill_azul = PatternFill(start_color="0F293A", end_color="0F293A", fill_type="solid")
    borda_dourada = Border(bottom=Side(style="medium", color="C59B27"))
    font_titulo = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    font_subtitulo = Font(name="Calibri", size=9.5, bold=True, color="C59B27")
    font_contato = Font(name="Calibri", size=8.5, italic=True, color="E2E8F0")

    for ws in wb.worksheets:
        ws.insert_rows(1, amount=3)
        ws.row_dimensions[1].height = 20
        ws.row_dimensions[2].height = 17
        ws.row_dimensions[3].height = 16
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

        ws["B3"] = "WhatsApp: +55 21 96880-7965  |  E-mail: inteligenciaprevidenciariaaudi@gmail.com"
        ws["B3"].font = font_contato
        ws["B3"].alignment = Alignment(vertical="center", horizontal="left")

        if logo_path:
            img = XLImage(logo_path)
            img.width = 58
            img.height = 58
            ws.add_image(img, "A1")

        ws.freeze_panes = "A5"

    wb.save(caminho_excel)

LOGO_SITE = _achar_logo_oficial()

st.set_page_config(
    page_title="Inteligência Previdenciária & Auditoria - Análise Detalhada do CNIS",
    page_icon=LOGO_SITE if LOGO_SITE else "⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

LINK_MP_AVULSO_19 = "https://mpago.la"
LINK_MP_DESCONTO_50 = "https://mpago.la"
LINK_MP_PASSE_7D = "https://mpago.la"
LINK_MP_PASSE_30D = "https://mpago.la"
EMAIL_SUPORTE = "inteligenciaprevidenciariaaudi@gmail.com"

CODIGOS_DOSSIE_AVULSO = {"DOSSIE2026", "VIP19"}
CODIGOS_PASSE_ILIMITADO = {"VIP7", "VIP30", "ILIMITADO30", "PRO2026", "ADMIN777"}
CODIGOS_DESCONTO_50 = {"DESC50", "METADE50", "PROMO50"}
CODIGOS_BRINDE_GRATIS = {"BRINDE2026", "CORTESIAVIP", "AMIGO100"}

if "plano_ativo" not in st.session_state:
    st.session_state["plano_ativo"] = "GRATUITO"
if "desconto_50" not in st.session_state:
    st.session_state["desconto_50"] = False

param_acesso = st.query_params.get("acesso", "").strip().upper()
if param_acesso in CODIGOS_PASSE_ILIMITADO:
    st.session_state["plano_ativo"] = "ILIMITADO"
elif param_acesso in CODIGOS_DOSSIE_AVULSO:
    st.session_state["plano_ativo"] = "DOSSIE"
elif param_acesso in CODIGOS_BRINDE_GRATIS:
    st.session_state["plano_ativo"] = "BRINDE"
elif param_acesso in CODIGOS_DESCONTO_50:
    st.session_state["desconto_50"] = True

st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    .block-container {
        padding-top: 1.1rem !important;
        padding-bottom: 2.5rem !important;
        max-width: 96% !important;
    }

    /* Barra lateral fixa em Azul-Petróleo com borda Dourada */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0F293A 0%, #16364D 100%) !important;
        border-right: 2px solid #C59B27 !important;
        min-width: 310px !important;
        max-width: 310px !important;
    }

    /* Remove o botão X de fechar e a seta de recolher a barra lateral */
    [data-testid="stSidebarCollapseButton"],
    [data-testid="collapsedControl"],
    button[kind="headerNoPadding"] {
        display: none !important;
        visibility: hidden !important;
    }

    /* Ajusta a cor dos títulos e textos fora dos cards brancos na sidebar */
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] > div > div > div > div > div > label,
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"] p {
        color: #FFFFFF !important;
    }

    /* Mantém os textos dentro dos cards brancos (.sidebar-card) em azul escuro */
    [data-testid="stSidebar"] .sidebar-card label,
    [data-testid="stSidebar"] .sidebar-card p,
    [data-testid="stSidebar"] .sidebar-card span {
        color: #0F293A !important;
    }
    [data-testid="stSidebar"] .sidebar-card [data-testid="stCaptionContainer"] p {
        color: #475569 !important;
    }
    .sidebar-card {
        background: linear-gradient(180deg, #FFFDF5 0%, #FFFFFF 100%);
        padding: 0.75rem 0.9rem;
        border-radius: 10px;
        border: 1.5px solid rgba(197, 155, 39, 0.55);
        border-left: 5px solid #C59B27;
        margin-bottom: 0.75rem;
        box-shadow: 0 3px 10px rgba(197, 155, 39, 0.12);
    }
    .sidebar-title {
        font-size: 0.85rem;
        font-weight: 700;
        color: #0F293A;
        margin-bottom: 0.3rem;
        text-transform: uppercase;
        letter-spacing: 0.4px;
    }
    .security-badge {
        background: linear-gradient(180deg, #FFFDF5 0%, #FFFFFF 100%);
        border: 1.5px solid rgba(197, 155, 39, 0.55);
        border-left: 5px solid #C59B27;
        color: #0F293A;
        padding: 0.75rem 0.85rem;
        border-radius: 10px;
        font-size: 0.8rem;
        line-height: 1.35;
        margin-top: 0.75rem;
    }
    .hero-box {
        background: linear-gradient(135deg, #0F293A 0%, #1B3B52 100%);
        padding: 1.15rem 1.8rem;
        border-radius: 10px;
        border-bottom: 3px solid #C59B27;
        color: #FFFFFF;
        margin-bottom: 0.8rem;
        box-shadow: 0 6px 18px -5px rgba(15, 41, 58, 0.15);
    }
    .hero-badge {
        display: inline-block;
        background-color: rgba(197, 155, 39, 0.2);
        color: #F3D06B;
        padding: 2px 10px;
        border-radius: 20px;
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.7px;
        text-transform: uppercase;
        margin-bottom: 0.35rem;
        border: 1px solid rgba(197, 155, 39, 0.4);
    }
    .hero-title {
        font-size: 1.45rem;
        font-weight: 700;
        margin: 0 0 0.2rem 0;
        letter-spacing: -0.4px;
    }
    .hero-subtitle {
        font-size: 0.88rem;
        color: #CBD5E1;
        margin: 0;
    }
    [data-testid="stExpander"] {
        background: linear-gradient(90deg, #FFFDF5 0%, #FFFFFF 100%);
        border: 1.5px solid #C59B27;
        border-left: 5px solid #C59B27;
        border-radius: 10px;
        margin-bottom: 0.65rem;
    }
    [data-testid="stExpander"] summary {
        font-weight: 700 !important;
        color: #0F293A !important;
        font-size: 0.94rem !important;
        padding: 0.55rem 0.9rem !important;
    }
    [data-testid="stFileUploader"] section {
        background: linear-gradient(180deg, #FFFDF5 0%, #FFFFFF 100%);
        border: 2px dashed #C59B27;
        border-radius: 10px;
        padding: 0.85rem 1.2rem;
    }
    [data-testid="stFileUploader"] button {
        background-color: #C59B27 !important;
        color: #0F293A !important;
        border: none !important;
        font-weight: 700 !important;
        padding: 0.4rem 1.1rem !important;
        border-radius: 6px !important;
    }
    .pricing-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
        gap: 0.9rem;
        margin: 0.5rem 0 1rem 0;
    }
    .plan-card {
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 10px;
        padding: 0.95rem 1.1rem;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        box-shadow: 0 3px 8px rgba(15, 41, 58, 0.04);
    }
    .plan-card.featured {
        background: linear-gradient(180deg, #0F293A 0%, #16364D 100%);
        border: 2px solid #C59B27;
        color: #FFFFFF;
    }
    .plan-top {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 0.2rem;
    }
    .plan-badge {
        background-color: #C59B27;
        color: #0F293A;
        font-size: 0.68rem;
        font-weight: 800;
        padding: 2px 8px;
        border-radius: 10px;
        text-transform: uppercase;
    }
    .plan-name {
        font-size: 1rem;
        font-weight: 700;
    }
    .plan-price {
        font-size: 1.4rem;
        font-weight: 800;
        margin: 0.15rem 0 0.4rem 0;
    }
    .plan-price span {
        font-size: 0.78rem;
        font-weight: 400;
        opacity: 0.85;
    }
    .plan-list {
        list-style: none;
        padding: 0;
        margin: 0 0 0.7rem 0;
        font-size: 0.8rem;
        line-height: 1.48;
    }
    .plan-cta {
        display: block;
        text-align: center;
        padding: 0.48rem 0.8rem;
        border-radius: 6px;
        font-weight: 700;
        text-decoration: none !important;
        font-size: 0.84rem;
        margin-top: 0.35rem;
    }
    .cta-outline {
        border: 1.5px solid #0F293A;
        color: #0F293A !important;
        background-color: #F8FAFC;
    }
    .cta-gold {
        background-color: #C59B27;
        color: #0F293A !important;
    }
    .cta-gold-outline {
        border: 1.5px solid #C59B27;
        color: #F3D06B !important;
        background-color: rgba(197, 155, 39, 0.12);
    }
    .feature-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(185px, 1fr));
        gap: 0.7rem;
        margin-top: 0.4rem;
        margin-bottom: 1.2rem;
    }
    .feature-card {
        background: linear-gradient(180deg, #FFFFFF 0%, #FFFDF5 100%);
        border: 1px solid #E2E8F0;
        border-top: 3px solid #C59B27;
        border-radius: 8px;
        padding: 0.75rem 0.85rem;
        box-shadow: 0 2px 6px rgba(15, 41, 58, 0.04);
    }
    .feature-title {
        font-size: 0.85rem;
        font-weight: 700;
        color: #0F293A;
        margin-bottom: 0.25rem;
    }
    .feature-desc {
        font-size: 0.76rem;
        color: #475569;
        line-height: 1.35;
        margin: 0;
    }
    .metric-card {
        background-color: #FFFFFF;
        padding: 1.2rem;
        border-radius: 10px;
        border: 1px solid #E2E8F0;
        border-left: 4px solid #0F293A;
    }
    .metric-card.highlight {
        border-left: 5px solid #C59B27;
        background: linear-gradient(to right, #FFFFFF, #FFFDF5);
    }
    .metric-label {
        font-size: 0.8rem;
        color: #64748B;
        text-transform: uppercase;
        font-weight: 600;
        margin-bottom: 0.3rem;
    }
    .metric-value {
        font-size: 1.45rem;
        font-weight: 700;
        color: #0F293A;
    }
    .metric-sub {
        font-size: 0.8rem;
        color: #059669;
        font-weight: 500;
        margin-top: 0.2rem;
    }
    .upsell-banner {
        background: linear-gradient(135deg, #0F293A 0%, #1B3B52 100%);
        border: 2px solid #C59B27;
        border-radius: 10px;
        padding: 1.2rem 1.5rem;
        color: #FFFFFF;
        margin: 1rem 0;
    }
    .stButton > button[kind="primary"] {
        background-color: #0F293A;
        border: 2px solid #C59B27;
        color: #FFFFFF;
        font-weight: 600;
        padding: 0.55rem 1.5rem;
        border-radius: 8px;
        font-size: 0.96rem;
    }
    .floating-container {
        position: fixed;
        bottom: 22px;
        right: 22px;
        display: flex;
        flex-direction: column;
        gap: 10px;
        z-index: 9999;
    }
    .float-btn {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 10px 16px;
        border-radius: 50px;
        font-size: 0.85rem;
        font-weight: 700;
        text-decoration: none !important;
        box-shadow: 0 6px 16px rgba(15, 41, 58, 0.25);
        transition: transform 0.2s ease;
    }
    .float-btn:hover {
        transform: translateY(-2px);
    }
    .float-gemini {
        background: linear-gradient(135deg, #0F293A 0%, #1B3B52 100%);
        color: #F3D06B !important;
        border: 1.5px solid #C59B27;
    }
    .float-support {
        background-color: #C59B27;
        color: #0F293A !important;
    }
</style>
""", unsafe_allow_html=True)

import base64

_logo_b64 = ""
if LOGO_SITE and os.path.exists(LOGO_SITE):
    with open(LOGO_SITE, "rb") as _f_logo:
        _logo_b64 = base64.b64encode(_f_logo.read()).decode("utf-8")

_img_hero_html = (
    f'<img src="data:image/png;base64,{_logo_b64}" style="width:74px; height:74px; border-radius:12px; border:1.5px solid #C59B27; flex-shrink:0;" />'
    if _logo_b64 else ""
)

st.markdown(f"""
<div id="topo-pagina" class="hero-box" style="display:flex; align-items:center; gap:1.2rem;">
    {_img_hero_html}
    <div>
        <div class="hero-badge">Inteligência Previdenciária & Auditoria</div>
        <h1 class="hero-title">Diagnóstico Estratégico de Aposentadoria e CNIS</h1>
        <p class="hero-subtitle">Análise automatizada de vínculos, auditoria de pendências no INSS, enquadramento PCD (LC 142/2013) e projeção de renda.</p>
    </div>
</div>
""", unsafe_allow_html=True)

with st.expander("📌 NÃO POSSUI O ARQUIVO PDF DO CNIS? CLIQUE AQUI PARA VER O PASSO A PASSO OFICIAL (2 MINUTOS)"):
    st.markdown("""
    1. **Acesse o Portal:** Entre em [meu.inss.gov.br](https://inss.gov.br) ou abra o aplicativo **Meu INSS** com sua conta **Gov.br**.
    2. **Busque o Extrato:** Clique em **"Extratos e Comprovantes"** e selecione **"Extrato de Contribuições (CNIS)"**.
    3. **Baixe a Versão Completa:** Role até o final, clique em **"Baixar Documento"** e escolha **"Vínculos, Contribuições e Remunerações"**.
    """)

with st.sidebar:
    _img_side_html = (
        f'<img src="data:image/png;base64,{_logo_b64}" style="width:54px; height:54px; border-radius:10px; border:1.5px solid #C59B27; flex-shrink:0;" />'
        if _logo_b64 else ""
    )
    st.markdown(f"""
    <div style="display:flex; align-items:center; gap:0.75rem; padding-bottom:0.85rem; margin-bottom:0.85rem; border-bottom:1.5px solid rgba(197, 155, 39, 0.45);">
        {_img_side_html}
        <div style="line-height:1.25;">
            <div style="color:#F3D06B; font-size:0.72rem; font-weight:700; text-transform:uppercase; letter-spacing:0.5px;">Auditoria Oficial</div>
            <div style="color:#FFFFFF; font-size:0.95rem; font-weight:800;">Inteligência Previdenciária</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("### ⚙️ Configure seu Perfil")
    sexo = st.selectbox("Sexo biológico (Regra INSS):", options=["M", "F"], format_func=lambda x: "Masculino" if x == "M" else "Feminino")

    st.markdown('<div class="sidebar-card"><div class="sidebar-title">1. Ambiente de Trabalho</div>', unsafe_allow_html=True)
    ambiente = st.radio(
        "Selecione sua atuação predominante:",
        options=["Administrativo / Escritório", "Operacional / Área de Risco (Insalubridade/Periculosidade)"],
        index=0,
        label_visibility="collapsed"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    ambiente_operacional = (ambiente == "Operacional / Área de Risco (Insalubridade/Periculosidade)")

    st.markdown('<div class="sidebar-card"><div class="sidebar-title">2. Enquadramento PCD / TEA</div>', unsafe_allow_html=True)
    possui_pcd = st.checkbox("Possui deficiência / TEA (LC 142/2013)?", value=True)
    st.caption("Marque se possui laudo médico de deficiência física, sensorial, intelectual ou TEA.")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="sidebar-card"><div class="sidebar-title">🔑 Cupom ou Código de Acesso</div>', unsafe_allow_html=True)
    codigo_digitado = st.text_input("Digite seu cupom ou código:", placeholder="Insira seu cupom aqui...", label_visibility="collapsed").strip().upper()
    if st.button("Ativar Código", use_container_width=True):
        if codigo_digitado in CODIGOS_PASSE_ILIMITADO:
            st.session_state["plano_ativo"] = "ILIMITADO"
            st.success("Passe Ilimitado ativado!")
        elif codigo_digitado in CODIGOS_DOSSIE_AVULSO:
            st.session_state["plano_ativo"] = "DOSSIE"
            st.success("Dossiê Completo ativado!")
        elif codigo_digitado in CODIGOS_BRINDE_GRATIS:
            st.session_state["plano_ativo"] = "BRINDE"
            st.success("🎁 Cortesia de 1 Dossiê Completo ativada!")
        elif codigo_digitado in CODIGOS_DESCONTO_50:
            st.session_state["desconto_50"] = True
            st.success("🎉 Cupom de 50% OFF aplicado!")
        else:
            st.error("Código ou cupom inválido.")

    plano_atual = st.session_state["plano_ativo"]
    if plano_atual == "BRINDE":
        badge_plano = "🎁 Cortesia VIP Liberada"
    elif plano_atual == "DOSSIE":
        badge_plano = "🟡 Dossiê Completo Liberado"
    elif plano_atual == "ILIMITADO":
        badge_plano = "👑 Passe Ilimitado Ativo"
    elif st.session_state["desconto_50"]:
        badge_plano = "🏷️ Cupom 50% OFF Ativo (R$ 9,95)"
    else:
        badge_plano = "🟢 Plano Gratuito (Prévia)"
    st.caption(f"Status atual: **{badge_plano}**")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="security-badge">
        <strong>🔒 Privacidade & Segurança LGPD</strong><br>
        Processamento isolado em memória. Nenhum dado pessoal ou PDF fica armazenado após a análise.
    </div>
    """, unsafe_allow_html=True)

st.markdown("#### 📤 Envie seu Extrato Oficial (CNIS)")
arquivo_pdf = st.file_uploader("Selecione ou arraste o arquivo PDF completo baixado do Meu INSS:", type=["pdf"], label_visibility="collapsed")
preco_avulso_txt = "R$ 9,95 (50% OFF)" if st.session_state["desconto_50"] else "R$ 19,90"
link_avulso_atual = LINK_MP_DESCONTO_50 if st.session_state["desconto_50"] else LINK_MP_AVULSO_19

if arquivo_pdf is not None:
    label_btn = "🔍 Gerar Diagnóstico Gratuito (Prévia)" if st.session_state["plano_ativo"] == "GRATUITO" else "🚀 Executar Auditoria e Gerar Dossiê Completo"
    if st.button(label_btn, type="primary"):
        with st.spinner("Processando vínculos, cruzando tetos históricos e auditando indicadores..."):
            sessao_id = str(uuid.uuid4())
            caminho_temp_pdf = f"temp_{sessao_id}.pdf"
            caminho_temp_xlsx = f"dossie_{sessao_id}.xlsx"
            caminho_pdf_relatorio = f"relatorio_{sessao_id}.pdf"

            try:
                with open(caminho_temp_pdf, "wb") as f:
                    f.write(arquivo_pdf.getbuffer())

                parecer_texto = extrair_cnis_v5(
                    caminho_temp_pdf, caminho_temp_xlsx, sexo=sexo,
                    ambiente_operacional=ambiente_operacional, possui_pcd=possui_pcd
                )

                abas_brutas = pd.read_excel(caminho_temp_xlsx, sheet_name=None, skiprows=1)
                aplicar_logo_todas_abas_excel(caminho_temp_xlsx)
                abas = {nome: df.dropna(how="all", axis=1) for nome, df in abas_brutas.items()}
                df_vinculos = abas["1_CNIS_Extraido"]
                df_res = abas["2_Tempo_Liquido_Resumo"]
                df_fin = abas["5_Estimativa_Valor_RMI"]

                tempo_liq = str(df_res.loc[df_res["Indicador"] == "TEMPO LÍQUIDO DE CONTRIBUIÇÃO (COMUM)", "Valor"].tolist().pop())
                rmi_comum = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Regra Comum"), "Valor_Estimado"].tolist().pop())
                rmi_pcd = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Aposentadoria PCD"), "Valor_Estimado"].tolist().pop())
                dif_mensal = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Diferença Mensal"), "Valor_Estimado"].tolist().pop())
                tempo_card = tempo_liq.split("(").pop().replace(")", "") if "(" in tempo_liq else tempo_liq

                modo_pago = st.session_state["plano_ativo"] in ("DOSSIE", "ILIMITADO", "BRINDE")

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f'<div class="metric-card"><div class="metric-label">Tempo Líquido Apurado</div><div class="metric-value">{tempo_card}</div><div class="metric-sub">Períodos concomitantes já descontados</div></div>', unsafe_allow_html=True)
                with col2:
                    val_c2 = rmi_comum if modo_pago else "R$ ••••,••"
                    sub_c2 = "Projeção com coeficiente atual" if modo_pago else "🔒 Disponível no Dossiê Completo"
                    st.markdown(f'<div class="metric-card"><div class="metric-label">Estimativa Regra Geral</div><div class="metric-value">{val_c2}</div><div class="metric-sub">{sub_c2}</div></div>', unsafe_allow_html=True)
                with col3:
                    val_c3 = rmi_pcd if modo_pago else "R$ ••••,••"
                    sub_c3 = f"{dif_mensal} vs. Regra Geral" if modo_pago else "🔒 Cálculo LC 142/2013 bloqueado"
                    st.markdown(f'<div class="metric-card highlight"><div class="metric-label">Potencial Regra PCD (LC 142)</div><div class="metric-value">{val_c3}</div><div class="metric-sub">{sub_c3}</div></div>', unsafe_allow_html=True)

                if not modo_pago:
                    qtd_alertas = df_vinculos["Status_Auditoria"].str.contains("ERRO|ALERTA|ATENÇÃO").sum()
                    parecer_gratuito = (
                        f"DIAGNÓSTICO INICIAL (VERSÃO GRATUITA)\n\n"
                        f"1. TEMPO LÍQUIDO APURADO: {tempo_liq}\n"
                        f"2. AUDITORIA DE ALERTAS NO CNIS: Foram detectados {qtd_alertas} ponto(s) de atenção/pendência nos seus vínculos.\n"
                        f"   -> Os alertas estão listados na aba 'Alertas de Vínculos', mas o passo a passo de como resolver cada pendência no Meu INSS está retido na versão gratuita.\n\n"
                        f"3. CÁLCULOS FINANCEIROS, ESPECIAL E PCD: Bloqueados nesta prévia.\n"
                        f"   -> Desbloqueie o Dossiê Completo por {preco_avulso_txt} (Pix ou Cartão) para liberar automaticamente a solução de todas as pendências, a memória de cálculo desde 1994 e os downloads em PDF e Excel."
                    )

                    st.markdown(f"""
                    <div class="upsell-banner">
                        <h4 style="margin:0 0 0.3rem 0; color:#F3D06B;">⚠️ Encontramos {qtd_alertas} alerta(s) no seu CNIS — Desbloqueie agora e veja como corrigir + valor do benefício</h4>
                        <p style="margin:0 0 0.8rem 0; font-size:0.88rem; color:#E2E8F0;">O <strong>Plano de Ação no INSS (como resolver)</strong>, o cálculo desde 1994 e os arquivos <strong>PDF Executivo + Excel</strong> são liberados automaticamente após o pagamento via Pix ou Cartão no Mercado Pago.</p>
                        <a href="{link_avulso_atual}" target="_blank" class="plan-cta cta-gold" style="display:inline-block; padding:0.55rem 1.3rem;">🔓 Desbloquear Dossiê Completo Agora — {preco_avulso_txt}</a>
                    </div>
                    """, unsafe_allow_html=True)

                    tab0, tab1, tab2 = st.tabs(["📄 Parecer Executivo (Prévia)", "📋 Alertas de Vínculos", "⏱️ Resumo de Tempo"])
                    with tab0:
                        st.text_area("Diagnóstico Resumido:", value=parecer_gratuito, height=230)
                    with tab1:
                        df_vinculos_free = df_vinculos.copy()
                        df_vinculos_free["Plano_Acao_INSS"] = f"🔒 Disponível no Dossiê Completo ({preco_avulso_txt})"
                        st.dataframe(df_vinculos_free, use_container_width=True)
                    with tab2:
                        st.dataframe(df_res, use_container_width=True)
                else:
                    tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
                        "📄 Parecer Executivo", "📋 Auditoria de Vínculos", "⏱️ Resumo de Tempo",
                        "🏭 Especial & GILRAT", "♿ Enquadramento PCD", "💰 Memória Financeira"
                    ])

                    with tab0:
                        st.text_area("Relatório Técnico Estruturado:", value=parecer_texto, height=380)
                    with tab1:
                        st.dataframe(df_vinculos, use_container_width=True)
                    with tab2:
                        st.dataframe(df_res, use_container_width=True)
                    with tab3:
                        st.dataframe(abas["3_Aposentadoria_Especial_PPP"], use_container_width=True)
                    with tab4:
                        st.dataframe(abas["4_Aposentadoria_PCD_LC142"], use_container_width=True)
                    with tab5:
                        st.dataframe(df_fin, use_container_width=True)

                    criar_pdf_executivo(
                        caminho_pdf_relatorio,
                        df_res,
                        df_fin,
                        parecer_texto
                    )

                    with open(caminho_pdf_relatorio, "rb") as f_pdf:
                        bytes_pdf = f_pdf.read()
                    with open(caminho_temp_xlsx, "rb") as f_excel:
                        bytes_excel = f_excel.read()

                    col_dl1, col_dl2, col_dl3 = st.columns(3)
                    with col_dl1:
                        st.download_button(
                            label="📑 Baixar Relatório Executivo (.pdf)",
                            data=bytes_pdf,
                            file_name="Relatorio_Executivo_Previdenciario.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    with col_dl2:
                        st.download_button(
                            label="📊 Baixar Planilha Detalhada (.xlsx)",
                            data=bytes_excel,
                            file_name="Dossie_Auditoria_INSS.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    with col_dl3:
                        msg_whats = urllib.parse.quote(f"Olá! Acabei de gerar minha Auditoria Previdenciária (Tempo Líquido: {tempo_card} | Potencial PCD: {rmi_pcd}) e gostaria de agendar uma análise com o escritório parceiro.")
                        link_whats = f"https://wa.me{msg_whats}"
                        st.link_button("⚖️ Enviar para Escritório Parceiro", url=link_whats, use_container_width=True)

            finally:
                for arq in [caminho_temp_pdf, caminho_temp_xlsx, caminho_pdf_relatorio]:
                    if os.path.exists(arq):
                        os.remove(arq)
st.markdown("""
<div style="font-size: 1.05rem; font-weight: 700; color: #0F293A; margin-top: 0.8rem;">💎 Escolha o Plano Ideal (Liberação Automática via Pix ou Cartão)</div>
<div class="pricing-grid">
    <div class="plan-card">
        <div>
            <div class="plan-top"><span class="plan-name" style="color:#0F293A;">Diagnóstico Inicial</span></div>
            <div class="plan-price" style="color:#0F293A;">Grátis <span>/ prévia</span></div>
            <ul class="plan-list" style="color:#334155;">
                <li>✅ Tempo líquido e leitura do CNIS</li>
                <li>✅ Alerta de pendências nos vínculos</li>
                <li>❌ Como resolver erros, RMI e PDF/Excel</li>
            </ul>
        </div>
        <div style="text-align:center; font-size:0.78rem; color:#64748B; font-weight:600; padding:0.45rem; background:#F1F5F9; border-radius:6px;">Ativo — Envie seu PDF acima</div>
    </div>
    <div class="plan-card">
        <div>
            <div class="plan-top"><span class="plan-name" style="color:#0F293A;">Dossiê Completo Avulso</span></div>
            <div class="plan-price" style="color:#0F293A;">__PRECO_19__ <span>/ único</span></div>
            <ul class="plan-list" style="color:#334155;">
                <li>✅ Plano de ação (como corrigir erros)</li>
                <li>✅ Cálculo RMI (1994–2026), Especial e PCD</li>
                <li>✅ Download PDF Executivo + Planilha Excel</li>
            </ul>
        </div>
        <a href="__LINK_19__" target="_blank" class="plan-cta cta-outline">⚡ Liberar Dossiê (Pix ou Cartão)</a>
    </div>
    <div class="plan-card featured">
        <div>
            <div class="plan-top">
                <span class="plan-name">Passe Ilimitado</span>
                <span class="plan-badge">Mais Escolhido</span>
            </div>
            <div class="plan-price" style="color:#F3D06B;">R$ 39,90 <span>/ 7d ou R$ 69,90 / 30d</span></div>
            <ul class="plan-list" style="color:#E2E8F0;">
                <li>✅ <strong>Simulações e PDFs/Excel ilimitados</strong></li>
                <li>✅ Teste cenários (PCD, Especial, futuro)</li>
                <li>✅ Analise também familiares ou clientes</li>
            </ul>
        </div>
        <div>
            <a href="__LINK_7D__" target="_blank" class="plan-cta cta-gold">⚡ Passe 7 Dias — R$ 39,90 (Pix/Cartão)</a>
            <a href="__LINK_30D__" target="_blank" class="plan-cta cta-gold-outline">👑 Passe 30 Dias — R$ 69,90 (Pix/Cartão)</a>
        </div>
    </div>
</div>

<div style="font-size: 1.05rem; font-weight: 700; color: #0F293A; margin-top: 0.4rem;">📋 O que está incluso no seu Dossiê Completo de Auditoria (6 Módulos):</div>
<div class="feature-grid">
    <div class="feature-card">
        <div class="feature-title">📄 1. Parecer Executivo</div>
        <p class="feature-desc">Diagnóstico jurídico-técnico com riscos, oportunidades e plano de ação no Meu INSS.</p>
    </div>
    <div class="feature-card">
        <div class="feature-title">📋 2. Auditoria de Vínculos</div>
        <p class="feature-desc">Varredura de indicadores ocultos do INSS, faltas de baixa e guias concomitantes.</p>
    </div>
    <div class="feature-card">
        <div class="feature-title">⏱️ 3. Resumo de Tempo</div>
        <p class="feature-desc">Cálculo exato do tempo líquido, descontando períodos sobrepostos automaticamente.</p>
    </div>
    <div class="feature-card">
        <div class="feature-title">🏭 4. Especial (Risco)</div>
        <p class="feature-desc">Identifica períodos operacionais até 2019 (fator 1.4 H / 1.2 M) e orienta PPP/GILRAT.</p>
    </div>
    <div class="feature-card">
        <div class="feature-title">♿ 5. Enquadramento PCD</div>
        <p class="feature-desc">Simula LC 142/2013 (incl. TEA Nível 1) nos graus grave, moderado e leve sem idade mínima.</p>
    </div>
    <div class="feature-card">
        <div class="feature-title">💰 6. Memória Financeira</div>
        <p class="feature-desc">Atualiza salários desde 07/1994 e compara Regra Geral vs. 80% maiores na Regra PCD.</p>
    </div>
</div>
""".replace("__PRECO_19__", preco_avulso_txt).replace("__LINK_19__", link_avulso_atual).replace("__LINK_7D__", LINK_MP_PASSE_7D).replace("__LINK_30D__", LINK_MP_PASSE_30D), unsafe_allow_html=True)

st.markdown('<h3 id="assistente-ia" style="text-align: center; color: #0F293A; margin-top: 1.8rem; margin-bottom: 0.3rem;">🤖 Assistente Virtual Previdenciário (IA Gemini)</h3>', unsafe_allow_html=True)
st.markdown('<p style="text-align: center; color: #64748B; font-size: 0.9rem; margin-bottom: 1rem;">Tem dúvidas sobre regras do INSS, tempo de contribuição, PPP ou aposentadoria PCD (LC 142)? Pergunte abaixo:</p>', unsafe_allow_html=True)

_, col_ia, _ = st.columns([1, 2.2, 1])

with col_ia:
    if "GEMINI_API_KEY" in st.secrets:
        from google import genai
        client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

        if "mensagens_ia" not in st.session_state:
            st.session_state["mensagens_ia"] = [
                {"role": "assistant", "content": "Olá! Sou o assistente virtual da **Inteligência Previdenciária & Auditoria**. Qual é a sua dúvida sobre aposentadoria, CNIS ou regras do INSS?"}
            ]

        for msg in st.session_state["mensagens_ia"]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        if pergunta_usuario := st.chat_input("Digite sua dúvida previdenciária aqui..."):
            st.session_state["mensagens_ia"].append({"role": "user", "content": pergunta_usuario})
            with st.chat_message("user"):
                st.markdown(pergunta_usuario)

            with st.chat_message("assistant"):
                with st.spinner("Consultando legislação previdenciária..."):
                    prompt_sistema = (
                        "Você é um assistente especialista em Direito Previdenciário Brasileiro (INSS, EC 103/2019, "
                        "Aposentadoria Especial e LC 142/2013 PCD/TEA). Responda de forma clara, objetiva, educada e em "
                        "português do Brasil. Sempre que fizer sentido, lembre o usuário de que ele pode enviar o PDF do "
                        "CNIS no topo desta página para auditar o caso exato dele."
                    )
                    resposta = client.models.generate_content(
                        model="gemini-2.5-flash",
                        contents=f"{prompt_sistema}\n\nPergunta do usuário: {pergunta_usuario}"
                    )
                    st.markdown(resposta.text)
                    st.session_state["mensagens_ia"].append({"role": "assistant", "content": resposta.text})
    else:
        st.info("💡 **Quer tirar dúvidas gerais sobre aposentadoria com Inteligência Artificial?** Clique abaixo para conversar com o Gemini ou use o botão dourado no canto inferior direito para suporte técnico do site.")
        st.link_button("✨ Tirar Dúvidas sobre Aposentadoria no Gemini", url="https://google.com", use_container_width=True)

st.markdown('<h3 style="text-align: center; color: #0F293A; margin-top: 2rem; margin-bottom: 1rem;">❓ Perguntas Frequentes</h3>', unsafe_allow_html=True)

_, col_faq, _ = st.columns([1, 2.2, 1])

with col_faq:
    with st.expander("Por que é tão importante entender o seu CNIS e ajustar possíveis problemas antes de entrar com o pedido de aposentadoria?", expanded=True):
        st.markdown("Obter acesso aos detalhes do seu CNIS e realizar uma auditoria prévia antes de dar entrada na aposentadoria é crucial para evitar dor de cabeça, indeferimentos ou perda de tempo de contribuição em um momento em que você quer descansar. Dependendo do caso, os relatórios da **Inteligência Previdenciária & Auditoria** entregam o diagnóstico técnico completo por uma fração do custo de uma consulta particular.")

    with st.expander("Qual é a diferença entre Cálculo Previdenciário e Planejamento Previdenciário?"):
        st.markdown("O **Cálculo Previdenciário** é voltado à apuração do tempo líquido e do valor estimado do benefício a partir dos seus salários desde julho de 1994. Já o **Planejamento Previdenciário** compara regras e cenários distintos (Regra Geral, Aposentadoria Especial por exposição a risco e Aposentadoria PCD pela LC 142/2013) para mostrar qual caminho é mais vantajoso para você.")

    with st.expander("Onde encontro essas funcionalidades?"):
        st.markdown("Basta enviar o PDF do seu CNIS no quadro de envio acima. Você pode escolher o **Dossiê Completo Avulso** (para analisar 1 extrato com downloads em PDF e Excel) ou os **Passes Ilimitados de 7 dias ou 30 dias** para realizar simulações distintas conforme a sua situação.")

    with st.expander("Preciso ter o CNIS para começar?"):
        st.markdown("Sim, esse é o único documento necessário (**CNIS em PDF** — [saiba como emitir gratuitamente no Meu INSS clicando aqui](#topo-pagina)) para entregar uma análise detalhada, explicando ponto a ponto o que você precisa fazer para seguir com o processo de aposentadoria.")

    with st.expander("Posso obter uma demonstração gratuita para entender como o serviço funciona?"):
        st.markdown("Sim! Basta enviar seu PDF normalmente sem código de ativação — o sistema executa o **Diagnóstico Inicial (Grátis)**, calculando seu tempo líquido e apontando se há alertas nos seus vínculos.")

    with st.expander("Tenho diagnóstico de TEA (Autismo) ou outra deficiência. O sistema calcula a regra PCD?"):
        st.markdown("Sim. Ao marcar a opção **\"Enquadramento PCD / TEA (LC 142/2013)\"** na barra lateral, o sistema simula o tempo exigido para deficiência leve (como TEA Nível 1), moderada e grave — sem idade mínima — e calcula a média dos 80% maiores salários (100% da média).")

    with st.expander("Meus dados pessoais e meu extrato CNIS estão seguros (LGPD)?"):
        st.markdown("Sim. O processamento ocorre de forma isolada e temporária apenas durante a leitura do arquivo. Assim que o relatório é exibido, os arquivos temporários são apagados automaticamente do servidor, em total conformidade com a Lei Geral de Proteção de Dados (LGPD).")

assunto_suporte = urllib.parse.quote("Suporte Técnico - Inteligência Previdenciária & Auditoria")
corpo_suporte = urllib.parse.quote("Olá! Preciso de suporte técnico no site Inteligência Previdenciária & Auditoria.")

st.markdown(f"""
<div class="floating-container">
    <a href="https://google.com" target="_blank" class="float-btn float-gemini">
        <span>✨ Dúvidas INSS (IA)</span>
    </a>
    <a href="mailto:{EMAIL_SUPORTE}?subject={assunto_suporte}&body={corpo_suporte}" class="float-btn float-support">
        <span>✉️ Suporte Técnico</span>
    </a>
</div>
""", unsafe_allow_html=True)
