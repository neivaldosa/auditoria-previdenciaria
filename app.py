import os
import uuid
import urllib.parse
import pandas as pd
import streamlit as st
from extrator_cnis import extrair_cnis_v5
from gerador_pdf import criar_pdf_executivo

st.set_page_config(
    page_title="Auditoria Previdenciária & Planejamento - Análise Detalhada do CNIS",
    page_icon="⚖️",
    layout="wide"
)

LINK_CHECKOUT_DOSSIE = "https://mpago.la"
LINK_CHECKOUT_PASSE = "https://mpago.la"

CODIGOS_DOSSIE_AVULSO = {"DOSSIE2026", "VIP19"}
CODIGOS_PASSE_ILIMITADO = {"ILIMITADO30", "PRO2026", "ADMIN777"}

if "plano_ativo" not in st.session_state:
    st.session_state["plano_ativo"] = "GRATUITO"

param_acesso = st.query_params.get("acesso", "").strip().upper()
if param_acesso in CODIGOS_PASSE_ILIMITADO:
    st.session_state["plano_ativo"] = "ILIMITADO"
elif param_acesso in CODIGOS_DOSSIE_AVULSO:
    st.session_state["plano_ativo"] = "DOSSIE"

st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    .block-container {
        padding-top: 1.1rem !important;
        padding-bottom: 1.5rem !important;
        max-width: 96% !important;
    }

    [data-testid="stSidebar"] {
        background-color: #F4F6F9;
        border-right: 1px solid #CBD5E1;
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
        border-radius: 8px;
        margin-bottom: 0.8rem;
    }
    [data-testid="stExpander"] summary {
        font-weight: 700 !important;
        color: #0F293A !important;
        font-size: 0.9rem !important;
        padding: 0.45rem 0.8rem !important;
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
        padding: 0.45rem 0.8rem;
        border-radius: 6px;
        font-weight: 700;
        text-decoration: none !important;
        font-size: 0.84rem;
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
    .feature-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(185px, 1fr));
        gap: 0.7rem;
        margin-top: 0.4rem;
        margin-bottom: 0.8rem;
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
        .faq-container {
        max-width: 920px;
        margin: 1.8rem auto 2rem auto;
    }
    .faq-header {
        text-align: center;
        color: #0F293A;
        font-size: 1.45rem;
        font-weight: 800;
        margin-bottom: 1rem;
    }
    details.faq-item {
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-left: 4px solid #C59B27;
        border-radius: 10px;
        margin-bottom: 0.65rem;
        padding: 0.85rem 1.2rem;
        box-shadow: 0 2px 6px rgba(15, 41, 58, 0.04);
        transition: all 0.2s ease;
    }
    details.faq-item[open] {
        background: linear-gradient(180deg, #FFFDF5 0%, #FFFFFF 100%);
        border-color: #C59B27;
    }
    details.faq-item summary {
        font-weight: 700;
        color: #0F293A;
        font-size: 0.96rem;
        cursor: pointer;
        list-style: none;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    details.faq-item summary::-webkit-details-marker {
        display: none;
    }
    details.faq-item summary::after {
        content: "⌄";
        font-size: 1.3rem;
        color: #C59B27;
        font-weight: 800;
        transition: transform 0.2s ease;
    }
    details.faq-item[open] summary::after {
        transform: rotate(180deg);
    }
    .faq-answer {
        margin-top: 0.65rem;
        padding-top: 0.65rem;
        border-top: 1px solid #E2E8F0;
        color: #475569;
        font-size: 0.88rem;
        line-height: 1.55;
    }
</style>
""", unsafe_allow_html=True)
st.markdown("""
<div id="topo-pagina" class="hero-box">
    <div class="hero-badge">Inteligência Previdenciária & Auditoria</div>
    <h1 class="hero-title">Diagnóstico Estratégico de Aposentadoria e CNIS</h1>
    <p class="hero-subtitle">Análise automatizada de vínculos, auditoria de pendências no INSS, enquadramento PCD (LC 142/2013) e projeção de renda.</p>
</div>
""", unsafe_allow_html=True)

with st.expander("📌 NÃO POSSUI O ARQUIVO PDF DO CNIS? CLIQUE AQUI PARA VER O PASSO A PASSO OFICIAL (2 MINUTOS)"):
    st.markdown("""
    1. **Acesse o Portal:** Entre em [meu.inss.gov.br](https://inss.gov.br) ou abra o aplicativo **Meu INSS** com sua conta **Gov.br**.
    2. **Busque o Extrato:** Clique em **"Extratos e Comprovantes"** e selecione **"Extrato de Contribuições (CNIS)"**.
    3. **Baixe a Versão Completa:** Role até o final, clique em **"Baixar Documento"** e escolha **"Vínculos, Contribuições e Remunerações"**.
    """)

with st.sidebar:
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

    st.markdown('<div class="sidebar-card"><div class="sidebar-title">🔑 Já possui Código de Acesso?</div>', unsafe_allow_html=True)
    codigo_digitado = st.text_input("Insira o código recebido após o Pix/Cartão:", placeholder="Ex: DOSSIE2026", label_visibility="collapsed").strip().upper()
    if st.button("Ativar Código", use_container_width=True):
        if codigo_digitado in CODIGOS_PASSE_ILIMITADO:
            st.session_state["plano_ativo"] = "ILIMITADO"
            st.success("Passe Ilimitado ativado!")
        elif codigo_digitado in CODIGOS_DOSSIE_AVULSO:
            st.session_state["plano_ativo"] = "DOSSIE"
            st.success("Dossiê Completo ativado!")
        else:
            st.error("Código inválido ou expirado.")

    plano_atual = st.session_state["plano_ativo"]
    badge_plano = "🟢 Plano Gratuito (Prévia)" if plano_atual == "GRATUITO" else ("🟡 Dossiê Completo Liberado" if plano_atual == "DOSSIE" else "👑 Passe Ilimitado Ativo")
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

                abas = pd.read_excel(caminho_temp_xlsx, sheet_name=None)
                df_vinculos = abas["1_CNIS_Extraido"]
                df_res = abas["2_Tempo_Liquido_Resumo"]
                df_fin = abas["5_Estimativa_Valor_RMI"]

                tempo_liq = str(df_res.loc[df_res["Indicador"] == "TEMPO LÍQUIDO DE CONTRIBUIÇÃO (COMUM)", "Valor"].tolist().pop())
                rmi_comum = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Regra Comum"), "Valor_Estimado"].tolist().pop())
                rmi_pcd = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Aposentadoria PCD"), "Valor_Estimado"].tolist().pop())
                dif_mensal = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Diferença Mensal"), "Valor_Estimado"].tolist().pop())
                tempo_card = tempo_liq.split("(").pop().replace(")", "") if "(" in tempo_liq else tempo_liq

                modo_pago = st.session_state["plano_ativo"] in ("DOSSIE", "ILIMITADO")

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
                        f"   -> Desbloqueie o Dossiê Completo por R$ 19,90 para liberar a solução de todas as pendências, a memória de cálculo desde 1994 e os downloads em PDF e Excel."
                    )

                    st.markdown(f"""
                    <div class="upsell-banner">
                        <h4 style="margin:0 0 0.3rem 0; color:#F3D06B;">⚠️ Encontramos {qtd_alertas} alerta(s) no seu CNIS — Saiba como corrigir e veja o valor do benefício</h4>
                        <p style="margin:0 0 0.8rem 0; font-size:0.88rem; color:#E2E8F0;">O <strong>Plano de Ação no INSS (como resolver)</strong>, o cálculo desde 1994 e os arquivos <strong>PDF Executivo + Excel</strong> são exclusivos do Dossiê Completo.</p>
                        <a href="{LINK_CHECKOUT_DOSSIE}" target="_blank" class="plan-cta cta-gold" style="display:inline-block; padding:0.55rem 1.3rem;">🔓 Desbloquear Dossiê Completo Agora (R$ 19,90)</a>
                    </div>
                    """, unsafe_allow_html=True)

                    tab0, tab1, tab2 = st.tabs(["📄 Parecer Executivo (Prévia)", "📋 Alertas de Vínculos", "⏱️ Resumo de Tempo"])
                    with tab0:
                        st.text_area("Diagnóstico Resumido:", value=parecer_gratuito, height=230)
                    with tab1:
                        df_vinculos_free = df_vinculos.copy()
                        df_vinculos_free["Plano_Acao_INSS"] = "🔒 Disponível no Dossiê Completo (R$ 19,90)"
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
<div style="font-size: 1.05rem; font-weight: 700; color: #0F293A; margin-top: 0.8rem;">💎 Escolha o Plano Ideal para sua Análise</div>
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
        <div style="text-align:center; font-size:0.78rem; color:#64748B; font-weight:600; padding:0.4rem; background:#F1F5F9; border-radius:6px;">Ativo — Envie seu PDF acima</div>
    </div>
    <div class="plan-card">
        <div>
            <div class="plan-top"><span class="plan-name" style="color:#0F293A;">Dossiê Completo Avulso</span></div>
            <div class="plan-price" style="color:#0F293A;">R$ 19,90 <span>/ único</span></div>
            <ul class="plan-list" style="color:#334155;">
                <li>✅ Plano de ação (como corrigir erros)</li>
                <li>✅ Cálculo RMI (1994–2026), Especial e PCD</li>
                <li>✅ Download PDF Executivo + Planilha Excel</li>
            </ul>
        </div>
        <a href="__LINK_DOSSIE__" target="_blank" class="plan-cta cta-outline">Adquirir Dossiê Avulso</a>
    </div>
    <div class="plan-card featured">
        <div>
            <div class="plan-top">
                <span class="plan-name">Passe Ilimitado</span>
                <span class="plan-badge">Mais Escolhido</span>
            </div>
            <div class="plan-price" style="color:#F3D06B;">R$ 39,90 <span>/ 7 dias (ou R$ 69,90 / 30d)</span></div>
            <ul class="plan-list" style="color:#E2E8F0;">
                <li>✅ <strong>Simulações e PDFs/Excel ilimitados</strong></li>
                <li>✅ Teste cenários (PCD, Especial, futuro)</li>
                <li>✅ Analise também familiares ou clientes</li>
            </ul>
        </div>
        <a href="__LINK_PASSE__" target="_blank" class="plan-cta cta-gold">Liberar Passe Ilimitado</a>
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
""".replace("__LINK_DOSSIE__", LINK_CHECKOUT_DOSSIE).replace("__LINK_PASSE__", LINK_CHECKOUT_PASSE), unsafe_allow_html=True)
st.markdown("""
<div class="faq-container">
    <div class="faq-header">Perguntas Frequentes</div>

    <details class="faq-item" open>
        <summary>Por que é tão importante entender o seu CNIS e ajustar possíveis problemas antes de entrar com o pedido de aposentadoria?</summary>
        <div class="faq-answer">
            Obter acesso aos detalhes do seu CNIS e realizar uma auditoria prévia antes de dar entrada na aposentadoria é crucial para evitar indeferimentos, perda de tempo de contribuição ou cálculo de benefício menor do que você tem direito — garantindo tranquilidade em um momento em que o objetivo é descansar. Dependendo do caso, os relatórios da <strong>Inteligência Previdenciária &amp; Auditoria</strong> entregam o diagnóstico técnico mastigado por uma fração do custo de uma consulta particular.
        </div>
    </details>

    <details class="faq-item">
        <summary>Qual a diferença entre Cálculo Previdenciário e Planejamento Previdenciário?</summary>
        <div class="faq-answer">
            O <strong>Cálculo Previdenciário</strong> apura o tempo líquido comprovado e estima o valor atual do benefício com base nos seus salários desde julho de 1994. Já o <strong>Planejamento Previdenciário</strong> compara diferentes regras e cenários (Regra Geral, Aposentadoria Especial por exposição a risco e Aposentadoria PCD pela LC 142/2013) para identificar qual caminho antecipa sua aposentadoria ou gera a maior renda mensal.
        </div>
    </details>

    <details class="faq-item">
        <summary>Onde encontro essas funcionalidades?</summary>
        <div class="faq-answer">
            Basta enviar o arquivo PDF do seu extrato CNIS no quadro de envio acima. Você pode adquirir o <strong>Dossiê Completo Avulso</strong> (análise definitiva de 1 extrato com downloads em PDF e Excel) ou contratar o <strong>Passe Ilimitado de 7 dias ou 30 dias</strong> para testar diferentes cenários e analisar extratos de familiares.
        </div>
    </details>

    <details class="faq-item">
        <summary>Preciso do CNIS para começar?</summary>
        <div class="faq-answer">
            Sim, esse é o único documento necessário (<strong>CNIS em PDF</strong> — <a href="#topo-pagina" style="color: #C59B27; font-weight: 700; text-decoration: underline;">saiba como emitir gratuitamente no Meu INSS clicando aqui</a>) para que nosso motor faça a leitura automática e entregue uma análise detalhada, explicando ponto a ponto o que você precisa fazer para regularizar pendências e seguir com sua aposentadoria.
        </div>
    </details>

    <details class="faq-item">
        <summary>Posso obter uma demonstração gratuita para entender como o serviço funciona?</summary>
        <div class="faq-answer">
            Sim! Basta enviar seu PDF normalmente sem inserir nenhum código — o sistema executa automaticamente o <strong>Diagnóstico Inicial (Grátis)</strong>, exibindo seu tempo líquido de contribuição já descontados os períodos concomitantes e apontando se existem alertas ou pendências nos seus vínculos.
        </div>
    </details>

    <details class="faq-item">
        <summary>Tenho diagnóstico de TEA (Autismo) ou outra deficiência. O sistema calcula a regra PCD?</summary>
        <div class="faq-answer">
            Sim. Ao marcar a opção <strong>"Enquadramento PCD / TEA (LC 142/2013)"</strong> na barra lateral, o sistema simula automaticamente o tempo exigido para deficiência leve (como TEA Nível 1), moderada e grave — sem exigência de idade mínima — e calcula a média dos 80% maiores salários (100% da média), comparando o ganho financeiro em relação à Regra Geral.
        </div>
    </details>

    <details class="faq-item">
        <summary>Meus dados pessoais e meu extrato CNIS ficam armazenados no site (LGPD)?</summary>
        <div class="faq-answer">
            Não. Todo o processamento é feito de forma isolada e temporária apenas durante os segundos necessários para gerar suas tabelas e relatórios. Assim que a tela é carregada, os arquivos temporários são excluídos automaticamente do servidor, em total conformidade com a Lei Geral de Proteção de Dados (LGPD).
        </div>
    </details>
</div>
""", unsafe_allow_html=True)
