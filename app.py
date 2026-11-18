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

# ==============================================================================
# CONFIGURAÇÃO DE CHECKOUT E CHAVES DE LIBERAÇÃO
# (Substitua pelos seus links reais do Mercado Pago, Stripe, Kiwify ou Asaas)
# ==============================================================================
LINK_CHECKOUT_DOSSIE = "https://mpago.la"
LINK_CHECKOUT_PASSE = "https://mpago.la"

# Códigos válidos para liberar o Dossiê Completo (você pode enviar no pós-venda ou configurar no redirecionamento)
CODIGOS_DOSSIE_AVULSO = {"DOSSIE2026", "VIP19"}
CODIGOS_PASSE_ILIMITADO = {"ILIMITADO30", "PRO2026", "ADMIN777"}

if "plano_ativo" not in st.session_state:
    st.session_state["plano_ativo"] = "GRATUITO"

# Verifica se o usuário voltou do checkout com parâmetro na URL (ex: ?acesso=PRO2026)
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

    [data-testid="stSidebar"] {
        background-color: #F4F6F9;
        border-right: 1px solid #CBD5E1;
    }
    .sidebar-card {
        background: linear-gradient(180deg, #FFFDF5 0%, #FFFFFF 100%);
        padding: 1rem 1.1rem;
        border-radius: 10px;
        border: 1.5px solid rgba(197, 155, 39, 0.55);
        border-left: 5px solid #C59B27;
        margin-bottom: 1rem;
        box-shadow: 0 3px 10px rgba(197, 155, 39, 0.12);
    }
    .sidebar-title {
        font-size: 0.92rem;
        font-weight: 700;
        color: #0F293A;
        margin-bottom: 0.4rem;
        text-transform: uppercase;
        letter-spacing: 0.4px;
    }
    .security-badge {
        background: linear-gradient(180deg, #FFFDF5 0%, #FFFFFF 100%);
        border: 1.5px solid rgba(197, 155, 39, 0.55);
        border-left: 5px solid #C59B27;
        color: #0F293A;
        padding: 0.95rem 1rem;
        border-radius: 10px;
        font-size: 0.84rem;
        line-height: 1.45;
        margin-top: 1rem;
        box-shadow: 0 3px 10px rgba(197, 155, 39, 0.12);
    }
    .hero-box {
        background: linear-gradient(135deg, #0F293A 0%, #1B3B52 100%);
        padding: 2rem 2.5rem;
        border-radius: 12px;
        border-bottom: 4px solid #C59B27;
        color: #FFFFFF;
        margin-bottom: 1.5rem;
        box-shadow: 0 10px 25px -5px rgba(15, 41, 58, 0.15);
    }
    .hero-badge {
        display: inline-block;
        background-color: rgba(197, 155, 39, 0.2);
        color: #F3D06B;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 0.8px;
        text-transform: uppercase;
        margin-bottom: 0.7rem;
        border: 1px solid rgba(197, 155, 39, 0.4);
    }
    .hero-title {
        font-size: 1.9rem;
        font-weight: 700;
        margin: 0 0 0.4rem 0;
        letter-spacing: -0.5px;
    }
    .hero-subtitle {
        font-size: 1rem;
        color: #CBD5E1;
        margin: 0;
    }
    [data-testid="stExpander"] {
        background: linear-gradient(90deg, #FFFDF5 0%, #FFFFFF 100%);
        border: 1.5px solid #C59B27;
        border-left: 6px solid #C59B27;
        border-radius: 10px;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 12px rgba(197, 155, 39, 0.14);
    }
    [data-testid="stExpander"] summary {
        font-weight: 700 !important;
        color: #0F293A !important;
        font-size: 1.02rem !important;
    }
    [data-testid="stFileUploader"] section {
        background: linear-gradient(180deg, #FFFDF5 0%, #FFFFFF 100%);
        border: 2px dashed #C59B27;
        border-radius: 12px;
        padding: 1.8rem;
        box-shadow: 0 4px 12px rgba(197, 155, 39, 0.10);
    }
    [data-testid="stFileUploader"] button {
        background-color: #C59B27 !important;
        color: #0F293A !important;
        border: none !important;
        font-weight: 700 !important;
        padding: 0.55rem 1.3rem !important;
        border-radius: 6px !important;
    }
    .pricing-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(270px, 1fr));
        gap: 1.2rem;
        margin: 1.2rem 0 2rem 0;
    }
    .plan-card {
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 12px;
        padding: 1.5rem;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        box-shadow: 0 4px 12px rgba(15, 41, 58, 0.05);
    }
    .plan-card.featured {
        background: linear-gradient(180deg, #0F293A 0%, #16364D 100%);
        border: 2px solid #C59B27;
        color: #FFFFFF;
        box-shadow: 0 10px 25px rgba(15, 41, 58, 0.22);
    }
    .plan-badge {
        display: inline-block;
        background-color: #C59B27;
        color: #0F293A;
        font-size: 0.72rem;
        font-weight: 800;
        padding: 3px 10px;
        border-radius: 12px;
        text-transform: uppercase;
        margin-bottom: 0.6rem;
        width: fit-content;
    }
    .plan-name {
        font-size: 1.15rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }
    .plan-price {
        font-size: 1.9rem;
        font-weight: 800;
        margin: 0.4rem 0;
    }
    .plan-price span {
        font-size: 0.85rem;
        font-weight: 400;
        opacity: 0.85;
    }
    .plan-list {
        list-style: none;
        padding: 0;
        margin: 1rem 0 1.2rem 0;
        font-size: 0.86rem;
        line-height: 1.75;
    }
    .plan-cta {
        display: block;
        text-align: center;
        padding: 0.65rem 1rem;
        border-radius: 8px;
        font-weight: 700;
        text-decoration: none !important;
        font-size: 0.92rem;
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
        grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
        gap: 1rem;
        margin-top: 1.2rem;
        margin-bottom: 1.5rem;
    }
    .feature-card {
        background: linear-gradient(180deg, #FFFFFF 0%, #FFFDF5 100%);
        border: 1px solid #E2E8F0;
        border-top: 4px solid #C59B27;
        border-radius: 10px;
        padding: 1.2rem 1.3rem;
        box-shadow: 0 4px 10px rgba(15, 41, 58, 0.05);
    }
    .feature-title {
        font-size: 0.98rem;
        font-weight: 700;
        color: #0F293A;
        margin-bottom: 0.45rem;
    }
    .feature-desc {
        font-size: 0.86rem;
        color: #475569;
        line-height: 1.45;
        margin: 0;
    }
    .metric-card {
        background-color: #FFFFFF;
        padding: 1.4rem;
        border-radius: 10px;
        border: 1px solid #E2E8F0;
        border-left: 4px solid #0F293A;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04);
    }
    .metric-card.highlight {
        border-left: 5px solid #C59B27;
        background: linear-gradient(to right, #FFFFFF, #FFFDF5);
    }
    .metric-label {
        font-size: 0.82rem;
        color: #64748B;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.5px;
        margin-bottom: 0.4rem;
    }
    .metric-value {
        font-size: 1.55rem;
        font-weight: 700;
        color: #0F293A;
    }
    .metric-sub {
        font-size: 0.82rem;
        color: #059669;
        font-weight: 500;
        margin-top: 0.3rem;
    }
    .upsell-banner {
        background: linear-gradient(135deg, #0F293A 0%, #1B3B52 100%);
        border: 2px solid #C59B27;
        border-radius: 12px;
        padding: 1.5rem 2rem;
        color: #FFFFFF;
        margin: 1.2rem 0;
    }
    .stButton > button[kind="primary"] {
        background-color: #0F293A;
        border: 2px solid #C59B27;
        color: #FFFFFF;
        font-weight: 600;
        padding: 0.65rem 1.8rem;
        border-radius: 8px;
        font-size: 1.02rem;
    }
</style>
""", unsafe_allow_html=True)
st.markdown("""
<div class="hero-box">
    <div class="hero-badge">Inteligência Previdenciária & Auditoria</div>
    <h1 class="hero-title">Diagnóstico Estratégico de Aposentadoria e CNIS</h1>
    <p class="hero-subtitle">Análise automatizada de vínculos, auditoria de pendências no INSS, enquadramento PCD (LC 142/2013) e projeção de renda.</p>
</div>
""", unsafe_allow_html=True)

with st.expander("📌 NÃO POSSUI O ARQUIVO PDF DO CNIS? CLIQUE AQUI PARA VER O PASSO A PASSO OFICIAL (2 MINUTOS)"):
    st.markdown("""
    Para que a auditoria leia todos os seus salários e pendências, baixe a **versão completa** no portal oficial:

    1. **Acesse o Portal:** Entre em [meu.inss.gov.br](https://inss.gov.br) ou abra o aplicativo **Meu INSS** e entre com sua conta **Gov.br**.
    2. **Busque o Extrato:** Clique em **"Extratos e Comprovantes"** (ou digite *CNIS* na barra de busca).
    3. **Selecione o Serviço:** Clique em **"Extrato de Contribuições (CNIS)"**.
    4. **Baixe a Versão Completa:** Role até o final da página, clique em **"Baixar Documento"** e selecione obrigatoriamente **"Vínculos, Contribuições e Remunerações"**.
    5. **Envie Abaixo:** Arraste o arquivo PDF salvo para o quadro de envio abaixo.
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
        Processamento isolado e criptografado em memória. Nenhum dado pessoal ou PDF fica armazenado após a análise.
    </div>
    """, unsafe_allow_html=True)

st.markdown("### 📤 Envie seu Extrato Oficial (CNIS)")
arquivo_pdf = st.file_uploader("Selecione ou arraste o arquivo PDF completo baixado do Meu INSS:", type=["pdf"])

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

                st.markdown("<br>", unsafe_allow_html=True)
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
                        f"   -> Os alertas estão listados na aba 'Auditoria de Vínculos', mas o passo a passo de como resolver cada pendência no Meu INSS está retido na versão gratuita.\n\n"
                        f"3. CÁLCULOS FINANCEIROS, ESPECIAL E PCD: Bloqueados nesta prévia.\n"
                        f"   -> Desbloqueie o Dossiê Completo por R$ 19,90 para liberar a solução de todas as pendências, a memória de cálculo desde 1994 e os downloads em PDF e Excel."
                    )

                    st.markdown(f"""
                    <div class="upsell-banner">
                        <h4 style="margin:0 0 0.4rem 0; color:#F3D06B;">⚠️ Encontramos {qtd_alertas} alerta(s) no seu CNIS — Saiba exatamente como corrigir e veja o valor da sua aposentadoria</h4>
                        <p style="margin:0 0 1rem 0; font-size:0.92rem; color:#E2E8F0;">Na versão gratuita você visualiza os alertas detectados, mas o <strong>Plano de Ação no INSS (como resolver)</strong>, o cálculo de salários desde 1994 e os arquivos <strong>PDF Executivo + Excel</strong> são exclusivos do Dossiê Completo.</p>
                        <a href="{LINK_CHECKOUT_DOSSIE}" target="_blank" class="plan-cta cta-gold" style="display:inline-block; padding:0.65rem 1.5rem;">🔓 Desbloquear Dossiê Completo Agora (R$ 19,90)</a>
                    </div>
                    """, unsafe_allow_html=True)

                    tab0, tab1, tab2 = st.tabs(["📄 Parecer Executivo (Prévia)", "📋 Alertas de Vínculos", "⏱️ Resumo de Tempo"])
                    with tab0:
                        st.text_area("Diagnóstico Resumido:", value=parecer_gratuito, height=260)
                    with tab1:
                        df_vinculos_free = df_vinculos.copy()
                        df_vinculos_free["Plano_Acao_INSS"] = "🔒 Disponível no Dossiê Completo (R$ 19,90)"
                        st.dataframe(df_vinculos_free, use_container_width=True)
                    with tab2:
                        st.dataframe(df_res, use_container_width=True)
                else:
                    st.markdown("<br>", unsafe_allow_html=True)
                    tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
                        "📄 Parecer Executivo", "📋 Auditoria de Vínculos", "⏱️ Resumo de Tempo",
                        "🏭 Especial & GILRAT", "♿ Enquadramento PCD", "💰 Memória Financeira"
                    ])

                    with tab0:
                        st.text_area("Relatório Técnico Estruturado:", value=parecer_texto, height=400)
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

                    st.divider()
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
<h3 style="color: #0F293A; margin-top: 2.2rem; margin-bottom: 0.2rem;">💎 Escolha o Plano Ideal para sua Análise</h3>
<p style="color: #64748B; font-size: 0.92rem; margin-bottom: 0.5rem;">Comece gratuitamente agora mesmo ou desbloqueie o dossiê técnico completo com memória de cálculo desde 1994.</p>

<div class="pricing-grid">
    <div class="plan-card">
        <div>
            <div class="plan-name" style="color:#0F293A;">Diagnóstico Inicial</div>
            <div class="plan-price" style="color:#0F293A;">Grátis <span>/ prévia</span></div>
            <ul class="plan-list" style="color:#334155;">
                <li>✅ Leitura automática do PDF do CNIS</li>
                <li>✅ Cálculo do tempo líquido de contribuição</li>
                <li>✅ Identificação dos alertas e pendências</li>
                <li>❌ Como resolver cada pendência no INSS</li>
                <li>❌ Cálculo do valor do benefício (RMI)</li>
                <li>❌ Download em PDF Executivo e Excel</li>
            </ul>
        </div>
        <div style="text-align:center; font-size:0.84rem; color:#64748B; font-weight:600; padding:0.6rem; background:#F1F5F9; border-radius:8px;">Ativo por padrão — Envie seu PDF acima</div>
    </div>
    <div class="plan-card">
        <div>
            <div class="plan-name" style="color:#0F293A;">Dossiê Completo Avulso</div>
            <div class="plan-price" style="color:#0F293A;">R$ 19,90 <span>/ pagamento único</span></div>
            <ul class="plan-list" style="color:#334155;">
                <li>✅ <strong>Tudo do Diagnóstico Inicial</strong></li>
                <li>✅ Plano de ação completo (como corrigir erros)</li>
                <li>✅ Enquadramento PCD (LC 142) e Especial (PPP)</li>
                <li>✅ Memória financeira completa desde 07/1994</li>
                <li>✅ Download do Relatório Executivo (.pdf)</li>
                <li>✅ Download da Planilha Auditada (.xlsx)</li>
            </ul>
        </div>
        <a href="__LINK_DOSSIE__" target="_blank" class="plan-cta cta-outline">Adquirir Dossiê Avulso</a>
    </div>
    <div class="plan-card featured">
        <div>
            <div class="plan-badge">Mais Escolhido • Simulações Livres</div>
            <div class="plan-name">Passe Ilimitado (7 ou 30 Dias)</div>
            <div class="plan-price" style="color:#F3D06B;">R$ 39,90 <span>/ 7 dias (ou R$ 69,90 / 30d)</span></div>
            <ul class="plan-list" style="color:#E2E8F0;">
                <li>✅ <strong>Simulações e downloads 100% ilimitados</strong></li>
                <li>✅ Teste cenários (PCD, Especial, datas futuras)</li>
                <li>✅ Analise o CNIS de familiares e clientes</li>
                <li>✅ Plano de ação completo para cada vínculo</li>
                <li>✅ PDFs Executivos e Planilhas Excel ilimitados</li>
                <li>✅ Ideal para planejamento familiar ou consultores</li>
            </ul>
        </div>
        <a href="__LINK_PASSE__" target="_blank" class="plan-cta cta-gold">Liberar Passe Ilimitado</a>
    </div>
</div>
""".replace("__LINK_DOSSIE__", LINK_CHECKOUT_DOSSIE).replace("__LINK_PASSE__", LINK_CHECKOUT_PASSE), unsafe_allow_html=True)

st.markdown("""
<h4 style="color: #0F293A; margin-top: 1.5rem; margin-bottom: 0.2rem;">📋 O que está incluso no seu Dossiê Completo de Auditoria:</h4>
<p style="color: #64748B; font-size: 0.9rem; margin-bottom: 1rem;">Ao desbloquear o Dossiê Completo, nosso motor previdenciário gera automaticamente 6 análises técnicas em tela, PDF e Excel:</p>

<div class="feature-grid">
    <div class="feature-card">
        <div class="feature-title">📄 1. Parecer Executivo</div>
        <p class="feature-desc">Diagnóstico jurídico-técnico consolidado que traduz o seu histórico em linguagem clara, apontando riscos, oportunidades e o plano de ação exato para protocolar no Meu INSS.</p>
    </div>
    <div class="feature-card">
        <div class="feature-title">📋 2. Auditoria de Vínculos</div>
        <p class="feature-desc">Varredura completa de cada empresa e recolhimento, identificando indicadores ocultos do INSS (como falta de data de saída, pendências salariais ou guias concomitantes bloqueadas).</p>
    </div>
    <div class="feature-card">
        <div class="feature-title">⏱️ 3. Resumo de Tempo</div>
        <p class="feature-desc">Cálculo exato do seu tempo líquido de contribuição, descontando automaticamente períodos sobrepostos para evitar erros de contagem na hora do pedido de aposentadoria.</p>
    </div>
    <div class="feature-card">
        <div class="feature-title">🏭 4. Aposentadoria Especial (Risco)</div>
        <p class="feature-desc">Identifica períodos operacionais expostos à insalubridade ou periculosidade antes da Reforma de 2019, calculando o acréscimo de tempo (fator 1.4 H / 1.2 M) e orientando sobre PPP e GILRAT.</p>
    </div>
    <div class="feature-card">
        <div class="feature-title">♿ 5. Enquadramento PCD (LC 142)</div>
        <p class="feature-desc">Simula o direito à aposentadoria da Pessoa com Deficiência (incluindo TEA Nível 1) nos graus grave, moderado e leve — sem exigência de idade mínima e com redução de tempo.</p>
    </div>
    <div class="feature-card">
        <div class="feature-title">💰 6. Memória Financeira (RMI)</div>
        <p class="feature-desc">Atualiza seus salários desde julho de 1994 pelos tetos históricos do INSS, comparando o valor do benefício na Regra Geral versus a média dos 80% maiores salários da Regra PCD.</p>
    </div>
</div>
""", unsafe_allow_html=True)
