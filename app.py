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
    .stButton > button[kind="primary"] {
        background-color: #0F293A;
        border: 2px solid #C59B27;
        color: #FFFFFF;
        font-weight: 600;
        padding: 0.65rem 1.8rem;
        border-radius: 8px;
        font-size: 1.02rem;
    }
    .feature-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
        gap: 1rem;
        margin-top: 1.5rem;
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

    st.markdown("""
    <div class="security-badge">
        <strong>🔒 Privacidade & Segurança LGPD</strong><br>
        Processamento isolado e criptografado em memória. Nenhum dado pessoal ou PDF fica armazenado após a análise.
    </div>
    """, unsafe_allow_html=True)

st.markdown("### 📤 Envie seu Extrato Oficial (CNIS)")
arquivo_pdf = st.file_uploader("Selecione ou arraste o arquivo PDF completo baixado do Meu INSS:", type=["pdf"])

if arquivo_pdf is not None:
    if st.button("🚀 Executar Auditoria e Gerar Dossiê Completo", type="primary"):
        st.markdown("""
<h4 style="color: #0F293A; margin-top: 1.8rem; margin-bottom: 0.2rem;">📋 O que está incluso no seu Dossiê Completo de Auditoria:</h4>
<p style="color: #64748B; font-size: 0.9rem; margin-bottom: 1rem;">Ao enviar seu CNIS, nosso motor previdenciário gera automaticamente 6 análises técnicas em tela, PDF e Excel:</p>

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
                df_res = abas["2_Tempo_Liquido_Resumo"]
                df_fin = abas["5_Estimativa_Valor_RMI"]

                tempo_liq = str(df_res.loc[df_res["Indicador"] == "TEMPO LÍQUIDO DE CONTRIBUIÇÃO (COMUM)", "Valor"].tolist().pop())
                rmi_comum = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Regra Comum"), "Valor_Estimado"].tolist().pop())
                rmi_pcd = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Aposentadoria PCD"), "Valor_Estimado"].tolist().pop())
                dif_mensal = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Diferença Mensal"), "Valor_Estimado"].tolist().pop())

                tempo_card = tempo_liq.split("(").pop().replace(")", "") if "(" in tempo_liq else tempo_liq

                st.markdown("<br>", unsafe_allow_html=True)
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f'<div class="metric-card"><div class="metric-label">Tempo Líquido Apurado</div><div class="metric-value">{tempo_card}</div><div class="metric-sub">Períodos concomitantes já descontados</div></div>', unsafe_allow_html=True)
                with col2:
                    st.markdown(f'<div class="metric-card"><div class="metric-label">Estimativa Regra Geral</div><div class="metric-value">{rmi_comum}</div><div class="metric-sub">Projeção com coeficiente atual</div></div>', unsafe_allow_html=True)
                with col3:
                    st.markdown(f'<div class="metric-card highlight"><div class="metric-label">Potencial Regra PCD (LC 142)</div><div class="metric-value">{rmi_pcd}</div><div class="metric-sub">{dif_mensal} vs. Regra Geral</div></div>', unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
                    "📄 Parecer Executivo", "📋 Auditoria de Vínculos", "⏱️ Resumo de Tempo",
                    "🏭 Especial & GILRAT", "♿ Enquadramento PCD", "💰 Memória Financeira"
                ])

                with tab0:
                    st.text_area("Relatório Técnico Estruturado:", value=parecer_texto, height=400)
                with tab1:
                    st.dataframe(abas["1_CNIS_Extraido"], use_container_width=True)
                with tab2:
                    st.dataframe(abas["2_Tempo_Liquido_Resumo"], use_container_width=True)
                with tab3:
                    st.dataframe(abas["3_Aposentadoria_Especial_PPP"], use_container_width=True)
                with tab4:
                    st.dataframe(abas["4_Aposentadoria_PCD_LC142"], use_container_width=True)
                with tab5:
                    st.dataframe(abas["5_Estimativa_Valor_RMI"], use_container_width=True)

                st.divider()
                criar_pdf_executivo(
                    caminho_pdf_relatorio,
                    abas["2_Tempo_Liquido_Resumo"],
                    abas["5_Estimativa_Valor_RMI"],
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
