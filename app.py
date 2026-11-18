import os
import uuid
import urllib.parse
import pandas as pd
import streamlit as st
from extrator_cnis import extrair_cnis_v5
from gerador_pdf import criar_pdf_executivo

st.set_page_config(page_title="Auditoria Previdenciaria & Planejamento", page_icon="⚖️", layout="wide")

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
</style>
""", unsafe_allow_html=True)
st.markdown("""
<div class="hero-box">
    <div class="hero-badge">Inteligencia Previdenciaria & Auditoria</div>
    <h1 class="hero-title">Diagnostico Estrategico de Aposentadoria e CNIS</h1>
    <p class="hero-subtitle">Analise automatizada de vinculos, auditoria de pendencias no INSS, enquadramento PCD (LC 142/2013) e projecao de renda.</p>
</div>
""", unsafe_allow_html=True)

with st.expander("📌 NAO POSSUI O ARQUIVO PDF DO CNIS? CLIQUE AQUI PARA VER O PASSO A PASSO OFICIAL (2 MINUTOS)"):
    st.markdown("""
    Para que a auditoria leia todos os seus salarios e pendencias, baixe a **versao completa** no portal oficial:

    1. **Acesse o Portal:** Entre em [meu.inss.gov.br](https://inss.gov.br) ou abra o aplicativo **Meu INSS** e entre com sua conta **Gov.br**.
    2. **Busque o Extrato:** Clique em **"Extratos e Comprovantes"** (ou digite *CNIS* na barra de busca).
    3. **Selecione o Servico:** Clique em **"Extrato de Contribuicoes (CNIS)"**.
    4. **Baixe a Versao Completa:** Role ate o final da pagina, clique em **"Baixar Documento"** e selecione obrigatoriamente **"Vinculos, Contribuicoes e Remuneracoes"**.
    5. **Envie Abaixo:** Arraste o arquivo PDF salvo para o quadro de envio abaixo.
    """)

with st.sidebar:
    st.markdown("### ⚙️ Configure seu Perfil")
    sexo = st.selectbox("Sexo biologico (Regra INSS):", options=["M", "F"], format_func=lambda x: "Masculino" if x == "M" else "Feminino")

    st.markdown('<div class="sidebar-card"><div class="sidebar-title">1. Ambiente de Trabalho</div>', unsafe_allow_html=True)
    ambiente = st.radio(
        "Selecione sua atuacao predominante:",
        options=["Administrativo / Escritorio", "Operacional / Area de Risco (Insalubridade/Periculosidade)"],
        index=0,
        label_visibility="collapsed"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    ambiente_operacional = (ambiente == "Operacional / Area de Risco (Insalubridade/Periculosidade)")

    st.markdown('<div class="sidebar-card"><div class="sidebar-title">2. Enquadramento PCD / TEA</div>', unsafe_allow_html=True)
    possui_pcd = st.checkbox("Possui deficiencia / TEA (LC 142/2013)?", value=True)
    st.caption("Marque se possui laudo medico de deficiencia fisica, sensorial, intelectual ou TEA.")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="security-badge">
        <strong>🔒 Privacidade & Seguranca LGPD</strong><br>
        Processamento isolado e criptografado em memoria. Nenhum dado pessoal ou PDF fica armazenado apos a analise.
    </div>
    """, unsafe_allow_html=True)

st.markdown("### 📤 Envie seu Extrato Oficial (CNIS)")
arquivo_pdf = st.file_uploader("Selecione ou arraste o arquivo PDF completo baixado do Meu INSS:", type=["pdf"])

if arquivo_pdf is not None:
    if st.button("🚀 Executar Auditoria e Gerar Dossie Completo", type="primary"):
        with st.spinner("Processando vinculos, cruzando tetos historicos e auditando indicadores..."):
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

                tempo_liq = str(df_res.loc[df_res["Indicador"] == "TEMPO LIQUIDO DE CONTRIBUICAO (COMUM)", "Valor"].tolist().pop())
                rmi_comum = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Regra Comum"), "Valor_Estimado"].tolist().pop())
                rmi_pcd = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Aposentadoria PCD"), "Valor_Estimado"].tolist().pop())
                dif_mensal = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Diferenca Mensal"), "Valor_Estimado"].tolist().pop())

                tempo_card = tempo_liq.split("(").pop().replace(")", "") if "(" in tempo_liq else tempo_liq

                st.markdown("<br>", unsafe_allow_html=True)
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f'<div class="metric-card"><div class="metric-label">Tempo Liquido Apurado</div><div class="metric-value">{tempo_card}</div><div class="metric-sub">Periodos concomitantes ja descontados</div></div>', unsafe_allow_html=True)
                with col2:
                    st.markdown(f'<div class="metric-card"><div class="metric-label">Estimativa Regra Geral</div><div class="metric-value">{rmi_comum}</div><div class="metric-sub">Projecao com coeficiente atual</div></div>', unsafe_allow_html=True)
                with col3:
                    st.markdown(f'<div class="metric-card highlight"><div class="metric-label">Potencial Regra PCD (LC 142)</div><div class="metric-value">{rmi_pcd}</div><div class="metric-sub">{dif_mensal} vs. Regra Geral</div></div>', unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
                    "📄 Parecer Executivo", "📋 Auditoria de Vinculos", "⏱️ Resumo de Tempo",
                    "🏭 Especial & GILRAT", "♿ Enquadramento PCD", "💰 Memoria Financeira"
                ])

                with tab0:
                    st.text_area("Relatorio Tecnico Estruturado:", value=parecer_texto, height=400)
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
                        label="📑 Baixar Relatorio Executivo (.pdf)",
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
                    msg_whats = urllib.parse.quote(f"Ola! Acabei de gerar minha Auditoria Previdenciaria (Tempo Liquido: {tempo_card} | Potencial PCD: {rmi_pcd}) e gostaria de agendar uma analise com o escritorio parceiro.")
                    link_whats = f"https://wa.me{msg_whats}"
                    st.link_button("⚖️ Enviar para Escritorio Parceiro", url=link_whats, use_container_width=True)

            finally:
                for arq in [caminho_temp_pdf, caminho_temp_xlsx, caminho_pdf_relatorio]:
                    if os.path.exists(arq):
                        os.remove(arq)
