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
    verificar_codigo_otp_recuperacao,
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
                        top=borda_ouro if r_dica == linha_ini else None,
                        bottom=borda_ouro if r_dica == linha_fim else None,
                    )

    saida_final = io.BytesIO()
    wb.save(saida_final)
    saida_final.seek(0)
    return saida_final.getvalue()


import json
import urllib.error
import urllib.request


def _carregar_prova_dossie_b64() -> str:
    candidatos = [
        "prova_resultados_dossie.png",
        "prova_resultados_dossie.jpg",
        "prova_resultados_dossie.jpeg",
        "resultados_dossie.png",
        "resultados_dossie.jpg",
    ]
    for nome in candidatos:
        if os.path.exists(nome):
            try:
                with open(nome, "rb") as f:
                    return base64.b64encode(f.read()).decode("utf-8")
            except Exception:
                pass
    return ""


def _carregar_prova_seguranca_b64() -> str:
    candidatos = [
        "prova_seguranca_gemini.png",
        "prova_seguranca_gemini.jpg",
        "prova_seguranca_gemini.jpeg",
        "seguranca_gemini.png",
        "seguranca_lgpd.png",
    ]
    for nome in candidatos:
        if os.path.exists(nome):
            try:
                with open(nome, "rb") as f:
                    return base64.b64encode(f.read()).decode("utf-8")
            except Exception:
                pass
    return ""


def _obter_gemini_api_key() -> str:
    chave = ""
    try:
        if "GEMINI_API_KEY" in st.secrets:
            chave = str(st.secrets["GEMINI_API_KEY"]).strip()
    except Exception:
        chave = ""
    if not chave:
        chave = str(os.environ.get("GEMINI_API_KEY", "")).strip()
    return chave


def _parece_pedido_calculo_ou_extrato(texto: str) -> bool:
    t = (texto or "").strip().lower()
    if len(t) > 650:
        return True
    gatilhos_extrato = [
        "seq.", "nit:", "código emp.", "remunerações", "indicadores",
        "data início", "data fim", "últ. remun.", "vínculos e contribuições",
    ]
    if sum(1 for g in gatilhos_extrato if g in t) >= 2:
        return True
    pedidos_servico = [
        "faça um levantamento", "fazer um levantamento", "faz um levantamento",
        "audite meu cnis", "auditar meu cnis", "analise meu cnis", "analisar meu cnis",
        "calcule meu tempo", "calcular meu tempo", "quanto tempo eu tenho",
        "quanto vou receber", "calcule minha aposentadoria", "calcular minha aposentadoria",
        "calcule minha rmi", "calcular minha rmi", "simule minha aposentadoria",
    ]
    return any(p in t for p in pedidos_servico)


def _responder_duvida_inss(pergunta: str, historico: list[dict] | None = None) -> str:
    if _parece_pedido_calculo_ou_extrato(pergunta):
        return (
            "Esse tipo de levantamento, contagem de tempo líquido ou cálculo de valor (RMI) "
            "não é feito aqui no chat — ele exige as regras determinísticas do motor oficial da página.\n\n"
            "Para fazer isso com precisão:\n"
            "1. Envie o **PDF oficial do seu CNIS** no campo de upload acima para ver o diagnóstico gratuito.\n"
            "2. Para liberar o plano de ação completo, cálculo desde 1994, Especial, PCD e os arquivos **PDF + Excel**, "
            "crie sua conta na **Área do Assinante** e escolha um dos planos do site."
        )

    api_key = _obter_gemini_api_key()
    if not api_key:
        return "⚠️ Chave GEMINI_API_KEY não configurada nos Secrets do Streamlit Cloud."

    instrucoes = (
        "Você é o assistente conversacional do site Inteligência Previdenciária & Auditoria. "
        "Converse em português do Brasil de forma natural, humana, direta e acolhedora — como um diálogo entre duas pessoas, "
        "sem respostas engessadas ou listas gigantes desnecessárias.\n"
        "REGRAS RÍGIDAS DE ESCOPO (NUNCA QUEBRE):\n"
        "1. Responda apenas dúvidas conceituais sobre INSS, significado de siglas/indicadores do CNIS (ex.: PREC-MENOR-MIN, PEXT, IEAN), "
        "regras gerais (tempo comum, especial, PCD LC 142/2013) e como baixar o extrato no Meu INSS.\n"
        "2. NUNCA faça levantamento de vínculos, contagem de tempo de contribuição, análise de trechos do extrato do usuário "
        "nem cálculo/estimativa de valor de benefício (RMI) — esses são serviços exclusivos do motor pago do site.\n"
        "3. Se o usuário pedir levantamento, análise do próprio CNIS ou cálculo de benefício, recuse com educação e oriente "
        "a enviar o PDF do CNIS no campo de upload da página e adquirir o Dossiê Completo ou Passe na Área do Assinante.\n"
        "4. Nunca substitua aconselhamento jurídico formal nem o atendimento oficial do INSS."
    )

    contexto_conversa = ""
    if historico:
        ultimas = historico[-6:]
        linhas = []
        for m in ultimas:
            papel = "Usuário" if m.get("role") == "user" else "Gemini"
            conteudo = str(m.get("content", "")).strip()
            if conteudo:
                linhas.append(f"{papel}: {conteudo}")
        if linhas:
            contexto_conversa = "Histórico recente da conversa:\n" + "\n".join(linhas) + "\n\n"

    prompt = f"{instrucoes}\n\n{contexto_conversa}Usuário: {pergunta}\nGemini:"

    # 1) Plano A: SDK google-genai
    try:
        from google import genai

        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=prompt,
        )
        texto = getattr(response, "text", None)
        if texto and texto.strip():
            return texto.strip()
    except Exception:
        pass

    # 2) Plano B: chamada REST direta com x-goog-api-key
    url = (
        "https://generativelanguage.googleapis.com/v1beta/"
        "models/gemini-3.8-flash:generateContent"
    )
    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt}
                ]
            }
        ]
    }
    req = urllib.request.Request(
        url=url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "x-goog-api-key": api_key,
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            dados = json.loads(resp.read().decode("utf-8"))
            candidatos = dados.get("candidates", [])
            if candidatos:
                partes = candidatos[0].get("content", {}).get("parts", [])
                texto_rest = "".join(p.get("text", "") for p in partes if "text" in p).strip()
                if texto_rest:
                    return texto_rest
        return "⚠️ A API do Gemini não devolveu texto na resposta."
    except urllib.error.HTTPError as e:
        detalhe = e.read().decode("utf-8", errors="ignore")
        return f"⚠️ Erro HTTP {e.code} na Gemini API: {detalhe}"
    except Exception as e:
        return f"⚠️ Falha ao consultar a Gemini API: {e}"

LOGO_SITE = _achar_logo_oficial()

st.set_page_config(
    page_title="Inteligência Previdenciária & Auditoria - Análise Detalhada do CNIS",
    page_icon=LOGO_SITE if LOGO_SITE else "⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

URL_APP = "https://auditoria-previdenciaria-4ibrpwssuthl96ga3kh974.streamlit.app"

query_params = st.query_params
access_token_url = query_params.get("access_token")
refresh_token_url = query_params.get("refresh_token")
token_hash = query_params.get("token_hash")
tipo_param = (query_params.get("type") or "").lower()
codigo_auth = query_params.get("code")


def _guardar_sessao_recuperacao(res):
    sess_obj = getattr(res, "session", None)
    user_obj = getattr(res, "user", None)
    if sess_obj:
        st.session_state["user_session"] = sess_obj
        acc_tok = getattr(sess_obj, "access_token", None)
        ref_tok = getattr(sess_obj, "refresh_token", None)
        if isinstance(sess_obj, dict):
            acc_tok = acc_tok or sess_obj.get("access_token")
            ref_tok = ref_tok or sess_obj.get("refresh_token")
        if acc_tok:
            st.session_state["recovery_access_token"] = acc_tok
        if ref_tok:
            st.session_state["recovery_refresh_token"] = ref_tok
    if user_obj:
        st.session_state["user_email"] = getattr(user_obj, "email", st.session_state.get("user_email", ""))
        st.session_state["user_id"] = getattr(user_obj, "id", st.session_state.get("user_id", ""))
    st.session_state["modo_redefinir_senha"] = True


if access_token_url and (tipo_param == "recovery" or not tipo_param):
    try:
        client = get_supabase()
        res = client.auth.set_session(access_token_url, refresh_token_url or "")
        st.session_state["recovery_access_token"] = access_token_url
        if refresh_token_url:
            st.session_state["recovery_refresh_token"] = refresh_token_url
        _guardar_sessao_recuperacao(res)
        st.query_params.clear()
        st.rerun()
    except Exception:
        st.query_params.clear()
elif token_hash and (tipo_param == "recovery" or not tipo_param):
    try:
        client = get_supabase()
        res = client.auth.verify_otp({"token_hash": token_hash, "type": "recovery"})
        _guardar_sessao_recuperacao(res)
        st.query_params.clear()
        st.rerun()
    except Exception:
        st.query_params.clear()
elif codigo_auth:
    try:
        client = get_supabase()
        res = client.auth.exchange_code_for_session({"auth_code": codigo_auth})
        _guardar_sessao_recuperacao(res)
        st.query_params.clear()
        st.rerun()
    except Exception:
        st.query_params.clear()

if st.session_state.get("modo_redefinir_senha"):
    st.subheader("🔑 Criar Nova Senha")
    st.caption("Sua conta e seu histórico de pagamentos/créditos permanecem preservados.")

    with st.form("form_criar_nova_senha", clear_on_submit=False):
        nova_senha = st.text_input(
            "Digite sua nova senha (mínimo 6 caracteres)",
            type="password",
            key="nova_senha_input",
        )
        confirmar_senha = st.text_input(
            "Confirme sua nova senha",
            type="password",
            key="conf_senha_input",
        )
        enviou_nova_senha = st.form_submit_button("Salvar Nova Senha", type="primary")

    if enviou_nova_senha:
        senha_limpa = (nova_senha or "").strip()
        conf_limpa = (confirmar_senha or "").strip()
        if len(senha_limpa) < 6:
            st.warning("A senha deve ter pelo menos 6 caracteres.")
        elif senha_limpa != conf_limpa:
            st.warning("As senhas não coincidem.")
        else:
            acc_tok = st.session_state.get("recovery_access_token")
            ref_tok = st.session_state.get("recovery_refresh_token", "")
            if acc_tok:
                try:
                    get_supabase().auth.set_session(acc_tok, ref_tok)
                except Exception:
                    pass
            sucesso, msg = atualizar_nova_senha(senha_limpa)
            if sucesso:
                st.success(msg)
                st.session_state["modo_redefinir_senha"] = False
                st.session_state.pop("recovery_access_token", None)
                st.session_state.pop("recovery_refresh_token", None)
                st.rerun()
            else:
                st.error(msg)

    if st.button("Voltar ao início", key="btn_voltar_redefinir"):
        st.session_state["modo_redefinir_senha"] = False
        st.session_state.pop("recovery_access_token", None)
        st.session_state.pop("recovery_refresh_token", None)
        st.rerun()
    st.stop()


def _renderizar_bloco_esqueci_senha(sufixo_key: str):
    url_base = globals().get("URL_APP", "https://auditoria-previdenciaria-4ibrpwssuthl96ga3kh974.streamlit.app")
    with st.expander("Esqueci minha senha", expanded=False):
        email_rec = st.text_input(
            "Digite o e-mail cadastrado",
            placeholder="seu@email.com",
            key=f"email_recuperacao_{sufixo_key}",
        )
        if st.button("Enviar código de recuperação", use_container_width=True, key=f"btn_enviar_recuperacao_{sufixo_key}"):
            if email_rec.strip():
                ok, mensagem = enviar_email_recuperacao(email_rec.strip(), url_base)
                if ok:
                    st.success(mensagem)
                else:
                    st.error(mensagem)
            else:
                st.warning("Por favor, digite seu e-mail.")

        st.markdown("---")
        codigo_otp = st.text_input(
            "Código recebido por e-mail",
            placeholder="Digite o código recebido",
            key=f"codigo_otp_{sufixo_key}",
        )
        if st.button("Validar código e trocar senha", use_container_width=True, key=f"btn_validar_otp_{sufixo_key}"):
            ok_otp, msg_otp = verificar_codigo_otp_recuperacao(email_rec.strip(), codigo_otp.strip())
            if ok_otp:
                st.session_state["modo_redefinir_senha"] = True
                st.rerun()
            else:
                st.error(msg_otp)


PLANOS = {
    "AVULSO": {
        "titulo": "Dossiê Avulso - Auditoria CNIS",
        "valor": 19.90,
    },
    "PASSE_7D": {
        "titulo": "Passe 7 Dias - Auditoria CNIS",
        "valor": 39.90,
    },
    "PASSE_30D": {
        "titulo": "Passe 30 Dias - Auditoria CNIS",
        "valor": 69.90,
    },
}


def criar_preferencia_checkout(plano_id: str, email_usuario: str) -> str:
    """Cria a preferência no Mercado Pago Checkout Pro com back_urls configuradas
    e retorna a URL de pagamento (init_point)."""
    import mercadopago

    if plano_id not in PLANOS:
        raise ValueError(f"Plano inválido: {plano_id}")

    plano_info = PLANOS[plano_id]
    valor = plano_info["valor"]
    titulo = plano_info["titulo"]

    access_token = st.secrets["MERCADOPAGO_ACCESS_TOKEN"]
    base_url = st.secrets.get("APP_BASE_URL", "https://inteligenciapreviden.streamlit.app").rstrip("/")

    sdk = mercadopago.SDK(access_token)

    url_sucesso = f"{base_url}/?pagamento=aprovado&plano={plano_id}&valor={valor:.2f}"
    url_falha = f"{base_url}/?pagamento=falha&plano={plano_id}"
    url_pendente = f"{base_url}/?pagamento=pendente&plano={plano_id}"

    preference_data = {
        "items": [
            {
                "id": plano_id,
                "title": titulo,
                "quantity": 1,
                "currency_id": "BRL",
                "unit_price": valor,
            }
        ],
        "payer": {
            "email": email_usuario,
        },
        "back_urls": {
            "success": url_sucesso,
            "failure": url_falha,
            "pending": url_pendente,
        },
        "auto_return": "approved",
        "external_reference": f"{email_usuario}|{plano_id}",
    }

    resposta = sdk.preference().create(preference_data)
    return resposta["response"]["init_point"]


_secrets_acesso = st.secrets.get("codigos_acesso", {})
CODIGOS_DOSSIE_AVULSO = {c.strip().upper() for c in _secrets_acesso.get("dossie_avulso", [])}
CODIGOS_PASSE_ILIMITADO = {c.strip().upper() for c in _secrets_acesso.get("passe_ilimitado", [])}
CODIGOS_DESCONTO_50 = {c.strip().upper() for c in _secrets_acesso.get("desconto_50", [])}
CODIGOS_BRINDE_GRATIS = {c.strip().upper() for c in _secrets_acesso.get("brinde_gratis", [])}
TODOS_CODIGOS_PAGOS = CODIGOS_DOSSIE_AVULSO | CODIGOS_PASSE_ILIMITADO

acesso_url = st.query_params.get("acesso", "").strip().upper()
pagina_url = st.query_params.get("pagina", "").strip().lower()
status_pgto = st.query_params.get("pagamento", "").strip().lower()
status_mp = st.query_params.get("status", "").strip().lower()
plano_url = st.query_params.get("plano", "").strip().upper()
payment_id_url = (st.query_params.get("payment_id") or st.query_params.get("collection_id") or "").strip()

retorno_mp_aprovado = (status_pgto == "aprovado") or (status_mp == "approved")
is_conversao_paga = (pagina_url == "obrigado") or (acesso_url in TODOS_CODIGOS_PAGOS) or retorno_mp_aprovado

valor_conversao = 0.0
if is_conversao_paga:
    valor_padrao_plano = PLANOS.get(plano_url, {}).get("valor", 19.90)
    try:
        valor_conversao = float(st.query_params.get("valor", valor_padrao_plano))
    except (ValueError, TypeError):
        valor_conversao = valor_padrao_plano

chave_conv_sessao = f"conv_gads_disparada_{payment_id_url or plano_url or pagina_url or acesso_url}"
script_evento_compra = ""
if is_conversao_paga and not st.session_state.get(chave_conv_sessao):
    st.session_state[chave_conv_sessao] = True
    script_evento_compra = f"""
      window.parent.gtag('event', 'conversion', {{
          'send_to': 'AW-18585331316/Fwq0CLj10L0dEPTMlp5F',
          'value': {valor_conversao:.2f},
          'currency': 'BRL',
          'transaction_id': '{payment_id_url}'
      }});
    """

components.html(f"""
<script>
  try {{
    var parentDoc = window.parent.document;
    if (!parentDoc.getElementById('gtm-script-5296H925')) {{
      var s = parentDoc.createElement('script');
      s.id = 'gtm-script-5296H925';
      s.async = true;
      s.src = 'https://googletagmanager.com/gtm.js?id=GTM-5296H925';
      parentDoc.head.appendChild(s);
      window.parent.dataLayer = window.parent.dataLayer || [];
      window.parent.gtag = function(){{window.parent.dataLayer.push(arguments);}};
      window.parent.gtag('js', new Date());
      window.parent.gtag('config', 'AW-18585331316');
    }}
    if (window.parent.gtag) {{
      {script_evento_compra}
    }}
  }} catch (e) {{
    console.log(e);
  }}
</script>
""", height=0, width=0)

if retorno_mp_aprovado:
    nome_plano_retornado = PLANOS.get(plano_url, {}).get("titulo", plano_url or "Dossiê Completo")
    st.query_params.clear()
    st.success(f"✅ Pagamento aprovado! Seu acesso ({nome_plano_retornado} — R$ {valor_conversao:.2f}) já foi confirmado. Clique em '🔄 Atualizar' na Área do Assinante se necessário.")
elif status_pgto == "pendente" or status_mp == "pending":
    st.query_params.clear()
    st.info("⏳ Seu pagamento está em processamento. Assim que for confirmado, seu acesso será liberado automaticamente.")
elif status_pgto == "falha" or status_mp == "rejected":
    st.query_params.clear()
    st.warning("⚠️ O pagamento não foi concluído. Você pode tentar novamente quando quiser.")

if pagina_url == "obrigado":
    cod_vip = acesso_url
    plano_txt = "Passe Ilimitado VIP" if cod_vip in CODIGOS_PASSE_ILIMITADO else "Dossiê Completo Avulso"
    st.markdown(f"""
    <div style="max-width:620px; margin:2.5rem auto; background:linear-gradient(135deg,#0F293A 0%,#1B3B52 100%); border:2px solid #C59B27; border-radius:14px; padding:2.2rem 2rem; text-align:center; color:#FFFFFF; box-shadow:0 10px 25px rgba(15,41,58,0.25);">
        <div style="display:inline-block; background:#25D366; color:#0F293A; font-weight:800; font-size:0.78rem; padding:4px 14px; border-radius:20px; text-transform:uppercase; margin-bottom:1rem;">✅ Pagamento Confirmado com Sucesso</div>
        <div style="font-size:1.75rem; font-weight:800; color:#F3D06B; margin-bottom:0.6rem;">Obrigado por escolher a Inteligência Previdenciária!</div>
        <p style="font-size:1rem; color:#E2E8F0; line-height:1.55; margin-bottom:1.4rem;">
            Seu acesso ao <strong>{plano_txt}</strong> já está liberado.<br>
            Clique no botão dourado abaixo para acessar o sistema com todos os cálculos e downloads desbloqueados.
        </p>
        <a href="/" target="_self" style="display:inline-block; background:#C59B27; color:#0F293A !important; font-weight:800; font-size:1.05rem; padding:0.85rem 2rem; border-radius:8px; text-decoration:none !important;">
            🚀 Acessar Meu Dossiê Completo Agora
        </a>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

_links_secrets = st.secrets.get("links_pagamento", {})
LINK_MP_AVULSO_19 = _links_secrets.get("avulso", "https://mpago.la/1XEo7nL")
LINK_MP_DESCONTO_50 = LINK_MP_AVULSO_19
LINK_MP_PASSE_7D = _links_secrets.get("passe_7d", "https://mpago.la/2aGyJn9")
LINK_MP_PASSE_30D = _links_secrets.get("passe_30d", "https://mpago.la/2iXqtct")
EMAIL_SUPORTE = "inteligenciaprevidenciariaaudi@gmail.com"

if "plano_ativo" not in st.session_state:
    st.session_state["plano_ativo"] = "GRATUITO"
if "desconto_50" not in st.session_state:
    st.session_state["desconto_50"] = False
if "aviso_cadastro_obrigatorio" not in st.session_state:
    st.session_state["aviso_cadastro_obrigatorio"] = False
if "gaveta_ia_aberta" not in st.session_state:
    st.session_state["gaveta_ia_aberta"] = False
if "mensagens_ia" not in st.session_state:
    st.session_state["mensagens_ia"] = [
        {
            "role": "assistant",
            "content": "✦ **Gemini** aqui! Posso conversar sobre dúvidas de **CNIS**, pendências no INSS, tempo concomitante, aposentadoria especial ou regra PCD (LC 142/2013). *(O cálculo oficial do seu extrato continua no motor da página.)*"
        }
    ]

if acesso_url in CODIGOS_PASSE_ILIMITADO:
    st.session_state["plano_ativo"] = "ILIMITADO"
elif acesso_url in CODIGOS_DOSSIE_AVULSO:
    st.session_state["plano_ativo"] = "DOSSIE"
elif acesso_url in CODIGOS_BRINDE_GRATIS:
    st.session_state["plano_ativo"] = "BRINDE"
elif acesso_url in CODIGOS_DESCONTO_50:
    st.session_state["desconto_50"] = True

st.markdown("""
<style>
    @media (min-width: 769px) {
        .st-key-menu_celular {
            display: none !important;
        }
    }
    #MainMenu {visibility: hidden;}
    [data-testid="stToolbar"] {display: none !important;}
    [data-testid="stDecoration"] {display: none !important;}
    [data-testid="stStatusWidget"] {display: none !important;}
    [class^="viewerBadge_"] {display: none !important;}
    div[class*="StatusWidget"] {display: none !important;}

    footer {visibility: hidden;}
    header {visibility: hidden;}

    div[data-testid="stColumn"]:nth-of-type(3) a[data-testid="stBaseLinkButton-secondary"] {
        background-color: #25D366 !important;
        color: #FFFFFF !important;
        border: 1px solid #1EBE5D !important;
        font-weight: 700 !important;
    }
    div[data-testid="stColumn"]:nth-of-type(3) a[data-testid="stBaseLinkButton-secondary"]:hover {
        background-color: #1EBE5D !important;
    }

    .block-container {
        padding-top: 1.1rem !important;
        padding-bottom: 2.5rem !important;
        max-width: 96% !important;
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0F293A 0%, #16364D 100%) !important;
        border-right: 2px solid #C59B27 !important;
    }

    @media (min-width: 769px) {
        section[data-testid="stSidebar"] {
            width: 315px !important;
            min-width: 315px !important;
            max-width: 315px !important;
            transform: none !important;
            visibility: visible !important;
        }
        section[data-testid="stSidebar"] > div {
            width: 315px !important;
            padding-top: 1.2rem !important;
        }
        [data-testid="stSidebarCollapseButton"],
        [data-testid="collapsedControl"] {
            display: none !important;
        }
    }

    @media (max-width: 768px) {
        [data-testid="stSidebarCollapseButton"],
        [data-testid="collapsedControl"] {
            display: none !important;
        }
        .block-container {
            padding-top: 0.55rem !important;
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
            padding-bottom: 7.5rem !important;
            max-width: 100% !important;
        }
    }

    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
    [data-testid="stSidebar"] [data-testid="stRadio"] label,
    [data-testid="stSidebar"] [data-testid="stRadio"] p,
    [data-testid="stSidebar"] [data-testid="stRadio"] span,
    [data-testid="stSidebar"] [data-testid="stCheckbox"] label,
    [data-testid="stSidebar"] [data-testid="stCheckbox"] p,
    [data-testid="stSidebar"] [data-testid="stCheckbox"] span {
        color: #FFFFFF !important;
        font-weight: 500 !important;
    }

    [data-testid="stSidebar"] [data-testid="stCaptionContainer"] p,
    [data-testid="stSidebar"] small {
        color: #E2E8F0 !important;
        opacity: 1 !important;
    }

    [data-testid="stSidebar"] button[kind="primary"] {
        background-color: #0F293A !important;
        border: 1.5px solid #C59B27 !important;
        color: #FFFFFF !important;
    }

    [data-testid="stSidebar"] [data-testid="stExpander"] summary,
    [data-testid="stSidebar"] [data-testid="stExpander"] summary p,
    [data-testid="stSidebar"] [data-testid="stExpander"] summary span,
    [data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stWidgetLabel"] p,
    [data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stCaptionContainer"] p,
    [data-testid="stSidebar"] [data-testid="stExpander"] label,
    [data-testid="stSidebar"] [data-testid="stExpander"] p {
        color: #0F293A !important;
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
    .sidebar-title-gold {
        font-size: 0.85rem;
        font-weight: 700;
        color: #F3D06B !important;
        margin-bottom: 0.3rem;
        text-transform: uppercase;
        letter-spacing: 0.4px;
    }
    .sidebar-title {
        font-size: 0.85rem;
        font-weight: 800;
        color: #0F293A !important;
        margin-bottom: 0.2rem;
        text-transform: uppercase;
        letter-spacing: 0.4px;
    }
    .security-badge {
        background: linear-gradient(180deg, #FFFDF5 0%, #FFFFFF 100%);
        border: 1.5px solid rgba(197, 155, 39, 0.55);
        border-left: 5px solid #C59B27;
        color: #0F293A !important;
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
        display: flex;
        align-items: flex-start;
        justify-content: flex-start;
        gap: 1rem;
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
        line-height: 1.25;
        overflow-wrap: break-word;
        word-break: normal;
    }
    .hero-subtitle {
        font-size: 0.88rem;
        color: #CBD5E1;
        margin: 0;
        line-height: 1.4;
        overflow-wrap: break-word;
        word-break: normal;
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
        margin-bottom: 0.5rem;
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
    .save-warning-box {
        background: linear-gradient(135deg, #0F293A 0%, #16364D 100%);
        border: 2px solid #C59B27;
        border-left: 6px solid #F3D06B;
        border-radius: 12px;
        padding: 1.05rem 1.25rem;
        color: #FFFFFF;
        font-size: 0.88rem;
        line-height: 1.5;
        margin: 1rem 0 0.85rem 0;
        box-shadow: 0 6px 18px rgba(15, 41, 58, 0.16);
    }
    .save-warning-badge {
        display: inline-block;
        background: #C59B27;
        color: #0F293A;
        font-size: 0.7rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        padding: 2px 9px;
        border-radius: 20px;
        margin-bottom: 0.45rem;
    }
    .save-warning-title {
        font-size: 1rem;
        font-weight: 800;
        color: #F3D06B;
        margin-bottom: 0.35rem;
    }
    .save-warning-support {
        margin-top: 0.65rem;
        padding-top: 0.6rem;
        border-top: 1px solid rgba(197, 155, 39, 0.35);
        font-size: 0.83rem;
        color: #F8FAFC;
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
        bottom: 72px;
        right: 22px;
        display: flex;
        flex-direction: column;
        align-items: flex-end;
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

    /* Caixa lateral interativa de Dúvidas INSS */
    .st-key-caixa_lateral_inss {
        position: fixed !important;
        top: 24px !important;
        right: 22px !important;
        width: 370px !important;
        max-width: calc(100vw - 24px) !important;
        max-height: calc(100vh - 48px) !important;
        overflow-y: auto !important;
        background: linear-gradient(180deg, #0F293A 0%, #15344A 100%) !important;
        border: 2px solid #C59B27 !important;
        border-radius: 14px !important;
        padding: 0.95rem 1rem !important;
        z-index: 10002 !important;
        box-shadow: 0 14px 34px rgba(15, 41, 58, 0.35) !important;
        color: #FFFFFF !important;
    }
    .st-key-caixa_lateral_inss [data-testid="stWidgetLabel"] p,
    .st-key-caixa_lateral_inss [data-testid="stCaptionContainer"] p,
    .st-key-caixa_lateral_inss p,
    .st-key-caixa_lateral_inss label {
        color: #F8FAFC !important;
    }
    .st-key-caixa_lateral_inss [data-testid="stChatMessage"] {
        background: rgba(255, 255, 255, 0.06) !important;
        border: 1px solid rgba(197, 155, 39, 0.28) !important;
        border-radius: 10px !important;
        padding: 0.5rem 0.65rem !important;
        margin-bottom: 0.45rem !important;
    }
    .st-key-caixa_lateral_inss [data-testid="stChatMessage"],
    .st-key-caixa_lateral_inss [data-testid="stChatMessage"] p,
    .st-key-caixa_lateral_inss [data-testid="stChatMessage"] li,
    .st-key-caixa_lateral_inss [data-testid="stChatMessage"] span,
    .st-key-caixa_lateral_inss [data-testid="stChatMessage"] strong,
    .st-key-caixa_lateral_inss [data-testid="stChatMessage"] em,
    .st-key-caixa_lateral_inss [data-testid="stChatMessage"] div,
    .st-key-caixa_lateral_inss [data-testid="stChatMessage"] ul,
    .st-key-caixa_lateral_inss [data-testid="stChatMessage"] ol {
        font-size: 0.84rem !important;
        line-height: 1.45 !important;
        color: #FFFFFF !important;
        opacity: 1 !important;
    }
    .st-key-caixa_lateral_inss [data-testid="stChatMessage"] li::marker {
        color: #F3D06B !important;
    }

    /* Pop-up de Resultados do Dossiê / Planos (hover + clique/foco) */
    .dossie-trigger-card {
        position: relative;
        cursor: pointer;
    }
    .dossie-popover-card {
        display: none;
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        width: 390px;
        max-width: calc(100vw - 24px);
        max-height: 88vh;
        overflow-y: auto;
        background: linear-gradient(180deg, #0F293A 0%, #15344A 100%);
        border: 2px solid #C59B27;
        border-radius: 14px;
        padding: 0.9rem 1rem;
        color: #FFFFFF;
        box-shadow: 0 18px 44px rgba(15, 41, 58, 0.55);
        z-index: 10025;
        text-align: left;
    }
    .dossie-trigger-card:hover .dossie-popover-card,
    .dossie-trigger-card:focus-within .dossie-popover-card {
        display: block;
    }
    .dossie-pill-hint {
        display: inline-block;
        font-size: 0.72rem;
        font-weight: 700;
        color: #8A6810;
        background: #FFFDF5;
        border: 1px solid #C59B27;
        border-radius: 20px;
        padding: 2px 9px;
        margin-bottom: 0.45rem;
    }
    .plan-card.featured .dossie-pill-hint {
        color: #F3D06B;
        background: rgba(197, 155, 39, 0.16);
        border-color: rgba(197, 155, 39, 0.55);
    }

    /* Pop-up de Segurança & LGPD (hover + clique/foco) */
    .lgpd-trigger-wrap {
        position: relative;
        display: block;
    }
    .lgpd-trigger-inline {
        position: relative;
        display: inline-block;
    }
    .lgpd-popover-card {
        display: none;
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        width: 390px;
        max-width: calc(100vw - 24px);
        max-height: 88vh;
        overflow-y: auto;
        background: linear-gradient(180deg, #0F293A 0%, #15344A 100%);
        border: 2px solid #C59B27;
        border-radius: 14px;
        padding: 0.9rem 1rem;
        color: #FFFFFF;
        box-shadow: 0 18px 44px rgba(15, 41, 58, 0.55);
        z-index: 10020;
    }
    .lgpd-trigger-wrap:hover .lgpd-popover-card,
    .lgpd-trigger-wrap:focus-within .lgpd-popover-card,
    .lgpd-trigger-inline:hover .lgpd-popover-card,
    .lgpd-trigger-inline:focus-within .lgpd-popover-card {
        display: block;
    }
    .lgpd-pop-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
        border-bottom: 1px solid rgba(197, 155, 39, 0.45);
        padding-bottom: 0.45rem;
        margin-bottom: 0.55rem;
    }
    .lgpd-pop-title {
        font-size: 0.88rem;
        font-weight: 800;
        color: #F3D06B;
    }
    .lgpd-pop-hint {
        font-size: 0.72rem;
        color: #CBD5E1;
    }
    .lgpd-pop-body {
        font-size: 0.8rem;
        color: #F8FAFC;
        line-height: 1.45;
    }
    .lgpd-pop-img {
        width: 100%;
        height: auto;
        border-radius: 10px;
        border: 1.5px solid #C59B27;
        margin-top: 0.45rem;
        display: block;
    }
    .lgpd-upload-header {
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 8px;
        margin: 0.35rem 0 0.45rem 0;
    }
    .lgpd-upload-title {
        font-size: 1.12rem;
        font-weight: 700;
        color: #0F293A;
        margin: 0;
    }
    .lgpd-pill-trigger {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        background: #FFFDF5;
        border: 1.5px solid #C59B27;
        color: #0F293A;
        font-size: 0.76rem;
        font-weight: 700;
        padding: 3px 10px;
        border-radius: 20px;
        cursor: pointer;
    }
    .st-key-caixa_lateral_inss button[kind="secondary"],
    .st-key-caixa_lateral_inss [data-testid="stBaseButton-secondary"] {
        background: #F8E7B2 !important;
        color: #0F293A !important;
        border: 1.5px solid #C59B27 !important;
        font-weight: 800 !important;
        border-radius: 8px !important;
    }
    .st-key-caixa_lateral_inss button[kind="secondary"] p,
    .st-key-caixa_lateral_inss button[kind="secondary"] span,
    .st-key-caixa_lateral_inss button[kind="secondary"] div,
    .st-key-caixa_lateral_inss [data-testid="stBaseButton-secondary"] p,
    .st-key-caixa_lateral_inss [data-testid="stBaseButton-secondary"] span,
    .st-key-caixa_lateral_inss [data-testid="stBaseButton-secondary"] div {
        color: #0F293A !important;
        font-weight: 800 !important;
        opacity: 1 !important;
    }
    .st-key-caixa_lateral_inss button[kind="secondary"]:hover,
    .st-key-caixa_lateral_inss [data-testid="stBaseButton-secondary"]:hover {
        background: #C59B27 !important;
        color: #0F293A !important;
    }
    .st-key-caixa_lateral_inss input {
        color: #0F293A !important;
        background-color: #FFFFFF !important;
        font-weight: 600 !important;
    }

    @media (max-width: 768px) {
        .hero-box {
            flex-direction: column;
            align-items: center;
            text-align: center !important;
            padding: 0.95rem 0.9rem !important;
            gap: 0.55rem !important;
        }
        .hero-box > div {
            text-align: center !important;
            width: 100%;
        }
        .hero-logo {
            width: 58px !important;
            height: 58px !important;
        }
        .hero-badge {
            font-size: 0.66rem !important;
        }
        .hero-title {
            font-size: 1.18rem !important;
            line-height: 1.24 !important;
            letter-spacing: -0.2px !important;
        }
        .hero-subtitle {
            font-size: 0.81rem !important;
            line-height: 1.38 !important;
        }
        [data-testid="stExpander"] summary {
            font-size: 0.84rem !important;
            padding: 0.5rem 0.75rem !important;
            line-height: 1.3 !important;
        }
        .pricing-grid,
        .feature-grid {
            grid-template-columns: 1fr !important;
            gap: 0.7rem !important;
        }
        .metric-card {
            padding: 0.9rem 1rem !important;
        }
        .metric-value {
            font-size: 1.25rem !important;
        }
        .upsell-banner {
            padding: 0.95rem 1rem !important;
        }
        .upsell-banner .plan-cta {
            display: block !important;
            width: 100% !important;
            box-sizing: border-box !important;
        }
        .stButton > button,
        .stDownloadButton > button,
        a[data-testid="stBaseLinkButton-secondary"] {
            width: 100% !important;
        }
        [data-testid="stTabs"] [role="tablist"] {
            overflow-x: auto !important;
            flex-wrap: nowrap !important;
            -webkit-overflow-scrolling: touch;
        }
        [data-testid="stTabs"] button[role="tab"] {
            white-space: nowrap !important;
            font-size: 0.8rem !important;
            padding: 0.45rem 0.65rem !important;
        }
        .floating-container {
            bottom: 64px !important;
            right: 10px !important;
            gap: 6px !important;
        }
        .float-btn {
            padding: 6px 11px !important;
            font-size: 0.73rem !important;
            gap: 5px !important;
        }
        .support-menu {
            width: 220px !important;
            bottom: 40px !important;
        }
        .st-key-caixa_lateral_inss {
            top: 12px !important;
            right: 10px !important;
            left: 10px !important;
            width: auto !important;
            max-width: none !important;
            max-height: calc(100vh - 24px) !important;
        }
    }
    @media (min-width: 769px) {
        .apenas-celular {
            display: none !important;
        }
    }
</style>
""", unsafe_allow_html=True)

_logo_b64 = ""
if LOGO_SITE and os.path.exists(LOGO_SITE):
    with open(LOGO_SITE, "rb") as _f_logo:
        _logo_b64 = base64.b64encode(_f_logo.read()).decode("utf-8")

_prova_seg_b64 = _carregar_prova_seguranca_b64()
_img_prova_seg_html = (
    f'<img src="data:image/png;base64,{_prova_seg_b64}" class="lgpd-pop-img" alt="Verificação de segurança e LGPD no Gemini" />'
    if _prova_seg_b64
    else (
        '<div style="background:rgba(255,255,255,0.07); border:1px solid rgba(197,155,39,0.4); border-radius:10px; padding:0.75rem; margin-top:0.5rem;">'
        '<div style="color:#F3D06B; font-weight:700; margin-bottom:0.35rem;">✦ Resposta do Gemini sobre a segurança do CNIS:</div>'
        '<div style="color:#FFFFFF; font-size:0.79rem; line-height:1.45;">'
        '"Pode confiar sim! A plataforma foi desenvolvida com protocolos modernos de segurança e criptografia, seguindo rigorosamente a Lei Geral de Proteção de Dados (LGPD). '
        'Os dados do seu arquivo são processados em memória com sigilo absoluto e utilizados exclusivamente para gerar o seu relatório."'
        '</div></div>'
    )
)

POPUP_SEGURANCA_HTML = (
    '<div class="lgpd-popover-card">'
    '<div class="lgpd-pop-header">'
    '<span class="lgpd-pop-title">🛡️ Sigilo Absoluto &amp; Conformidade LGPD</span>'
    '<span class="lgpd-pop-hint">Passe fora ou toque fora para fechar</span>'
    '</div>'
    '<div class="lgpd-pop-body">'
    'Seu PDF do CNIS é lido <strong>100% em memória</strong> apenas durante a auditoria e <strong>nunca fica gravado em disco</strong>.'
    + _img_prova_seg_html
    + '</div>'
    '</div>'
)

_prova_dossie_b64 = _carregar_prova_dossie_b64()
_img_prova_dossie_html = (
    f'<img src="data:image/jpeg;base64,{_prova_dossie_b64}" class="lgpd-pop-img" alt="Como o Dossiê Completo ajuda na sua aposentadoria" />'
    if _prova_dossie_b64
    else (
        '<div style="background:rgba(255,255,255,0.07); border:1px solid rgba(197,155,39,0.4); border-radius:10px; padding:0.75rem; margin-top:0.5rem;">'
        '<div style="color:#F3D06B; font-weight:700; margin-bottom:0.35rem;">✦ Como o Dossiê ajuda na sua aposentadoria:</div>'
        '<div style="color:#FFFFFF; font-size:0.79rem; line-height:1.45;">'
        'O Dossiê funciona como um <strong>raio-X completo</strong> da sua vida profissional: revela pendências escondidas no CNIS, '
        'mostra o que está regular, onde estão as travas e quais comprovantes reunir antes de pedir o benefício no INSS.'
        '</div></div>'
    )
)

POPUP_DOSSIE_HTML = (
    '<div class="dossie-popover-card">'
    '<div class="lgpd-pop-header">'
    '<span class="lgpd-pop-title">📋 Por que o Dossiê faz diferença no INSS</span>'
    '<span class="lgpd-pop-hint">Passe fora ou toque fora para fechar</span>'
    '</div>'
    '<div class="lgpd-pop-body">'
    'Veja como a auditoria completa evita exigências demoradas e benefício concedido com valor menor:'
    + _img_prova_dossie_html
    + '</div>'
    '</div>'
)


_img_hero_html = (
    f'<img src="data:image/png;base64,{_logo_b64}" class="hero-logo" style="width:76px; height:76px; border-radius:10px; border:1.5px solid #C59B27; flex-shrink:0;" />'
    if _logo_b64 else ""
)

st.markdown(f"""
<div id="topo-pagina" class="hero-box">
    {_img_hero_html}
    <div>
        <div class="hero-badge">Inteligência Previdenciária & Auditoria</div>
        <h1 class="hero-title">Diagnóstico Estratégico de Aposentadoria e CNIS</h1>
        <p class="hero-subtitle">Análise automatizada de vínculos, auditoria de pendências no INSS, enquadramento PCD (LC 142/2013) e projeção de renda.</p>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div style="background:linear-gradient(90deg, #FFFDF5 0%, #FFFFFF 100%); border:1.5px solid #C59B27; border-left:5px solid #0F293A; border-radius:10px; padding:0.75rem 1.1rem; margin-bottom:0.75rem; font-size:0.86rem; color:#0F293A; line-height:1.45;">
    <strong>🛡️ Por que não usar o ChatGPT ou IAs genéricas para ler seu CNIS?</strong><br>
    Ferramentas genéricas de texto costumam errar datas, ignorar indicadores ocultos do INSS (como <em>PREC-MENOR-MIN</em>, <em>IEAN</em> e <em>PEXT</em>) e somar períodos concomitantes de forma incorreta. Aqui, cada vínculo e remuneração passa por <strong>regras determinísticas da legislação previdenciária</strong> com sigilo total (LGPD).
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

    acesso_supabase = renderizar_painel_acesso(chave_prefixo="side")
    _renderizar_bloco_esqueci_senha("side")

    st.markdown("---")
    st.markdown("### ⚙️ Configure seu Perfil")
    sexo = st.selectbox("Sexo biológico (Regra INSS):", options=["M", "F"], format_func=lambda x: "Masculino" if x == "M" else "Feminino")

    st.markdown('<div class="sidebar-card"><div class="sidebar-title">1. Ambiente de Trabalho</div></div>', unsafe_allow_html=True)
    ambiente = st.radio(
        "Selecione sua atuação predominante:",
        options=["Administrativo / Escritório", "Operacional / Área de Risco (Insalubridade/Periculosidade)"],
        index=0,
        label_visibility="collapsed"
    )
    ambiente_operacional = (ambiente == "Operacional / Área de Risco (Insalubridade/Periculosidade)")

    st.markdown('<div class="sidebar-card"><div class="sidebar-title">2. Enquadramento PCD / TEA</div></div>', unsafe_allow_html=True)
    possui_pcd = st.checkbox("Possui deficiência / TEA (LC 142/2013)?", value=True)
    st.caption("Marque se possui laudo médico de deficiência física, sensorial, intelectual ou TEA.")

    with st.expander("🔑 Possui Cupom ou Código VIP?"):
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

    _html_lgpd_side = (
        '<div class="lgpd-trigger-wrap" tabindex="0">'
        '<div class="security-badge" style="cursor:pointer;">'
        '<strong>🔒 Privacidade &amp; Segurança LGPD</strong> '
        '<span style="font-size:0.72rem; color:#8A6810; font-weight:700;">(ver detalhes)</span><br>'
        'Processamento 100% isolado em memória. Nenhum dado pessoal ou PDF fica gravado em disco.'
        '</div>'
        + POPUP_SEGURANCA_HTML
        + '</div>'
    )
    st.markdown(_html_lgpd_side, unsafe_allow_html=True)

modo_pago = acesso_supabase["tem_acesso_completo"] or (st.session_state["plano_ativo"] in ("DOSSIE", "ILIMITADO", "BRINDE"))
email_logado = acesso_supabase.get("email")

with st.container(key="menu_celular"):
    with st.expander("🔐 ENTRAR NA CONTA OU CONFIGURAR PERFIL (SEXO, AMBIENTE, PCD)"):
        acesso_mob = renderizar_painel_acesso(chave_prefixo="mob")
        _renderizar_bloco_esqueci_senha("mob")
        if acesso_mob["tem_acesso_completo"]:
            modo_pago = True
            email_logado = acesso_mob.get("email")
        elif acesso_mob.get("email") and not email_logado:
            email_logado = acesso_mob.get("email")

        st.markdown("---")
        sexo_mob = st.selectbox(
            "Sexo biológico (Regra INSS):",
            options=["M", "F"],
            index=0 if sexo == "M" else 1,
            format_func=lambda x: "Masculino" if x == "M" else "Feminino",
            key="mob_sexo"
        )
        if sexo_mob != sexo:
            sexo = sexo_mob

        st.markdown("**1. AMBIENTE DE TRABALHO**")
        amb_mob = st.radio(
            "Ambiente de Trabalho predominante:",
            options=["Administrativo / Escritório", "Operacional / Área de Risco (Insalubridade/Periculosidade)"],
            index=1 if ambiente_operacional else 0,
            label_visibility="collapsed",
            key="mob_amb"
        )
        ambiente_operacional = (amb_mob == "Operacional / Área de Risco (Insalubridade/Periculosidade)")

        st.markdown("**2. ENQUADRAMENTO PCD / TEA**")
        possui_pcd = st.checkbox("Possui deficiência / TEA (LC 142/2013)?", value=possui_pcd, key="mob_pcd")
        st.caption("Marque se possui laudo médico de deficiência física, sensorial, intelectual ou TEA.")

        st.markdown("**🔑 CUPOM OU CÓDIGO VIP**")
        cod_mob = st.text_input("Cupom ou Código VIP:", placeholder="Insira seu cupom aqui...", label_visibility="collapsed", key="mob_cod").strip().upper()
        if st.button("Ativar Código VIP", use_container_width=True, key="mob_btn"):
            if cod_mob in CODIGOS_PASSE_ILIMITADO:
                st.session_state["plano_ativo"] = "ILIMITADO"
                st.success("Passe Ilimitado ativado!")
                st.rerun()
            elif cod_mob in CODIGOS_DOSSIE_AVULSO:
                st.session_state["plano_ativo"] = "DOSSIE"
                st.success("Dossiê Completo ativado!")
                st.rerun()
            elif cod_mob in CODIGOS_BRINDE_GRATIS:
                st.session_state["plano_ativo"] = "BRINDE"
                st.success("🎁 Cortesia de 1 Dossiê Completo ativada!")
                st.rerun()
            elif cod_mob in CODIGOS_DESCONTO_50:
                st.session_state["desconto_50"] = True
                st.success("🎉 Cupom de 50% OFF aplicado!")
                st.rerun()
            else:
                st.error("Código ou cupom inválido.")

_html_lgpd_upload = (
    '<div class="lgpd-upload-header">'
    '<div class="lgpd-trigger-inline" tabindex="0" style="cursor:pointer;">'
    '<span class="lgpd-upload-title">📤 Envie seu Extrato Oficial (CNIS)</span>'
    '<span class="lgpd-pill-trigger">🔒 Sigilo LGPD</span>'
    + POPUP_SEGURANCA_HTML
    + '</div>'
    '</div>'
)
st.markdown(_html_lgpd_upload, unsafe_allow_html=True)
arquivo_pdf = st.file_uploader("Selecione ou arraste o arquivo PDF completo baixado do Meu INSS:", type=["pdf"], label_visibility="collapsed")
preco_avulso_txt = "R$ 9,95 (50% OFF)" if st.session_state["desconto_50"] else "R$ 19,90"

link_avulso_base = LINK_MP_DESCONTO_50 if st.session_state["desconto_50"] else LINK_MP_AVULSO_19
link_avulso_atual = gerar_link_checkout_personalizado(link_avulso_base, email_logado)
link_7d_atual = gerar_link_checkout_personalizado(LINK_MP_PASSE_7D, email_logado)
link_30d_atual = gerar_link_checkout_personalizado(LINK_MP_PASSE_30D, email_logado)

if arquivo_pdf is not None:
    label_btn = "🔍 Gerar Diagnóstico Gratuito (Prévia)" if not modo_pago else "🚀 Executar Auditoria e Gerar Dossiê Completo"
    if st.button(label_btn, type="primary"):
        with st.spinner("Processando vínculos, cruzando tetos históricos e auditando indicadores..."):
            pdf_em_memoria = io.BytesIO(arquivo_pdf.getvalue())
            excel_bruto_mem = io.BytesIO()

            parecer_texto = extrair_cnis_v5(
                pdf_em_memoria,
                excel_bruto_mem,
                sexo=sexo,
                ambiente_operacional=ambiente_operacional,
                possui_pcd=possui_pcd
            )

            excel_bruto_mem.seek(0)
            abas_brutas = pd.read_excel(excel_bruto_mem, sheet_name=None, skiprows=1)
            bytes_excel = aplicar_logo_todas_abas_excel(excel_bruto_mem)

            abas = {nome: df.dropna(how="all", axis=1) for nome, df in abas_brutas.items()}
            df_res = abas["2. ⏱️ TEMPO LÍQUIDO"]
            df_fin = abas["5. 💰 VALOR ESTIMADO (RMI)"]

            tmp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            caminho_pdf_relatorio = tmp_pdf.name
            tmp_pdf.close()

            try:
                criar_pdf_executivo(caminho_pdf_relatorio, df_res, df_fin, parecer_texto)
                with open(caminho_pdf_relatorio, "rb") as f_pdf:
                    bytes_pdf = f_pdf.read()
            finally:
                if os.path.exists(caminho_pdf_relatorio):
                    os.remove(caminho_pdf_relatorio)

            if modo_pago and acesso_supabase.get("tipo_acesso") == "AVULSO":
                registrar_uso_credito_avulso(acesso_supabase)

            st.session_state["resultado_auditoria"] = {
                "parecer_texto": parecer_texto,
                "abas": abas,
                "bytes_pdf": bytes_pdf,
                "bytes_excel": bytes_excel,
            }
else:
    st.session_state.pop("resultado_auditoria", None)

if "resultado_auditoria" in st.session_state:
    res = st.session_state["resultado_auditoria"]
    parecer_texto = res["parecer_texto"]
    abas = res["abas"]
    bytes_pdf = res["bytes_pdf"]
    bytes_excel = res["bytes_excel"]
    df_vinculos = abas["1. 📋 CNIS EXTRAÍDO"]
    df_res = abas["2. ⏱️ TEMPO LÍQUIDO"]
    df_fin = abas["5. 💰 VALOR ESTIMADO (RMI)"]
    tempo_liq = str(df_res.loc[df_res["Indicador"] == "TEMPO LÍQUIDO DE CONTRIBUIÇÃO (COMUM)", "Valor"].tolist().pop())
    rmi_comum = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Regra Comum"), "Valor_Estimado"].tolist().pop())
    rmi_pcd = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Aposentadoria PCD"), "Valor_Estimado"].tolist().pop())
    dif_mensal = str(df_fin.loc[df_fin["Metrica_Financeira"].str.contains("Diferença Mensal"), "Valor_Estimado"].tolist().pop())
    tempo_card = tempo_liq.split("(").pop().replace(")", "") if "(" in tempo_liq else tempo_liq

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
            f"   -> Cadastre-se primeiro em '🔐 Área do Assinante' (barra lateral esquerda no computador ou menu '🔐 ENTRAR NA CONTA OU CONFIGURAR PERFIL' no celular) para liberar o pagamento e o Dossiê Completo por {preco_avulso_txt}."
        )

        if email_logado:
            st.markdown(f"""
            <div class="upsell-banner">
                <h4 style="margin:0 0 0.3rem 0; color:#F3D06B;">⚠️ Encontramos {qtd_alertas} alerta(s) no seu CNIS — Desbloqueie agora e veja como corrigir + valor do benefício</h4>
                <p style="margin:0 0 0.6rem 0; font-size:0.88rem; color:#E2E8F0;">Conta conectada: <strong>{email_logado}</strong>. Ao concluir o pagamento via Pix ou Cartão, o <strong>Plano de Ação no INSS</strong>, o cálculo desde 1994 e os arquivos <strong>PDF + Excel</strong> serão liberados automaticamente para este e-mail.</p>
                <a href="{link_avulso_atual}" target="_blank" class="plan-cta cta-gold" style="display:inline-block; padding:0.55rem 1.3rem;">🔓 Desbloquear Dossiê Completo Agora — {preco_avulso_txt}</a>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="upsell-banner">
                <h4 style="margin:0 0 0.3rem 0; color:#F3D06B;">⚠️ Encontramos {qtd_alertas} alerta(s) no seu CNIS — Cadastre-se primeiro para liberar o pagamento</h4>
                <p style="margin:0 0 0.5rem 0; font-size:0.88rem; color:#E2E8F0;">Para garantir que o pagamento fique vinculado à sua conta e libere o <strong>Dossiê Completo ({preco_avulso_txt})</strong> automaticamente, <strong>você precisa criar sua conta gratuita antes de pagar</strong>:</p>
                <p style="margin:0 0 0.85rem 0; font-size:0.84rem; color:#F3D06B;">
                    ⬅️ <strong>No computador:</strong> vá na <strong>barra lateral esquerda</strong> em <strong>"🔐 Área do Assinante" → "Criar Conta Grátis"</strong>.<br>
                    📱 <strong>No celular:</strong> abra o menu <strong>"🔐 ENTRAR NA CONTA OU CONFIGURAR PERFIL"</strong> acima e clique em <strong>"Criar Conta Grátis"</strong>.
                </p>
                <a href="#topo-pagina" class="plan-cta cta-gold" style="display:inline-block; padding:0.55rem 1.3rem;">👆 Ir para o Cadastro (Área do Assinante)</a>
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
            m_seg = re.search(r"Segurado\(a\):\s*([^|\n]+)", parecer_texto)
            m_idade = re.search(r"Idade Atual:\s*([^\n]+)", parecer_texto)
            nome_seg = m_seg.group(1).strip() if m_seg else "SEGURADO AUDITADO"
            idade_seg = m_idade.group(1).strip() if m_idade else "-"

            b1 = re.search(r"1\.\s*TEMPO LÍQUIDO DE CONTRIBUIÇÃO\s*\n(.*?)(?=\n\s*2\.|\Z)", parecer_texto, re.S)
            b2 = re.search(r"2\.\s*PENDÊNCIAS NO MEU INSS\s*\n(.*?)(?=\n\s*3\.|\Z)", parecer_texto, re.S)
            b3 = re.search(r"3\.\s*ANÁLISE ESPECIAL E PCD.*?\n(.*)", parecer_texto, re.S)

            txt_b1 = b1.group(1).strip().replace("\n", "<br>") if b1 else parecer_texto.replace("\n", "<br>")
            txt_b2 = b2.group(1).strip().replace("\n", "<br>") if b2 else "Nenhuma pendência crítica localizada."
            txt_b3 = b3.group(1).strip().replace("\n", "<br>") if b3 else "-"

            c1 = f'<div style="background:#F3F6FA; border:1px solid #DCE3EC; border-left:5px solid #0F293A; border-radius:8px; padding:0.9rem 1.1rem; margin-bottom:0.85rem;"><div style="font-size:0.8rem; font-weight:800; color:#0F293A; margin-bottom:0.3rem; text-transform:uppercase;"><span style="background:#0F293A; color:#FFFFFF; padding:2px 7px; border-radius:4px; font-size:0.7rem; margin-right:6px;">TEMPO COMPROVADO</span>1. Tempo Líquido de Contribuição Apurado</div><div style="font-size:0.88rem; color:#334155; line-height:1.5;">{txt_b1}</div></div>'
            c2 = f'<div style="background:#FDF5F5; border:1px solid #F5D0D0; border-left:5px solid #B92D2D; border-radius:8px; padding:0.9rem 1.1rem; margin-bottom:0.85rem;"><div style="font-size:0.8rem; font-weight:800; color:#B92D2D; margin-bottom:0.3rem; text-transform:uppercase;"><span style="background:#B92D2D; color:#FFFFFF; padding:2px 7px; border-radius:4px; font-size:0.7rem; margin-right:6px;">AÇÃO NECESSÁRIA</span>2. Pendências e Indicadores no Meu INSS</div><div style="font-size:0.88rem; color:#334155; line-height:1.5;">{txt_b2}</div></div>'
            c3 = f'<div style="background:#FDFAF0; border:1px solid #EFE3B8; border-left:5px solid #C59B27; border-radius:8px; padding:0.9rem 1.1rem;"><div style="font-size:0.8rem; font-weight:800; color:#8A6810; margin-bottom:0.3rem; text-transform:uppercase;"><span style="background:#C59B27; color:#0F293A; padding:2px 7px; border-radius:4px; font-size:0.7rem; margin-right:6px;">ENQUADRAMENTO ESTRATÉGICO</span>3. Análise Especial e Regra PCD (LC 142/2013)</div><div style="font-size:0.88rem; color:#334155; line-height:1.5;">{txt_b3}</div></div>'
            topo_p = f'<div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.5rem; border-bottom:2px solid #C59B27; padding-bottom:0.75rem; margin-bottom:1rem;"><div><div style="font-size:0.75rem; font-weight:700; color:#C59B27; text-transform:uppercase; letter-spacing:0.5px;">Parecer Técnico Previdenciário & Auditoria de CNIS</div><div style="font-size:1.1rem; font-weight:800; color:#0F293A;">👤 {nome_seg}</div></div><div style="background:#0F293A; color:#F3D06B; padding:0.35rem 0.85rem; border-radius:20px; font-size:0.82rem; font-weight:700; border:1px solid #C59B27;">Idade Atual: {idade_seg}</div></div>'
            st.markdown(f'<div style="background:#FFFFFF; border:1px solid #CBD5E1; border-radius:12px; padding:1.2rem 1.4rem; margin-top:0.4rem; box-shadow:0 4px 12px rgba(15,41,58,0.05);">{topo_p}{c1}{c2}{c3}</div>', unsafe_allow_html=True)

        with tab1:
            st.dataframe(df_vinculos, use_container_width=True)
        with tab2:
            st.dataframe(df_res, use_container_width=True)
        with tab3:
            st.dataframe(abas["3. ⚠️ APOSENT. ESPECIAL"], use_container_width=True)
        with tab4:
            st.dataframe(abas["4. ♿ REGRA PCD (LC 142)"], use_container_width=True)
        with tab5:
            st.dataframe(df_fin, use_container_width=True)

        st.markdown("""
        <div class="save-warning-box">
            <div class="save-warning-badge">⚠️ Salve seus arquivos agora</div>
            <div class="save-warning-title">📥 Baixe e guarde seu dossiê antes de sair ou atualizar a página</div>
            <div style="color:#E2E8F0;">
                Por segurança e privacidade (<strong>LGPD</strong>), a leitura do CNIS é feita em memória e o dossiê gerado aqui <strong>não fica salvo no sistema para reabertura posterior</strong>. Baixe agora o <strong>Relatório Executivo (.pdf)</strong> e a <strong>Planilha Detalhada (.xlsx)</strong> no seu celular ou computador.
            </div>
            <div class="save-warning-support">
                💬 <strong>Suporte:</strong> WhatsApp <strong style="color:#F3D06B; white-space:nowrap;">+55 21 96880-7955</strong> &nbsp;|&nbsp; E-mail: <strong style="color:#F3D06B;">inteligenciaprevidenciariaaudi@gmail.com</strong>
            </div>
        </div>
        """, unsafe_allow_html=True)

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
            st.caption("💡 Dica: Ao abrir a planilha no Excel, clique nas **6 abas coloridas** na parte inferior da tela para ver cada módulo.")
        with col_dl3:
            texto_bruto = (
                f"Olá! Acabei de gerar minha Auditoria Previdenciária "
                f"(Tempo Líquido: {tempo_card} | Potencial PCD: {rmi_pcd}) "
                f"e gostaria de agendar uma análise com o escritório parceiro."
            )
            params_whats = urllib.parse.urlencode({"phone": "5521968807955", "text": texto_bruto})
            dominio_wa = "web." + "whatsapp" + ".com"
            link_whats = f"https://{dominio_wa}/send?{params_whats}"

            st.link_button(
                label="⚖️ Enviar para Escritório Parceiro",
                url=link_whats,
                use_container_width=True,
            )

if not email_logado and st.session_state["aviso_cadastro_obrigatorio"]:
    st.warning(
        "🔐 **Cadastre-se primeiro antes de ir para o pagamento:** "
        "no computador, abra a **barra lateral esquerda** em **'🔐 Área do Assinante' → 'Criar Conta Grátis'**. "
        "No celular, abra o menu **'🔐 ENTRAR NA CONTA OU CONFIGURAR PERFIL'** e clique em **'Criar Conta Grátis'**."
    )

if email_logado:
    cta_avulso_html = f'<a href="{link_avulso_atual}" target="_blank" class="plan-cta cta-outline">⚡ Liberar Dossiê (Pix ou Cartão)</a>'
    cta_passes_html = f'<a href="{link_7d_atual}" target="_blank" class="plan-cta cta-gold">⚡ Passe 7 Dias — R$ 39,90 (Pix/Cartão)</a><a href="{link_30d_atual}" target="_blank" class="plan-cta cta-gold-outline">👑 Passe 30 Dias — R$ 69,90 (Pix/Cartão)</a>'
else:
    cta_avulso_html = '<a href="#topo-pagina" class="plan-cta cta-outline">🔐 Cadastre-se na Área do Assinante Primeiro</a>'
    cta_passes_html = '<a href="#topo-pagina" class="plan-cta cta-gold">🔐 Cadastre-se na Área do Assinante Primeiro</a>'

st.markdown(f"""
<div style="font-size: 1.05rem; font-weight: 700; color: #0F293A; margin-top: 0.8rem;">💎 Escolha o Plano Ideal (Liberação Automática via Pix ou Cartão)</div>
<div class="pricing-grid">
    <div class="plan-card dossie-trigger-card" tabindex="0">
        <div>
            <div class="plan-top"><span class="plan-name" style="color:#0F293A;">Diagnóstico Inicial</span></div>
            <div class="plan-price" style="color:#0F293A;">Grátis <span>/ prévia</span></div>
            <div class="dossie-pill-hint">✦ Veja o que o Dossiê destrava</div>
            <ul class="plan-list" style="color:#334155;">
                <li>✅ Tempo líquido e leitura do CNIS</li>
                <li>✅ Alerta de pendências nos vínculos</li>
                <li>❌ Como resolver erros, RMI e PDF/Excel</li>
            </ul>
        </div>
        <div style="text-align:center; font-size:0.78rem; color:#64748B; font-weight:600; padding:0.45rem; background:#F1F5F9; border-radius:6px;">Ativo — Envie seu PDF acima</div>
        {POPUP_DOSSIE_HTML}
    </div>
    <div class="plan-card dossie-trigger-card" tabindex="0">
        <div>
            <div class="plan-top"><span class="plan-name" style="color:#0F293A;">Dossiê Completo Avulso</span></div>
            <div class="plan-price" style="color:#0F293A;">{preco_avulso_txt} <span>/ único</span></div>
            <div class="dossie-pill-hint">✦ Raio-X completo antes do pedido</div>
            <ul class="plan-list" style="color:#334155;">
                <li>✅ Plano de ação (como corrigir erros)</li>
                <li>✅ Cálculo RMI (1994–2026), Especial e PCD</li>
                <li>✅ Download PDF Executivo + Planilha Excel</li>
            </ul>
        </div>
        {cta_avulso_html}
        {POPUP_DOSSIE_HTML}
    </div>
    <div class="plan-card featured dossie-trigger-card" tabindex="0">
        <div>
            <div class="plan-top">
                <span class="plan-name">Passe Ilimitado</span>
                <span class="plan-badge">Mais Escolhido</span>
            </div>
            <div class="plan-price" style="color:#F3D06B;">R$ 39,90 <span>/ 7d ou R$ 69,90 / 30d</span></div>
            <div class="dossie-pill-hint">✦ Raio-X completo antes do pedido</div>
            <ul class="plan-list" style="color:#E2E8F0;">
                <li>✅ <strong>Simulações e PDFs/Excel ilimitados</strong></li>
                <li>✅ Teste cenários (PCD, Especial, futuro)</li>
                <li>✅ Analise também familiares ou clientes</li>
            </ul>
        </div>
        <div>
            {cta_passes_html}
        </div>
        {POPUP_DOSSIE_HTML}
    </div>
</div>
<div style="background-color:#EFF6FF; border:1px solid #BFDBFE; border-left:4px solid #0F293A; color:#0F293A; padding:0.65rem 0.9rem; border-radius:8px; font-size:0.83rem; margin-bottom:0.8rem;">
    💡 <strong>Como funciona a liberação automática:</strong> Crie sua conta gratuita em <strong>"🔐 Área do Assinante"</strong> (na barra lateral esquerda no computador ou em <strong>"🔐 ENTRAR NA CONTA OU CONFIGURAR PERFIL"</strong> no celular) antes de pagar e conclua o Pix ou Cartão no Mercado Pago. Assim que o pagamento for aprovado, clique em <strong>"🔄 Atualizar"</strong> para liberar o acesso na hora!
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
""", unsafe_allow_html=True)

if not email_logado:
    if st.button("ℹ️ Onde faço meu cadastro antes de pagar?", key="btn_onde_cadastrar"):
        st.session_state["aviso_cadastro_obrigatorio"] = True
        st.rerun()

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
        st.markdown("Sim. O processamento ocorre de forma isolada e 100% em memória apenas durante a leitura do ficheiro. Nenhum PDF enviado fica gravado no disco do servidor, em total conformidade com a Lei Geral de Proteção de Dados (LGPD).")

st.markdown(
    """
<div style="text-align:center; color:#475569; font-size:0.82rem; margin:2rem auto 4.5rem auto; padding:1.25rem 1rem 0.5rem 1rem; max-width:860px; border-top:1px solid #CBD5E1; line-height:1.6;">
    <div style="color:#0F293A; font-weight:700; font-size:0.86rem; margin-bottom:2px;">
        Inteligência Previdenciária & Auditoria · Idealizado e desenvolvido por Neivaldo Barbosa de Sá
    </div>
    <div style="color:#64748B; font-size:0.79rem; margin-bottom:0.75rem;">
        Automação, análise de dados e otimização de processos aplicados à revisão previdenciária.
    </div>
    <div style="display:flex; justify-content:center; gap:12px; flex-wrap:wrap; margin-bottom:0.9rem;">
        <a href="https://www.youtube.com/@inteligenciaprevidenciaria" target="_blank" rel="noopener noreferrer"
           style="color:#0F293A; background:#F8FAFC; border:1px solid #CBD5E1; padding:5px 14px; border-radius:20px; text-decoration:none; font-size:0.78rem; font-weight:700; display:inline-flex; align-items:center; gap:6px;">
            <span style="color:#DC2626;">▶</span> Canal no YouTube
        </a>
        <a href="https://www.linkedin.com/in/neivaldo-barbosa-de-s%C3%A1-4867ba33/" target="_blank" rel="noopener noreferrer"
           style="color:#0F293A; background:#F8FAFC; border:1px solid #C59B27; padding:5px 14px; border-radius:20px; text-decoration:none; font-size:0.78rem; font-weight:700; display:inline-flex; align-items:center; gap:6px;">
            <span style="color:#0A66C2; font-weight:900;">in</span> LinkedIn · Projetos & Otimização
        </a>
    </div>
    <div style="font-size:0.76rem; color:#64748B;">
        Ferramenta privada e independente de auditoria previdenciária. Não possuímos vínculo com o INSS ou órgãos governamentais.
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# Caixa lateral interativa quando o usuário clica em "Dúvidas INSS"
if st.session_state.get("gaveta_ia_aberta"):
    with st.container(key="caixa_lateral_inss"):
        col_ia_tit, col_ia_fechar = st.columns([5, 1])
        with col_ia_tit:
            st.markdown(
                '<div style="display:flex; align-items:center; gap:6px; color:#F3D06B; font-weight:800; font-size:0.95rem;">'
                '<span style="background:#C59B27; color:#0F293A; border-radius:6px; padding:1px 6px; font-size:0.8rem;">✦</span>'
                '<span>Dúvidas INSS • Gemini</span>'
                '</div>',
                unsafe_allow_html=True,
            )
            st.caption("Conversa com Gemini — cálculo exato no motor do CNIS.")
        with col_ia_fechar:
            if st.button("✕", key="btn_fechar_gaveta_ia", use_container_width=True):
                st.session_state["gaveta_ia_aberta"] = False
                st.rerun()

        c_rap1, c_rap2, c_rap3 = st.columns([2, 2, 1.3])
        with c_rap1:
            if st.button("Como baixar CNIS?", key="ia_rap_cnis", use_container_width=True):
                perg = "Como baixar o CNIS completo no Meu INSS?"
                resp = _responder_duvida_inss(perg, st.session_state["mensagens_ia"])
                st.session_state["mensagens_ia"].append({"role": "user", "content": perg})
                st.session_state["mensagens_ia"].append({"role": "assistant", "content": resp})
                st.rerun()
        with c_rap2:
            if st.button("Regra PCD / TEA", key="ia_rap_pcd", use_container_width=True):
                perg = "Como funciona a aposentadoria PCD e TEA pela LC 142/2013?"
                resp = _responder_duvida_inss(perg, st.session_state["mensagens_ia"])
                st.session_state["mensagens_ia"].append({"role": "user", "content": perg})
                st.session_state["mensagens_ia"].append({"role": "assistant", "content": resp})
                st.rerun()
        with c_rap3:
            if st.button("Limpar", key="ia_limpar_chat", use_container_width=True):
                st.session_state["mensagens_ia"] = [
                    {
                        "role": "assistant",
                        "content": "✦ **Gemini** aqui! Posso conversar sobre dúvidas de **CNIS**, pendências no INSS, tempo concomitante, aposentadoria especial ou regra PCD (LC 142/2013). *(O cálculo oficial do seu extrato continua no motor da página.)*"
                    }
                ]
                st.rerun()

        for msg in st.session_state["mensagens_ia"][-8:]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        with st.form("form_gaveta_ia", clear_on_submit=True):
            pergunta_ia = st.text_input(
                "Sua dúvida sobre CNIS ou INSS:",
                placeholder="Ex.: O que significa PREC-MENOR-MIN?",
                key="input_gaveta_ia",
            )
            enviou_ia = st.form_submit_button("Enviar", type="primary", use_container_width=True)

        if enviou_ia:
            if pergunta_ia.strip():
                perg_limpa = pergunta_ia.strip()
                resp = _responder_duvida_inss(perg_limpa, st.session_state["mensagens_ia"])
                st.session_state["mensagens_ia"].append({"role": "user", "content": perg_limpa})
                st.session_state["mensagens_ia"].append({"role": "assistant", "content": resp})
                st.rerun()
            else:
                st.warning("Digite uma dúvida antes de enviar.")

if st.query_params.get("abrir_ia") == "1":
    st.query_params.pop("abrir_ia", None)
    st.session_state["gaveta_ia_aberta"] = True
    st.rerun()

assunto_suporte = urllib.parse.quote("Suporte Técnico - Inteligência Previdenciária & Auditoria")
corpo_suporte = urllib.parse.quote("Olá! Preciso de suporte técnico no site Inteligência Previdenciária & Auditoria.")
params_sup = urllib.parse.urlencode({"phone": "5521968807955", "text": "Olá! Preciso de suporte técnico no site Inteligência Previdenciária & Auditoria."})
dominio_sup = "web." + "whatsapp" + ".com"
whats_suporte = f"https://{dominio_sup}/send?{params_sup}"

st.markdown(f"""
<style>
    .support-wrapper {{
        position: relative;
        display: inline-block;
    }}
    .support-menu {{
        display: none;
        position: absolute;
        bottom: 48px;
        right: 0;
        background: #0F293A;
        border: 1.5px solid #C59B27;
        border-radius: 12px;
        padding: 10px;
        width: 245px;
        box-shadow: 0 8px 20px rgba(0,0,0,0.35);
        z-index: 10000;
    }}
    .support-wrapper:hover .support-menu,
    .support-wrapper:focus-within .support-menu {{
        display: block;
    }}
    .support-item {{
        display: block;
        padding: 8px 10px;
        margin-bottom: 6px;
        border-radius: 8px;
        font-size: 0.82rem;
        font-weight: 700;
        text-decoration: none !important;
        text-align: center;
    }}
    .sup-wa {{ background: #25D366; color: #FFFFFF !important; }}
    .sup-mail {{ background: #C59B27; color: #0F293A !important; }}
    .sup-tip {{
        background: rgba(255,255,255,0.08);
        color: #E2E8F0;
        font-size: 0.74rem;
        padding: 7px 9px;
        border-radius: 8px;
        border-left: 3px solid #C59B27;
        line-height: 1.35;
    }}
</style>

<div class="floating-container">
    <a href="?abrir_ia=1" target="_self" class="float-btn float-gemini">
        <span>✦ Dúvidas INSS (Gemini)</span>
    </a>
    <div class="support-wrapper" tabindex="0">
        <div class="support-menu">
            <a href="{whats_suporte}" target="_blank" rel="noopener noreferrer" class="support-item sup-wa">💬 Suporte via WhatsApp</a>
            <a href="mailto:{EMAIL_SUPORTE}?subject={assunto_suporte}&body={corpo_suporte}" class="support-item sup-mail">✉️ Suporte via E-mail</a>
            <div class="sup-tip">
                <strong>💡 Dica de Entrega:</strong><br>
                Se o navegador corporativo bloquear o WhatsApp, use o e-mail ou salve o número: <strong>+55 21 96880-7955</strong>.
            </div>
        </div>
        <div class="float-btn float-support" style="cursor:pointer;">
            <span>🛠️ Suporte Técnico</span>
            <span style="background:#0F293A; color:#F3D06B; border-radius:50%; width:20px; height:20px; display:inline-flex; align-items:center; justify-content:center; font-size:0.75rem; margin-left:4px;" title="Opções de contato e ajuda na entrega">💡</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)
