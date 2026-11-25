import streamlit as st
from supabase import create_client, Client, ClientOptions
from datetime import datetime, timezone
from urllib.parse import urlencode


def get_supabase(access_token: str | None = None) -> Client:
    """Cria o cliente Supabase com o token JWT do utilizador quando autenticado."""
    url = st.secrets["supabase"]["url"].strip()
    key = st.secrets["supabase"]["key"].strip()
    if access_token:
        options = ClientOptions(headers={"Authorization": f"Bearer {access_token}"})
        return create_client(url, key, options=options)
    return create_client(url, key)


def renderizar_painel_acesso(chave_prefixo: str = "side") -> dict:
    """
    Renderiza o login/cadastro e status da assinatura.
    Retorna um dicionário com o estado atual de acesso do utilizador.
    """
    estado = {
        "autenticado": False,
        "tem_acesso_completo": False,
        "tipo_acesso": "GRATUITO",
        "email": None,
        "dias_restantes": 0,
        "creditos": 0,
    }

    st.markdown('<div class="sidebar-title">🔐 Área do Assinante</div>', unsafe_allow_html=True)

    if "user_session" not in st.session_state or st.session_state["user_session"] is None:
        supabase = get_supabase()
        modo_acesso = st.radio(
            "Escolha uma opção:",
            options=["Entrar na Conta", "Criar Conta Grátis"],
            horizontal=True,
            label_visibility="collapsed",
            key=f"{chave_prefixo}_modo_acesso",
        )

        if modo_acesso == "Entrar na Conta":
            email_login = st.text_input("E-mail", key=f"{chave_prefixo}_login_email").strip().lower()
            senha_login = st.text_input("Senha", type="password", key=f"{chave_prefixo}_login_senha")
            if st.button("Entrar", use_container_width=True, type="primary", key=f"{chave_prefixo}_btn_entrar"):
                if not email_login or not senha_login:
                    st.warning("Preencha e-mail e senha.")
                else:
                    try:
                        res = supabase.auth.sign_in_with_password({"email": email_login, "password": senha_login})
                        st.session_state["user_session"] = res.session
                        st.session_state["user_email"] = res.user.email
                        st.session_state["user_id"] = res.user.id
                        st.rerun()
                    except Exception as e:
                        st.error(f"E-mail ou senha inválidos ({e}).")
        else:
            email_cad = st.text_input("Seu melhor e-mail", key=f"{chave_prefixo}_cad_email").strip().lower()
            senha_cad = st.text_input("Crie uma senha (mín. 6 dígitos)", type="password", key=f"{chave_prefixo}_cad_senha")
            if st.button("Cadastrar Grátis", use_container_width=True, type="primary", key=f"{chave_prefixo}_btn_cad"):
                if not email_cad or len(senha_cad) < 6:
                    st.warning("Informe um e-mail válido e senha com ao menos 6 caracteres.")
                else:
                    try:
                        res = supabase.auth.sign_up({"email": email_cad, "password": senha_cad})
                        if res.session:
                            st.session_state["user_session"] = res.session
                            st.session_state["user_email"] = res.user.email
                            st.session_state["user_id"] = res.user.id
                            st.success("Conta criada com sucesso!")
                            st.rerun()
                        else:
                            st.info("Conta criada! Selecione 'Entrar na Conta' acima para acessar.")
                    except Exception as e:
                        st.error(f"Não foi possível cadastrar: {e}")

        st.caption("Crie sua conta gratuita para vincular seu pagamento automaticamente.")
        return estado

    # Utilizador autenticado
    session = st.session_state["user_session"]
    supabase = get_supabase(session.access_token)
    email_usuario = st.session_state.get("user_email", "")

    dados = None
    try:
        resp = (
            supabase.table("assinaturas")
            .select("*")
            .eq("id", st.session_state["user_id"])
            .execute()
        )
        if resp.data:
            dados = resp.data[0]
    except Exception as e:
        st.error(f"Erro ao ler assinatura: {e}")

    st.success(f"Conectado: **{email_usuario}**")

    if dados:
        plano = dados.get("plano", "GRATUITO")
        creditos = int(dados.get("creditos_avulsos", 0))
        expira_str = dados.get("expira_em")
        agora = datetime.now(timezone.utc)

        dias_restantes = 0
        passe_ativo = False

        if expira_str:
            expira_dt = datetime.fromisoformat(expira_str.replace("Z", "+00:00"))
            if expira_dt > agora:
                passe_ativo = True
                dias_restantes = max(1, (expira_dt - agora).days)

        if passe_ativo:
            nome_plano = "Passe 30 Dias" if plano == "PASSE_30D" else "Passe 7 Dias"
            st.info(f"✅ **{nome_plano} Ativo** — Válido por mais **{dias_restantes} dia(s)**.")
            estado.update({
                "autenticado": True,
                "tem_acesso_completo": True,
                "tipo_acesso": plano,
                "email": email_usuario,
                "dias_restantes": dias_restantes,
                "creditos": creditos,
            })
        elif creditos > 0:
            st.info(f"🎟️ **Dossiê Avulso** — **{creditos} crédito(s)** disponível(is).")
            estado.update({
                "autenticado": True,
                "tem_acesso_completo": True,
                "tipo_acesso": "AVULSO",
                "email": email_usuario,
                "creditos": creditos,
            })
        else:
            st.warning("🔒 **Plano Gratuito (Prévia)** — Escolha um plano para liberar o PDF e o Excel completo:")
            _exibir_botoes_checkout(email_usuario)
            estado.update({
                "autenticado": True,
                "tem_acesso_completo": False,
                "tipo_acesso": "GRATUITO",
                "email": email_usuario,
            })
    else:
        st.warning("⚠️ Registro de assinatura não localizado para esta conta.")

    col_atualizar, col_sair = st.columns(2)
    with col_atualizar:
        if st.button("🔄 Atualizar", use_container_width=True, key=f"{chave_prefixo}_btn_att", help="Clique após concluir um pagamento"):
            st.rerun()
    with col_sair:
        if st.button("Sair", use_container_width=True, key=f"{chave_prefixo}_btn_sair"):
            try:
                supabase.auth.sign_out()
            except Exception:
                pass
            st.session_state["user_session"] = None
            st.session_state.pop("credito_consumido_nesta_sessao", None)
            st.rerun()

    return estado


def gerar_link_checkout_personalizado(url_base: str, email_usuario: str | None) -> str:
    """Anexa o e-mail do utilizador logado ao link do Mercado Pago para liberação automática."""
    if not email_usuario:
        return url_base
    separador = "&" if "?" in url_base else "?"
    param = urlencode({"payer_email": email_usuario, "external_reference": email_usuario})
    return f"{url_base}{separador}{param}"


def _exibir_botoes_checkout(email_usuario: str):
    links = st.secrets.get("links_pagamento", {})
    if links.get("avulso"):
        st.link_button(
            "📄 Dossiê Avulso (R$ 19,90)",
            gerar_link_checkout_personalizado(links["avulso"], email_usuario),
            use_container_width=True,
        )
    if links.get("passe_7d"):
        st.link_button(
            "⚡ Passe 7 Dias (R$ 39,90)",
            gerar_link_checkout_personalizado(links["passe_7d"], email_usuario),
            use_container_width=True,
        )
    if links.get("passe_30d"):
        st.link_button(
            "👑 Passe 30 Dias (R$ 69,90)",
            gerar_link_checkout_personalizado(links["passe_30d"], email_usuario),
            use_container_width=True,
        )


def registrar_uso_credito_avulso(acesso: dict):
    """Desconta 1 crédito avulso quando o utilizador gera um Dossiê Avulso."""
    if acesso.get("tipo_acesso") == "AVULSO" and not st.session_state.get("credito_consumido_nesta_sessao"):
        session = st.session_state.get("user_session")
        if session:
            try:
                supabase = get_supabase(session.access_token)
                supabase.rpc("consumir_credito_avulso").execute()
                st.session_state["credito_consumido_nesta_sessao"] = True
            except Exception:
                pass
                
def enviar_email_recuperacao(email: str, redirect_url: str):
    """Envia o link de redefinição de senha para o e-mail do usuário."""
    try:
        client = get_supabase()
        client.auth.reset_password_for_email(
            email,
            options={"redirect_to": redirect_url}
        )
        return True, "Enviamos um link de recuperação para o seu e-mail. Verifique sua caixa de entrada e spam."
    except Exception as e:
        return False, f"Erro ao enviar e-mail de recuperação: {str(e)}"

def atualizar_nova_senha(nova_senha: str):
    """Atualiza a senha do usuário após ele clicar no link do e-mail."""
    try:
        client = get_supabase()
        client.auth.update_user({"password": nova_senha})
        return True, "Senha alterada com sucesso! Você já pode acessar seu dossiê."
    except Exception as e:
        return False, f"Erro ao atualizar a senha: {str(e)}"
        
