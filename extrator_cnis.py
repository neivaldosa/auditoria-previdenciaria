import re
from datetime import datetime, timedelta
import pandas as pd
import pdfplumber
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

TETOS_INSS = {
    1994: 582.86, 1995: 832.66, 1996: 957.56, 1997: 1031.87, 1998: 1081.50,
    1999: 1255.32, 2000: 1328.25, 2001: 1430.00, 2002: 1561.56, 2003: 1869.34,
    2004: 2508.72, 2005: 2668.15, 2006: 2801.82, 2007: 2894.28, 2008: 3038.99,
    2009: 3218.90, 2010: 3467.40, 2011: 3691.74, 2012: 3916.20, 2013: 4159.00,
    2014: 4390.24, 2015: 4663.75, 2016: 5189.82, 2017: 5531.31, 2018: 5645.80,
    2019: 5839.45, 2020: 6101.06, 2021: 6433.57, 2022: 7087.22, 2023: 7507.49,
    2024: 7786.02, 2025: 8157.41, 2026: 8475.55
}

FATOR_ATUALIZACAO_APROX = {
    1994: 7.45, 1995: 6.10, 1996: 5.25, 1997: 4.90, 1998: 4.70, 1999: 4.45,
    2000: 4.15, 2001: 3.85, 2002: 3.50, 2003: 3.05, 2004: 2.85, 2005: 2.68,
    2006: 2.58, 2007: 2.48, 2008: 2.34, 2009: 2.22, 2010: 2.10, 2011: 1.97,
    2012: 1.86, 2013: 1.76, 2014: 1.66, 2015: 1.52, 2016: 1.38, 2017: 1.33,
    2018: 1.29, 2019: 1.24, 2020: 1.18, 2021: 1.09, 2022: 1.01, 2023: 1.00,
    2024: 1.00, 2025: 1.00, 2026: 1.00
}


def calcular_tempo_liquido(vinculos):
    intervalos = []
    for v in vinculos:
        if v["Data_Admissao"]:
            dt_ini = datetime.strptime(v["Data_Admissao"], "%d/%m/%Y")
            if v["Data_Demissao"]:
                dt_fim = datetime.strptime(v["Data_Demissao"], "%d/%m/%Y")
            elif v["Ultima_Remun"]:
                mes, ano = map(int, v["Ultima_Remun"].split("/"))
                dt_fim = datetime(ano, mes, 28)
            else:
                continue
            if dt_fim >= dt_ini:
                intervalos.append((dt_ini, dt_fim))

    if not intervalos:
        return 0

    intervalos.sort()
    unificados = [intervalos.pop(0)]

    for atual_ini, atual_fim in intervalos:
        ult_ini, ult_fim = unificados[-1]
        if atual_ini <= ult_fim + timedelta(days=1):
            unificados[-1] = (ult_ini, max(ult_fim, atual_fim))
        else:
            unificados.append((atual_ini, atual_fim))

    return sum((fim - ini).days + 1 for ini, fim in unificados)


def formatar_ymd(dias_totais):
    anos = int(dias_totais // 365.25)
    resto_ano = dias_totais % 365.25
    meses = int(resto_ano // 30.4375)
    dias = int(round(resto_ano % 30.4375))
    return f"{anos} anos, {meses} meses e {dias} dias"


def extrair_cnis_v5(caminho_pdf, caminho_excel_saida, sexo="M", ambiente_operacional=False, possui_pcd=True):
    vinculos = []
    remuneracoes_dict = {}
    vinculo_atual = None
    data_nascimento = ""
    nome_segurado = ""

    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text(layout=True)
            if not texto:
                continue

            for i, linha in enumerate(texto.split("\n")):
                linha_limpa = " ".join(linha.split())
                if not linha_limpa:
                    continue

                if "Nome:" in linha_limpa and not nome_segurado:
                    m_nome = re.search(r"Nome:\s*(.+?)(?:Data|$)", linha_limpa)
                    if m_nome:
                        nome_segurado = m_nome.group(1).strip()

                if "Data de nascimento:" in linha_limpa and not data_nascimento:
                    m_nasc = re.search(r"Data de nascimento:\s*(\d{2}/\d{2}/\d{4})", linha_limpa)
                    if m_nasc:
                        data_nascimento = m_nasc.group(1)

                for m_rem in re.finditer(r"\b(\d{2}/\d{4})\s+([\d\.]+,\d{2})\b", linha_limpa):
                    comp = m_rem.group(1)
                    val_str = m_rem.group(2).replace(".", "").replace(",", ".")
                    try:
                        val_float = float(val_str)
                        mes_c, ano_c = map(int, comp.split("/"))
                        if ano_c > 1994 or (ano_c == 1994 and mes_c >= 7):
                            remuneracoes_dict[comp] = remuneracoes_dict.get(comp, 0.0) + val_float
                    except ValueError:
                        pass

                match_inicio = re.match(r"^(\d{1,2})\s+(\d{3}\.\d{5}\.\d{2}-\d)\s+(.+)$", linha_limpa)
                if match_inicio:
                    seq = int(match_inicio.group(1))
                    resto = match_inicio.group(3)

                    if "Benefício" in resto or "AUXÍLIO" in resto or "AUXILIO" in resto:
                        vinculo_atual = None
                        continue

                    if vinculo_atual:
                        vinculos.append(vinculo_atual)

                    match_cnpj = re.search(r"(\d{2}\.\d{3}\.\d{3}(?:/\d{4}-\d{2})?)", resto)
                    if match_cnpj:
                        cnpj = match_cnpj.group(1)
                    elif "RECOLHIMENTO" in resto or "Contribuinte Individual" in resto or "Facultativo" in resto:
                        cnpj = "RECOLHIMENTO (NIT)"
                    else:
                        cnpj = "Não informado"

                    datas = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", resto)
                    data_ini = datas.pop(0) if datas else ""
                    data_fim = datas.pop(0) if datas else ""

                    competencias_linha = re.findall(r"\b\d{2}/\d{4}\b", resto)
                    ult_remun = competencias_linha.pop() if competencias_linha else ""

                    if "Empregado ou Agente" in resto:
                        tipo_encontrado = "Empregado ou Agente Público"
                    elif "Contribuinte Individual" in resto:
                        tipo_encontrado = "Contribuinte Individual"
                    elif "Facultativo" in resto:
                        tipo_encontrado = "Facultativo"
                    elif "Segurado Especial" in resto:
                        tipo_encontrado = "Segurado Especial"
                    else:
                        tipo_encontrado = "Não identificado"

                    empresa = resto.replace(cnpj, "") if match_cnpj else resto
                    for corte in ["Empregado ou Agente", "Contribuinte Individual", "Facultativo", "Segurado Especial"]:
                        if corte in empresa:
                            empresa = empresa.split(corte).pop(0)
                            break

                    empresa = re.sub(r"\b[A-Z0-9]*\d{4,}[A-Z0-9]*\b", "", empresa)
                    empresa = " ".join(empresa.split()).strip(" -")

                    if not empresa and cnpj == "RECOLHIMENTO (NIT)":
                        empresa = "RECOLHIMENTO INDIVIDUAL / MEI"

                    vinculo_atual = {
                        "Seq": seq, "Nome_Empresa": empresa, "CNPJ_Codigo": cnpj,
                        "Tipo_Filiado": tipo_encontrado, "Data_Admissao": data_ini,
                        "Data_Demissao": data_fim, "Ultima_Remun": ult_remun,
                        "Indicadores_Lista": [], "Dias_Totais": 0, "Dias_Pre_Reforma_2019": 0,
                        "Potencial_Especial": "Não", "Status_Auditoria": "OK",
                        "Plano_Acao_INSS": "Nenhuma pendência identificada neste vínculo.",
                    }
                    continue

                if vinculo_atual:
                    if "Indicadores:" in linha_limpa:
                        parte_ind = linha_limpa.split("Indicadores:").pop()
                        for ind in parte_ind.split(","):
                            ind_limpo = ind.strip()
                            if ind_limpo and ind_limpo not in vinculo_atual["Indicadores_Lista"]:
                                vinculo_atual["Indicadores_Lista"].append(ind_limpo)

                    for sigla in ["IREM-ACD", "IREM-INDPEND", "IVIN-JORN-DIFERENCIADA", "PREM-BLOQ-EC103", "IREC-MEI", "IREC-LC123", "IREC-INDPEND", "PREC-FACULTCONC", "EXTEMP", "PEXT"]:
                        if sigla in linha_limpa and sigla not in vinculo_atual["Indicadores_Lista"]:
                            vinculo_atual["Indicadores_Lista"].append(sigla)

    if vinculo_atual:
        vinculos.append(vinculo_atual)
    max_seq_clt = max((x["Seq"] for x in vinculos if "Empregado" in x["Tipo_Filiado"]), default=0)
    dt_reforma = datetime(2019, 11, 13)

    for v in vinculos:
        v["Indicadores"] = ", ".join(v["Indicadores_Lista"]) if v["Indicadores_Lista"] else "Sem Pendência"
        del v["Indicadores_Lista"]

        if v["Data_Admissao"]:
            dt_ini = datetime.strptime(v["Data_Admissao"], "%d/%m/%Y")
            if v["Data_Demissao"]:
                dt_fim = datetime.strptime(v["Data_Demissao"], "%d/%m/%Y")
            elif v["Ultima_Remun"]:
                mes, ano = map(int, v["Ultima_Remun"].split("/"))
                dt_fim = datetime(ano, mes, 28)
            else:
                dt_fim = dt_ini

            v["Dias_Totais"] = max(0, (dt_fim - dt_ini).days + 1)
            if dt_ini <= dt_reforma:
                fim_pre = min(dt_fim, dt_reforma)
                v["Dias_Pre_Reforma_2019"] = (fim_pre - dt_ini).days + 1

            if not v["Data_Demissao"] and "Empregado" in v["Tipo_Filiado"]:
                if v["Seq"] < max_seq_clt:
                    v["Status_Auditoria"] = "ERRO CRÍTICO: Sem Data de Saída"
                    v["Plano_Acao_INSS"] = "Solicitar Acerto de Vínculos e Remunerações no Meu INSS e anexar baixa na CTPS ou TRCT."
                else:
                    v["Status_Auditoria"] = "VÍNCULO ATIVO (Em aberto)"
                    v["Plano_Acao_INSS"] = "Vínculo atual em curso. Contado até a última remuneração recolhida."

        palavras_risco = ["PETROLEO", "OLEO E GAS", "ENERGIA", "OFFSHORE", "QUIMICA", "MINERACAO", "METALURGICA", "HOSPITAL", "VIGILANCIA", "ELETRICA"]
        tem_risco_ramo = any(p in v["Nome_Empresa"].upper() for p in palavras_risco)

        if "PREM-BLOQ-EC103" in v["Indicadores"] or "IREC-MEI" in v["Indicadores"]:
            v["Status_Auditoria"] = "ALERTA: Recolhimento Bloqueado/Concomitante"
            v["Plano_Acao_INSS"] = "Guia MEI/Autônomo paga no mesmo mês de vínculo CLT. Não soma tempo extra."
        elif "IVIN-JORN-DIFERENCIADA" in v["Indicadores"] or tem_risco_ramo:
            if ambiente_operacional:
                v["Potencial_Especial"] = "Sim (Operacional / PPP)"
                v["Status_Auditoria"] = "OPORTUNIDADE: Atividade Especial / Risco"
                v["Plano_Acao_INSS"] = "Solicitar PPP e LTCAT à empresa e conferir código de ocorrência GFIP/eSocial."
            else:
                v["Potencial_Especial"] = "Não (Área Administrativa)"
                if v["Status_Auditoria"] == "OK":
                    v["Status_Auditoria"] = "INFO: Setor Industrial (Cargo Administrativo)"
                    v["Plano_Acao_INSS"] = "Marcação de jornada diferenciada/setor de risco, mas sem conversão especial por atuar em área administrativa."
        elif "IREM-INDPEND" in v["Indicadores"] and v["Status_Auditoria"] == "OK":
            v["Status_Auditoria"] = "ATENÇÃO: Pendência na Remuneração"
            v["Plano_Acao_INSS"] = "Guardar holerites e comprovantes de dissídio (IREM-ACD)."

    df_vinculos = pd.DataFrame(vinculos)
    dias_brutos = sum(v["Dias_Totais"] for v in vinculos)
    dias_liquidos = calcular_tempo_liquido(vinculos)
    dias_concomitantes = dias_brutos - dias_liquidos

    idade_atual = "Não informada"
    if data_nascimento:
        dt_nasc = datetime.strptime(data_nascimento, "%d/%m/%Y")
        idade_atual = f"{int((datetime.now() - dt_nasc).days // 365.25)} anos"

    df_resumo = pd.DataFrame([
        {"Indicador": "Segurado", "Valor": nome_segurado},
        {"Indicador": "Data de Nascimento / Idade Atual", "Valor": f"{data_nascimento} ({idade_atual})"},
        {"Indicador": "Ambiente de Trabalho Informado", "Valor": "Operacional / Área de Risco" if ambiente_operacional else "Administrativo / Escritório"},
        {"Indicador": "Soma Bruta de Todos os Vínculos", "Valor": f"{dias_brutos} dias ({formatar_ymd(dias_brutos)})"},
        {"Indicador": "Períodos Concomitantes Descontados", "Valor": f"{dias_concomitantes} dias"},
        {"Indicador": "TEMPO LÍQUIDO DE CONTRIBUIÇÃO (COMUM)", "Valor": f"{dias_liquidos} dias ({formatar_ymd(dias_liquidos)})"},
    ])

    fator_conv = 1.4 if sexo == "M" else 1.2
    linhas_especial = []
    total_dias_bonus_especial = 0

    for v in vinculos:
        if "Sim" in v["Potencial_Especial"]:
            dias_pre = v["Dias_Pre_Reforma_2019"]
            dias_bonus = int(round((dias_pre * fator_conv) - dias_pre))
            total_dias_bonus_especial += dias_bonus
            linhas_especial.append({
                "Seq": v["Seq"], "Empresa": v["Nome_Empresa"],
                "Periodo": f"{v['Data_Admissao']} a {v['Data_Demissao'] or v['Ultima_Remun']}",
                "Dias_Ate_13_11_2019": dias_pre, "Fator_Conversao": fator_conv,
                "Dias_Extras_Ganhos_Com_PPP": dias_bonus,
                "Auditoria_Recolhimento_Risco_GILRAT": "Conferir Código Ocorrência 01 a 04 na GFIP/eSocial (adicional GILRAT 6%, 9% ou 12%).",
            })

    if not linhas_especial:
        linhas_especial.append({
            "Seq": "-", "Empresa": "Nenhum período operacional com exposição a risco selecionado",
            "Periodo": "-", "Dias_Ate_13_11_2019": 0, "Fator_Conversao": fator_conv,
            "Dias_Extras_Ganhos_Com_PPP": 0,
            "Auditoria_Recolhimento_Risco_GILRAT": "Como o ambiente informado foi Administrativo/Escritório, não há acréscimo de tempo especial.",
        })

    novo_total_dias_especial = dias_liquidos + total_dias_bonus_especial
    df_especial = pd.DataFrame(linhas_especial)

    metas_pcd = [
        ("Deficiência Grave", 25 if sexo == "M" else 20),
        ("Deficiência Moderada", 29 if sexo == "M" else 24),
        ("Deficiência Leve (Ex: TEA Nível 1)", 33 if sexo == "M" else 28),
        ("Regra Comum Pós-Reforma", 35 if sexo == "M" else 30),
    ]

    linhas_pcd = []
    for grau, anos_exigidos in metas_pcd:
        dias_exigidos = int(round(anos_exigidos * 365.25))
        dias_falt_comum = max(0, dias_exigidos - dias_liquidos)
        dias_falt_esp = max(0, dias_exigidos - novo_total_dias_especial)
        linhas_pcd.append({
            "Modalidade": grau, "Tempo_Exigido": f"{anos_exigidos} anos",
            "Falta_Hoje_Tempo_Comum": "CUMPRIDO!" if dias_falt_comum == 0 else formatar_ymd(dias_falt_comum),
            "Falta_Hoje_Com_Bonus_PPP": "CUMPRIDO!" if dias_falt_esp == 0 else formatar_ymd(dias_falt_esp),
        })
    df_pcd = pd.DataFrame(linhas_pcd)

    teto_2026 = 8475.55
    piso_2026 = 1621.00
    salarios_corrigidos = []

    for comp, val_orig in sorted(remuneracoes_dict.items(), key=lambda x: list(reversed(next(iter(x)).split("/")))):
        mes_c, ano_c = map(int, comp.split("/"))
        teto_epoca = TETOS_INSS.get(ano_c, teto_2026)
        val_limitado = min(val_orig, teto_epoca)
        fator = FATOR_ATUALIZACAO_APROX.get(ano_c, 1.0)
        val_atual = min(val_limitado * fator, teto_2026)
        salarios_corrigidos.append(val_atual)

    if salarios_corrigidos:
        media_100 = sum(salarios_corrigidos) / len(salarios_corrigidos)
        salarios_ord = sorted(salarios_corrigidos, reverse=True)
        qtd_80 = max(1, int(round(len(salarios_ord) * 0.8)))
        media_80_maiores = sum(salarios_ord[:qtd_80]) / qtd_80
    else:
        media_100 = 0.0
        media_80_maiores = 0.0

    anos_contrib_atual = int(dias_liquidos // 365.25)
    excedente_anos = max(0, anos_contrib_atual - (20 if sexo == "M" else 15))
    coef_reforma = min(1.0, 0.60 + (0.02 * excedente_anos))

    rmi_regra_comum_hoje = max(piso_2026, min(teto_2026, media_100 * coef_reforma)) if media_100 > 0 else 0.0
    rmi_regra_pcd_100 = max(piso_2026, min(teto_2026, media_80_maiores)) if media_80_maiores > 0 else 0.0

    df_financeiro = pd.DataFrame([
        {"Metrica_Financeira": "Competências Extraídas (Desde 07/1994)", "Valor_Estimado": f"{len(salarios_corrigidos)} meses"},
        {"Metrica_Financeira": "Teto Oficial INSS Vigente (2026)", "Valor_Estimado": f"R$ {teto_2026:,.2f}"},
        {"Metrica_Financeira": "Média de 100% dos Salários (Regra Geral)", "Valor_Estimado": f"R$ {media_100:,.2f}"},
        {"Metrica_Financeira": "Média dos 80% Maiores Salários (Regra PCD LC 142)", "Valor_Estimado": f"R$ {media_80_maiores:,.2f}"},
        {"Metrica_Financeira": f"Estimativa RMI Regra Comum (Coeficiente {int(coef_reforma*100)}%)", "Valor_Estimado": f"R$ {rmi_regra_comum_hoje:,.2f}"},
        {"Metrica_Financeira": "Estimativa RMI Aposentadoria PCD (100% da Média)", "Valor_Estimado": f"R$ {rmi_regra_pcd_100:,.2f}"},
        {"Metrica_Financeira": "Diferença Mensal a Favor da Regra PCD", "Valor_Estimado": f"+ R$ {max(0, rmi_regra_pcd_100 - rmi_regra_comum_hoje):,.2f} / mês"},
    ])

    erros_lista = [v for v in vinculos if "ERRO" in v["Status_Auditoria"] or "ALERTA" in v["Status_Auditoria"]]
    texto_parecer = f"PARECER TÉCNICO PREVIDENCIÁRIO E AUDITORIA DE CNIS\nSegurado(a): {nome_segurado} | Idade Atual: {idade_atual}\n\n1. TEMPO LÍQUIDO DE CONTRIBUIÇÃO\nTempo líquido comprovado hoje: {formatar_ymd(dias_liquidos)} ({dias_liquidos} dias), já descontados {dias_concomitantes} dias concomitantes.\n\n2. PENDÊNCIAS NO MEU INSS\n"
    if erros_lista:
        for err in erros_lista:
            texto_parecer += f"- Vínculo {err['Seq']} ({err['Nome_Empresa']}): {err['Status_Auditoria']}. Ação: {err['Plano_Acao_INSS']}\n"
    else:
        texto_parecer += "- Nenhuma pendência impeditiva crítica encontrada.\n"

    texto_parecer += "\n3. ANÁLISE ESPECIAL E PCD (LC 142/2013)\n"
    if ambiente_operacional and total_dias_bonus_especial > 0:
        texto_parecer += f"- Conversão Especial (PPP): Acréscimo estimado de {formatar_ymd(total_dias_bonus_especial)}. Solicitar PPP/LTCAT e conferir código de ocorrência (01 a 04) na GFIP/eSocial.\n"
    else:
        texto_parecer += "- Conversão Especial (PPP): Perfil Administrativo/Escritório selecionado (sem acréscimo automático de insalubridade/periculosidade).\n"

    if possui_pcd:
        texto_parecer += f"- Regra PCD (LC 142/2013): Direito à redução de tempo (25, 29 ou 33 anos H / 20, 24 ou 28 anos M) sem idade mínima e com média dos 80% maiores salários (100%), elevando a estimativa de R$ {rmi_regra_comum_hoje:,.2f} para R$ {rmi_regra_pcd_100:,.2f} (+ R$ {max(0, rmi_regra_pcd_100 - rmi_regra_comum_hoje):,.2f}/mês).\n"

    df_parecer = pd.DataFrame([{"Parecer_Tecnico_Completo": texto_parecer}])

    with pd.ExcelWriter(caminho_excel_saida, engine="openpyxl") as writer:
        abas_dados = [
            ("1. 📋 CNIS EXTRAÍDO", df_vinculos),
            ("2. ⏱️ TEMPO LÍQUIDO", df_resumo),
            ("3. ⚠️ APOSENT. ESPECIAL", df_especial),
            ("4. ♿ REGRA PCD (LC 142)", df_pcd),
            ("5. 💰 VALOR ESTIMADO (RMI)", df_financeiro),
            ("6. ✅ PARECER FINAL", df_parecer),
        ]

        fill_header = PatternFill(start_color="0F293A", end_color="0F293A", fill_type="solid")
        font_header = Font(name="Calibri", size=11, bold=True, color="F3D06B")
        fill_zebra = PatternFill(start_color="FFFDF5", end_color="FFFDF5", fill_type="solid")
        font_dados = Font(name="Calibri", size=10, color="1E293B")
        borda_fina = Side(border_style="thin", color="CBD5E1")
        borda_grade = Border(left=borda_fina, right=borda_fina, top=borda_fina, bottom=borda_fina)

        for nome_aba, df_atual in abas_dados:
            df_atual.to_excel(writer, index=False, sheet_name=nome_aba, startrow=1, startcol=1)
            ws = writer.sheets[nome_aba]
            ws.views.sheetView[0].showGridLines = False
            ws.column_dimensions["A"].width = 4
            ws.row_dimensions[1].height = 14
            ws.row_dimensions[2].height = 26

            for col_idx in range(2, len(df_atual.columns) + 2):
                cell_h = ws.cell(row=2, column=col_idx)
                cell_h.fill = fill_header
                cell_h.font = font_header
                cell_h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell_h.border = borda_grade

            for row_idx in range(3, len(df_atual) + 3):
                max_linhas = 1
                for col_idx in range(2, len(df_atual.columns) + 2):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.font = font_dados
                    cell.border = borda_grade
                    if row_idx % 2 == 1:
                        cell.fill = fill_zebra
                    alin_h = "left" if nome_aba == "6. ✅ PARECER FINAL" else "center"
                    cell.alignment = Alignment(horizontal=alin_h, vertical="center", wrap_text=True)

                    texto_cel = str(cell.value or "")
                    linhas_texto = texto_cel.count("\n") + 1 + (len(texto_cel) // 65)
                    max_linhas = max(max_linhas, linhas_texto)

                if nome_aba == "6. ✅ PARECER FINAL":
                    ws.row_dimensions[row_idx].height = max(220, max_linhas * 16)
                else:
                    ws.row_dimensions[row_idx].height = max(22, max_linhas * 16)

            for col_idx in range(2, len(df_atual.columns) + 2):
                col_letter = ws.cell(row=2, column=col_idx).column_letter
                if nome_aba == "6. ✅ PARECER FINAL":
                    ws.column_dimensions[col_letter].width = 95
                else:
                    max_len = max(
                        (len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(2, len(df_atual) + 3)),
                        default=12
                    )
                    ws.column_dimensions[col_letter].width = min(max(max_len + 5, 16), 48)

        # Colorir as abas inferiores e destacar o dossiê
        cores_abas = {
            "1. 📋 CNIS EXTRAÍDO": "0F293A",        # Azul-Petróleo
            "2. ⏱️ TEMPO LÍQUIDO": "C59B27",        # Dourado
            "3. ⚠️ APOSENT. ESPECIAL": "D97706",    # Laranja/Âmbar
            "4. ♿ REGRA PCD (LC 142)": "1D4ED8",   # Azul Royal
            "5. 💰 VALOR ESTIMADO (RMI)": "15803D", # Verde Esmeralda
            "6. ✅ PARECER FINAL": "7F1D1D",        # Bordô/Vinho
        }

        for titulo_aba, cor_hex in cores_abas.items():
            if titulo_aba in writer.sheets:
                writer.sheets[titulo_aba].sheet_properties.tabColor = cor_hex

        # Define a aba de Tempo Líquido como a primeira a abrir no Excel
        writer.book.active = writer.sheets["2. ⏱️ TEMPO LÍQUIDO"]

    return texto_parecer
