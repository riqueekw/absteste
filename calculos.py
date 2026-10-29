"""Regras de negócio dos orçamentos. Não usa Streamlit: dá pra testar e reaproveitar em qualquer lugar.

Preços, taxas e constantes NÃO ficam aqui: vêm do parametros.xlsx (objeto `par`, ver dados.py).
"""
import math

import pandas as pd

# Nomes das constantes na tabela "Constantes" do parametros.xlsx
K_LOTE = "Lote econômico (pç)"
K_ABS_CICLO = "ABS: ciclo de termoformagem (min/chapa)"
K_ABS_CC_TERMO = "ABS: centro de custo da termoformagem"
K_ABS_CC_ACAB = "ABS: centro de custo do acabamento"
K_ABS_PREP = "ABS: preparação (fração do tempo de termo)"
K_ABS_DESTACAR = "ABS: destacar (min por m de perímetro)"
K_ABS_RECORTAR = "ABS: recortar (min por m de perímetro)"
K_ABS_ACABAMENTO = "ABS: acabamento (min por m de perímetro)"
K_CHAPA_C = "ABS: chapa - comprimento (m)"
K_CHAPA_L = "ABS: chapa - largura (m)"
K_PINT_HORA = "Pintura: valor hora (R$/h)"
K_PINT_CONSUMO = "Pintura: consumo de tinta (L/mm², um lado)"
K_PINT_TEMPO = "Pintura: tempo dos dois lados (min/mm²)"
K_MONT_CC = "Montagem: centro de custo padrão"

POLIMEROS = {
    "ABS-s": ("ABS", 1.06),
    "ABS-PMMA": ("ABS-PMMA", 1.18),
    "ABS-TPE": ("ABS-TPE", 1.11),
}

CORES = sorted([
    "EVO1", "EVO2", "EVO4",
    "ECO1", "ECO2", "ECO3",
    "PRETO", "BLACK PIANO", "BRANCO GEADA",
])

DENSIDADES_PESO = {"ABS": 1.06, "PEAD": 0.96, "PP": 0.95}


def _col(df, nome):
    return pd.to_numeric(df[nome], errors="coerce")


def _txt(v):
    return "" if pd.isna(v) else str(v)


# -----------------------------------------------------------------------------
# ABS (termoformagem)
# -----------------------------------------------------------------------------

def calcular_abs(par, cod_item, comprimento, largura, espessura, cor,
                 peso_liquido, polimero_label, tipo_corte, conceito):
    """Devolve (linha para salvar, custo por etapa)."""
    k = par.k
    polimero, densidade = POLIMEROS[polimero_label]
    cto, adc = (200, 0.3) if conceito == "Mesa" else (100, 0.0)  # Mesa / Caixa

    blank_c, blank_l = comprimento + cto, largura + cto
    comp, lng = blank_c / 1000, blank_l / 1000
    perimetro = 2 * comp + 2 * lng
    peso_bruto = blank_c * blank_l * espessura * densidade * (1 + adc) / 1_000_000

    chapa_c, chapa_l = k[K_CHAPA_C], k[K_CHAPA_L]
    qtd = max(math.floor(chapa_c / comp) * math.floor(chapa_l / lng),
              math.floor(chapa_c / lng) * math.floor(chapa_l / comp))
    if qtd <= 0:
        raise ValueError("Não é possível encaixar nenhuma peça na chapa com essas medidas.")

    taxa_termo = par.taxa(k[K_ABS_CC_TERMO])
    taxa_acab = par.taxa(k[K_ABS_CC_ACAB])

    termo = k[K_ABS_CICLO] / qtd
    prep = termo * k[K_ABS_PREP]
    destacar = perimetro * k[K_ABS_DESTACAR]
    recortar = perimetro * k[K_ABS_RECORTAR]
    acabamento = perimetro * k[K_ABS_ACABAMENTO]

    custos = {
        "Termoformagem": termo * taxa_termo / 60,
        "Preparação": prep * taxa_acab / 60,
        "Destacar": destacar * taxa_acab / 60,
        "Recortar": recortar * taxa_acab / 60,
        "Acabamento": acabamento * taxa_acab / 60,
    }
    custo_total = sum(custos.values())

    item = {
        "Código Item": cod_item,
        "Polímero": polimero,
        "Cor": cor,
        "Tipo Corte": tipo_corte,
        "Conceito": conceito,
        "Blank (mm)": f"{blank_c} x {blank_l}",
        "Peso Líquido (kg)": round(peso_liquido, 3),
        "Peso Bruto (kg)": round(peso_bruto, 3),
        "Qtd Peças/Chapa": qtd,
        "Tempo Termo (min)": round(termo, 2),
        "Tempo Prep (min)": round(prep, 2),
        "Tempo Destacar (min)": round(destacar, 2),
        "Tempo Recortar (min)": round(recortar, 2),
        "Tempo Acabamento (min)": round(acabamento, 2),
        "Custo Recorte (R$)": round(custos["Recortar"], 2),
        "Custo Total (R$)": round(custo_total, 2),
    }
    return item, {nome: round(v, 2) for nome, v in custos.items()}

def recomendar_chapa(par, comprimento, largura, espessura, cor, polimero_label, conceito):
    """Melhor chapa cadastrada (aba Parâmetros > Chapas ABS) para o blank da peça.

    Filtra por cor, polímero (célula em branco = serve para qualquer um) e espessura exata,
    e só mantém as chapas onde o blank cabe. Critério: maior aproveitamento da chapa;
    empate -> mais peças por chapa. Devolve (tabela ordenada, aviso); tabela vazia = sem chapa.
    """
    cto = 200 if conceito == "Mesa" else 100
    comp, lng = (comprimento + cto) / 1000, (largura + cto) / 1000

    ch = par.chapas.dropna(subset=["Código", "Comprimento (mm)", "Largura (mm)", "Espessura (mm)"])
    ch = ch.drop_duplicates("Código")
    ch = ch[ch["Cor"].astype(str).str.strip().str.upper() == str(cor).strip().upper()]
    pol = ch["Polímero"]
    ch = ch[pol.isna() | (pol.astype(str).str.strip() == "") | (pol == polimero_label)]
    if ch.empty:
        return pd.DataFrame(), f"Nenhuma chapa cadastrada para a cor {cor} com {polimero_label}."

    f = lambda v: f"{v:g}".replace(".", ",")
    disponiveis = ", ".join(f(e) for e in sorted(ch["Espessura (mm)"].unique()))
    ch = ch[(ch["Espessura (mm)"] - espessura).abs() < 0.01]
    if ch.empty:
        return pd.DataFrame(), (f"Não há chapa de {f(espessura)} mm para {cor} ({polimero_label}). "
                                f"Espessuras cadastradas: {disponiveis} mm.")

    linhas = []
    for _, x in ch.iterrows():
        c, l = x["Comprimento (mm)"] / 1000, x["Largura (mm)"] / 1000
        qtd = max(math.floor(c / comp) * math.floor(l / lng), math.floor(c / lng) * math.floor(l / comp))
        if qtd > 0:
            linhas.append({
                "Código": x["Código"],
                "Descrição": _txt(x["Descrição"]),
                "Chapa (mm)": f'{x["Comprimento (mm)"]:.0f} x {x["Largura (mm)"]:.0f}',
                "Espessura (mm)": x["Espessura (mm)"],
                "Peças/chapa": qtd,
                "Aproveitamento (%)": round(qtd * comp * lng / (c * l) * 100, 1),
            })
    if not linhas:
        return pd.DataFrame(), (f"O blank ({comprimento + cto:.0f} x {largura + cto:.0f} mm) não cabe em "
                                f"nenhuma chapa de {f(espessura)} mm para {cor}.")
    df = pd.DataFrame(linhas).sort_values(["Aproveitamento (%)", "Peças/chapa"], ascending=False, kind="stable")
    return df.reset_index(drop=True), ""

# -----------------------------------------------------------------------------
# INSERTOS (chapas, operações, terceiros)
# -----------------------------------------------------------------------------

def calcular_insertos(par, mp, terc, ops, ferram, ensaios, lote):
    """Devolve (resumo de custos, tabelas de detalhe).

    Setup é rateado no lote econômico; tempo de operação multiplica pelo nº de homens.
    Ferramental e ensaios são mostrados à parte (não entram no custo da peça).
    """
    lote = max(float(lote), 1.0)

    mp = mp.dropna(subset=["Material"]).copy()
    cat = par.materiais.drop_duplicates("Descrição").set_index("Descrição")
    mp["Un"] = mp["Material"].map(cat["Un"])
    preco_tabela = mp["Material"].map(cat["Preço (R$)"])
    mp["Preço unit. (R$)"] = _col(mp, "Preço unit. (R$)").fillna(preco_tabela).fillna(0.0)
    mp["Total (R$)"] = _col(mp, "Qtd").fillna(0.0) * mp["Preço unit. (R$)"]

    terc = terc.dropna(subset=["Descrição"]).copy()
    terc["Total (R$)"] = _col(terc, "Qtd").fillna(0.0) * _col(terc, "Preço unit. (R$)").fillna(0.0)

    ops = ops.dropna(subset=["Operação"]).copy()
    cat_ops = par.operacoes.drop_duplicates("Operação").set_index("Operação")
    ops["C.C."] = ops["Operação"].map(cat_ops["C.C."])
    ops["R$/h"] = [par.taxa(cc) for cc in ops["C.C."]]
    homens = _col(ops, "Nº homens").fillna(1.0)
    ops["Custo setup (R$)"] = ops["R$/h"] * _col(ops, "Setup (min)").fillna(0.0) / lote / 60
    ops["Custo tempo (R$)"] = ops["R$/h"] * _col(ops, "Tempo (min)").fillna(0.0) / 60 * homens

    ferram = ferram.dropna(subset=["Descrição"]).copy()
    ensaios = ensaios.dropna(subset=["Descrição"]).copy()

    r = {
        "Matéria-prima (R$)": float(mp["Total (R$)"].sum()),
        "Terceiros (R$)": float(terc["Total (R$)"].sum()),
        "Mão de obra - tempo (R$)": float(ops["Custo tempo (R$)"].sum()),
        "Mão de obra - setup (R$)": float(ops["Custo setup (R$)"].sum()),
    }
    r["Custo da Peça (R$)"] = sum(r.values())
    r["Ferramental (R$)"] = float(_col(ferram, "Valor (R$)").sum())
    r["Ensaios (R$)"] = float(_col(ensaios, "Valor (R$)").sum())
    return r, dict(mp=mp, terc=terc, ops=ops, ferram=ferram, ensaios=ensaios)


def linha_insertos(cod, descricao, lote, r):
    linha = {"Código Item": cod, "Descrição": descricao, "Lote Econômico (pç)": int(lote)}
    linha.update({nome: round(v, 2) for nome, v in r.items()})
    return linha


def detalhe_insertos(cod, det):
    """Composição do orçamento em formato 'longo' (uma linha por item) para a aba de detalhe."""
    linhas = []

    def add(bloco, desc, qtd, un, preco, total):
        linhas.append({"Código Item": cod, "Bloco": bloco, "Descrição": desc, "Qtd / Tempo": qtd,
                       "Un": un, "Preço / Taxa (R$)": preco, "Total (R$)": round(float(total), 2)})

    for _, x in det["mp"].iterrows():
        add("Matéria-prima", x["Material"], x["Qtd"], _txt(x["Un"]), x["Preço unit. (R$)"], x["Total (R$)"])
    for _, x in det["terc"].iterrows():
        forn = _txt(x.get("Fornecedor"))
        add("Terceiros", f'{x["Descrição"]}' + (f" ({forn})" if forn else ""),
            x["Qtd"], "", x["Preço unit. (R$)"], x["Total (R$)"])
    for _, x in det["ops"].iterrows():
        setup = 0 if pd.isna(x["Setup (min)"]) else x["Setup (min)"]
        homens = 1 if pd.isna(x["Nº homens"]) else x["Nº homens"]
        add("Operação", f'{x["Operação"]} (C.C. {int(float(x["C.C."]))}; setup {setup} min; {homens} h)',
            x["Tempo (min)"], "min", x["R$/h"], x["Custo setup (R$)"] + x["Custo tempo (R$)"])
    for bloco, chave in (("Ferramental", "ferram"), ("Ensaio", "ensaios")):
        for _, x in det[chave].iterrows():
            add(bloco, x["Descrição"], "", "", "", x["Valor (R$)"])
    return linhas


# -----------------------------------------------------------------------------
# PINTURA
# -----------------------------------------------------------------------------

def calcular_pintura(par, cod_item, cod_marcopolo, comprimento, largura, area_manual, tinta, lados):
    """Custo = tinta (L x R$/L) + tempo de cabine (h x R$/h), por lado, x nº de lados.

    Use a medida REAL da peça (sem os 200 mm de cada lado do blank da termoformagem).
    """
    k = par.k
    area = area_manual if area_manual > 0 else comprimento * largura  # mm²
    preco = par.preco_tinta(tinta)

    tinta_l = area * k[K_PINT_CONSUMO]
    tempo_lado = area * k[K_PINT_TEMPO] / 2  # o parâmetro é o tempo dos dois lados
    custo_lado = tinta_l * preco + tempo_lado / 60 * k[K_PINT_HORA]
    custo = custo_lado * lados

    return {
        "Código Item": cod_item,
        "Código Marcopolo": cod_marcopolo,
        "Tinta": tinta,
        "Lados": lados,
        "Área (m²)": round(area / 1_000_000, 4),
        "Tinta (L)": round(tinta_l * lados, 4),
        "Tempo Pintura (min)": round(tempo_lado * lados, 2),
        "Custo 1 Lado (R$)": round(custo_lado, 2),
        "Custo Total (R$)": round(custo, 2),
    }


# -----------------------------------------------------------------------------
# MONTAGEM e PESO
# -----------------------------------------------------------------------------

def calcular_montagem(par, cod_item, modo, tempo, homens, cc, valor_direto):
    if modo == "Valor direto (R$)":
        custo, taxa = valor_direto, None
    else:
        taxa = par.taxa(cc)
        custo = taxa * tempo / 60 * homens
    return {
        "Código Item": cod_item,
        "Modo": modo,
        "Centro de Custo": cc if taxa is not None else "",
        "Tempo (min)": tempo if taxa is not None else "",
        "Nº Homens": homens if taxa is not None else "",
        "Custo Montagem (R$)": round(custo, 2),
    }


def calcular_peso(espessura, largura, comprimento, tipo):
    """mm x mm x mm x g/cm³ / 1.000.000 = kg. Devolve (peso, peso +30%)."""
    peso = espessura * largura * comprimento * DENSIDADES_PESO[tipo] / 1_000_000
    return peso, peso * 1.3
