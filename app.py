# =============================================================================
# ORÇAMENTOS  |  rodar com:  streamlit run app.py
# app.py = telas | calculos.py = regras | dados.py = Excel (parâmetros e orçamentos)
# =============================================================================
import importlib
import re

import pandas as pd
import streamlit as st

import calculos as calc
import dados as d
import ficha as fi

VERSAO = "2026-10-31.2"

# recarrega os módulos a cada execução (evita ficar com versão antiga na memória)
for _mod in (calc, d, fi):
    importlib.reload(_mod)

st.set_page_config(page_title="Orçamentos", page_icon="🧮", layout="wide")

_versoes = {"app.py": VERSAO, "calculos.py": getattr(calc, "VERSAO", "antiga"),
            "dados.py": getattr(d, "VERSAO", "antiga"), "ficha.py": getattr(fi, "VERSAO", "antiga")}
if len(set(_versoes.values())) > 1:
    st.error("Os arquivos do programa estão em versões diferentes. Substitua os desatualizados pelos da última versão "
             "(todos na mesma pasta) e reinicie o app.")
    st.table(pd.DataFrame({"Arquivo": list(_versoes), "Versão": list(_versoes.values())}).set_index("Arquivo"))
    st.stop()

st.markdown("""
<style>
.block-container {padding-top: 2.5rem; max-width: 1300px;}
[data-testid="stMetric"] {
    background: rgba(128,128,128,0.08);
    border: 1px solid rgba(128,128,128,0.25);
    border-radius: 10px;
    padding: 12px 16px;
}
[data-testid="stMetricLabel"] {opacity: 0.75;}
</style>
""", unsafe_allow_html=True)

par = d.carregar_parametros()

if _msg := st.session_state.pop("_msg", None):
    st.toast(_msg, icon="✅")


# ------------------------------- utilidades ----------------------------------

def num(v, casas=2):
    return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def brl(v):
    return "R$ " + num(v)


def iniciar(padrao):
    for chave, valor in padrao.items():
        st.session_state.setdefault(chave, valor)


def resetar(padrao, versao=None):
    for chave, valor in padrao.items():
        st.session_state[chave] = valor
    if versao:  # muda a chave das tabelas editáveis para limpá-las
        st.session_state[versao] = st.session_state.get(versao, 0) + 1


def salvar(aba, cod, linhas, ok):
    if not cod:
        st.warning("Informe o código do item.")
    elif d.codigo_existe(aba, cod):
        st.warning(f"Código **{cod}** já cadastrado nesta aba. Use outro ou exclua o anterior no Histórico.")
    else:
        try:
            d.salvar_linhas(linhas)
            st.toast(f"Orçamento {cod} salvo!", icon="✅")
            st.success(ok)
        except PermissionError:
            st.error("Não consegui salvar: feche o orcamentos.xlsx no Excel e tente de novo.")
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")


def botoes(pref, padrao, versao=None):
    b1, b2, _ = st.columns([1.3, 1, 1.5])
    clicou = b1.button("💾 Salvar orçamento", type="primary", key=f"{pref}_salvar")
    b2.button("Limpar", key=f"{pref}_limpar", on_click=resetar, args=(padrao, versao))
    return clicou


def tabela_vazia(**colunas):
    return pd.DataFrame({nome: pd.Series(dtype=tipo) for nome, tipo in colunas.items()})


def cfg_num(rotulo, **kw):
    return st.column_config.NumberColumn(rotulo, min_value=0.0, **kw)


materiais = par.materiais["Descrição"].dropna().tolist()
operacoes = par.operacoes["Operação"].dropna().tolist()
tintas = par.tintas["Tinta"].dropna().tolist()
cores = sorted(set(calc.CORES) | set(par.chapas["Cor"].dropna().astype(str)))
centros = [int(float(x)) for x in par.centros["Código"].dropna()]
nome_cc = {int(float(cod)): f"{int(float(cod))} · {desc} ({brl(float(v))}/h)"
           for cod, desc, v in zip(par.centros["Código"], par.centros["Descrição"], par.centros["R$/h"])
           if pd.notna(cod) and pd.notna(v)}

st.markdown("## 🧮 Orçamentos")
tab_abs, tab_ins, tab_pin, tab_mon, tab_peso, tab_aprov, tab_ficha, tab_hist, tab_par = st.tabs(
    ["ABS", "Insertos", "Pintura", "Montagem", "Peso", "Aproveitamento", "Ficha", "Histórico", "Parâmetros"])

# =============================== ABS =========================================
with tab_abs:
    PAD = {"abs_cod": "", "abs_desc": "", "abs_lote": int(par.k[calc.K_LOTE]), "abs_comp": 1000, "abs_larg": 1000, "abs_esp": 3.0, "abs_cor": calc.CORES[0],
           "abs_peso": 1.0, "abs_pol": list(calc.POLIMEROS)[0], "abs_corte": "Corte Robô", "abs_conc": "Mesa"}
    iniciar(PAD)
    s = st.session_state
    col_f, col_r = st.columns([1.1, 1], gap="large")

    def params_abs():
        return dict(cod_item=s["abs_cod"].strip(), descricao=s["abs_desc"].strip(), lote=s["abs_lote"],
                    comprimento=s["abs_comp"], largura=s["abs_larg"],
                    espessura=s["abs_esp"], cor=s["abs_cor"], peso_liquido=s["abs_peso"],
                    polimero_label=s["abs_pol"], tipo_corte=s["abs_corte"], conceito=s["abs_conc"])

    with col_f:
        with st.container(border=True):
            st.markdown("**Identificação**")
            i1, i2, i3 = st.columns([1, 2, 1])
            i1.text_input("Código do item", key="abs_cod", placeholder="Código único do item")
            i2.text_input("Descrição da peça", key="abs_desc")
            i3.number_input("Lote econômico (pç)", min_value=1, step=10, key="abs_lote")
        with st.container(border=True):
            st.markdown("**Dimensões**")
            c1, c2, c3 = st.columns(3)
            c1.number_input("Comprimento (mm)", min_value=1, max_value=2700, step=1, key="abs_comp")
            c2.number_input("Largura (mm)", min_value=1, max_value=1700, step=1, key="abs_larg")
            c3.number_input("Espessura (mm)", min_value=3.0, max_value=7.0, step=0.1, format="%.1f", key="abs_esp")
        with st.container(border=True):
            st.markdown("**Material**")
            c4, c5, c6 = st.columns(3)
            c4.selectbox("Cor", cores, key="abs_cor")
            c5.selectbox("Polímero", list(calc.POLIMEROS), key="abs_pol")
            c6.number_input("Peso líquido (kg)", min_value=0.001, max_value=1000.0, step=0.1,
                            format="%.3f", key="abs_peso")
        with st.container(border=True):
            st.markdown("**Processo**")
            c7, c8 = st.columns(2)
            c7.radio("Tipo de corte", ["Corte Robô", "Corte Manual"], key="abs_corte", horizontal=True)
            c8.radio("Conceito", ["Mesa", "Caixa"], key="abs_conc", horizontal=True)

        if botoes("abs", PAD):
            p = params_abs()
            try:
                item, _ = calc.calcular_abs(par, **p)
                rec, _ = calc.recomendar_chapa(par, p["comprimento"], p["largura"], p["espessura"],
                                               p["cor"], p["polimero_label"], p["conceito"])
                item["Código Chapa"] = rec.iloc[0]["Código"] if not rec.empty else ""
                salvar("Orçamentos", p["cod_item"], {"Orçamentos": [item]},
                       f"Orçamento **{p['cod_item']}** salvo: {brl(item['Custo Total (R$)'])}")
            except Exception as e:
                st.error(f"Erro no cálculo: {e}")

    with col_r:
        st.markdown("**Chapa recomendada**")
        rec, aviso = calc.recomendar_chapa(par, s["abs_comp"], s["abs_larg"], s["abs_esp"],
                                           s["abs_cor"], s["abs_pol"], s["abs_conc"])
        if rec.empty:
            st.warning(aviso)
        else:
            melhor = rec.iloc[0]
            st.code(melhor["Código"], language=None)
            st.caption(f'{melhor["Descrição"]} · {melhor["Chapa (mm)"]} mm · {num(melhor["Espessura (mm)"])} mm de espessura')
            n1, n2 = st.columns(2)
            n1.metric("Peças por chapa", int(melhor["Peças/chapa"]))
            n2.metric("Aproveitamento", num(melhor["Aproveitamento (%)"], 1) + "%")
            if len(rec) > 1:
                with st.expander("Outras opções"):
                    st.dataframe(rec, hide_index=True)
        st.divider()
        st.markdown("**Prévia do cálculo**")
        try:
            prev, custos = calc.calcular_abs(par, **params_abs())
        except Exception as e:
            prev = None
            st.warning(str(e))
        if prev:
            m1, m2 = st.columns(2)
            m1.metric("Custo total", brl(prev["Custo Total (R$)"]))
            m2.metric("Peças por chapa", prev["Qtd Peças/Chapa"])
            m3, m4 = st.columns(2)
            m3.metric("Peso bruto", num(prev["Peso Bruto (kg)"], 3) + " kg")
            m4.metric("Blank (mm)", prev["Blank (mm)"])
            st.markdown("**Custo por etapa (R$)**")
            st.bar_chart(pd.DataFrame({"Etapa": list(custos), "Custo (R$)": list(custos.values())}).set_index("Etapa"))
            with st.expander("Ver todos os campos"):
                st.table(pd.DataFrame([(k, str(v)) for k, v in prev.items()], columns=["Campo", "Valor"]).set_index("Campo"))

# ============================== INSERTOS =====================================
with tab_ins:
    PAD = {"ins_cod": "", "ins_desc": "", "ins_lote": int(par.k[calc.K_LOTE])}
    iniciar(PAD)
    st.session_state.setdefault("ver_ins", 0)
    s, v = st.session_state, st.session_state["ver_ins"]
    col_f, col_r = st.columns([1.4, 1], gap="large")

    with col_f:
        with st.container(border=True):
            st.markdown("**Identificação**")
            i1, i2, i3 = st.columns([1, 2, 1])
            i1.text_input("Código do item", key="ins_cod")
            i2.text_input("Descrição da peça", key="ins_desc")
            i3.number_input("Lote econômico (pç)", min_value=1, step=10, key="ins_lote")

        with st.container(border=True):
            st.markdown("**Matéria-prima**")
            st.caption("Escolha o material e a quantidade (na unidade do material: kg, M², L). "
                       "Preço vazio = usa o da aba Parâmetros.")
            mp = st.data_editor(
                tabela_vazia(**{"Material": "object", "Qtd": "float64", "Preço unit. (R$)": "float64"}),
                num_rows="dynamic", hide_index=True, key=f"ins_mp_{v}",
                column_config={"Material": st.column_config.SelectboxColumn("Material", options=materiais, width="large"),
                               "Qtd": cfg_num("Qtd", format="%.4f"),
                               "Preço unit. (R$)": cfg_num("Preço unit. (R$)", format="%.4f")})

        with st.container(border=True):
            st.markdown("**Operações (roteiro)**")
            st.caption("O centro de custo vem da operação. Setup é dividido pelo lote; tempo é multiplicado pelo nº de homens.")
            ops = st.data_editor(
                tabela_vazia(**{"Operação": "object", "Setup (min)": "float64", "Tempo (min)": "float64", "Nº homens": "float64"}),
                num_rows="dynamic", hide_index=True, key=f"ins_ops_{v}",
                column_config={"Operação": st.column_config.SelectboxColumn("Operação", options=operacoes, width="large"),
                               "Setup (min)": cfg_num("Setup (min)"), "Tempo (min)": cfg_num("Tempo (min)"),
                               "Nº homens": cfg_num("Nº homens", help="Vazio = 1")})

        with st.expander("Terceiros, ferramental e ensaios (opcional)"):
            st.markdown("**Operações externas (terceiros)**")
            terc = st.data_editor(
                tabela_vazia(**{"Descrição": "object", "Fornecedor": "object", "Qtd": "float64", "Preço unit. (R$)": "float64"}),
                num_rows="dynamic", hide_index=True, key=f"ins_terc_{v}",
                column_config={"Qtd": cfg_num("Qtd"), "Preço unit. (R$)": cfg_num("Preço unit. (R$)", format="%.4f")})
            st.markdown("**Ferramental** (fora do custo da peça)")
            ferram = st.data_editor(
                tabela_vazia(**{"Descrição": "object", "Valor (R$)": "float64"}),
                num_rows="dynamic", hide_index=True, key=f"ins_ferr_{v}",
                column_config={"Valor (R$)": cfg_num("Valor (R$)")})
            st.markdown("**Ensaios** (fora do custo da peça)")
            ensaios = st.data_editor(
                tabela_vazia(**{"Descrição": "object", "Valor (R$)": "float64"}),
                num_rows="dynamic", hide_index=True, key=f"ins_ens_{v}",
                column_config={"Valor (R$)": cfg_num("Valor (R$)")})

        try:
            resumo, det = calc.calcular_insertos(par, mp, terc, ops, ferram, ensaios, s["ins_lote"])
            erro = None
        except Exception as e:
            resumo = det = None
            erro = str(e)

        if botoes("ins", PAD, "ver_ins"):
            if erro:
                st.error(f"Erro no cálculo: {erro}")
            elif resumo["Custo da Peça (R$)"] <= 0:
                st.warning("Adicione matéria-prima ou operações antes de salvar.")
            else:
                cod = s["ins_cod"].strip()
                salvar("Insertos", cod,
                       {"Insertos": [calc.linha_insertos(cod, s["ins_desc"], s["ins_lote"], resumo)],
                        "Insertos - Detalhe": calc.detalhe_insertos(cod, det)},
                       f"Orçamento **{cod}** salvo: {brl(resumo['Custo da Peça (R$)'])}")

    with col_r:
        st.markdown("**Prévia do cálculo**")
        if erro:
            st.warning(erro)
        else:
            m1, m2 = st.columns(2)
            m1.metric("Custo da peça", brl(resumo["Custo da Peça (R$)"]))
            m2.metric("Matéria-prima", brl(resumo["Matéria-prima (R$)"]))
            m3, m4 = st.columns(2)
            m3.metric("Terceiros", brl(resumo["Terceiros (R$)"]))
            m4.metric("Mão de obra (tempo + setup)",
                      brl(resumo["Mão de obra - tempo (R$)"] + resumo["Mão de obra - setup (R$)"]))
            st.markdown("**Composição do custo (R$)**")
            comp = {"Matéria-prima": resumo["Matéria-prima (R$)"], "Terceiros": resumo["Terceiros (R$)"],
                    "MO - tempo": resumo["Mão de obra - tempo (R$)"], "MO - setup": resumo["Mão de obra - setup (R$)"]}
            st.bar_chart(pd.DataFrame({"Bloco": list(comp), "R$": list(comp.values())}).set_index("Bloco"))
            st.caption(f"Ferramental: {brl(resumo['Ferramental (R$)'])} · Ensaios: {brl(resumo['Ensaios (R$)'])} "
                       "(não entram no custo da peça)")
            if not det["ops"].empty:
                with st.expander("Custo por operação"):
                    st.dataframe(det["ops"][["Operação", "C.C.", "R$/h", "Custo setup (R$)", "Custo tempo (R$)"]],
                                 hide_index=True)

# =============================== PINTURA =====================================
with tab_pin:
    PAD = {"pin_cod": "", "pin_mar": "", "pin_comp": 500.0, "pin_larg": 300.0, "pin_area": 0.0,
           "pin_tinta": tintas[0], "pin_lados": 1}
    iniciar(PAD)
    s = st.session_state
    if s["pin_tinta"] not in tintas:
        s["pin_tinta"] = tintas[0]
    col_f, col_r = st.columns([1.1, 1], gap="large")

    def params_pin():
        return dict(cod_item=s["pin_cod"].strip(), cod_marcopolo=s["pin_mar"].strip(),
                    comprimento=s["pin_comp"], largura=s["pin_larg"], area_manual=s["pin_area"],
                    tinta=s["pin_tinta"], lados=s["pin_lados"])

    with col_f:
        with st.container(border=True):
            st.markdown("**Identificação**")
            p1, p2 = st.columns(2)
            p1.text_input("Código Amalcaburio", key="pin_cod")
            p2.text_input("Código Marcopolo (opcional)", key="pin_mar")
        with st.container(border=True):
            st.markdown("**Peça**")
            st.info("Use a medida **real** da peça, sem os 200 mm de cada lado do blank.")
            q1, q2, q3 = st.columns(3)
            q1.number_input("Comprimento (mm)", min_value=0.0, step=1.0, key="pin_comp")
            q2.number_input("Largura (mm)", min_value=0.0, step=1.0, key="pin_larg")
            q3.number_input("Área manual (mm²)", min_value=0.0, step=100.0, key="pin_area",
                            help="Para peças irregulares. Deixe 0 para usar comprimento × largura.")
        with st.container(border=True):
            st.markdown("**Tinta**")
            r1, r2 = st.columns(2)
            r1.selectbox("Tinta / cor", tintas, key="pin_tinta")
            r2.radio("Lados pintados", [1, 2], key="pin_lados", horizontal=True)

        if botoes("pin", PAD):
            p = params_pin()
            try:
                linha = calc.calcular_pintura(par, **p)
                salvar("Pintura", p["cod_item"], {"Pintura": [linha]},
                       f"Orçamento **{p['cod_item']}** salvo: {brl(linha['Custo Total (R$)'])}")
            except Exception as e:
                st.error(f"Erro no cálculo: {e}")

    with col_r:
        st.markdown("**Prévia do cálculo**")
        try:
            pv = calc.calcular_pintura(par, **params_pin())
        except Exception as e:
            pv = None
            st.warning(str(e))
        if pv:
            m1, m2 = st.columns(2)
            m1.metric("Custo total", brl(pv["Custo Total (R$)"]))
            m2.metric("Custo de 1 lado", brl(pv["Custo 1 Lado (R$)"]))
            m3, m4 = st.columns(2)
            m3.metric("Tinta", num(pv["Tinta (L)"], 4) + " L")
            m4.metric("Tempo de pintura", num(pv["Tempo Pintura (min)"]) + " min")
            st.caption(f"Área: {num(pv['Área (m²)'], 4)} m²")

# ============================== MONTAGEM =====================================
with tab_mon:
    padrao_cc = int(par.k[calc.K_MONT_CC])
    PAD = {"mon_cod": "", "mon_modo": "Tempo × centro de custo", "mon_tempo": 8.0, "mon_homens": 1,
           "mon_cc": padrao_cc if padrao_cc in centros else centros[0], "mon_valor": 0.0}
    iniciar(PAD)
    s = st.session_state
    if s["mon_cc"] not in centros:
        s["mon_cc"] = centros[0]
    col_f, col_r = st.columns([1.1, 1], gap="large")

    def params_mon():
        return dict(cod_item=s["mon_cod"].strip(), modo=s["mon_modo"], tempo=s["mon_tempo"],
                    homens=s["mon_homens"], cc=s["mon_cc"], valor_direto=s["mon_valor"])

    with col_f:
        with st.container(border=True):
            st.markdown("**Identificação**")
            st.text_input("Código do item", key="mon_cod")
        with st.container(border=True):
            st.markdown("**Montagem**")
            st.radio("Como informar", ["Tempo × centro de custo", "Valor direto (R$)"], key="mon_modo", horizontal=True)
            if s["mon_modo"] == "Valor direto (R$)":
                st.number_input("Valor da montagem (R$)", min_value=0.0, step=1.0, key="mon_valor")
            else:
                t1, t2 = st.columns(2)
                t1.number_input("Tempo de montagem (min)", min_value=0.0, step=0.5, key="mon_tempo")
                t2.number_input("Nº de homens", min_value=1, step=1, key="mon_homens")
                st.selectbox("Centro de custo", centros, format_func=lambda x: nome_cc.get(x, str(x)), key="mon_cc")
        if botoes("mon", PAD):
            p = params_mon()
            try:
                linha = calc.calcular_montagem(par, **p)
                salvar("Montagem", p["cod_item"], {"Montagem": [linha]},
                       f"Montagem **{p['cod_item']}** salva: {brl(linha['Custo Montagem (R$)'])}")
            except Exception as e:
                st.error(f"Erro no cálculo: {e}")

    with col_r:
        st.markdown("**Prévia do cálculo**")
        try:
            mv = calc.calcular_montagem(par, **params_mon())
            st.metric("Custo de montagem", brl(mv["Custo Montagem (R$)"]))
            st.caption("Some ao custo da peça na aba Histórico → Resumo por código (mesmo código do item).")
        except Exception as e:
            st.warning(str(e))

# ================================= PESO ======================================
with tab_peso:
    PAD = {"peso_esp": 4.0, "peso_larg": 15.0, "peso_comp": 550.0, "peso_tipo": "ABS"}
    iniciar(PAD)
    s = st.session_state
    col_f, col_r = st.columns([1.1, 1], gap="large")
    with col_f:
        with st.container(border=True):
            st.markdown("**Peso de peça plástica (consulta rápida, não salva)**")
            w1, w2, w3 = st.columns(3)
            w1.number_input("Espessura (mm)", min_value=0.0, step=0.5, key="peso_esp")
            w2.number_input("Largura (mm)", min_value=0.0, step=1.0, key="peso_larg")
            w3.number_input("Comprimento (mm)", min_value=0.0, step=1.0, key="peso_comp")
            st.radio("Material", list(calc.DENSIDADES_PESO), key="peso_tipo", horizontal=True)
    with col_r:
        st.markdown("**Resultado**")
        peso, peso30 = calc.calcular_peso(s["peso_esp"], s["peso_larg"], s["peso_comp"], s["peso_tipo"])
        m1, m2 = st.columns(2)
        m1.metric("Peso da peça", num(peso, 3) + " kg")
        m2.metric("Peso + 30%", num(peso30, 3) + " kg")

# ============================ APROVEITAMENTO =================================
with tab_aprov:
    mat_chapa = par.densidades["Material"].dropna().tolist()
    PAD = {"ap_mat": mat_chapa[0], "ap_esp": 1.5, "ap_cl": 1000.0, "ap_cc": 2810.0,
           "ap_pl": 453.0, "ap_pc": 496.0, "ap_qt": 1}
    iniciar(PAD)
    s = st.session_state
    if s["ap_mat"] not in mat_chapa:
        s["ap_mat"] = mat_chapa[0]
    col_f, col_r = st.columns([1, 1.7], gap="large")

    with col_f:
        with st.container(border=True):
            st.markdown("**Chapa**")
            st.selectbox("Material", mat_chapa, key="ap_mat")
            a1, a2, a3 = st.columns(3)
            a1.number_input("Espessura (mm)", min_value=0.0, step=0.1, key="ap_esp")
            a2.number_input("Largura (mm)", min_value=0.0, step=10.0, key="ap_cl")
            a3.number_input("Comprimento (mm)", min_value=0.0, step=10.0, key="ap_cc")
        with st.container(border=True):
            st.markdown("**Blank a ser cortado**")
            b1, b2, b3 = st.columns(3)
            b1.number_input("Largura (mm)", min_value=0.0, step=1.0, key="ap_pl")
            b2.number_input("Comprimento (mm)", min_value=0.0, step=1.0, key="ap_pc")
            b3.number_input("Peças por blank", min_value=1, step=1, key="ap_qt",
                            help="Ex.: 2 quando saem as peças D/E do mesmo blank")
        st.button("Limpar", key="ap_limpar", on_click=resetar, args=(PAD,))

    with col_r:
        try:
            tot, modelos, melhor = calc.calcular_aproveitamento(
                par, s["ap_mat"], s["ap_esp"], s["ap_cl"], s["ap_cc"], s["ap_pl"], s["ap_pc"], s["ap_qt"])
        except Exception as e:
            modelos = None
            st.warning(str(e))
        if modelos:
            m1, m2, m3 = st.columns(3)
            m1.metric("Peso da chapa", num(tot["Peso chapa (kg)"], 3) + " kg")
            m2.metric("Peso da peça", num(tot["Peso peça (kg)"], 3) + " kg")
            m3.metric("Melhor modelo", melhor or "Não cabe")

            for col, (nome, m) in zip(st.columns(2), modelos.items()):
                with col.container(border=True):
                    st.markdown(f"**{nome}**" + ("  ·  ✅ **É ESSE**" if nome == melhor else ""))
                    st.caption(f'Blank {num(m["Largura (mm)"], 0)} × {num(m["Comprimento (mm)"], 0)} mm (largura × comprimento)')
                    if m["Peças por chapa"] == 0:
                        st.error("Não cabe na chapa.")
                        continue
                    k1, k2 = st.columns(2)
                    k1.metric("Total de peças", m["Peças por chapa"])
                    k2.metric("Aproveitamento", num(m["Aproveitamento (%)"], 1) + "%")
                    st.metric("Consumo por peça (kg)", num(m["Consumo por peça (kg)"], 4))
                    st.metric("Retalho por peça (kg)", num(m["Retalho por peça (kg)"], 4))
                    st.caption(f'{m["Tiras por chapa"]} tiras por chapa × {m["Peças por tira"]} peças por tira')

            if melhor:
                m = modelos[melhor]
                resumo = (
                    f'Chapa {s["ap_mat"]} {num(s["ap_esp"])} x {num(s["ap_cl"], 0)} x {num(s["ap_cc"], 0)} mm '
                    f'({num(tot["Peso chapa (kg)"], 3)} kg). Blank {num(m["Largura (mm)"], 0)} x {num(m["Comprimento (mm)"], 0)} mm, '
                    f'{s["ap_qt"]} peça(s) por blank. {m["Tiras por chapa"]} tiras x {m["Peças por tira"]} peças por tira '
                    f'= {m["Peças por chapa"]} peças por chapa. Peso da peça {num(tot["Peso peça (kg)"], 3)} kg; '
                    f'retalho {num(m["Retalho por peça (kg)"], 3)} kg por peça; aproveitamento {num(m["Aproveitamento (%)"], 1)}%.')
                st.text_area("Resumo para as observações do orçamento", resumo, height=120)

# ================================== FICHA ====================================
with tab_ficha:
    df_abs, df_ins = d.carregar("Orçamentos"), d.carregar("Insertos")
    det_ins, df_pin = d.carregar("Insertos - Detalhe"), d.carregar("Pintura")
    opcoes = ([f"{c} · ABS" for c in df_abs["Código Item"].astype(str)] if not df_abs.empty else []) + \
             ([f"{c} · Insertos" for c in df_ins["Código Item"].astype(str)] if not df_ins.empty else [])

    cc_fi = int(par.k[calc.K_MONT_CC])
    PAD_FI = {"fi_num": "", "fi_cli": "", "fi_cod": "", "fi_desc": "", "fi_cons": 0, "fi_lote": int(par.k[calc.K_LOTE]),
              "fi_sim": "", "fi_mont": True, "fi_tempo": 8.0, "fi_homens": 1,
              "fi_cc": cc_fi if cc_fi in centros else centros[0], "fi_pint": True, "fi_folha_mont": True}
    iniciar(PAD_FI)
    st.session_state.setdefault("ver_fi", 0)
    s, v = st.session_state, st.session_state["ver_fi"]
    if s["fi_cc"] not in centros:
        s["fi_cc"] = centros[0]

    def limpar_ficha():
        resetar(PAD_FI, "ver_fi")
        st.session_state.pop("fi_arquivo", None)

    col_f, col_r = st.columns([1.2, 1], gap="large")
    with col_f:
        with st.container(border=True):
            st.markdown("**Conjunto**")
            f1, f2 = st.columns(2)
            f1.text_input("Nº do orçamento", key="fi_num", placeholder="Ex.: 055-10-2026")
            f2.text_input("Cliente", key="fi_cli")
            f3, f4 = st.columns([1, 2])
            f3.text_input("Código / desenho", key="fi_cod")
            f4.text_input("Descrição da peça", key="fi_desc")
            f5, f6, f7 = st.columns(3)
            f5.number_input("Consumo peças/mês", min_value=0, step=10, key="fi_cons")
            f6.number_input("Lote econômico (pç)", min_value=1, step=10, key="fi_lote")
            f7.text_input("Item similar", key="fi_sim")

        with st.container(border=True):
            st.markdown("**Componentes do conjunto**")
            if not opcoes:
                st.info("Salve orçamentos nas abas ABS ou Insertos para poder montar a ficha.")
            st.caption("Cada linha vira uma folha FICHA. Escolha o orçamento salvo e a quantidade por conjunto (vazio = 1).")
            editor = st.data_editor(
                tabela_vazia(**{"Componente": "object", "Qtd por conjunto": "float64"}),
                num_rows="dynamic", hide_index=True, key=f"fi_comp_{v}",
                column_config={"Componente": st.column_config.SelectboxColumn("Orçamento salvo", options=opcoes, width="large"),
                               "Qtd por conjunto": cfg_num("Qtd por conjunto")})

        with st.container(border=True):
            st.markdown("**Montagem e opções**")
            st.checkbox("Incluir montagem do conjunto", key="fi_mont")
            if s["fi_mont"]:
                g1, g2 = st.columns(2)
                g1.number_input("Tempo de montagem (min)", min_value=0.0, step=0.5, key="fi_tempo")
                g2.number_input("Nº de homens", min_value=1, step=1, key="fi_homens")
                st.selectbox("Centro de custo", centros, format_func=lambda x: nome_cc.get(x, str(x)), key="fi_cc")
            st.checkbox("Levar a pintura salva (aba Pintura, mesmo código) como operação externa", key="fi_pint")
            st.checkbox("Gerar também a folha MONT (conjunto)", key="fi_folha_mont")
        st.button("Limpar", key="fi_limpar", on_click=limpar_ficha)

    # ----- monta os componentes a partir dos orçamentos salvos -----
    comps, erros = [], []
    for _, linha in editor.dropna(subset=["Componente"]).iterrows():
        cod, tipo = linha["Componente"].rsplit(" · ", 1)
        qtd = linha["Qtd por conjunto"]
        qtd = 1.0 if pd.isna(qtd) or qtd <= 0 else float(qtd)
        pint = None
        if s["fi_pint"] and not df_pin.empty:
            achou = df_pin[df_pin["Código Item"].astype(str) == cod]
            pint = achou.iloc[-1] if not achou.empty else None
        try:
            if tipo == "ABS":
                comp = fi.componente_abs(par, df_abs[df_abs["Código Item"].astype(str) == cod].iloc[-1], pint)
            else:
                comp = fi.componente_insertos(par, df_ins[df_ins["Código Item"].astype(str) == cod].iloc[-1], det_ins, pint)
            comp["qtd"], comp["lote"] = qtd, s["fi_lote"]
            comps.append(comp)
        except Exception as e:
            erros.append(f"{cod}: {e}")

    mont = ({"cc": s["fi_cc"], "grupo": fi.GRUPO_MONTAGEM, "setup": 0, "tempo": s["fi_tempo"], "homens": s["fi_homens"]}
            if s["fi_mont"] else None)
    cab = {"numero": s["fi_num"].strip(), "cliente": s["fi_cli"].strip(), "codigo": s["fi_cod"].strip(),
           "descricao": s["fi_desc"].strip(), "consumo": s["fi_cons"], "lote": s["fi_lote"], "similar": s["fi_sim"].strip()}

    with col_r:
        st.markdown("**Prévia do conjunto**")
        for e in erros:
            st.warning(e)
        if comps:
            try:
                linhas_prev = []
                for c_ in comps:
                    x = fi.custos_peca(par, c_)
                    linhas_prev.append({"Componente": f'{c_["codigo"]} ({c_["tipo"]})', "Qtd": c_["qtd"],
                                        "Custo unit. (R$)": round(x["Total"], 2), "Total (R$)": round(x["Total"] * c_["qtd"], 2)})
                tot = fi.totais_conjunto(par, comps, mont, s["fi_lote"])
                st.metric("Custo do conjunto", brl(tot["Total"]))
                p1, p2, p3 = st.columns(3)
                p1.metric("Mão de obra", brl(tot["Mão de obra"]))
                p2.metric("Set-up", brl(tot["Set-up"]))
                p3.metric("Matéria-prima + externo", brl(tot["Matéria-prima"] + tot["Externo"]))
                st.dataframe(pd.DataFrame(linhas_prev), hide_index=True)
                st.caption("Mesmas somas da folha MONT (o set-up não multiplica pela quantidade, como no modelo). "
                           "Matéria-prima do cliente fica sem preço, igual à ficha.")
            except Exception as e:
                st.error(f"Erro no cálculo: {e}")
            else:
                if st.button("📄 Gerar ficha", type="primary", key="fi_gerar"):
                    try:
                        arq, avisos = fi.gerar(par, cab, comps, mont, s["fi_folha_mont"])
                        nome = re.sub(r"[^\w.-]", "_", cab["codigo"] or comps[0]["codigo"])
                        s["fi_arquivo"] = (arq, f"FICHA_{nome}.xlsx", avisos)
                    except Exception as e:
                        st.error(f"Erro ao gerar a ficha: {e}")
                if "fi_arquivo" in s:
                    arq, nome_arq, avisos = s["fi_arquivo"]
                    for av in avisos:
                        st.warning(av)
                    st.download_button("⬇️ Baixar ficha (.xlsx)", data=arq, file_name=nome_arq,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.caption("A ficha baixada é a gerada no último clique em Gerar ficha.")
        elif not erros:
            st.caption("Adicione ao menos um componente para ver o custo e gerar a ficha.")

# =============================== HISTÓRICO ===================================
with tab_hist:
    ABAS = {"ABS": ("Orçamentos", "Custo Total (R$)"), "Insertos": ("Insertos", "Custo da Peça (R$)"),
            "Pintura": ("Pintura", "Custo Total (R$)"), "Montagem": ("Montagem", "Custo Montagem (R$)")}
    escolha = st.radio("Ver", list(ABAS) + ["Resumo por código"], horizontal=True, key="hist_aba")

    if escolha == "Resumo por código":
        partes = []
        for nome, (aba, col) in ABAS.items():
            df = d.carregar(aba)
            if not df.empty:
                partes.append(df[["Código Item", col]].rename(columns={col: f"{nome} (R$)"}))
        if partes:
            resumo_cod = pd.concat(partes).groupby("Código Item", as_index=False).sum(min_count=1)
            resumo_cod["Total (R$)"] = resumo_cod.drop(columns="Código Item").sum(axis=1)
            st.dataframe(resumo_cod, hide_index=True)
        else:
            st.info("Nenhum orçamento salvo ainda.")
    else:
        aba, col_custo = ABAS[escolha]
        df = d.carregar(aba)
        if df.empty:
            st.info("Nenhum orçamento salvo nesta aba ainda.")
        else:
            h1, h2, h3 = st.columns(3)
            h1.metric("Orçamentos salvos", len(df))
            h2.metric("Custo médio", brl(df[col_custo].mean()))
            h3.metric("Maior custo", brl(df[col_custo].max()))
            st.dataframe(df, hide_index=True)

            codigos = df["Código Item"].astype(str).tolist()
            if escolha == "Insertos":
                with st.expander("Ver composição de um orçamento"):
                    cod_v = st.selectbox("Código do item", codigos, key="hist_ver")
                    det_df = d.carregar("Insertos - Detalhe")
                    if not det_df.empty:
                        st.dataframe(det_df[det_df["Código Item"].astype(str) == cod_v], hide_index=True)

            with st.expander("Excluir um orçamento"):
                cod_x = st.selectbox("Código do item", codigos, key="hist_del")
                apagar = st.button("Excluir", key="hist_excluir")
                if apagar:
                    try:
                        d.excluir(aba, cod_x)
                    except PermissionError:
                        st.error("Feche o orcamentos.xlsx no Excel e tente de novo.")
                        apagar = False
                if apagar:
                    st.rerun()

    if d.ARQ_ORC.exists():
        st.download_button("⬇️ Baixar planilha completa (.xlsx)", data=d.ARQ_ORC.read_bytes(),
                           file_name="orcamentos.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# =============================== PARÂMETROS ==================================
with tab_par:
    st.caption("Preços, taxas e constantes usados em todos os cálculos. Ficam em parametros.xlsx.")
    sub = st.tabs(["Centros de custo", "Materiais (insertos)", "Tintas", "Operações", "Constantes", "Chapas ABS", "Densidades", "Grupos máquina"])
    novas = {}
    with sub[0]:
        novas["Centros de custo"] = st.data_editor(par.centros, num_rows="dynamic", hide_index=True, key="par_cc")
    with sub[1]:
        novas["Materiais"] = st.data_editor(par.materiais, num_rows="dynamic", hide_index=True, key="par_mat")
    with sub[2]:
        novas["Tintas"] = st.data_editor(par.tintas, num_rows="dynamic", hide_index=True, key="par_tinta")
    with sub[3]:
        st.caption("Cada operação já define o centro de custo (C.C.) usado no cálculo.")
        novas["Operações"] = st.data_editor(par.operacoes, num_rows="dynamic", hide_index=True, key="par_ops")
    with sub[4]:
        const = pd.DataFrame({"Parâmetro": list(par.k), "Valor": list(par.k.values())})
        obs = dict(d._padrao()["Constantes"][["Parâmetro", "Obs."]].values)
        const["Obs."] = const["Parâmetro"].map(obs)
        novas["Constantes"] = st.data_editor(
            const, hide_index=True, key="par_const", disabled=["Parâmetro", "Obs."], num_rows="fixed",
            column_config={"Valor": st.column_config.NumberColumn("Valor", format="%g")})

    with sub[5]:
        st.caption("Polímero em branco = a chapa serve para qualquer polímero (ex.: ABS/PC).")
        novas["Chapas"] = st.data_editor(
            par.chapas, num_rows="dynamic", hide_index=True, key="par_chapas",
            column_config={"Polímero": st.column_config.SelectboxColumn("Polímero", options=list(calc.POLIMEROS))})

    with sub[6]:
        st.caption("Densidade em g/cm³ (usada na aba Aproveitamento).")
        novas["Densidades"] = st.data_editor(par.densidades, num_rows="dynamic", hide_index=True, key="par_dens")

    with sub[7]:
        st.caption("Descrição de cada grupo de máquina (usada nas fichas).")
        novas["Grupos máquina"] = st.data_editor(par.grupos_df, num_rows="dynamic", hide_index=True, key="par_grupos")

    if st.button("💾 Salvar parâmetros", type="primary", key="par_salvar"):
        try:
            d.salvar_parametros(novas)
            st.session_state["_msg"] = "Parâmetros salvos"
            salvo = True
        except PermissionError:
            st.error("Não consegui salvar: feche o parametros.xlsx no Excel e tente de novo.")
            salvo = False
        if salvo:
            st.rerun()
