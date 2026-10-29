# =============================================================================
# ORÇAMENTOS  |  rodar com:  streamlit run app.py
# app.py = telas | calculos.py = regras | dados.py = Excel (parâmetros e orçamentos)
# =============================================================================
import pandas as pd
import streamlit as st

import calculos as calc
import dados as d

st.set_page_config(page_title="ORÇAMENTOS", page_icon="▶️", layout="wide")

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
centros = [int(float(x)) for x in par.centros["Código"].dropna()]
nome_cc = {int(float(cod)): f"{int(float(cod))} · {desc} ({brl(float(v))}/h)"
           for cod, desc, v in zip(par.centros["Código"], par.centros["Descrição"], par.centros["R$/h"])
           if pd.notna(cod) and pd.notna(v)}

st.markdown("## ▶️ ORÇAMENTOS")
tab_abs, tab_ins, tab_pin, tab_mon, tab_peso, tab_hist, tab_par = st.tabs(
    ["ABS", "Insertos", "Pintura", "Montagem", "Peso", "Histórico", "Parâmetros"])

# =============================== ABS =========================================
with tab_abs:
    PAD = {"abs_cod": "", "abs_comp": 1000, "abs_larg": 1000, "abs_esp": 3.0, "abs_cor": calc.CORES[0],
           "abs_peso": 1.0, "abs_pol": list(calc.POLIMEROS)[0], "abs_corte": "Corte Robô", "abs_conc": "Mesa"}
    iniciar(PAD)
    s = st.session_state
    col_f, col_r = st.columns([1.1, 1], gap="large")

    def params_abs():
        return dict(cod_item=s["abs_cod"].strip(), comprimento=s["abs_comp"], largura=s["abs_larg"],
                    espessura=s["abs_esp"], cor=s["abs_cor"], peso_liquido=s["abs_peso"],
                    polimero_label=s["abs_pol"], tipo_corte=s["abs_corte"], conceito=s["abs_conc"])

    with col_f:
        with st.container(border=True):
            st.markdown("**Identificação**")
            st.text_input("Código do item", key="abs_cod", placeholder="Código único do item")
        with st.container(border=True):
            st.markdown("**Dimensões**")
            c1, c2, c3 = st.columns(3)
            c1.number_input("Comprimento (mm)", min_value=1, max_value=2700, step=1, key="abs_comp")
            c2.number_input("Largura (mm)", min_value=1, max_value=1700, step=1, key="abs_larg")
            c3.number_input("Espessura (mm)", min_value=3.0, max_value=7.0, step=0.1, format="%.1f", key="abs_esp")
        with st.container(border=True):
            st.markdown("**Material**")
            c4, c5, c6 = st.columns(3)
            c4.selectbox("Cor", calc.CORES, key="abs_cor")
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
                salvar("Orçamentos", p["cod_item"], {"Orçamentos": [item]},
                       f"Orçamento **{p['cod_item']}** salvo: {brl(item['Custo Total (R$)'])}")
            except Exception as e:
                st.error(f"Erro no cálculo: {e}")

    with col_r:
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
    sub = st.tabs(["Centros de custo", "Materiais (insertos)", "Tintas", "Operações", "Constantes"])
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
