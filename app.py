# =============================================================================
# ORÇAMENTO ABS — Calculadora de termoformagem + custo de montagem
# Executar: streamlit run app.py
# Dependências: streamlit, pandas, openpyxl
# Mantenha app.py e orcamentos.xlsx na mesma pasta.
# =============================================================================

import math
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ARQUIVO = Path(__file__).parent / "orcamentos.xlsx"
CORES = sorted(["EVO1", "EVO2", "EVO4", "ECO1", "ECO2", "ECO3", "PRETO", "BLACK PIANO", "BRANCO GEADA"])
POLIMEROS = {"ABS-s": ("ABS", 1.06), "ABS-PMMA": ("ABS-PMMA", 1.18), "ABS-TPE": ("ABS-TPE", 1.11)}
CHAPA_COMPRIMENTO_M = 2.7
CHAPA_LARGURA_M = 1.7

COLUNAS = [
    "Código Item", "Polímero", "Cor", "Tipo Corte", "Conceito", "Blank (mm)",
    "Peso Líquido (kg)", "Peso Bruto (kg)", "Qtd Peças/Chapa", "Tempo Termo (min)",
    "Tempo Prep (min)", "Tempo Destacar (min)", "Tempo Recortar (min)",
    "Tempo Acabamento (min)", "Custo Recorte (R$)", "Custo Total (R$)",
    "Custo Montagem (R$)", "Custo Geral (R$)"
]


def parametros_conceito(conceito):
    return (200, 0.3) if conceito == "Mesa" else (100, 0.0)


def calcular_peso_bruto(comprimento, largura, cto, espessura, densidade, adicional):
    return ((comprimento + cto) * (largura + cto) * espessura * densidade * (1 + adicional)) / 1_000_000


def calcular_perimetro(comprimento, largura, cto):
    comp, larg = (comprimento + cto) / 1000, (largura + cto) / 1000
    return comp, larg, 2 * comp + 2 * larg


def quantidade_pecas(comp, larg):
    q1 = math.floor(CHAPA_COMPRIMENTO_M / comp) * math.floor(CHAPA_LARGURA_M / larg)
    q2 = math.floor(CHAPA_COMPRIMENTO_M / larg) * math.floor(CHAPA_LARGURA_M / comp)
    return max(q1, q2)


def calcular_orcamento(cod_item, comprimento, largura, espessura, cor, peso_liquido, polimero_label, tipo_corte, conceito):
    polimero, densidade = POLIMEROS[polimero_label]
    cto, adicional = parametros_conceito(conceito)
    blank_comp, blank_larg = comprimento + cto, largura + cto
    peso_bruto = calcular_peso_bruto(comprimento, largura, cto, espessura, densidade, adicional)
    comp, larg, perimetro = calcular_perimetro(comprimento, largura, cto)
    qtd = quantidade_pecas(comp, larg)
    if qtd <= 0:
        raise ValueError("Não é possível encaixar nenhuma peça na chapa com essas medidas.")
    tempo_termo = 5.75 / qtd
    tempo_prep = tempo_termo * 0.1
    tempo_destacar = perimetro * 0.25
    tempo_recortar = perimetro * 0.5
    tempo_acabamento = perimetro * 0.9
    custos = {
        "Termoformagem": (tempo_termo * 252) / 60,
        "Preparação": (tempo_prep * 105) / 60,
        "Destacar": (tempo_destacar * 105) / 60,
        "Recortar": (tempo_recortar * 105) / 60,
        "Acabamento": (tempo_acabamento * 105) / 60,
    }
    total = sum(custos.values())
    item = {
        "Código Item": str(cod_item).strip(), "Polímero": polimero, "Cor": cor,
        "Tipo Corte": tipo_corte, "Conceito": conceito,
        "Blank (mm)": f"{blank_comp} x {blank_larg}",
        "Peso Líquido (kg)": round(float(peso_liquido), 3), "Peso Bruto (kg)": round(peso_bruto, 3),
        "Qtd Peças/Chapa": qtd, "Tempo Termo (min)": round(tempo_termo, 2),
        "Tempo Prep (min)": round(tempo_prep, 2), "Tempo Destacar (min)": round(tempo_destacar, 2),
        "Tempo Recortar (min)": round(tempo_recortar, 2), "Tempo Acabamento (min)": round(tempo_acabamento, 2),
        "Custo Recorte (R$)": round(custos["Recortar"], 2), "Custo Total (R$)": round(total, 2),
        "Custo Montagem (R$)": 0.0, "Custo Geral (R$)": round(total, 2),
    }
    return item, {k: round(v, 2) for k, v in custos.items()}


def brl(valor):
    return "R$ " + f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def garantir_arquivo():
    if not ARQUIVO.exists():
        with pd.ExcelWriter(ARQUIVO, engine="openpyxl") as writer:
            pd.DataFrame(columns=COLUNAS).to_excel(writer, sheet_name="Orçamentos", index=False)
            pd.DataFrame(columns=["Código Item", "Custo Montagem (R$)"]).to_excel(writer, sheet_name="Montagem", index=False)
    wb = load_workbook(ARQUIVO)
    if "Orçamentos" not in wb.sheetnames:
        ws = wb.create_sheet("Orçamentos", 0)
        ws.append(COLUNAS)
    ws = wb["Orçamentos"]
    headers = [c.value for c in ws[1]]
    for col in COLUNAS:
        if col not in headers:
            ws.cell(row=1, column=len(headers) + 1, value=col)
            headers.append(col)
    if "Montagem" not in wb.sheetnames:
        wsm = wb.create_sheet("Montagem")
        wsm.append(["Código Item", "Custo Montagem (R$)"])
    wb.save(ARQUIVO)
    formatar_planilha()


def carregar_df():
    garantir_arquivo()
    df = pd.read_excel(ARQUIVO, sheet_name="Orçamentos", dtype={"Código Item": str})
    for c in COLUNAS:
        if c not in df.columns:
            df[c] = 0 if "Custo" in c else ""
    df["Custo Montagem (R$)"] = pd.to_numeric(df["Custo Montagem (R$)"], errors="coerce").fillna(0)
    df["Custo Total (R$)"] = pd.to_numeric(df["Custo Total (R$)"], errors="coerce").fillna(0)
    df["Custo Geral (R$)"] = df["Custo Total (R$)"] + df["Custo Montagem (R$)"]
    return df[COLUNAS]


def salvar_df(df):
    with pd.ExcelWriter(ARQUIVO, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df[COLUNAS].to_excel(writer, sheet_name="Orçamentos", index=False)
    formatar_planilha()


def salvar_item(item):
    df = carregar_df()
    if item["Código Item"] in df["Código Item"].astype(str).values:
        raise ValueError("Esse código já está cadastrado. Use a aba Montagem para atualizar somente a montagem.")
    df = pd.concat([df, pd.DataFrame([item])], ignore_index=True)
    salvar_df(df)


def salvar_montagem(codigo, valor):
    df = carregar_df()
    mask = df["Código Item"].astype(str) == str(codigo)
    if not mask.any():
        raise ValueError("Código não encontrado no histórico. Salve primeiro o orçamento do item.")
    df.loc[mask, "Custo Montagem (R$)"] = round(float(valor), 2)
    df.loc[mask, "Custo Geral (R$)"] = df.loc[mask, "Custo Total (R$)"] + round(float(valor), 2)
    salvar_df(df)


def excluir_item(codigo):
    df = carregar_df()
    salvar_df(df[df["Código Item"].astype(str) != str(codigo)].reset_index(drop=True))


def formatar_planilha():
    wb = load_workbook(ARQUIVO)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 30
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = Font(name="Calibri", size=11)
                cell.border = Border(*(Side(style="thin", color="D9D9D9"),) * 4)
                if cell.row % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="F2F6FA")
        for col_idx in range(1, ws.max_column + 1):
            name = str(ws.cell(1, col_idx).value or "")
            vals = [len(str(ws.cell(r, col_idx).value or "")) for r in range(1, min(ws.max_row, 100) + 1)]
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(vals + [12]) + 3, 32)
            if "(R$)" in name:
                for r in range(2, ws.max_row + 1):
                    ws.cell(r, col_idx).number_format = 'R$ #,##0.00'
            elif "(kg)" in name or "(min)" in name:
                for r in range(2, ws.max_row + 1):
                    ws.cell(r, col_idx).number_format = '0.00'
    wb.save(ARQUIVO)


st.set_page_config(page_title="Orçamento ABS", page_icon="🧮", layout="wide")
st.markdown("""<style>
.block-container {padding-top: 1.5rem; max-width: 1400px;}
[data-testid="stMetric"] {background: rgba(128,128,128,.08); border: 1px solid rgba(128,128,128,.2); border-radius: 10px; padding: 12px 16px;}
</style>""", unsafe_allow_html=True)

st.title("🧮 Orçamento ABS")
st.caption("Preencha os dados para calcular o custo de termoformagem; informe a montagem em uma aba separada.")
garantir_arquivo()
tab_calc, tab_montagem, tab_hist = st.tabs(["Novo orçamento", "Custo de montagem", "Histórico"])

with tab_calc:
    left, right = st.columns([1.05, 1], gap="large")
    with left:
        with st.form("form_orcamento"):
            st.subheader("1. Identificação")
            cod = st.text_input("Código do item", placeholder="Ex.: 13437033")
            st.subheader("2. Dimensões (mm)")
            a, b, c = st.columns(3)
            comprimento = a.number_input("Comprimento", min_value=1, max_value=2700, value=1000, step=1)
            largura = b.number_input("Largura", min_value=1, max_value=1700, value=1000, step=1)
            espessura = c.number_input("Espessura", min_value=3.0, max_value=7.0, value=3.0, step=0.1, format="%.1f")
            st.subheader("3. Material")
            a, b, c = st.columns(3)
            cor = a.selectbox("Cor", CORES)
            polimero = b.selectbox("Polímero", list(POLIMEROS.keys()))
            peso_liquido = c.number_input("Peso líquido (kg)", min_value=0.001, max_value=1000.0, value=1.0, step=0.1, format="%.3f")
            st.subheader("4. Processo")
            a, b = st.columns(2)
            corte = a.selectbox("Tipo de corte", ["Corte Robô", "Corte Manual"])
            conceito = b.radio("Conceito", ["Mesa", "Caixa"], horizontal=True)
            submitted = st.form_submit_button("Calcular e salvar orçamento", type="primary", use_container_width=True)
    params = dict(cod_item=cod.strip() or "Prévia", comprimento=comprimento, largura=largura, espessura=espessura,
                  cor=cor, peso_liquido=peso_liquido, polimero_label=polimero, tipo_corte=corte, conceito=conceito)
    try:
        preview, costs = calcular_orcamento(**params)
        with right:
            st.subheader("Resultado estimado")
            m1, m2 = st.columns(2)
            m1.metric("Custo de termoformagem", brl(preview["Custo Total (R$)"]))
            m2.metric("Peças por chapa", preview["Qtd Peças/Chapa"])
            m1, m2 = st.columns(2)
            m1.metric("Peso bruto", f'{preview["Peso Bruto (kg)"]:.3f} kg'.replace(".", ","))
            m2.metric("Blank", preview["Blank (mm)"])
            st.markdown("**Composição do custo**")
            st.dataframe(pd.DataFrame([{"Etapa": k, "Custo": brl(v)} for k, v in costs.items()]), hide_index=True, use_container_width=True)
            with st.expander("Ver dados calculados"):
                st.dataframe(pd.DataFrame({"Campo": list(preview.keys()), "Valor": [str(v) for v in preview.values()]}), hide_index=True, use_container_width=True)
    except Exception as e:
        preview = None
        with right:
            st.warning(str(e))
    if submitted:
        if not cod.strip():
            st.error("Informe o código do item antes de salvar.")
        elif cod.strip() in carregar_df()["Código Item"].astype(str).values:
            st.error("Esse código já existe. Para alterar o custo de montagem, use a aba Custo de montagem.")
        else:
            try:
                item, _ = calcular_orcamento(**params)
                salvar_item(item)
                st.success(f'Orçamento salvo para {cod.strip()}. Custo de termoformagem: {brl(item["Custo Total (R$)"])}')
            except Exception as e:
                st.error(f"Não foi possível salvar: {e}")

with tab_montagem:
    st.subheader("Custo de montagem")
    st.caption("Informe apenas o custo de montagem. O sistema soma esse valor ao custo de termoformagem do item.")
    df = carregar_df()
    if df.empty:
        st.info("Cadastre primeiro um item na aba Novo orçamento.")
    else:
        codigos = df["Código Item"].astype(str).tolist()
        codigo = st.selectbox("Código do item", codigos, key="codigo_montagem")
        linha = df[df["Código Item"].astype(str) == codigo].iloc[0]
        a, b, c = st.columns(3)
        a.metric("Custo de termoformagem", brl(linha["Custo Total (R$)"]))
        b.metric("Montagem atual", brl(linha["Custo Montagem (R$)"]))
        c.metric("Custo geral atual", brl(linha["Custo Geral (R$)"]))
        with st.form("form_montagem"):
            valor = st.number_input("Valor de montagem (R$)", min_value=0.0, value=float(linha["Custo Montagem (R$)"]), step=1.0, format="%.2f")
            salvar_montagem_btn = st.form_submit_button("Salvar custo de montagem", type="primary")
        if salvar_montagem_btn:
            try:
                salvar_montagem(codigo, valor)
                st.success(f"Montagem atualizada. Custo geral: {brl(float(linha['Custo Total (R$)']) + valor)}")
                st.rerun()
            except Exception as e:
                st.error(str(e))

with tab_hist:
    df = carregar_df()
    if df.empty:
        st.info("Nenhum orçamento cadastrado ainda.")
    else:
        a, b, c, d = st.columns(4)
        a.metric("Itens cadastrados", len(df))
        b.metric("Média termoformagem", brl(df["Custo Total (R$)"].mean()))
        c.metric("Total de montagem", brl(df["Custo Montagem (R$)"].sum()))
        d.metric("Maior custo geral", brl(df["Custo Geral (R$)"].max()))
        busca = st.text_input("Buscar por código, cor, polímero ou conceito")
        filtrado = df
        if busca.strip():
            mask = filtrado.astype(str).apply(lambda col: col.str.contains(busca.strip(), case=False, na=False)).any(axis=1)
            filtrado = filtrado[mask]
        st.dataframe(filtrado, hide_index=True, use_container_width=True)
        st.download_button("Baixar histórico em Excel", data=ARQUIVO.read_bytes(), file_name="orcamentos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with st.expander("Excluir item do histórico"):
            apagar = st.selectbox("Selecione o código", df["Código Item"].astype(str).tolist(), key="excluir_codigo")
            if st.button("Excluir item", type="secondary"):
                excluir_item(apagar)
                st.success(f"Item {apagar} excluído.")
                st.rerun()
