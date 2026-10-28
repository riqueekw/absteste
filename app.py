# =============================================================================
# ORÇAMENTO ABS - Calculadora de Termoformagem (versão Streamlit)
# Rodar com:  streamlit run app.py
# =============================================================================

import math
from pathlib import Path

import pandas as pd
import streamlit as st

# O arquivo fica na mesma pasta do app.py
ARQUIVO = Path(__file__).parent / "orcamentos.xlsx"

CORES = sorted([
    "EVO1", "EVO2", "EVO4",
    "ECO1", "ECO2", "ECO3",
    "PRETO", "BLACK PIANO", "BRANCO GEADA",
])

POLIMEROS = {
    "ABS-s": ("ABS", 1.06),
    "ABS-PMMA": ("ABS-PMMA", 1.18),
    "ABS-TPE": ("ABS-TPE", 1.11),
}

CHAPA_COMPRIMENTO_M = 2.7
CHAPA_LARGURA_M = 1.7


# -----------------------------------------------------------------------------
# REGRAS DE NEGÓCIO (idênticas ao código do Colab)
# -----------------------------------------------------------------------------

def tipo_conceito_params(conceito: str):
    """Retorna (cto_mm, adicional) conforme o conceito de fabricação."""
    if conceito == "Mesa":
        return 200, 0.3
    return 100, 0.0  # Caixa


def calcular_peso_bruto(comprimento, largura, cto, espessura, densidade, adc):
    return ((comprimento + cto) * (largura + cto) * espessura * densidade * (1 + adc)) / 1_000_000


def calcular_perimetro(comprimento, largura, cto):
    comp = (comprimento + cto) / 1000
    lng = (largura + cto) / 1000
    perimetro = (2 * comp) + (2 * lng)
    return comp, lng, perimetro


def quantidade_pecas(comp, lng):
    qtd1 = math.floor(CHAPA_COMPRIMENTO_M / comp) * math.floor(CHAPA_LARGURA_M / lng)
    qtd2 = math.floor(CHAPA_COMPRIMENTO_M / lng) * math.floor(CHAPA_LARGURA_M / comp)
    return max(qtd1, qtd2)


def calcular_termoformagem(comp, lng):
    qtd = quantidade_pecas(comp, lng)
    if qtd <= 0:
        raise ValueError(
            "Não é possível encaixar nenhuma peça na chapa com essas medidas."
        )
    termo = 5.75 / qtd
    custo_termo = (termo * 252) / 60
    return termo, custo_termo, qtd


def calcular_preparacao(termo):
    prep = termo * 0.1
    return prep, (prep * 105) / 60


def calcular_destacar(perimetro):
    destacar = perimetro * 0.25
    return destacar, (destacar * 105) / 60


def calcular_recortar(perimetro):
    recortar = perimetro * 0.5
    return recortar, (recortar * 105) / 60


def calcular_acabamento(perimetro):
    acabamento = perimetro * 0.9
    return acabamento, (acabamento * 105) / 60


def calcular_orcamento(cod_item, comprimento, largura, espessura, cor,
                       peso_liquido, polimero_label, tipo_corte, conceito):
    """Executa toda a cadeia de cálculo e devolve um dicionário pronto para salvar."""
    polimero, densidade = POLIMEROS[polimero_label]
    cto, adc = tipo_conceito_params(conceito)

    comprimento_blank = comprimento + cto
    largura_blank = largura + cto

    peso_bruto = calcular_peso_bruto(comprimento, largura, cto, espessura, densidade, adc)
    comp, lng, perimetro = calcular_perimetro(comprimento, largura, cto)
    termo, custo_termo, qtd_pecas = calcular_termoformagem(comp, lng)
    prep, custo_prep = calcular_preparacao(termo)
    destacar, custo_destacar = calcular_destacar(perimetro)
    recortar, custo_recortar = calcular_recortar(perimetro)
    acabamento, custo_acabamento = calcular_acabamento(perimetro)

    custo_total = (custo_termo + custo_prep + custo_destacar
                   + custo_recortar + custo_acabamento)

    return {
        "Código Item": cod_item,
        "Polímero": polimero,
        "Cor": cor,
        "Tipo Corte": tipo_corte,
        "Conceito": conceito,
        "Blank (mm)": f"{comprimento_blank} x {largura_blank}",
        "Peso Líquido (kg)": round(peso_liquido, 3),
        "Peso Bruto (kg)": round(peso_bruto, 3),
        "Qtd Peças/Chapa": qtd_pecas,
        "Tempo Termo (min)": round(termo, 2),
        "Tempo Prep (min)": round(prep, 2),
        "Tempo Destacar (min)": round(destacar, 2),
        "Tempo Recortar (min)": round(recortar, 2),
        "Tempo Acabamento (min)": round(acabamento, 2),
        "Custo Recorte (R$)": round(custo_recortar, 2),
        "Custo Total (R$)": round(custo_total, 2),
    }


# -----------------------------------------------------------------------------
# PERSISTÊNCIA (Excel)
# -----------------------------------------------------------------------------

def carregar_df():
    if ARQUIVO.exists():
        return pd.read_excel(ARQUIVO)
    return pd.DataFrame()


def codigo_existe(df, cod_item):
    if df.empty or "Código Item" not in df.columns:
        return False
    return cod_item in df["Código Item"].astype(str).values


def salvar_item(novo_item):
    df = carregar_df()
    df = pd.concat([df, pd.DataFrame([novo_item])], ignore_index=True)
    df.to_excel(ARQUIVO, index=False)
    formatar_planilha(ARQUIVO, df)
    return df


COLUNAS_MOEDA = {"Custo Recorte (R$)", "Custo Total (R$)"}
COLUNAS_DECIMAL = {
    "Peso Líquido (kg)", "Peso Bruto (kg)",
    "Tempo Termo (min)", "Tempo Prep (min)", "Tempo Destacar (min)",
    "Tempo Recortar (min)", "Tempo Acabamento (min)",
}


def formatar_planilha(caminho, df):
    """Aplica formatação profissional (cabeçalho, cores, larguras, moeda, filtro)."""
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = load_workbook(caminho)
    ws = wb.active
    ws.title = "Orçamentos"

    n_linhas = ws.max_row
    n_colunas = ws.max_column

    fonte_padrao = Font(name="Calibri", size=11)
    for linha in ws.iter_rows(min_row=1, max_row=n_linhas, max_col=n_colunas):
        for celula in linha:
            celula.font = fonte_padrao

    fundo_cabecalho = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for celula in ws[1]:
        celula.fill = fundo_cabecalho
        celula.font = fonte_cabecalho
        celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    borda = Border(*(Side(style="thin", color="D9D9D9"),) * 4)
    for linha in ws.iter_rows(min_row=1, max_row=n_linhas, max_col=n_colunas):
        for celula in linha:
            celula.border = borda

    fundo_zebra = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid")
    colunas = list(df.columns)
    for idx_linha in range(2, n_linhas + 1):
        if idx_linha % 2 == 0:
            for celula in ws[idx_linha]:
                celula.fill = fundo_zebra
        for idx_col, nome_col in enumerate(colunas, start=1):
            celula = ws.cell(row=idx_linha, column=idx_col)
            if nome_col in COLUNAS_MOEDA:
                celula.number_format = 'R$ #,##0.00'
            elif nome_col in COLUNAS_DECIMAL:
                celula.number_format = '0.00'
            if idx_col > 1:
                celula.alignment = Alignment(horizontal="center")

    for idx_col, nome_col in enumerate(colunas, start=1):
        maior = len(str(nome_col))
        for idx_linha in range(2, n_linhas + 1):
            valor = ws.cell(row=idx_linha, column=idx_col).value
            maior = max(maior, len(str(valor)) if valor is not None else 0)
        ws.column_dimensions[get_column_letter(idx_col)].width = min(maior + 4, 32)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(caminho)


# -----------------------------------------------------------------------------
# INTERFACE (Streamlit)
# -----------------------------------------------------------------------------

st.set_page_config(page_title="Orçamento ABS", page_icon="🧮", layout="centered")

DEFAULTS = {
    "cod": "",
    "comprimento": 1000,
    "largura": 1000,
    "espessura": 3.0,
    "cor": CORES[0],
    "peso_liquido": 1.0,
    "polimero": list(POLIMEROS.keys())[0],
    "corte": "Corte Robô",
    "conceito": "Mesa",
}

for chave, valor in DEFAULTS.items():
    st.session_state.setdefault(chave, valor)


def resetar_formulario():
    for chave, valor in DEFAULTS.items():
        st.session_state[chave] = valor
    st.session_state.pop("ultimo", None)


st.title("🧮 Orçamento de Peças ABS Termoformadas")

st.text_input("Código do item", key="cod")

col1, col2, col3 = st.columns(3)
col1.number_input("Comprimento (mm)", min_value=1, max_value=2700, step=1, key="comprimento")
col2.number_input("Largura (mm)", min_value=1, max_value=1700, step=1, key="largura")
col3.number_input("Espessura (mm)", min_value=3.0, max_value=7.0, step=0.1,
                  format="%.1f", key="espessura")

col4, col5, col6 = st.columns(3)
col4.selectbox("Cor", CORES, key="cor")
col5.selectbox("Polímero", list(POLIMEROS.keys()), key="polimero")
col6.number_input("Peso líquido (kg)", min_value=0.001, max_value=1000.0, step=0.1,
                  format="%.3f", key="peso_liquido")

col7, col8 = st.columns(2)
col7.radio("Tipo de corte", ["Corte Robô", "Corte Manual"], key="corte")
col8.radio("Conceito", ["Mesa", "Caixa"], key="conceito")

b1, b2, _ = st.columns([1, 1, 2])
calcular = b1.button("Calcular e Salvar", type="primary", use_container_width=True)
b2.button("Novo orçamento", on_click=resetar_formulario, use_container_width=True)

if calcular:
    cod_item = st.session_state["cod"].strip()

    if not cod_item:
        st.warning("Informe o código do item.")
    elif codigo_existe(carregar_df(), cod_item):
        st.warning(f"Código **{cod_item}** já cadastrado. Use outro código.")
    else:
        try:
            novo_item = calcular_orcamento(
                cod_item=cod_item,
                comprimento=st.session_state["comprimento"],
                largura=st.session_state["largura"],
                espessura=st.session_state["espessura"],
                cor=st.session_state["cor"],
                peso_liquido=st.session_state["peso_liquido"],
                polimero_label=st.session_state["polimero"],
                tipo_corte=st.session_state["corte"],
                conceito=st.session_state["conceito"],
            )
            salvar_item(novo_item)
            st.session_state["ultimo"] = novo_item
        except PermissionError:
            st.error("Não consegui salvar: feche o orcamentos.xlsx no Excel e tente de novo.")
        except Exception as e:
            st.error(f"Erro no cálculo: {e}")

if "ultimo" in st.session_state:
    st.success("Orçamento salvo com sucesso")
    resultado = pd.DataFrame(
        [(k, str(v)) for k, v in st.session_state["ultimo"].items()],
        columns=["Campo", "Valor"],
    )
    st.table(resultado.set_index("Campo"))

st.divider()

df_salvo = carregar_df()
if df_salvo.empty:
    st.caption("Nenhum orçamento salvo ainda.")
else:
    st.subheader(f"Orçamentos salvos ({len(df_salvo)})")
    st.dataframe(df_salvo, use_container_width=True, hide_index=True)
    st.download_button(
        "Baixar planilha (.xlsx)",
        data=ARQUIVO.read_bytes(),
        file_name="orcamentos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon=":material/download:",
    )
