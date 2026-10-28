# =============================================================================
# ORÇAMENTO ABS - Calculadora de Termoformagem + Reforços + Montagem + Pintura
# Rodar com:  streamlit run app.py
# =============================================================================

import math
from pathlib import Path

import pandas as pd
import streamlit as st

from calculos_engenharia import (
    DENSIDADE, peso_tubo_redondo, peso_tubo_quadrado, peso_tubo_retangular,
    peso_barra_redonda, peso_barra_quadrada, peso_barra_sextavada,
    peso_chapa_circular, peso_chapa_plastica, peso_plastico_bolha,
    desenvolvimento_chapa, comprimento_dobra_raio_medio,
    diametro_para_perimetro, perimetro_para_diametro,
    tonelagem_prensa, volume_recipiente,
    aproveitamento_chapa, tempo_corte_laser,
    pintura_tunel, pintura_pistola_quantidade, pintura_pistola_tempo,
    tinta_agco, tinta_fundo_agco, tinta_po, custo_mo_plastico,
)

ARQUIVO = Path(__file__).parent / "orcamentos.xlsx"
PLANILHA_REF = Path(__file__).parent / "Planilha ABS-PINTURA-REFORÇO.xlsx"

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
# REGRAS DE NEGÓCIO - TERMOFORMAGEM (inalterado)
# -----------------------------------------------------------------------------

def tipo_conceito_params(conceito: str):
    if conceito == "Mesa":
        return 200, 0.3
    return 100, 0.0


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
        raise ValueError("Não é possível encaixar nenhuma peça na chapa com essas medidas.")
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

    item = {
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
    custos = {
        "Termoformagem": round(custo_termo, 2),
        "Preparação": round(custo_prep, 2),
        "Destacar": round(custo_destacar, 2),
        "Recortar": round(custo_recortar, 2),
        "Acabamento": round(custo_acabamento, 2),
    }
    return item, custos


# -----------------------------------------------------------------------------
# PERSISTÊNCIA (inalterado)
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
# NOVO: BASE DE DADOS DE REFERÊNCIA (lida da Planilha ABS-PINTURA-REFORÇO.xlsx)
# -----------------------------------------------------------------------------

# Fallback caso a planilha não exista (valores extraídos da planilha)
CENTROS_CUSTO_FALLBACK = {
    312: ("MANUTENÇÃO", 80),
    322: ("MODELARIA", 80),
    323: ("PINTURA DE PEÇAS - TUNEL", 1400),
    325: ("FERRAMENTARIA", 80),
    338: ("PREPARAÇÃO PINTURA CABINE", 188),
    339: ("ACABAMENTO TERMO FORMAGEM", 105),
    343: ("TERMO FORMAGEM", 230),
    345: ("PINTURA PÓ", 170),
    348: ("CORTE LASER 3D", 340),
    358: ("P&D", 80),
    359: ("CORTE E DOBRA", 140),
    360: ("PRENSAS", 240),
    366: ("MONTAGEM FINAL", 130),
    367: ("SOLDA E MONTAGEM", 170),
    368: ("TERMO FORMAGEM - MARCOPOLO", 252),
    369: ("ACABAMENTO TERMO FORMAGEM - MARCOPOLO", 105),
    370: ("CORTE LASER 2D", 180),
}

# Grupos de máquina relevantes para os processos (descrição = nome que aparece)
GRUPOS_MAQUINA_FALLBACK = [
    "GUILHOTINA MECANICA NEWTON",
    "CORTE LASER 2D",
    "DOBRADEIRA CNC 160 TN - YAWEI",
    "PRENSA EXC. JUTMANN 130 T",
    "MÁQ. TERMO FORMAGEM (MAQUINA 1)",
    "DESTACAR PEÇA 1",
    "RECORTAR PEÇA FRESA 1",
    "LIXADEIRA 1",
    "CABINE DE PINTURA DE PECAS",
    "MONTAGEM FINAL - DIVERSOS",
]

# Preços/kg da matéria-prima (extraídos da aba ORÇAMENTO INSERTOS)
MATERIAS_PRIMA = [
    ("109299000", "CHAPA AÇO 1,90x1000x3000 NBR7008 ZEE G1 Z100", 7.31),
    ("1064003000", "CHAPA ALUMÍNIO 1,50x1000x2000 LIGA 1200 H14", 28.75),
    ("107679000", "CHAPA AÇO 1,55x1000x2680 NBR7008 ZEE G2 Z100", 6.76),
    ("107676000", "CHAPA AÇO 1,20x1500x1795 ST05 Z/NBR7008 ZEE G2", 8.54),
    ("106546000", "CHAPA AÇO 1,55mm 1200x3000 NBR7008 ZEE GR1 Z100", 4.2525),
    ("100755000", "CHAPA AÇO (diversos)", 6.87),
    ("100106000", "CHAPA AÇO 4,75mm FQD 1200x3000 SAE1006 DEC", 5.8373),
    ("109796000", "CHAPA AÇO (diversos)", 5.0438),
    ("106039000", "CHAPA AÇO (diversos)", 5.849),
    ("103085000", "CHAPA AÇO (diversos)", 4.22),
    ("109389000", "CHAPA AÇO 1,00x1500x2700 GI EN10346 DX56D Z140", 6.87),
    ("103209000", "CHAPA AÇO (diversos)", 5.1084),
    ("106475000", "MANTA DE ISOLAMENTO", 39.283),
    ("100634000", "LÍQUIDO (diversos)", 23.53),
    ("107670000", "CHAPA AÇO 6,35x1200x2000 NBR6656 LNE-38 DECOL", 5.91),
]

# Fatores de tinta por tipo de acabamento (aba ORÇAMENTO PINTURA)
FATOR_TINTA = {
    "PRETO FOSCO": 1.89,
    "PRETO BRILHO": 1.5,
    "CINZA TITÂNIO": 1.89,
    "CINZA CRISTAL": 1.89,
    "FUNDO CINZA": 1.89,
    "AZUL MP2": 1.89,
    "EVO1": 1.89,
    "EVO2": 1.89,
    "EVO4 FOSCO": 1.89,
    "ECO1": 1.89,
    "ECO2": 1.89,
    "ECO3": 1.89,
    "PRIMER": 1.89,
    "PRIMER PRETO": 1.89,
    "RETARDADOR": 1.5,
    "PRATA": 1.89,
}

# Preços por litro de tinta (aba ORÇAMENTO PINTURA coluna M — valores unitários)
PRECO_TINTA = {
    "PRETO FOSCO": 79.72,
    "PRETO BRILHO": 33.078,
    "CINZA TITÂNIO": 59.95,
    "CINZA CRISTAL": 79.72,
    "FUNDO CINZA": 79.72,
    "AZUL MP2": 79.72,
    "EVO1": 79.72,
    "EVO2": 48.98,
    "EVO4 FOSCO": 52.03,
    "ECO1": 79.72,
    "ECO2": 48.57,
    "ECO3": 71.28,
    "PRIMER": 44.56,
    "PRIMER PRETO": 23.53,
    "RETARDADOR": 26.46,
    "PRATA": 103.4,
}

# Constantes de pintura (aba ORÇAMENTO PINTURA)
TINTA_POR_MM2 = 3.26e-07           # coluna G
CATALISADOR_DIVISOR = 8            # coluna I
TEMPO_PINTURA_POR_MM2 = 1.8181e-05 # coluna J
CUSTO_HORA_PINTURA = 175.0         # usado na fórmula M
PRECO_TINTA_REF = 57.01             # default quando cor não está no dicionário


def carregar_centros_custo():
    """Lê a aba CENTRO DE CUSTO da planilha de referência; fallback embutido."""
    if PLANILHA_REF.exists():
        try:
            df = pd.read_excel(PLANILHA_REF, sheet_name="CENTRO DE CUSTO", header=1)
            df = df.dropna(subset=["C.C."])
            centros = {}
            for _, row in df.iterrows():
                try:
                    cc = int(row["C.C."])
                    centros[cc] = (str(row["DESCRIÇÃO"]), float(row["CUSTO"]))
                except (ValueError, TypeError):
                    continue
            if centros:
                return centros
        except Exception:
            pass
    return dict(CENTROS_CUSTO_FALLBACK)


def carregar_grupos_maquina():
    """Lê a aba GRUPO MÁQUINA 2025; retorna lista de descrições únicas."""
    if PLANILHA_REF.exists():
        try:
            df = pd.read_excel(PLANILHA_REF, sheet_name="GRUPO MÁQUINA 2025")
            nomes = df["Descrição"].dropna().astype(str).unique().tolist()
            if nomes:
                return sorted(nomes)
        except Exception:
            pass
    return list(GRUPOS_MAQUINA_FALLBACK)


# -----------------------------------------------------------------------------
# NOVAS REGRAS DE NEGÓCIO - REFORÇOS (INSERTOS)
# -----------------------------------------------------------------------------

def calcular_custo_hora(custo_hora: float, tempo_min: float, num_homens: int = 1) -> float:
    """Fórmula da planilha: custo_hora * tempo / 60 * nº_homens."""
    return (custo_hora * tempo_min / 60) * num_homens


def calcular_reforco(
    cod_item: str,
    materiais: list,       # [(cod, desc, preco_kg, qtd, un), ...]
    terceiros: list,       # [(cod, desc, qtd, preco_unit), ...]
    operacoes: list,       # [(num_op, descricao, centro_custo, tempo_min, num_homens), ...]
) -> dict:
    """
    Retorna dict com linhas detalhadas e totais:
      - total_materia_prima
      - total_terceiros
      - total_mao_obra_setup  (soma R$ setup)
      - total_mao_obra_operacao (soma R$ operação)
      - custo_total (matéria-prima + terceiros + mão-de-obra)
    """
    linhas_mat = []
    total_mp = 0.0
    for cod, desc, preco, qtd, un in materiais:
        subtotal = float(qtd) * float(preco)
        total_mp += subtotal
        linhas_mat.append({
            "Código": cod, "Descrição": desc, "Qtd": qtd, "Un": un,
            "Preço Unit. (R$)": round(preco, 4),
            "Total (R$)": round(subtotal, 2),
        })

    linhas_ter = []
    total_ter = 0.0
    for cod, desc, qtd, preco_unit in terceiros:
        subtotal = float(qtd) * float(preco_unit)
        total_ter += subtotal
        linhas_ter.append({
            "Código": cod, "Descrição": desc, "Qtd": qtd,
            "Preço Unit. (R$)": round(preco_unit, 2),
            "Total (R$)": round(subtotal, 2),
        })

    centros = carregar_centros_custo()
    linhas_op = []
    total_setup = 0.0
    total_oper = 0.0
    for num_op, descricao, cc_codigo, setup_min, tempo_min, num_homens in operacoes:
        desc_cc, valor_hora = centros.get(int(cc_codigo), ("—", 0.0))
        rs_setup = calcular_custo_hora(valor_hora, float(setup_min), int(num_homens))
        rs_oper = calcular_custo_hora(valor_hora, float(tempo_min), int(num_homens))
        total_setup += rs_setup
        total_oper += rs_oper
        linhas_op.append({
            "Nº Op.": num_op,
            "Descrição": descricao,
            "Centro Custo": f"{int(cc_codigo)} - {desc_cc}",
            "Custo Hora (R$)": round(valor_hora, 2),
            "Setup (min)": setup_min,
            "Tempo Op. (min)": tempo_min,
            "Nº Homens": num_homens,
            "R$ Setup": round(rs_setup, 2),
            "R$ Operação": round(rs_oper, 2),
        })

    custo_total = total_mp + total_ter + total_setup + total_oper

    return {
        "cod_item": cod_item,
        "linhas_materiais": linhas_mat,
        "linhas_terceiros": linhas_ter,
        "linhas_operacoes": linhas_op,
        "total_materia_prima": round(total_mp, 2),
        "total_terceiros": round(total_ter, 2),
        "total_setup": round(total_setup, 2),
        "total_operacao": round(total_oper, 2),
        "custo_total": round(custo_total, 2),
    }


# -----------------------------------------------------------------------------
# NOVA REGRA - MONTAGEM (custo-hora do setor)
# -----------------------------------------------------------------------------

def calcular_montagem(cc_codigo: int, tempo_min: float, num_homens: int = 1) -> dict:
    centros = carregar_centros_custo()
    desc, valor_hora = centros.get(int(cc_codigo), ("—", 0.0))
    custo = calcular_custo_hora(valor_hora, float(tempo_min), int(num_homens))
    return {
        "centro_custo": f"{int(cc_codigo)} - {desc}",
        "custo_hora": round(valor_hora, 2),
        "tempo_min": tempo_min,
        "num_homens": num_homens,
        "custo_montagem": round(custo, 2),
    }


# -----------------------------------------------------------------------------
# NOVA REGRA - PINTURA
# -----------------------------------------------------------------------------

def calcular_pintura(comprimento_mm, largura_mm, cor, num_lados=2) -> dict:
    area_mm2 = float(comprimento_mm) * float(largura_mm)
    fator = FATOR_TINTA.get(cor, 1.89)
    tinta_ml = area_mm2 * TINTA_POR_MM2 * fator
    tinta_l = tinta_ml / 1000
    catalisador_l = tinta_l / CATALISADOR_DIVISOR
    tempo_pintura_min = (area_mm2 * TEMPO_PINTURA_POR_MM2) / 2  # já dividido por 2 como na planilha
    preco_tinta = PRECO_TINTA.get(cor, PRECO_TINTA_REF)
    custo_tinta = tinta_l * preco_tinta
    custo_mao_obra = (tempo_pintura_min / 60) * CUSTO_HORA_PINTURA
    custo_um_lado = custo_tinta + custo_mao_obra
    custo_dois_lados = custo_um_lado * 2
    return {
        "area_mm2": round(area_mm2, 2),
        "fator_tinta": fator,
        "tinta_ml": round(tinta_ml, 2),
        "tinta_l": round(tinta_l, 4),
        "catalisador_l": round(catalisador_l, 4),
        "tempo_pintura_min": round(tempo_pintura_min, 2),
        "preco_tinta_l": preco_tinta,
        "custo_tinta": round(custo_tinta, 2),
        "custo_mao_obra": round(custo_mao_obra, 2),
        "custo_um_lado": round(custo_um_lado, 2),
        "custo_dois_lados": round(custo_dois_lados, 2) if num_lados == 2 else None,
        "num_lados": num_lados,
    }


# -----------------------------------------------------------------------------
# INTERFACE
# -----------------------------------------------------------------------------

def brl(valor):
    return "R$ " + f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def excluir_item(cod_item):
    df = carregar_df()
    df = df[df["Código Item"].astype(str) != str(cod_item)]
    df.to_excel(ARQUIVO, index=False)
    formatar_planilha(ARQUIVO, df)


st.set_page_config(page_title="Orçamento ABS", page_icon="🧮", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 2.5rem; max-width: 1200px;}
[data-testid="stMetric"] {
    background: rgba(128,128,128,0.08);
    border: 1px solid rgba(128,128,128,0.25);
    border-radius: 10px;
    padding: 12px 16px;
}
[data-testid="stMetricLabel"] {opacity: 0.75;}
</style>
""", unsafe_allow_html=True)

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


def ler_parametros():
    s = st.session_state
    return dict(
        cod_item=s["cod"].strip(),
        comprimento=s["comprimento"],
        largura=s["largura"],
        espessura=s["espessura"],
        cor=s["cor"],
        peso_liquido=s["peso_liquido"],
        polimero_label=s["polimero"],
        tipo_corte=s["corte"],
        conceito=s["conceito"],
    )


st.markdown("## 🧮 Orçamento ABS")
st.caption("Calculadora de peças termoformadas + reforços + montagem + pintura")

tab_novo, tab_hist, tab_reforco, tab_montagem, tab_pintura, \
    tab_eng, tab_aprov, tab_laser = st.tabs([
    "Novo orçamento", "Histórico", "Reforços (Insertos)",
    "Montagem", "Pintura", "Engenharia",
    "Aproveitamento", "Corte Laser"
])
# ------------------------------- NOVO ORÇAMENTO ------------------------------
with tab_novo:
    col_form, col_prev = st.columns([1.1, 1], gap="large")

    with col_form:
        with st.container(border=True):
            st.markdown("**Identificação**")
            st.text_input("Código do item", key="cod", placeholder="Código único do item")

        with st.container(border=True):
            st.markdown("**Dimensões**")
            c1, c2, c3 = st.columns(3)
            c1.number_input("Comprimento (mm)", min_value=1, max_value=2700, step=1, key="comprimento")
            c2.number_input("Largura (mm)", min_value=1, max_value=1700, step=1, key="largura")
            c3.number_input("Espessura (mm)", min_value=3.0, max_value=7.0, step=0.1, format="%.1f", key="espessura")

        with st.container(border=True):
            st.markdown("**Material**")
            c4, c5, c6 = st.columns(3)
            c4.selectbox("Cor", CORES, key="cor")
            c5.selectbox("Polímero", list(POLIMEROS.keys()), key="polimero")
            c6.number_input("Peso líquido (kg)", min_value=0.001, max_value=1000.0, step=0.1, format="%.3f", key="peso_liquido")

        with st.container(border=True):
            st.markdown("**Processo**")
            c7, c8 = st.columns(2)
            c7.radio("Tipo de corte", ["Corte Robô", "Corte Manual"], key="corte", horizontal=True)
            c8.radio("Conceito", ["Mesa", "Caixa"], key="conceito", horizontal=True)

        b1, b2, _ = st.columns([1.3, 1, 1.5])
        salvar = b1.button("💾 Salvar orçamento", type="primary")
        b2.button("Limpar", on_click=resetar_formulario)

        if salvar:
            params = ler_parametros()
            if not params["cod_item"]:
                st.warning("Informe o código do item.")
            elif codigo_existe(carregar_df(), params["cod_item"]):
                st.warning(f"Código **{params['cod_item']}** já cadastrado.")
            else:
                try:
                    item, _ = calcular_orcamento(**params)
                    salvar_item(item)
                    st.toast(f"Orçamento {params['cod_item']} salvo!", icon="✅")
                    st.success(f"Orçamento **{params['cod_item']}** salvo: {brl(item['Custo Total (R$)'])}")
                except PermissionError:
                    st.error("Feche o orcamentos.xlsx no Excel e tente de novo.")
                except Exception as e:
                    st.error(f"Erro ao salvar: {e}")

    with col_prev:
        st.markdown("**Prévia do cálculo**")
        try:
            prev, custos = calcular_orcamento(**ler_parametros())
        except Exception as e:
            prev = None
            st.warning(str(e))

        if prev:
            m1, m2 = st.columns(2)
            m1.metric("Custo total", brl(prev["Custo Total (R$)"]))
            m2.metric("Peças por chapa", prev["Qtd Peças/Chapa"])
            m3, m4 = st.columns(2)
            m3.metric("Peso bruto", f"{prev['Peso Bruto (kg)']:.3f} kg".replace(".", ","))
            m4.metric("Blank (mm)", prev["Blank (mm)"])

            st.markdown("**Custo por etapa (R$)**")
            df_custos = pd.DataFrame({"Etapa": list(custos.keys()), "Custo (R$)": list(custos.values())}).set_index("Etapa")
            st.bar_chart(df_custos)

            with st.expander("Ver todos os campos"):
                detalhe = pd.DataFrame([(k, str(v)) for k, v in prev.items()], columns=["Campo", "Valor"])
                st.table(detalhe.set_index("Campo"))

# --------------------------------- HISTÓRICO ---------------------------------
with tab_hist:
    df_salvo = carregar_df()
    if df_salvo.empty:
        st.info("Nenhum orçamento salvo ainda.")
    else:
        h1, h2, h3 = st.columns(3)
        h1.metric("Orçamentos salvos", len(df_salvo))
        h2.metric("Custo médio", brl(df_salvo["Custo Total (R$)"].mean()))
        h3.metric("Maior custo", brl(df_salvo["Custo Total (R$)"].max()))

        st.dataframe(df_salvo, hide_index=True)

        st.download_button(
            "⬇️ Baixar planilha (.xlsx)",
            data=ARQUIVO.read_bytes(),
            file_name="orcamentos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Excluir um orçamento"):
            cod_del = st.selectbox("Código do item", df_salvo["Código Item"].astype(str).tolist())
            if st.button("Excluir"):
                try:
                    excluir_item(cod_del)
                    st.rerun()
                except PermissionError:
                    st.error("Feche o orcamentos.xlsx no Excel e tente de novo.")

# ------------------------------- REFORÇOS (INSERTOS) -------------------------

ARQUIVO_MATERIAIS = Path(__file__).parent / "materiais_reforco.xlsx"

# Base inicial (extraída da planilha ORÇAMENTO INSERTOS) — usada só na 1ª execução
MATERIAIS_INICIAIS = [
    # (código, descrição, preço R$/un, unidade)
    ("109299000",  "CHAPA DE ACO 1,90 x 1000 x 3000mm - NBR7008 ZEE GRAU 1 Z100", 7.31,   "kg"),
    ("1064003000", "CHAPA DE ALUMINIO 1,50 X 1000 X 2000 mm - LIGA 1200 - H14",  28.75,  "kg"),
    ("107679000",  "CHAPA DE ACO 1,55X1000X2680 - NBR 7008 ZEE - GRAU 2 - Z100",  6.76,   "kg"),
    ("107676000",  "CHAPA DE ACO 1,20X1500X1795 - ST 05 Z/NBR 7008 ZEE GRAU 2,+Z", 8.54, "kg"),
    ("106546000",  "CHAPA DE ACO 1,55mm 1200 x 3000mm NBR 7008 ZEE GR1 (Z100)",   4.2525, "kg"),
    ("100755000",  "CHAPA DE ACO (código 100755000)",                             6.87,   "kg"),
    ("100106000",  "CHAPA DE ACO 4,75mm FQD 1200 x 3000 - SAE 1006 DEC",          5.8373, "kg"),
    ("109796000",  "CHAPA DE ACO (código 109796000)",                             5.0438, "kg"),
    ("106039000",  "CHAPA DE ACO (código 106039000)",                             5.849,  "kg"),
    ("103085000",  "CHAPA DE ACO (código 103085000)",                             4.22,   "kg"),
    ("109389000",  "CHAPA DE ACO 1,00 x 1500 x 2700mm GI EN10346 DX56D Z140-M-C-", 6.87,  "kg"),
    ("103209000",  "CHAPA DE ACO (código 103209000)",                             5.1084, "kg"),
    ("106475000",  "MANTA DE ISOLAMENTO",                                         39.283, "m²"),
    ("100634000",  "LÍQUIDO (código 100634000)",                                  23.53,  "L"),
    ("100106000",  "CHAPA DE ACO 4,75mm FQD 1200 x 3000 - SAE 1006 DEC",          5.8373, "kg"),
    ("107670000",  "CHAPA DE AÇO 6,35 x 1200 x 2000mm - NBR 6656 LNE-38 DECOL",   5.91,   "kg"),
]


def _inicializar_banco_materiais():
    """Cria o arquivo de materiais com os dados iniciais se ainda não existir."""
    if not ARQUIVO_MATERIAIS.exists():
        df = pd.DataFrame(MATERIAIS_INICIAIS,
                          columns=["Código", "Descrição", "Preço Unit. (R$)", "Un"])
        # remove códigos duplicados mantendo o primeiro
        df = df.drop_duplicates(subset=["Código"], keep="first").reset_index(drop=True)
        df.to_excel(ARQUIVO_MATERIAIS, index=False)


def carregar_materiais() -> pd.DataFrame:
    """Retorna DataFrame com Código, Descrição, Preço Unit. (R$), Un."""
    _inicializar_banco_materiais()
    try:
        df = pd.read_excel(ARQUIVO_MATERIAIS)
        df["Código"] = df["Código"].astype(str)
        return df
    except Exception:
        return pd.DataFrame(MATERIAIS_INICIAIS,
                            columns=["Código", "Descrição", "Preço Unit. (R$)", "Un"])


def adicionar_material(codigo: str, descricao: str, preco: float, unidade: str) -> bool:
    """Adiciona material ao banco. Retorna False se o código já existir."""
    df = carregar_materiais()
    if str(codigo) in df["Código"].values:
        return False
    novo = pd.DataFrame([{
        "Código": str(codigo),
        "Descrição": descricao,
        "Preço Unit. (R$)": float(preco),
        "Un": unidade,
    }])
    df = pd.concat([df, novo], ignore_index=True)
    df.to_excel(ARQUIVO_MATERIAIS, index=False)
    return True


with tab_reforco:
    st.markdown("### 🔩 Reforços / Insertos")
    st.caption("Selecione o material no catálogo — o preço é preenchido automaticamente.")

    centros = carregar_centros_custo()
    grupos = carregar_grupos_maquina()
    df_materiais = carregar_materiais()

    # Mapeamento código -> (descrição, preço, unidade) para lookup rápido
    mapa_materiais = {
        str(row["Código"]): (row["Descrição"], float(row["Preço Unit. (R$)"]), row["Un"])
        for _, row in df_materiais.iterrows()
    }

    # -------- Painel para cadastrar novo material (expansível) --------
    with st.expander("➕ Cadastrar novo material no banco"):
        with st.form("form_novo_material", clear_on_submit=True):
            c1, c2, c3, c4 = st.columns([1.2, 2.5, 1, 1])
            novo_cod = c1.text_input("Código")
            nova_desc = c2.text_input("Descrição")
            novo_preco = c3.number_input("Preço (R$)", min_value=0.0, value=0.0, format="%.4f")
            nova_un = c4.selectbox("Unidade", ["kg", "m²", "L", "un", "m"])
            salvar_mat = st.form_submit_button("Salvar material", type="primary")

        if salvar_mat:
            if not novo_cod.strip() or not nova_desc.strip():
                st.warning("Preencha código e descrição.")
            elif adicionar_material(novo_cod.strip(), nova_desc.strip(), novo_preco, nova_un):
                st.success(f"Material **{novo_cod}** adicionado ao banco.")
                st.rerun()
            else:
                st.warning(f"Código **{novo_cod}** já existe no banco.")

    st.divider()

    # -------- Lista de códigos com "—" (vazio) no topo --------
    opcoes_material = ["— selecione —"] + [
        f"{cod} | {desc} | R$ {preco:.4f}/{un}"
        for cod, (desc, preco, un) in mapa_materiais.items()
    ]
    codigos_ordenados = ["— selecione —"] + sorted(mapa_materiais.keys())

    # -------- 1) Matéria-prima --------
    st.markdown("**1) Matéria-prima**")
    n_mp = st.number_input("Quantas linhas de matéria-prima?",
                           min_value=1, max_value=20, value=1, step=1, key="n_mp")

    materiais_input = []
    for i in range(int(n_mp)):
        with st.container(border=True):
            st.markdown(f"*Linha {i+1}*")
            c1, c2, c3 = st.columns([3, 1, 1])
            sel_idx = c1.selectbox(
                "Material",
                options=range(len(opcoes_material)),
                format_func=lambda j: opcoes_material[j],
                key=f"mp_sel_{i}",
            )
            qtd = c2.number_input("Qtd", min_value=0.0, value=0.0,
                                  step=0.01, key=f"mp_qtd_{i}")
            # Mostra preço/unidade do material selecionado
            if sel_idx > 0:
                cod_sel = codigos_ordenados[sel_idx]
                desc, preco, un = mapa_materiais[cod_sel]
                c3.metric(f"R$/{un}", f"{preco:.4f}")
                if qtd > 0:
                    st.caption(f"Subtotal: **{brl(qtd * preco)}**")
                materiais_input.append((cod_sel, desc, preco, qtd, un))
            else:
                c3.metric("R$/un", "—")
                if qtd > 0:
                    st.caption("⚠ Selecione um material para calcular.")

    # -------- 2) Operações externas (terceiros) --------
    st.markdown("**2) Operações externas (terceiros)**")
    n_ter = st.number_input("Quantas linhas de terceiros?",
                            min_value=0, max_value=10, value=0, step=1, key="n_ter")
    terceiros_input = []
    for i in range(int(n_ter)):
        with st.container(border=True):
            c1, c2, c3, c4 = st.columns([1.5, 2.5, 1, 1])
            cod = c1.text_input("Código", key=f"ter_cod_{i}")
            desc = c2.text_input("Descrição", key=f"ter_desc_{i}")
            qtd = c3.number_input("Qtd", min_value=0.0, value=1.0, key=f"ter_qtd_{i}")
            preco = c4.number_input("Preço unit. (R$)", min_value=0.0,
                                    value=0.0, key=f"ter_preco_{i}")
            terceiros_input.append((cod, desc, qtd, preco))

    # -------- 3) Processo de fabricação (mão-de-obra) --------
    st.markdown("**3) Processo de fabricação (mão-de-obra)**")
    n_op = st.number_input("Quantas operações?", min_value=1, max_value=20,
                           value=1, step=1, key="n_op")

    opcoes_cc = {f"{k} - {v[0]} (R$ {v[1]:.2f}/h)": (k, v[1])
                 for k, v in sorted(centros.items())}
    operacoes_input = []
    for i in range(int(n_op)):
        with st.container(border=True):
            st.markdown(f"*Operação {i+1}*")
            c1, c2, c3 = st.columns([1, 3, 2])
            num_op = c1.number_input("Nº op.", min_value=1, value=10 * (i + 1),
                                     step=10, key=f"op_num_{i}")
            descricao = c2.text_input("Descrição da operação", key=f"op_desc_{i}")
            cc_label = c3.selectbox("Centro de custo",
                                    list(opcoes_cc.keys()), key=f"op_cc_{i}")
            c4, c5, c6, c7 = st.columns(4)
            grupo_maq = c4.selectbox("Grupo máquina", grupos, key=f"op_grupo_{i}")
            setup = c5.number_input("Setup (min)", min_value=0.0, value=0.0,
                                    key=f"op_setup_{i}")
            tempo = c6.number_input("Tempo op. (min)", min_value=0.0, value=0.0,
                                    key=f"op_tempo_{i}")
            homens = c7.number_input("Nº homens", min_value=1, value=1,
                                     key=f"op_homens_{i}")
            operacoes_input.append((
                int(num_op), descricao, opcoes_cc[cc_label][0],
                setup, tempo, int(homens),
            ))

    st.divider()

    if st.button("🔍 Calcular reforço", type="primary", use_container_width=True):
        res = calcular_reforco(
            cod_item="—",
            materiais=materiais_input,
            terceiros=terceiros_input,
            operacoes=operacoes_input,
        )

        st.markdown("#### Resultado")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Matéria-prima", brl(res["total_materia_prima"]))
        k2.metric("Terceiros", brl(res["total_terceiros"]))
        k3.metric("Setup", brl(res["total_setup"]))
        k4.metric("Operação", brl(res["total_operacao"]))
        k5.metric("**Custo total**", brl(res["custo_total"]))

        if res["linhas_materiais"]:
            st.markdown("**Matéria-prima**")
            df_mp = pd.DataFrame(res["linhas_materiais"])
            df_mp["Total (R$)"] = df_mp["Total (R$)"].map(brl)
            df_mp["Preço Unit. (R$)"] = df_mp["Preço Unit. (R$)"].map(
                lambda v: f"{v:.4f}"
            )
            st.dataframe(df_mp, hide_index=True, use_container_width=True)

        if res["linhas_terceiros"]:
            st.markdown("**Terceiros**")
            st.dataframe(pd.DataFrame(res["linhas_terceiros"]),
                         hide_index=True, use_container_width=True)

        if res["linhas_operacoes"]:
            st.markdown("**Operações**")
            df_op = pd.DataFrame(res["linhas_operacoes"])
            for col in ["Custo Hora (R$)", "R$ Setup", "R$ Operação"]:
                df_op[col] = df_op[col].map(brl)
            st.dataframe(df_op, hide_index=True, use_container_width=True)

        st.info(f"💡 Custo total do reforço: **{brl(res['custo_total'])}**")

# ---------------------------------- MONTAGEM ---------------------------------
with tab_montagem:
    st.markdown("### 🛠 Montagem")
    st.caption("Custo de montagem pelo custo-hora do setor")

    centros = carregar_centros_custo()
    opcoes_cc = {f"{k} - {v[0]}  (R$ {v[1]:.2f}/h)": k for k, v in sorted(centros.items())}

    with st.form("form_montagem"):
        c1, c2, c3 = st.columns(3)
        cc_label = c1.selectbox("Centro de custo / setor", list(opcoes_cc.keys()))
        tempo = c2.number_input("Tempo de montagem (min)", min_value=0.0, value=10.0, step=1.0)
        homens = c3.number_input("Nº de homens", min_value=1, max_value=20, value=1, step=1)
        calc = st.form_submit_button("Calcular montagem", type="primary")

    if calc:
        res = calcular_montagem(opcoes_cc[cc_label], tempo, int(homens))
        m1, m2, m3 = st.columns(3)
        m1.metric("Custo-hora", brl(res["custo_hora"]))
        m2.metric("Tempo × homens", f"{tempo:.0f} min × {homens}")
        m3.metric("**Custo de montagem**", brl(res["custo_montagem"]))
        with st.expander("Detalhes"):
            st.json(res)

# ----------------------------------- PINTURA ---------------------------------
with tab_pintura:
    st.markdown("### 🎨 Pintura")
    st.caption("Custo de tinta + mão-de-obra por lado")

    cores_pintura = sorted(PRECO_TINTA.keys())

    with st.form("form_pintura"):
        c1, c2, c3, c4 = st.columns(4)
        comp = c1.number_input("Comprimento (mm)", min_value=1, value=500, step=1)
        larg = c2.number_input("Largura (mm)", min_value=1, value=400, step=1)
        cor_sel = c3.selectbox("Cor / acabamento", cores_pintura)
        lados = c4.radio("Lados", [1, 2], horizontal=True, index=1)
        calc_pint = st.form_submit_button("Calcular pintura", type="primary")

    if calc_pint:
        res = calcular_pintura(comp, larg, cor_sel, int(lados))
        p1, p2, p3, p4 = st.columns(4)
        p1.metric("Área", f"{res['area_mm2']:,.0f} mm²".replace(",", "."))
        p2.metric("Tinta", f"{res['tinta_ml']:.2f} mL")
        p3.metric("Catalisador", f"{res['catalisador_l']:.4f} L")
        p4.metric("Tempo pintura", f"{res['tempo_pintura_min']:.2f} min")

        q1, q2, q3, q4 = st.columns(4)
        q1.metric("Preço tinta (R$/L)", brl(res["preco_tinta_l"]))
        q2.metric("Custo tinta", brl(res["custo_tinta"]))
        q3.metric("Custo mão-de-obra", brl(res["custo_mao_obra"]))
        q4.metric(f"**Custo {lados} lado(s)**",
                  brl(res["custo_dois_lados"] if lados == 2 else res["custo_um_lado"]))

        with st.expander("Detalhes do cálculo"):
            st.json(res)

#----------------------------------- ENGENHARIA ---------------------------------

with tab_eng:
    st.markdown("### 📐 Cálculos de Engenharia")
    st.caption("Pesos, desenvolvimento de chapa, conversões, prensa, volume.")

    sub_peso, sub_chapa, sub_conv, sub_misc = st.tabs([
        "Pesos de matéria-prima", "Desenvolvimento de chapa",
        "Conversões", "Prensa / Volume"
    ])

    with sub_peso:
        tipo = st.selectbox("Tipo de material", [
            "Tubo redondo", "Tubo quadrado", "Tubo retangular",
            "Barra redonda", "Barra quadrada", "Barra sextavada",
            "Chapa circular", "Chapa plástica", "Plástico bolha",
        ], key="eng_tipo_peso")

        if tipo == "Tubo redondo":
            c1, c2, c3 = st.columns(3)
            d = c1.number_input("Diâmetro externo (mm)", value=30.0, key="eng_tr_d")
            e = c2.number_input("Espessura (mm)", value=2.0, key="eng_tr_e")
            L = c3.number_input("Comprimento (mm)", value=1395.0, key="eng_tr_L")
            st.metric("Peso", f"{peso_tubo_redondo(d, e, L):.3f} kg")

        elif tipo == "Tubo quadrado":
            c1, c2, c3 = st.columns(3)
            lado = c1.number_input("Lado (mm)", value=30.0, key="eng_tq_lado")
            e = c2.number_input("Espessura (mm)", value=3.0, key="eng_tq_e")
            L = c3.number_input("Comprimento (mm)", value=6000.0, key="eng_tq_L")
            st.metric("Peso", f"{peso_tubo_quadrado(lado, e, L):.3f} kg")

        elif tipo == "Tubo retangular":
            c1, c2, c3, c4 = st.columns(4)
            b = c1.number_input("Base (mm)", value=30.0, key="eng_tret_b")
            h = c2.number_input("Altura (mm)", value=60.0, key="eng_tret_h")
            e = c3.number_input("Espessura (mm)", value=3.0, key="eng_tret_e")
            L = c4.number_input("Comprimento (mm)", value=6000.0, key="eng_tret_L")
            st.metric("Peso", f"{peso_tubo_retangular(b, h, e, L):.3f} kg")

        elif tipo == "Barra redonda":
            c1, c2 = st.columns(2)
            d = c1.number_input("Diâmetro (mm)", value=131.0, key="eng_br_d")
            L = c2.number_input("Comprimento (mm)", value=24.0, key="eng_br_L")
            st.metric("Peso", f"{peso_barra_redonda(d, L):.3f} kg")

        elif tipo == "Barra quadrada":
            c1, c2 = st.columns(2)
            b = c1.number_input("Base (mm)", value=27.5, key="eng_bq_b")
            L = c2.number_input("Comprimento (mm)", value=48.0, key="eng_bq_L")
            st.metric("Peso", f"{peso_barra_quadrada(b, L):.3f} kg")

        elif tipo == "Barra sextavada":
            c1, c2 = st.columns(2)
            d = c1.number_input("Diâmetro (mm)", value=12.7, key="eng_bs_d")
            L = c2.number_input("Comprimento (mm)", value=1000.0, key="eng_bs_L")
            st.metric("Peso", f"{peso_barra_sextavada(d, L):.3f} kg")

        elif tipo == "Chapa circular":
            c1, c2 = st.columns(2)
            d = c1.number_input("Diâmetro (mm)", value=134.0, key="eng_cc_d")
            e = c2.number_input("Espessura (mm)", value=4.0, key="eng_cc_e")
            st.metric("Peso", f"{peso_chapa_circular(d, e):.3f} kg")

        elif tipo == "Chapa plástica":
            c1, c2, c3, c4 = st.columns(4)
            t = c1.selectbox("Tipo", ["ABS", "PEAD", "PP"], key="eng_cp_tipo")
            e = c2.number_input("Espessura (mm)", value=4.0, key="eng_cp_e")
            w = c3.number_input("Largura (mm)", value=50.0, key="eng_cp_w")
            L = c4.number_input("Comprimento (mm)", value=515.0, key="eng_cp_L")
            st.metric("Peso", f"{peso_chapa_plastica(e, w, L, t):.3f} kg")

        elif tipo == "Plástico bolha":
            c1, c2 = st.columns(2)
            L = c1.number_input("Comprimento (mm)", value=2853.0, key="eng_pb_L")
            w = c2.number_input("Largura (mm)", value=1500.0, key="eng_pb_w")
            st.metric("Peso", f"{peso_plastico_bolha(L, w):.4f} kg")

    with sub_chapa:
        st.markdown("**Desenvolvimento de chapas dobradas**")
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        ang = c1.number_input("Ângulo (°)", value=90.0, key="eng_dev_ang")
        e = c2.number_input("Espessura (mm)", value=3.0, key="eng_dev_e")
        a = c3.number_input("a (mm)", value=20.0, key="eng_dev_a")
        b = c4.number_input("b (mm)", value=32.0, key="eng_dev_b")
        coef = c5.number_input("Coef.", value=3.0, key="eng_dev_coef")
        r = c6.number_input("Raio (mm)", value=6.0, key="eng_dev_r")
        st.metric("Comprimento necessário", f"{desenvolvimento_chapa(ang, e, a, b, coef, r):.2f} mm")

        st.markdown("**Raio médio da dobra**")
        c1, c2 = st.columns(2)
        rmed = c1.number_input("Raio médio (mm)", value=147.0, key="eng_rmed")
        ang2 = c2.number_input("Ângulo (°)", value=90.0, key="eng_ang2")
        st.metric("Comprimento da dobra", f"{comprimento_dobra_raio_medio(rmed, ang2):.2f} mm")

    with sub_conv:
        st.markdown("**Ø → Perímetro**")
        c1, c2 = st.columns(2)
        d = c1.number_input("Diâmetro externo (mm)", value=900.0, key="eng_conv_d")
        e = c2.number_input("Espessura (mm)", value=25.4, key="eng_conv_e")
        st.json(diametro_para_perimetro(d, e))

        st.markdown("**Perímetro → Ø**")
        c1, c2 = st.columns(2)
        p = c1.number_input("Perímetro externo (mm)", value=2827.43, key="eng_conv_p")
        e2 = c2.number_input("Espessura (mm)", value=25.4, key="eng_conv_e2")
        st.json(perimetro_para_diametro(p, e2))

    with sub_misc:
        st.markdown("**Definir prensa (tonelagem)**")
        c1, c2, c3 = st.columns(3)
        w = c1.number_input("Largura peça (mm)", value=200.0, key="eng_pr_w")
        L = c2.number_input("Comprimento peça (mm)", value=200.0, key="eng_pr_L")
        e = c3.number_input("Espessura (mm)", value=2.0, key="eng_pr_e")
        st.metric("Prensa necessária", f"{tonelagem_prensa(w, L, e):.1f} Toneladas")

        st.markdown("**Volume de recipiente**")
        c1, c2, c3 = st.columns(3)
        comp = c1.number_input("Comprimento (mm)", value=500.0, key="eng_vol_c")
        larg = c2.number_input("Largura (mm)", value=200.0, key="eng_vol_l")
        alt = c3.number_input("Altura (mm)", value=100.0, key="eng_vol_a")
        st.metric("Volume", f"{volume_recipiente(comp, larg, alt):.2f} L")

#------------------------------- APROVEITAMENTO -------------------------------------------
with tab_aprov:
    st.markdown("### ✂️ Aproveitamento de chapa + Guilhotina")
    with st.form("form_aprov"):
        c1, c2, c3 = st.columns(3)
        e_chapa = c1.number_input("Espessura chapa (mm)", value=1.5, key="ap_e_chapa")
        l_chapa = c2.number_input("Largura chapa (mm)", value=1000.0, key="ap_l_chapa")
        c_chapa = c3.number_input("Comprimento chapa (mm)", value=2810.0, key="ap_c_chapa")
        c4, c5, c6 = st.columns(3)
        e_peca = c4.number_input("Espessura peça (mm)", value=1.5, key="ap_e_peca")
        l_peca = c5.number_input("Largura peça (mm)", value=453.0, key="ap_l_peca")
        c_peca = c6.number_input("Comprimento peça (mm)", value=496.0, key="ap_c_peca")
        qtd_blank = st.number_input("Qtd peças por blank", min_value=1, value=1, key="ap_qtd_blank")
        calc = st.form_submit_button("Calcular", type="primary")

    if calc:
        r = aproveitamento_chapa(e_chapa, l_chapa, c_chapa,
                                 e_peca, l_peca, c_peca, int(qtd_blank))
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Melhor modelo", r["escolhido"])
        m2.metric("Peças por chapa", r["total_pecas"])
        m3.metric("Peso da chapa", f"{r['peso_chapa_kg']} kg")
        m4.metric("Retalho", f"{r['retalho_kg']} kg")
        st.json(r)


