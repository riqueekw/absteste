"""Armazenamento em Excel.

parametros.xlsx -> preços, taxas e constantes (editável no app, aba Parâmetros, ou direto no Excel)
orcamentos.xlsx -> orçamentos salvos, uma aba por tipo
"""
from pathlib import Path

import pandas as pd

import calculos as calc

PASTA = Path(__file__).parent
ARQ_ORC = PASTA / "orcamentos.xlsx"
ARQ_PAR = PASTA / "parametros.xlsx"

# ---- valores iniciais (vieram da planilha ABS-PINTURA-REFORÇO; depois edite em Parâmetros) ----
CENTROS = [
    ('312', 'MANUTENÇÃO', 80.0),
    ('322', 'MODELARIA', 80.0),
    ('323', 'PINTURA DE PEÇAS - TUNEL', 1400.0),
    ('325', 'FERRAMENTARIA', 80.0),
    ('338', 'PREPARAÇÃO PINTURA CABINE', 188.0),
    ('339', 'ACABAMENTO TERMO FORMAGEM', 105.0),
    ('343', 'TERMO FORMAGEM', 230.0),
    ('345', 'PINTURA PÓ', 170.0),
    ('348', 'CORTE LASER 3D', 340.0),
    ('358', 'P&D', 80.0),
    ('359', 'CORTE E DOBRA', 140.0),
    ('360', 'PRENSAS', 240.0),
    ('366', 'MONTAGEM FINAL', 130.0),
    ('367', 'SOLDA E MONTAGEM', 170.0),
    ('368', 'TERMO FORMAGEM - MARCOPOLO', 252.0),
    ('369', 'ACABAMENTO TERMO FORMAGEM - MARCOPOLO', 105.0),
    ('370', 'CORTE LASER 2D', 180.0),
]

MATERIAIS = [
    ('109299000', 'CHAPA DE ACO 1,90 x 1000 x 3000mm - NBR7008 ZEE GRAU 1 Z100', 'kg', 7.31),
    ('1064003000', 'CHAPA DE ALUMINIO 1,50 X 1000 X 2000 mm - LIGA 1200 - H14', 'kg', 28.75),
    ('107679000', 'CHAPA DE ACO 1,55X1000X2680 - NBR 7008 ZEE - GRAU 2 - Z100', 'kg', 6.76),
    ('107676000', 'CHAPA DE ACO 1,20X1500X1795 - ST 05 Z/NBR 7008 ZEE GRAU 2,+Z', 'kg', 8.54),
    ('106546000', 'CHAPA DE ACO 1,55mm 1200 x 3000mm NBR 7008 ZEE GR1 (Z100)', 'kg', 4.2525),
    ('100755000', 'CHAPA DE ACO 3,00 mm FQ 1200 x 3000 - SAE 1008 DEC', 'kg', 6.87),
    ('100106000', 'CHAPA DE ACO 1,50mm FF 1200 x 3000 NBR5915 EM', 'kg', 5.91),
    ('109796000', 'CHAPA DE ACO 1,90mm FF 1500 x 3000', 'kg', 5.0438),
    ('106039000', 'CHAPA DE ACO 4,25mm FQ 1200 x 3000 - NBR 6656 LNE - 38', 'kg', 5.849),
    ('103085000', 'CHAPA DE ACO 6,35mm FQ 1200 x 3000 - SAE 1010/20', 'kg', 4.22),
    ('109389000', 'CHAPA DE ACO 1,00 x 1500 x 2700mm GI EN10346 DX56D Z140-M-C-', 'kg', 6.87),
    ('103209000', 'CHAPA DE ACO 1,20mm FF 1200 x 3000 EM', 'kg', 5.1084),
    ('106475000', 'MANTA DE ISOLAMENTO', 'M²', 39.283),
    ('100634000', 'PRIMER PRETO IMERSAO BASE D. AGUA', 'L', 23.53),
    ('100106000', 'CHAPA DE ACO 4,75mm FQD 1200 x 3000 - SAE 1006 DEC', 'kg', 5.8373),
    ('107670000', 'CHAPA DE AÇO 6,35 x 1200 x 2000mm - NBR 6656 LNE-38 DECOL', 'kg', 5.91),
]

TINTAS = [
    ('PRETO FOSCO', '105479000', 79.72),
    ('PRETO BRILHO', '106580000 / 106586000', 33.078),
    ('CINZA TITÂNIO', '105475000', 59.95),
    ('EVO2', '105515000', 48.98),
    ('RETARDADOR', '106605000', 26.46),
    ('ECO2', '105861000', 48.57),
    ('ECO3', '105816000', 71.28),
    ('PRIMER', '106613000', 44.56),
    ('CINZA CRISTAL', '106613000', 79.72),
    ('EVO 1', '106186000', 79.72),
    ('ECO 1', '105976000', 79.72),
    ('EVO4 FOSCO', '105473000', 52.03),
    ('PRIMER PRETO', '100634000', 23.53),
    ('PRATA', '107723000', 103.4),
]

OPERACOES = [
    ('CORTAR CHAPA EM TIRA', 359, '301-002'),
    ('CORTAR CHAPA EM BLANK', 359, '301-002'),
    ('CORTAR EM LASER', 370, '348-004'),
    ('DOBRAR PEÇA', 359, '303-003'),
    ('PINTAR PEÇA (CABINE PÓ)', 345, '350-001'),
    ('ESTAMPAR', 360, '336-001'),
    ('SOLDAR PINOS', 367, '305-008'),
    ('MONTAR CONJUNTO', 369, '351-001'),
    ('SOLDAR ULTRASSOM', 369, '352-007'),
    ('PINTAR PEÇA (TÚNEL)', 323, '323M'),
]


CONSTANTES = [
    (calc.K_LOTE, 100, "Divide o setup entre as peças do lote"),
    (calc.K_ABS_CICLO, 5.75, "min por chapa"),
    (calc.K_ABS_CC_TERMO, 368, "Código do centro de custo (R$/h vem da tabela Centros de custo)"),
    (calc.K_ABS_CC_ACAB, 369, "Código do centro de custo"),
    (calc.K_ABS_PREP, 0.1, "Preparação = fração do tempo de termoformagem"),
    (calc.K_ABS_DESTACAR, 0.25, "min por metro de perímetro"),
    (calc.K_ABS_RECORTAR, 0.5, "min por metro de perímetro"),
    (calc.K_ABS_ACABAMENTO, 0.9, "min por metro de perímetro"),
    (calc.K_CHAPA_C, 2.7, "m"),
    (calc.K_CHAPA_L, 1.7, "m"),
    (calc.K_PINT_HORA, 175, "R$/h"),
    (calc.K_PINT_CONSUMO, 3.26e-07, "L de tinta por mm² (um lado)"),
    (calc.K_PINT_TEMPO, 1.8181e-05, "min por mm² somando os dois lados (um lado = metade)"),
    (calc.K_MONT_CC, 369, "Código do centro de custo da montagem"),
]

DENSIDADES = [  # (material, g/cm³)
    ("Aço Carbono", 7.85), ("Aço Inox", 7.85), ("Alumínio", 2.7), ("Borracha Nat", 0.92),
    ("Bronze", 8.83), ("Cobre", 8.93), ("Latão", 8.5), ("Níquel", 8.91), ("Titânio", 4.55),
]

CHAPAS = [  # (código, comprimento mm, largura mm, espessura mm, polímero, cor, descrição)
    ('105751000', 2700, 1700, 4, 'ABS-s', 'EVO2', 'ABS C/ TEXTURA EVO II'),
    ('105753000', 2700, 1700, 4, 'ABS-s', 'EVO4', 'ABS C/ TEXTURA EVO IV'),
    ('105310000', 2400, 1600, 4, 'ABS-s', 'EVO2', 'ABES C/ TEXTURA EVO II'),
    ('105624000', 1200, 1600, 4, 'ABS-s', 'EVO1', 'ABS C/ TEXTURA EVO I'),
    ('105397000', 1600, 1200, 5.75, 'ABS-s', 'PRETO', 'ABS LISO PRETO'),
    ('105395000', 1600, 1200, 4, 'ABS-s', 'PRETO', 'ABS LISO PRETO'),
    ('106271000', 1600, 1200, 4, 'ABS-s', 'ECO2', 'ABS TEXT ECO II'),
    ('106235000', 1500, 2100, 6.5, 'ABS-s', 'ECO3', 'ABS TEXT ECO III'),
    ('106290000', 1200, 1600, 4, 'ABS-s', 'ECO1', 'ABS TEXT ECO I'),
    ('105905000', 1200, 1400, 4.5, 'ABS-PMMA', 'BRANCO', 'CHAPA ACRÍLICA MONDGLASS BRANCO'),
    ('105909000', 800, 850, 4, 'ABS-s', 'PRETO', 'ABS LISO PRETO'),
    ('105308000', 1600, 1200, 4, 'ABS-s', 'EVO2', 'ABS C/ TEXTURA EVO II'),
    ('105307000', 1600, 1200, 4, 'ABS-s', 'EVO4', 'ABS C/ TEXTURA EVO1V'),
    ('105904000', 1600, 1200, 4.5, 'ABS-PMMA', 'BRANCO', 'CHAPA ABS ACRILICO BRANCO'),
    ('107439000', 2100, 1100, 6, 'ABS-s', 'EVO4', 'CHAPA ABS C/ TEXTURA EVO IV'),
    ('105556000', 2700, 1400, 5.75, 'ABS-s', 'EVO1', 'CHAPA ABS C/ TEXTURA EVO I'),
    ('106297000', 2200, 1050, 4, 'ABS-s', 'EVO2', 'CHAPA ABS C/ TEXTURA EVO II'),
    ('106292000', 2700, 1700, 4, 'ABS-s', 'PRETO', 'CHAPA ABS LISO PRETO'),
    ('105398000', 2270, 1600, 5.75, 'ABS-s', 'PRETO', 'CHAPA ABS LISO PRETO'),
    ('106419000', 3610, 1930, 6, 'ABS-s', 'PRETO', 'CHAPA ABS LISO PRETO'),
    ('105782000', 2700, 1400, 5.75, 'ABS-s', 'EVO1', 'CHAPA ABS TEXT. EVO I'),
    ('106278000', 2700, 1700, 4, 'ABS-s', 'ECO1', 'CHAPA ABS TEXT. ECO 1'),
    ('106279000', 2700, 1700, 5.75, 'ABS-s', 'ECO1', 'CHAPA ABS TEXT. ECO 1'),
    ('106291000', 2700, 1700, 6, 'ABS-s', 'ECO1', 'CHAPA ABS TEXT. ECO 1'),
    ('106280000', 2700, 1700, 4, 'ABS-s', 'ECO2', 'CHAPA ABS TEXT. ECO 2'),
    ('106281000', 2700, 1700, 5.75, 'ABS-s', 'ECO2', 'CHAPA ABS TEXT. ECO 2'),
    ('106299000', 1600, 1200, 6.5, 'ABS-s', 'ECO3', 'CHAPA ABS TEXT. ECO 3'),
    ('106294000', 1600, 1200, 4, 'ABS-s', 'ECO3', 'CHAPA ABS TEXT. ECO 3'),
    ('106283000', 2700, 1700, 5.75, 'ABS-s', 'ECO3', 'CHAPA ABS TEXT. ECO 3'),
    ('106282000', 2700, 1700, 4, 'ABS-s', 'ECO3', 'CHAPA ABS TEXT. ECO 3'),
    ('106285000', 2700, 1700, 5.75, 'ABS-s', 'EVO1', 'CHAPA ABS TEXT. EVO I'),
    ('106284000', 2700, 1700, 4, 'ABS-s', 'EVO1', 'CHAPA ABS TEXT. EVO I'),
    ('106302000', 2850, 1950, 5.75, 'ABS-s', 'EVO1', 'CHAPA ABS TEXT. EVO 1'),
    ('106287000', 2700, 1700, 5.75, 'ABS-s', 'EVO2', 'CHAPA ABS TEXT. EVO II'),
    ('106289000', 2700, 1700, 5.75, 'ABS-s', 'EVO4', 'CHAPA ABS TEXT. EVO IV'),
    ('106483000', 1500, 1300, 6, 'ABS-PMMA', 'BRANCO', 'CHAPA ACRILICA MONDGLASS BRANCO'),
    ('106288000', 2700, 1700, 4, 'ABS-s', 'EVO4', 'CHAPA ABS TEXT. EVO IV'),
    ('105942000', 2050, 1460, 6, 'ABS-PMMA', 'BRANCO', 'CHAPA ACRILICA MONDGLASS BRANCO'),
    ('105941000', 2700, 1700, 6, 'ABS-PMMA', 'BRANCO', 'CHAPA ACRILICA MONDGLASS BRANCO'),
    ('105986000', 850, 600, 4.5, 'ABS-PMMA', 'BRANCO', 'CHAPA ACRILICA MONDGLASS BRANCO'),
    ('106451000', 1320, 1060, 6, 'ABS-PMMA', 'BRANCO', 'CHAPA ACRILICA MONDGLASS BRANCO'),
    ('106452000', 1700, 1700, 6, 'ABS-PMMA', 'BRANCO', 'CHAPA ACRILICA MONDGLASS BRANCO'),
    ('106453000', 2000, 1500, 6, 'ABS-PMMA', 'PRETO GLITTER', 'CHAPA ACRILICA MONDGLASS PRETO GLITTER'),
    ('106484000', 1400, 1000, 6, 'ABS-PMMA', 'BRANCO', 'CHAPA ACRILICA MONDGLASS BRANCO'),
    ('106485000', 1500, 1220, 6, 'ABS-PMMA', 'BRANCO', 'CHAPA ACRILICA MONDGLASS BRANCO'),
    ('106486000', 1500, 1300, 6, 'ABS-PMMA', 'BRANCO', 'CHAPA ACRILICA MONDGLASS BRANCO'),
    ('106487000', 1300, 950, 6, 'ABS-PMMA', 'BRANCO', 'CHAPA ACRILICA MONDGLASS BRANCO'),
    ('190300000', 1600, 2700, 5, 'ABS-TPE', 'ECO1', 'CHAPA ABS-TPE ECO1 SENOPLAST'),
    ('190301000', 1600, 2700, 4, 'ABS-TPE', 'ECO2', 'CHAPA ABS-TPE ECO2 SENOPLAST'),
    ('190302000', 1600, 2700, 4, 'ABS-TPE', 'RHINO', 'CHAPA ABS-TPE RHINO SENOPLAST'),
    ('190303000', 1600, 2700, 6, 'ABS-PMMA', 'BRANCO GEADA', 'CHAPA ABS/PC BRANCO GEADA SENOPLAST'),
    ('190304000', 1600, 2700, 4, 'ABS-PMMA', 'BLACK PIANO', 'CHAPA ABS/PC BLACK PIANO SENOPLAST'),
    ('190305000', 1600, 2700, 4, 'ABS-TPE', 'ECO3', 'CHAPA ABS-TPE ECO3 SENOPLAST'),
    ('110355000', 1600, 2700, 6, 'ABS-TPE', 'ECO1', 'CHAPA ABS-TPE ECO 1 SENOPLAST'),
    ('110356000', 1600, 2700, 4, 'ABS-TPE', 'ECO3', 'CHAPA ABS-TPE ECO 3 SENOPLAST'),
    ('110357000', 1600, 2700, 4, 'ABS-PMMA', 'BLACK PIANO', 'CHAPA ABS/PC BLACK PIANO SENOPLAST'),
    ('110358000', 1600, 2700, 4, 'ABS-TPE', 'ECO2', 'CHAPA ABS-TPE ECO 2 CORAL SENOPLAST'),
    ('110359000', 1600, 2700, 6, 'ABS-PMMA', 'BRANCO GEADA', 'CHAPA ABS/PC BRANCO GEADA SENOPLAST'),
    ('110501000', 1600, 2400, 4, 'ABS-PMMA', 'BLACK PIANO', 'CHAPA ABS/PC BLACK PIANO SENOPLAST'),
    ('110360000', 2400, 1600, 6, 'ABS-PMMA', 'BRANCO GEADA', 'CHAPA ABS/PC BRANCO GEADA SENOPLAST'),
    ('110361000', 2150, 1050, 6, 'ABS-PMMA', 'BRANCO GEADA', 'CHAPA ABS/PC BRANCO GEADA SENOPLAST'),
]

def _padrao():
    return {
        "Centros de custo": pd.DataFrame(CENTROS, columns=["Código", "Descrição", "R$/h"]),
        "Materiais": pd.DataFrame(MATERIAIS, columns=["Código", "Descrição", "Un", "Preço (R$)"]),
        "Tintas": pd.DataFrame(TINTAS, columns=["Tinta", "Código", "Preço (R$/L)"]),
        "Operações": pd.DataFrame(OPERACOES, columns=["Operação", "C.C.", "Grupo máquina"]),
        "Constantes": pd.DataFrame(CONSTANTES, columns=["Parâmetro", "Valor", "Obs."]),
        "Densidades": pd.DataFrame(DENSIDADES, columns=["Material", "Densidade (g/cm³)"]),
        "Chapas": pd.DataFrame(CHAPAS, columns=["Código", "Comprimento (mm)", "Largura (mm)", "Espessura (mm)", "Polímero", "Cor", "Descrição"]),
    }


# -----------------------------------------------------------------------------
# PARÂMETROS
# -----------------------------------------------------------------------------

class Parametros:
    def __init__(self, tabelas):
        self.centros = tabelas["Centros de custo"]
        self.materiais = tabelas["Materiais"]
        self.tintas = tabelas["Tintas"]
        self.operacoes = tabelas["Operações"]
        self.chapas = tabelas["Chapas"]
        const = tabelas["Constantes"].dropna(subset=["Parâmetro", "Valor"])
        self.k = dict(zip(const["Parâmetro"], const["Valor"].astype(float)))
        cc = self.centros.dropna(subset=["Código", "R$/h"])
        self._cc = {int(float(cod)): float(v) for cod, v in zip(cc["Código"], cc["R$/h"])}

    def taxa(self, cc):
        """R$/h do centro de custo."""
        try:
            return self._cc[int(float(cc))]
        except (KeyError, ValueError, TypeError):
            raise ValueError(f"Centro de custo '{cc}' não encontrado na aba Parâmetros.") from None

    def preco_tinta(self, tinta):
        linha = self.tintas[self.tintas["Tinta"] == tinta]
        if linha.empty:
            raise ValueError(f"Tinta '{tinta}' não encontrada na aba Parâmetros.")
        return float(linha["Preço (R$/L)"].iloc[0])


def carregar_parametros():
    padrao = _padrao()
    existe = ARQ_PAR.exists()
    lidas = pd.read_excel(ARQ_PAR, sheet_name=None, dtype={"Código": str}) if existe else {}
    tabelas = {nome: lidas.get(nome, df) for nome, df in padrao.items()}
    # constantes novas que ainda não existem num arquivo antigo
    const, base = tabelas["Constantes"], padrao["Constantes"]
    faltam = base[~base["Parâmetro"].isin(const["Parâmetro"])]
    tabelas["Constantes"] = pd.concat([const, faltam], ignore_index=True)
    if not existe:
        salvar_parametros(tabelas)
    return Parametros(tabelas)


def salvar_parametros(tabelas):
    with pd.ExcelWriter(ARQ_PAR, engine="openpyxl") as w:
        for nome, df in tabelas.items():
            df.dropna(how="all").to_excel(w, sheet_name=nome, index=False)
    _formatar(ARQ_PAR)


# -----------------------------------------------------------------------------
# ORÇAMENTOS SALVOS
# -----------------------------------------------------------------------------

def _ler_todas():
    if not ARQ_ORC.exists():
        return {}
    return pd.read_excel(ARQ_ORC, sheet_name=None, dtype={"Código Item": str})


def _gravar(abas):
    with pd.ExcelWriter(ARQ_ORC, engine="openpyxl") as w:
        for nome, df in abas.items():
            df.to_excel(w, sheet_name=nome, index=False)
    _formatar(ARQ_ORC)


def carregar(aba):
    return _ler_todas().get(aba, pd.DataFrame())


def codigo_existe(aba, cod):
    df = carregar(aba)
    return (not df.empty) and "Código Item" in df.columns and str(cod) in df["Código Item"].astype(str).values


def salvar_linhas(por_aba):
    """por_aba = {"nome da aba": [linha(dict), ...]}. Grava tudo de uma vez."""
    abas = _ler_todas()
    for aba, linhas in por_aba.items():
        if linhas:
            abas[aba] = pd.concat([abas.get(aba, pd.DataFrame()), pd.DataFrame(linhas)], ignore_index=True)
    _gravar(abas)


def excluir(aba, cod):
    abas = _ler_todas()
    for nome in (aba, aba + " - Detalhe"):
        df = abas.get(nome)
        if df is not None and "Código Item" in df.columns:
            abas[nome] = df[df["Código Item"].astype(str) != str(cod)]
    _gravar(abas)


def _formatar(caminho):
    """Formatação das abas: cabeçalho, zebra, moeda/decimais pelo texto do cabeçalho, filtro."""
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = load_workbook(caminho)
    fonte = Font(name="Calibri", size=11)
    cab_fundo = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    cab_fonte = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid")
    borda = Border(*(Side(style="thin", color="D9D9D9"),) * 4)

    for ws in wb.worksheets:
        n_linhas, n_cols = ws.max_row, ws.max_column
        nomes = [str(cel.value or "") for cel in ws[1]]

        for linha in ws.iter_rows(min_row=1, max_row=n_linhas, max_col=n_cols):
            for cel in linha:
                cel.font = fonte
                cel.border = borda
        for cel in ws[1]:
            cel.fill, cel.font = cab_fundo, cab_fonte
            cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 30

        for i in range(2, n_linhas + 1):
            for j, nome in enumerate(nomes, start=1):
                cel = ws.cell(row=i, column=j)
                if i % 2 == 0:
                    cel.fill = zebra
                if "R$" in nome:
                    cel.number_format = "R$ #,##0.00"
                elif "(L)" in nome:
                    cel.number_format = "0.0000"
                elif any(u in nome for u in ("(kg)", "(min)", "(m²)")):
                    cel.number_format = "0.00"
                if j > 1 and not nome.startswith("Descrição"):
                    cel.alignment = Alignment(horizontal="center")

        for j, nome in enumerate(nomes, start=1):
            maior = len(nome)
            for i in range(2, n_linhas + 1):
                valor = ws.cell(row=i, column=j).value
                maior = max(maior, len(str(valor)) if valor is not None else 0)
            ws.column_dimensions[get_column_letter(j)].width = min(maior + 4, 40)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    wb.save(caminho)
