# =============================================================================
# Cálculos de Engenharia - extraídos das planilhas
# =============================================================================
import math

# Densidades (g/cm³)
DENSIDADE = {
    "ABS": 1.06, "PEAD": 0.96, "PP": 0.95,
    "AÇO CARBONO": 7.85, "AÇO INOX": 7.85, "ALUMÍNIO": 2.7,
    "BORRACHA NATURAL": 0.92, "BRONZE": 8.83, "COBRE": 8.93,
    "LATÃO": 8.5, "NÍQUEL": 8.91, "TITÂNIO": 4.55,
}

# ---------------------------------------------------------------------------
# PESOS DE MATÉRIA-PRIMA
# ---------------------------------------------------------------------------

def peso_tubo_redondo(diametro_mm, espessura_mm, comprimento_mm, densidade=7.85):
    """Fórmula da planilha: ((D - e) * e) * (0.02466 * L) / 1000"""
    return ((diametro_mm - espessura_mm) * espessura_mm) * (0.02466 * comprimento_mm) / 1000

def peso_tubo_quadrado(lado_mm, espessura_mm, comprimento_mm):
    """((lado*1.27 - e) * e) * (0.02466 * L) / 1000  (mesma estrutura)"""
    # a planilha usa diametro = lado * 1.27
    d = lado_mm * 1.27
    return ((d - espessura_mm) * espessura_mm) * (0.02466 * comprimento_mm) / 1000

def peso_tubo_retangular(lado_mm, altura_mm, espessura_mm, comprimento_mm):
    d = (lado_mm + altura_mm) * 0.635
    return ((d - espessura_mm) * espessura_mm) * (0.02466 * comprimento_mm) / 1000

def peso_barra_redonda(diametro_mm, comprimento_mm):
    return (((diametro_mm ** 2) * 0.62) / 100) * comprimento_mm / 1000

def peso_barra_quadrada(base_mm, comprimento_mm):
    return (((base_mm ** 2) * 0.79) / 100) * comprimento_mm / 1000

def peso_barra_sextavada(diametro_mm, comprimento_mm):
    return (((diametro_mm ** 2) * 0.68) / 100) * (comprimento_mm / 1000)

def peso_chapa_circular(diametro_mm, espessura_mm):
    area_m2 = (math.pi * (diametro_mm ** 2) / 4) / 1_000_000
    return area_m2 * espessura_mm * 7.85

def peso_chapa_plastica(espessura_mm, largura_mm, comprimento_mm, tipo="ABS"):
    dens = DENSIDADE.get(tipo, 1.06)
    return (espessura_mm * largura_mm * comprimento_mm * dens) / 1_000_000

def peso_plastico_bolha(comprimento_mm, largura_mm=1500):
    return (comprimento_mm * 0.098) / 1000

# ---------------------------------------------------------------------------
# DESENVOLVIMENTO DE CHAPA (dobra)
# ---------------------------------------------------------------------------

def desenvolvimento_chapa(angulo_graus, espessura_mm, a_mm, b_mm, coef, raio_mm):
    """Fórmula da planilha: a + b + ((e/coef)+raio)*(angulo*pi/180)"""
    return round(
        a_mm + b_mm + (((espessura_mm / coef) + raio_mm) * (angulo_graus * math.pi / 180)), 2
    )

def comprimento_dobra_raio_medio(raio_mm, angulo_graus):
    return (math.pi * raio_mm) * (angulo_graus / 180)

def raio_medio_diametro_raio_interno(diametro_mm, raio_interno_mm):
    return (diametro_mm / 2) + raio_interno_mm

def comprimento_dobra_raio_interno(raio_interno_mm, angulo_graus, diametro_mm):
    return ((raio_interno_mm + diametro_mm / 2) * math.pi) * (angulo_graus / 180)

# ---------------------------------------------------------------------------
# DIÂMETRO ↔ PERÍMETRO
# ---------------------------------------------------------------------------

def diametro_para_perimetro(diametro_externo_mm, espessura_mm):
    pi = math.pi
    interno = diametro_externo_mm - 2 * espessura_mm
    medio = diametro_externo_mm - espessura_mm
    return {
        "interno_mm": interno,
        "medio_mm": medio,
        "externo_mm": diametro_externo_mm,
        "perimetro_interno_mm": pi * interno,
        "perimetro_medio_mm": pi * medio,
        "perimetro_externo_mm": pi * diametro_externo_mm,
    }

def perimetro_para_diametro(perimetro_externo_mm, espessura_mm):
    pi = math.pi
    diam_ext = perimetro_externo_mm / pi
    return {
        "diametro_externo_mm": diam_ext,
        "diametro_medio_mm": diam_ext - espessura_mm,
        "diametro_interno_mm": diam_ext - 2 * espessura_mm,
    }

# ---------------------------------------------------------------------------
# PRENSA / VOLUME
# ---------------------------------------------------------------------------

def tonelagem_prensa(largura_mm, comprimento_mm, espessura_mm, fator=40):
    return ((((comprimento_mm + comprimento_mm + largura_mm + largura_mm) / 1000) * espessura_mm) * fator)

def volume_recipiente(comprimento_mm, largura_mm, altura_mm):
    return (comprimento_mm / 1000) * (largura_mm / 1000) * (altura_mm / 1000) * 1000  # litros

# ---------------------------------------------------------------------------
# APROVEITAMENTO DE CHAPA + GUILHOTINA
# ---------------------------------------------------------------------------

def aproveitamento_chapa(esp_chapa_mm, larg_chapa_mm, comp_chapa_mm,
                         esp_peca_mm, larg_peca_mm, comp_peca_mm,
                         qtd_pecas_por_blank=1):
    """
    Calcula quantas peças cabem na chapa (modelo 1 e modelo 2) e escolhe o melhor.
    Retorna dict com quantidades, peso da chapa, peso da peça, retalho.
    """
    # peso chapa (kg) = esp * larg * comp * densidade/1000
    peso_chapa = esp_chapa_mm * larg_chapa_mm * comp_chapa_mm * 7.85 / 1_000_000

    # Modelo 1: tiras no sentido da largura
    n_tiras_1 = math.floor(larg_chapa_mm / larg_peca_mm)
    n_pecas_tira_1 = math.floor(comp_chapa_mm / comp_peca_mm)
    total_1 = n_tiras_1 * n_pecas_tira_1 * qtd_pecas_por_blank

    # Modelo 2: tiras giradas
    n_tiras_2 = math.floor(larg_chapa_mm / comp_peca_mm)
    n_pecas_tira_2 = math.floor(comp_chapa_mm / larg_peca_mm)
    total_2 = n_tiras_2 * n_pecas_tira_2 * qtd_pecas_por_blank

    escolhido = "Modelo 1" if total_1 >= total_2 else "Modelo 2"
    total = max(total_1, total_2)

    peso_peca = (esp_peca_mm * larg_peca_mm * comp_peca_mm * 7.85 / 1_000_000) / max(total, 1) if total > 0 else 0

    # retalho
    if escolhido == "Modelo 1":
        retalho = peso_chapa - total * peso_peca * qtd_pecas_por_blank
    else:
        retalho = peso_chapa - total * peso_peca * qtd_pecas_por_blank

    return {
        "escolhido": escolhido,
        "total_1": total_1,
        "total_2": total_2,
        "total_pecas": total,
        "peso_chapa_kg": round(peso_chapa, 3),
        "peso_peca_kg": round(peso_peca, 3),
        "retalho_kg": round(max(retalho, 0), 3),
        "n_tiras_1": n_tiras_1, "n_pecas_tira_1": n_pecas_tira_1,
        "n_tiras_2": n_tiras_2, "n_pecas_tira_2": n_pecas_tira_2,
    }

# ---------------------------------------------------------------------------
# CORTE LASER 2D / 3D
# ---------------------------------------------------------------------------

VELOCIDADE_LASER = {
    "3kW": 5000, "12kW": 25000, "3D": 4000,
}

def tempo_corte_laser(largura_mm, comprimento_mm, potencia="3kW",
                      num_furos=0, diam_furos_mm=0,
                      num_oblongos=0, comp_oblongo_mm=0, raio_oblongo_mm=0,
                      tempo_perfuracao_base_min=10, tempo_perfuracao_por_furo_min=0.2):
    """
    Replica a lógica da planilha: perímetro de corte / velocidade + tempos de perfuração.
    """
    vel = VELOCIDADE_LASER.get(potencia, 5000)

    perimetro_corte = (largura_mm * 2) + (comprimento_mm * 2)
    perimetro_furos = num_furos * (math.pi * diam_furos_mm)
    perimetro_oblongos = num_oblongos * (
        (2 * math.pi * (raio_oblongo_mm)) + (2 * comp_oblongo_mm)
    ) if raio_oblongo_mm > 0 else 0

    perimetro_total = perimetro_corte + perimetro_furos + perimetro_oblongos
    tempo_corte_min = (perimetro_total * 60) / vel
    tempo_perfuracao = tempo_perfuracao_base_min + tempo_perfuracao_por_furo_min * (num_furos + num_oblongos)
    tempo_total = math.ceil(tempo_corte_min + tempo_perfuracao)

    return {
        "perimetro_total_mm": round(perimetro_total, 2),
        "perimetro_corte_mm": round(perimetro_corte, 2),
        "perimetro_furos_mm": round(perimetro_furos, 2),
        "perimetro_oblongos_mm": round(perimetro_oblongos, 2),
        "velocidade_mm_min": vel,
        "tempo_corte_min": round(tempo_corte_min, 2),
        "tempo_perfuracao_min": round(tempo_perfuracao, 2),
        "tempo_total_min": tempo_total,
        "tempo_total_h": round(tempo_total / 60, 2),
    }

# ---------------------------------------------------------------------------
# PINTURA (túnel, pistola, pó)
# ---------------------------------------------------------------------------

def pintura_tunel(area_m2, camada_seca_micras=20, rendimento_m2_l=18,
                  perda=0.05, reducao_agua_pct=20):
    rendimento_liquido = rendimento_m2_l - (rendimento_m2_l * perda)
    rendimento_final = rendimento_liquido - (rendimento_liquido * reducao_agua_pct / 100)
    litros_2_lados = (area_m2 / rendimento_final) * 2
    return {
        "rendimento_liquido_m2_l": round(rendimento_liquido, 2),
        "rendimento_final_m2_l": round(rendimento_final, 2),
        "litros_2_lados": round(litros_2_lados, 4),
    }

def pintura_pistola_quantidade(area_mm2, lados=2, fator_perda=1.2):
    """Fórmula: ((((area_mm² * lados) / 1_000_000) / 4) * 1.2)"""
    return round((((area_mm2 / 1_000_000) * lados) / 4) * fator_perda, 4)

def pintura_pistola_tempo(area_mm2, lados=2, tempo_por_m2=3.6):
    return round((area_mm2 / 1_000_000) * tempo_por_m2 * lados, 2)

def tinta_agco(area_m2, tinta_por_m2=7, proporcao_catalisador=2, proporcao_diluente=2.5):
    """Tinta PU acrílica brilhante + catalisador + diluente"""
    tinta = area_m2 / tinta_por_m2
    catalisador = tinta / proporcao_catalisador
    diluente = tinta / proporcao_diluente
    return {
        "tinta_l": round(tinta, 3),
        "catalisador_l": round(catalisador, 3),
        "diluente_l": round(diluente, 3),
        "total_l": round(tinta + catalisador + diluente, 3),
    }

def tinta_fundo_agco(quantidade_fundo_ml, proporcoes=(5, 1, 2.5)):
    """proporções primer / catalisador / diluente"""
    p, c, d = proporcoes
    total = p + c + d
    return {
        "primer_ml": round(quantidade_fundo_ml / total * p, 2),
        "catalisador_ml": round(quantidade_fundo_ml / total * c, 2),
        "diluente_ml": round(quantidade_fundo_ml / total * d, 2),
    }

def tinta_po(area_mm2, espessura_micras=60, peso_especifico_g_cm3=1.8):
    """m²/kg = 1000/(espessura_mm * peso_especifico); qtd pó = area_m2 / (m²/kg)"""
    esp_mm = espessura_micras / 1000
    m2_por_kg = 1000 / (esp_mm * peso_especifico_g_cm3)
    area_m2 = area_mm2 / 1_000_000
    qtd_1_lado = area_m2 / m2_por_kg
    return {
        "m2_por_kg": round(m2_por_kg, 2),
        "quantidade_po_kg_1_lado": round(qtd_1_lado, 4),
        "quantidade_po_kg_2_lados": round(qtd_1_lado * 2, 4),
    }

# ---------------------------------------------------------------------------
# CUSTO MÃO DE OBRA PLÁSTICOS (R$/kg)
# ---------------------------------------------------------------------------

CUSTO_MO_PLASTICOS = {
    "ABS CINZA": 1.6, "ABS CINZA TEXT AGCO": 1.6,
    "ABS PRETO": 1.36, "ABS PRETO AGCO": 1.36,
    "NORYL": 1.575,
    "PEAD CINZA": 1.11, "PEAD CINZA ESCURO": 1.11,
    "PEAD CINZA GRAFITE": 1.11, "PEAD PRETO 1ª LINHA": 1.11,
    "PEAD PRETO 2ª LINHA": 1.11, "PP": 1.11,
}

def custo_mo_plastico(tipo, peso_kg):
    return round(CUSTO_MO_PLASTICOS.get(tipo, 0) * peso_kg, 2)
