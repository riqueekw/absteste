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
    ('190303000', 1600, 2700, 6, None, 'BRANCO GEADA', 'CHAPA ABS/PC BRANCO GEADA SENOPLAST'),
    ('190304000', 1600, 2700, 4, None, 'BLACK PIANO', 'CHAPA ABS/PC BLACK PIANO SENOPLAST'),
    ('190305000', 1600, 2700, 4, 'ABS-TPE', 'ECO3', 'CHAPA ABS-TPE ECO3 SENOPLAST'),
    ('110355000', 1600, 2700, 6, 'ABS-TPE', 'ECO1', 'CHAPA ABS-TPE ECO 1 SENOPLAST'),
    ('110356000', 1600, 2700, 4, 'ABS-TPE', 'ECO3', 'CHAPA ABS-TPE ECO 3 SENOPLAST'),
    ('110357000', 1600, 2700, 4, None, 'BLACK PIANO', 'CHAPA ABS/PC BLACK PIANO SENOPLAST'),
    ('110358000', 1600, 2700, 4, 'ABS-TPE', 'ECO2', 'CHAPA ABS-TPE ECO 2 CORAL SENOPLAST'),
    ('110359000', 1600, 2700, 6, None, 'BRANCO GEADA', 'CHAPA ABS/PC BRANCO GEADA SENOPLAST'),
    ('110501000', 1600, 2400, 4, None, 'BLACK PIANO', 'CHAPA ABS/PC BLACK PIANO SENOPLAST'),
    ('110360000', 2400, 1600, 6, None, 'BRANCO GEADA', 'CHAPA ABS/PC BRANCO GEADA SENOPLAST'),
    ('110361000', 2150, 1050, 6, None, 'BRANCO GEADA', 'CHAPA ABS/PC BRANCO GEADA SENOPLAST'),
]

DENSIDADES = [  # (material, g/cm³)
    ("Aço Carbono", 7.85), ("Aço Inox", 7.85), ("Alumínio", 2.7), ("Borracha Nat", 0.92),
    ("Bronze", 8.83), ("Cobre", 8.93), ("Latão", 8.5), ("Níquel", 8.91), ("Titânio", 4.55),
]

GRUPOS = [  # (código do grupo de máquina, descrição)
    ('300-001', 'ACABAMENTO - ESTILETE'),
    ('301-001', 'BANCADA 13 - GUILHOTINA'),
    ('301-002', 'GUILHOTINA MECANICA NEWTON'),
    ('301-003', 'GUILHOTINA MECANICA'),
    ('301-004', 'GUILHOTINA HIDR. GASPARIN'),
    ('301-005', 'GUILHOTINA NEWTON GHNL'),
    ('301-006', 'GUILHOTINA NEWTON GMN 2005'),
    ('301A', 'MAQ. SOLDA OXIACETILENO'),
    ('302-001', 'PRENSA EXC. JUNDIAÍ 40 T (A)'),
    ('302-002', 'PRENSA EXC. GUTMANN 130 T (A)'),
    ('302-004', 'PRENSA EXC. HARLO 130 T (A)'),
    ('302-005', 'PRENSA EXC. VITOR 60 T (A)'),
    ('302-008', 'PRENSA EXC. JUNDIAÍ 160 T (A)'),
    ('302-009', 'PRENSA YAWEI SERVO - 63TN'),
    ('302-010', 'PRENSA YAWEI SERVO - 125 TN'),
    ('303-001', 'DOBRADEIRA MECANICA CALVI-40 TN'),
    ('303-002', 'DOBRADEIRA MECANICA NEWTON-75 TON'),
    ('303-003', 'DOBRADEIRA CNC 160 TN - YAWEI'),
    ('303-006', 'DOBRADEIRA PNEUMATICA GRADE IVECO'),
    ('304-001', 'GAB. P/ SOLDAR P. LAMA'),
    ('304-003', 'CALANDRA'),
    ('304-006', 'PRENSA EXCENTRICA NR - 80 TON'),
    ('304-009', 'GAB. MONT. PORTA DIR. E ESQ.'),
    ('304-010', 'MAQUINA DE SOLDA'),
    ('304-011', 'GABARITO DE FURAÇÃO'),
    ('304-012', 'GABARITO DE MONTAGEM'),
    ('304-013', 'MAQ. SOLDA CEBORA'),
    ('304-014', 'MAQUINA SOLDA PONTO PNEUMATICA'),
    ('304-015', 'GABARITO DE MONTAR E SOLDAR ESTRIBO'),
    ('304-016', 'GABARITO DE MONTAR GRADE AGL'),
    ('304-017', 'GAB. FECHAMENTO PORTA DIREITA'),
    ('304-018', 'GAB. FECHAMENTO PORTA ESQ.'),
    ('304-019', 'FURADEIRA'),
    ('304-020', 'MAQUINA DE SOLDA CEBORA'),
    ('304-022', 'PONTEADEIRA ULTRASOLDA'),
    ('304-023', 'PONTEADEIRA CASENORTE'),
    ('304-024', 'GABARITO DE MONTAR PORTA AGL (Lado D)'),
    ('304-025', 'GABARITO DE MONTAR PORTA AGL (Lado E)'),
    ('304-026', 'PONTEADEIRA - MONT. PORTA AGL'),
    ('304-027', 'PONTEADEIRA FIXA'),
    ('304-028', 'GABARITO DE MONTAGEM'),
    ('304-029', 'MAQUINA DE SOLDA CEBORA'),
    ('304-030', 'PONTEADEIRA CASENORTE'),
    ('304-031', 'GAB. VIROLAR ABAS PORTA D/E'),
    ('304-032', 'BANCADA - PONTEADEIRA'),
    ('304-033', 'MAQ. SOLDA OXIACETILENO'),
    ('304-034', 'GABARITO DE MONTAR CANTO TRAS'),
    ('304-036', 'MAQUINA DE SOLDA WHITE MARTINS'),
    ('304-037', 'PONTEADEIRA AEREA'),
    ('304-038', 'GABARITO MONT. PORTA VOLKS L/D'),
    ('304-039', 'GABARITO MONT. PORTA VOLKS L/E'),
    ('304-041', 'MAQUINA SOLDA PONTO'),
    ('304-042', 'GABARITO DE MONTAR CAPO 608'),
    ('304-043', 'GABARITO DE MONTAR GRADE 608'),
    ('304-044', 'MÁQUINA CURVADEIRA MANUAL'),
    ('304-045', 'MAQUINA SOLDA MIG'),
    ('304-046', 'PONTEADEIRA COVER'),
    ('304-048', 'BANCADA - PONTEADEIRA'),
    ('305-008', 'EQUIPAMENTO SOLDA MIG'),
    ('305-009', 'MAQUINA DE SOLDA'),
    ('305-010', 'TRANSFORMADOR SOLDA ESAB'),
    ('305-012', 'MAQUINA DE SOLDA BALAO DE AR'),
    ('305-013', 'MAQUINA SOLDA CONJUNTO 2'),
    ('305-014', 'MAQUINA DE SOLDA'),
    ('305-015', 'GABARITO DE MONTAR PARA-CHOQUE ANTIGO'),
    ('305-016', 'MAQUINA DE SOLDA'),
    ('305-017', 'CALANDRA SALINHER'),
    ('305-018', 'GABARITO DE MONTAGEM PARA-CHOQUE IVECO'),
    ('305-019', 'GABARITO DE MONTAGEM IVECO - SOLDA'),
    ('305-021', 'FURADEIRA KHROME'),
    ('305-023', 'MONTAGEM CONEXAO'),
    ('305-025', 'BANCADA - FURADEIRAS'),
    ('305-026', 'FURADEIRA TIEKO'),
    ('305-027', 'PRENSA PNEUMATICA MONT. DIVISORIA'),
    ('305-030', 'MÁQ. SOLDA - ROBÔ'),
    ('305-031', 'GAB. SOLDA CJ CHASSI'),
    ('305-032', 'GAB. SOLDA CJ ARTICULADOR'),
    ('305-033', 'GAB. SOLDA CJ SUPORTE GERAL'),
    ('305-034', 'GAB. SOLDA CJ EIXO DIANT.'),
    ('305-035', 'GAB. SOLDA CJ BASCULANTE TAMBOR'),
    ('305-036', 'GAB. SOLDA CJ ESTRUT INF. CHASSI'),
    ('305-037', 'GAB. SOLDA CJ RESERVATORIO HIDR.'),
    ('305-038', 'GAB. SOLDA CJ CAMBÃO'),
    ('305-039', 'GAB. SOLDA CJ CAIXA DE ÁGUA'),
    ('305-040', 'GAB. SOLDA MONT. CAIXA DE ÁGUA'),
    ('305-041', 'TESTE ESTANQUEIDADE'),
    ('305-042', 'MÁQ. SOLDA - ROBÔ NOVO'),
    ('305-043', 'CÉLULA CROSSBEAN VW'),
    ('306-001', 'TESOURA UNIVERSAL FRANHO (A)'),
    ('306-002', 'GUILHOTINA MEC - MARCA SORG'),
    ('306-003', 'SERRA FITA'),
    ('306-004', 'CALANDRA'),
    ('306-005', 'BANCADA 08 - TESOURA ELETRICA'),
    ('307-002', 'ESTAGIO 1 AGL'),
    ('307-003', 'ESTAGIO 2 AGL'),
    ('307-004', 'ESTAGIO 3 AGL'),
    ('307-005', 'AJUSTE CABINE AGL'),
    ('307-006', 'GAB. MONT. OCULOS PARABRISA'),
    ('307-007', 'GAB. MONT. FRONTAL-BOLEIA 1519/25'),
    ('307-008', 'GAB. MONT. CORPO MOLE'),
    ('307-009', 'GAB. P/ ACABAMENTO'),
    ('307-012', 'GAB. MONTAGEM FOLHA PORTA INT/EXT'),
    ('307-013', 'GAB. MONT. FRONTAL CABINE'),
    ('307-014', 'GABAR. MONTAGEM BOLEIA'),
    ('307-016', 'GAB. MONTAGEM PAREDE TRASEIRA'),
    ('307-017', 'GAB. MONTAGEM PARTE TRAS.'),
    ('307-018', 'GAB. MONTAGEM ASSOALHO'),
    ('307-019', 'GAB. MONTAGEM ASSOALHO'),
    ('307-020', 'GAB. MONT. LONGARINA ESQUERDA'),
    ('307-021', 'GAB. MONT. LONGARINA DIREITA'),
    ('307-022', 'GAB. VIROLAR ABAS'),
    ('307-023', 'ESTAGIO 1 MONTAGEM HPN'),
    ('307-024', 'MONTAGEM ASSOALHO HPN'),
    ('307-025', 'ESTAGIO 2 MONTAGEM HPN'),
    ('307-026', 'MONTAGEM FRENTE 608'),
    ('307-027', 'MONTAGEM 608 ESTAGIO 1'),
    ('307-028', 'MONTAGEM 608 ESTAGIO 2'),
    ('307-029', 'MONTAGEM 608 ESTAGIO 3'),
    ('307-030', 'MONTAGEM / AJUSTE 608'),
    ('307-031', 'GAB.MONT. ESTRUT. INTERNA CAPO'),
    ('307-032', 'GAB. CHAPEAR CAPO'),
    ('307-034', 'GAB. PONTEAMENTO CAPO'),
    ('307-035', 'GAB. DOBRAR ABAS CAPO'),
    ('307-036', 'GAB. ACABAMENTO CAPO'),
    ('307-037', 'GAB.MONT. PARTE EXT. CAPO'),
    ('307-038', 'GAB. MONT. GRADE CAPO'),
    ('307-039', 'GAB. MONT. HASTE CAPO'),
    ('307-040', 'MESA MONTAGEM 1'),
    ('307-041', 'MESA MONTAGEM 6'),
    ('307-042', 'MESA 11 - CAPO'),
    ('307-043', 'GAB. MONTAGEM ASSOALHO 608/708'),
    ('307-044', 'GAB. MONTAGEM PARA-CHOQUE'),
    ('307-045', 'GAB. MONT. PAREDE TRAS. HPN'),
    ('307-046', 'GAB. MONTAGEM PORTA'),
    ('307-047', 'GAB. MONTAGEM PAINEL 608/708'),
    ('307-048', 'GABAR. MONTAR CABINE PASSAGEIRO'),
    ('307-049', 'GAB. MONTAGEM PAR. TRASEIRA 608/708'),
    ('307-050', 'GAB. MONTAGEM PAR. TRASEIRA INTERNA'),
    ('307-051', 'GAB. ALONGAR ASSOALHO'),
    ('307-052', 'GABARITO DE MONTAGEM PAREDE TRASEIRA'),
    ('307-053', 'GABARITO CHAPEAR ABAS SUPORTE INT. GRADE'),
    ('307-054', 'GAB. MONT. HASTE CAPO'),
    ('307-055', 'PONTEADEIRA'),
    ('307-056', 'RECORTAR FRONTAL DO CAPO'),
    ('307-057', 'RECORTAR FRONTAL E TRAS. DO CAPO'),
    ('307-058', 'ESTAMPAR ABAS LATERAIS E FURAR'),
    ('307-059', 'ESTAMPAR FRONTAL DO CAPO'),
    ('307-060', 'GABARITO MONT. PARA-CHOQUE'),
    ('307-073', 'GABARITO DE MONTAR CAPO'),
    ('307-075', 'PONTEADEIRA FIXA'),
    ('307-076', 'GABARITO DE MONTAGEM HPN - DIR.'),
    ('307-077', 'GABARITO DE MONTAGEM HPN - ESQ.'),
    ('307-078', 'GABARITO DE MONTAGEM HPN - DIR.'),
    ('307-079', 'GABARITO DE MONTAGEM HPN - ESQ.'),
    ('307-080', 'BANCADA - AJUSTE CHAPA CAPO HPN'),
    ('307-081', 'ÁREA MONTAGEM PARA-LAMA HPN'),
    ('307-082', 'GABARITO DE MONTAGEM CAPO 710'),
    ('307-083', 'GAB. DE MONT. CAPO FORD CARGO'),
    ('307-084', 'MESA MONTAGEM 28'),
    ('307-085', 'MESA MONTAGEM 19'),
    ('307-086', 'MESA MONTAGEM 39'),
    ('309-001', 'CABINE DE DE PINTURA FINAL'),
    ('309-002', 'PREPARACAO DE TINTAS'),
    ('310-001', 'MÁQ. TERMO FORMAGEM (MAQUINA BALAO)'),
    ('310A', 'MÁQ. TERMO FORMAGEM (MAQUINA 1)'),
    ('310B', 'MÁQ. TERMO FORMAGEM (MAQUINA 2)'),
    ('316-001', 'BANCADA MONT.'),
    ('316-002', 'ROSQUEADEIRA'),
    ('316-003', 'FURADEIRA DE BANCADA'),
    ('316-004', 'ALMOXARIFADO 1'),
    ('316-005', 'EMBALAGEM - EXPEDICAO'),
    ('316-006', 'GABARITO DE MONTAGEM FINAL IVECO'),
    ('317-001', 'MONT. FINAL CABINES - MONTAR INTERNO (A)'),
    ('317-002', 'MONT. FINAL CABINES - MONTAR FRENTE (A)'),
    ('317-003', 'MONT. FINAL CABINES - MONTAR PORTAS (A)'),
    ('317-004', 'MONT. FINAL CABINES - MONTAR ACESSOR.(A)'),
    ('317-006', 'MONT. FINAL - DIVERSOS'),
    ('317-007', 'MESA22 - BANCADA MONT CAIXA DE ÁGUA'),
    ('317-008', 'MESA23 - BANCADA MONT CJ BASCULANTE'),
    ('317-009', 'MESA24 - BANCADA MONT MOTOR ELETRICO'),
    ('317-010', 'MESA25 - BANCADA MONT RESERVATORIO HIDR'),
    ('317-011', 'POSTO1 - POSTO1 MONT. BETONEIRA'),
    ('317-012', 'POSTO2 - POSTO2 MONT. BETONEIRA'),
    ('317-013', 'POSTO3 - POSTO3 REVISÃO BETONEIRA'),
    ('317-014', 'MESA26 - BANCADA MONT TAMBOR'),
    ('317-015', 'MESA27 - BANCADA MONT CAÇAMBA'),
    ('323M', 'TUNEL DE PINTURA'),
    ('324-001', 'PRENSA HIDR. HYDRAMA 300 T (A)'),
    ('324-002', 'PRENSA HIDR. MARKUSE 600 T (A)'),
    ('324-003', 'PRENSA HIDR. 830 T (A)'),
    ('324-005', 'PRENSA HIDR. MÜLLER 1000 T (A)'),
    ('325M', 'FRESA KENDALL & GENT'),
    ('325T', 'TORNO ROMI'),
    ('332-002', 'MAQUINA DE EXPANDIR E REDUZIR TUBO'),
    ('332-013', 'ABADEIRA E RECRAVADEIRA LATA D´AGUA'),
    ('332-014', 'MAQUINA REBARBAR TUBO'),
    ('332-015', "AGRAFADEIRA LATA D'AGUA"),
    ('332-016', 'MESA 14 - ESCAPAMENTO MONTADORAS'),
    ('332-018', 'VIRTUAL ESCAPAMENTO'),
    ('332-019', 'TORNO ORBITAL'),
    ('332-022', 'MAQUINA DE ACALANAR'),
    ('332-024', 'RECRAVADEIRA'),
    ('336-001', 'PRENSA EXC. GUTMANN 250 T (A)'),
    ('336-002', 'PRENSA EXC. WEINGARTEN 250 T (A)'),
    ('336-003', 'PRENSA EXC. MAHNKE 200 T (A)'),
    ('336-004', 'PRENSA EXC. SCHULER 400 T (A)'),
    ('336-005', 'PRENSA EXC. SCHULER 350 T (A)'),
    ('336-006', 'PRENSA EXC. ERFURT DP. EF. 500 T (A)'),
    ('336-007', 'PRENSA EXC. CLEARING 500 TN'),
    ('336-008', 'PRENSA EXCENTRICA 1000 TN - SCHULER'),
    ('336-009', 'PRENSA EXCENTRICA 400 TN - SCHULER'),
    ('336-010', 'GENERICO PRENSA VERSON 2500 TN'),
    ('337-001', 'PRENSA VERSON DP. EF. 1600 T (A)'),
    ('338-001', 'CABINE DE PREPARACAO'),
    ('338-002', 'CABINE DE LAVAGEM À JATO'),
    ('338-003', 'CABINE DE ESMERILHAMENTO'),
    ('338-004', 'CABINE DE PREPARACAO DE PECAS'),
    ('338-005', 'AREA DE REFORMA E RETRABALHO (1)'),
    ('338-006', 'CABINE DE PINTURA PARACHOQUE MAN'),
    ('339-001', 'SERRA 1'),
    ('339-006', 'ACABAMENTO AC2'),
    ('339-007', 'ACABAMENTO - FURACAO'),
    ('339-008', 'ACABAMENTO AC10'),
    ('339-013', 'LIXADEIRA 1'),
    ('339-014', 'DESTACAR PEÇA 1'),
    ('339-015', 'CORTE TUPIA'),
    ('339-016', 'SERRA 3'),
    ('340-004', 'BOX DE MONTAGEM 6 - XBC'),
    ('340-005', 'BOX DE MONTAGEM 7 - HSK'),
    ('340-006', 'BOX DE MONTAGEM 8 - 140L'),
    ('340-007', 'PONTEADEIRA FIXA'),
    ('340-009', 'BOX DE MONTAGEM 12'),
    ('340-010', 'BOX DE MONTAGEM 10'),
    ('340-011', 'BOX DE MONTAGEM 11'),
    ('340-012', 'BOX DE MONTAGEM 5 - 140L'),
    ('341-001', 'PINTURA ESCAPAMENTO GALVANIZADO'),
    ('341-002', 'CABINE DE PINTURA ESCAPAMENTOS'),
    ('341-003', 'MONOVIA - PENDURAR PECA'),
    ('341-004', 'MONOVIA - LAVAR PECA'),
    ('341-005', 'MESA 12 - ESCAPAMENTO'),
    ('341-006', 'MESA 20 - ESCAPAMENTO'),
    ('341-007', 'MESA 10 - ESCAPAMENTO'),
    ('341-008', 'LAVAR GALVANIZADOS'),
    ('343A', 'MÁQ. TERMO FORMAGEM (A) MARCOPOLO'),
    ('343B', 'MÁQ. TERMO FORMAGEM (MAQUINA 4)'),
    ('343C', 'MÁQ. TERMO FORMAGEM (MAQUINA 5)'),
    ('343D', 'MÁQ. TERMO FORMAGEM (A)'),
    ('343E', 'MAQUINA TERMOFORMAGEM'),
    ('345-001', 'CABINE PINTURA EPOXI'),
    ('345-002', 'BANCADA 06 - PINTURA A PO'),
    ('346-002', 'CURVADEIRA HORIZONTAL BLM'),
    ('348-001', 'MAQUINA CORTE LASER 3D'),
    ('348-002', 'MAQUINA CORTE LASER 3D TRUMPF'),
    ('348-003', 'ROBO CORTE LASER 3D GWEIKE'),
    ('348-004', 'MAQUINA CORTE LASER 2D GWEIKE'),
    ('349-005', 'MAQUINA P/ SERRAR TUBO'),
    ('350-001', 'CABINE DE PINTURA DE PECAS'),
    ('351-001', 'SETOR DE COLAGEM 1'),
    ('351-003', 'MON/COL 1 - MARCOPOLO'),
    ('352-001', 'CENTRO ROBOTIZADO'),
    ('352-003', 'CENTRO ROBOTIZADO'),
    ('352-004', 'CENTRO ROBOTIZADO'),
    ('352-005', 'FRESADORA'),
    ('352-006', 'CENTRO ROBOTIZADO 8'),
    ('352-007', 'MAQ. SOLDA ULTRASSOM 1'),
    ('361-001', 'GABARITO DE SOLDA TALA SUSPENSYS - 2P'),
    ('361-002', 'BANCADA MONT MANITOWOC'),
    ('361-005', 'GABARITO DE SOLDA TALA SUSPENSYS - 1P'),
    ('361-006', 'MESA MONTAGEM 37'),
    ('361-007', 'BANCADA DE SOLDA PEÇAS BETONEIRA'),
    ('362-001', 'MONTAGEM FRONTAL MB'),
    ('362-002', 'PONTEADEIRA AEREA'),
    ('362-003', 'DISP FECH LATERAL CAPO MARRUA'),
    ('362-004', 'DISP FECH TRASEIRO CAPO MARRUA'),
    ('362-005', 'DISP FECH DIANTEIRO PORTA MARRUA'),
    ('362-006', 'DISP FECH TRASEIRO PORTA MARRUA'),
    ('362-007', 'DISP FECH CANALETAS PORTA MARRUA'),
    ('362-008', 'MESA MONTAGEM MARRUA'),
    ('364-001', 'INJETORA DE ALTA PRESSÃO - POLIURETANO'),
    ('364-002', 'ACABAMENTO PU'),
    ('367-001', '3101'),
    ('367-002', 'TANQUE ESTANQUIEDADE 1 e 2'),
    ('367-003', 'MAQUINA SOLDA OXIGENIO'),
    ('367-004', 'BANCADA DE AJUSTE ABERTURA'),
    ('367-005', 'BANCADA DE AJUSTE PLANICIDADE'),
    ('367-007', 'GABARITO MONTAGEM ASSOALHO MAN'),
    ('367-008', 'GABARITO MONTAGEM PAINEL TRAS. MAN'),
    ('367-009', 'GABARITO MONTAGEM LAT. TRAVES MAN'),
    ('367-010', 'GABARITO MONTAGEM TETO E FRONTAL MAN'),
    ('367-011', 'GABARITO MONTAGEM PORTAS MAN'),
    ('367-012', 'GABARITO MONTAGEM BACIÃO E FECHAM. MAN'),
    ('367-013', 'MAQ. SOLDA PORCAS QUADRADAS MAN'),
    ('367-014', 'MAQ. SOLDA PORCAS SEXTAVADA MAN'),
    ('367-015', 'MAQ. SOLDA REFORÇOS 1 MAN'),
    ('367-016', 'MAQ. SOLDA REFORÇOS 2 MAN'),
    ('TRN1', 'TREINAMENTO 1'),
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


def _padrao():
    return {
        "Centros de custo": pd.DataFrame(CENTROS, columns=["Código", "Descrição", "R$/h"]),
        "Materiais": pd.DataFrame(MATERIAIS, columns=["Código", "Descrição", "Un", "Preço (R$)"]),
        "Tintas": pd.DataFrame(TINTAS, columns=["Tinta", "Código", "Preço (R$/L)"]),
        "Operações": pd.DataFrame(OPERACOES, columns=["Operação", "C.C.", "Grupo máquina"]),
        "Constantes": pd.DataFrame(CONSTANTES, columns=["Parâmetro", "Valor", "Obs."]),
        "Grupos máquina": pd.DataFrame(GRUPOS, columns=["Código", "Descrição"]),
        "Densidades": pd.DataFrame(DENSIDADES, columns=["Material", "Densidade (g/cm³)"]),
        "Chapas": pd.DataFrame(CHAPAS, columns=["Código", "Comprimento (mm)", "Largura (mm)", "Espessura (mm)",
                                                "Polímero", "Cor", "Descrição"]),
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
        self.densidades = tabelas["Densidades"]
        self.grupos_df = tabelas["Grupos máquina"]
        gr = self.grupos_df.dropna(subset=["Código"])
        self._grupos = {str(c).strip(): ("" if pd.isna(d) else str(d)) for c, d in zip(gr["Código"], gr["Descrição"])}
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

    def grupo(self, codigo):
        """Descrição do grupo de máquina ('' se não estiver na tabela)."""
        return self._grupos.get(str(codigo).strip(), "")

    def densidade(self, material):
        linha = self.densidades[self.densidades["Material"] == material]
        if linha.empty:
            raise ValueError(f"Material '{material}' não encontrado na aba Parâmetros.")
        return float(linha["Densidade (g/cm³)"].iloc[0])

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
