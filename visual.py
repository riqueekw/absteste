"""Desenho proporcional dos layouts de guilhotina, sem dependências extras."""
from html import escape
import math


def desenho_chapa(chapa_larg, chapa_comp, modelo, titulo="Plano de corte"):
    """SVG: comprimento no eixo horizontal; cada retângulo representa um blank."""
    valores = [chapa_larg, chapa_comp, modelo['Largura (mm)'], modelo['Comprimento (mm)']]
    if not all(math.isfinite(v) and v > 0 for v in valores):
        raise ValueError('As dimensões devem ser positivas e finitas.')
    tiras, colunas = modelo['Tiras por chapa'], modelo['Peças por tira']
    largura, comprimento = modelo['Largura (mm)'], modelo['Comprimento (mm)']
    escala = min(780 / chapa_comp, 330 / chapa_larg)
    w, h = chapa_comp * escala, chapa_larg * escala
    bw, bh = comprimento * escala, largura * escala
    x, y = 76, 78
    altura = max(260, h + 166)
    total = tiras * colunas
    labels = ''
    if 0 < total <= 80 and min(bw, bh) >= 26:
        labels = ''.join(
            f'<text x="{x+(c+.5)*bw:.3f}" y="{y+(r+.5)*bh:.3f}" class="blank">{r*colunas+c+1:02}</text>'
            for r in range(tiras) for c in range(colunas))
    sobra_l = chapa_larg - tiras * largura
    sobra_c = chapa_comp - colunas * comprimento
    vazio = (f'<text x="{x+w/2}" y="{y+h/2}" class="dim">O blank não cabe nesta orientação</text>'
             if total == 0 else '')
    return f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 920 {altura}" role="img" aria-label="{escape(titulo, quote=True)}">
    <title>{escape(titulo)}</title>
    <desc>{tiras} tiras com {colunas} blanks por tira. Sobra lateral {sobra_l:g} mm e longitudinal {sobra_c:g} mm.</desc>
    <style>text{{font-family:Arial,sans-serif;fill:#334155}}.dim{{font-size:13px;text-anchor:middle}}.blank{{font-size:12px;fill:#fff;text-anchor:middle;dominant-baseline:middle}}.small{{font-size:12px;fill:#64748b}}</style>
    <defs>
      <pattern id="retalho" width="9" height="9" patternUnits="userSpaceOnUse"><rect width="9" height="9" fill="#eef2f6"/><path d="M-2 2L2 -2M0 9L9 0M7 11L11 7" stroke="#cbd5e1" stroke-width="1"/></pattern>
      <pattern id="blanks" x="{x}" y="{y}" width="{bw}" height="{bh}" patternUnits="userSpaceOnUse"><rect width="{bw}" height="{bh}" fill="#168b80" stroke="#fff" stroke-width="1.5"/></pattern>
    </defs>
    <rect width="920" height="{altura}" rx="18" fill="#fff"/>
    <text x="30" y="29" font-size="15" font-weight="700">{escape(titulo)}</text>
    <line x1="{x}" y1="53" x2="{x+w}" y2="53" stroke="#94a3b8"/>
    <path d="M{x} 48v10 M{x+w} 48v10" stroke="#94a3b8"/>
    <text x="{x+w/2}" y="47" class="dim">{chapa_comp:g} mm · comprimento</text>
    <text transform="translate(30,{y+h/2}) rotate(-90)" class="dim">{chapa_larg:g} mm · largura</text>
    <rect x="{x}" y="{y}" width="{w}" height="{h}" fill="url(#retalho)"/>
    <rect x="{x}" y="{y}" width="{colunas*bw}" height="{tiras*bh}" fill="url(#blanks)"/>
    {labels}{vazio}
    <rect x="{x}" y="{y}" width="{w}" height="{h}" fill="none" stroke="#475569" stroke-width="1.4"/>
    <rect x="76" y="{y+h+24}" width="12" height="12" rx="2" fill="#168b80"/>
    <text x="96" y="{y+h+34}" class="small">Blanks {largura:g} × {comprimento:g} mm</text>
    <rect x="380" y="{y+h+24}" width="12" height="12" fill="url(#retalho)"/>
    <text x="400" y="{y+h+34}" class="small">Sobra de chapa</text>
    <text x="76" y="{y+h+59}" class="small">{total} blanks · sobra lateral: {sobra_l:g} mm · sobra longitudinal: {sobra_c:g} mm</text>
    </svg>'''
