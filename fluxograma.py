"""Editor FO-030. Altera apenas células no XLSX para conservar os desenhos do Excel."""
from io import BytesIO
from contextlib import closing
from pathlib import Path
import posixpath
import re
import shutil
import subprocess
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from lxml import etree as ET
from openpyxl import load_workbook

MODELO = Path(__file__).with_name("modelo_fluxograma.xlsx")
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
OPERACOES = dict(A="Nº operação", E="Fabricar", G="Mover", I="Armazenar",
                 K="Inspecionar", M="Interno", N="Externo", O="Simbologia", P="Descrição")
REVISOES = dict(A="Nº revisão", C="Código", D="Descrição", L="Link SAE/RNC",
                N="Responsável", P="Data")
SIMBOLOS = {"E": "◇", "G": "●", "I": "◢", "K": "■"}


def abrir(conteudo):
    return load_workbook(BytesIO(conteudo), read_only=True, keep_links=False)


def campos(aba):
    if aba.title.strip().upper() == "REVISÃO":
        return {}, REVISOES, list(range(5, max(8, aba.max_row) + 1, 2))
    if "FLUXOGRAMA" not in str(aba["G1"].value).upper():
        return {}, {}, []
    legenda = next((r for r in range(7, min(aba.max_row, 500) + 1)
                    if str(aba[f"A{r}"].value).strip().upper() == "SIMBOLOGIA"), None)
    if legenda is None:
        return {}, {}, []
    return {"W1": "Data", "G5": "Código", "X5": "Versão",
            f"A{legenda + 2}": "Elaborado", f"X{legenda + 2}": "Revisado"}, OPERACOES, list(range(7, legenda))


def ler(aba, coord):
    valor = aba[coord].value
    if valor is None:
        return ""
    if hasattr(valor, "strftime"):
        return valor.strftime("%d/%m/%Y")
    return str(valor)


TIPOS = ["", "Fabricar", "Mover", "Armazenar", "Inspecionar"]
ESPECIAIS = ["", "Segurança", "Crítica", "Importante"]
COLS = ["Nº operação", "Tipo", "Processo", "Característica", "Descrição"]
GLIFOS = {"Fabricar": "◆", "Mover": "●", "Armazenar": "▼", "Inspecionar": "■",
          "Segurança": "Ⓢ", "Crítica": "▱", "Importante": "—"}


def importar(conteudo):
    import json
    # Os arquivos gerados guardam os dados editáveis para reabertura fiel.
    with ZipFile(BytesIO(conteudo)) as z:
        if "fluxograma/dados.json" in z.namelist():
            return json.loads(z.read("fluxograma/dados.json"))
        desenhos = {}
        for n in z.namelist():
            if n.startswith("xl/drawings/drawing") and n.endswith(".xml"):
                desenhos[n] = ET.fromstring(z.read(n))
        wbxml = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        caminhos = {r.get("Id"): posixpath.normpath(posixpath.join("xl", r.get("Target"))) for r in rels}
        simbolos = {}
        for sheet in wbxml.find(NS + "sheets"):
            path = caminhos[sheet.get(RID)]
            relpath = posixpath.join(posixpath.dirname(path), "_rels", posixpath.basename(path)+".rels")
            if relpath not in z.namelist():
                continue
            for rel in ET.fromstring(z.read(relpath)):
                dest = posixpath.normpath(posixpath.join(posixpath.dirname(path), rel.get("Target")))
                if dest not in desenhos:
                    continue
                for anchor in desenhos[dest]:
                    row = anchor.findtext(".//{*}from/{*}row")
                    col = anchor.findtext(".//{*}from/{*}col")
                    geom = anchor.find(".//{*}prstGeom")
                    if row is None or col is None or geom is None:
                        continue
                    tipo = {4:"Fabricar",6:"Mover",8:"Armazenar",10:"Inspecionar"}.get(int(col))
                    if tipo:
                        simbolos.setdefault((sheet.get("name"), int(row)+1), []).append(tipo)
    resultado = {}
    with closing(abrir(conteudo)) as w:
        for aba in w:
            cab, cols, linhas = campos(aba)
            if cols == REVISOES:
                resultado[aba.title] = {"revisao": True, "cab": {}, "linhas":
                    [{rot:ler(aba,f"{c}{r}") for c,rot in cols.items()} for r in linhas]}
            else:
                registros=[]
                for r in linhas:
                    tipos=simbolos.get((aba.title,r),[])
                    registros.append({"Nº operação":ler(aba,f"A{r}"),
                        "Tipo":tipos[0] if tipos else "",
                        "Processo":"Externo" if ler(aba,f"N{r}").upper()=="X" else "Interno" if ler(aba,f"M{r}").upper()=="X" else "",
                        "Característica":"", "Descrição":ler(aba,f"P{r}")})
                resultado[aba.title]={"revisao":False,"cab":{rot:ler(aba,c) for c,rot in cab.items()},
                                     "linhas":registros}
    return resultado


def gerar_excel(dados):
    import json
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.pagebreak import Break
    w=Workbook()
    w.remove(w.active)
    for nome, folha in dados.items():
        s=w.create_sheet(nome)
        revisao=folha["revisao"]
        cols=list(REVISOES.values()) if revisao else COLS
        n=len(cols)
        s.merge_cells(start_row=1,start_column=1,end_row=1,end_column=n)
        s.cell(1,1,"CONTROLE DE REVISÕES" if revisao else "FLUXOGRAMA DO PROCESSO · FO-030")
        s.row_dimensions[1].height=30
        for i,(k,v) in enumerate(folha["cab"].items(),2):
            s.merge_cells(start_row=i,start_column=1,end_row=i,end_column=n)
            s.cell(i,1,f"{k}: {v}")
            s.row_dimensions[i].height=22
        h=len(folha["cab"])+3
        for c,t in enumerate(cols,1): s.cell(h,c,t)
        for r,linha in enumerate(folha["linhas"],h+1):
            for c,k in enumerate(cols,1):
                valor=str(linha.get(k,"") or "")
                if k in ("Tipo","Característica") and valor:
                    valor=GLIFOS.get(valor,"")+" "+valor
                cell=s.cell(r,c,valor)
                cell.data_type="s"
            comprimento=len(str(linha.get("Descrição","")))
            s.row_dimensions[r].height=max(32, 15*(1+comprimento//65))
        fim=s.max_row
        for row in s:
            for c in row:
                c.font=Font(name="DejaVu Sans",size=10)
                c.alignment=Alignment(vertical="center",wrap_text=True)
                c.border=Border(bottom=Side(style="hair",color="CBD5E1"))
                if c.row in (1,h):
                    c.fill=PatternFill("solid",fgColor="14594F")
                    c.font=Font(name="DejaVu Sans",size=11,bold=True,color="FFFFFF")
        larguras=[13,20,16,22,85] if not revisao else [13,22,65,24,24,18]
        for i,largura in enumerate(larguras,1):
            from openpyxl.utils import get_column_letter
            s.column_dimensions[get_column_letter(i)].width=largura
        s.sheet_view.showGridLines=False
        s.freeze_panes=f"A{h+1}"
        s.print_title_rows=f"1:{h}"
        s.print_area=f"A1:{get_column_letter(n)}{fim}"
        s.page_setup.orientation="landscape"
        s.page_setup.paperSize=s.PAPERSIZE_A4
        s.page_setup.fitToWidth=1
        s.page_setup.fitToHeight=0
        s.sheet_properties.pageSetUpPr.fitToPage=True
    buf=BytesIO();w.save(buf)
    with ZipFile(buf,"a") as z:
        z.writestr("fluxograma/dados.json",json.dumps(dados,ensure_ascii=False))
    return buf.getvalue()


def gerar_pdf(dados):
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from xml.sax.saxutils import escape
    font=Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
    if font.exists():
        pdfmetrics.registerFont(TTFont("Fluxo",str(font)))
    nomefont="Fluxo" if font.exists() else "Helvetica"
    estilos=getSampleStyleSheet()
    for estilo in estilos.byName.values(): estilo.fontName=nomefont
    estilos["BodyText"].fontSize=8
    estilos["BodyText"].leading=11
    def par(v):return Paragraph(escape(str(v or "")).replace("\n","<br/>"),estilos["BodyText"])
    historia=[]
    for nome,f in dados.items():
        if historia:historia.append(PageBreak())
        historia.append(Paragraph(escape(nome),estilos["Heading2"]))
        for k,v in f["cab"].items():historia.append(par(f"{k}: {v}"))
        cols=list(REVISOES.values()) if f["revisao"] else COLS
        linhas=[[par(c) for c in cols]]
        for item in f["linhas"]:
            linhas.append([par((GLIFOS.get(str(item.get(c,"")),"")+" " if c in ("Tipo","Característica") else "")+str(item.get(c,"") or "")) for c in cols])
        larguras=[55,90,65,90,482] if not f["revisao"] else [50,95,310,110,120,97]
        t=Table(linhas,colWidths=larguras,repeatRows=1,splitByRow=1,splitInRow=1)
        t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#D7EAE5")),
                              ("GRID",(0,0),(-1,-1),.4,colors.HexColor("#94A3B8")),
                              ("VALIGN",(0,0),(-1,-1),"TOP"),
                              ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7)]))
        historia.append(t)
    buf=BytesIO()
    SimpleDocTemplate(buf,pagesize=landscape(A4),leftMargin=30,rightMargin=30,topMargin=28,bottomMargin=28).build(historia)
    return buf.getvalue()


def interface():
    import hashlib
    import pandas as pd
    import streamlit as st
    st.subheader("Fluxograma de processo")
    arquivo=st.file_uploader("Abrir Excel para editar",type=["xlsx"],key="fluxo_v2_upload")
    conteudo=arquivo.getvalue() if arquivo else MODELO.read_bytes()
    chave=hashlib.sha256(conteudo).hexdigest()
    if st.session_state.get("fluxo_v2_hash")!=chave:
        try: dados=importar(conteudo)
        except Exception as e:
            st.error(f"Não consegui importar a planilha: {e}")
            return
        st.session_state.update(fluxo_v2_hash=chave,fluxo_v2_dados=dados,fluxo_v2_export=None)
    dados=st.session_state["fluxo_v2_dados"]
    nome=st.selectbox("Folha",list(dados),key="fluxo_v2_folha")
    f=dados[nome]
    with st.form("fluxo_v2_form_"+chave+nome):
        cab={k:st.text_input(k,value=str(v)) for k,v in f["cab"].items()}
        config={}
        if not f["revisao"]:
            config={"Tipo":st.column_config.SelectboxColumn(options=TIPOS),
                    "Processo":st.column_config.SelectboxColumn(options=["","Interno","Externo"]),
                    "Característica":st.column_config.SelectboxColumn(options=ESPECIAIS)}
        cols=list(REVISOES.values()) if f["revisao"] else COLS
        tabela=st.data_editor(pd.DataFrame(f["linhas"],columns=cols).fillna(""),
                             column_config=config,num_rows="dynamic",hide_index=True,
                             use_container_width=True,key="fluxo_v2_edit_"+chave+nome)
        st.caption("Use a linha vazia no final para adicionar operações. Selecione linhas para excluir.")
        aplicar=st.form_submit_button("Salvar alterações")
        exportar=st.form_submit_button("Salvar e gerar Excel + PDF",type="primary")
    if aplicar or exportar:
        f["cab"]=cab
        f["linhas"]=tabela.fillna("").to_dict("records")
        st.session_state["fluxo_v2_export"]=None
        if exportar:
            try:
                st.session_state["fluxo_v2_export"]=(gerar_excel(dados),gerar_pdf(dados))
            except Exception as e:st.error(f"Falha na exportação: {e}")
        else:st.success("Folha salva. Você pode trocar de folha.")
    st.caption("Salve antes de trocar de folha. As exportações incluem todas as folhas salvas e a edição atual.")
    if st.session_state["fluxo_v2_export"]:
        x,p=st.session_state["fluxo_v2_export"]
        st.download_button("Baixar Excel",x,"Fluxograma.xlsx",key="fluxo_v2_xlsx")
        st.download_button("Baixar PDF",p,"Fluxograma.pdf",key="fluxo_v2_pdf")
