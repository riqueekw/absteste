from fluxograma import *
from copy import deepcopy
import json
def _destino(base, alvo):
    return alvo.lstrip("/") if alvo.startswith("/") else posixpath.normpath(posixpath.join(base, alvo))


def _celula(root, coord, valor):
    dados = root.find(NS + "sheetData")
    linha = int(re.search(r"\d+$", coord).group())
    row = dados.find(f"{NS}row[@r='{linha}']")
    if row is None:
        row = ET.Element(NS + "row", r=str(linha))
        anterior = next((r for r in dados if int(r.get("r", 0)) > linha), None)
        dados.append(row) if anterior is None else anterior.addprevious(row)
    c = row.find(f"{NS}c[@r='{coord}']")
    if c is None:
        c = ET.Element(NS + "c", r=coord)
        coluna = coord.rstrip("0123456789")
        anterior = next((x for x in row if x.tag == NS + "c" and
                         (len(x.get("r").rstrip("0123456789")), x.get("r").rstrip("0123456789")) >
                         (len(coluna), coluna)), None)
        row.append(c) if anterior is None else anterior.addprevious(c)
    for filho in list(c):
        if filho.tag in {NS + "v", NS + "f", NS + "is"}:
            c.remove(filho)
    c.attrib.pop("t", None)
    if valor != "":
        c.set("t", "inlineStr")
        texto = ET.SubElement(ET.SubElement(c, NS + "is"), NS + "t")
        texto.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        texto.text = valor



def simbolo(drawing, row, col, tipo):
    xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    a="http://schemas.openxmlformats.org/drawingml/2006/main"
    geo,cor={"Fabricar":("diamond","FF0000"),"Mover":("ellipse","4F81BD"),
             "Armazenar":("rtTriangle","00B050"),"Inspecionar":("rect","FFFF00"),
             "Crítica":("roundRect","FFFFFF"),"Importante":("line","000000"),
             "Segurança":("ellipse","FFFFFF")}[tipo]
    node=ET.SubElement(drawing,f"{{{xdr}}}twoCellAnchor")
    for tag,offset in [("from",40000),("to",180000)]:
        p=ET.SubElement(node,f"{{{xdr}}}{tag}")
        for k,v in [("col",col),("colOff",offset),("row",row),("rowOff",offset)]:
            ET.SubElement(p,f"{{{xdr}}}{k}").text=str(v)
    sp=ET.SubElement(node,f"{{{xdr}}}sp")
    nv=ET.SubElement(sp,f"{{{xdr}}}nvSpPr")
    ET.SubElement(nv,f"{{{xdr}}}cNvPr",id=str(1000+len(drawing)),name="Operação")
    ET.SubElement(nv,f"{{{xdr}}}cNvSpPr")
    props=ET.SubElement(sp,f"{{{xdr}}}spPr")
    ET.SubElement(ET.SubElement(props,f"{{{a}}}prstGeom",prst=geo),f"{{{a}}}avLst")
    ET.SubElement(ET.SubElement(props,f"{{{a}}}solidFill"),f"{{{a}}}srgbClr",val=cor)
    ET.SubElement(ET.SubElement(ET.SubElement(props,f"{{{a}}}ln",w="10000"),f"{{{a}}}solidFill"),f"{{{a}}}srgbClr",val="000000")
    if tipo=="Segurança":
        body=ET.SubElement(sp,f"{{{xdr}}}txBody")
        ET.SubElement(body,f"{{{a}}}bodyPr")
        ET.SubElement(body,f"{{{a}}}lstStyle")
        p=ET.SubElement(body,f"{{{a}}}p");r=ET.SubElement(p,f"{{{a}}}r")
        ET.SubElement(r,f"{{{a}}}t").text="S"
    ET.SubElement(node,f"{{{xdr}}}clientData")


def gerar_excel(dados):
    with ZipFile(MODELO) as z:
        arquivos={n:z.read(n) for n in z.namelist()}
    livro=ET.fromstring(arquivos["xl/workbook.xml"])
    sheets=livro.find(NS+"sheets");sheets.clear()
    rels=ET.fromstring(arquivos["xl/_rels/workbook.xml.rels"])
    for r in list(rels):
        if r.get("Type","").endswith("/worksheet"):rels.remove(r)
    tipos=ET.fromstring(arquivos["[Content_Types].xml"])
    ct="{http://schemas.openxmlformats.org/package/2006/content-types}"
    names=livro.find(NS+"definedNames")
    if names is not None:livro.remove(names)
    names=ET.SubElement(livro,NS+"definedNames")
    styles=ET.fromstring(arquivos["xl/styles.xml"])
    fonts=styles.find(NS+"fonts")
    font=deepcopy(fonts[0])
    font.find(NS+"name").set("val","Arial")
    fonts.append(font);fonts.set("count",str(len(fonts)))
    xfs=styles.find(NS+"cellXfs")
    mark=deepcopy(xfs[0]);mark.set("fontId",str(len(fonts)-1))
    alignment=mark.find(NS+"alignment")
    if alignment is None:alignment=ET.SubElement(mark,NS+"alignment")
    alignment.set("horizontal","center");alignment.set("vertical","center")
    xfs.append(mark);xfs.set("count",str(len(xfs)))
    arquivos["xl/styles.xml"]=ET.tostring(styles)
    idx=0;usados=set()
    for nome,f in dados.items():
        capacidade=2 if f["revisao"] else 5
        registros=f["linhas"]
        for inicio in range(0,max(1,len(registros)),capacidade):
            idx+=1
            titulo=nome if inicio==0 else f"{nome[:23]} ({inicio//capacidade+1})"
            titulo=titulo[:31]
            while titulo in usados:titulo=titulo[:27]+f"-{idx}"
            usados.add(titulo)
            rid=f"rIdFluxo{idx}"
            ET.SubElement(sheets,NS+"sheet",name=titulo,sheetId=str(idx),attrib={RID:rid})
            ET.SubElement(rels,rels[0].tag,Id=rid,Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",Target=f"worksheets/fluxo{idx}.xml")
            original=5 if f["revisao"] else 1
            root=ET.fromstring(arquivos[f"xl/worksheets/sheet{original}.xml"])
            if f["revisao"]:
                for j in range(2):
                    item=registros[inicio+j] if inicio+j<len(registros) else {}
                    for c,k in REVISOES.items():_celula(root,f"{c}{5+j*2}",str(item.get(k,"") or ""))
            else:
                for coord,k in {"W1":"Data","G5":"Código","X5":"Versão","A14":"Elaborado","X14":"Revisado"}.items():
                    _celula(root,coord,str(f["cab"].get(k,"")))
                drawing=ET.fromstring(arquivos["xl/drawings/drawing1.xml"])
                for anchor in list(drawing):
                    row=anchor.findtext(".//{*}from/{*}row")
                    if row is not None and 6<=int(row)<=10:drawing.remove(anchor)
                for j in range(5):
                    item=registros[inicio+j] if inicio+j<len(registros) else {}
                    r=7+j
                    for c in ["A","D","E","G","I","K","M","N","O","P"]:_celula(root,f"{c}{r}","")
                    _celula(root,f"A{r}",str(item.get("Nº operação","") or ""))
                    _celula(root,f"P{r}",str(item.get("Descrição","") or ""))
                    processo=item.get("Processo")
                    if processo in ["Interno","Externo"]:_celula(root,f"{'M' if processo=='Interno' else 'N'}{r}","X")
                    for c in ("M","N"):
                        cell=root.find(f".//{NS}c[@r='{c}{r}']")
                        cell.set("s",str(len(xfs)-1))
                    tipo=item.get("Tipo")
                    if tipo in TIPOS[1:]:simbolo(drawing,r-1,{"Fabricar":4,"Mover":6,"Armazenar":8,"Inspecionar":10}[tipo],tipo)
                    especial=item.get("Característica")
                    if especial in ESPECIAIS[1:]:simbolo(drawing,r-1,14,especial)
                drawname=f"drawfluxo{idx}.xml"
                arquivos[f"xl/drawings/{drawname}"]=ET.tostring(drawing)
                arquivos[f"xl/drawings/_rels/{drawname}.rels"]=arquivos["xl/drawings/_rels/drawing1.xml.rels"]
                sr=ET.fromstring(arquivos["xl/worksheets/_rels/sheet1.xml.rels"])
                for rel in sr:
                    if rel.get("Type","").endswith("/drawing"):rel.set("Target",f"../drawings/{drawname}")
                arquivos[f"xl/worksheets/_rels/fluxo{idx}.xml.rels"]=ET.tostring(sr)
                ET.SubElement(tipos,ct+"Override",PartName=f"/xl/drawings/{drawname}",ContentType="application/vnd.openxmlformats-officedocument.drawing+xml")
            setup=root.find(NS+"pageSetup")
            if setup is None:setup=ET.SubElement(root,NS+"pageSetup")
            setup.attrib.clear();setup.attrib.update({"paperSize":"9","orientation":"landscape","fitToWidth":"1","fitToHeight":"1"})
            props=root.find(NS+"sheetPr")
            if props is None:props=ET.Element(NS+"sheetPr");root.insert(0,props)
            ps=props.find(NS+"pageSetUpPr")
            if ps is None:ps=ET.SubElement(props,NS+"pageSetUpPr")
            ps.set("fitToPage","1")
            dn=ET.SubElement(names,NS+"definedName",name="_xlnm.Print_Area",localSheetId=str(idx-1))
            dn.text="'"+titulo.replace("'","''")+"'!"+("$A$1:$P$8" if f["revisao"] else "$A$1:$X$28")
            arquivos[f"xl/worksheets/fluxo{idx}.xml"]=ET.tostring(root)
            ET.SubElement(tipos,ct+"Override",PartName=f"/xl/worksheets/fluxo{idx}.xml",ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    arquivos["xl/workbook.xml"]=ET.tostring(livro)
    arquivos["xl/_rels/workbook.xml.rels"]=ET.tostring(rels)
    arquivos["[Content_Types].xml"]=ET.tostring(tipos)
    arquivos["fluxograma/dados.json"]=json.dumps(dados,ensure_ascii=False).encode()
    buf=BytesIO()
    with ZipFile(buf,"w") as z:
        for n,b in arquivos.items():z.writestr(n,b)
    return buf.getvalue()


def gerar_pdf(dados):
    comando=shutil.which("soffice") or shutil.which("libreoffice")
    if not comando:raise RuntimeError("Inclua packages.txt no GitHub para instalar LibreOffice.")
    with TemporaryDirectory(prefix="fluxo-") as d:
        p=Path(d);(p/"fluxo.xlsx").write_bytes(gerar_excel(dados))
        subprocess.run([comando,f"-env:UserInstallation={(p/'perfil').as_uri()}","--headless","--convert-to","pdf","--outdir",str(p),str(p/"fluxo.xlsx")],check=True,capture_output=True,timeout=120)
        return (p/"fluxo.pdf").read_bytes()
