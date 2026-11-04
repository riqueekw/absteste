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


def gerar_excel(conteudo, alteracoes):
    saida = BytesIO()
    with ZipFile(BytesIO(conteudo)) as origem, ZipFile(saida, "w") as destino:
        livro = ET.fromstring(origem.read("xl/workbook.xml"))
        relacoes = ET.fromstring(origem.read("xl/_rels/workbook.xml.rels"))
        caminhos = {r.get("Id"): _destino("xl", r.get("Target")) for r in relacoes}
        novos = {}
        for aba in livro.find(NS + "sheets"):
            nome = aba.get("name")
            aba.attrib.pop("state", None)
            caminho = caminhos[aba.get(RID)]
            root = ET.fromstring(origem.read(caminho))
            for c in root.iter(NS + "c"):
                f = c.find(NS + "f")
                if c.get("r", "").startswith("D") and f is not None and "#REF!" in (f.text or ""):
                    _celula(root, c.get("r"), "")
            for coord, valor in alteracoes.get(nome, {}).items():
                _celula(root, coord, valor)
            novos[caminho] = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
        novos["xl/workbook.xml"] = ET.tostring(livro, encoding="UTF-8", xml_declaration=True)
        for info in origem.infolist():
            destino.writestr(info, novos.get(info.filename, origem.read(info)))
    return saida.getvalue()


def gerar_pdf(excel):
    comando = shutil.which("soffice") or shutil.which("libreoffice")
    if not comando:
        raise RuntimeError("Instale o LibreOffice Calc para gerar PDF. No Streamlit Cloud, inclua packages.txt no GitHub.")
    with TemporaryDirectory(prefix="fo030-") as pasta:
        pasta = Path(pasta)
        arquivo = pasta / "fluxograma.xlsx"
        arquivo.write_bytes(excel)
        subprocess.run(
            [comando, f"-env:UserInstallation={(pasta / 'perfil').as_uri()}", "--headless",
             "--convert-to", 'pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":"true"}}',
             "--outdir", str(pasta), str(arquivo)],
            check=True, capture_output=True, timeout=120,
        )
        pdf = pasta / "fluxograma.pdf"
        if not pdf.is_file():
            raise RuntimeError("O LibreOffice não retornou o PDF.")
        return pdf.read_bytes()


def interface():
    import hashlib
    import pandas as pd
    import streamlit as st

    st.subheader("Fluxograma de processo · FO-030")
    st.caption("Comece pelo modelo ou importe um Excel FO-030 para continuar a edição.")
    arquivo = st.file_uploader("Abrir fluxograma existente", type=["xlsx"], key="fo030_arquivo")
    conteudo = arquivo.getvalue() if arquivo else MODELO.read_bytes()
    chave = hashlib.sha256(conteudo).hexdigest()[:16]
    if st.session_state.get("fo030_fonte") != chave:
        st.session_state.update(fo030_fonte=chave, fo030_edicoes={}, fo030_pdf=None)

    try:
        livro = abrir(conteudo)
    except Exception:
        st.error("Não consegui abrir esse arquivo. Selecione uma planilha .xlsx no formato FO-030.")
        return
    with closing(livro):
        nome = st.selectbox("Aba do fluxograma", livro.sheetnames, key=f"fo030_aba_{chave}")
        aba = livro[nome]
        cabecalho, colunas, linhas = campos(aba)
        edicoes = st.session_state["fo030_edicoes"]
        anterior = edicoes.get(nome, {})
        if linhas:
            versao = hashlib.sha256(repr(anterior).encode()).hexdigest()[:12]
            prefixo = f"fo030_{chave}_{livro.sheetnames.index(nome)}_{versao}"
            with st.form(prefixo):
                valores = {}
                for coord, rotulo in cabecalho.items():
                    valores[coord] = st.text_input(rotulo, value=anterior.get(coord, ler(aba, coord)))
                tabela = pd.DataFrame([
                    {rotulo: anterior.get(f"{col}{r}", ler(aba, f"{col}{r}"))
                     for col, rotulo in colunas.items()} for r in linhas
                ])
                editada = st.data_editor(tabela, hide_index=True, use_container_width=True,
                                         num_rows="fixed", key=f"{prefixo}_tabela")
                if st.form_submit_button("Aplicar edição desta aba", type="primary"):
                    for i, r in enumerate(linhas):
                        for col, rotulo in colunas.items():
                            valor = editada.iloc[i][rotulo]
                            valores[f"{col}{r}"] = "" if pd.isna(valor) else str(valor)
                    novas = {coord: valor for coord, valor in valores.items()
                             if valor != ler(aba, coord)}
                    if novas != anterior:
                        edicoes[nome] = novas
                        st.session_state["fo030_pdf"] = None
                        st.toast("Alterações aplicadas.")
                        st.rerun()
            if colunas == OPERACOES:
                st.caption(f"{len(linhas)} linhas de operação nesta folha. "
                           "Os desenhos e conectores originais ficam no Excel e no PDF. "
                           "Para preencher as colunas de operação no modelo, use ◇ Fabricar, ● Mover, ◢ Armazenar ou ■ Inspecionar.")
        else:
            st.info("Esta aba não segue a estrutura editável do FO-030. Ela será incluída nas exportações.")

        st.caption(f"{len(livro.sheetnames)} abas incluídas no Excel e no PDF. Aplique a edição antes de trocar de aba ou exportar.")
        excel = gerar_excel(conteudo, edicoes)

    base = Path(arquivo.name).stem if arquivo else "Fluxograma_FO-030"
    col_excel, col_pdf = st.columns(2)
    col_excel.download_button("Baixar Excel", excel, f"{base}.xlsx",
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if col_pdf.button("Gerar PDF", key="fo030_gerar_pdf"):
        try:
            with st.spinner("Convertendo todas as abas em PDF…"):
                st.session_state["fo030_pdf"] = gerar_pdf(excel)
        except (RuntimeError, subprocess.SubprocessError, OSError) as erro:
            st.error(f"Não consegui gerar o PDF. {erro}")
    if st.session_state["fo030_pdf"]:
        col_pdf.download_button("Baixar PDF", st.session_state["fo030_pdf"],
                                f"{base}.pdf", "application/pdf")
