"""Gera o modelo em branco do RCF no padrão SGS (cliente AXIA Energia).

Uso:
    python3 gerar_rcf_sgs.py <saida.docx>

Logos lidos de ../assets/. Todo campo de preenchimento é um controle de
conteúdo do Word (texto, data, lista suspensa, caixa de combinação ou foto),
para o inspetor preencher sem desconfigurar o layout. Nº, revisão e data de
emissão ficam só no quadro de identificação (pág. 1); cabeçalho e rodapé
não têm campos a preencher.
"""
import copy
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ASSETS = Path(__file__).resolve().parent.parent / "assets"

CINZA = "76777A"
LARANJA = "F69138"
TEXTO = RGBColor(0x2B, 0x2C, 0x2E)
SECUND = RGBColor(0x5F, 0x60, 0x64)
FUNDO = "F2F2F3"
BORDA = "C9CACC"

SITUACOES_ITEM = [
    ("AP", "AP – Aprovado"),
    ("APC", "APC – Aprovado após correções"),
    ("NAP", "NAP – Não aprovado"),
    ("NR", "NR – Não realizado"),
    ("R", "R – Realizado"),
    ("AL", "AL – Aguardando laudo"),
    ("AU", "AU – Autorizado"),
    ("NAU", "NAU – Não autorizado"),
    ("IMP", "IMP – Improdutivo"),
    ("N/A", "N/A – Não aplicável"),
]
SITUACOES_GERAL = ["APROVADO", "APROVADO APÓS CORREÇÕES", "NÃO APROVADO", "IMPRODUTIVO"]
STATUS_PENDENCIA = ["Aberta", "Em tratamento", "Encerrada", "N/A"]
TIPOS_ENSAIO = ["de rotina", "de tipo", "especiais", "de rotina e de tipo"]
RESULTADO_PARECER = ["conformes", "não conformes", "conformes com ressalvas"]

# Textos-padrão dos ensaios testemunhados (seção 5); a caixa de combinação
# também oferece todos os controles da tabela e aceita texto livre.
TESTEMUNHOS = [
    "Fator de potência do isolamento e capacitância das buchas",
    "Resistência de isolamento dos enrolamentos",
    "Resistência de isolamento do núcleo",
    "Relação de transformação: AT × BT; BT × terciário; AT × terciário",
    "Resistência ôhmica dos enrolamentos de AT, BT e terciário",
    "Relação de transformação dos TCs",
    "Resistência de enrolamento dos TCs",
    "Resistência de isolamento dos TCs",
]

# (item, controle, situação) — None em situação = linha de grupo
CONTROLES = [
    ("1", "TRANSFORMADOR MONTADO – ENSAIOS FINAIS", None),
    ("1.1", "Aspecto visual", ""),
    ("1.2", "Dimensional", ""),
    ("1.3", "Resistência dos enrolamentos em todas as derivações do CDC", ""),
    ("1.4", "Relação de transformação", ""),
    ("1.5", "Polaridade", ""),
    ("1.6", "Deslocamento angular", ""),
    ("1.7", "Sequência de fases", ""),
    ("1.8", "Resistência de isolamento dos enrolamentos", ""),
    ("1.9", "Resistência de isolamento do núcleo e armadura", ""),
    ("1.10", "Perdas em carga e impedância de curto-circuito com ciclo de comutação", ""),
    ("1.11", "Perdas em vazio e corrente de excitação", ""),
    ("1.12", "Perdas em vazio e corrente de excitação (antes dos dielétricos)", ""),
    ("1.13", "Perdas em vazio e corrente de excitação (após os dielétricos) com ciclo de comutação a vazio", ""),
    ("1.14", "Levantamento da curva de saturação", ""),
    ("1.15", "Impulso atmosférico alta tensão e H0", ""),
    ("1.16", "Impulso atmosférico baixa tensão e X0", ""),
    ("1.17", "Impulso atmosférico terciário", ""),
    ("1.18", "Impulso de manobra", ""),
    ("1.19", "Tensão induzida de longa duração com medição de descargas parciais", ""),
    ("1.20", "Tensão aplicada à frequência industrial", ""),
    ("1.21", "Estanqueidade e resistência à pressão", ""),
    ("1.22", "Capacitância entre enrolamentos e terra e entre enrolamentos (fator de potência)", ""),
    ("1.23", "Capacitância e fator de potência das buchas", ""),
    ("1.24", "Resposta em frequência", ""),
    ("1.25", "Impedância terminal", ""),
    ("1.26", "Ponto de orvalho (URSI)", ""),
    ("1.27", "Vácuo e deformação", ""),
    ("2", "TRANSFORMADORES DE CORRENTE DE BUCHA", None),
    ("2.1", "Relação de transformação", ""),
    ("2.2", "Polaridade", ""),
    ("2.3", "Resistência elétrica dos enrolamentos", ""),
    ("2.4", "Tensão aplicada à frequência industrial", ""),
    ("2.5", "Resistência de isolamento", ""),
    ("3", "ARMÁRIO DE CONTROLE / ACESSÓRIOS E FIAÇÃO AUXILIAR", None),
    ("3.1", "Aspecto dos acessórios e fiação", ""),
    ("3.2", "Funcional dos acessórios", ""),
    ("3.3", "Tensão aplicada à frequência industrial nos componentes e circuitos", ""),
    ("3.4", "Continuidade dos circuitos elétricos de comando, controle, proteção e medição", ""),
    ("3.5", "Funcional do comutador", ""),
    ("3.6", "Resistência de isolamento", ""),
    ("3.7", "Endereçamento da fiação (anilhamento)", ""),
    ("3.8", "Ajuste da imagem térmica", ""),
    ("3.9", "Verificação dimensional", ""),
    ("3.10", "Verificação da lista de materiais", ""),
    ("3.11", "Verificação da aderência da camada de pintura", ""),
    ("3.12", "Verificação da espessura total da camada de pintura", ""),
    ("4", "ÓLEO ISOLANTE", None),
    ("4.1", "Físico-químico – inicial", ""),
    ("4.2", "Físico-químico – final", ""),
    ("4.3", "Cromatografia – antes dos dielétricos", ""),
    ("4.4", "Cromatografia – após os dielétricos", ""),
    ("5", "PINTURA", None),
    ("5.1", "Espessura de película seca", ""),
    ("5.2", "Aderência", ""),
]

# ---------- utilitários de formatação ----------

def sombrear(cell, cor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), cor)
    tcPr.append(shd)


def bordas_tabela(table, cor=BORDA, sz=4):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{lado}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(sz))
        e.set(qn("w:color"), cor)
        b.append(e)
    tblPr.append(b)


def sem_bordas(table):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{lado}")
        e.set(qn("w:val"), "nil")
        b.append(e)
    tblPr.append(b)


def borda_celula(cell, lado, cor, sz):
    tcPr = cell._tc.get_or_add_tcPr()
    tcB = tcPr.find(qn("w:tcBorders"))
    if tcB is None:
        tcB = OxmlElement("w:tcBorders")
        tcPr.append(tcB)
    e = OxmlElement(f"w:{lado}")
    e.set(qn("w:val"), "single")
    e.set(qn("w:sz"), str(sz))
    e.set(qn("w:color"), cor)
    tcB.append(e)


def larguras(table, cms):
    """Fixa largura das colunas (grade + células) para Word e LibreOffice."""
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(sum(cms) * 567)))
    for gc, w in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), cms):
        gc.set(qn("w:w"), str(int(w * 567)))
    for row in table.rows:
        for cell, w in zip(row.cells, cms):
            cell.width = Cm(w)


def repetir_cabecalho(row):
    trPr = row._tr.get_or_add_trPr()
    e = OxmlElement("w:tblHeader")
    e.set(qn("w:val"), "true")
    trPr.append(e)


def nao_quebrar(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


def escrever(cell_or_par, texto, tam=9, negrito=False, cor=TEXTO, alinh=None, italico=False):
    p = cell_or_par.paragraphs[0] if hasattr(cell_or_par, "paragraphs") else cell_or_par
    run = p.add_run(texto)
    run.font.size = Pt(tam)
    run.bold = negrito
    run.italic = italico
    run.font.color.rgb = cor
    if alinh is not None:
        p.alignment = alinh
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    return run


def campo(paragraph, instr, tam=8, negrito=False, resultado="1", cor=TEXTO):
    """Insere campo do Word (PAGE, NUMPAGES)."""
    def r(tipo=None, texto=None):
        run = paragraph.add_run()
        run.font.size = Pt(tam)
        run.bold = negrito
        run.font.color.rgb = cor
        if tipo:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), tipo)
            run._r.append(fc)
        if texto:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = texto
            run._r.append(it)
        return run
    r("begin")
    r(texto=f" {instr} ")
    r("separate")
    run = r()
    run.text = resultado
    r("end")


# ---------- controles de conteúdo ----------

BRANCO = RGBColor(0xFF, 0xFF, 0xFF)
PLACEHOLDER = "8A8B8E"

_sdt_id = [1000]


def _rpr(tam, negrito=False, cor=None, italico=False):
    rpr = OxmlElement("w:rPr")
    if negrito:
        rpr.append(OxmlElement("w:b"))
    if italico:
        rpr.append(OxmlElement("w:i"))
    if cor:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), cor)
        rpr.append(c)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(tam * 2)))
    rpr.append(sz)
    return rpr


def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(f"w:{k}"), v)
    return e


def controle(par, tipo, titulo, guia, tam=9, negrito=False, cor="2B2C2E", valor=None,
             opcoes=()):
    """Insere um controle de conteúdo (sdt) no parágrafo.

    tipo: "texto" | "data" | "lista" | "combo".
    guia: texto cinza exibido enquanto o campo está vazio.
    """
    _sdt_id[0] += 1
    sdt = OxmlElement("w:sdt")
    pr = OxmlElement("w:sdtPr")
    pr.append(_rpr(tam, negrito, cor))
    pr.append(_el("w:alias", val=titulo))
    pr.append(_el("w:tag", val=titulo.lower().replace(" ", "_")))
    pr.append(_el("w:id", val=str(_sdt_id[0])))
    pr.append(_el("w:lock", val="sdtLocked"))
    if valor is None:
        pr.append(OxmlElement("w:showingPlcHdr"))
    if tipo == "texto":
        pr.append(OxmlElement("w:text"))
    elif tipo == "data":
        d = OxmlElement("w:date")
        d.append(_el("w:dateFormat", val="dd/MM/yyyy"))
        d.append(_el("w:lid", val="pt-BR"))
        d.append(_el("w:storeMappedDataAs", val="text"))
        d.append(_el("w:calendar", val="gregorian"))
        pr.append(d)
    elif tipo in ("lista", "combo"):
        lst = OxmlElement("w:dropDownList" if tipo == "lista" else "w:comboBox")
        for o in opcoes:
            lst.append(_el("w:listItem", displayText=o, value=o))
        pr.append(lst)
    sdt.append(pr)

    content = OxmlElement("w:sdtContent")
    run = OxmlElement("w:r")
    if valor is None:
        run.append(_rpr(tam, negrito, PLACEHOLDER, italico=True))
    else:
        run.append(_rpr(tam, negrito, cor))
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = valor if valor is not None else guia
    run.append(t)
    content.append(run)
    sdt.append(content)
    par._p.append(sdt)
    return sdt


def controle_foto(par, largura_cm, altura_cm):
    """Controle de imagem: clique no ícone e a foto se ajusta ao quadro."""
    _sdt_id[0] += 1
    run = par.add_run()
    run.add_picture(str(ASSETS / "foto_placeholder.png"), width=Cm(largura_cm), height=Cm(altura_cm))
    r = run._r
    par._p.remove(r)
    sdt = OxmlElement("w:sdt")
    pr = OxmlElement("w:sdtPr")
    pr.append(_el("w:alias", val="Foto"))
    pr.append(_el("w:tag", val="foto"))
    pr.append(_el("w:id", val=str(_sdt_id[0])))
    pr.append(_el("w:lock", val="sdtLocked"))
    pr.append(OxmlElement("w:showingPlcHdr"))
    pr.append(OxmlElement("w:picture"))
    sdt.append(pr)
    content = OxmlElement("w:sdtContent")
    content.append(r)
    sdt.append(content)
    par._p.append(sdt)


def lista_suspensa(cell, opcoes, valor=None, tam=8.5, titulo="Situação", alinh=WD_ALIGN_PARAGRAPH.CENTER):
    opcoes = [o[0] if isinstance(o, tuple) else o for o in opcoes]
    p = cell.paragraphs[0]
    p.alignment = alinh
    controle(p, "lista", titulo, "Escolher ▾", tam=tam, negrito=True, valor=valor, opcoes=opcoes)


def campo_celula(cell, guia, titulo=None, tipo="texto", tam=9, negrito=False, **kw):
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    controle(p, tipo, titulo or guia.strip("[]"), guia, tam=tam, negrito=negrito, **kw)


# ---------- blocos de layout ----------

def titulo_secao(doc, numero, texto):
    t = doc.add_table(rows=1, cols=2)
    sem_bordas(t)
    larguras(t, [0.9, 16.1])
    c0, c1 = t.rows[0].cells
    sombrear(c0, CINZA)
    escrever(c0, numero, tam=10, negrito=True, cor=BRANCO, alinh=WD_ALIGN_PARAGRAPH.CENTER)
    escrever(c1, texto.upper(), tam=10, negrito=True)
    borda_celula(c1, "bottom", LARANJA, 8)
    c0.vertical_alignment = c1.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    espaco(doc, 3)


def espaco(doc, pts=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.add_run().font.size = Pt(pts)
    return p


def tabela_simples(doc, cabecalho, linhas, cms, tam=9):
    t = doc.add_table(rows=1, cols=len(cabecalho))
    bordas_tabela(t)
    repetir_cabecalho(t.rows[0])
    for c, txt in zip(t.rows[0].cells, cabecalho):
        sombrear(c, FUNDO)
        escrever(c, txt, tam=tam - 0.5, negrito=True)
    for linha in linhas:
        row = t.add_row()
        row.height = Cm(0.5)
        cells = row.cells
        for c in cells:
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for c, txt in zip(cells, linha):
            if callable(txt):
                txt(c)
            else:
                escrever(c, txt, tam=tam)
    larguras(t, cms)
    return t


def rotulo(cell, texto):
    escrever(cell, texto.upper(), tam=6.5, cor=SECUND)
    p = cell.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    return p


def paginacao(par, tam=7.5):
    escrever(par, "Página ", tam=tam, cor=SECUND)
    campo(par, "PAGE", tam=tam, negrito=True)
    escrever(par, " de ", tam=tam, cor=SECUND)
    campo(par, "NUMPAGES", tam=tam, negrito=True)


def filete_inferior(row, cor=LARANJA, sz="18"):
    for tc in row._tr.findall(qn("w:tc")):
        tcPr = tc.get_or_add_tcPr()
        tcB = OxmlElement("w:tcBorders")
        tcB.append(_el("w:bottom", val="single", sz=sz, color=cor))
        tcPr.append(tcB)


# ---------- montagem ----------

def gerar(destino):
    doc = Document()

    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    st.font.size = Pt(10)
    st.font.color.rgb = TEXTO
    st.paragraph_format.space_after = Pt(0)
    st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.line_spacing = 1.0

    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2)
    sec.top_margin = Cm(1.2)
    sec.bottom_margin = Cm(1.5)
    sec.header_distance = Cm(0.8)
    sec.footer_distance = Cm(0.6)

    # Cabeçalho fixo: logos | título (paginação fica no rodapé)
    hdr = sec.header
    ht = hdr.add_table(rows=1, cols=4, width=Cm(17))
    sem_bordas(ht)
    larguras(ht, [3.1, 2.5, 6.9, 4.5])
    a, c, b, d = ht.rows[0].cells
    a.paragraphs[0].add_run().add_picture(str(ASSETS / "logo_sgs.png"), height=Cm(1.25))
    c.paragraphs[0].add_run().add_picture(str(ASSETS / "logo_cliente_axia.png"), height=Cm(1.15))
    c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    borda_celula(c, "left", BORDA, 8)
    borda_celula(c, "right", BORDA, 8)
    escrever(b, "INSPEÇÃO EM FÁBRICA", tam=7, cor=SECUND)
    escrever(b.add_paragraph(), "RCF – Relatório de Controle", tam=11, negrito=True)
    escrever(b.add_paragraph(), "na Fábrica", tam=11, negrito=True)
    for par in b.paragraphs:
        par.paragraph_format.left_indent = Cm(0.2)
    for cell in (a, c, b, d):
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    filete_inferior(ht.rows[0])
    hdr.paragraphs[0].paragraph_format.space_after = Pt(0)

    # Rodapé: nome e rubrica dos responsáveis em todas as páginas
    ftr = sec.footer
    ft = ftr.add_table(rows=2, cols=3, width=Cm(17))
    sem_bordas(ft)
    larguras(ft, [6.6, 6.6, 3.8])
    r0, r1 = ft.rows
    for cell, quem in zip(r0.cells[:2], ["Inspetor SGS", "Representante do fornecedor"]):
        borda_celula(cell, "top", BORDA, 6)
        escrever(cell, quem.upper(), tam=6.5, negrito=True, cor=SECUND)
        escrever(cell.add_paragraph(), "Nome: ________________________", tam=7.5)
        escrever(cell.add_paragraph(), "Rubrica: ______________________", tam=7.5)
    pg = r0.cells[2]
    borda_celula(pg, "top", BORDA, 6)
    p = pg.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paginacao(p)
    linha = r1.cells[0].merge(r1.cells[1]).merge(r1.cells[2])
    escrever(linha, "SGS do Brasil Ltda. · Inspeção Industrial  —  Documento controlado; reprodução parcial "
                    "proibida sem autorização da SGS.  ·  Cliente: AXIA Energia", tam=6.5, cor=SECUND)
    linha.paragraphs[0].paragraph_format.space_before = Pt(3)

    espaco(doc, 4)

    # Quadro de identificação — único lugar onde nº, revisão e emissão são digitados
    idt = doc.add_table(rows=5, cols=3)
    bordas_tabela(idt)
    larguras(idt, [5.6, 5.6, 5.8])
    faixa = idt.rows[0].cells[0].merge(idt.rows[0].cells[1]).merge(idt.rows[0].cells[2])
    sombrear(faixa, CINZA)
    escrever(faixa, "IDENTIFICAÇÃO DO RELATÓRIO", tam=8, negrito=True, cor=BRANCO)
    r = idt.rows[1].cells
    for cell in r:
        sombrear(cell, FUNDO)
    controle(rotulo(r[0], "Número do relatório"), "texto", "Nº do relatório", "RCF-[SIGLA]-[0000]-[AAAA]",
             tam=11, negrito=True)
    controle(rotulo(r[1], "Revisão"), "texto", "Revisão", "00", tam=11, negrito=True)
    controle(rotulo(r[2], "Data de emissão"), "data", "Data de emissão", "Escolher data ▾",
             tam=11, negrito=True)
    campos = [
        ("Documento de aquisição", "Contrato / pedido nº", "texto"),
        ("Item / lista de material", "Item ou LM", "texto"),
        ("PICQ nº", "Código do PICQ", "texto"),
        ("Empreendimento", "Nome do empreendimento", "texto"),
        ("Fornecedor / subfornecedor", "Razão social", "texto"),
        ("Data da inspeção", "Escolher data ▾", "data"),
    ]
    cells = list(idt.rows[2].cells) + list(idt.rows[3].cells)
    for cell, (rot, guia, tipo) in zip(cells, campos):
        controle(rotulo(cell, rot), tipo, rot, guia, tam=9.5)
    mat = idt.rows[4].cells[0].merge(idt.rows[4].cells[1])
    controle(rotulo(mat, "Material"), "texto", "Material", "Descrição do equipamento (tensões, potência)", tam=9.5)
    controle(rotulo(idt.rows[4].cells[2], "Nº de série / PEP"), "texto", "Nº de série", "Nº de série / PEP", tam=9.5)
    espaco(doc)

    # Situação geral (lista suspensa)
    sg = doc.add_table(rows=1, cols=2)
    bordas_tabela(sg)
    larguras(sg, [5.0, 12.0])
    s0, s1 = sg.rows[0].cells
    sombrear(s0, FUNDO)
    escrever(s0, "SITUAÇÃO DA INSPEÇÃO", tam=9, negrito=True)
    lista_suspensa(s1, SITUACOES_GERAL, tam=11, titulo="Situação geral", alinh=WD_ALIGN_PARAGRAPH.LEFT)
    for cell in (s0, s1):
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    espaco(doc)

    # Tabela de controles com lista suspensa
    tc = doc.add_table(rows=1, cols=3)
    bordas_tabela(tc)
    h = tc.rows[0]
    repetir_cabecalho(h)
    for cell, txt in zip(h.cells, ["ITEM", "CONTROLE / ENSAIO", "SITUAÇÃO ▾"]):
        sombrear(cell, CINZA)
        escrever(cell, txt, tam=8, negrito=True, cor=BRANCO,
                 alinh=None if txt == "CONTROLE / ENSAIO" else WD_ALIGN_PARAGRAPH.CENTER)
    for item, ctrl, sit in CONTROLES:
        row = tc.add_row()
        nao_quebrar(row)
        c0, c1, c2 = row.cells
        if sit is None:
            m = c0.merge(c1).merge(c2)
            sombrear(m, FUNDO)
            borda_celula(m, "left", LARANJA, 24)
            escrever(m, f"{item}  ·  {ctrl}", tam=8.5, negrito=True)
            m.paragraphs[0].paragraph_format.keep_with_next = True
        else:
            row.height = Cm(0.5)
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            escrever(c0, item, tam=8.5, cor=SECUND, alinh=WD_ALIGN_PARAGRAPH.CENTER)
            escrever(c1, ctrl, tam=8.5)
            lista_suspensa(c2, SITUACOES_ITEM, titulo=f"Situação {item}")
    larguras(tc, [1.3, 12.9, 2.8])
    espaco(doc, 2)
    leg = doc.add_paragraph()
    leg.paragraph_format.space_after = Pt(2)
    escrever(leg, "Legenda: ", tam=7.5, negrito=True, cor=SECUND)
    escrever(leg, "AP Aprovado · APC Aprovado após correções · NAP Não aprovado · NR Não realizado · "
                  "R Realizado · AL Aguardando laudo · AU Autorizado · NAU Não autorizado · IMP Improdutivo · "
                  "N/A Não aplicável", tam=7.5, cor=SECUND)
    espaco(doc, 10)

    titulo_secao(doc, "1", "Objetivo")
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.3
    escrever(p, "Atender à convocação ", tam=10)
    controle(p, "texto", "Convocação nº", "nº da convocação", tam=10)
    escrever(p, ", de ", tam=10)
    controle(p, "data", "Data da convocação", "data ▾", tam=10)
    escrever(p, ", para testemunhar a realização dos ensaios ", tam=10)
    controle(p, "combo", "Tipo de ensaio", "escolher ▾", tam=10, opcoes=TIPOS_ENSAIO)
    escrever(p, ", conforme PICQ aprovado.", tam=10)
    espaco(doc)

    def txt(guia, titulo=None, tipo="texto", **kw):
        return lambda c: campo_celula(c, guia, titulo, tipo, **kw)

    titulo_secao(doc, "2", "Documentos de referência")
    tabela_simples(doc, ["Código", "Título", "Revisão"], [
        (txt("Código do PICQ"), "Plano de Inspeção e Controle da Qualidade", txt("Rev.", "Rev. PICQ")),
        (txt("Norma ABNT / IEC"), txt("Título da norma"), txt("Ano", "Ano norma")),
        (txt("ET do cliente"), txt("Título", "Título ET"), txt("Rev.", "Rev. ET")),
        (txt("Desenho / lista de materiais"), txt("Título", "Título desenho"), txt("Rev.", "Rev. desenho")),
    ], [5.0, 9.5, 2.5])
    espaco(doc)

    titulo_secao(doc, "3", "Material inspecionado")
    tabela_simples(doc, ["Descrição", "Nº de série / PEP", "Quantidade"], [
        (txt("Descrição do equipamento"), txt("Nº de série"), txt("Qtd.")),
    ], [10.0, 4.5, 2.5])
    espaco(doc)

    titulo_secao(doc, "4", "Pendências")
    tabela_simples(doc, ["Nº", "Descrição da pendência", "Responsável", "Status ▾"], [
        (f"{n:02d}", txt("Descrever a pendência ou “Não há”", f"Pendência {n}"),
         txt("Responsável", f"Responsável {n}"),
         lambda c: lista_suspensa(c, STATUS_PENDENCIA, titulo="Status pendência"))
        for n in (1, 2, 3)
    ], [1.0, 10.0, 3.2, 2.8])
    espaco(doc)

    titulo_secao(doc, "5", "Ensaios testemunhados")
    opcoes_ensaio = TESTEMUNHOS + [c for _, c, s in CONTROLES if s is not None and c not in TESTEMUNHOS]
    linhas = [(lambda c, t=t: campo_celula(c, t, "Ensaio testemunhado", "combo", valor=t, opcoes=opcoes_ensaio),
               txt("Foto nº / boletim", "Evidência"),
               lambda c: lista_suspensa(c, SITUACOES_ITEM, titulo="Resultado"))
              for t in TESTEMUNHOS]
    linhas += [(txt("Escolher ou digitar o ensaio ▾", "Ensaio testemunhado", "combo", opcoes=opcoes_ensaio),
                txt("Foto nº / boletim", "Evidência"),
                lambda c: lista_suspensa(c, SITUACOES_ITEM, titulo="Resultado"))
               for _ in range(2)]
    tabela_simples(doc, ["Ensaio", "Evidência (foto / boletim)", "Resultado ▾"], linhas, [10.0, 4.2, 2.8])
    espaco(doc)

    titulo_secao(doc, "6", "Parecer técnico")
    pt_ = doc.add_table(rows=1, cols=1)
    bordas_tabela(pt_)
    larguras(pt_, [17])
    nao_quebrar(pt_.rows[0])
    cell = pt_.rows[0].cells[0]
    sombrear(cell, "F7F7F8")
    borda_celula(cell, "left", LARANJA, 24)
    pp = cell.paragraphs[0]
    pp.paragraph_format.line_spacing = 1.3
    escrever(pp, "Os ensaios relacionados no item 5, realizados em ", tam=10)
    controle(pp, "data", "Data dos ensaios", "data ▾", tam=10)
    escrever(pp, ", foram testemunhados pela SGS e apresentaram resultados ", tam=10)
    controle(pp, "lista", "Resultado do parecer", "escolher ▾", tam=10, negrito=True, opcoes=RESULTADO_PARECER)
    escrever(pp, " aos critérios do PICQ e das normas de referência, conforme situação registrada na tabela "
                 "de controles. ", tam=10)
    controle(pp, "texto", "Observações do parecer",
             "Ensaios pendentes, próximas datas e observações.", tam=10)
    espaco(doc)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    titulo_secao(doc, "7", "Registro fotográfico")
    ft_ = doc.add_table(rows=0, cols=2)
    sem_bordas(ft_)
    for k in range(0, 6, 2):
        r_img = ft_.add_row()
        r_leg = ft_.add_row()
        nao_quebrar(r_img)
        for j in range(2):
            ci = r_img.cells[j]
            ci.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            ci.paragraphs[0].paragraph_format.keep_with_next = True
            controle_foto(ci.paragraphs[0], 8.0, 4.5)
            cl = r_leg.cells[j]
            pl = cl.paragraphs[0]
            escrever(pl, f"Foto {k + j + 1} – ", tam=8, negrito=True, alinh=WD_ALIGN_PARAGRAPH.CENTER)
            controle(pl, "texto", f"Legenda foto {k + j + 1}", "Objeto: ensaio", tam=8)
            pl.paragraph_format.space_after = Pt(8)
    larguras(ft_, [8.5, 8.5])
    p = doc.add_paragraph()
    escrever(p, "Para mais registros: selecione as duas últimas linhas (foto + legenda), copie e cole abaixo.",
             tam=7.5, cor=RGBColor(0x8A, 0x8B, 0x8E), italico=True)
    espaco(doc, 12)

    titulo_secao(doc, "8", "Controle de revisões")
    tabela_simples(doc, ["Rev.", "Data", "Descrição", "Elaborado", "Aprovado"], [
        ("00", txt("data ▾", "Data rev. 00", "data"), "Emissão inicial",
         txt("Inspetor", "Elaborado rev. 00"), txt("Resp. técnico", "Aprovado rev. 00")),
        (txt("01", "Rev."), txt("data ▾", "Data rev.", "data"), txt("Descrição da alteração", "Descrição rev."),
         txt("Inspetor", "Elaborado rev."), txt("Resp. técnico", "Aprovado rev.")),
    ], [1.2, 2.7, 7.1, 3.0, 3.0])

    doc.save(destino)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    gerar(sys.argv[1])
