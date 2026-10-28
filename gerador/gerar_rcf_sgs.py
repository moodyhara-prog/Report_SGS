"""Gera o modelo em branco do RCF no padrão SGS (cliente AXIA Energia).

Uso:
    python3 gerar_rcf_sgs.py <saida.docx>

Logos lidos de ../assets/. Campos entre colchetes são para preenchimento;
a coluna Situação e o resultado geral são listas suspensas.
"""
import copy
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
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


def rotulo_valor(cell, rotulo, valor):
    escrever(cell, rotulo.upper(), tam=6.5, cor=SECUND)
    p = cell.add_paragraph()
    escrever(p, valor, tam=9, negrito=True)


def campo(paragraph, instr, tam=8, negrito=False):
    """Insere campo do Word (PAGE, NUMPAGES)."""
    def r(tipo=None, texto=None):
        run = paragraph.add_run()
        run.font.size = Pt(tam)
        run.bold = negrito
        run.font.color.rgb = TEXTO
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
    run.text = "1"
    r("end")


_sdt_id = [1000]


def lista_suspensa(cell, opcoes, valor, tam=8.5, negrito=True, titulo="Situação", alinh=WD_ALIGN_PARAGRAPH.CENTER):
    """Controle de conteúdo do tipo lista suspensa dentro da célula.

    opcoes: lista de (valor_exibido_na_celula, texto_na_lista) ou strings.
    """
    opcoes = [(o, o) if isinstance(o, str) else o for o in opcoes]
    p = cell.paragraphs[0]
    p.alignment = alinh
    p.paragraph_format.space_after = Pt(0)
    _sdt_id[0] += 1

    sdt = OxmlElement("w:sdt")
    pr = OxmlElement("w:sdtPr")
    rpr = OxmlElement("w:rPr")
    if negrito:
        rpr.append(OxmlElement("w:b"))
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(tam * 2)))
    rpr.append(sz)
    pr.append(rpr)
    alias = OxmlElement("w:alias")
    alias.set(qn("w:val"), titulo)
    pr.append(alias)
    tag = OxmlElement("w:tag")
    tag.set(qn("w:val"), titulo.lower().replace(" ", "_"))
    pr.append(tag)
    i = OxmlElement("w:id")
    i.set(qn("w:val"), str(_sdt_id[0]))
    pr.append(i)
    dd = OxmlElement("w:dropDownList")
    exibido = None
    for val, txt in opcoes:
        li = OxmlElement("w:listItem")
        li.set(qn("w:displayText"), val)
        li.set(qn("w:value"), val)
        dd.append(li)
        if val == valor:
            exibido = val
    dd.set(qn("w:lastValue"), exibido or "")
    pr.append(dd)
    sdt.append(pr)

    content = OxmlElement("w:sdtContent")
    run = OxmlElement("w:r")
    rpr2 = copy.deepcopy(rpr)
    run.append(rpr2)
    t = OxmlElement("w:t")
    t.text = exibido if exibido else "Escolher ▾"
    run.append(t)
    content.append(run)
    sdt.append(content)
    p._p.append(sdt)


def titulo_secao(doc, numero, texto):
    t = doc.add_table(rows=1, cols=2)
    sem_bordas(t)
    larguras(t, [0.9, 16.1])
    c0, c1 = t.rows[0].cells
    sombrear(c0, CINZA)
    escrever(c0, numero, tam=10, negrito=True, cor=RGBColor(0xFF, 0xFF, 0xFF), alinh=WD_ALIGN_PARAGRAPH.CENTER)
    escrever(c1, texto.upper(), tam=10, negrito=True)
    borda_celula(c1, "bottom", LARANJA, 8)
    c0.vertical_alignment = c1.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    espaco(doc, 3)


def espaco(doc, pts=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    r = p.add_run()
    r.font.size = Pt(pts)
    return p


def paragrafo(doc, texto, tam=10, negrito=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    escrever(p, texto, tam=tam, negrito=negrito)
    p.paragraph_format.space_after = Pt(4)
    return p


def tabela_simples(doc, cabecalho, linhas, cms, tam=9):
    t = doc.add_table(rows=1, cols=len(cabecalho))
    bordas_tabela(t)
    hdr = t.rows[0]
    repetir_cabecalho(hdr)
    for c, txt in zip(hdr.cells, cabecalho):
        sombrear(c, FUNDO)
        escrever(c, txt, tam=tam - 0.5, negrito=True)
    for linha in linhas:
        cells = t.add_row().cells
        for c, txt in zip(cells, linha):
            if callable(txt):
                txt(c)
            else:
                escrever(c, txt, tam=tam)
    larguras(t, cms)
    return t


# ---------- montagem ----------

BRANCO = RGBColor(0xFF, 0xFF, 0xFF)
PLACEHOLDER = RGBColor(0x8A, 0x8B, 0x8E)


def preencher(cell_or_par, texto="[preencher]", tam=9):
    """Texto-guia em cinza itálico, para ser substituído no preenchimento."""
    return escrever(cell_or_par, texto, tam=tam, cor=PLACEHOLDER, italico=True)


def rotulo_campo(cell, rotulo, guia):
    escrever(cell, rotulo.upper(), tam=6.5, cor=SECUND)
    preencher(cell.add_paragraph(), guia, tam=9)


def paginacao(par, tam=7.5):
    escrever(par, "Página ", tam=tam, cor=SECUND)
    campo(par, "PAGE", tam=tam, negrito=True)
    escrever(par, " de ", tam=tam, cor=SECUND)
    campo(par, "NUMPAGES", tam=tam, negrito=True)


def gerar(destino):
    doc = Document()

    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    st.font.size = Pt(10)
    st.font.color.rgb = TEXTO

    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2)
    sec.top_margin = Cm(1.2)
    sec.bottom_margin = Cm(1.5)
    sec.header_distance = Cm(0.8)
    sec.footer_distance = Cm(0.6)

    # Cabeçalho: logo SGS | título | logo cliente | bloco de controle
    hdr = sec.header
    ht = hdr.add_table(rows=2, cols=4, width=Cm(17))
    sem_bordas(ht)
    larguras(ht, [3.0, 8.4, 2.3, 3.3])
    a, b, c, d = ht.rows[0].cells
    a = a.merge(ht.rows[1].cells[0])
    b = b.merge(ht.rows[1].cells[1])
    c = c.merge(ht.rows[1].cells[2])
    a.paragraphs[0].add_run().add_picture(str(ASSETS / "logo_sgs.png"), width=Cm(3.0))
    escrever(b, "INSPEÇÃO EM FÁBRICA", tam=7, cor=SECUND)
    escrever(b.add_paragraph(), "RCF – RELATÓRIO DE CONTROLE NA FÁBRICA", tam=11.5, negrito=True)
    c.paragraphs[0].add_run().add_picture(str(ASSETS / "logo_cliente_axia.png"), width=Cm(1.9))
    c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for cell in (a, b, c):
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        borda_celula(cell, "bottom", LARANJA, 18)
    sombrear(d, CINZA)
    escrever(d, "Nº RCF-[SIGLA]-[0000]-[AAAA]", tam=8.5, negrito=True, cor=BRANCO)
    d2 = ht.rows[1].cells[3]
    borda_celula(d2, "bottom", LARANJA, 18)
    escrever(d2, "Rev. [00]  ·  Emissão [DD/MM/AAAA]", tam=7.5, cor=SECUND)
    paginacao(d2.add_paragraph())
    hdr.paragraphs[0].paragraph_format.space_after = Pt(0)

    # Rodapé: rubricas (inspetor SGS e fornecedor) em todas as páginas
    ftr = sec.footer
    ft = ftr.add_table(rows=2, cols=3, width=Cm(17))
    sem_bordas(ft)
    larguras(ft, [6.5, 6.5, 4.0])
    r0, r1 = ft.rows
    for cell, quem in zip(r0.cells[:2], ["Inspetor SGS", "Representante do fornecedor"]):
        borda_celula(cell, "top", BORDA, 6)
        escrever(cell, f"{quem.upper()}", tam=6.5, cor=SECUND)
        escrever(cell.add_paragraph(), "Rubrica: ______________________", tam=7.5)
    pg = r0.cells[2]
    borda_celula(pg, "top", BORDA, 6)
    pg.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    escrever(pg, "Nº RCF-[SIGLA]-[0000]-[AAAA]", tam=6.5, cor=SECUND)
    pp = pg.add_paragraph()
    pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paginacao(pp)
    linha = r1.cells[0].merge(r1.cells[1]).merge(r1.cells[2])
    escrever(linha, "SGS do Brasil Ltda. · Inspeção Industrial  —  Documento controlado; reprodução parcial "
                    "proibida sem autorização da SGS.  ·  Cliente: AXIA Energia", tam=6.5, cor=SECUND)
    linha.paragraphs[0].paragraph_format.space_before = Pt(3)

    # Identificação
    idt = doc.add_table(rows=2, cols=3)
    bordas_tabela(idt)
    larguras(idt, [5.4, 5.4, 6.2])
    campos = [
        ("Documento de aquisição", "[Contrato / pedido nº]"),
        ("Empreendimento", "[Nome do empreendimento]"),
        ("Fornecedor / subfornecedor", "[Razão social]"),
        ("PICQ nº", "[Código do PICQ]"),
        ("Material", "[Descrição do equipamento]"),
        ("Nº de série", "[Nº de série / PEP]"),
    ]
    for cell, (r_, g) in zip(idt.rows[0].cells + idt.rows[1].cells, campos):
        rotulo_campo(cell, r_, g)
    espaco(doc)

    # Situação geral (lista suspensa)
    sg = doc.add_table(rows=1, cols=2)
    bordas_tabela(sg)
    larguras(sg, [5.0, 12.0])
    s0, s1 = sg.rows[0].cells
    sombrear(s0, FUNDO)
    escrever(s0, "SITUAÇÃO DA INSPEÇÃO", tam=9, negrito=True)
    lista_suspensa(s1, SITUACOES_GERAL, None, tam=11, titulo="Situação geral", alinh=WD_ALIGN_PARAGRAPH.LEFT)
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
        else:
            escrever(c0, item, tam=8.5, cor=SECUND, alinh=WD_ALIGN_PARAGRAPH.CENTER)
            escrever(c1, ctrl, tam=8.5)
            lista_suspensa(c2, SITUACOES_ITEM, None)
    larguras(tc, [1.3, 12.9, 2.8])
    espaco(doc, 2)
    leg = doc.add_paragraph()
    leg.paragraph_format.space_after = Pt(2)
    escrever(leg, "Legenda: ", tam=7.5, negrito=True, cor=SECUND)
    escrever(leg, "AP Aprovado · APC Aprovado após correções · NAP Não aprovado · NR Não realizado · "
                  "R Realizado · AL Aguardando laudo · AU Autorizado · NAU Não autorizado · IMP Improdutivo · "
                  "N/A Não aplicável", tam=7.5, cor=SECUND)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    titulo_secao(doc, "1", "Objetivo")
    p = doc.add_paragraph()
    escrever(p, "Atender à convocação ", tam=10)
    preencher(p, "[nº da convocação]", tam=10)
    escrever(p, ", de ", tam=10)
    preencher(p, "[DD/MM/AAAA]", tam=10)
    escrever(p, ", para testemunhar a realização dos ensaios ", tam=10)
    preencher(p, "[de rotina / de tipo / especiais]", tam=10)
    escrever(p, ", conforme PICQ aprovado.", tam=10)
    espaco(doc)

    titulo_secao(doc, "2", "Documentos de referência")
    vazio = lambda g: (lambda c: preencher(c, g))
    tabela_simples(doc, ["Código", "Título", "Revisão"], [
        (vazio("[Código do PICQ]"), "Plano de Inspeção e Controle da Qualidade", vazio("[Rev.]")),
        (vazio("[Norma ABNT / IEC]"), vazio("[Título da norma]"), vazio("[Ano]")),
        (vazio("[ET do cliente]"), vazio("[Título]"), vazio("[Rev.]")),
        (vazio("[Desenho / lista de materiais]"), vazio("[Título]"), vazio("[Rev.]")),
    ], [5.0, 9.5, 2.5])
    espaco(doc)

    titulo_secao(doc, "3", "Material inspecionado")
    tabela_simples(doc, ["Descrição", "Nº de série / PEP", "Quantidade"], [
        (vazio("[Descrição do equipamento]"), vazio("[Nº de série]"), vazio("[Qtd.]")),
    ], [10.0, 4.5, 2.5])
    espaco(doc)

    titulo_secao(doc, "4", "Pendências")
    tabela_simples(doc, ["Nº", "Descrição da pendência", "Responsável", "Status ▾"], [
        (f"{n:02d}", vazio("[Descrição]"), vazio("[Responsável]"),
         lambda c: lista_suspensa(c, STATUS_PENDENCIA, None, titulo="Status pendência"))
        for n in (1, 2, 3)
    ], [1.0, 10.0, 3.2, 2.8])
    espaco(doc)

    titulo_secao(doc, "5", "Ensaios testemunhados")
    tabela_simples(doc, ["Ensaio", "Evidência (foto / boletim)", "Resultado ▾"], [
        (vazio("[Ensaio testemunhado]"), vazio("[Foto nº / boletim]"),
         lambda c: lista_suspensa(c, SITUACOES_ITEM, None, titulo="Resultado"))
        for _ in range(6)
    ], [10.0, 4.2, 2.8])
    espaco(doc)

    titulo_secao(doc, "6", "Parecer técnico")
    pt_ = doc.add_table(rows=1, cols=1)
    bordas_tabela(pt_)
    larguras(pt_, [17])
    cell = pt_.rows[0].cells[0]
    sombrear(cell, "F7F7F8")
    borda_celula(cell, "left", LARANJA, 24)
    pp = cell.paragraphs[0]
    pp.paragraph_format.line_spacing = 1.2
    escrever(pp, "Os ensaios relacionados no item 5, realizados em ", tam=10)
    preencher(pp, "[DD/MM/AAAA]", tam=10)
    escrever(pp, ", foram testemunhados pela SGS e apresentaram resultados ", tam=10)
    preencher(pp, "[conformes / não conformes]", tam=10)
    escrever(pp, " aos critérios do PICQ e das normas de referência, conforme situação registrada na tabela "
                 "de controles. ", tam=10)
    preencher(pp, "[Ensaios pendentes, próximas datas e observações.]", tam=10)
    espaco(doc)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    titulo_secao(doc, "7", "Registro fotográfico")
    ft_ = doc.add_table(rows=0, cols=2)
    sem_bordas(ft_)
    for k in range(0, 6, 2):
        r_img = ft_.add_row()
        r_leg = ft_.add_row()
        nao_quebrar(r_img)
        r_img.height = Cm(5.6)
        for j in range(2):
            ci = r_img.cells[j]
            bordas = ("top", "left", "bottom", "right")
            for lado in bordas:
                borda_celula(ci, lado, BORDA, 6)
            ci.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            preencher(ci, "[Inserir foto – 8 cm de largura]", tam=8)
            ci.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            ci.paragraphs[0].paragraph_format.keep_with_next = True
            cl = r_leg.cells[j]
            escrever(cl, f"Foto {k + j + 1} – ", tam=8, negrito=True, alinh=WD_ALIGN_PARAGRAPH.CENTER)
            preencher(cl, "[Objeto: ensaio]", tam=8)
            cl.paragraphs[0].paragraph_format.space_after = Pt(8)
    larguras(ft_, [8.5, 8.5])
    p = doc.add_paragraph()
    escrever(p, "Copie o par de linhas foto + legenda para incluir mais registros.", tam=7.5,
             cor=PLACEHOLDER, italico=True)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    titulo_secao(doc, "8", "Aprovação")
    at = doc.add_table(rows=2, cols=2)
    sem_bordas(at)
    larguras(at, [8.5, 8.5])
    at.rows[0].height = Cm(1.8)
    for cell, quem in zip(at.rows[1].cells, ["Representante do fornecedor", "Inspetor SGS"]):
        borda_celula(cell, "top", "2B2C2E", 6)
        escrever(cell, quem.upper(), tam=7.5, negrito=True, cor=SECUND, alinh=WD_ALIGN_PARAGRAPH.CENTER)
        for rot in ("Nome", "Função / registro", "Data"):
            q = cell.add_paragraph()
            q.alignment = WD_ALIGN_PARAGRAPH.CENTER
            escrever(q, f"{rot}: ", tam=8, cor=SECUND)
            preencher(q, "[preencher]", tam=8)
    espaco(doc, 12)

    titulo_secao(doc, "9", "Controle de revisões")
    tabela_simples(doc, ["Rev.", "Data", "Descrição", "Elaborado", "Aprovado"], [
        ("00", vazio("[DD/MM/AAAA]"), "Emissão inicial", vazio("[Inspetor]"), vazio("[Resp. técnico]")),
        ("", "", "", "", ""),
    ], [1.2, 2.7, 7.1, 3.0, 3.0])

    doc.save(destino)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    gerar(sys.argv[1])
