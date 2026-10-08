import zipfile, re, sys, pickle, datetime
from xml.sax.saxutils import escape
# Acrescenta uma aba nova (no fim) a um .xlsx SEM regravar o resto do arquivo.
# Por que não openpyxl: as pastas de Tratamento têm validações x14, Power Query, comentários
# encadeados e tabelas que o openpyxl apaga ao salvar (ver CLAUDE.md, armadilha 9). Aqui todas
# as partes do zip são copiadas byte a byte; só mudam workbook.xml, workbook.xml.rels,
# [Content_Types].xml, docProps/app.xml e styles.xml (estilos novos no fim), e entra
# xl/worksheets/sheetN.xml. Textos vão como inlineStr (não mexe no sharedStrings).
# Uso: python3 adicionar_aba_xlsx.py entrada.xlsx saida.xlsx linhas.pkl marcacoes|afastamentos ["Nome da aba"]
# (linhas.pkl = saída do formatar_marcacoes.py ou do formatar_afastamentos.py)

# Layout de cada aba: cabeçalho, larguras, colunas de texto (códigos com zero à esquerda),
# colunas de hora (h:mm) e de duração ([h]:mm:ss). Datas são reconhecidas pelo tipo do valor.
LAYOUTS = {
    'marcacoes': dict(nome='Marcações Digitadas',
        cab=["Matrícula", "Colaborador", "Cód. Cargo", "Cargo", "Cód. Local", "Local", "Origem",
             "Data", "Hora", "Dia da semana", "Coletor", "Função", "Motivo", "Justificativa"],
        larg=[11, 34, 9, 26, 18, 28, 8, 11, 8, 9, 8, 8, 36, 60], texto={3, 5, 11, 12}, hora={9}, dur=set()),
    'afastamentos': dict(nome='Afastamentos',
        cab=["Matrícula", "Colaborador", "Admissão", "Cód. Situação", "Situação", "Início", "Hora início",
             "Término", "Hora término", "Dias", "Horas", "Prev. Término", "Exame"],
        larg=[11, 36, 11, 9, 30, 11, 11, 11, 11, 8, 8, 12, 7], texto={4}, hora={7, 9}, dur={11}),
}

# Número da coluna -> letras do Excel (1 = "A", 28 = "AB").
def letra(n):
    s = ''
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s

# Para a aba nova ter cabeçalho cinza em negrito e bordas, acrescenta no FIM do styles.xml
# uma fonte, um preenchimento, uma borda e 6 formatos de célula (os índices dos estilos que já
# existem não mudam, então as outras abas continuam iguais). add() insere um item numa lista
# do XML e corrige o atributo count dela. numFmtId: 14 = data, 20 = h:mm, 46 = [h]:mm:ss,
# 49 = texto (mantém zeros à esquerda).
def acrescentar_estilos(styles):
    """Devolve (styles novo, ids) com 5 xfs novos: cabeçalho, data, hora, número, texto."""
    def add(tag, item):
        nonlocal styles
        m = re.search(r'<%s count="(\d+)"([^>]*)>' % tag, styles)
        n = int(m.group(1))
        fim = styles.index('</%s>' % tag, m.end())
        styles = styles[:fim] + item + styles[fim:]
        styles = styles[:m.start()] + '<%s count="%d"%s>' % (tag, n + 1, m.group(2)) + styles[m.end():]
        return n
    fonte = add('fonts', '<font><b/><sz val="10"/><name val="Calibri"/><family val="2"/></font>')
    fill = add('fills', '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>')
    borda = add('borders', '<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right>'
                           '<top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>')
    ids = {}
    ids['cab'] = add('cellXfs', '<xf numFmtId="0" fontId="%d" fillId="%d" borderId="%d" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
                                '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' % (fonte, fill, borda))
    ids['data'] = add('cellXfs', '<xf numFmtId="14" fontId="0" fillId="0" borderId="%d" xfId="0" applyNumberFormat="1" applyBorder="1"/>' % borda)
    ids['hora'] = add('cellXfs', '<xf numFmtId="20" fontId="0" fillId="0" borderId="%d" xfId="0" applyNumberFormat="1" applyBorder="1"/>' % borda)
    ids['dur'] = add('cellXfs', '<xf numFmtId="46" fontId="0" fillId="0" borderId="%d" xfId="0" applyNumberFormat="1" applyBorder="1"/>' % borda)
    ids['num'] = add('cellXfs', '<xf numFmtId="0" fontId="0" fillId="0" borderId="%d" xfId="0" applyBorder="1"/>' % borda)
    ids['txt'] = add('cellXfs', '<xf numFmtId="49" fontId="0" fillId="0" borderId="%d" xfId="0" applyNumberFormat="1" applyBorder="1"/>' % borda)
    return styles, ids

# Gera o XML da aba nova: linha 1 = cabeçalho; depois uma linha por item de "linhas".
# O tipo de cada valor decide a célula: datetime -> número de série com formato de data;
# colunas de hora/duração -> fração do dia; número -> número; o resto -> texto (inlineStr,
# que não precisa mexer no sharedStrings.xml). Também congela a linha 1 e liga o filtro.
def montar_sheet(linhas, ids, lay):
    CAB, LARG, TEXTO_COLS = lay['cab'], lay['larg'], lay['texto']
    base = datetime.datetime(1899, 12, 30)
    ult = len(linhas) + 1
    rows = []
    cab = ''.join('<c r="%s1" t="inlineStr" s="%d"><is><t>%s</t></is></c>' % (letra(i + 1), ids['cab'], escape(h)) for i, h in enumerate(CAB))
    rows.append('<row r="1">%s</row>' % cab)
    for n, l in enumerate(linhas, start=2):
        cel = []
        for i, v in enumerate(l, start=1):
            ref = '%s%d' % (letra(i), n)
            if v is None or v == '':
                cel.append('<c r="%s" s="%d"/>' % (ref, ids['txt']))
            elif isinstance(v, datetime.datetime):
                cel.append('<c r="%s" s="%d"><v>%d</v></c>' % (ref, ids['data'], (v - base).days))
            elif i in lay['hora'] or i in lay['dur']:
                cel.append('<c r="%s" s="%d"><v>%r</v></c>' % (ref, ids['hora' if i in lay['hora'] else 'dur'], float(v)))
            elif isinstance(v, (int, float)) and i not in TEXTO_COLS:
                cel.append('<c r="%s" s="%d"><v>%s</v></c>' % (ref, ids['num'], v))
            else:
                cel.append('<c r="%s" t="inlineStr" s="%d"><is><t xml:space="preserve">%s</t></is></c>' % (ref, ids['txt'], escape(str(v))))
        rows.append('<row r="%d">%s</row>' % (n, ''.join(cel)))
    cols = ''.join('<col min="%d" max="%d" width="%d" customWidth="1"/>' % (i, i, w) for i, w in enumerate(LARG, start=1))
    ref = 'A1:%s%d' % (letra(len(CAB)), ult)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<dimension ref="%s"/>'
            '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/><cols>%s</cols><sheetData>%s</sheetData>'
            '<autoFilter ref="%s"/></worksheet>') % (ref, cols, ''.join(rows), ref)

# Copia o .xlsx de entrada para o de saída parte por parte, trocando só:
#   workbook.xml (registra a aba), workbook.xml.rels (aponta para o arquivo da aba),
#   [Content_Types].xml (declara o tipo), styles.xml (estilos novos no fim) e
#   docProps/app.xml (lista de nomes de abas), e acrescenta xl/worksheets/sheetN.xml.
# Recusa se já existir uma aba com o mesmo nome. Nunca altera a planilha de entrada.
def adicionar_aba(entrada, saida, linhas, nome_aba, lay):
    zin = zipfile.ZipFile(entrada)
    nomes = zin.namelist()
    wb = zin.read('xl/workbook.xml').decode('utf-8')
    if re.search(r'<sheet [^>]*name="%s"' % re.escape(escape(nome_aba, {'"': '&quot;'})), wb):
        raise SystemExit('A pasta já tem uma aba "%s".' % nome_aba)
    rels = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    ct = zin.read('[Content_Types].xml').decode('utf-8')
    n = 1
    while 'xl/worksheets/sheet%d.xml' % n in nomes: n += 1
    parte = 'xl/worksheets/sheet%d.xml' % n
    rid_n = max(int(x) for x in re.findall(r'Id="rId(\d+)"', rels)) + 1
    sheet_id = max(int(x) for x in re.findall(r'<sheet [^>]*sheetId="(\d+)"', wb)) + 1
    wb = wb.replace('</sheets>', '<sheet name="%s" sheetId="%d" r:id="rId%d"/></sheets>' % (escape(nome_aba), sheet_id, rid_n))
    rels = rels.replace('</Relationships>', '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/></Relationships>' % (rid_n, n))
    ct = ct.replace('</Types>', '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>' % parte)
    styles, ids = acrescentar_estilos(zin.read('xl/styles.xml').decode('utf-8'))
    novos = {'xl/workbook.xml': wb, 'xl/_rels/workbook.xml.rels': rels, '[Content_Types].xml': ct, 'xl/styles.xml': styles}
    if 'docProps/app.xml' in nomes:
        app = zin.read('docProps/app.xml').decode('utf-8')
        m = re.search(r'<TitlesOfParts><vt:vector size="(\d+)"', app)
        if m:
            k = int(m.group(1))
            app = app.replace(m.group(0), '<TitlesOfParts><vt:vector size="%d"' % (k + 1))
            app = app.replace('</vt:vector></TitlesOfParts>', '<vt:lpstr>%s</vt:lpstr></vt:vector></TitlesOfParts>' % escape(nome_aba))
            app = re.sub(r'(<vt:lpstr>Planilhas</vt:lpstr></vt:variant><vt:variant><vt:i4>)(\d+)', lambda mm: mm.group(1) + str(int(mm.group(2)) + 1), app)
            novos['docProps/app.xml'] = app
    with zipfile.ZipFile(saida, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            dados = novos[info.filename].encode('utf-8') if info.filename in novos else zin.read(info.filename)
            zout.writestr(info, dados)
        zout.writestr(parte, montar_sheet(linhas, ids, lay).encode('utf-8'))
    return parte

# Argumentos: entrada.xlsx saida.xlsx linhas.pkl marcacoes|afastamentos ["Nome da aba"].
if __name__ == '__main__':
    linhas = pickle.load(open(sys.argv[3], 'rb'))
    lay = LAYOUTS[sys.argv[4]]
    nome = sys.argv[5] if len(sys.argv) > 5 else lay['nome']
    print('aba gravada em', adicionar_aba(sys.argv[1], sys.argv[2], linhas, nome, lay))
