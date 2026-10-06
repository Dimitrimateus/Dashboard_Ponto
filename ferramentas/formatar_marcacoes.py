import zipfile, re, html, sys, datetime, pickle, unicodedata
# Reproduz em Python o FormatarMarcacoesDigitadas.bas: lê o relatório "Marcações digitadas"
# exportado pelo sistema (o .xlsx do Senior tem caminhos com "\" dentro do zip, por isso o
# XML é lido na mão) e devolve a mesma tabela que a macro deixa na aba.
# Uso: python3 formatar_marcacoes.py "Marcações digitadas.xlsx" [saida.pkl]
CAB = ["Matrícula", "Colaborador", "Cód. Cargo", "Cargo", "Cód. Local", "Local", "Origem",
       "Data", "Hora", "Dia da semana", "Coletor", "Função", "Motivo", "Justificativa"]
DIAS = ["SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM"]

def col_num(letras):
    n = 0
    for ch in letras: n = n * 26 + ord(ch) - 64
    return n

def ler_planilha(caminho):
    z = zipfile.ZipFile(caminho)
    nome = next(n for n in z.namelist() if n.replace('\\', '/').endswith('sheet1.xml'))
    x = z.read(nome).decode('utf-8')
    linhas = {}
    for m in re.finditer(r'<row[^>]*r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        cel = {}
        for c in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)', m.group(2), re.S):
            letras, attrs, dentro = c.groups(); v = None
            if dentro:
                t = re.search(r'<t[^>]*>(.*?)</t>', dentro, re.S); vv = re.search(r'<v>(.*?)</v>', dentro)
                if t:
                    v = re.sub(r'^<!\[CDATA\[(.*)\]\]>$', r'\1', t.group(1), flags=re.S); v = html.unescape(v)
                elif vv and vv.group(1).strip():
                    v = float(vv.group(1))
            cel[col_num(letras)] = v
        linhas[int(m.group(1))] = cel
    return linhas

def simples(v):
    t = str(v or '').strip().upper()
    return unicodedata.normalize('NFD', t).encode('ascii', 'ignore').decode()

def texto(v):
    if v is None: return ''
    if isinstance(v, float) and v.is_integer(): v = int(v)
    return str(v).replace('\r', ' ').replace('\n', ' ').strip()

def formatar(linhas):
    pos = dict(cracha=1, nome=3, cargocod=6, cargo=7, local=9, origem=10, data=11, hora=12,
               coletor=13, funcao=14, motivo=15, texto=17)
    for r in sorted(linhas)[:30]:
        cel = linhas[r]
        if any(simples(v) == 'CRACHA' for v in cel.values()):
            for c, v in cel.items():
                t = simples(v)
                if t == 'CRACHA': pos['cracha'] = c
                elif t == 'COLABORADOR': pos['nome'] = c + 1
                elif t == 'CARGO': pos['cargocod'] = c; pos['cargo'] = c + 1
                elif t == 'LOCAL': pos['local'] = c
                elif t in ('ORI.', 'ORI', 'ORIGEM'): pos['origem'] = c
                elif t == 'DATA': pos['data'] = c
                elif t == 'HORA': pos['hora'] = c
                elif t == 'COLETOR': pos['coletor'] = c
                elif t == 'FUNCAO': pos['funcao'] = c
                elif t == 'JUSTIFICATIVA': pos['motivo'] = c; pos['texto'] = c + 2
            break
    base = datetime.datetime(1899, 12, 30)
    out = []; continuacoes = 0
    for r in sorted(linhas):
        cel = linhas[r]; g = lambda k: cel.get(pos[k])
        cr, dt = g('cracha'), g('data')
        if isinstance(cr, float) and isinstance(dt, float) and dt > 0:
            loc = texto(g('local')); cod, desc = (loc.split(' - ', 1) + [''])[:2] if ' - ' in loc else ('', loc)
            d = base + datetime.timedelta(days=int(dt))
            h = g('hora'); h = (h - int(h)) if isinstance(h, float) else None
            cod_txt = lambda v, n: f"{int(v):0{n}d}" if isinstance(v, float) else texto(v)
            out.append([int(cr), texto(g('nome')), texto(g('cargocod')), texto(g('cargo')), cod.strip(), desc.strip(),
                        texto(g('origem')), d, h, DIAS[d.weekday()], cod_txt(g('coletor'), 3), cod_txt(g('funcao'), 2),
                        texto(g('motivo')), texto(g('texto'))])
        elif out and texto(cr) == '':
            m, t = texto(g('motivo')), texto(g('texto'))
            if m and simples(m) != 'JUSTIFICATIVA': out[-1][12] = (out[-1][12] + ' ' + m).strip(); continuacoes += 1
            if t: out[-1][13] = (out[-1][13] + ' ' + t).strip()
    return out, continuacoes

if __name__ == '__main__':
    tab, cont = formatar(ler_planilha(sys.argv[1]))
    pickle.dump(tab, open(sys.argv[2] if len(sys.argv) > 2 else 'marcacoes.pkl', 'wb'))
    print(len(tab), 'marcações;', cont, 'motivos juntados')
