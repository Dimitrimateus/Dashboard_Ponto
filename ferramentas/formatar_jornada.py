import sys, re, pickle, datetime
from formatar_marcacoes import ler_planilha, simples, texto
# Reproduz em Python o FormatarJornada.bas: lê um dos 4 relatórios de jornada exportados pelo
# sistema (Intrajornada, Interjornada, Interjornada semanal, Horas excedentes), descobre qual é
# pelo cabeçalho e devolve a mesma tabela que a macro deixa na aba.
# Onde se encaixa: CLAUDE.md, seção 10 (conferência sem Excel). O .pkl gerado vai para o
# adicionar_aba_xlsx.py (layouts "intrajornada", "interjornada", "intersemanal", "excedentes").
# Uso: PYTHONPATH=ferramentas python3 formatar_jornada.py relatorio.xlsx [saida.pkl]
#      -> imprime o tipo achado e grava (tipo, linhas) no .pkl

BASE = datetime.datetime(1899, 12, 30)   # dia 0 das datas do Excel

# Cabeçalho de saída de cada tipo (o mesmo que a macro escreve e que o GerarCSVPonto procura).
CAB = {
    'intrajornada': ["Matrícula", "Colaborador", "Local", "Admissão", "Cargo", "Data", "Dia", "Cód. Horário",
                     "Marcações", "Carga horária", "Intervalo 1", "Intervalo 2", "Intervalo 3", "Intervalo total"],
    'interjornada': ["Matrícula", "Colaborador", "Cód. Local", "Local", "Cargo", "C.C.", "Data apuração",
                     "Marcação anterior", "Marcação atual", "Horas descansadas", "Ocorrência"],
    'intersemanal': ["Matrícula", "Colaborador", "Cód. Local", "Local", "Data DSR/Feriado", "Ocorrência"],
    'excedentes':   ["Matrícula", "Colaborador", "Cargo", "C.C.", "Filial", "Cód. Local", "Local", "Data",
                     "Carga horária", "Marcações", "Extras", "Horas trabalhadas", "Conv. noturna", "Total"],
}

def num(v): return isinstance(v, float)
def data(v): return BASE + datetime.timedelta(days=int(v)) if num(v) and v > 1000 else None
def hora(v): return round((v - int(v)) * 1440) / 1440 if num(v) else None   # fração do dia, arredondada ao minuto (0,5 = 12:00)
def hm(f):                                                   # fração do dia -> "hh:mm"
    m = int(round(f * 1440)); return '%02d:%02d' % (m // 60 % 24, m % 60)
def local(t):                                                # "2.028.30.306 - Setor X" -> ("2.028.30.306", "Setor X")
    t = texto(t)
    return (t.split(' - ', 1)[0].strip(), t.split(' - ', 1)[1].strip()) if ' - ' in t else ('', t)

def tipo_relatorio(linhas):
    """Descobre o relatório pelos títulos das primeiras linhas (mesma regra da macro)."""
    tudo = ' | '.join(simples(v) for r in sorted(linhas)[:12] for v in linhas[r].values() if isinstance(v, str))
    if 'REGISTROS DE PONTO' in tudo and 'INTERVALO' in tudo: return 'intrajornada'
    if 'HORAS DESCANSADAS' in tudo: return 'interjornada'
    if 'HORAS DSR' in tudo or 'MARCACAO ANTES DSR' in tudo: return 'intersemanal'
    if 'CARGA HOR' in tudo and 'EXTRAS' in tudo: return 'excedentes'
    return None

def cols(linhas, primeira):
    """{TITULO SIMPLES: coluna} da 1ª linha cujo 1º título é 'primeira' (ex.: 'CADASTRO')."""
    for r in sorted(linhas):
        c = linhas[r]
        if simples(c.get(1)) == primeira:
            out = {}
            for k, v in sorted(c.items()):
                t = simples(v)
                if t and t not in out: out[t] = k
            return r, out
    return None, {}

def intrajornada(linhas):
    # Títulos: Cadastro | Nome | Local | Admissão | Cargo | Data | Dia | Registros de Ponto | Car. Horaria | Intervalo(s)
    # Valores: as batidas começam NA coluna de "Registros de Ponto" e ocupam 8 colunas; logo depois
    # vêm a carga horária e até 3 intervalos.
    _, h = cols(linhas, 'CADASTRO')
    cMat, cNome, cLocal, cAdm, cCargo = h.get('CADASTRO', 1), h.get('NOME', 2), h.get('LOCAL', 4), h.get('ADMISSAO', 7), h.get('CARGO', 8)
    cData, cDia, cReg = h.get('DATA', 10), h.get('DIA', 11), h.get('REGISTROS DE PONTO', 12)
    cCarga = cReg + 8
    out = []
    for r in sorted(linhas):
        c = linhas[r]; g = c.get
        if not (num(g(cMat)) and data(g(cData))): continue
        regs = [hm(hora(g(k))) for k in range(cReg, cReg + 8) if num(g(k))]
        dia = texto(g(cDia)); m = re.match(r'(.*?)\s*\((\w+)\)', dia)
        ints = [hora(g(k)) if num(g(k)) else None for k in (cCarga + 1, cCarga + 2, cCarga + 3)]
        out.append([int(g(cMat)), texto(g(cNome)), texto(g(cLocal)), data(g(cAdm)), texto(g(cCargo)), data(g(cData)),
                    m.group(1) if m else dia, m.group(2) if m else '', ' '.join(regs),
                    hora(g(cCarga)), ints[0], ints[1], ints[2], sum(x for x in ints if x) or 0.0])
    return out

def interjornada(linhas):
    # Títulos (linha "Data Apuração ..."): Data Apuração | Marcação Anterior | Marcação Atual |
    # Horas Descansadas | Ocorrência | Local | Cargo | C.C.  — valores deslocados: data = título+1;
    # marcação = data em título+1 e hora em título+2; horas descansadas e ocorrência = título+1.
    # A pessoa (matrícula na col. A, nome na B) vem na mesma linha; linha sem matrícula = mesma pessoa.
    h = {}
    for r in sorted(linhas):
        c = linhas[r]
        if any(simples(v) == 'HORAS DESCANSADAS' for v in c.values()):
            h = {simples(v): k for k, v in c.items() if isinstance(v, str)}; break
    cAp = h.get('DATA APURACAO', 5) + 1
    cAnt, cAtu = h.get('MARCACAO ANTERIOR', 7) + 1, h.get('MARCACAO ATUAL', 9) + 1
    cDesc, cOc = h.get('HORAS DESCANSADAS', 11) + 1, h.get('OCORRENCIA', 12) + 1
    cCargo, cCC = h.get('CARGO', 17), h.get('C.C.', 20)
    out = []; mat = nome = loc = None
    for r in sorted(linhas):
        c = linhas[r]; g = c.get
        if num(g(1)) and texto(g(2)): mat, nome, loc = int(g(1)), texto(g(2)), texto(g(5))
        if mat is None or not data(g(cAp)) or not num(g(cDesc)): continue
        ant = data(g(cAnt)); atu = data(g(cAtu))
        ant = ant + datetime.timedelta(days=hora(g(cAnt + 1)) or 0) if ant else None
        atu = atu + datetime.timedelta(days=hora(g(cAtu + 1)) or 0) if atu else None
        cod, desc = local(loc)
        out.append([mat, nome, cod, desc, texto(g(cCargo)), texto(g(cCC)), data(g(cAp)), ant, atu,
                    hora(g(cDesc)), texto(g(cOc))])
    return out

def intersemanal(linhas):
    # Cada linha de dado tem um texto "Trabalhou no Dia do DSR/Feriado (16/08/2026)" em alguma
    # coluna (H ou I: a célula é mesclada). Matrícula na col. A, nome na B, local na col. de "Local";
    # linha sem matrícula = outra data da mesma pessoa.
    h = {}
    for r in sorted(linhas):
        c = linhas[r]
        if simples(c.get(1)) == 'COLABORADOR': h = {simples(v): k for k, v in c.items() if isinstance(v, str)}; break
    cLoc = h.get('LOCAL', 7)
    out = []; mat = nome = None
    for r in sorted(linhas):
        c = linhas[r]; g = c.get
        if num(g(1)) and texto(g(2)): mat, nome = int(g(1)), texto(g(2))
        if mat is None: continue
        for k, v in sorted(c.items()):
            m = re.match(r'(.*?)\s*\((\d{2})/(\d{2})/(\d{4})\)\s*$', texto(v)) if isinstance(v, str) else None
            if m:
                cod, desc = local(g(cLoc))
                out.append([mat, nome, cod, desc, datetime.datetime(int(m.group(4)), int(m.group(3)), int(m.group(2))), m.group(1)])
                break
    return out

def excedentes(linhas):
    # Títulos: Data | Colaborador/Nome | Cargo | C. C. | Fil | Local | Carga Hor. | Local | Marcações |
    # Extras | Horas + Cv.Not. = Total. Valores: matrícula na col. de "Colaborador/Nome" e nome +1;
    # carga = título+1; local completo ("código - descrição") = 2º "Local"+1; batidas = 8 colunas a
    # partir de "Marcações"+1; logo depois: extras, horas, conversão noturna e total.
    hr, h = cols(linhas, 'DATA')
    titulos = sorted((k, simples(v)) for k, v in linhas.get(hr, {}).items() if isinstance(v, str))
    locais = [k for k, t in titulos if t == 'LOCAL']
    cData, cMat = h.get('DATA', 1), h.get('COLABORADOR/NOME', 2)
    cCargo, cCC, cFil = h.get('CARGO', 5), h.get('C. C.', 7), h.get('FIL', 8)
    cLocal1 = locais[0] if locais else 9
    cCarga = h.get('CARGA HOR.', 10) + 1
    cLocal2 = (locais[1] if len(locais) > 1 else 11) + 1
    cMarc = h.get('MARCACOES', 12) + 1
    cExtra = cMarc + 8
    out = []
    for r in sorted(linhas):
        c = linhas[r]; g = c.get
        if not (data(g(cData)) and num(g(cMat))): continue
        regs = [hm(hora(g(k))) for k in range(cMarc, cMarc + 8) if num(g(k))]
        cod, desc = local(g(cLocal2))
        cc = g(cCC); fil = g(cFil)
        out.append([int(g(cMat)), texto(g(cMat + 1)), texto(g(cCargo)), texto(cc), ('%03d' % fil) if num(fil) else texto(fil),
                    cod, desc or texto(g(cLocal1)), data(g(cData)), hora(g(cCarga)), ' '.join(regs),
                    hora(g(cExtra)), hora(g(cExtra + 1)), hora(g(cExtra + 2)), hora(g(cExtra + 3))])
    return out

FORMATAR = dict(intrajornada=intrajornada, interjornada=interjornada, intersemanal=intersemanal, excedentes=excedentes)

if __name__ == '__main__':
    linhas = ler_planilha(sys.argv[1])
    tipo = tipo_relatorio(linhas)
    if not tipo: raise SystemExit('Relatório não reconhecido')
    tab = FORMATAR[tipo](linhas)
    pickle.dump((tipo, tab), open(sys.argv[2] if len(sys.argv) > 2 else tipo + '.pkl', 'wb'))
    print(tipo, len(tab), 'linha(s) de', len({x[0] for x in tab}), 'pessoa(s)')
