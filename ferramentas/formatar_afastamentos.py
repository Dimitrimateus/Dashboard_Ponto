import sys, pickle, datetime
from formatar_marcacoes import ler_planilha, simples, texto
# Reproduz em Python o FormatarAfastamentos.bas: lê o relatório "Histórico de Afastamentos"
# (HRCL006) exportado pelo sistema e devolve a mesma tabela que a macro deixa na aba.
# Uso: PYTHONPATH=ferramentas python3 formatar_afastamentos.py "Afastamentos.xlsx" [saida.pkl]
CAB = ["Matrícula", "Colaborador", "Admissão", "Cód. Situação", "Situação", "Início", "Hora início",
       "Término", "Hora término", "Dias", "Horas", "Prev. Término", "Exame"]

def parece_numero(v):
    t = texto(v)
    try: float(t); return t != ''
    except ValueError: return False

def formatar(linhas):
    p = dict(tipo=1, mat=2, nome=3, adm=8, ini=1, hini=2, cod=3, desc=4, fim=9, hfim=10, prev=11, exame=12)
    achou_p = achou_a = False
    for r in sorted(linhas):
        cel = linhas[r]
        if not achou_p and simples(cel.get(1)) == 'TIPO':
            for c, v in cel.items():
                t = simples(v)
                if t == 'TIPO': p['tipo'] = c
                if t == 'COLABORADOR': p['mat'] = c; p['nome'] = c + 1
                if t == 'ADMISSAO': p['adm'] = c
            achou_p = True
        if not achou_a and simples(cel.get(1)) == 'AFASTAMENTO':
            for c, v in cel.items():
                t = simples(v)
                if t == 'AFASTAMENTO': p['ini'] = c; p['hini'] = c + 1
                if t == 'SITUACAO': p['cod'] = c + 1; p['desc'] = c + 2
                if t == 'DT.TERMINO': p['fim'] = c; p['hfim'] = c + 1
                if t == 'PREV TERMINO': p['prev'] = c + 1
                if t == 'EXAME': p['exame'] = c + 1
            achou_a = True
        if achou_p and achou_a: break
    base = datetime.datetime(1899, 12, 30)
    num = lambda v: isinstance(v, float)
    out = []; mat = None
    for r in sorted(linhas):
        cel = linhas[r]; g = lambda k: cel.get(p[k])
        if parece_numero(g('tipo')) and parece_numero(g('mat')) and texto(g('nome')) and not parece_numero(g('nome')):
            if float(texto(g('tipo'))) < 1000 and float(texto(g('mat'))) > 0:
                mat = int(float(texto(g('mat')))); nome = texto(g('nome')); adm = g('adm'); continue
        vi = g('ini')
        if num(vi) and mat is not None and vi > 1000 and texto(g('desc')):
            hi, vf, hf, vp = g('hini'), g('fim'), g('hfim'), g('prev')
            ini = base + datetime.timedelta(days=int(vi))
            fim = base + datetime.timedelta(days=int(vf)) if num(vf) and vf > 1000 else None
            dias = (fim - ini).days + 1 if fim else None
            hif = hi - int(hi) if num(hi) else None
            hff = hf - int(hf) if num(hf) else None
            horas = (hff - hif) if dias == 1 and hif is not None and hff is not None and hff > hif else None
            cod = f"{int(g('cod')):03d}" if num(g('cod')) else texto(g('cod'))
            out.append([mat, nome, base + datetime.timedelta(days=int(adm)) if num(adm) else None, cod, texto(g('desc')),
                        ini, hif, fim, hff, dias, horas,
                        base + datetime.timedelta(days=int(vp)) if num(vp) and vp > 1000 else None, texto(g('exame'))])
    return out

if __name__ == '__main__':
    tab = formatar(ler_planilha(sys.argv[1]))
    pickle.dump(tab, open(sys.argv[2] if len(sys.argv) > 2 else 'afastamentos.pkl', 'wb'))
    print(len(tab), 'afastamentos de', len({x[0] for x in tab}), 'pessoas')
