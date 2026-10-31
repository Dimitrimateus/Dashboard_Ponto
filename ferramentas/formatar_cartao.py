import sys, re, pickle, datetime
from formatar_marcacoes import ler_planilha
# Reproduz em Python o ConsolidarCartaoPonto.bas: lê o HRCP102 (Cartão Ponto) exportado pelo
# sistema e devolve as linhas da aba consolidada (uma por colaborador e dia, 29 colunas, na
# ordem de CAB). Uso: python3 formatar_cartao.py "cartão ponto outubro.xlsx" saida.pkl
#   saida.pkl = {ini, fim, linhas}; o adicionar_aba_xlsx.py (layout "cartao") monta a aba com os
#   rótulos de grupo na linha 1, o cabeçalho na linha 2 e os dados a partir da linha 3.
# Mesmas regras da macro: matrícula/nome da linha "Empregado:", cargo de "Cargo:", setor = col. D
# de "Localização:", lista "Horários:" (código -> descrição), dia = data serial com formato "dd"
# (o dia do mês), resolvida pelo dia da semana dentro do período; marcações = os hh:mm do texto da
# coluna D e o resto vira a descrição; horas G..O arredondadas ao minuto; colunas de auditoria
# (Y..AC) calculadas igual.
CAB = ["Matrícula", "Nome", "Cargo", "Setor", "DT", "Sem", "Cód. Horário", "Horário",
       "Ponto 1", "Ponto 2", "Ponto 3", "Ponto 4", "Ponto 5", "Ponto 6", "Descrição Marcação",
       "Trabalho", "BH", "Just.", "Injus.", "50%", "60%", "100%", "120%", "Ad. Not.",
       "Banco de horas e extra no mesmo dia", "Ponto faltando", "Mais de 4 pontos", "Tolerância < 15min", "Batidas < 30min"]
BASE = datetime.datetime(1899, 12, 30)
SEM = {"SEG": 0, "TER": 1, "QUA": 2, "QUI": 3, "SEX": 4, "SAB": 5, "SÁB": 5, "DOM": 6}

def txt(v):
    if v is None: return ''
    if isinstance(v, float) and v == int(v): return str(int(v))
    return str(v).replace('\xa0', ' ').strip()

def resolver_data(dia, sem, ini, fim):
    alvo = SEM.get(sem.upper().strip())
    if alvo is None: return None
    achou = []
    d = ini - datetime.timedelta(days=1)
    while d <= fim + datetime.timedelta(days=1):
        if d.day == dia and d.weekday() == alvo: achou.append(d)
        d += datetime.timedelta(days=1)
    return achou[0] if len(achou) == 1 else None

def formatar(caminho):
    L = ler_planilha(caminho)
    ini = fim = None
    for r in sorted(L)[:20]:
        datas = [v for c, v in sorted(L[r].items()) if isinstance(v, float) and 36526 < v < 73415]
        if len(datas) >= 2:
            a, b = sorted(datas[:2]); ini, fim = BASE + datetime.timedelta(days=a), BASE + datetime.timedelta(days=b); break
    out = []; mat = nome = cargo = setor = ''; horarios = {}; lendo = False; captura = False
    for r in sorted(L):
        c = L[r]
        A = txt(c.get(1)).upper(); B = txt(c.get(2)).upper()
        if 'EMPREGADO:' in A:
            mat, nome, setor, horarios, lendo = txt(c.get(2)), txt(c.get(4)), '', {}, False
        if 'CARGO:' in A: cargo = txt(c.get(2))
        if 'LOCALIZAÇÃO:' in A or 'LOCALIZACAO:' in A: setor = txt(c.get(4))
        if 'HORÁRIOS:' in A or 'HORARIOS:' in A: lendo = True
        elif A or not B: lendo = False
        if lendo and B and B not in horarios: horarios[B] = txt(c.get(3))
        if A == 'DT': captura, lendo = True, False; continue
        if 'HORAS NORMAIS:' in A: captura = False
        if not (captura and A and B and isinstance(c.get(1), float)): continue
        dia = (BASE + datetime.timedelta(days=c[1])).day
        dt = resolver_data(dia, B, ini, fim)
        cod = txt(c.get(3)).upper()
        texto = txt(c.get(4))
        horas = re.findall(r'(?:0[0-9]|1[0-9]|2[0-3]):[0-5][0-9]', texto)
        desc = texto
        for h in horas: desc = desc.replace(h, '')
        desc = desc.strip()
        pontos = [(int(h[:2]) * 60 + int(h[3:])) / 1440 for h in horas[:6]] + [None] * (6 - min(6, len(horas)))
        vals = []
        for j in range(7, 16):
            v = c.get(j)
            vals.append(round(v * 1440) / 1440 if isinstance(v, float) else (txt(v) or None))
        bh, extra50 = vals[1], vals[4]
        tem_extra = any(v not in (None, '') for v in vals[4:8])
        banco = 'Sim' if bh not in (None, '') and tem_extra else ''
        p = ['' if x is None else x for x in pontos]
        falta = 'Falta ponto' if ((p[0] == '' and p[1] == '' and p[2] != '' and p[3] != '') or (p[0] != '' and p[1] != '' and p[2] == '' and p[3] == '')) else ''
        mais4 = 'Sim' if sum(1 for x in pontos if x is not None) > 4 else ''
        q = 15 / 1440
        em = isinstance(extra50, float) and 0 < extra50 < q; fm = isinstance(bh, float) and 0 < bh < q
        tol = 'Extra e Falta < 15min' if em and fm else 'Extra < 15min' if em else 'Falta < 15min' if fm else ''
        prox = ''
        for j in range(5):
            if pontos[j] is not None and pontos[j + 1] is not None and round(abs(pontos[j + 1] - pontos[j]) * 1440) <= 30:
                f = lambda x: '%02d:%02d' % divmod(round(x * 1440), 60)
                prox = 'Sim (%s e %s)' % (f(pontos[j]), f(pontos[j + 1])); break
        m = int(mat) if mat.isdigit() else mat
        out.append([m, nome, cargo, setor, dt, txt(c.get(2)), cod, horarios.get(cod, ''), *pontos, desc, *vals, banco, falta, mais4, tol, prox])
    return {'ini': ini, 'fim': fim, 'linhas': out}

if __name__ == '__main__':
    d = formatar(sys.argv[1])
    print(len(d['linhas']), 'linhas;', d['ini'].date(), 'a', d['fim'].date(), '; sem data:', sum(1 for l in d['linhas'] if l[4] is None))
    if len(sys.argv) > 2: pickle.dump(d, open(sys.argv[2], 'wb'))
