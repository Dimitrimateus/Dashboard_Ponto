"""Reproduz em Python o FormatarBancoHoras.bas (04/11): lê o relatório de banco de horas do
sistema (HRES105, o arquivo "BH" da pasta do mês) e devolve a mesma tabela que a macro deixa
na aba "Banco de Horas". Serve para conferir sem Excel (CLAUDE.md, seção 10).

Uso: python3 ferramentas/formatar_bh.py "BH.xlsx" [saida.pkl]
O .pkl é um dict {ini, fim, meses, linhas} (linhas na ordem de CAB, horas em MINUTOS).
"""
import datetime, pickle, sys

from formatar_marcacoes import ler_planilha, simples, texto

CAB = ["Matrícula", "Colaborador", "Filial", "Local", "Início do período", "Fim do período",
       "Meses do período", "Banco fechado", "Saldo anterior positivo", "Saldo anterior negativo",
       "Crédito", "Débito", "Horas anistiadas", "Horas pagas", "Saldo final positivo", "Saldo final negativo"]
BASE = datetime.date(1899, 12, 30)


def minutos(v):
    """Hora do relatório -> minutos, sempre positivo ("25:40-" = 1540; 1.0701389 = 1541)."""
    if v is None:
        return 0
    if isinstance(v, float):
        return int(round(abs(v) * 1440))
    t = str(v).replace("-", "").replace(" ", "").strip()
    if not t:
        return 0
    if ":" not in t:
        try:
            return int(round(abs(float(t)) * 1440))
        except ValueError:
            return 0
    h, m = t.split(":", 1)
    return int(h or 0) * 60 + int(m or 0)


def meses_do_periodo(ini, fim):
    """Ciclos de 28 a 27 cobertos: 28/06 a 27/10 = 4."""
    a = ini.year * 12 + ini.month + (1 if ini.day >= 28 else 0)
    b = fim.year * 12 + fim.month + (1 if fim.day >= 28 else 0)
    return b - a + 1


def formatar(linhas):
    # colunas pela linha "Locais | Positivo | Negativo | Credito | Débito | Hr. Anist. | Hr. Pagas | Positivo | Negativo"
    pos = dict(antpos=5, antneg=6, cred=7, deb=8, anist=9, pagas=10, finpos=11, finneg=12)
    for r in sorted(linhas):
        cel = linhas[r]
        if simples(cel.get(1)) == "LOCAIS":
            npos = nneg = 0
            for c in sorted(cel):
                t = simples(cel[c])
                if t == "POSITIVO":
                    npos += 1; pos["antpos" if npos == 1 else "finpos"] = c
                elif t == "NEGATIVO":
                    nneg += 1; pos["antneg" if nneg == 1 else "finneg"] = c
                elif t == "CREDITO": pos["cred"] = c
                elif t == "DEBITO": pos["deb"] = c
                elif t.startswith("HR. ANIST") or t.startswith("HR ANIST"): pos["anist"] = c
                elif t.startswith("HR. PAGAS") or t.startswith("HR PAGAS"): pos["pagas"] = c
            break
    ini = fim = None; filial = local = ""; saida = []
    for r in sorted(linhas):
        cel = linhas[r]
        for c in sorted(cel):
            t = simples(cel[c])
            if t == "PERIODO:":
                for k in sorted(cel):
                    v = cel[k]
                    if k > c and isinstance(v, float) and 30000 < v < 80000:
                        ini = v if ini is None or v < ini else ini
                        fim = v if fim is None or v > fim else fim
            elif t == "FILIAL:":
                for k in sorted(cel):
                    if k > c and texto(cel[k]):
                        filial = texto(cel[k]); break
        a = cel.get(1); ta = texto(a)
        if simples(a).startswith("AURORA,"):
            local = ta.rsplit(",", 1)[-1].strip()
        elif ta.upper() == "AURORA" or ta.upper().startswith("AURORA "):
            local = ta
        elif ta and ta.replace(".", "", 1).isdigit() and texto(cel.get(2)):
            saida.append([str(int(float(ta))), texto(cel.get(2)), filial, local, None, None, None, None] +
                         [minutos(cel.get(pos[k])) for k in ("antpos", "antneg", "cred", "deb", "anist", "pagas", "finpos", "finneg")])
    di = BASE + datetime.timedelta(days=int(ini)); df = BASE + datetime.timedelta(days=int(fim))
    meses = meses_do_periodo(di, df)
    for l in saida:
        l[4], l[5], l[6], l[7] = di, df, meses, "Sim" if meses >= 4 else "Não"
    return dict(ini=di, fim=df, meses=meses, linhas=saida)


if __name__ == "__main__":
    res = formatar(ler_planilha(sys.argv[1]))
    L = res["linhas"]
    print("Período %s a %s (%d meses, %s); %d pessoas" % (res["ini"].strftime("%d/%m/%Y"), res["fim"].strftime("%d/%m/%Y"),
          res["meses"], "fechado" if res["meses"] >= 4 else "aberto", len(L)))
    tot = lambda i: sum(l[i] for l in L)
    h = lambda m: "%d:%02d" % divmod(m, 60)
    print("Crédito %s · Débito %s · Anistiadas %s (%d pessoas) · Pagas %s (%d pessoas)" % (
        h(tot(10)), h(tot(11)), h(tot(12)), sum(1 for l in L if l[12]), h(tot(13)), sum(1 for l in L if l[13])))
    erradas = [l for l in L if (l[8] - l[9] + l[10] - l[11]) - (l[13] - l[12]) != (l[14] - l[15])]
    print("Pessoas em que anterior + crédito − débito − pagas + anistiadas ≠ saldo final:", len(erradas))
    if len(sys.argv) > 2:
        pickle.dump(res, open(sys.argv[2], "wb"))
