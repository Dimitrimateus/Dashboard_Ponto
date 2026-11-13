# -*- coding: utf-8 -*-
"""
CONFERÊNCIA AMPLIADA DO PAINEL (13/11/2026)
================================================================================
Complemento de ferramentas/auditar_csv_painel.py. Lá ficam as contas principais de cada
seção; aqui, TODO o resto que aparece na tela: as tabelas de tendência (por gestor e por
pessoa, com a frase "comparado ao mês anterior"), os quadros de afastamentos que faltavam
(evolução por tipo, dia da semana, encostados em folga, mapa, sinais de atestado), banco de
horas (saldo dia a dia, fechamento, tabela por pessoa inteira), pausas (evolução, dias sem
pausa), PontoNet (lista, demora por pessoa, tendência), as listas completas, "Arquivos
carregados", o resumo de cada pessoa (a janela do nome) e as visões com filtro de gestor.

Mesma ideia da conferência principal: nada aqui usa o código do painel. Cada conta é refeita
do zero, em Python, a partir do CSV, seguindo as regras do RH, e comparada com o que a tela
mostrou (lido por extrair_tudo_painel.js).

É chamado no fim de auditar_csv_painel.conferir(J, R) por conferir_mais(J, R, A), onde A é o
próprio módulo auditar_csv_painel (para reaproveitar a leitura, item(), tabela() etc.).
"""
import math, re
from collections import Counter, defaultdict, OrderedDict
from datetime import date, datetime, timedelta

A = None   # o módulo auditar_csv_painel (preenchido por conferir_mais)


# ------------------------------------------------------------------ formatação igual à da tela
def jsround(x):
    """Arredonda como o navegador (Math.round): 2,5 -> 3; -2,5 -> -2."""
    return math.floor(x + 0.5)


def fmt_num(n):
    """1234 -> '1.234' (número inteiro com ponto de milhar)."""
    n = jsround(n)
    s = "{:,}".format(abs(n)).replace(",", ".")
    return ("-" if n < 0 else "") + s


def fmt_dec(x):
    """Uma casa decimal, vírgula e ponto de milhar: 1234,56 -> '1.234,6'."""
    v = jsround((x or 0) * 10) / 10
    inteiro = int(abs(v)); dec = int(round((abs(v) - inteiro) * 10))
    return ("-" if v < 0 else "") + fmt_num(inteiro) + "," + str(dec)


def to_fixed1(x):
    """Uma casa decimal como o toFixed(1) do navegador (arredonda pelo valor exato do número)."""
    from decimal import Decimal, ROUND_HALF_UP
    return str(Decimal(x).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)).replace(".", ",")


def espacos(t): return re.sub(r"\s+", " ", t or "").strip()


def pct_txt(v, total):
    if not total: return "0%"
    p = v * 100 / total
    return "<1%" if 0 < p < 1 else "%d%%" % jsround(p)


def dias_txt(v):
    if v is None: return "–"
    return fmt_dec(v).replace(",0", "", 1) + (" dia" if abs(v - 1) < 0.05 else " dias")


def horas_txt(m):
    """Minutos -> '2h 05min' / '45min' / '3h' (sem sinal)."""
    m = jsround(abs(m or 0)); h, mi = divmod(m, 60)
    if h == 0: return "%dmin" % mi
    if mi == 0: return "%dh" % h
    return "%dh %02dmin" % (h, mi)


def saldo_txt(m):
    m = jsround(m or 0)
    if m == 0: return "0h"
    return ("+" if m > 0 else "−") + horas_txt(m)


DIAS = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]


def dia_curto(d): return DIAS[d.weekday()]


def ddmm(d): return d.strftime("%d/%m")


# ------------------------------------------------------------------ regras comuns
def chave_nome(t): return re.sub(r"\s+", " ", A.sem_acento(t or "")).upper().strip()


def chave_pessoa(r): return "M:" + r["matricula"] if r["matricula"] else "N:" + chave_nome(r["colaborador"])


def dias_do(r):
    """Os dias corridos que um afastamento cobre: do início, 'dias de ausência' dias (mínimo 1)."""
    n = max(1, jsround(r["dias_aus"] or 1))
    return [r["dia"] + timedelta(days=i) for i in range(n)]


def eh_atestado(r):
    return r["tipo"] == A.TIPO_ABS or (r["tipo"] == A.TIPO_AFAST and "atest" in A.sem_acento(r["situacao"]).lower())


def por_horas(r): return r["tipo"] == A.TIPO_AFAST and r["duracao"] > 0 and not eh_atestado(r)


GRUPOS = [("Atestado", ["atest"]), ("Férias", ["ferias"]), ("Viagem a serviço", ["viagem"]),
          ("Reunião, curso e treinamento", ["reuni", "curso", "trein", "integra"]),
          ("Saída médica e horas", ["medico", "horas falta"]), ("Acidente, doença e falta", ["acidente", "doenca", "falta"]),
          ("Doação de sangue", ["sangue"])]


def grupo_afast(r):
    if eh_atestado(r): return "Atestado"
    t = A.sem_acento(r["situacao"]).lower()
    for nome, tem in GRUPOS[1:]:
        if any(x in t for x in tem): return nome
    return "Outros (licenças, dispensas…)"


def fim_cartao(R):
    d = [r["dia"] for r in R if r["tipo"] in (A.TIPO_RES_HORAS, A.TIPO_AUD)]
    return max(d) if d else None


def soma_ciclo(k, n):
    y, m = map(int, k.split("-")); m += n
    while m > 12: m -= 12; y += 1
    while m < 1: m += 12; y -= 1
    return "%d-%02d" % (y, m)


def ciclos_tend(R):
    f = fim_cartao(A.R_FULL)
    if not f: return []
    u = A.ciclo(f)
    return [soma_ciclo(u, -2), soma_ciclo(u, -1), u]


def num_tela(t):
    t = str(t).strip()
    if t in ("–", "", "-"): return None
    m = re.match(r"^-?[\d.]+", t.replace("−", "-"))
    return int(m.group(0).replace(".", "")) if m else None


def tab(J, id_, cheia=True):
    return A.tabela(J, id_, cheia)


def el(J, id_):
    """Um elemento da tela, se estiver visível (quadro escondido não conta)."""
    return A.visivel(J["normal"].get(id_))


def ok_item(secao, card, oque, como, esp, pain, ok=None, obs=""):
    A.item(secao, card, oque, como, esp, pain, ok, obs)


def compara_linhas(secao, card, oque, como, esp, pain, unidade="linhas", obs=""):
    """esp e pain: dict chave -> tupla de textos. Mostra só as diferenças quando há."""
    if esp == pain:
        ok_item(secao, card, oque, como, "%d %s iguais" % (len(esp), unidade), "%d %s iguais" % (len(pain), unidade), True, obs)
    else:
        dif = {k: (esp.get(k), pain.get(k)) for k in set(esp) | set(pain) if esp.get(k) != pain.get(k)}
        ok_item(secao, card, oque, como, {k: v[0] for k, v in list(dif.items())[:8]},
                {k: v[1] for k, v in list(dif.items())[:8]}, False, obs + " (%d diferença(s))" % len(dif))


# ------------------------------------------------------------------ tendências (10 tabelas)
def indicadores(R):
    """Cada assunto de tendência: lista de (registro, ciclo, num, den)."""
    dig = [r for r in R if r["tipo"] == A.TIPO_DIG]
    res = [r for r in R if r["tipo"] == A.TIPO_RES_PAU]
    pn = [r for r in R if r["tipo"] == A.TIPO_PN and A.norm(r["situacao"]) == "integrado"]
    af = [r for r in R if r["tipo"] == A.TIPO_AFAST] or [r for r in R if r["tipo"] == A.TIPO_ABS]
    tp = lambda t: [(r, A.ciclo(r["dia"]), 1, 0) for r in R if r["tipo"] == t]
    out = OrderedDict()
    out["dig"] = ("Batidas digitadas", "sobe", "", [(r, A.ciclo(r["dia"]), 1, 0) for r in dig])
    out["esq"] = ("Batidas digitadas por esquecimento", "sobe", "",
                  [(r, A.ciclo(r["dia"]), 1, 0) for r in dig if A.norm(r["situacao"]) == "inclusao_por_esquecimento"])
    out["atest"] = ("Dias de atestado", "sobe", "", [(r, A.ciclo(d), 1, 0) for r in af if eh_atestado(r) for d in dias_do(r)])
    out["he"] = ("Horas excedentes (dias com mais de 10h)", "sobe", "", tp("Horas Excedentes"))
    out["intra"] = ("Intrajornada (intervalo de almoço curto)", "sobe", "", tp("Intrajornada"))
    out["folga"] = ("Interjornada semanal (trabalho na folga ou feriado)", "sobe", "", tp("Interjornada Semanal"))
    out["inter"] = ("Interjornada (menos de 11h de descanso)", "sobe", "", tp("Interjornada"))
    out["jor"] = ("Ocorrências de jornada (os 4 tipos)", "sobe", "",
                  [(r, A.ciclo(r["dia"]), 1, 0) for r in R if r["tipo"] in A.TIPOS_JORNADA])
    out["pauCertas"] = ("Pausas certas (20 min)", "desce", "pct", [(r, A.ciclo(r["dia"]), r["pc"], r["pc"] + r["pm"] + r["pM"]) for r in res])
    out["pauAcima"] = ("Trabalho acima de 1h40 sem pausa", "sobe", "pct", [(r, A.ciclo(r["dia"]), r["tM"], r["tc"] + r["tM"] + r["tm"]) for r in res])
    out["pnAprov"] = ("Justificativas aprovadas no PontoNet", "sobe", "", [(r, A.ciclo(r["dia"]), 1, 0) for r in pn])
    out["pnGestor"] = ("Dias para o gestor avaliar (média)", "sobe", "media",
                       [(r, A.ciclo(r["dia"]), r["dias_gest"], 1) for r in pn if r["dias_gest"] is not None])
    out["pauImpar"] = ("Batida de pausa faltando", "sobe", "", [(r, A.ciclo(r["dia"]), r["imp"], 0) for r in res])
    return out


def agrega(itens, chave, ciclos):
    m = OrderedDict()
    for r, c, num, den in itens:
        if c not in ciclos: continue
        k = chave(r)
        a = m.setdefault(k, {"label": r["colaborador"], "gestor": r["gestor"], "c": {}})
        v = a["c"].setdefault(c, [0, 0]); v[0] += num or 0; v[1] += den or 0
    return m


def valor(tipo, a, c):
    v = a["c"].get(c)
    if tipo in ("pct", "media"):
        if not v or not v[1]: return None
        return v[0] * 100 / v[1] if tipo == "pct" else v[0] / v[1]
    return v[0] if v else 0


def fmt_valor(tipo, v):
    if v is None: return "–"
    if tipo == "pct": return fmt_dec(v).replace(",0", "", 1) + "%"
    if tipo == "media": return dias_txt(v)
    return fmt_num(v)


def mudanca(tipo, pior, a, ciclos):
    """A frase 'comparado ao mês anterior' e se piorou/melhorou."""
    u, p = valor(tipo, a, ciclos[-1]), valor(tipo, a, ciclos[-2])
    if u is None or p is None: return "–", 0, False, False
    d = u - p
    if abs(d) < (0.05 if tipo else 0.5): return "igual", 0, False, False
    piorou = d > 0 if pior == "sobe" else d < 0
    seta = "▲ " if d > 0 else "▼ "
    if tipo:
        fp = (lambda v: dias_txt(v)) if tipo == "media" else (lambda v: fmt_dec(v).replace(",0", "", 1) + "%")
        txt = seta + ("subiu" if d > 0 else "caiu") + " de " + fp(p) + " para " + fp(u)
    else:
        txt = seta + fmt_num(abs(d)) + (" a mais" if d > 0 else " a menos") + " que em " + A.rotulo_ciclo(ciclos[-2]) + \
              ((" (" + ("+" if d > 0 else "−") + str(jsround(abs(d) * 100 / p)) + "%)") if p > 0 else "")
    return txt, d, piorou, not piorou


def linhas_tend(ind, por, ciclos):
    nome, pior, tipo, itens = ind
    chave = (lambda r: r["gestor"]) if por == "gestor" else chave_pessoa
    m = agrega(itens, chave, ciclos)
    if por == "gestor":
        for k, a in m.items(): a["label"] = k
    return m


def confere_tabela_tend(secao, card, J, id_, ind, por, ciclos, filtro=None, limite_vazio=False):
    """Uma tabela de tendência: para cada linha, o valor de cada mês e a frase de comparação."""
    nome, pior, tipo, itens = ind
    m = linhas_tend(ind, por, ciclos)
    linhas = list(m.values())
    if filtro:
        linhas = filtro(linhas)
    t = tab(J, id_)
    if not t["rows"] and not linhas: return
    n_c = len(ciclos)
    pain, esp = {}, {}
    for row in t["rows"]:
        if len(row) < 2 + n_c: continue
        base = 2 if por == "pessoa" else 1
        k = row[0] + ("|" + row[1] if por == "pessoa" else "")
        pain[k] = tuple(row[base:base + n_c]) + (row[base + n_c + 1],)
    for a in linhas:
        k = a["label"] + ("|" + a["gestor"] if por == "pessoa" else "")
        if k not in pain: continue   # lista cortada na tela (os primeiros); a falta de alguém é vista abaixo
        esp[k] = tuple(fmt_valor(tipo, valor(tipo, a, c)) for c in ciclos) + (mudanca(tipo, pior, a, ciclos)[0],)
    faltando = len(linhas) - len(pain)
    compara_linhas(secao, card, "Valores de cada mês e a frase \"comparado ao mês anterior\" (%s)" % nome,
                   "Assunto \"%s\", somado por %s em cada mês do ponto (%s). A frase compara o último mês com o anterior."
                   % (nome, "gestor" if por == "gestor" else "pessoa", ", ".join(A.rotulo_ciclo(c) for c in ciclos)),
                   esp, pain, "linhas")
    ok_item(secao, card, "Quantidade de linhas (%s)" % nome,
            "Uma linha por %s que teve o assunto em algum dos 3 meses." % ("gestor" if por == "gestor" else "pessoa"),
            len(linhas), len(t["rows"]) if not faltando or faltando < 0 else len(t["rows"]),
            len(linhas) == len(t["rows"]))


def tendencias(J, R):
    ciclos = ciclos_tend(R)
    if len(ciclos) < 2: return
    inds = indicadores(R)
    tabelas = [("Marcações digitadas", "Tendência por pessoa", "tableDigTend", "dig", "pessoa"),
               ("Pausas térmicas", "Tendência por gestor", "tablePauTendGestor", "pauAcima", "gestor"),
               ("Pausas térmicas", "Tendência por pessoa", "tablePauTend", "pauAcima", "pessoa"),
               ("Afastamentos", "Tendência por gestor", "tableAbsTendGestor", "atest", "gestor"),
               ("Jornada", "Tendência por gestor", "tableJorTendGestor", "jor", "gestor"),
               ("Jornada", "Tendência por pessoa", "tableJorTendPessoa", "he", "pessoa"),
               ("PontoNet", "Tendência por gestor", "tablePnTendGestor", "pnGestor", "gestor")]
    for secao, card, id_, ind_id, por in tabelas:
        if not tab(J, id_)["rows"]: continue
        confere_tabela_tend(secao, card, J, id_, inds[ind_id], por, ciclos)
    # seção Tendências: assunto do seletor (o primeiro da lista, no padrão)
    nome_sel = J["estado"].get("tendIndicador") or ""
    ind_id = next((k for k, v in inds.items() if v[0] == nome_sel), "dig")
    ind = inds[ind_id]
    confere_tabela_tend("Tendências", "Por gestor (assunto escolhido)", J, "tableTendGestor", ind, "gestor", ciclos)

    def variacoes(quer):
        def f(linhas):
            out = []
            for a in linhas:
                txt, d, piorou, melhorou = mudanca(ind[2], ind[1], a, ciclos)
                if not (piorou if quer == "piorou" else melhorou): continue
                if ind[2]:
                    u, p = a["c"].get(ciclos[-1]), a["c"].get(ciclos[-2])
                    if not (u and p and u[1] >= 10 and p[1] >= 10): continue
                out.append(a)
            return out
        return f
    confere_tabela_tend("Tendências", "Quem mais piorou", J, "tableTendPiorou", ind, "pessoa", ciclos, filtro=variacoes("piorou"))
    confere_tabela_tend("Tendências", "Quem mais melhorou", J, "tableTendMelhorou", ind, "pessoa", ciclos, filtro=variacoes("melhorou"))


# ------------------------------------------------------------------ afastamentos (o que faltava)
def afastamentos(J, R):
    S = "Afastamentos"
    todos = [r for r in R if r["tipo"] == A.TIPO_AFAST] or [r for r in R if r["tipo"] == A.TIPO_ABS]
    if not todos: return
    est = J["estado"]
    mostrar = est.get("absMostrar") or "Todos os afastamentos"
    so_atest = mostrar == "Só atestados"
    so_grupo = mostrar[3:] if mostrar.startswith("Só ") and not so_atest else None
    passa = lambda r: (eh_atestado(r) if so_atest else (grupo_afast(r).lower() == so_grupo.lower() if so_grupo else True))
    todos = [r for r in todos if passa(r)] if (so_atest or so_grupo) else todos
    vis = [r for r in todos if not por_horas(r)]
    # meses do seletor: do arquivo inteiro (sem filtro de pessoa), até o fim do cartão ponto
    full = [r for r in A.R_FULL if r["tipo"] == A.TIPO_AFAST] or [r for r in A.R_FULL if r["tipo"] == A.TIPO_ABS]
    f = fim_cartao(A.R_FULL)
    cmax = A.ciclo(f) if f else None
    ciclos = sorted({A.ciclo(d) for r in full if passa(r) and not por_horas(r) for d in dias_do(r) if not cmax or A.ciclo(d) <= cmax}, reverse=True)
    rot_sel = (est.get("absCiclo") or "").split(" ")[0]
    ciclo_sel = next((c for c in ciclos if A.rotulo_ciclo(c) == rot_sel), None)
    no_ciclo = (lambda d: A.ciclo(d) == ciclo_sel) if ciclo_sel else (lambda d: A.ciclo(d) in ciclos)

    # resumo em frases
    vis_c = [r for r in vis if any(no_ciclo(d) for d in dias_do(r))]
    dias_c = sum(1 for r in vis_c for d in dias_do(r) if no_ciclo(d))
    pessoas = {chave_pessoa(r) for r in vis_c}
    at = [r for r in vis_c if eh_atestado(r)]
    dias_at = sum(1 for r in at for d in dias_do(r) if no_ciclo(d))
    ph = [r for r in todos if por_horas(r) and no_ciclo(r["dia"])]
    txt = (el(J, "absResumo") or {}).get("texto", "")
    nums = lambda: [int(x.replace(".", "")) for x in re.findall(r"(?<![\d/])\d[\d.]*(?![\d/])", txt.split("RESUMO")[-1].split(":", 1)[-1])]
    oque = "atestado(s) médico(s)" if so_atest else ("afastamento(s) do tipo “%s”" % so_grupo.lower() if so_grupo else "afastamento(s) registrado(s)")
    esp = "%s %s, de %s pessoa(s) diferente(s), somando %s dia(s)" % (fmt_num(len(vis_c)), oque, fmt_num(len(pessoas)), fmt_num(dias_c))
    ok_item(S, "Resumo em frases", "Afastamentos, pessoas e dias do período",
            "Afastamentos com pelo menos um dia no período; pessoas diferentes; total de dias (do início ao fim de cada afastamento).",
            esp, (re.search(r"\d[\d.]* (?:afastamento|atestado).*?somando [\d.]+ dia\(s\)", txt) or [""])[0], esp in txt)
    if not so_atest and not so_grupo:
        esp2 = "Desses, %s foram atestados médicos, com %s dia(s): %d%%" % (fmt_num(len(at)), fmt_num(dias_at), jsround(dias_at * 100 / dias_c) if dias_c else 0)
        ok_item(S, "Resumo em frases", "Atestados, dias de atestado e % dos dias",
                "Dos afastamentos do período, os que são atestado médico, seus dias e a parte que representam do total de dias.",
                esp2, (re.search(r"Desses, .*?%", txt) or [""])[0], esp2 in txt)
    if ph:
        esp3 = "%s saída(s) de poucas horas no mesmo dia (consulta médica, horas de falta, reunião curta), somando %s" % (fmt_num(len(ph)), horas_txt(sum(r["duracao"] for r in ph)))
        ok_item(S, "Resumo em frases", "Saídas de poucas horas", "Afastamentos de poucas horas no mesmo dia, em quantidade e horas.",
                esp3, (re.search(r"\d[\d.]* saída.*?somando [^.]*", txt) or [""])[0], esp3 in txt)

    # evolução (gráfico): dias por mês
    ped = [(d, grupo_afast(r), r) for r in vis for d in dias_do(r)]
    if not ciclo_sel:
        esp = Counter(A.rotulo_ciclo(A.ciclo(d)) for d, g, r in ped if A.ciclo(d) in ciclos)
        A.compara_dict(S, "Gráfico de evolução", "Dias de afastamento em cada mês",
                       "Cada dia de afastamento contado no mês do ponto em que cai.", {k: v for k, v in esp.items() if v},
                       {k: v for k, v in A.tips(J, "chartAbsEvolucao").items() if v})
        # pedaços por tipo (legenda)
        leg = {}
        for x in (el(J, "chartAbsEvolucao") or {}).get("legendaMini", []):
            m = re.match(r"^(.*) ([\d.]+)$", x)
            if m: leg[m.group(1)] = int(m.group(2).replace(".", ""))
        A.compara_dict(S, "Gráfico de evolução", "Dias de cada tipo (legenda)", "Dias somados por tipo de afastamento nos meses mostrados.",
                       dict(Counter(g for d, g, r in ped if A.ciclo(d) in ciclos)), leg)
    # dia da semana do início dos atestados
    todos_af = [r for r in R if r["tipo"] == A.TIPO_AFAST] or [r for r in R if r["tipo"] == A.TIPO_ABS]
    abs_ = [r for r in todos_af if eh_atestado(r)]
    ini = [r for r in abs_ if no_ciclo(r["dia"])]
    A.compara_dict(S, "Atestados por dia da semana", "Atestados que começaram em cada dia",
                   "Atestados contados pelo dia da semana em que começaram, nos meses mostrados.",
                   dict(Counter(dia_curto(r["dia"]) for r in ini)), {k: v for k, v in A.tips(J, "chartAbsDiaSemana").items() if v})
    em = [r for r in ini if r["emenda"]]
    t = (el(J, "absEmendaResumo") or {}).get("texto", "")
    esp = "%d de %d atestados (%d%%) encostam numa folga" % (len(em), len(ini), jsround(len(em) * 100 / len(ini))) if ini else "Nenhum atestado"
    ok_item(S, "Atestados encostados em folga", "Resumo", "Atestados que começam logo depois ou terminam logo antes de uma folga, sobre os que começaram no período.",
            esp, t, esp in t)
    pain = {}
    for row in tab(J, "tableAbsEmendas")["rows"]:
        pain.setdefault(row[0] + "|" + row[2], 0); pain[row[0] + "|" + row[2]] += 1
    esp = Counter()
    for r in em:
        dd = dias_do(r); fim = dd[-1]
        txt_at = dia_curto(r["dia"]) + " " + ddmm(r["dia"]) + ((" a " + dia_curto(fim) + " " + ddmm(fim)) if len(dd) > 1 else "") + \
                 " (%d %s)" % (len(dd), "dias" if len(dd) > 1 else "dia")
        esp[r["colaborador"] + "|" + txt_at] += 1
    compara_linhas(S, "Atestados encostados em folga", "Lista: pessoa, datas e dias de cada atestado",
                   "Cada atestado que encosta numa folga, com o dia da semana e a data de começo e de fim.", dict(esp), pain, "atestados")
    # mapa: dias no ciclo e em 3 meses por pessoa
    cmapa = ciclo_sel or (ciclos[0] if ciclos else None)
    pp = {}
    for r in vis:
        k = chave_pessoa(r)
        x = pp.setdefault(k, {"nome": r["colaborador"], "dias": set(), "tot": 0})
        x["tot"] += sum(1 for d in dias_do(r) if A.ciclo(d) in ciclos)
        for d in dias_do(r):
            if A.ciclo(d) == cmapa: x["dias"].add(d)
    esp = {x["nome"]: (len(x["dias"]), x["tot"]) for x in pp.values() if x["tot"] > 0}
    pain = {row[0]: (num_tela(row[-2]) or 0, num_tela(row[-1]) or 0) for row in tab(J, "absHeat")["rows"]}
    compara_linhas(S, "Mapa colaborador × dia", "Dias no mês e dias nos 3 meses, por pessoa",
                   "Para cada pessoa com afastamento nos meses mostrados: os dias no mês do mapa e o total nos 3 meses.", esp, pain, "pessoas")
    cores = (tab(J, "absHeat").get("cores") or [])
    ok_item(S, "Mapa colaborador × dia", "Quadrados coloridos = dias no mês",
            "Em cada linha do mapa, a quantidade de quadrados pintados tem que ser igual aos dias no mês.",
            sum(v[0] for v in pain.values()), sum(sum(c.values()) for c in cores))
    # sinais: mesmo dia da semana (2+), 1 dia na seg/sex, perto das férias
    pd_ = defaultdict(list)
    for r in abs_: pd_[(chave_pessoa(r), r["dia"].weekday())].append(r)
    tot = Counter(chave_pessoa(r) for r in abs_)
    esp = {}
    for (k, wd), l in pd_.items():
        if len(l) >= 2:
            l = sorted(l, key=lambda r: r["dia"])
            esp[l[0]["colaborador"] + "|" + DIAS[wd]] = ("%d de %d" % (len(l), tot[k]), ", ".join(dia_curto(r["dia"]) + " " + ddmm(r["dia"]) for r in l))
    pain = {row[0] + "|" + row[1]: (row[2], row[3]) for row in tab(J, "tableAbsMesmoDia")["rows"] if len(row) >= 4}
    compara_linhas(S, "Sinais de atenção", "2 ou mais atestados no mesmo dia da semana",
                   "Pessoas com 2+ atestados começando no mesmo dia da semana; quantos e as datas.", esp, pain, "pessoas")
    ss = lambda r: r["dia"].weekday() in (0, 4)
    um = lambda r: jsround(r["dias_aus"] or 1) <= 1
    n_ss = sum(1 for r in abs_ if ss(r)); um_ss = [r for r in abs_ if ss(r) and um(r)]
    t = (el(J, "absSegSexResumo") or {}).get("texto", "")
    esp = ("%s de %s atestados (%s) começam na segunda ou na sexta; %s deles são de 1 dia só" % (fmt_num(n_ss), fmt_num(len(abs_)), pct_txt(n_ss, len(abs_)), fmt_num(len(um_ss)))
           if abs_ else "Nenhum atestado nos últimos 3 meses para o filtro atual.")
    ok_item(S, "Sinais de atenção", "Atestados na segunda ou sexta (resumo)", "Atestados que começam na segunda ou sexta, e quantos deles são de 1 dia.", esp, t, esp in t)
    p2 = defaultdict(list)
    for r in um_ss: p2[(chave_pessoa(r), r["colaborador"])].append(r)
    esp = {k[1]: (str(len(l)), ", ".join(dia_curto(r["dia"]) + " " + ddmm(r["dia"]) for r in sorted(l, key=lambda r: r["dia"]))) for k, l in p2.items()}
    pain = {row[0]: (row[1], row[2]) for row in tab(J, "tableAbsSegSex")["rows"] if len(row) >= 3}
    compara_linhas(S, "Sinais de atenção", "Atestado de 1 dia na segunda ou sexta, por pessoa", "Pessoas, quantos e as datas.", esp, pain, "pessoas")
    g2 = defaultdict(lambda: [0, 0])
    for r in abs_:
        g2[r["gestor"]][0] += 1
        if ss(r) and um(r): g2[r["gestor"]][1] += 1
    esp = {g: (str(v[1]), str(v[0]), pct_txt(v[1], v[0])) for g, v in g2.items() if v[1] > 0}
    pain = {row[0]: (row[1], row[2], row[3]) for row in tab(J, "tableAbsSegSexGestor")["rows"] if len(row) >= 4}
    compara_linhas(S, "Sinais de atenção", "Atestado de 1 dia na segunda ou sexta, por equipe", "Por gestor: atestados de 1 dia na seg/sex, total de atestados e %.", esp, pain, "equipes")
    fer = defaultdict(list)
    for r in todos_af:
        if "ferias" in A.sem_acento(r["situacao"]).lower(): fer[chave_pessoa(r)].append(r)
    casos = Counter()
    for a in abs_:
        da = dias_do(a); ai, af = da[0], da[-1]
        for fr in fer.get(chave_pessoa(a), []):
            df = dias_do(fr); fi, ff = df[0], df[-1]
            antes = (fi - af).days - 1; depois = (ai - ff).days - 1; q = None
            if ai <= ff and fi <= af: q = "durante as férias"
            elif 0 <= antes <= 3: q = "termina na véspera das férias" if antes == 0 else "termina %d dia(s) antes das férias" % antes
            elif 0 <= depois <= 3: q = "começa no dia seguinte ao fim das férias" if depois == 0 else "começa %d dia(s) depois das férias" % depois
            if q: casos[a["colaborador"] + "|" + q] += 1
    pain = Counter()
    for row in tab(J, "tableAbsFerias")["rows"]:
        if len(row) >= 5: pain[row[0] + "|" + row[4]] += 1
    compara_linhas(S, "Sinais de atenção", "Atestado perto das férias", "Atestados a até 3 dias das férias da mesma pessoa (antes, depois ou durante).",
                   dict(casos), dict(pain), "casos")


def conferir_mais(J, R, modulo):
    global A
    A = modulo
    tendencias(J, R)
    afastamentos(J, R)
    banco_horas(J, R)
    jornada(J, R)
    pontonet(J, R)
    digitadas(J, R)
    pausas(J, R)
    arquivos_e_lista(J, R)
    selos(J, R)


# ------------------------------------------------------------------ banco de horas (tudo)
def num_ciclo(d): return d.year * 12 + d.month + (1 if d.day >= 28 else 0)


def fim_ciclo(n): return date((n - 1) // 12, (n - 1) % 12 + 1, 27)


def ini_ciclo(n):
    f = fim_ciclo(n - 1); return date(f.year, f.month, 28)


def ini_periodo(n):
    """1º mês do período do banco (nov–fev, mar–jun, jul–out) do ciclo n."""
    m = (n - 1) % 12 + 1
    desloc = m - 11 if m >= 11 else (m + 1 if m <= 2 else (m - 3 if m <= 6 else m - 7))
    return n - desloc


def eh_fechamento(n): return (n - 1) % 12 + 1 in (2, 6, 10)


def dados_bh(R):
    """Pessoas com movimento: dias (extras = coluna 50%, falta = coluna BH) e relatórios do sistema."""
    pes = OrderedDict(); ciclos = set(); dmin = dmax = None
    for r in R:
        dia = r["tipo"] == A.TIPO_BH_DIA; rel = r["tipo"] == A.TIPO_BH
        if not dia and not rel: continue
        k = chave_pessoa(r)
        p = pes.setdefault(k, {"nome": r["colaborador"], "gestor": r["gestor"], "dias": {}, "rels": [], "tem_dia": False})
        if dia:
            x = p["dias"].setdefault(r["dia"], [0, 0, r["situacao"]]); x[0] += r["he"]; x[1] += r["bh"]
            ciclos.add(num_ciclo(r["dia"])); p["gestor"] = r["gestor"]; p["tem_dia"] = True
            dmin = min(dmin or r["dia"], r["dia"]); dmax = max(dmax or r["dia"], r["dia"])
        else:
            p["rels"].append(r)
            if not p["tem_dia"]: p["gestor"] = r["gestor"]
    cl = sorted(ciclos)
    if cl: dmin = min(dmin, ini_ciclo(cl[0])); dmax = max(dmax, fim_ciclo(cl[-1]))
    return pes, cl, dmin, dmax


def saldos(p, dias, cl, dmin, dmax, R):
    """Saldo do banco no fim de cada dia. Período com o cartão ponto de todos os meses: parte de zero
    no começo do período. Faltando mês: parte do relatório do sistema que termina no período."""
    out = {}
    por_q = defaultdict(list)
    for t in dias: por_q[ini_periodo(num_ciclo(t))].append(t)
    rel_fins = [num_ciclo(r["dia"]) for r in R if r["tipo"] == A.TIPO_BH]
    for q, ts in por_q.items():
        qi, qf = ini_ciclo(q), fim_ciclo(q + 3)
        ult = max([n for n in cl if q <= n <= q + 3] + [n for n in rel_fins if q <= n <= q + 3] + [q])
        falta = [n for n in range(q, min(ult, q + 3) + 1) if n not in cl]
        anc = None
        if not falta: anc = (qi - timedelta(days=1), 0)
        else:
            for r in p["rels"]:
                if not r["bh_ini"]: continue
                tf, ti = r["dia"], r["bh_ini"].date()
                if ti < qi or tf > qf: continue
                if tf < dmin - timedelta(days=1) or tf > dmax + timedelta(days=1): continue
                if not anc or tf > anc[0]: anc = (tf, r["bh_ant"] + r["bh_cred"] - r["bh_deb"])
            if not anc: anc = (max(qi, dmin) - timedelta(days=1), 0)
        movs = {t: v for t, v in p["dias"].items() if qi <= t <= qf}
        for t in ts:
            v = anc[1]
            for mt, x in movs.items():
                dl = x[0] - x[1]
                if anc[0] < mt <= t: v += dl
                elif t < mt <= anc[0]: v -= dl
            out[t] = v
    return out


def banco_horas(J, R):
    S = "Banco de horas"
    pes, _cl, _a, _b = dados_bh(R)
    _p, cl, dmin, dmax = dados_bh(A.R_FULL)
    if not cl: return
    marcados = [m.split(" ")[0] for m, v in J["estado"].get("bhMeses", []) if v]
    escolhidos = [n for n in cl if "%s/%s" % (A.MESES[(n - 1) % 12], str((n - 1) // 12)[2:]) in marcados]
    if not escolhidos: return
    dias = []
    for n in escolhidos:
        t = ini_ciclo(n)
        while t <= fim_ciclo(n):
            if dmin <= t <= dmax: dias.append(t)
            t += timedelta(days=1)
    a, b = escolhidos[0], escolhidos[-1]
    comeco, fech = a == ini_periodo(a), eh_fechamento(b)
    linhas, soma_dia = {}, defaultdict(lambda: [0, 0, 0])
    tot_ini = tot_fim = 0
    for k, p in pes.items():
        s = saldos(p, dias, cl, dmin, dmax, R)
        d0 = dias[0]; m0 = p["dias"].get(d0)
        v_ini = s.get(d0, 0) - ((m0[0] - m0[1]) if m0 else 0)
        v_fim = s.get(dias[-1], 0)
        tot_ini += v_ini; tot_fim += v_fim
        c = d = dc = dd = folgas = 0; menor = None
        for t in dias:
            v = s.get(t)
            if v is not None and (menor is None or v < menor): menor = v
            sd = soma_dia[t]; sd[2] += v or 0
            m = p["dias"].get(t)
            if not m: continue
            c += m[0]; d += m[1]; dc += m[0] > 0; dd += m[1] > 0
            sd[0] += m[0]; sd[1] += m[1]
            if re.search(r"falta \(banco horas\)", A.sem_acento(m[2] or ""), re.I): folgas += 1
        if not (c or d or v_ini or v_fim): continue
        leit = ("Começou zerado" if comeco else "Começou " + texto_saldo(v_ini)) + \
               ((" e fechou devendo " + horas_txt(v_fim) if v_fim < 0 else (" e fechou com " + horas_txt(v_fim) + " no banco" if v_fim > 0 else " e fechou zerado"))
                if fech else " e em " + fim_ciclo(b).strftime("%d/%m/%Y") + " estava " + texto_saldo(v_fim))
        linhas[p["nome"] + "|" + p["gestor"]] = (saldo_txt(v_ini), horas_txt(c), horas_txt(d), saldo_txt(v_fim),
                                                 "–" if menor is None else saldo_txt(menor), str(dc), str(dd), str(folgas), leit)
    pain = {}
    for row in tab(J, "tableBHPessoas")["rows"]:
        if len(row) >= 11: pain[row[0] + "|" + row[1]] = tuple(row[2:11])
    compara_linhas(S, "Banco de cada pessoa (tabela inteira)",
                   "Banco no começo e no fim, horas extras, horas falta, pior momento, dias com extra e com falta, folgas e a frase",
                   "Saldo de cada pessoa no fim de cada dia, somando dia a dia (horas extras − horas falta) desde o começo do período "
                   "(todo mundo zerado). Pior momento = o menor saldo de um fim de dia. Folga no banco = dia inteiro descontado do banco.",
                   linhas, pain, "pessoas")
    # gráfico do saldo, dia a dia (soma de todos)
    esp, pn = {}, {}
    for t in dias:
        sd = soma_dia[t]
        esp[t.strftime("%d/%m/%Y")] = (horas_txt(sd[0]), horas_txt(sd[1]), saldo_txt(sd[2]))
    for tp in (el(J, "chartBHLinha") or {}).get("tips", []):
        m = re.match(r"^\w+ (\d\d/\d\d/\d{4}) · Horas extras: (.*?) · Horas falta: (.*?) · Banco no fim do dia: (\S+(?: \S+min)?)", tp)
        if m: pn[m.group(1)] = (m.group(2), m.group(3), m.group(4))
    compara_linhas(S, "Saldo do banco dia a dia (gráfico)", "Horas extras, horas falta e banco no fim de cada dia (soma de todos)",
                   "Para cada dia dos meses escolhidos: horas extras e horas falta do dia e o banco de todos somado no fim do dia.",
                   esp, pn, "dias")
    kp = A.numeros(J, "bhKpis")
    ok_item(S, "Números do topo", "Banco no começo e no fim (soma de todos)",
            "Soma do banco de cada pessoa no começo e no fim dos meses escolhidos.",
            (saldo_txt(tot_ini), saldo_txt(tot_fim)), (kp.get("BANCO NO COMEÇO", ("",))[0], kp.get("BANCO NO FIM", ("",))[0]))
    # quem colocou e tirou horas (barras por pessoa)
    esp = {}
    for p in pes.values():
        c = sum(m[0] for t, m in p["dias"].items() if t in set(dias)); d = sum(m[1] for t, m in p["dias"].items() if t in set(dias))
        if c or d: esp[p["nome"]] = (horas_txt(c), horas_txt(d))
    pn = {}
    for tp in ((A.visivel(J["telaCheia"].get("chartBancoHoras")) or el(J, "chartBancoHoras")).get("tips") or []):
        m = re.match(r"^(.*?) · Horas extras: (.*?) · Horas falta: (.*)$", tp)
        if m: pn[m.group(1)] = (m.group(2), m.group(3))
    compara_linhas(S, "Horas extras e horas falta por pessoa (gráfico)", "Horas extras e horas falta de cada pessoa",
                   "Soma, nos meses escolhidos, das horas extras e das horas falta do cartão ponto de cada pessoa.", esp, pn, "pessoas")
    # fechamento (relatório do sistema)
    rels = defaultdict(list)
    for r in R:
        if r["tipo"] == A.TIPO_BH: rels[((r["bh_ini"].date() if r["bh_ini"] else None), r["dia"])].append(r)
    if not rels: return
    chave = max(rels, key=lambda k: (k[1], -(k[0].toordinal() if k[0] else 0)))
    lst = rels[chave]
    na, nb = (num_ciclo(chave[0]) if chave[0] else num_ciclo(chave[1])), num_ciclo(chave[1])
    fechado = any(eh_fechamento(n) for n in range(na, nb + 1))
    sit = lambda r: ("pago" if r["bh_pag"] > 0 else "descontado" if r["bh_ani"] > 0 else "zerado") if fechado else \
                    ("positivo" if r["bh_fin"] > 0 else "negativo" if r["bh_fin"] < 0 else "zerado")
    txt = (el(J, "bhResumo") or {}).get("texto", "")
    c = Counter(sit(r) for r in lst)
    esp1 = "%s pessoa(s) no relatório · horas extras %s, horas falta %s" % (fmt_num(len(lst)), horas_txt(sum(abs(r["bh_cred"]) for r in lst)), horas_txt(sum(abs(r["bh_deb"]) for r in lst)))
    ok_item(S, "Como cada pessoa fechou o banco (resumo)", "Pessoas, horas extras e horas falta do relatório",
            "Soma das colunas do relatório de banco de horas do sistema.", esp1, (re.search(r"\d[\d.]* pessoa\(s\) no relatório[^.]*", txt) or [""])[0], esp1 in txt)
    if fechado:
        esp2 = "%s tiveram horas pagas no fechamento (%s no total). %s fecharam devendo. %s fecharam zerados" % (
            fmt_num(c["pago"]), horas_txt(sum(r["bh_pag"] for r in lst if sit(r) == "pago")), fmt_num(c["descontado"]), fmt_num(c["zerado"]))
        pain2 = re.sub(r"\s+", " ", txt)
        ok_item(S, "Como cada pessoa fechou o banco (resumo)", "Quantos tiveram horas pagas, fecharam devendo e fecharam zerados",
                "Pela coluna de horas pagas e de horas descontadas de cada pessoa no relatório do sistema.", esp2,
                (re.search(r"\d[\d.]* tiveram horas pagas.*?zerados", pain2) or [""])[0], esp2 in pain2)
    esp = {}
    for r in lst:
        s_ = sit(r)
        fimtxt = ((horas_txt(r["bh_pag"]) + " pagas") if s_ == "pago" else ("fechou devendo " + horas_txt(r["bh_ani"])) if s_ == "descontado" else "zerado") if fechado \
            else texto_saldo(r["bh_fin"])
        esp[r["colaborador"]] = (horas_txt(r["bh_cred"]), horas_txt(r["bh_deb"]), fimtxt)
    pn = {}
    for tp in ((A.visivel(J["telaCheia"].get("chartFechamentoBH")) or el(J, "chartFechamentoBH")).get("tips") or []):
        m = re.match(r"^(.*?) · Horas extras: (.*?) · Horas falta: (.*?)(?: · Saldo que trazia de antes: .*?)? · (?:Fechamento|No fim do relatório): (.*)$", tp)
        if m: pn[m.group(1)] = (m.group(2), m.group(3), m.group(4))
    compara_linhas(S, "Como cada pessoa fechou o banco (gráfico)", "Horas extras, horas falta e o fechamento de cada pessoa",
                   "Do relatório do sistema: horas extras, horas falta e, no fechamento, horas pagas, horas que ficou devendo ou zerado.",
                   esp, pn, "pessoas")


def texto_saldo(v):
    return ("com " + horas_txt(v) + " no banco") if v > 0 else (("devendo " + horas_txt(v)) if v < 0 else "zerado")


# ------------------------------------------------------------------ jornada, PontoNet, digitadas, pausas (o que faltava)
NOME_JOR = {"Horas Excedentes": "Horas excedentes", "Intrajornada": "Intrajornada", "Interjornada Semanal": "Interjornada semanal", "Interjornada": "Interjornada"}


def detalhe_jornada(r):
    if r["tipo"] == "Horas Excedentes" and r["min_trab"] is not None:
        return horas_txt(r["min_trab"]) + " no dia (" + horas_txt(max(0, r["min_trab"] - 600)) + " acima de 10h)"
    if r["tipo"] == "Interjornada" and r["min_desc"] is not None:
        return "Descansou " + horas_txt(r["min_desc"]) + " (faltaram " + horas_txt(max(0, 660 - r["min_desc"])) + " para 11h)"
    if r["tipo"] == "Interjornada Semanal": return "Trabalhou no dia de folga ou feriado"
    if r["tipo"] == "Intrajornada" and r["min_trab"] is not None:
        return (r["situacao"] or "Intervalo curto") + " · " + horas_txt(r["min_trab"]) + " de trabalho no dia"
    return r["situacao"] or ""


def ciclo_sel(J, chave, todos_ciclos):
    rot = (J["estado"].get(chave) or "").split(" ")[0]
    return next((c for c in todos_ciclos if A.rotulo_ciclo(c) == rot), None)


def equipe_org(R):
    vis, eq = set(), Counter()
    for r in R:
        if r["tipo"] != A.TIPO_ORG: continue
        k = chave_pessoa(r)
        if k in vis: continue
        vis.add(k); eq[r["gestor"]] += 1
    return eq


def jornada(J, R):
    S = "Jornada"
    todos = [r for r in R if r["tipo"] in A.TIPOS_JORNADA]
    if not todos: return
    cs = sorted({A.ciclo(r["dia"]) for r in A.R_FULL if r["tipo"] in A.TIPOS_JORNADA})
    c = ciclo_sel(J, "jorCiclo", cs)
    rec = [r for r in todos if not c or A.ciclo(r["dia"]) == c]
    kp = A.numeros(J, "jorKpis")
    sem = sum(1 for r in rec if r["tipo"] == "Intrajornada" and re.search("sem intervalo", r["situacao"], re.I))
    ok_item(S, "Números do topo", "Intrajornada: dias sem intervalo nenhum", "Dias de intrajornada em que o sistema diz \"sem intervalo\".",
            "%d sem intervalo nenhum" % sem, kp.get("INTRAJORNADA", ("", ""))[1].split("· ")[-1])
    eq = equipe_org(R); g = defaultdict(lambda: [0, set()])
    for r in rec: g[r["gestor"]][0] += 1; g[r["gestor"]][1].add(chave_pessoa(r))
    esp = {k: ("%d%% de %d" % (jsround(len(v[1]) * 100 / eq[k]), eq[k])) if eq.get(k) else "fora da lista" for k, v in g.items()}
    pain = {row[0]: row[7] for row in tab(J, "tableJorGestor")["rows"] if len(row) >= 8}
    compara_linhas(S, "Tabela por gestor", "% da equipe", "Pessoas da equipe com pelo menos uma ocorrência ÷ pessoas da equipe na lista de equipes.", esp, pain, "gestores")
    esp = Counter((r["dia"].strftime("%d/%m/%Y"), r["colaborador"], r["gestor"], NOME_JOR[r["tipo"]], espacos(detalhe_jornada(r)), espacos(r["marc"])) for r in rec)
    pain = Counter(tuple(row[:6]) for row in tab(J, "tableJorLista")["rows"] if len(row) >= 6)
    compara_linhas(S, "Lista de ocorrências", "Cada linha da lista: data, pessoa, gestor, tipo, o que houve e batidas",
                   "Uma linha por ocorrência dos relatórios de jornada, com o texto do que houve calculado das horas do dia.",
                   {"|".join(k): v for k, v in esp.items()}, {"|".join(k): v for k, v in pain.items()}, "linhas")
    leg = {}
    for x in (el(J, "jorListaLeg") or {}).get("legendaMini", []):
        m = re.match(r"^(.*) ([\d.]+)$", x)
        if m: leg[m.group(1)] = int(m.group(2).replace(".", ""))
    A.compara_dict(S, "Lista de ocorrências", "Quantidade de cada tipo (legenda)", "Ocorrências da lista contadas por tipo.",
                   dict(Counter(NOME_JOR[r["tipo"]] for r in rec)), leg)


def pontonet(J, R):
    S = "PontoNet"
    ap = [r for r in R if r["tipo"] == A.TIPO_PN and A.norm(r["situacao"]) == "integrado"]
    if not ap: return
    cs = sorted({A.ciclo(r["dia"]) for r in A.R_FULL if r["tipo"] == A.TIPO_PN})
    c = None if (J["estado"].get("pnCiclo") or "").startswith("Todos") else (ciclo_sel(J, "pnCiclo", cs) or cs[-1])
    rec = [r for r in ap if not c or A.ciclo(r["dia"]) == c]
    media = lambda vs: (sum(vs) / len(vs)) if vs else None
    kp = A.numeros(J, "pnKpis")
    ok_item(S, "Números do topo", "Tempo do colaborador e do gestor (média)",
            "Média dos dias contados pelo PontoNet, nos pedidos aprovados do mês: do colaborador (pedido aberto → justificativa) e do gestor (justificativa → avaliação).",
            (dias_txt(media([r["dias_col"] for r in rec if r["dias_col"] is not None])), dias_txt(media([r["dias_gest"] for r in rec if r["dias_gest"] is not None]))),
            (kp.get("TEMPO DO COLABORADOR PARA JUSTIFICAR", ("",))[0], kp.get("TEMPO DO GESTOR PARA AVALIAR", ("",))[0]))
    # mês a mês por motivo (pedaços)
    tot_c = Counter(A.ciclo(r["dia"]) for r in ap)
    esp = {}
    for r in ap:
        k = A.rotulo_ciclo(A.ciclo(r["dia"])) + " · " + A.nome_tipo_pn(r["tipo_just"])
        esp[k] = esp.get(k, 0) + 1
    esp = {k: "%d (%s)" % (v, pct_txt(v, tot_c[next(c_ for c_ in tot_c if A.rotulo_ciclo(c_) == k.split(" · ")[0])])) for k, v in esp.items()}
    pain = {}
    for sgm in (el(J, "chartPnCiclo") or {}).get("segs", []):
        m = re.match(r"^(.*?): (\d+ \([^)]*\))", sgm)
        if m: pain[m.group(1)] = m.group(2)
    compara_linhas(S, "Mês a mês", "Aprovados de cada motivo em cada mês (pedaços das barras)", "Pedidos aprovados contados por mês e motivo, com a % da barra do mês.", esp, pain, "pedaços")
    # lista completa
    dn = lambda v: "–" if v is None else fmt_num(v)
    esp = Counter((r["data"].strftime("%d/%m/%Y") + (" " + r["data"].strftime("%H:%M") if r["data"].strftime("%H:%M") != "00:00" else ""),
                   r["colaborador"], r["gestor"], A.nome_tipo_pn(r["tipo_just"]), espacos(r["just"]), dn(r["dias_col"]), dn(r["dias_gest"])) for r in rec)
    pain = Counter(tuple(row[:7]) for row in tab(J, "tablePnLista")["rows"] if len(row) >= 7)
    compara_linhas(S, "Lista de aprovados", "Cada linha: batida que faltou, pessoa, gestor, motivo, texto, dias do colaborador e do gestor",
                   "Uma linha por pedido aprovado do mês, com os campos do PontoNet.", {"|".join(k): v for k, v in esp.items()},
                   {"|".join(k): v for k, v in pain.items()}, "linhas")
    leg = {}
    for x in (el(J, "pnListaLeg") or {}).get("legendaMini", []):
        m = re.match(r"^(.*) ([\d.]+)$", x)
        if m: leg[m.group(1)] = int(m.group(2).replace(".", ""))
    A.compara_dict(S, "Lista de aprovados", "Quantidade de cada motivo (legenda)", "Pedidos da lista contados por motivo.",
                   dict(Counter(A.nome_tipo_pn(r["tipo_just"]) for r in rec)), leg)
    # demora: média de todos e por pessoa
    dem = [(r, (r["tratativa"] - r["data"]).total_seconds() / 3600) for r in rec if r["tratativa"]]
    dem = [(r, h) for r, h in dem if h >= 0]
    fd = lambda h: (to_fixed1(h) + "h") if h < 24 else (to_fixed1(h / 24) + " dias")
    t = (el(J, "demoraGeral") or {}).get("texto", "")
    esp = "Média de todos: %s (%s justificativa(s) aprovada(s)" % (fd(sum(h for _, h in dem) / len(dem)), fmt_num(len(dem))) if dem else "Média de todos: –"
    ok_item(S, "Demora até a aprovação", "Média de todos", "Média, em todos os pedidos aprovados do mês, do dia e hora da batida que faltou até a aprovação.", esp, t, t.startswith(esp))
    por = defaultdict(list)
    for r, h in dem: por[r["colaborador"]].append(h)
    esp = {k: (str(len(v)), fd(sum(v) / len(v)), fd(max(v))) for k, v in por.items()}
    pain = {row[0].split(" ⚠")[0]: tuple(row[1:4]) for row in tab(J, "tableDemoraColaborador")["rows"] if len(row) >= 4}
    compara_linhas(S, "Demora por colaborador", "Aprovadas, demora média e maior demora de cada pessoa", "Mesma conta da demora, por pessoa.", esp, pain, "pessoas")


def digitadas(J, R):
    S = "Marcações digitadas"
    dig = [r for r in R if r["tipo"] == A.TIPO_DIG]
    if not dig: return
    cs = sorted({A.ciclo(r["dia"]) for r in A.R_FULL if r["tipo"] == A.TIPO_DIG})
    c = ciclo_sel(J, "digCiclo", cs)
    rec = [r for r in dig if not c or A.ciclo(r["dia"]) == c]
    kp = A.numeros(J, "digKpis")
    esq = [r for r in rec if A.norm(r["situacao"]) == "inclusao_por_esquecimento"]
    org_k = {chave_pessoa(r) for r in R if r["tipo"] == A.TIPO_ORG}
    sem_mat = {chave_nome(r["colaborador"]) for r in R if r["tipo"] == A.TIPO_ORG and not r["matricula"]}
    eq = len(org_k | {chave_pessoa(r) for r in rec if chave_pessoa(r) not in org_k and chave_nome(r["colaborador"]) not in sem_mat})
    pess = len({chave_pessoa(r) for r in rec})
    sub = kp.get("COLABORADORES", ("", ""))[1]
    ok_item(S, "Números do topo", "% de pessoas com batida digitada", "Pessoas com batida digitada ÷ pessoas da lista de equipes.",
            "%s de %s" % (pct_txt(pess, eq), fmt_num(eq)), " ".join(sub.split(" ")[:3]), sub.startswith("%s de %s" % (pct_txt(pess, eq), fmt_num(eq))))
    ok_item(S, "Números do topo", "% das digitadas que são por esquecimento", "Batidas por esquecimento ÷ total de batidas digitadas.",
            "%s das digitadas" % pct_txt(len(esq), len(rec)), kp.get("POR ESQUECIMENTO", ("", ""))[1])
    esp = Counter((r["data"].strftime("%d/%m/%Y"), r["data"].strftime("%H:%M"), r["colaborador"], r["gestor"], espacos(r["situacao"]), espacos(r["just"])) for r in rec)
    pain = Counter(tuple(row[:6]) for row in tab(J, "tableDigLista")["rows"] if len(row) >= 6)
    compara_linhas(S, "Lista de batidas", "Cada linha: data, hora, pessoa, gestor, motivo e justificativa",
                   "Uma linha por batida digitada, com os campos do relatório.", {"|".join(k): v for k, v in esp.items()},
                   {"|".join(k): v for k, v in pain.items()}, "linhas")
    # esquecimento repetido: frase de comparação e quem entra
    ct = sorted({A.ciclo(r["dia"]) for r in A.R_FULL if r["tipo"] == A.TIPO_DIG})[-3:]   # os 3 últimos meses com digitadas
    por = defaultdict(lambda: {"c": {}})
    for r in dig:
        if A.norm(r["situacao"]) != "inclusao_por_esquecimento" or A.ciclo(r["dia"]) not in ct: continue
        a = por[chave_pessoa(r)]; a.setdefault("label", r["colaborador"]); a.setdefault("gestor", r["gestor"])
        v = a["c"].setdefault(A.ciclo(r["dia"]), [0, 0]); v[0] += 1
    esp = {a["label"] + "|" + a["gestor"]: tuple(str(valor("", a, k)) for k in ct) + (mudanca("", "sobe", a, ct)[0],) for a in por.values()}
    pain = {row[0] + "|" + row[1]: tuple(row[2:2 + len(ct)]) + (row[2 + len(ct)],) for row in tab(J, "tableDigEsquecimento")["rows"] if len(row) >= 3 + len(ct)}
    compara_linhas(S, "Esquecimento repetido", "Esquecimentos de cada mês e a frase \"comparado ao mês anterior\" (lista inteira)",
                   "Batidas por esquecimento de cada pessoa em cada um dos 3 meses e a comparação do último com o anterior.", esp, pain, "pessoas")
    n5 = sum(1 for a in por.values() if valor("", a, ct[-1]) >= 5)
    t = (el(J, "digEsqHint") or {}).get("texto", "")
    ok_item(S, "Esquecimento repetido", "Pessoas com 5 ou mais esquecimentos no último mês", "Quem teve 5 ou mais batidas por esquecimento no último mês (em vermelho).",
            "%d pessoa(s) com 5 ou mais" % n5, t, ("%d pessoa(s) com 5 ou mais" % n5) in t)


def pausas(J, R):
    S = "Pausas térmicas"
    res = [r for r in R if r["tipo"] == A.TIPO_RES_PAU]
    if res:
        ult = max(r["dia"] for r in res)
        mes = [r for r in res if r["dia"] == ult]
        kp = A.numeros(J, "pauKpis")
        blocos = lambda r: r["tc"] + r["tM"] + r["tm"]
        acima = sum(r["tM"] for r in mes); bl = sum(blocos(r) for r in mes)
        a30 = sum(1 for r in mes if blocos(r) and r["tM"] / blocos(r) > 0.30)
        ok_item(S, "Números do topo", "Trabalho acima de 1h40: % dos trechos e pessoas acima de 30%",
                "Trechos acima de 1h40 ÷ todos os trechos do mês; pessoas com mais de 30% dos trechos acima de 1h40.",
                (fmt_num(acima), "%s dos trechos · %d pessoa(s) acima de 30%%" % (pct_txt(acima, bl), a30)), kp.get("TRABALHO ACIMA DE 1H40", ("", "")))
        imp = sum(r["imp"] for r in mes); com = sum(1 for r in mes if r["imp"] > 0)
        ok_item(S, "Números do topo", "Batida de pausa faltando: total e pessoas", "Soma das batidas de pausa faltando; pessoas com pelo menos uma.",
                (fmt_num(imp), "%d pessoa(s) esqueceram de bater a saída ou a volta da pausa" % com), kp.get("BATIDA DE PAUSA FALTANDO", ("", "")))
        por_mes = defaultdict(list)
        for r in res: por_mes[r["dia"]].append(r)
        if len(por_mes) >= 2:
            esp = {A.rotulo_ciclo(A.ciclo(d)): fmt_num(sum(r["tM"] for r in l) * 100 / max(1, sum(blocos(r) for r in l))) for d, l in por_mes.items()}
            pain = {k: str(v) for k, v in A.tips(J, "chartPauEvolucao").items()}
            compara_linhas(S, "Evolução mês a mês (gráfico)", "Trabalho acima de 1h40 em % dos trechos, em cada mês",
                           "Trechos acima de 1h40 ÷ todos os trechos, somando as pessoas de cada mês.", esp, pain, "meses")
    dias = [r for r in R if r["tipo"] == A.TIPO_PAU_DIA]
    if not dias: return
    cs = sorted({A.ciclo(r["dia"]) for r in A.R_FULL if r["tipo"] == A.TIPO_PAU_DIA})
    c = ciclo_sel(J, "pauBatCiclo", cs) or cs[-1]
    rec = [r for r in dias if A.ciclo(r["dia"]) == c]
    an = {id(r): A.analisa_batidas(r["marc"]) for r in rec}
    sem = [r for r in rec if an[id(r)]["sem"]]
    t = (el(J, "pauSemPausaHint") or {}).get("texto", "")
    esp = "%s dia(s) de %s (%s) com mais de 1h40 de trabalho e nenhuma pausa térmica" % (fmt_num(len(sem)), fmt_num(len(rec)), pct_txt(len(sem), len(rec)))
    ok_item(S, "Batidas: dias sem nenhuma pausa", "Resumo", "Dias com mais de 1h40 de trabalho e nenhuma pausa térmica (duas batidas F seguidas).", esp, t, t.startswith(esp))
    nd = Counter(chave_pessoa(r) for r in rec); ns = defaultdict(list)
    for r in sem: ns[chave_pessoa(r)].append(r)
    esp = {l[0]["colaborador"]: "%d de %d" % (len(l), nd[k]) for k, l in ns.items()}
    pain = {}
    for tp in ((A.visivel(J["telaCheia"].get("chartPauSemPausa")) or el(J, "chartPauSemPausa")).get("tips") or []):
        m = re.match(r"^(.*?): (\d+ de \d+) dias", tp)
        if m: pain[m.group(1)] = m.group(2)
    compara_linhas(S, "Batidas: dias sem nenhuma pausa", "Por pessoa: dias sem pausa / dias trabalhados", "Por pessoa, no mês das batidas.", esp, pain, "pessoas")
    prim = sorted(an[id(r)]["primeira"] for r in rec if an[id(r)]["primeira"] is not None)
    if prim:
        med = prim[(len(prim) - 1) // 2]; tard = sum(1 for x in prim if x > 100)
        hm = lambda m: "%dh%02d" % divmod(jsround(m), 60)
        esp = "Metade dos dias teve a 1ª pausa com até %s de trabalho · %s de %s dias (%s) passaram de 1h40" % (hm(med), fmt_num(tard), fmt_num(len(prim)), pct_txt(tard, len(prim)))
        t = (el(J, "pauPrimeiraHint") or {}).get("texto", "")
        ok_item(S, "Batidas: trabalho até a 1ª pausa térmica", "Frase do resumo (metade dos dias e quantos passaram de 1h40)",
                "Mediana do tempo de trabalho até a 1ª pausa e quantos dias passaram de 1h40.", esp, t)


def arquivos_e_lista(J, R):
    S = "Lista de colaboradores"
    org = {}
    for r in R:
        if r["tipo"] == A.TIPO_ORG and chave_pessoa(r) not in org: org[chave_pessoa(r)] = r
    t = (el(J, "orgResumo") or {}).get("texto", "")
    esp = "%d na lista de equipes %d sem matrícula no cadastro %d sem gestor %d sem cargo %d sem área" % (
        len(org), sum(1 for r in org.values() if not r["matricula"]), sum(1 for r in org.values() if r["gestor"] in ("Sem gestor", "Desligados")),
        sum(1 for r in org.values() if r["cargo"] == "Sem cargo"), sum(1 for r in org.values() if not r["situacao"]))
    ok_item(S, "Lista de colaboradores", "Resumo do cadastro (sem matrícula, sem gestor, sem cargo, sem área)", "Contagem direta na lista de equipes.", esp, t)
    # Arquivos carregados: linhas e período de cada informação
    S = "Arquivos carregados"
    fontes = OrderedDict([("Cartão ponto", lambda t_: t_ in (A.TIPO_BH_DIA, A.TIPO_RES_HORAS, A.TIPO_AUD) or t_ in A.TIPOS_CURTAS),
                          ("Banco de horas", lambda t_: t_ == A.TIPO_BH), ("Pausas térmicas", lambda t_: t_ in (A.TIPO_RES_PAU, A.TIPO_PAU_DIA)),
                          ("Marcações digitadas", lambda t_: t_ == A.TIPO_DIG), ("Afastamentos", lambda t_: t_ == A.TIPO_AFAST),
                          ("PontoNet", lambda t_: t_ == A.TIPO_PN), ("Horas excedentes", lambda t_: t_ == "Horas Excedentes"),
                          ("Intrajornada", lambda t_: t_ == "Intrajornada"), ("Interjornada", lambda t_: t_ == "Interjornada"),
                          ("Interjornada semanal", lambda t_: t_ == "Interjornada Semanal"), ("Lista de equipes", lambda t_: t_ == A.TIPO_ORG)])
    esp = {}
    for nome, f in fontes.items():
        l = [r for r in A.R_FULL if f(r["tipo"])]   # a seção mostra o arquivo inteiro, sem filtro
        if not l: continue
        if nome == "Lista de equipes": per = "%d pessoa(s) (não tem período)" % len({chave_pessoa(r) for r in l})
        elif nome == "Banco de horas":   # o relatório de banco de horas tem o seu período (início e fim)
            per = "relatório de " + min(r["bh_ini"] for r in l if r["bh_ini"]).strftime("%d/%m/%Y") + " a " + max(r["dia"] for r in l).strftime("%d/%m/%Y")
        else: per = min(r["dia"] for r in l).strftime("%d/%m/%Y") + " a " + max(r["dia"] for r in l).strftime("%d/%m/%Y")
        esp[nome] = (fmt_num(len(l)), per)
    t = tab(J, "arqPorArquivo")
    pain = {row[0]: (row[2], row[3]) for row in t["rows"] if len(row) >= 4}
    compara_linhas(S, "O que veio em cada arquivo", "Linhas de cada informação", "Quantas linhas do arquivo são de cada relatório do sistema e de que data a que data elas vão.", esp, pain, "informações")


# ------------------------------------------------------------------ etiquetas 📅 (o período de cada quadro)
def selos(J, R):
    S = "Etiquetas de período (📅)"
    sl = J.get("selos") or {}
    if not sl: return
    br = lambda d: d.strftime("%d/%m/%Y")
    def faixa(c1, c2): return br(A.ciclo_ini_fim(c1)[0]) + " a " + br(A.ciclo_ini_fim(c2)[1])
    full = A.R_FULL
    esp = {}
    f = fim_cartao(full)
    if f:
        k = A.ciclo(f)
        esp["tableCurtas"] = "Mês do cartão ponto: " + faixa(k, k)
        esp["tableExtraFalta"] = "Mês do cartão ponto: " + faixa(k, k) + " + os 2 meses anteriores"
    paus = max((r["dia"] for r in full if r["tipo"] == A.TIPO_RES_PAU), default=None)
    if paus:
        for i in ["tablePausasTermicas", "chartPauAcima", "tablePauAcimaGestor", "chartPauGrupo"]:
            esp[i] = "Relatório de pausas do mês: " + faixa(A.ciclo(paus), A.ciclo(paus))
    ct = ciclos_tend(R)
    if ct:
        t = "Ciclos " + ", ".join(A.rotulo_ciclo(c) for c in ct) + ": " + faixa(ct[0], ct[-1])
        for i in ["tableTendUnidade", "tableTendGestor", "chartDigCiclo", "tableDigTend", "tablePauTendGestor", "tablePauTend",
                  "tableAbsTendGestor", "tableJorTendGestor", "tableJorTendPessoa", "tableDigEsquecimento", "tablePnTendGestor"]:
            esp[i] = t
    def por_seletor(tipos, chave, ids_sel, ids_todos=()):
        cs = sorted({A.ciclo(r["dia"]) for r in full if r["tipo"] in tipos})
        if not cs: return
        c = A.sel_ciclo(J, chave, cs)
        todos = "Todos os ciclos do arquivo: " + faixa(cs[0], cs[-1])
        txt = ("Ciclo " + A.rotulo_ciclo(c) + ": " + faixa(c, c)) if c else todos
        for i in ids_sel: esp[i] = txt
        for i in ids_todos: esp[i] = todos
    por_seletor({A.TIPO_DIG}, "digCiclo", ["chartDigMotivo", "chartDigDiaSemana", "chartDigData", "chartDigTop", "chartDigGestor", "chartDigHora", "tableDigJust", "tableDigLista"])
    por_seletor(set(A.TIPOS_JORNADA), "jorCiclo", ["chartJorDezHoras", "chartJorGestor", "tableJorGestor", "tableJorLista"], ["chartJorCiclo"])
    cs = sorted({A.ciclo(r["dia"]) for r in full if r["tipo"] == A.TIPO_PN})
    if cs:
        c = None if (J["estado"].get("pnCiclo") or "").startswith("Todos") else (A.sel_ciclo(J, "pnCiclo", cs) or cs[-1])
        for i in ["chartPnTipo", "chartPnGestor", "tablePnLista", "tableDemoraGestor", "tableDemoraColaborador"]:
            esp[i] = ("Ciclo " + A.rotulo_ciclo(c) + ": " + faixa(c, c)) if c else ("Todos os ciclos do arquivo: " + faixa(cs[0], cs[-1]))
    cs = sorted({A.ciclo(r["dia"]) for r in full if r["tipo"] == A.TIPO_PAU_DIA})
    if cs:
        c = A.sel_ciclo(J, "pauBatCiclo", cs) or cs[-1]
        for i in ["chartPauHora", "chartPauDiaSemana", "chartPauPrimeira", "chartPauSemPausa", "chartPauMaior", "pauHeat"]:
            esp[i] = "Batidas do ciclo " + A.rotulo_ciclo(c) + ": " + faixa(c, c)
    meses = [m.split(" ")[0] for m, v in J["estado"].get("bhMeses", []) if v]
    if meses:
        for i in ["chartBHLinha", "tableBHPessoas", "chartBancoHoras"]: esp[i] = "Meses escolhidos: " + ", ".join(meses)
    pain = {k: v.replace("📅 ", "") for k, v in sl.items() if isinstance(v, str) and k in esp}
    compara_linhas(S, "Etiqueta 📅 de cada quadro", "O período escrito na etiqueta de cada quadro",
                   "O período que cada quadro considera: o mês do cartão ponto, o relatório de pausas do mês, os 3 meses das tendências, "
                   "o mês escolhido no seletor de cada seção ou os meses escolhidos do banco de horas.",
                   {k: v for k, v in esp.items() if k in pain}, pain, "quadros")
    ok_item(S, "Etiqueta 📅 de cada quadro", "Quadros com etiqueta conferidos", "Quantos quadros com etiqueta tiveram o período conferido.",
            len(pain), len([k for k in esp if k in sl]))
