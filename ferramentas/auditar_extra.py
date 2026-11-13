# -*- coding: utf-8 -*-
"""
CONFERÊNCIA DO QUE APARECE DEPOIS DE UM CLIQUE (13/11/2026)
================================================================================
Usa o JSON de ferramentas/extrair_extra_painel.js e refaz as contas, sem usar o código do painel:
  - o resumo de CADA pessoa (a janela que abre ao clicar no nome): números, tabelas e textos;
  - o painel inteiro com o filtro de CADA gestor (roda a conferência completa no recorte);
  - o painel com CADA opção dos seletores de mês/tipo das seções (idem);
  - as janelas "de onde vem": o número de linhas da janela bate com o número da barra.

Uso:
  python3 auditar_extra.py extra.json arquivo.csv [mais.csv ...] [--json resultado.json]
(junte o resultado com o de auditar_csv_painel.py para o relatório)
"""
import json, re, sys, os
from collections import Counter, defaultdict
from datetime import timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import auditar_csv_painel as A
import auditar_mais as M

M.A = A


def primeiro(todas, campo, vazio=None):
    for r in todas:
        v = r[campo]
        if v and v != vazio: return v
    return ""


def folgas_txt(r):
    """Folga(s) encostada(s) num atestado, como a janela da pessoa escreve."""
    t = r["emenda"] or ""
    out = []
    dd = M.dias_do(r); ini, fim = dd[0], dd[-1]
    mA = re.match(r"^antes \((.*?)\)(?= e depois \(|$)", t); mD = re.search(r"(?:^| e )depois \((.*)\)$", t)
    def nome_folga(x):
        n = A.sem_acento(x).lower()
        return "folga da semana" if "dsr" in n else "folga compensada" if "compens" in n else "feriado" if "feriado" in n else x
    if mA:
        d = ini - timedelta(days=1); out.append("%s %s · %s — dia antes do início" % (M.dia_curto(d), M.ddmm(d), nome_folga(mA.group(1))))
    if mD:
        d = fim + timedelta(days=1); out.append("%s %s · %s — dia depois do fim" % (M.dia_curto(d), M.ddmm(d), nome_folga(mD.group(1))))
    if not out and t: out.append(nome_folga(t))
    return " ".join(out)


def cards(X, R):
    S = "Resumo de cada pessoa"
    org_existe = any(r["tipo"] == A.TIPO_ORG for r in R)
    p_ini = A.datetime.strptime(X.get("periodoIni") or "1900-01-01", "%Y-%m-%d").date() if X.get("periodoIni") else None
    f_cart = M.fim_cartao(R)
    ci = M.ini_ciclo(M.num_ciclo(f_cart)) if f_cart else None
    # relatório de banco de horas mais recente (o do resumo)
    rels = defaultdict(list)
    for r in R:
        if r["tipo"] == A.TIPO_BH: rels[((r["bh_ini"].date() if r["bh_ini"] else None), r["dia"])].append(r)
    per = max(rels, key=lambda k: (k[1], -(k[0].toordinal() if k[0] else 0))) if rels else None
    fechado = per and any(M.eh_fechamento(n) for n in range(M.num_ciclo(per[0] or per[1]), M.num_ciclo(per[1]) + 1))
    pes, cl, dmin, dmax = M.dados_bh(R)
    res = [r for r in R if r["tipo"] == A.TIPO_RES_PAU]
    mes_pau = max((r["dia"] for r in res), default=None)
    inds = M.indicadores(R); ct = M.ciclos_tend(R)
    falhas = []; n_ok = 0; n_itens = 0
    for nome, c in X["cards"].items():
        todas = [r for r in R if r["colaborador"] == nome]
        if not todas: continue
        esp = {}
        horas = next((r for r in todas if r["tipo"] == A.TIPO_RES_HORAS), None)
        esp["HORA EXTRA"] = M.horas_txt(horas["he"]) if horas else "–"
        esp["BANCO DE HORAS"] = M.horas_txt(horas["bh"]) if horas else "–"
        esp["SALDO"] = M.saldo_txt(horas["he"] - horas["bh"]) if horas else "–"
        if per:
            bh = next((r for r in todas if r["tipo"] == A.TIPO_BH and r["dia"] == per[1]), None)
            rot = "FECHAMENTO DO BANCO DE HORAS" if fechado else "BANCO DE HORAS (RELATÓRIO)"
            if not bh: esp[rot] = "–"
            elif fechado:
                esp[rot] = ("+" + M.horas_txt(bh["bh_pag"])) if bh["bh_pag"] > 0 else (("Devendo " + M.horas_txt(bh["bh_ani"])) if bh["bh_ani"] > 0 else "Zerado")
            else: esp[rot] = M.saldo_txt(bh["bh_fin"])
        k = M.chave_pessoa(todas[0])
        p = pes.get(k) or next((x for x in pes.values() if x["nome"] == nome), None)
        if p and p["dias"] and cl:
            fim = max(dmax, M.fim_ciclo(cl[-1]))
            esp["SALDO DO BANCO NO ÚLTIMO DIA"] = M.saldo_txt(M.saldos(p, [fim], cl, dmin, dmax, R)[fim])
        aud = next((r for r in todas if r["tipo"] == A.TIPO_AUD), None)
        esp["EXTRA E FALTA NO MESMO DIA"] = ("%d / %d" % (aud["mes"], aud["tres"])) if aud else "0"
        cur = [r for r in todas if r["tipo"] in A.TIPOS_CURTAS and ci and ci <= r["dia"] <= f_cart]
        esp["OCORRÊNCIAS CURTAS"] = str(len(cur))
        pau = next((r for r in todas if r["tipo"] == A.TIPO_RES_PAU and r["dia"] == mes_pau), None)
        esp["PAUSAS TÉRMICAS"] = M.fmt_num(pau["pc"]) if pau else "–"
        # textos de baixo de alguns números
        esp_sub = {"OCORRÊNCIAS CURTAS": "%d falta(s) · %d extra(s)" % (sum(1 for r in cur if r["tipo"] == "Falta < 15min"), sum(1 for r in cur if r["tipo"] != "Falta < 15min")),
                   "PAUSAS TÉRMICAS": ("corretas · %d curta(s) · %d marcação(ões) ímpar(es)" % (pau["pm"], pau["imp"])) if pau else "fora do relatório",
                   "EXTRA E FALTA NO MESMO DIA": "este mês / 3 meses", "HORA EXTRA": "cartão ponto do mês", "BANCO DE HORAS": "cartão ponto do mês"}
        if per and fechado:
            bh = next((r for r in todas if r["tipo"] == A.TIPO_BH and r["dia"] == per[1]), None)
            if bh:
                sit = "pago" if bh["bh_pag"] > 0 else ("descontado" if bh["bh_ani"] > 0 else "zerado")
                esp_sub["FECHAMENTO DO BANCO DE HORAS"] = ("horas pagas no fechamento" if sit == "pago" else "fechou devendo" if sit == "descontado" else "fechou zerado") + \
                    " · horas extras " + M.horas_txt(bh["bh_cred"]) + ", horas falta " + M.horas_txt(bh["bh_deb"])
        pain_sub = {t[0]: t[2] for t in c["tiles"]}
        # pausas da pessoa comparadas com a média da equipe (mesmo mês e mesmo gestor, arquivo inteiro)
        pau_esp = []
        if pau:
            eq = [r for r in A.R_FULL if r["tipo"] == A.TIPO_RES_PAU and r["dia"] == pau["dia"] and r["gestor"] == pau["gestor"]]
            tot = lambda r: r["pc"] + r["pm"] + r["pM"]; blo = lambda r: r["tc"] + r["tM"] + r["tm"]
            pcc = lambda r: (r["pc"] / tot(r)) if tot(r) else 0; pac = lambda r: (r["tM"] / blo(r)) if blo(r) else 0
            med = lambda fn: (sum(fn(r) for r in eq) / len(eq)) if eq else 0
            pau_esp = [("Pausas certas (20 min)", "%s (%s de %s)" % (M.pct_txt(pcc(pau) * 100, 100), M.fmt_num(pau["pc"]), M.fmt_num(tot(pau))), M.pct_txt(med(pcc) * 100, 100)),
                       ("Pausas curtas (menos de 20 min)", M.fmt_num(pau["pm"]), M.fmt_dec(med(lambda r: r["pm"]))),
                       ("Trabalho acima de 1h40 sem pausa", "%s (%s dos trechos)" % (M.fmt_num(pau["tM"]), M.pct_txt(pau["tM"], blo(pau))), M.pct_txt(med(pac) * 100, 100) + " dos trechos"),
                       ("Batida de pausa faltando", M.fmt_num(pau["imp"]), M.fmt_dec(med(lambda r: r["imp"])))]
        pain = {t[0]: t[1] for t in c["tiles"]}
        # cabeçalho
        meta = ["Matrícula: " + primeiro(todas, "matricula")] if primeiro(todas, "matricula") else []
        meta.append("Gestor: " + (primeiro(todas, "gestor", "Sem gestor") or "Sem gestor"))
        meta.append("Cargo: " + (primeiro(todas, "cargo", "Sem cargo") or "Sem cargo"))
        setor = primeiro(todas, "setor", "Sem setor") or "Sem setor"
        if setor != "Desligado": meta.append("Setor: " + setor)
        o = next((r for r in todas if r["tipo"] == A.TIPO_ORG), None)
        if o and o["situacao"]: meta.append("Área: " + o["situacao"])
        if not o and org_existe: meta.append("Não está na lista de equipes")
        if any(r.get("desligado") for r in todas): meta.append("Desligado (não está no cadastro de ativos)")
        esp_meta = " ".join(meta)
        # tabelas: tendência, afastamentos, digitadas
        tend = {}
        for iid, ind in inds.items():
            if iid == "jor": continue
            its = [x for x in ind[3] if x[0]["colaborador"] == nome and x[1] in ct]
            if not its: continue
            if not ind[2] and not any(x[2] for x in its): continue
            a = {"c": {}}
            for r_, cc, num, den in its:
                v = a["c"].setdefault(cc, [0, 0]); v[0] += num or 0; v[1] += den or 0
            tend[ind[0]] = tuple(M.fmt_valor(ind[2], M.valor(ind[2], a, x)) for x in ct) + (M.mudanca(ind[2], ind[1], a, ct)[0],)
        sec = {s["titulo"].split(" (")[0]: s for s in c["secoes"]}
        st = sec.get("Evolução da pessoa, ciclo a ciclo")
        tend_p = {row[0]: tuple(row[1:1 + len(ct)]) + (row[-1],) for row in (st["rows"] if st else [])}
        af = sorted([r for r in todas if r["tipo"] == A.TIPO_AFAST], key=lambda r: r["data"], reverse=True)
        afs = [(r["situacao"], r["dia"].strftime("%d/%m/%Y"), M.dias_do(r)[-1].strftime("%d/%m/%Y"),
                M.horas_txt(r["duracao"]) if r["duracao"] > 0 else "%s dia(s)" % (M.fmt_num(r["dias_aus"]) if r["dias_aus"] else "1"),
                folgas_txt(r) if (M.eh_atestado(r) and r["emenda"]) else "–") for r in af]
        sa = next((s for t, s in sec.items() if t.startswith("Afastamentos")), None)
        afs_p = [tuple(row[:5]) for row in (sa["rows"] if sa else [])]
        dg = sorted([r for r in todas if r["tipo"] == A.TIPO_DIG], key=lambda r: r["data"], reverse=True)
        dgs = [(r["data"].strftime("%d/%m/%Y"), r["data"].strftime("%H:%M"), M.espacos(r["situacao"]), M.espacos(r["just"])) for r in dg]
        sd = next((s for t, s in sec.items() if t.startswith("Marcações digitadas")), None)
        dgs_p = [tuple(row[:4]) for row in (sd["rows"] if sd else [])]
        sp = next((s_ for t, s_ in sec.items() if t.startswith("Pausas térmicas no mês")), None)
        pau_p = [tuple(row[:3]) for row in (sp["rows"] if sp else [])]
        ok = True; det = []
        if pau_esp != pau_p: ok = False; det.append("pausas: %s × %s" % (pau_esp, pau_p))
        for k_, v in esp_sub.items():
            if k_ in pain_sub and pain_sub[k_] != v: ok = False; det.append("%s (texto de baixo): esperado %s, painel %s" % (k_, v, pain_sub[k_]))
        for k_, v in esp.items():
            if pain.get(k_) != v: ok = False; det.append("%s: esperado %s, painel %s" % (k_, v, pain.get(k_)))
        if c["meta"] != esp_meta: ok = False; det.append("cabeçalho: esperado [%s], painel [%s]" % (esp_meta, c["meta"]))
        if tend != tend_p: ok = False; det.append("tendência: %s × %s" % (tend, tend_p))
        if Counter(afs) != Counter(afs_p): ok = False; det.append("afastamentos diferentes")
        if Counter(dgs) != Counter(dgs_p): ok = False; det.append("digitadas diferentes")
        n_itens += 1
        if ok: n_ok += 1
        else: falhas.append(nome + ": " + "; ".join(det))
    A.item(S, "Resumo de cada pessoa (janela do nome)",
           "Todos os números (e o texto de baixo), o cabeçalho e as tabelas (evolução, pausas × média da equipe, afastamentos, digitadas) de cada pessoa",
           "Para cada pessoa do arquivo: horas do cartão ponto do mês, banco de horas (relatório e saldo no último dia), extra e "
           "falta no mesmo dia, ocorrências curtas, pausas, a evolução mês a mês, cada afastamento e cada batida digitada.",
           "%d pessoas, todas iguais" % n_itens, ("%d pessoas, todas iguais" % n_ok) if not falhas else falhas[:5],
           not falhas, obs=("%d com diferença" % len(falhas)) if falhas else "")


def resumo_rodada(secao, card, oque, como, res_antes, res_depois):
    """Junta numa linha só o resultado de uma rodada inteira de conferência (um gestor, um mês...)."""
    novos = A.RES[res_antes:]
    del A.RES[res_antes:]
    falhas = [r for r in novos if not r["ok"]]
    A.item(secao, card, oque, como, "%d contas iguais" % len(novos),
           ("%d contas iguais" % len(novos)) if not falhas else ["%s — %s: esperado %s, painel %s" % (f["card"], f["item"], f["esperado"][:150], f["painel"][:150]) for f in falhas[:6]],
           not falhas, obs=("%d de %d com diferença" % (len(falhas), len(novos))) if falhas else "")


def gestores(X, R):
    for g in X.get("gestores", []):
        nome = g["gestor"]
        Rg = [r for r in R if r["gestor"] == nome]
        J = {"normal": g["normal"], "telaCheia": g.get("telaCheia", {}), "estado": g["estado"], "erros": []}
        antes = len(A.RES)
        A.conferir(J, Rg, R)
        resumo_rodada("Visão de cada gestor (filtro)", "Filtro: " + nome,
                      "O painel inteiro, com o filtro deste gestor",
                      "A conferência completa refeita só com as pessoas da equipe deste gestor.", antes, None)


def estados(X, R):
    nomes = {"digCiclo": "Batidas digitadas", "jorCiclo": "Jornada", "pnCiclo": "PontoNet", "pauBatCiclo": "Pausas (batidas)",
             "absCiclo": "Afastamentos (mês)", "absMostrar": "Afastamentos (tipo)", "tendIndicador": "Tendências (assunto)",
             "bhMeses": "Banco de horas (mês)"}
    for e in X.get("estados", []):
        J = {"normal": e["normal"], "telaCheia": e.get("telaCheia", {}), "estado": e["estado"], "erros": []}
        antes = len(A.RES)
        A.conferir(J, R, R)
        sel = e["estado"].get(e["seletor"]) if e["seletor"] != "bhMeses" else ", ".join(m for m, v in e["estado"]["bhMeses"] if v)
        resumo_rodada("Outros meses e opções dos seletores", nomes.get(e["seletor"], e["seletor"]) + ": " + str(sel),
                      "O painel inteiro com esta opção escolhida",
                      "A conferência completa refeita com esta opção no seletor da seção.", antes, None)


def origem(X):
    """Janela 'de onde vem': a quantidade de registros da janela tem que ser o número da barra clicada."""
    S = "Janelas \"de onde vem\""
    n_int = lambda t: int(t.replace(".", ""))
    tot = iguais = so_abre = 0; falhas = []
    for o in X.get("origem", []):
        tot += 1
        cont = o["alvo"].split("#")[0]
        if not o.get("aberto") or not o.get("linhas"):
            m0 = re.search(r": 0(?:\D|$)", o.get("dica") or "")
            if m0: so_abre += 1; continue          # barra com 0: não há registros para mostrar
            falhas.append(cont + ": a janela não abriu ou veio vazia"); continue
        dica = o.get("dica") or ""
        m_sub = re.match(r"^([\d.]+)", o.get("sub") or "")
        esp = None
        if cont in ("absHeat", "pauHeat"): esp = 1                     # um quadrado do mapa = um registro
        elif cont == "chartPnGestor": m = re.search(r"([\d.]+) pedido", dica); esp = m and n_int(m.group(1))
        elif cont == "chartJorDezHoras": m = re.search(r": ([\d.]+) dias? acima", dica); esp = m and n_int(m.group(1))
        elif cont == "chartPauMaior": m = re.search(r"([\d.]+) dia\(s\) acima", dica); esp = m and n_int(m.group(1))
        elif cont in ("chartPauEvolucao", "chartPauGrupo", "chartPauHora"): esp = None   # barra em %, pausas ou trechos; a janela lista pessoas/dias
        else:
            m = re.search(r": ([\d.]+)", dica); esp = m and n_int(m.group(1))
        if esp is None or not m_sub: so_abre += 1; continue
        if n_int(m_sub.group(1)) == esp: iguais += 1
        else: falhas.append("%s: barra %s, janela %s" % (cont, esp, m_sub.group(1)))
    A.item(S, "Clique nas barras, pedaços, fatias, pontos e quadrados",
           "A janela abre e a quantidade de registros é a mesma do número clicado",
           "Cada elemento clicável dos gráficos (até 40 por gráfico) foi clicado. A janela tem que abrir com os registros e, quando "
           "o número da barra é uma contagem, a janela tem que listar exatamente essa quantidade.",
           "%d cliques: %d com a mesma quantidade, %d só abertura (barra em %% ou em outra unidade)" % (tot, tot - so_abre, so_abre),
           ("%d cliques: %d com a mesma quantidade, %d só abertura (barra em %% ou em outra unidade)" % (tot, iguais, so_abre)) if not falhas else falhas[:6],
           not falhas, obs=("%d com diferença" % len(falhas)) if falhas else "")


def main():
    args = sys.argv[1:]
    saida = None
    if "--json" in args:
        i = args.index("--json"); saida = args[i + 1]; del args[i:i + 2]
    X = {}
    for f in [a for a in args if a.endswith(".json")]:
        d = json.load(open(f, encoding="utf-8"))
        for k, v in d.items():
            if isinstance(v, list): X.setdefault(k, []).extend(v)
            elif isinstance(v, dict): X.setdefault(k, {}).update(v)
    R = A.ler_varios([a for a in args if a.endswith(".csv")])
    A.R_FULL = R
    if X.get("cards"): cards(X, R)
    if X.get("gestores"): gestores(X, R)
    if X.get("estados"): estados(X, R)
    if X.get("origem"): origem(X)
    for r in A.RES:
        print(("  OK    " if r["ok"] else "  FALHA ") + r["card"] + " — " + r["item"] + ("" if r["ok"] else "\n        " + r["painel"][:900]))
    n = sum(r["ok"] for r in A.RES)
    print("\nTotal: %d de %d itens conferem." % (n, len(A.RES)))
    if saida: json.dump(A.RES, open(saida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
