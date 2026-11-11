# -*- coding: utf-8 -*-
"""
CONFERÊNCIA INDEPENDENTE DO PAINEL (11/11/2026)
================================================================================
Para que serve: provar que cada número do painel (Dashboard/index.html) está certo.

Como prova: este script NÃO usa o código do painel. Ele lê o(s) mesmo(s) CSV(s) que o
painel abre, refaz cada conta do zero, em Python, seguindo as regras de negócio
combinadas com o RH (CLAUDE.md, seções 5.4 e 6), e compara com o que a tela mostrou.
O que a tela mostrou é lido por ferramentas/extrair_tudo_painel.js (abre o painel no
navegador, carrega os CSVs e grava num JSON tudo o que aparece: números, balões dos
gráficos, legendas e tabelas).

Uso:
  NODE_PATH=$(npm root -g) node ferramentas/extrair_tudo_painel.js "$PWD/Dashboard/index.html" painel.json arquivo.csv
  python3 ferramentas/auditar_csv_painel.py painel.json arquivo.csv [mais.csv ...] [--json resultado.json]

Saída: uma linha por item conferido, "OK" ou "FALHA", agrupada por seção do painel,
com a explicação em português de como o número é calculado. No fim, o total.
Com --json, grava também a lista de itens (usada para montar o relatório em Word/PDF).

Para um leigo: pense numa calculadora separada, feita por outra pessoa, que refaz todas
as contas a partir do mesmo arquivo. Se as duas chegam no mesmo número, o painel está
certo. Se não chegam, o item aparece como FALHA com os dois valores lado a lado.
"""
import csv, io, json, re, sys, unicodedata
from collections import Counter, defaultdict, OrderedDict
from datetime import date, datetime, timedelta

# ------------------------------------------------------------------ 1. leitura do CSV
TIPO_AUD = "Auditoria Extra e Falta (3 Meses)"
TIPO_RES_PAU = "Resumo Pausas Térmicas"
TIPO_RES_HORAS = "Resumo Horas Cartão"
TIPO_ORG = "Cadastro ORG"
TIPO_DIG = "Marcação Digitada"
TIPO_PAU_DIA = "Pausa Térmica Dia"
TIPO_PN = "Justificativa PontoNet"
TIPO_ABS = "Absenteísmo"
TIPO_AFAST = "Afastamento"
TIPO_BH = "Banco de Horas Período"
TIPO_BH_DIA = "BH Dia"
TIPOS_JORNADA = ["Horas Excedentes", "Intrajornada", "Interjornada Semanal", "Interjornada"]
TIPOS_CURTAS = ["Falta < 15min", "Extra < 15min"]
TIPOS_HE = ["Hora Extra", "Hora Extra 100%"]
TIPOS_FIXOS = [TIPO_AUD, TIPO_RES_PAU, TIPO_RES_HORAS, TIPO_ORG, TIPO_DIG, TIPO_PAU_DIA, TIPO_PN,
               TIPO_ABS, TIPO_AFAST, TIPO_BH, TIPO_BH_DIA] + TIPOS_JORNADA + TIPOS_CURTAS + TIPOS_HE
NAO_OCORRENCIA = set(TIPOS_FIXOS)
DIAS = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]          # weekday() 0 = segunda
MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def sem_acento(s):
    return "".join(c for c in unicodedata.normalize("NFD", s or "") if unicodedata.category(c) != "Mn")


def norm(s):
    """Texto para comparar: sem acento, minúsculo, só letras/números e '_'."""
    s = sem_acento(str(s or "")).lower().strip()
    return re.sub(r"[^a-z0-9]+", "_", s).strip("_")


def chave_tipo(t):
    t = str(t or "")
    if re.search("[ÃÂ]", t):                       # acento estragado (UTF-8 lido como ANSI)
        try: t = t.encode("latin-1").decode("utf-8")
        except Exception: pass
    return re.sub(r"[^a-z0-9<%()]", "", sem_acento(t).lower())


_CANON = {}
for _t in TIPOS_FIXOS:
    _CANON["a:" + chave_tipo(_t)] = _t
    _CANON["b:" + re.sub(r"[^a-z0-9<%()]", "", _t.lower())] = _t


def tipo_canonico(t):
    return _CANON.get("a:" + chave_tipo(t)) or _CANON.get("b:" + re.sub(r"[^a-z0-9<%()]", "", str(t).lower())) or t


def erro_excel(v):
    return bool(re.match(r"^(erro|error) \d+$|^#(n/d|n/a|ref!|valor!|value!|nome\?|name\?)$", str(v or "").strip(), re.I))


def texto(v):
    v = str(v or "").strip()
    return "" if erro_excel(v) else v


def numero(v):
    s = str(v or "").strip()
    if not s: return 0.0
    s = re.sub(r"[^\d,.\-]", "", s)
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".", 1)
    try: return float(s)
    except ValueError:
        m = re.match(r"-?\d+(\.\d+)?", s)
        return float(m.group(0)) if m else 0.0


def data(v):
    v = str(v or "").strip()
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})(?:[ T](\d{2}):(\d{2}))?", v) or None
    if m: return datetime(int(m[3]), int(m[2]), int(m[1]), int(m[4] or 0), int(m[5] or 0))
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?", v)
    if m: return datetime(int(m[1]), int(m[2]), int(m[3]), int(m[4] or 0), int(m[5] or 0))
    return None


def ler_csv(caminho):
    bruto = open(caminho, "rb").read()
    txt = bruto.decode("utf-8-sig", errors="replace")
    prim = txt.split("\n", 1)[0]
    sep = ";" if prim.count(";") > prim.count(",") else ","
    linhas = list(csv.reader(io.StringIO(txt), delimiter=sep))
    cab = [norm(c) for c in linhas[0]]
    col = lambda nome: cab.index(nome) if nome in cab else None
    pega = lambda r, nome: (r[col(nome)] if col(nome) is not None and col(nome) < len(r) else "")
    out = []
    for r in linhas[1:]:
        if not r: continue
        colab = pega(r, "colaborador").strip()
        dt = data(pega(r, "data"))
        if not colab or not dt: continue
        tipo = tipo_canonico(pega(r, "tipo_ocorrencia").strip())
        if norm(tipo) in ("problema_horario", "jussara"): continue
        # (11/11) as linhas da planilha de tratamento (ocorrências e "Hora Extra") não são mais usadas
        if tipo not in NAO_OCORRENCIA or tipo in TIPOS_HE: continue
        if tipo == TIPO_DIG and texto(pega(r, "origem_marcacao")).upper() == "E": continue
        vazio_none = lambda nome: None if texto(pega(r, nome)) == "" else numero(pega(r, nome))
        out.append(dict(
            data=dt, dia=dt.date(), colaborador=colab, matricula=pega(r, "matricula").strip(),
            gestor=texto(pega(r, "gestor")) or "Sem gestor", cargo=texto(pega(r, "cargo")) or "Sem cargo",
            setor=pega(r, "setor").strip() or "Sem setor", grupo=texto(pega(r, "grupo")),
            tipo=tipo or "Não informado", situacao=pega(r, "situacao").strip(), duracao=numero(pega(r, "duracao_minutos")),
            tratativa=data(pega(r, "data_tratativa_pontonet")),
            mes=numero(pega(r, "ocorrencias_extra_falta_mes_atual")), tres=numero(pega(r, "ocorrencias_extra_falta_3_meses")),
            pc=numero(pega(r, "pausas_corretas")), pm=numero(pega(r, "pausas_menor_20min")), pM=numero(pega(r, "pausas_maior_20min")),
            tc=numero(pega(r, "trabalho_correto_140")), tM=numero(pega(r, "trabalho_maior_140")), tm=numero(pega(r, "trabalho_menor_140")),
            imp=numero(pega(r, "pausas_marcacoes_impares")),
            he=numero(pega(r, "minutos_hora_extra_cartao")), bh=numero(pega(r, "minutos_banco_horas_cartao")),
            just=texto(pega(r, "justificativa_marcacao")), origem=texto(pega(r, "origem_marcacao")),
            dias_aus=numero(pega(r, "dias_ausencia")), emenda=texto(pega(r, "emenda_folga")),
            min_trab=vazio_none("minutos_trabalhados"), marc=texto(pega(r, "marcacoes_dia")),
            tipo_just=texto(pega(r, "tipo_justificativa")), dias_col=vazio_none("dias_colaborador"),
            dias_gest=vazio_none("dias_gestor"),
            bh_ini=data(pega(r, "bh_inicio_periodo")), bh_ant=numero(pega(r, "bh_saldo_anterior_minutos")),
            bh_cred=numero(pega(r, "bh_credito_minutos")), bh_deb=numero(pega(r, "bh_debito_minutos")),
            bh_fin=numero(pega(r, "bh_saldo_final_minutos")), bh_pag=numero(pega(r, "bh_pagas_minutos")),
            bh_ani=numero(pega(r, "bh_anistiadas_minutos")),
            _linha="\u0001".join(r)))
    return out


def ler_varios(caminhos):
    """Junta vários CSVs: linha idêntica que aparece em arquivos DIFERENTES conta uma vez."""
    vistos, todos = {}, []
    for i, c in enumerate(caminhos):
        for rec in ler_csv(c):
            dono = vistos.get(rec["_linha"])
            if dono is not None and dono != i: continue
            vistos[rec["_linha"]] = i
            rec["_arq"] = i
            todos.append(rec)
    return todos


# ------------------------------------------------------------------ 2. ajudas
def ciclo(d):
    """Ciclo de apuração 28 -> 27: o dia 28/09 já é do ciclo de outubro."""
    y, m = d.year, d.month
    if d.day >= 28:
        m += 1
        if m > 12: m, y = 1, y + 1
    return "%d-%02d" % (y, m)


def rotulo_ciclo(k):
    y, m = k.split("-")
    return MESES[int(m) - 1] + "/" + y[2:]


def ciclo_ini_fim(k):
    y, m = map(int, k.split("-"))
    ini = date(y if m > 1 else y - 1, m - 1 if m > 1 else 12, 28)
    return ini, date(y, m, 27)


def int_br(s):
    """'1.234' -> 1234 (número inteiro como o painel escreve)."""
    s = str(s).strip().split(" ")[0].replace(".", "")
    m = re.match(r"-?\d+", s.replace("−", "-"))
    return int(m.group(0)) if m else None


def fmt_dec(x):
    return ("%.1f" % x).replace(".", ",")


def dur_txt(m):
    """Minutos -> '5h 40min' (como o painel)."""
    m = int(round(abs(m)))
    h, mi = divmod(m, 60)
    if h and mi: return "%dh %02dmin" % (h, mi)
    if h: return "%dh" % h
    return "%dmin" % mi


def minutos_de(txt):
    """'+42h 20min' / '−7h 25min' / '35min' / '335h' -> minutos com sinal."""
    t = str(txt).replace("−", "-").strip()
    neg = t.startswith("-")
    h = re.search(r"(\d+)h", t); mi = re.search(r"(\d+)\s*min", t)
    v = (int(h.group(1)) * 60 if h else 0) + (int(mi.group(1)) if mi else 0)
    return -v if neg else v


def tips_dict(lista):
    """['Seg: 16', 'PESSOA 05: 5 · clique...'] -> {'Seg': 16, 'PESSOA 05': 5}"""
    out = {}
    for t in lista or []:
        m = re.match(r"^(.*?):\s*([\d.]+)", t)
        if m: out[m.group(1).strip()] = int(m.group(2).replace(".", ""))
    return out



def analisa_batidas(txt):
    """Batidas do dia 'H:MM L ...' -> trechos de trabalho e pausas (regra do RH).
    A hora que volta (passou da meia-noite) ganha +24h. Entre batidas, os trechos alternam:
    trabalho, pausa, trabalho... Pausa térmica = trecho entre duas batidas F."""
    tk = str(txt or "").split()
    bat, base, ant = [], 0, None
    for i, x in enumerate(tk):
        m = re.match(r"^(\d{1,2}):(\d{2})$", x)
        if not m: continue
        letra = tk[i + 1].upper() if i + 1 < len(tk) and re.match(r"^[A-Za-z]$", tk[i + 1]) else ""
        v = int(m.group(1)) * 60 + int(m.group(2))
        if ant is not None and v + base < ant: base += 1440
        bat.append((v + base, letra)); ant = v + base
    trab, ff = [], 0
    for j in range(len(bat) - 1):
        if j % 2 == 0: trab.append((bat[j][0], bat[j + 1][0] - bat[j][0], j + 1))
        elif bat[j][1] + bat[j + 1][1] == "FF": ff += 1
    total = sum(d for _, d, _ in trab)
    iF = next((k for k in range(1, len(bat)) if bat[k][1] == "F"), -1)
    primeira = None if iF < 0 else sum(d for _, d, f in trab if f <= iF)
    return dict(trab=[(a, b) for a, b, _ in trab], acima=[d for _, d, _ in trab if d > 100], primeira=primeira,
                impar=len(bat) % 2 == 1, sem=total > 100 and not ff)


def nome_tipo_pn(t):
    """Motivo do PontoNet em palavras simples (regra combinada para a tela): dispositivo -> 'Problema no
    relógio de ponto'; externa -> 'Atividade externa'; esquec... -> 'Esquecimento'; exclus... -> 'Exclusão de
    batida errada'; o resto com a 1ª letra maiúscula."""
    n = sem_acento(t or "").lower()
    if not n.strip(): return "Sem motivo"
    for chave, nome in (("dispositivo", "Problema no relógio de ponto"), ("externa", "Atividade externa"),
                        ("esquec", "Esquecimento"), ("exclus", "Exclusão de batida errada")):
        if chave in n: return nome
    t = t.strip(); return t[:1].upper() + t[1:].lower()

# ------------------------------------------------------------------ 3. registro dos resultados
RES = []


def item(secao, card, o_que, como, esperado, painel, ok=None, obs=""):
    if ok is None: ok = (esperado == painel)
    RES.append(OrderedDict(secao=secao, card=card, item=o_que, como=como,
                           esperado=resumo(esperado), painel=resumo(painel), ok=bool(ok), obs=obs))


def resumo(v):
    if isinstance(v, dict):
        itens = list(v.items())
        s = "; ".join("%s = %s" % (k, x) for k, x in itens[:12])
        return s + ("; ... (+%d)" % (len(itens) - 12) if len(itens) > 12 else "")
    if isinstance(v, (list, tuple)):
        s = "; ".join(str(x) for x in v[:12])
        return s + ("; ... (+%d)" % (len(v) - 12) if len(v) > 12 else "")
    return str(v)


def compara_dict(secao, card, o_que, como, esperado, painel, so_chaves_do_painel=False, obs=""):
    """Compara contagens por nome. so_chaves_do_painel: o gráfico mostra só os primeiros."""
    if so_chaves_do_painel:
        dif = {k: (esperado.get(k), v) for k, v in painel.items() if esperado.get(k) != v}
        ok = not dif and bool(painel)
        # os que ficaram de fora não podem ter número maior que o menor mostrado
        if painel:
            menor = min(painel.values())
            fora = [k for k, v in esperado.items() if k not in painel and v > menor]
            if fora: ok = False; dif.update({k: (esperado[k], "não mostrado") for k in fora})
        item(secao, card, o_que, como, {k: painel[k] for k in painel}, painel if ok else dif, ok, obs)
    else:
        item(secao, card, o_que, como, dict(sorted(esperado.items(), key=lambda x: -x[1])),
             dict(sorted(painel.items(), key=lambda x: -x[1])), esperado == painel, obs)


def tabela(J, id_, cheia=True):
    """Tabela de um card; com cheia=True prefere a leitura em tela cheia (lista inteira)."""
    fonte = (J["telaCheia"].get(id_) if cheia else None) or J["normal"].get(id_) or {}
    t = (fonte.get("tabelas") or [None])[0]
    return t or {"head": [], "rows": []}


def tips(J, id_):
    return tips_dict(((J["telaCheia"].get(id_) or {}).get("tips")) or (J["normal"].get(id_) or {}).get("tips"))


def numeros(J, id_):
    return {n[0].upper(): (n[1], n[2]) for n in (J["normal"].get(id_) or {}).get("numeros", [])}


# ------------------------------------------------------------------ 4. as conferências
def conferir(J, R):
    est = J["estado"]
    p_ini = datetime.strptime(est["periodoIni"], "%Y-%m-%d").date()
    p_fim = datetime.strptime(est["periodoFim"], "%Y-%m-%d").date()

    # ---------------------------------------------------------- período padrão
    S = "Período padrão do painel"
    fim_cartao = [r["dia"] for r in R if r["tipo"] in (TIPO_RES_HORAS, TIPO_AUD)]
    esp_fim = max(fim_cartao)
    esp_ini = ciclo_ini_fim(ciclo(esp_fim))[0]
    item(S, "Filtro de período (lateral)", "Data inicial",
         "Dia 28 do mês anterior ao fim do cartão ponto mais recente (começo do ciclo de 28 a 27).",
         esp_ini.strftime("%d/%m/%Y"), p_ini.strftime("%d/%m/%Y"))
    item(S, "Filtro de período (lateral)", "Data final",
         "Último dia do cartão ponto mais recente (data das linhas de resumo do cartão, dia 27).",
         esp_fim.strftime("%d/%m/%Y"), p_fim.strftime("%d/%m/%Y"))
    no_periodo = lambda r: p_ini <= r["dia"] <= p_fim
    oc = []   # (11/11) as ocorrências da planilha de tratamento não são mais usadas pelo painel
    org = {}
    for r in R:
        if r["tipo"] == TIPO_ORG and (r["matricula"] or r["colaborador"]) not in org:
            org[r["matricula"] or r["colaborador"]] = r

    # ---------------------------------------------------------- Marcações digitadas
    S = "Marcações digitadas"
    dig = [r for r in R if r["tipo"] == TIPO_DIG]
    kp = numeros(J, "digKpis")
    item(S, "Números do topo", "Marcações digitadas",
         "Batidas lançadas à mão no sistema (só as de origem \"D\"), em todos os ciclos do arquivo (seletor padrão).",
         len(dig), int_br(kp.get("MARCAÇÕES DIGITADAS", ("?",))[0]))
    item(S, "Números do topo", "Colaboradores",
         "Pessoas diferentes com ao menos uma batida digitada.",
         len({r["colaborador"] for r in dig}), int_br(kp.get("COLABORADORES", ("?",))[0]))
    sub = kp.get("COLABORADORES", ("", ""))[1]
    item(S, "Números do topo", "\"de N pessoas na lista de equipes\"",
         "Tamanho da equipe do recorte: pessoas da lista de equipes (uma por matrícula).",
         len(org), int_br((re.search(r"de ([\d.]+)", sub) or [0, "0"])[1]))
    esq = [r for r in dig if norm(r["situacao"]) == "inclusao_por_esquecimento"]
    k_esq = [k for k in kp if "ESQUEC" in k]
    if k_esq:
        item(S, "Números do topo", "Por esquecimento", "Batidas cujo motivo é \"Inclusão por esquecimento\".",
             len(esq), int_br(kp[k_esq[0]][0]))
    leg = {a: int_br(b) for a, b in (J["normal"]["chartDigMotivo"].get("legenda") or [])}
    m = Counter(r["situacao"] for r in dig)
    top = [k for k, _ in m.most_common()][:5]
    esp = {k: v for k, v in m.items() if k in top}
    if sum(v for k, v in m.items() if k not in top): esp["Outros"] = sum(v for k, v in m.items() if k not in top)
    compara_dict(S, "Por motivo (rosca)", "Quantidade por motivo", "Batidas digitadas contadas pelo motivo informado.", esp, leg)
    compara_dict(S, "Mês a mês", "Batidas por ciclo",
                 "Batidas digitadas agrupadas pelo ciclo 28 a 27 da data da batida (últimos 3 ciclos).",
                 Counter(rotulo_ciclo(ciclo(r["dia"])) for r in dig), tips(J, "chartDigCiclo"), so_chaves_do_painel=True)
    compara_dict(S, "Por dia da semana", "Batidas em cada dia", "Batidas digitadas pelo dia da semana.",
                 Counter(DIAS[r["dia"].weekday()] for r in dig), {k: v for k, v in tips(J, "chartDigDiaSemana").items() if v})
    compara_dict(S, "Por data", "Batidas em cada dia", "Batidas digitadas pela data.",
                 Counter(r["dia"].strftime("%d/%m") for r in dig), {k: v for k, v in tips(J, "chartDigData").items() if v})
    compara_dict(S, "Por hora do dia", "Batidas em cada hora", "Batidas digitadas pela hora da batida (5h = de 5:00 a 5:59).",
                 Counter("%dh" % r["data"].hour for r in dig), {k: v for k, v in tips(J, "chartDigHora").items() if v})
    compara_dict(S, "Por gestor", "Batidas por gestor (os 12 maiores)", "Batidas digitadas pelo gestor da pessoa.",
                 Counter(r["gestor"] for r in dig), tips(J, "chartDigGestor"), so_chaves_do_painel=True)
    compara_dict(S, "Quem mais teve batidas digitadas", "Batidas por pessoa (lista inteira em tela cheia)",
                 "Batidas digitadas por pessoa.", Counter(r["colaborador"] for r in dig), tips(J, "chartDigTop"))
    t = tabela(J, "tableDigJust")
    js = defaultdict(lambda: [0, set()])
    for r in dig:
        if r["just"]:
            k = re.sub(r"\s+", " ", sem_acento(r["just"]).lower().strip(" .;,"))
            js[k][0] += 1; js[k][1].add(r["colaborador"])
    esp = sorted([(v[0], len(v[1])) for v in js.values()], reverse=True)[:len(t["rows"])]
    pain = sorted([(int_br(x[1]), int_br(x[2])) for x in t["rows"]], reverse=True)
    item(S, "Justificativas mais usadas", "Batidas e pessoas de cada texto",
         "Textos de justificativa iguais (ignorando maiúscula, acento e pontuação no fim) somados.", esp, pain)
    item(S, "Lista de batidas", "Quantidade de linhas", "Uma linha por batida digitada.", len(dig), len(tabela(J, "tableDigLista")["rows"]))
    ciclos3 = sorted({ciclo(r["dia"]) for r in R if r["tipo"] in (TIPO_RES_HORAS, TIPO_AUD)})
    fim_k = ciclo(p_fim)
    todos_k = sorted({ciclo(r["dia"]) for r in R})
    c3 = [k for k in todos_k if k <= fim_k][-3:]
    por = defaultdict(Counter)
    for r in esq: por[r["colaborador"]][ciclo(r["dia"])] += 1
    t = tabela(J, "tableDigEsquecimento")
    pain = {row[0]: tuple(int_br(x) for x in row[2:2 + len(c3)]) for row in t["rows"]}
    esp = {k: tuple(por[k][c] for c in c3) for k in pain}
    item(S, "Esquecimento repetido", "Esquecimentos por pessoa nos 3 ciclos",
         "Batidas por esquecimento de cada pessoa, ciclo a ciclo (%s)." % ", ".join(rotulo_ciclo(c) for c in c3), esp, pain)

    # ---------------------------------------------------------- Afastamentos
    S = "Afastamentos"
    af = [r for r in R if r["tipo"] == TIPO_AFAST] or [r for r in R if r["tipo"] == TIPO_ABS]
    por_horas = lambda r: r["tipo"] == TIPO_AFAST and r["duracao"] > 0 and "atest" not in norm(r["situacao"])
    dias_af = defaultdict(list)          # ciclo -> [(pessoa, dia)]
    for r in af:
        if por_horas(r): continue
        n = max(1, int(round(r["dias_aus"] or 0)))
        for i in range(n):
            d = r["dia"] + timedelta(days=i)
            dias_af[ciclo(d)].append((r, d))
    t = tabela(J, "tableAbsEvolucao")
    pain, esp = {}, {}
    for row in t["rows"]:
        k = row[0].split(" ")[0]
        pain[k] = tuple(int_br(x) for x in row[1:5])
        c = [x for x in set(dias_af) | {ciclo(r["dia"]) for r in af} if rotulo_ciclo(x) == k]
        c = c[0] if c else None
        comeca = [r for r in af if not por_horas(r) and ciclo(r["dia"]) == c]
        esp[k] = (len(dias_af.get(c, [])), len(comeca), sum(1 for r in comeca if "atest" in norm(r["situacao"])),
                  len({r["matricula"] or r["colaborador"] for r, _ in dias_af.get(c, [])}))
    item(S, "Evolução por ciclo", "Dias, afastamentos, atestados e pessoas em cada ciclo",
         "Dias: cada afastamento vira os dias corridos que cobre (data de início + dias de ausência; mínimo 1 dia) e cada dia "
         "cai no ciclo 28 a 27 em que está. Afastamentos e atestados: contados no ciclo em que COMEÇAM. Pessoas: quem tem ao "
         "menos um dia afastado no ciclo. Afastamento por horas (saída médica, reunião curta) fica fora.", esp, pain)
    at = [r for r in af if "atest" in norm(r["situacao"])]
    rk = defaultdict(lambda: [0, 0.0, set(), 0])
    for r in at:
        x = rk[r["colaborador"]]
        x[0] += 1; x[1] += r["dias_aus"]
        for i in range(max(1, int(round(r["dias_aus"] or 0)))): x[2].add(ciclo(r["dia"] + timedelta(days=i)))
        if r["emenda"]: x[3] += 1
    t = tabela(J, "tableAbsRanking")
    pain = {row[0]: (int_br(row[2]), int_br(row[3]), int_br(row[4]), int_br(row[5])) for row in t["rows"]}
    esp = {k: (rk[k][0], int(round(rk[k][1])), len(rk[k][2]), rk[k][3]) for k in pain}
    item(S, "Ranking de atestados", "Atestados, dias, meses e encostados em folga por pessoa",
         "Por pessoa: número de atestados, soma dos dias de ausência informados, quantos ciclos diferentes tiveram dia de "
         "atestado e quantos atestados encostam numa folga (dia antes ou depois).", esp, pain)

    # ---------------------------------------------------------- Jornada
    S = "Jornada"
    jor = [r for r in R if r["tipo"] in TIPOS_JORNADA]
    kp = numeros(J, "jorKpis")
    nomes = {"Horas Excedentes": "HORAS EXCEDENTES", "Intrajornada": "INTRAJORNADA",
             "Interjornada Semanal": "INTERJORNADA SEMANAL", "Interjornada": "INTERJORNADA"}
    for tp, rot in nomes.items():
        lst = [r for r in jor if r["tipo"] == tp]
        item(S, "Números do topo", rot.capitalize() + " (ocorrências e pessoas)",
             "Linhas do relatório do sistema desse tipo, em todos os ciclos do arquivo; pessoas diferentes.",
             (len(lst), len({r["colaborador"] for r in lst})),
             (int_br(kp.get(rot, ("?", ""))[0]), int_br(kp.get(rot, ("", "?"))[1])))
    he_l = [r for r in jor if r["tipo"] == "Horas Excedentes"]
    acima = sum(max(0, (r["min_trab"] or 0) - 600) for r in he_l)
    sub = kp.get("HORAS EXCEDENTES", ("", ""))[1]
    item(S, "Números do topo", "Horas acima de 10h (soma)",
         "Para cada dia de horas excedentes: total de horas do dia (coluna Total do sistema) menos 10h; somado.",
         dur_txt(acima), (re.search(r"·\s*(.*?) acima", sub) or [None, ""])[1])
    compara_dict(S, "Mês a mês", "Ocorrências por ciclo", "Ocorrências de jornada pelo ciclo 28 a 27 da data.",
                 Counter(rotulo_ciclo(ciclo(r["dia"])) for r in jor), tips(J, "chartJorCiclo"), so_chaves_do_painel=True)
    compara_dict(S, "Por gestor", "Ocorrências por gestor", "Ocorrências de jornada pelo gestor da pessoa.",
                 Counter(r["gestor"] for r in jor), tips(J, "chartJorGestor"), so_chaves_do_painel=True)
    t = tabela(J, "tableJorGestor")
    pain = {row[0]: tuple(int_br(x) or 0 for x in row[1:7]) for row in t["rows"]}   # "–" = 0
    esp = {}
    for g in pain:
        lst = [r for r in jor if r["gestor"] == g]
        c = Counter(r["tipo"] for r in lst)
        esp[g] = (len(lst), c["Horas Excedentes"], c["Intrajornada"], c["Interjornada Semanal"], c["Interjornada"],
                  len({r["colaborador"] for r in lst}))
    item(S, "Tabela por gestor", "Ocorrências por tipo e pessoas, por gestor",
         "Por gestor: total, cada um dos 4 tipos e pessoas diferentes.", esp, pain)
    item(S, "Lista de ocorrências", "Quantidade de linhas", "Uma linha por ocorrência de jornada.",
         len(jor), len(tabela(J, "tableJorLista")["rows"]))
    dez = defaultdict(lambda: [0, 0])
    for r in he_l:
        dez[r["colaborador"]][0] += 1; dez[r["colaborador"]][1] += int(max(0, (r["min_trab"] or 0) - 600))
    pain = {}
    for tp in ((J["telaCheia"].get("chartJorDezHoras") or {}).get("tips") or J["normal"]["chartJorDezHoras"].get("tips") or []):
        m = re.match(r"^(.*?): (\d+) dias? acima de 10h, (.*?) além", tp)
        if m: pain[m.group(1)] = (int(m.group(2)), minutos_de(m.group(3)))
    esp = {k: tuple(dez[k]) for k in pain}
    item(S, "Mais de 10h no dia, por pessoa", "Dias e minutos acima de 10h por pessoa",
         "Por pessoa: dias com horas excedentes e soma do que passou de 10h.", esp, pain)

    aud = {}
    for r in R:
        if r["tipo"] == TIPO_AUD: aud[r["colaborador"]] = (int(r["mes"]), int(r["tres"]))
    t = tabela(J, "tableExtraFalta")
    pain = {row[0]: (int_br(row[2]), int_br(row[3])) for row in t["rows"]}
    esp = {k: aud.get(k) for k in pain}
    maiores = sorted(aud.values(), key=lambda x: (-x[1], -x[0]))[:len(pain)]
    ok = esp == pain and sorted(pain.values(), key=lambda x: (-x[1], -x[0])) == maiores
    alerta_ok = all(("⚠" in row[3]) == (int_br(row[3]) >= 10) for row in t["rows"])
    item("Jornada", "Do cartão ponto: extra e falta no mesmo dia", "Este mês e 3 meses, por pessoa (os 10 maiores)",
         "Do cartão ponto: dias em que a pessoa teve hora extra e banco de horas (falta) no mesmo dia, no mês atual e "
         "somando os 3 últimos cartões. O alerta ⚠ aparece a partir de 10 na soma dos 3 meses.",
         esp, pain, ok and alerta_ok)
    cur = defaultdict(lambda: [0, 0])
    for r in R:
        if r["tipo"] in TIPOS_CURTAS and no_periodo(r):
            cur[r["colaborador"]][0 if r["tipo"].startswith("Falta") else 1] += 1
    t = tabela(J, "tableCurtas")
    pain = {row[0]: (int_br(row[1]), int_br(row[2]), int_br(row[3])) for row in t["rows"]}
    esp = {k: (v[0], v[1], v[0] + v[1]) for k, v in cur.items()}
    item("Jornada", "Do cartão ponto: ocorrências curtas (menos de 15 min)", "Faltas, extras e total por pessoa (lista inteira)",
         "Do cartão ponto do mês: dias com falta ou extra de menos de 15 minutos (tolerância), contados por pessoa.",
         esp, pain)

    # ---------------------------------------------------------- Pausas térmicas
    S = "Pausas térmicas"
    res = [r for r in R if r["tipo"] == TIPO_RES_PAU]
    meses_p = sorted({ciclo(r["dia"]) for r in res})
    if res:
        atual = [r for r in res if ciclo(r["dia"]) == meses_p[-1]]
        kp = numeros(J, "pauKpis")
        item(S, "Números do topo", "Pessoas no relatório",
             "Pessoas na tabela de pausas do mês mais recente.", len(atual), int_br(kp.get("PESSOAS NO RELATÓRIO", ("?",))[0]))
        pc = sum(r["pc"] for r in atual); tot = pc + sum(r["pm"] + r["pM"] for r in atual)
        v = kp.get("PAUSAS CERTAS", ("?", ""))
        item(S, "Números do topo", "Pausas certas (%)",
             "Pausas de 20 minutos ÷ todas as pausas (certas + menores + maiores), no mês mais recente.",
             ("%d%%" % round(100 * pc / tot) if tot else "–", int(pc), int(tot)),
             (v[0], int_br((re.match(r"([\d.]+) de", v[1]) or [0, "0"])[1]), int_br((re.search(r"de ([\d.]+)", v[1]) or [0, "0"])[1])))
        t = tabela(J, "tablePausasTermicas")
        pain = {row[0]: tuple(int_br(x) for x in row[2:9]) for row in t["rows"]}
        esp = {r["colaborador"]: tuple(int(x) for x in (r["pc"], r["pm"], r["pM"], r["tc"], r["tM"], r["tm"], r["imp"])) for r in atual}
        item(S, "Tabela de pausas", "Os 7 números de cada pessoa (lista inteira)",
             "Copiados da tabela de pausas do sistema (mês mais recente), sem recálculo.", esp, pain)
        t = tabela(J, "tablePauEvolucao")
        pain = {row[0]: (int_br(row[1]), row[2], int_br(row[3]), int_br(row[4])) for row in t["rows"]}
        esp = {}
        for k in meses_p:
            l = [r for r in res if ciclo(r["dia"]) == k]
            pc = sum(r["pc"] for r in l); tp = pc + sum(r["pm"] + r["pM"] for r in l)
            esp[rotulo_ciclo(k)] = (len(l), "%d%%" % round(100 * pc / tp) if tp else "–", int(sum(r["tM"] for r in l)), int(sum(r["imp"] for r in l)))
        esp = {k: esp.get(k) for k in pain}
        item(S, "Evolução mês a mês", "Pessoas, % certas, trechos acima de 1h40 e batida faltando por mês",
             "Uma linha por mês de relatório de pausas: soma das pessoas do mês.", esp, pain)
        t = tabela(J, "tablePauAcimaGestor")
        pain = {row[0]: (int_br(row[1]), int_br(row[2])) for row in t["rows"]}
        esp = {g: (sum(1 for r in atual if r["gestor"] == g), int(sum(r["tM"] for r in atual if r["gestor"] == g))) for g in pain}
        item(S, "Acima de 1h40 por gestor", "Pessoas e trechos acima de 1h40",
             "Mês mais recente: pessoas da equipe no relatório e soma dos trechos de trabalho acima de 1h40.", esp, pain)

    # ---------------------------------------------------------- PontoNet
    S = "PontoNet"
    pn = [r for r in R if r["tipo"] == TIPO_PN and norm(r["situacao"]) == "integrado"]
    if pn:
        k_atual = max(ciclo(r["dia"]) for r in pn)
        cic = [r for r in pn if ciclo(r["dia"]) == k_atual]
        kp = numeros(J, "pnKpis")
        item(S, "Números do topo", "Justificativas aprovadas",
             "Pedidos do PontoNet com situação INTEGRADO (aprovados), no ciclo mais recente (seletor padrão).",
             len(cic), int_br(kp.get("JUSTIFICATIVAS APROVADAS", ("?",))[0]))
        item(S, "Números do topo", "Pessoas com justificativa aprovada", "Pessoas diferentes nesses pedidos.",
             len({r["colaborador"] for r in cic}), int_br(kp.get("PESSOAS COM JUSTIFICATIVA APROVADA", ("?",))[0]))
        compara_dict(S, "Mês a mês", "Aprovados por ciclo", "Pedidos aprovados pelo ciclo da batida que faltou.",
                     Counter(rotulo_ciclo(ciclo(r["dia"])) for r in pn), tips(J, "chartPnCiclo"), so_chaves_do_painel=True)
        t = tabela(J, "tablePnTempos")
        pain = {row[0]: (int_br(row[1]), row[2], row[3]) for row in t["rows"]}
        esp = {}
        for row in t["rows"]:
            l = [r for r in pn if rotulo_ciclo(ciclo(r["dia"])) == row[0]]
            md = lambda vs: (fmt_dec(sum(vs) / len(vs)) + " dias") if vs else "–"
            esp[row[0]] = (len(l), md([r["dias_col"] for r in l if r["dias_col"] is not None]),
                           md([r["dias_gest"] for r in l if r["dias_gest"] is not None]))
        item(S, "Tempos por ciclo", "Aprovados, tempo médio do colaborador e do gestor",
             "Média dos dias contados pelo próprio PontoNet (colaborador: da falta até pedir; gestor: do pedido até avaliar), "
             "só nos aprovados.", esp, pain)
        item(S, "Lista de aprovados", "Quantidade de linhas", "Um pedido aprovado do ciclo por linha.",
             len(cic), len(tabela(J, "tablePnLista")["rows"]))
        dem = defaultdict(list)
        for r in cic:
            if r["tratativa"]:
                h = (r["tratativa"] - r["data"]).total_seconds() / 86400
                if h >= 0: dem[r["gestor"]].append(h)
        t = tabela(J, "tableDemoraGestor")
        pain = {row[0]: (int_br(row[1]), row[2], row[3]) for row in t["rows"]}
        dtxt = lambda dd: fmt_dec(dd * 24) + "h" if dd < 1 else fmt_dec(dd) + " dias"   # menos de 1 dia: em horas
        esp = {g: (len(dem[g]), dtxt(sum(dem[g]) / len(dem[g])), dtxt(max(dem[g]))) if dem[g] else None for g in pain}
        item(S, "Demora por gestor", "Aprovadas, demora média e maior demora",
             "Do dia e hora da batida que faltou até a data em que o pedido foi aprovado, em dias, no ciclo escolhido.", esp, pain)

    # ---------------------------------------------------------- Banco de horas
    S = "Banco de horas"
    bd = [r for r in R if r["tipo"] == TIPO_BH_DIA]
    if bd:
        meses = [m for m, marc in est["bhMeses"] if marc]
        ks = [k for k in sorted({ciclo(r["dia"]) for r in bd}) if rotulo_ciclo(k) in [m.split(" ")[0] for m in meses]]
        sel = [r for r in bd if ciclo(r["dia"]) in ks]
        kp = numeros(J, "bhKpis")
        item(S, "Números do topo", "Horas extras",
             "Soma da hora extra 50% do cartão ponto, dia a dia, nos meses escolhidos (a 100% é paga à parte).",
             dur_txt(sum(r["he"] for r in sel)), kp.get("HORAS EXTRAS", ("?",))[0])
        k_falta = [k for k in kp if "FALTA" in k]
        if k_falta:
            item(S, "Números do topo", "Horas falta",
                 "Soma da coluna BH do cartão ponto (falta e atraso descontados do banco), nos meses escolhidos.",
                 dur_txt(sum(r["bh"] for r in sel)), kp[k_falta[0]][0])
        t = tabela(J, "tableBHPessoas")
        pain = {row[0]: (minutos_de(row[3]), minutos_de(row[4]), minutos_de(row[5]) - minutos_de(row[2]), int_br(row[7]), int_br(row[8])) for row in t["rows"]}
        esp = {}
        for k in pain:
            l = [r for r in sel if r["colaborador"] == k]
            e, f = int(sum(r["he"] for r in l)), int(sum(r["bh"] for r in l))
            esp[k] = (e, f, e - f, sum(1 for r in l if r["he"] > 0), sum(1 for r in l if r["bh"] > 0))
        item(S, "Banco de cada pessoa", "Horas extras, horas falta, variação do banco e dias com extra/falta",
             "Por pessoa, nos meses escolhidos: soma das horas extras e das horas falta do cartão ponto; o banco no fim menos o "
             "banco no começo tem que ser igual a extras − faltas.", esp, pain)
        rel = [r for r in R if r["tipo"] == TIPO_BH]
        if rel:
            item(S, "Relatório de banco de horas", "Contas fechadas no relatório do sistema",
                 "Para cada pessoa do relatório: saldo anterior + crédito − débito (− pagas + anistiadas) = saldo final.",
                 "%d de %d pessoas" % (len(rel), len(rel)),
                 "%d de %d pessoas" % (sum(1 for r in rel if abs(r["bh_ant"] + r["bh_cred"] - r["bh_deb"] - r["bh_pag"] + r["bh_ani"] - r["bh_fin"]) <= 1), len(rel)),
                 obs="Conferência dos dados de entrada (o relatório do sistema), não de uma conta do painel.")

    # ---------------------------------------------------------- Batidas das pausas (dia a dia)
    S = "Pausas térmicas"
    pdias = [r for r in R if r["tipo"] == TIPO_PAU_DIA]
    if pdias:
        k_b = max(ciclo(r["dia"]) for r in pdias)
        dias_b = [r for r in pdias if ciclo(r["dia"]) == k_b]
        an = {id(r): analisa_batidas(r["marc"]) for r in dias_b}
        n_ac = sum(1 for r in dias_b if an[id(r)]["acima"])
        txt_r = J["normal"].get("pauBatResumo", {}).get("texto", "")
        item(S, "Batidas: resumo do mês", "Dias trabalhados, pessoas e dias com trecho acima de 1h40",
             "Uma linha por pessoa e dia com as batidas do relógio. Entre batidas, os trechos alternam trabalho e pausa "
             "(1ª→2ª trabalho, 2ª→3ª pausa...). Trecho de trabalho acima de 1h40 (100 min) = acima do limite.",
             "%d dias trabalhados de %d pessoa(s) · %d dia(s) com algum trecho acima de 1h40" % (len(dias_b), len({r["matricula"] or r["colaborador"] for r in dias_b}), n_ac),
             txt_r.split(" (")[0])
        hrs = defaultdict(lambda: [0, 0])
        for r in dias_b:
            for ini, dur in an[id(r)]["trab"]:
                h = (ini % 1440) // 60; hrs[h][0] += 1; hrs[h][1] += dur > 100
        esp = {"%dh" % h: "%d de %d" % (v[1], v[0]) for h, v in hrs.items() if v[0] >= 10}
        pain = {}
        for tp in J["normal"].get("chartPauHora", {}).get("tips", []):
            m = re.match(r"^(\d+h): (\d+) de (\d+)", tp)
            if m: pain[m.group(1)] = "%s de %s" % (m.group(2), m.group(3))
        item(S, "Batidas: por hora de começo do trecho", "Trechos acima de 1h40 / trechos que começaram em cada hora",
             "Cada trecho de trabalho é contado na hora em que começou; só aparecem horas com 10 ou mais trechos.", esp, pain)
        dsem = defaultdict(lambda: [0, 0])
        for r in dias_b:
            x = dsem[DIAS[r["dia"].weekday()]]; x[0] += 1; x[1] += bool(an[id(r)]["acima"])
        esp = {k: "%d de %d" % (v[1], v[0]) for k, v in dsem.items()}
        pain = {}
        for tp in J["normal"].get("chartPauDiaSemana", {}).get("tips", []):
            m = re.match(r"^(\w+): (\d+) de (\d+)", tp)
            if m and m.group(3) != "0": pain[m.group(1)] = "%s de %s" % (m.group(2), m.group(3))
        item(S, "Batidas: por dia da semana", "Dias com trecho acima de 1h40 / dias trabalhados",
             "Dias trabalhados de cada dia da semana (data do turno) e quantos tiveram algum trecho acima de 1h40.", esp, pain)
        fx = [("Até 1h30", 90), ("1h31 a 1h40", 100), ("1h41 a 2h", 120), ("2h01 a 3h", 180), ("Mais de 3h", 10 ** 9)]
        cnt = Counter()
        for r in dias_b:
            p1 = an[id(r)]["primeira"]
            if p1 is None: continue
            cnt[next(l for l, mx in fx if p1 <= mx)] += 1
        esp = {l: cnt[l] for l, _ in fx}
        pain = {}
        for tp in J["normal"].get("chartPauPrimeira", {}).get("tips", []):
            m = re.match(r"^(.*?) de trabalho até a 1ª pausa: (\d+)", tp)
            if m: pain[m.group(1)] = int(m.group(2))
        item(S, "Batidas: trabalho até a 1ª pausa térmica", "Dias em cada faixa",
             "Soma dos trechos de trabalho do começo do dia até a 1ª saída para pausa térmica (letra F), em faixas.", esp, pain)
        mp = {}
        for r in dias_b:
            a = an[id(r)]
            if not a["acima"]: continue
            x = mp.setdefault(r["colaborador"], [0, 0])
            x[0] = max(x[0], max(d for _, d in a["trab"])); x[1] += 1
        pain = {}
        for tp in (J["telaCheia"].get("chartPauMaior") or J["normal"].get("chartPauMaior", {})).get("tips", []):
            m = re.match(r"^(.*?): maior trecho sem pausa de (\d+)h(\d+) .*?· (\d+) dia", tp)
            if m: pain[m.group(1)] = (int(m.group(2)) * 60 + int(m.group(3)), int(m.group(4)))
        esp = {k: tuple(mp.get(k, (None, None))) for k in pain}
        item(S, "Batidas: maior trecho sem pausa, por pessoa", "Maior trecho (minutos) e dias acima de 1h40",
             "Por pessoa, nos dias com algum trecho acima de 1h40: o maior trecho de trabalho seguido e quantos dias.", esp, pain,
             esp == pain and len(pain) == min(len(mp), 20) or (esp == pain and len(pain) == len(mp)))
        t = tabela(J, "pauHeat")
        pain = {row[0]: (int_br(row[-2]) or 0, int_br(row[-1]) or 0) for row in t["rows"]}   # "–" = 0
        pr = defaultdict(lambda: [0, 0])
        for r in dias_b:
            a = an[id(r)]; x = pr[(r["colaborador"], r["matricula"])]; x[1] += 1
            x[0] += bool(a["acima"] or a["sem"] or a["impar"])
        por_nome = defaultdict(list)
        for (nm, mt), v in pr.items(): por_nome[nm].append(tuple(v))
        esp = {k: (por_nome[k][0] if len(por_nome[k]) == 1 else por_nome[k]) for k in pain}
        item(S, "Batidas: mapa das pausas", "Dias com problema e dias no relatório, por pessoa",
             "Dia com problema = algum trecho acima de 1h40, ou mais de 1h40 de trabalho sem nenhuma pausa térmica, ou "
             "batida faltando (número ímpar de batidas).", esp, pain, esp == pain and len(pain) == len(pr),
             obs="Uma linha por pessoa (pela matrícula).")

    # ---------------------------------------------------------- outros gráficos de pausas, PontoNet e afastamentos
    if res:
        pain = tips(J, "chartPauGrupo")
        esp = {g: int(sum(r["pc"] + r["pm"] + r["pM"] for r in atual if r["gestor"] == g)) for g in pain}
        item("Pausas térmicas", "Pausas por gestor", "Total de pausas da equipe (certas + curtas + longas)",
             "Mês mais recente: soma das pausas de cada pessoa da equipe.", esp, pain)
        pain = {}
        for tp in (J["telaCheia"].get("chartPauAcima") or J["normal"].get("chartPauAcima", {})).get("tips", []):
            m = re.match(r"^(.*?): (\d+) de (\d+) trechos", tp)
            if m: pain[m.group(1)] = (int(m.group(2)), int(m.group(3)))
        esp = {r["colaborador"]: (int(r["tM"]), int(r["tc"] + r["tM"] + r["tm"])) for r in atual if r["colaborador"] in pain}
        item("Pausas térmicas", "Acima de 1h40 por pessoa", "Trechos acima de 1h40 / todos os trechos",
             "Do relatório de pausas do mês: trechos de trabalho acima de 1h40 ÷ (trechos certos + acima + abaixo).", esp, pain)
    if pn:
        pain = tips(J, "chartPnTipo")
        tp_c = Counter(nome_tipo_pn(r["tipo_just"]) for r in cic)
        tp_bruto = Counter(r["tipo_just"] for r in cic)
        item("PontoNet", "Por motivo", "Aprovados de cada motivo (ciclo escolhido)",
             "Pedidos aprovados do ciclo contados pelo motivo informado no PontoNet.",
             dict(sorted(tp_c.items(), key=lambda x: -x[1])), dict(sorted(pain.items(), key=lambda x: -x[1])),
             tp_c == Counter(pain),
             obs="Motivos como vêm do sistema: " + "; ".join("%s = %d" % kv for kv in tp_bruto.most_common()) +
                 ". O painel junta os que falam de esquecimento (inclusive \"esquecimento crachá\") em \"Esquecimento\".")
        pain = {}
        for tp in J["normal"].get("chartPnGestor", {}).get("tips", []):
            m = re.match(r"^(.*?): em média ([\d,]+) dias? para avaliar (\d+)", tp)
            if m: pain[m.group(1)] = (m.group(2), int(m.group(3)))
        esp = {}
        for g in pain:
            l = [r["dias_gest"] for r in cic if r["gestor"] == g and r["dias_gest"] is not None]
            esp[g] = (fmt_dec(sum(l) / len(l)).replace(",0", "") if l else "0", len(l))
        item("PontoNet", "Tempo do gestor para avaliar, por gestor", "Média de dias e pedidos",
             "Média dos dias que o gestor levou para avaliar (contagem do PontoNet), nos aprovados do ciclo.", esp, pain)
    S = "Afastamentos"
    leg = {}
    for x in J["normal"].get("chartAbsGrupo", {}).get("legendaMini", []):
        m = re.match(r"^(.*) ([\d.]+)$", x)
        if m: leg[m.group(1)] = int(m.group(2).replace(".", ""))
    ciclos_tela = {r[0].split(" ")[0] for r in tabela(J, "tableAbsEvolucao")["rows"]}
    dias_vis = {c: v for c, v in dias_af.items() if rotulo_ciclo(c) in ciclos_tela}
    dias_tot = sum(len(v) for v in dias_vis.values())
    item(S, "Por gestor (por tipo)", "Soma dos dias de todos os tipos = dias do período",
         "A legenda soma os dias de cada grupo de afastamento; o total tem que ser igual à soma dos dias dos ciclos mostrados "
         "(os 3 últimos ciclos até o fim do cartão ponto; férias já marcadas para depois ficam de fora).",
         dias_tot, sum(leg.values()))
    pain = tips(J, "chartAbsGrupo")
    esp = Counter(r["gestor"] for c in dias_vis.values() for r, _ in c)
    compara_dict(S, "Por gestor (por tipo)", "Dias de afastamento por gestor",
                 "Dias de afastamento (todos os tipos, 3 meses) somados pelo gestor da pessoa.", esp, pain, so_chaves_do_painel=True)

    # ---------------------------------------------------------- Tendências
    S = "Tendências"
    t = tabela(J, "tableTendUnidade")
    cab = [h.title() for h in t["head"][2:5]]
    cks = []
    for h in cab:
        mm = MESES.index(h.split("/")[0][:3].title()) + 1
        cks.append("20%s-%02d" % (h.split("/")[1], mm))
    def cnt(lst, f=lambda r: r["dia"]):
        c = Counter(ciclo(f(r)) for r in lst); return tuple(c[k] for k in cks)
    pau_c = lambda campo, den: tuple((fmt_dec(100 * sum(campo(r) for r in res if ciclo(r["dia"]) == k) /
                                               max(1, sum(den(r) for r in res if ciclo(r["dia"]) == k)))).replace(",0", "") + "%"
                                     if sum(den(r) for r in res if ciclo(r["dia"]) == k) else "–" for k in cks)
    dias_at = [d for r in at for d in [r["dia"] + timedelta(days=i) for i in range(max(1, int(round(r["dias_aus"] or 1))))]]
    gm = lambda k: [r["dias_gest"] for r in pn if ciclo(r["dia"]) == k and r["dias_gest"] is not None]
    esp_t = {
        "Batidas digitadas": cnt(dig), "Batidas digitadas por esquecimento": cnt(esq),
        "Dias de atestado": tuple(Counter(ciclo(d) for d in dias_at)[k] for k in cks),
        "Horas excedentes (dias com mais de 10h)": cnt([r for r in jor if r["tipo"] == "Horas Excedentes"]),
        "Intrajornada (intervalo de almoço curto)": cnt([r for r in jor if r["tipo"] == "Intrajornada"]),
        "Interjornada semanal (trabalho na folga ou feriado)": cnt([r for r in jor if r["tipo"] == "Interjornada Semanal"]),
        "Interjornada (menos de 11h de descanso)": cnt([r for r in jor if r["tipo"] == "Interjornada"]),
        "Pausas certas (20 min)": pau_c(lambda r: r["pc"], lambda r: r["pc"] + r["pm"] + r["pM"]),
        "Trabalho acima de 1h40 sem pausa": pau_c(lambda r: r["tM"], lambda r: r["tc"] + r["tM"] + r["tm"]),
        "Justificativas aprovadas no PontoNet": cnt(pn),
        "Dias para o gestor avaliar (média)": tuple((fmt_dec(sum(gm(k)) / len(gm(k))).replace(",0", "") + " dias") if gm(k) else "–" for k in cks),
        "Batida de pausa faltando": tuple(int(sum(r["imp"] for r in res if ciclo(r["dia"]) == k)) for k in cks),
    }
    for row in t["rows"]:
        nome = row[0]
        val = tuple(row[2:5])
        e = esp_t.get(nome)
        if e is None: continue
        e_txt = tuple(str(x) if not isinstance(x, str) else x for x in e)
        item(S, "Assuntos da unidade (3 meses)", nome,
             "Mesmo cálculo da seção de origem, por ciclo 28 a 27 (%s)." % ", ".join(rotulo_ciclo(k) for k in cks),
             e_txt, tuple(v.replace(".", "") if re.match(r"^[\d.]+$", v) else v for v in val))

    # ---------------------------------------------------------- Banco de horas: começo + extras − faltas = fim
    if bd:
        kp = numeros(J, "bhKpis")
        if "BANCO NO COMEÇO" in kp and "BANCO NO FIM" in kp:
            ini = minutos_de(kp["BANCO NO COMEÇO"][0]); fim = minutos_de(kp["BANCO NO FIM"][0])
            item("Banco de horas", "Números do topo", "Banco no fim = banco no começo + horas extras − horas falta",
                 "A conta do saldo tem que fechar: o que a equipe tinha no começo, mais o que fez de extra, menos o que faltou.",
                 dur_txt(ini + sum(r["he"] for r in sel) - sum(r["bh"] for r in sel)), dur_txt(fim),
                 ini + sum(r["he"] for r in sel) - sum(r["bh"] for r in sel) == fim)

    # ---------------------------------------------------------- Banco de horas: conferência cartão ponto x relatório
    if bd and [r for r in R if r["tipo"] == TIPO_BH]:
        t = tabela(J, "tableBHConf")
        if t["rows"]:
            rel = {}
            for r in R:
                if r["tipo"] == TIPO_BH and ciclo(r["dia"]) in ks:
                    if r["colaborador"] not in rel or r["dia"] > rel[r["colaborador"]]["dia"]: rel[r["colaborador"]] = r
            calc = defaultdict(lambda: [0, 0])
            for r in sel:
                calc[r["colaborador"]][0] += r["he"]; calc[r["colaborador"]][1] += r["bh"]
            pain, esp = {}, {}
            for row in t["rows"]:
                nm = row[0]
                pain[nm] = (minutos_de(row[2]) if row[2] not in ("–", "") else None, minutos_de(row[4]) if row[4] not in ("–", "") else None,
                            minutos_de(row[6]) if row[6] not in ("–", "") else None,
                            minutos_de(row[7]) if row[7] not in ("–", "") else None,
                            "bate" if "diferen" not in row[8] and "fora" not in row[8].lower() and "não está" not in row[8] else "não bate")
                rr = rel.get(nm)
                c = calc.get(nm, [0, 0] if rr else None)   # no relatório e sem nenhum dia no cartão = banco 0
                saldo_rel = int(rr["bh_ant"] + rr["bh_cred"] - rr["bh_deb"]) if rr else None
                saldo_calc = int(c[0] - c[1]) if c else None
                ok_p = rr is not None and c is not None and abs(saldo_rel - saldo_calc) <= 1
                esp[nm] = (int(c[0]) if c else None, int(c[1]) if c else None, saldo_calc, saldo_rel, "bate" if ok_p else "não bate")
            # o painel mostra "–" onde a pessoa não tem a informação; compara só o que tem número
            iguais = all(all(a == b or b is None for a, b in zip(esp[k][:4], pain[k][:4])) and esp[k][4] == pain[k][4] for k in pain)
            n_bate = sum(1 for v in esp.values() if v[4] == "bate")
            item("Banco de horas", "Conferência: cartão ponto × relatório de banco de horas",
                 "Banco de cada pessoa pelas duas contas e se bate (lista inteira)",
                 "Banco calculado = soma das horas extras − soma das horas falta do cartão ponto, dia a dia, desde o começo do "
                 "período (zerado). Banco no relatório = saldo anterior + crédito − débito do relatório de banco de horas do "
                 "sistema. Bate quando a diferença é de até 1 minuto.",
                 "%d pessoas na tabela; %d batem, %d não batem" % (len(esp), n_bate, len(esp) - n_bate),
                 "%d pessoas na tabela; %d batem, %d não batem" % (len(pain), sum(1 for v in pain.values() if v[4] == "bate"),
                                                                     sum(1 for v in pain.values() if v[4] != "bate")),
                 iguais and len(esp) == len(pain),
                 obs="Quando não bate, a diferença está nos dados de origem (cartão ponto × relatório do sistema), não numa conta do painel: o RH confere essas pessoas no sistema.")

    # ---------------------------------------------------------- Lista de colaboradores
    S = "Lista de colaboradores"
    t = tabela(J, "tableOrg")
    item(S, "Lista de colaboradores", "Pessoas na lista",
         "Uma linha por pessoa da lista de equipes (matrícula repetida conta uma vez).", len(org), len(t["rows"]))
    pain = {(row[0] if row[0] != "–" else "") + "|" + row[1].replace("sem matrícula", "").strip(): tuple(row[2:5]) for row in t["rows"]}
    esp = {(r["matricula"]) + "|" + r["colaborador"]: (r["situacao"], "" if r["cargo"] == "Sem cargo" else r["cargo"],
           "Sem gestor" if r["gestor"] == "Sem gestor" else r["gestor"]) for r in org.values()}
    item(S, "Lista de colaboradores", "Matrícula, nome, área, cargo e gestor de cada pessoa",
         "As colunas são as da lista de equipes, sem nenhuma conta (desde 11/11 a lista não mostra ocorrências, digitadas "
         "nem saldo de horas).", esp, pain)

    # ---------------------------------------------------------- erros de JavaScript
    item("Painel inteiro", "Funcionamento", "Erros de programa ao abrir e desenhar tudo",
         "O navegador registra qualquer erro de JavaScript enquanto o painel abre o arquivo e desenha todas as seções.",
         0, len(J.get("erros") or []))


def main():
    args = sys.argv[1:]
    saida_json = None
    if "--json" in args:
        i = args.index("--json"); saida_json = args[i + 1]; del args[i:i + 2]
    J = json.load(open(args[0], encoding="utf-8"))
    R = ler_varios(args[1:])
    conferir(J, R)
    secao = None
    for r in RES:
        if r["secao"] != secao:
            secao = r["secao"]; print("\n== " + secao)
        print(("  OK    " if r["ok"] else "  FALHA ") + r["card"] + " — " + r["item"])
        if not r["ok"]:
            print("        esperado: " + r["esperado"][:600]); print("        painel:   " + r["painel"][:600])
    n_ok = sum(r["ok"] for r in RES)
    print("\nTotal: %d de %d itens conferem." % (n_ok, len(RES)))
    if saida_json: json.dump(RES, open(saida_json, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    sys.exit(0 if n_ok == len(RES) else 1)


if __name__ == "__main__":
    main()
