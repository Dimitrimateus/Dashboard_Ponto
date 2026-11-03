"""Reproduz em Python as macros de Jornada e o cálculo das Pausas Térmicas por ciclo.

  - FormatarExcedentes / FormatarIntrajornada / FormatarInterjornada / FormatarInterjornadaSemanal
    (JornadaCLT/FormatarJornada.bas): cada relatório vira a tabela
    Matrícula | Colaborador | Cargo | Local | Data | Ocorrência | Minutos | Detalhe
  - FormatarPausasTermicas (PausasTermicas/FormatarPausasTermicas.bas): totais por colaborador e
    por ciclo (28 a 27), calculados das linhas de cada dia, e a conferência com o "Total do
    Colaborador" do próprio relatório (período inteiro).

Uso: python3 ferramentas/formatar_jornada.py <pasta com os relatórios>   (mostra o resumo)
"""
import datetime
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from formatar_marcacoes import ler_planilha  # noqa: E402

CAB = ["Matrícula", "Colaborador", "Cargo", "Local", "Data", "Ocorrência", "Minutos", "Detalhe"]
BASE = datetime.date(1899, 12, 30)


def data(serial):
    return BASE + datetime.timedelta(days=int(serial))


def hm(frac_ou_min, minutos=False):
    m = int(round(frac_ou_min if minutos else frac_ou_min * 1440))
    return "%d:%02d" % (m // 60, m % 60)


def hora(frac):
    m = int(round(frac * 1440)) % 1440
    return "%02d:%02d" % (m // 60, m % 60)


def num(v):
    return isinstance(v, float)


def txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def sem_codigo(local):
    """'2.028.30.304.3041 - Ativ. Operacional 1ºT' -> 'Ativ. Operacional 1ºT'"""
    return re.sub(r"^[\d.]+\s*-\s*", "", local.strip())


def achar(linha, rotulo):
    for c, v in sorted(linha.items()):
        if isinstance(v, str) and v.strip().lower() == rotulo.lower():
            return c
    return 0


def linha_cabecalho(L, *rotulos):
    for r in sorted(L):
        if all(achar(L[r], x) for x in rotulos):
            return r
    return 0


# --------------------------------------------------------------------------- Excedentes
def excedentes(L):
    h = L[linha_cabecalho(L, "Data", "Colaborador/Nome", "Marcações")]
    cData, cMat = achar(h, "Data"), achar(h, "Colaborador/Nome")
    cCargo, cLocal = achar(h, "Cargo"), achar(h, "Local")
    cCarga, cMarc = achar(h, "Carga Hor.") + 1, achar(h, "Marcações") + 1
    out = []
    for r in sorted(L):
        d = L[r]
        if not (num(d.get(cData)) and num(d.get(cMat)) and isinstance(d.get(cMat + 1), str)):
            continue
        nums = [d[c] for c in sorted(d) if c >= cMarc and num(d[c])]
        if len(nums) < 4:
            continue
        extras, horas, cvnot, total = nums[-4:]
        marc = nums[:-4]
        det = "Jornada %s (%s trabalhadas + %s de adicional noturno) · carga %s · extras %s · marcações %s" % (
            hm(total), hm(horas), hm(cvnot), hm(d.get(cCarga) or 0), hm(extras), " ".join(hora(x) for x in marc))
        out.append([txt(d[cMat]), txt(d[cMat + 1]), txt(d.get(cCargo)), txt(d.get(cLocal)), data(d[cData]),
                    "Horas excedentes", int(round(total * 1440)), det])
    return out


# --------------------------------------------------------------------------- Intrajornada
def intrajornada(L):
    h = L[linha_cabecalho(L, "Cadastro", "Registros de Ponto")]
    cMat, cNome, cLocal = achar(h, "Cadastro"), achar(h, "Nome"), achar(h, "Local")
    cCargo, cData, cDia, cRP = achar(h, "Cargo"), achar(h, "Data"), achar(h, "Dia"), achar(h, "Registros de Ponto")
    out = []
    for r in sorted(L):
        d = L[r]
        if not (num(d.get(cMat)) and isinstance(d.get(cNome), str) and num(d.get(cData))):
            continue
        marc = [d[c] for c in range(cRP, cRP + 8) if num(d.get(c))]
        trab = d.get(cRP + 8) or 0
        ints = [d[c] for c in range(cRP + 9, cRP + 13) if num(d.get(c))]
        soma = sum(ints)
        det = "Trabalhou %s %s · marcações %s · horário %s" % (
            hm(trab), ("com intervalo(s) de " + " + ".join(hm(x) for x in ints)) if ints else "sem intervalo",
            " ".join(hora(x) for x in marc), txt(d.get(cDia)))
        out.append([txt(d[cMat]), txt(d[cNome]), txt(d.get(cCargo)), txt(d.get(cLocal)), data(d[cData]),
                    "Intrajornada", int(round(soma * 1440)), det])
    return out


# --------------------------------------------------------------------------- Interjornada
def interjornada(L):
    rh = linha_cabecalho(L, "Data Apuração", "Horas Descansadas")
    h = L[rh]
    cApur = achar(h, "Data Apuração") + 1
    cAnt, cAtu = achar(h, "Marcação Anterior") + 1, achar(h, "Marcação Atual") + 1
    cDesc, cOc = achar(h, "Horas Descansadas") + 1, achar(h, "Ocorrência") + 1
    cLocal, cCargo = achar(h, "Local"), achar(h, "Cargo")
    out = []
    for r in sorted(L):
        d = L[r]
        if r <= rh or not (num(d.get(1)) and isinstance(d.get(2), str) and num(d.get(cApur))):
            continue
        desc = d.get(cDesc) or 0
        det = "Descansou %s entre %s %s e %s %s · %s" % (
            hm(desc), data(d[cAnt]).strftime("%d/%m"), hora(d.get(cAnt + 1) or 0),
            data(d[cAtu]).strftime("%d/%m"), hora(d.get(cAtu + 1) or 0), txt(d.get(cOc)))
        out.append([txt(d[1]), txt(d[2]), txt(d.get(cCargo)), txt(d.get(cLocal)), data(d[cApur]),
                    "Interjornada", int(round(desc * 1440)), det])
    return out


# --------------------------------------------------------------------------- Interjornada semanal
def interjornada_semanal(L):
    h = L[linha_cabecalho(L, "Colaborador", "Horas DSR")]
    cLocal, cTxt = achar(h, "Local"), achar(h, "Horas DSR")
    out, mat, nome = [], "", ""
    for r in sorted(L):
        d = L[r]
        t = txt(d.get(cTxt))
        m = re.search(r"\((\d{2})/(\d{2})/(\d{4})\)", t)
        if not m:
            continue
        if num(d.get(1)):
            mat, nome = txt(d[1]), txt(d.get(2))
        extra = " · ".join(txt(d[c]) if not num(d[c]) else hora(d[c]) for c in sorted(d) if c > cTxt and d[c] not in (None, ""))
        dia = datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        out.append([mat, nome, "", sem_codigo(txt(d.get(cLocal))), dia, "Interjornada semanal", 0,
                    t + ((" · " + extra) if extra else "")])
    return out


# --------------------------------------------------------------------------- Pausas Térmicas
def fim_do_ciclo(dia):
    y, m = dia.year, dia.month
    if dia.day >= 28:
        m += 1
        if m > 12:
            m, y = 1, y + 1
    return datetime.date(y, m, 27)


def classificar_dia(marc):
    """marc = [(fração do dia, 'P'|'F'|'G'), ...] -> [pausa ok, pausa <20, pausa >20, trab =1:40, >1:40, <1:40, ímpar]
    Regra (conferida com o "Total do Colaborador" do HRES114): as batidas se alternam saída/volta
    pela posição; trechos de posição par (1ª→2ª, 3ª→4ª...) são trabalho; os de posição ímpar são
    intervalo, e só contam como pausa térmica quando começam e terminam numa batida "F". Pausa de
    20 min ou mais = correta (o relatório nunca conta "pausa maior"). Trabalho de 1:40 exato =
    correto. Dia com número ímpar de batidas conta 1 "marcação ímpar" e, além disso, o último
    trecho do dia é contado como trabalho."""
    t = [0] * 7
    mins, base, ant = [], 0, None
    for f, tp in marc:
        x = int(round(f * 1440))
        if ant is not None and x + base < ant:
            base += 1440
        ant = x + base
        mins.append((x + base, tp))
    n = len(mins)
    impar = n % 2 == 1
    if impar:
        t[6] = 1
    for k in range(n - 1):
        dur = mins[k + 1][0] - mins[k][0]
        if k % 2 == 0 or (impar and k == n - 2):
            t[3 if dur == 100 else 4 if dur > 100 else 5] += 1
        elif mins[k][1] == "F" and mins[k + 1][1] == "F":
            t[0 if dur >= 20 else 1] += 1
    return t


def pausas(L):
    """Devolve (linhas por ciclo, conferência). linhas: [mat, nome, cargo, fim do ciclo, 7 números]."""
    por_ciclo, conf, cur = {}, [], None
    for r in sorted(L):
        d = L[r]
        a, b = d.get(1), d.get(2)
        if num(a) and isinstance(b, str) and b.strip() not in ("", "-") and d.get(3) not in ("P", "F", "G"):
            cur = {"mat": txt(a), "nome": b.strip(), "cargo": txt(d.get(6)), "soma": [0] * 7}
        elif num(a) and d.get(3) in ("P", "F", "G") and cur:
            marc, k = [], 2
            while num(d.get(k)) and d.get(k + 1) in ("P", "F", "G"):
                marc.append((d[k], d[k + 1]))
                k += 2
            t = classificar_dia(marc)
            ch = (cur["mat"], fim_do_ciclo(data(a)))
            if ch not in por_ciclo:
                por_ciclo[ch] = [cur["mat"], cur["nome"], cur["cargo"], ch[1]] + [0] * 7
            for i in range(7):
                por_ciclo[ch][4 + i] += t[i]
                cur["soma"][i] += t[i]
        elif isinstance(a, str) and a.strip() == "Total do Colaborador" and cur:
            v = L.get(r + 2, {})
            sist = [int(v.get(c) or 0) for c in (1, 2, 4, 6, 8, 10, 11)]
            conf.append([cur["mat"], cur["nome"], cur["cargo"]] + sist + [cur["soma"] == sist])
    return sorted(por_ciclo.values(), key=lambda x: (x[1], x[3])), conf


ARQUIVOS = {"Horas excedentes": ("Excedentes 3 meses.xlsx", excedentes),
            "Intrajornada": ("Intrajornada 3 meses.xlsx", intrajornada),
            "Interjornada": ("Interjornada 3 meses.xlsx", interjornada),
            "Interjornada semanal": ("Interjornada semanal 3 meses.xlsx", interjornada_semanal)}


def achar_arquivo(pasta, nome):
    alvo = nome.lower()
    for f in os.listdir(pasta):
        if f.lower() == alvo or f.lower().replace("#u00e9", "é").replace("#u00e3", "ã") == alvo:
            return os.path.join(pasta, f)
    raise FileNotFoundError(nome)


if __name__ == "__main__":
    pasta = sys.argv[1]
    for tipo, (arq, fn) in ARQUIVOS.items():
        linhas = fn(ler_planilha(achar_arquivo(pasta, arq)))
        print("%-22s %3d linhas, %3d pessoas" % (tipo, len(linhas), len({x[0] for x in linhas})))
    ciclos, conf = pausas(ler_planilha(achar_arquivo(pasta, "Pausas térmicas 3 meses.xlsx")))
    ok_p = sum(1 for c in conf if c[3:6] == [sum(x[4 + i] for x in ciclos if x[0] == c[0]) for i in range(3)])
    print("Pausas: %d linhas (colaborador x ciclo), %d colaboradores; pausas iguais ao sistema em %d; "
          "todos os 7 números iguais em %d" % (len(ciclos), len(conf), ok_p, sum(1 for c in conf if c[-1])))
