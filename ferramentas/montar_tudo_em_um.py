"""Monta o módulo único TudoEmUm/AtualizarPainelPonto.bas (e txt/AtualizarPainelPonto.txt).

Junta o orquestrador (TudoEmUm/Orquestrador.bas) com as macros do fluxo, sem mudar os .bas originais:
  - tira "Attribute VB_Name" e "Option Explicit" de cada parte (o módulo final tem um só);
  - deixa todos os Sub/Function Private (só AtualizarPainelPonto aparece no Alt+F8);
  - renomeia os nomes repetidos entre FormatarMarcacoesDigitadas e FormatarAfastamentos;
  - troca MsgBox por Avisar/Perguntar (no lote, as mensagens vão para o resumo final);
  - faz o GerarAbaCSV usar as abas recém-importadas (LoteAbasCartao etc.) e a pasta já escolhida.
Cada troca pontual é conferida (tem que acontecer exatamente uma vez), para o script falhar
alto se um .bas mudar.

Uso: python3 ferramentas/montar_tudo_em_um.py   (na raiz do repositório)
"""
import re
import sys

ORQ = "TudoEmUm/Orquestrador.bas"
PARTES = [
    ("CartaoPonto/ConsolidarCartaoPonto.bas", {}),
    ("PausasTermicas/FormatarPausasTermicas.bas", {}),
    ("MarcacoesDigitadas/FormatarMarcacoesDigitadas.bas", {
        "NOME_ABA_SAIDA": "NOME_ABA_DIGITADAS", "Celula": "CelulaDig"}),
    ("Afastamentos/FormatarAfastamentos.bas", {
        "NOME_ABA_SAIDA": "NOME_ABA_AFASTAMENTOS", "Celula": "CelulaAfa", "TextoCel": "TextoCelAfa",
        "TextoSimples": "TextoSimplesAfa", "CodigoTexto": "CodigoTextoAfa"}),
    ("GerarCSVPonto/GerarCSVPonto.bas", {}),
]
TROCAS_GERARCSV = [
    ("AbasMaisRecentes(EncontrarAbasCartao(wbTrat), MAXIMO_MESES_AUDITORIA)",
     "LoteAbasCartaoOrdenadas(wbTrat, MAXIMO_MESES_AUDITORIA)"),
    ("Set wsAfast = EncontrarAbaAfastamentos(wbTrat)", "Set wsAfast = LoteAbaAfastamentos(wbTrat)"),
    ("Set wsDigitadas = EncontrarAbaDigitadas(wbTrat)", "Set wsDigitadas = LoteAbaDigitadas(wbTrat)"),
    ("Set wsPausas = EncontrarTabelaPausas(wbTrat, linhaCabecalhoPausas)",
     "Set wsPausas = LoteTabelaPausas(wbTrat, linhaCabecalhoPausas)"),
    ("    pasta = EscolherPasta()\n",
     "    If mPastaSaida <> \"\" Then pasta = mPastaSaida Else pasta = EscolherPasta()\n"),
]


def partes_codigo(linha):
    """Divide a linha em trechos (é_código, texto): strings "..." e comentários ' não são código."""
    out, i, buf = [], 0, ""
    while i < len(linha):
        c = linha[i]
        if c == '"':
            if buf:
                out.append((True, buf))
                buf = ""
            j = i + 1
            while j < len(linha):
                if linha[j] == '"':
                    if j + 1 < len(linha) and linha[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            out.append((False, linha[i:j + 1]))
            i = j + 1
            continue
        if c == "'" or (linha[i:i + 4].lower() == "rem " and not buf.strip()):
            if buf:
                out.append((True, buf))
            out.append((False, linha[i:]))
            return out
        buf += c
        i += 1
    if buf:
        out.append((True, buf))
    return out


def no_codigo(texto, fn):
    linhas = []
    for linha in texto.split("\n"):
        linhas.append("".join(fn(t) if cod else t for cod, t in partes_codigo(linha)))
    return "\n".join(linhas)


def preparar(caminho, renomear):
    s = open(caminho, encoding="utf-8").read().replace("\r\n", "\n")
    s = re.sub(r"^Attribute VB_Name.*\n", "", s, flags=re.M)
    s = re.sub(r"^Option Explicit\s*\n", "", s, flags=re.M)
    # Sub/Function públicos (ou sem modificador) -> Private
    s = re.sub(r"^(?:Public\s+)?(Sub|Function)\b", r"Private \1", s, flags=re.M)
    s = re.sub(r"^Private Private ", "Private ", s, flags=re.M)
    for velho, novo in renomear.items():
        s = no_codigo(s, lambda t, v=velho, n=novo: re.sub(r"\b%s\b" % re.escape(v), n, t, flags=re.I))

    def trocar_msgbox(t):
        t = re.sub(r"^(\s*)MsgBox\b", r"\1Avisar", t)
        t = re.sub(r"\bThen MsgBox\b", "Then Avisar", t)
        t = re.sub(r"\bElse MsgBox\b", "Else Avisar", t)
        return re.sub(r"\bMsgBox\s*\(", "Perguntar(", t)
    s = no_codigo(s, trocar_msgbox)
    if caminho.endswith("GerarCSVPonto.bas"):
        for velho, novo in TROCAS_GERARCSV:
            n = s.count(velho)
            if n != 1:
                sys.exit("ERRO: '%s' aparece %d vez(es) em %s (esperado 1)" % (velho.strip(), n, caminho))
            s = s.replace(velho, novo)
    restante = [l for l in s.split("\n") if any(cod and re.search(r"\bMsgBox\b", t) for cod, t in partes_codigo(l))]
    if restante:
        sys.exit("ERRO: MsgBox não convertido em %s:\n%s" % (caminho, "\n".join(restante)))
    return s.strip("\n")


def main():
    orq = open(ORQ, encoding="utf-8").read().replace("\r\n", "\n").strip("\n")
    blocos = [orq]
    for caminho, ren in PARTES:
        sep = "' " + "=" * 110 + "\n' PARTE: " + caminho + " (gerado por ferramentas/montar_tudo_em_um.py)\n' " + "=" * 110
        blocos.append(sep + "\n\n" + preparar(caminho, ren))
    corpo = "Option Explicit\n\n" + "\n\n".join(blocos) + "\n"
    publicos = re.findall(r"^Public (?:Sub|Function) (\w+)", corpo, flags=re.M)
    if publicos != ["AtualizarPainelPonto"]:
        sys.exit("ERRO: procedimentos públicos inesperados: %s" % publicos)
    nomes = re.findall(r"^(?:Public |Private )?(?:Sub|Function|Const)\s+(\w+)", corpo, flags=re.M)
    vistos, dup = set(), set()
    for n in nomes:
        (dup if n.lower() in vistos else vistos).add(n.lower())
    if dup:
        sys.exit("ERRO: nomes repetidos no módulo: %s" % sorted(dup))
    open("TudoEmUm/AtualizarPainelPonto.bas", "w", encoding="utf-8", newline="\n").write(
        'Attribute VB_Name = "AtualizarPainelPonto"\n' + corpo)
    # .txt para colar: sem a linha Attribute (no editor ela vira erro de sintaxe), CRLF
    open("txt/AtualizarPainelPonto.txt", "w", encoding="utf-8", newline="\r\n").write(corpo)
    print("OK: TudoEmUm/AtualizarPainelPonto.bas e txt/AtualizarPainelPonto.txt (%d linhas)" % corpo.count("\n"))


if __name__ == "__main__":
    main()
