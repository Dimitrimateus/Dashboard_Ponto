"""Monta o MÓDULO ÚNICO (pedido do Dimitri, 03/11: "quero colar um só código").

Junta todas as macros (ConsolidarCartaoPonto, FormatarPausasTermicas, FormatarMarcacoesDigitadas,
FormatarAfastamentos, FormatarJornada, FormatarBancoHoras, GerarCSVPonto, PrepararAbasDoMes com a GerarCSVsDaPasta e
sem a PURO_Para_EDITADO, que é de outro assunto e fica num .txt à parte desde 04/11) num arquivo
só, sem mudar os .bas de origem:
  - tira "Attribute VB_Name" e "Option Explicit" de cada parte (o módulo final tem um só);
  - nomes de nível de módulo repetidos entre as partes (ex.: modoAuto, msgAuto, Aviso, helpers
    com o mesmo nome) ganham um sufixo da parte (modoAuto_Pau...), só fora de textos e comentários;
  - Application.Run("X_Auto") vira chamada direta (RodarMacroAuto): com o módulo único não
    depende de nome público e não fica "ambíguo" se os módulos antigos continuarem na planilha;
  - só ficam públicas (Alt+F8) as macras de usar: GerarCSVsDaPasta, PrepararAbasDoMes,
    (as do painel). O resto vira Private.
  - (04/11) as declarações do nível do módulo (variáveis Private/Public/Dim, Const, Type, Enum)
    de TODAS as partes vão para o TOPO do módulo: o VBA só aceita declaração antes da 1ª Sub/
    Function. Deixadas no meio (onde cada parte começava) davam "Variável não definida" ao rodar.
Cada troca pontual é conferida (o script para se não achar o texto esperado).

Uso (na raiz do repositório): python3 ferramentas/montar_modulo_unico.py
Saída: ModuloUnico/PainelPonto_TudoEmUm.bas e txt/PainelPonto_TudoEmUm.txt (CRLF, sem a linha
Attribute, pronto para colar). Rode de novo sempre que mudar um .bas.
"""
import re
import sys

PARTES = [  # (arquivo, sufixo para nomes repetidos)
    ("GerarCSVPonto/GerarCSVPonto.bas", "Csv"),
    ("PrepararAbas/PrepararAbasDoMes.bas", "Prep"),
    ("CartaoPonto/ConsolidarCartaoPonto.bas", "Car"),
    ("PausasTermicas/FormatarPausasTermicas.bas", "Pau"),
    ("MarcacoesDigitadas/FormatarMarcacoesDigitadas.bas", "Dig"),
    ("Afastamentos/FormatarAfastamentos.bas", "Afa"),
    ("Jornada/FormatarJornada.bas", "Jor"),
    ("BancoHoras/FormatarBancoHoras.bas", "BH"),
]
PUBLICAS = {"gerarcsvsdapasta", "prepararabasdomes"}
MACROS_AUTO = ["ConsolidarCartaoPontoDefinitivo_Auto", "FormatarPausasTermicas_Auto", "FormatarMarcacoesDigitadas_Auto",
               "FormatarAfastamentos_Auto", "FormatarJornada_Auto", "FormatarBancoHoras_Auto", "GerarAbaCSV_Auto"]


def partes_codigo(linha):
    """Divide a linha em (é_código, texto): strings "..." e comentários ' não são código."""
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
        if c == "'":
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
    return "\n".join("".join(fn(t) if cod else t for cod, t in partes_codigo(l)) for l in texto.split("\n"))


def nomes_de_modulo(s):
    """Nomes declarados no nível do módulo: Sub/Function/Property, Const, variáveis, Type, Enum."""
    nomes = set(re.findall(r"^(?:Public |Private |Friend )?(?:Static )?(?:Sub|Function|Property (?:Get|Let|Set))\s+(\w+)", s, re.M | re.I))
    nomes |= set(re.findall(r"^(?:Public |Private |Global )?Const\s+(\w+)", s, re.M | re.I))
    nomes |= set(re.findall(r"^(?:Private|Public|Dim|Global)\s+(?!Sub|Function|Const|Type|Enum|Property|Declare|Static)(\w+)\s+As\b", s, re.M | re.I))
    nomes |= set(re.findall(r"^(?:Public |Private )?(?:Type|Enum)\s+(\w+)", s, re.M | re.I))
    return nomes


def publicos(s):
    return {n.lower() for n in re.findall(r"^(?:Public\s+)?(?:Sub|Function)\s+(\w+)", s, re.M)
            if not re.match(r"^Private", n)}


def main():
    textos = []
    for caminho, suf in PARTES:
        s = open(caminho, encoding="utf-8").read().replace("\r\n", "\n")
        s = re.sub(r"^Attribute VB_Name.*\n", "", s, flags=re.M)
        s = re.sub(r"^Option Explicit\s*\n", "", s, flags=re.M)
        textos.append([caminho, suf, s])

    # nomes repetidos entre partes: a 1ª parte que declara fica com o nome; as outras ganham sufixo
    dono = {}
    renomear = {c: {} for c, _, _ in textos}
    for caminho, suf, s in textos:
        pub = {n.lower() for n in re.findall(r"^Public\s+(?:Sub|Function)\s+(\w+)", s, re.M)}
        for n in sorted(nomes_de_modulo(s)):
            k = n.lower()
            if k in dono:
                if k in pub:
                    sys.exit("ERRO: '%s' é público em %s e repetido em %s" % (n, caminho, dono[k]))
                renomear[caminho][n] = n + "_" + suf
            else:
                dono[k] = caminho

    # nome de módulo de uma parte igual a variável/parâmetro de OUTRA parte também dá erro de
    # compilação no módulo único (o VBA não diferencia maiúsculas): renomeia o da parte que declara
    def locais(s):
        nomes = set()
        for m in re.findall(r"^\s+(?:Dim|Static|Const)\s+([^\n']+)", s, re.M):
            for parte in m.split(","):
                if parte.strip():
                    nomes.add(parte.strip().split()[0].split("(")[0].lower())
        for m in re.findall(r"(?:Sub|Function)\s+\w+\s*\(([^)]*)\)", s.replace("_\n", " ")):
            for parte in m.split(","):
                w = [x for x in parte.split() if x not in ("Optional", "ByRef", "ByVal", "ParamArray")]
                if w:
                    nomes.add(w[0].split("(")[0].lower())
        return nomes
    loc = {c: locais(s) for c, _, s in textos}
    for caminho, suf, s in textos:
        for n in sorted(nomes_de_modulo(s)):
            atual = renomear[caminho].get(n, n)
            if any(atual.lower() in loc[o] for o, _, _ in textos if o != caminho):
                if n.lower() in {x.lower() for x in PUBLICAS}:
                    sys.exit("ERRO: macro pública %s colide com variável de outra parte" % n)
                renomear[caminho][n] = n + "_" + suf if atual == n else atual + "x"

    saida = []
    for caminho, suf, s in textos:
        for velho, novo in renomear[caminho].items():
            s = no_codigo(s, lambda t, v=velho, n=novo: re.sub(r"\b%s\b" % re.escape(v), n, t, flags=re.I))
        # Application.Run -> chamada direta
        s = no_codigo(s, lambda t: re.sub(r"Application\.Run\(", "RodarMacroAuto(", t))
        # visibilidade: só as macros de usar ficam públicas
        def vis(m):
            nome = m.group(2)
            return ("Public " if nome.lower() in PUBLICAS else "Private ") + m.group(1) + " " + nome
        s = re.sub(r"^(?:Public\s+|Private\s+)?(Sub|Function)\s+(\w+)", vis, s, flags=re.M)
        cab = "' " + "=" * 108 + "\n' PARTE: " + caminho + "  (gerado por ferramentas/montar_modulo_unico.py; edite o .bas de origem)\n' " + "=" * 108
        saida.append(cab + "\n\n" + s.strip("\n"))

    # (04/11) declarações do nível do módulo -> topo. Percorre cada parte linha a linha, sabendo
    # se está dentro de uma Sub/Function/Property; fora delas, linha que começa com Private/
    # Public/Dim/Global (variável), Const, Type...End Type ou Enum...End Enum (com as
    # continuações " _") sai do lugar e vai para o bloco de declarações.
    re_ini = re.compile(r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+", re.I)
    re_fim = re.compile(r"^End\s+(?:Sub|Function|Property)\b", re.I)
    re_dec = re.compile(r"^(?:(?:Public|Private|Global)\s+(?:Const\s+)?\w+|Dim\s+\w+|Const\s+\w+)", re.I)
    re_bloco = re.compile(r"^(?:Public\s+|Private\s+)?(Type|Enum)\s+\w+", re.I)
    declaracoes = []
    for i, parte in enumerate(saida):
        fica, sobe, dentro, bloco, cont = [], [], False, None, False
        for l in parte.split("\n"):
            t = l.strip()
            if cont or bloco:
                sobe.append(l)
                if bloco and re.match(r"^End\s+" + bloco + r"\b", t, re.I): bloco = None
                cont = l.rstrip().endswith(" _")
                continue
            if not dentro:
                if re_ini.match(t):
                    dentro = True
                elif re_bloco.match(t):
                    bloco = re_bloco.match(t).group(1); sobe.append(l); continue
                elif re_dec.match(t) and not re_ini.match(t):
                    sobe.append(l); cont = l.rstrip().endswith(" _"); continue
            elif re_fim.match(t):
                dentro = False
            fica.append(l)
        saida[i] = "\n".join(fica)
        if sobe:
            nome = re.search(r"PARTE: (\S+)", parte).group(1)
            declaracoes.append("' --- de " + nome + "\n" + "\n".join(sobe))
    decl_txt = "' " + "=" * 108 + "\n' DECLARAÇÕES DE TODAS AS PARTES (o VBA só aceita antes da 1ª Sub/Function)\n' " + "=" * 108 + "\n" + "\n".join(declaracoes)

    casos = "\n".join('        Case "%s": RodarMacroAuto = %s()' % (m.lower(), m) for m in MACROS_AUTO)
    disp = '''' ============================================================================================================
' MÓDULO ÚNICO DO PAINEL DE PONTO (03/11)
'
' Cole SÓ este código num módulo da planilha de Tratamento (Alt+F11 > Inserir > Módulo > colar).
' Se a planilha ainda tiver os módulos antigos (um por macro), apague-os: este já tem tudo.
' Macros (Alt+F8):
'   GerarCSVsDaPasta ............... a do mês: confere os 12 arquivos da pasta (11 + BH), pergunta onde salvar os
'                                    CSVs, traz os relatórios, formata, gera a aba CSV e grava os CSVs
'   PrepararAbasDoMes .............. só formata relatórios que você colou como abas (sem CSV)
' A PURO_Para_EDITADO (aba PURO -> abas de tratamento) NÃO está aqui: é de outro assunto e fica no
' txt/PURO_Para_EDITADO.txt, num módulo separado.
' Este arquivo é montado por ferramentas/montar_modulo_unico.py a partir dos .bas de cada macro.
' ============================================================================================================

' Chamada direta das etapas automáticas (no lugar do Application.Run "nome", que dependia de o
' módulo de cada macro estar colado com aquele nome público).
Private Function RodarMacroAuto(ByVal nome As String, Optional ByVal pasta As Variant) As Variant
    Select Case LCase$(nome)
%s
        Case "exportarcsvporgestor_napasta": RodarMacroAuto = ExportarCSVPorGestor_NaPasta(CStr(pasta))
        Case Else: Err.Raise vbObjectError + 513, , "Etapa desconhecida: " & nome
    End Select
End Function
''' % casos
    corpo = "Option Explicit\n\n" + decl_txt + "\n\n" + disp + "\n\n" + "\n\n".join(saida) + "\n"

    # conferências
    # depois da 1ª Sub/Function só pode haver procedimentos, comentários e linhas em branco
    dentro, primeira = False, False
    for n, l in enumerate(corpo.split("\n"), 1):
        t = l.strip()
        if re_ini.match(t): dentro = primeira = True; continue
        if re_fim.match(t): dentro = False; continue
        if primeira and not dentro and t and not t.startswith("'"):
            sys.exit("ERRO: linha %d fora de procedimento depois da 1ª Sub/Function: %s" % (n, t))
    if any(cod and "Application.Run" in t for l in corpo.split("\n") for cod, t in partes_codigo(l)):
        sys.exit("ERRO: sobrou Application.Run fora de comentário")
    pubs = sorted(n.lower() for n in re.findall(r"^Public (?:Sub|Function) (\w+)", corpo, re.M))
    if pubs != sorted(PUBLICAS):
        sys.exit("ERRO: públicas inesperadas: %s" % pubs)
    vistos, dup = set(), set()
    for n in re.findall(r"^(?:Public |Private )(?:Sub|Function)\s+(\w+)", corpo, re.M) + \
            re.findall(r"^(?:Public |Private )?Const\s+(\w+)", corpo, re.M) + \
            re.findall(r"^(?:Private|Public|Dim)\s+(?!Sub|Function|Const|Type|Enum)(\w+)\s+As\b", corpo, re.M):
        (dup if n.lower() in vistos else vistos).add(n.lower())
    if dup:
        sys.exit("ERRO: nomes repetidos no módulo: %s" % sorted(dup))

    import os
    os.makedirs("ModuloUnico", exist_ok=True)
    open("ModuloUnico/PainelPonto_TudoEmUm.bas", "w", encoding="utf-8", newline="\n").write(
        'Attribute VB_Name = "PainelPonto_TudoEmUm"\n' + corpo)
    open("txt/PainelPonto_TudoEmUm.txt", "w", encoding="utf-8", newline="\r\n").write(corpo)
    ren = {c: r for c, r in renomear.items() if r}
    print("OK: ModuloUnico/PainelPonto_TudoEmUm.bas e txt/PainelPonto_TudoEmUm.txt (%d linhas)" % corpo.count("\n"))
    for c, r in ren.items():
        print("  renomeados em %s: %s" % (c, ", ".join("%s->%s" % kv for kv in sorted(r.items()))))


if __name__ == "__main__":
    main()
