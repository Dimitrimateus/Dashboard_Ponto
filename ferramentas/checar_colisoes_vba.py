import re, sys
# Procura nomes repetidos entre procedimentos/constantes e variáveis/parâmetros num módulo VBA.
# O VBA não diferencia maiúsculas: uma função "EmendaFolga" e um parâmetro "emendaFolga" no mesmo
# módulo dão erro de compilação (CLAUDE.md, armadilha 1). Rode antes de entregar qualquer .bas.
# Uso: python3 ferramentas/checar_colisoes_vba.py GerarCSVPonto/GerarCSVPonto.bas [outros.bas ...]
# procs/consts = nomes de Sub/Function e de constantes; nomes = variáveis (Dim) e parâmetros.
# Tudo em minúsculas, porque é assim que o VBA compara. Qualquer nome nos dois grupos é colisão.
# Também acusa For dentro de For com a mesma variável (04/11).
# Sai com código 1 se houver colisão (dá para usar em script).
ok = True
for caminho in sys.argv[1:]:
    s = open(caminho, encoding='utf-8').read()
    procs = {m.lower() for m in re.findall(r'(?:Sub|Function)\s+(\w+)\s*\(', s)}
    consts = {m.lower() for m in re.findall(r'Const\s+(\w+)', s)}
    nomes = set()
    for m in re.findall(r'\bDim\s+([^\n]+)', s):
        for parte in m.split(','):
            if parte.strip(): nomes.add(parte.strip().split()[0].split('(')[0].lower())
    for m in re.findall(r'(?:Sub|Function)\s+\w+\s*\(([^)]*)\)', s, re.S):
        for parte in m.replace('_\n', ' ').split(','):
            w = [x for x in parte.split() if x not in ('Optional', 'ByRef', 'ByVal')]
            if w: nomes.add(w[0].lower())
    colisoes = sorted((procs | consts) & nomes)
    # (04/11) For dentro de For com a MESMA variável ("Variável de controle For já em uso"):
    # a pilha guarda as variáveis dos For abertos; Next fecha o último.
    forRep, pilha = [], []
    for n, l in enumerate(s.split('\n'), 1):
        t = l.split("'")[0].strip() if '"' not in l else l.strip()
        if re.match(r'^(?:Public |Private )?(?:Sub|Function|Property)\b', t): pilha = []
        for st in re.split(r':\s', t):
            st = st.strip()
            mf = re.match(r'^For\s+Each\s+(\w+)\s+In\b', st, re.I) or re.match(r'^For\s+(\w+)\s*=', st, re.I)
            if mf:
                v = mf.group(1).lower()
                if v in pilha: forRep.append('linha %d: For %s dentro de outro For %s' % (n, v, v))
                pilha.append(v)
            elif re.match(r'^Next\b', st, re.I) and pilha:
                pilha.pop()
    problemas = (['COLISÕES: ' + ', '.join(colisoes)] if colisoes else []) + forRep
    print(caminho, '->', 'sem colisões' if not problemas else '; '.join(problemas))
    ok = ok and not problemas
sys.exit(0 if ok else 1)
