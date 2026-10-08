import re, sys
# Procura nomes repetidos entre procedimentos/constantes e variáveis/parâmetros num módulo VBA.
# O VBA não diferencia maiúsculas: uma função "EmendaFolga" e um parâmetro "emendaFolga" no mesmo
# módulo dão erro de compilação (CLAUDE.md, armadilha 1). Rode antes de entregar qualquer .bas.
# Uso: python3 ferramentas/checar_colisoes_vba.py GerarCSVPonto/GerarCSVPonto.bas [outros.bas ...]
# procs/consts = nomes de Sub/Function e de constantes; nomes = variáveis (Dim) e parâmetros.
# Tudo em minúsculas, porque é assim que o VBA compara. Qualquer nome nos dois grupos é colisão.
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
    print(caminho, '->', 'sem colisões' if not colisoes else 'COLISÕES: ' + ', '.join(colisoes))
    ok = ok and not colisoes
sys.exit(0 if ok else 1)
