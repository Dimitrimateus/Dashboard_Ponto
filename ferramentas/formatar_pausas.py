import zipfile, re, sys, pickle, datetime
# Reproduz em Python o FormatarPausasTermicas.bas: lê o relatório "Analítico de Pausas Térmicas"
# (HRES114) exportado pelo sistema e devolve a tabela de colaboradores que a macro deixa na aba
# (uma linha por pessoa, com o bloco "Total do Colaborador"). Também devolve o período do
# relatório (linha "Período: ... até ..."), que diz de que mês é a aba.
# Uso: python3 formatar_pausas.py "pausa_termica.xlsx" [saida.pkl]
#   saida.pkl = {fim, linhas, geral, batidas}; o adicionar_aba_xlsx.py (layout "pausas") monta a aba.
#   batidas(caminho) devolve a 3ª tabela da macro (18/10): uma linha por batida, com matrícula,
#   nome, data, hora (fração do dia), nº da batida no dia, letra (P/G/F) e exceções do dia.
# Por que ler o XML na mão: o .xlsx do Senior tem "\" nos caminhos do zip e estilos que o
# openpyxl não abre (CLAUDE.md, armadilha 8). A macro separa a linha da pessoa (matrícula na
# coluna A) da linha de dia (data na coluna A) pelo formato da célula: dd/mm/aaaa = dia. Aqui o
# formato vem do styles.xml (cellXfs -> numFmt).
CAB = ["Matrícula", "Colaborador", "Cargo", "Pausa corretas", "Pausa menor 0:20", "Pausa maior 0:20",
       "Trabalho correto 1:40", "Trabalho maior 1:40", "Trabalho menor 1:40", "Marcações Ímpares"]
BASE = datetime.datetime(1899, 12, 30)

def col_num(letras):
    n = 0
    for ch in letras: n = n * 26 + ord(ch) - 64
    return n

# {linha: {coluna: (valor, é_data)}}: texto de <t>, número de <v>; é_data = o formato da célula
# é de data (dd/mm/aaaa).
def ler(caminho):
    z = zipfile.ZipFile(caminho)
    achar = lambda fim: next(n for n in z.namelist() if n.replace('\\', '/').endswith(fim))
    st = z.read(achar('styles.xml')).decode('utf-8')
    fmts = {int(i): c for i, c in re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', st)}
    xfs = re.findall(r'<xf [^>]*numFmtId="(\d+)"', re.search(r'<cellXfs.*?</cellXfs>', st, re.S).group(0))
    eh_data = [int(f) in (14, 15, 16, 17, 22) or 'DD/MM' in fmts.get(int(f), '').upper() for f in xfs]
    x = z.read(achar('sheet1.xml')).decode('utf-8')
    linhas = {}
    for m in re.finditer(r'<row[^>]*r="(\d+)"[^>]*>(.*?)</row>', x, re.S):
        cel = {}
        for c in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*)>(.*?)</c>', m.group(2), re.S):
            attrs, corpo = c.group(2), c.group(3)
            s = re.search(r' s="(\d+)"', attrs)
            data = bool(s) and int(s.group(1)) < len(eh_data) and eh_data[int(s.group(1))]
            t = re.search(r'<t[^>]*>(?:<!\[CDATA\[)?(.*?)(?:\]\]>)?</t>', corpo, re.S)
            v = re.search(r'<v>(.*?)</v>', corpo)
            if t: val = t.group(1).strip()
            elif v and v.group(1).strip():
                val = float(v.group(1))
                if val == int(val): val = int(val)
            else: continue
            cel[col_num(c.group(1))] = (val, data)
        linhas[int(m.group(1))] = cel
    return linhas

def num(v): return isinstance(v, (int, float))

def formatar(caminho):
    L = ler(caminho)
    fim_periodo = None
    for r in sorted(L):
        txts = {c: v for c, (v, _) in L[r].items()}
        if any(v == 'Período:' for v in txts.values()):
            datas = [v for c, (v, d) in sorted(L[r].items()) if d and num(v)]
            if datas: fim_periodo = BASE + datetime.timedelta(days=datas[-1])
            break
    saida, atual = [], None
    g = lambda r, c: L.get(r, {}).get(c, (None, False))[0]
    for r in sorted(L):
        a, a_data = L[r].get(1, (None, False))
        b, f = g(r, 2), g(r, 6)
        if num(a) and a_data:
            continue                               # linha de dia
        if num(a) and isinstance(b, str) and b and isinstance(f, str) and f:
            atual = (str(int(a)), b.strip(), f.strip())   # linha da pessoa
        elif a == 'Total do Colaborador' and atual:
            v = [g(r + 2, c) for c in (1, 2, 4, 6, 8, 10, 11)]
            saida.append([int(atual[0]), atual[1], atual[2]] + v)
    return fim_periodo, saida

# Batidas de cada dia, como GuardarBatidasDoDia da macro: depois da data (col. A) vêm pares
# horário + letra; as exceções do dia são o texto entre parênteses no fim da linha.
def batidas(caminho):
    L = ler(caminho); atual = None; out = []
    for r in sorted(L):
        a, a_data = L[r].get(1, (None, False))
        b, f = L[r].get(2, (None, False))[0], L[r].get(6, (None, False))[0]
        if num(a) and a_data:
            if not atual: continue
            exc = ''
            for c in range(2, 46):
                v = L[r].get(c, (None, False))[0]
                if isinstance(v, str) and v.strip().startswith('('): exc = v.strip()
            seq = 0
            for c in range(2, 45):
                v = L[r].get(c, (None, False))[0]; l = L[r].get(c + 1, (None, False))[0]
                if num(v) and isinstance(l, str) and len(l.strip()) == 1:
                    seq += 1
                    out.append((atual[0], atual[1], BASE + datetime.timedelta(days=a), v - int(v), seq, l.strip(), exc))
        elif num(a) and isinstance(b, str) and b and isinstance(f, str) and f:
            atual = (int(a), b.strip())
    return out

# Total Geral do relatório (7 números, 2 linhas abaixo de "Total | Geral"), como a macro copia.
def total_geral(caminho):
    L = ler(caminho)
    for r in sorted(L):
        if L[r].get(1, (None, 0))[0] == 'Total' and L[r].get(2, (None, 0))[0] == 'Geral':
            return [L.get(r + 2, {}).get(c, (0, 0))[0] for c in (1, 2, 4, 6, 8, 10, 11)]
    return None

# Saída: um dict com tudo o que a macro deixa na aba (fim do período, tabela de colaboradores,
# Total Geral e batidas); o adicionar_aba_xlsx.py (layout "pausas") monta as 3 tabelas com ele.
if __name__ == '__main__':
    fim, linhas = formatar(sys.argv[1])
    b = batidas(sys.argv[1])
    print('%d colaboradores, %d batidas; período até %s' % (len(linhas), len(b), fim.strftime('%d/%m/%Y') if fim else '?'))
    pickle.dump({'fim': fim, 'linhas': linhas, 'geral': total_geral(sys.argv[1]), 'batidas': b},
                open(sys.argv[2] if len(sys.argv) > 2 else 'pausas.pkl', 'wb'))
