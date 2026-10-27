import csv, sys, pickle, datetime
# Lê o relatório "Ausência de marcação" do PontoNet exportado em CSV (separador ";",
# acentos em Latin-1) e grava as linhas com os tipos certos (data, data e hora, hora,
# número), na ordem das colunas do layout "pontonet" do adicionar_aba_xlsx.py.
# Uso: python3 ler_pontonet_csv.py "Ausência de marcação.csv" pontonet.pkl
# Depois: adicionar_aba_xlsx.py entrada.xlsx saida.xlsx pontonet.pkl pontonet "Pontonet" --substituir
from adicionar_aba_xlsx import LAYOUTS

def valor(v, col):
    v = v.strip()
    if v == '': return None
    if col == 'Código terceiro': return v                      # "02": fica texto
    for fmt in ('%d/%m/%Y %H:%M', '%d/%m/%Y'):
        try: return datetime.datetime.strptime(v, fmt)
        except ValueError: pass
    if col == 'Hora justificada' and ':' in v:
        h, m = v.split(':')[:2]; return (int(h) * 60 + int(m)) / 1440   # fração do dia
    if v.lstrip('-').isdigit(): return int(v)
    return v

if __name__ == '__main__':
    texto = open(sys.argv[1], 'rb').read()
    try: texto = texto.decode('utf-8-sig')
    except UnicodeDecodeError: texto = texto.decode('latin-1')
    rd = csv.DictReader(texto.splitlines(), delimiter=';')
    cab = LAYOUTS['pontonet']['cab']
    linhas = [[valor(r.get(c) or '', c) for c in cab] for r in rd]
    pickle.dump(linhas, open(sys.argv[2] if len(sys.argv) > 2 else 'pontonet.pkl', 'wb'))
    print(len(linhas), 'linhas')
