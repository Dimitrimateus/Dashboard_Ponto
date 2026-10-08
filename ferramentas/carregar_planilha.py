import openpyxl, pickle, sys
# Uso: python3 carregar_planilha.py "Tratamento Ponto ... .xlsx"  -> grava planilha.pkl
# Lê SÓ valores (data_only=True) e nunca salva o .xlsx (openpyxl estraga validações x14,
# Power Query e comentários dessas pastas de trabalho - ver CLAUDE.md, armadilhas).
# Resultado: dicionário {nome da aba: [linhas]}, cada linha uma lista de valores (datas como
# datetime, horas como time/timedelta). É a entrada do simular_gerarcsv.py e do
# conferir_painel.py (variável de ambiente PLANILHA_PKL ou o arquivo planilha.pkl).
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
d = {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
pickle.dump(d, open(sys.argv[2] if len(sys.argv) > 2 else 'planilha.pkl', 'wb'))
print('abas:', list(d))
