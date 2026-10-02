import openpyxl, pickle, sys
# Uso: python3 carregar_planilha.py "Tratamento Ponto ... .xlsx"  -> grava planilha.pkl
# Lê SÓ valores (data_only=True) e nunca salva o .xlsx (openpyxl estraga validações x14,
# Power Query e comentários dessas pastas de trabalho - ver CLAUDE.md, armadilhas).
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
d = {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
pickle.dump(d, open(sys.argv[2] if len(sys.argv) > 2 else 'planilha.pkl', 'wb'))
print('abas:', list(d))
