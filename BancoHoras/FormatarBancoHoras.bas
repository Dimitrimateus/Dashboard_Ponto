Attribute VB_Name = "modFormatarBancoHoras"
Option Explicit

' Modo automático: quando a GerarCSVsDaPasta / PrepararAbasDoMes chama esta formatação
' (FormatarBancoHoras_Auto), nenhuma mensagem aparece; o texto de cada aviso vai para msgAuto e
' volta para o relatório da preparação. Ver Aviso, no fim do módulo.
Private modoAuto As Boolean
Private msgAuto As String
' ============================================================================
' FormatarBancoHoras  (04/11)
'
' DE ONDE VEM: a aba ATIVA, com o relatório de banco de horas do sistema (HRES105,
'   o arquivo "BH" da pasta do mês), exportado como "dump" de impressão:
'     - cabeçalho de página repetido: "Filial: <filial>   Período: <início> a <fim>",
'       os títulos "Período Anterior | Movimento Mensal | Fechamento Banco | Saldo
'       Final" e embaixo "Locais | Positivo | Negativo | Credito | Débito | Hr. Anist.
'       | Hr. Pagas | Positivo | Negativo";
'     - uma linha com o local ("Aurora, Unid Vendas ..., Setor X, Ativ. X");
'     - uma linha por pessoa: matrícula (coluna A), nome (coluna B) e as 8 horas;
'     - linhas de subtotal (só o nome do local na coluna A, sem nome de pessoa) e o
'       rodapé "HRES105.GER - data - hora", que são ignorados.
'   As horas vêm como hora do Excel (fração de dia, ex.: 1,0701389 = 25:41) ou como
'   texto com o sinal no fim quando são negativas ("25:40-", sempre na coluna Débito).
'
' O QUE SIGNIFICA CADA COLUNA (o banco de horas fecha de 4 em 4 meses):
'   Saldo anterior ... o que a pessoa trazia de antes do período (positivo/negativo)
'   Crédito .......... horas a mais trabalhadas no período (entraram no banco)
'   Débito ........... horas a menos (atraso, saída antes, folga tirada do banco)
'   Horas anistiadas . no fechamento, o saldo NEGATIVO que fica zerado (na análise do RH de
'                      05/11, é o que vai como desconto na folha)
'   Horas pagas ...... no fechamento, o saldo POSITIVO que a empresa pagou
'   Saldo final ...... o que sobra no banco depois do fechamento (zero quando fecha)
'   Num período que inclui um fechamento: crédito − débito = horas pagas − horas anistiadas.
'   Os períodos são fixos: novembro a fevereiro, março a junho, julho a outubro (05/11).
'
' PARA ONDE VAI: reescreve a MESMA aba como uma tabela (cabeçalho na linha 1, uma
'   linha por pessoa) e renomeia a aba para "Banco de Horas". As horas ficam em
'   formato [h]:mm, sempre POSITIVAS (o lado diz o sinal: débito, saldo negativo).
'   O GerarCSVPonto acha esta aba pelo cabeçalho (Matrícula + Crédito + Horas pagas)
'   e gera uma linha "Banco de Horas Período" por pessoa no CSV; o painel mostra
'   na seção "Colaboradores em destaque" como cada pessoa fechou o banco.
'
' Este módulo não depende de nenhum outro. Nunca rodou no Excel: rode numa CÓPIA.
' ============================================================================

Private Const NOME_ABA_SAIDA As String = "Banco de Horas"
Private Const N_COLS As Long = 16

Sub FormatarBancoHoras()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Lê da A1 até a última célula usada (o relatório pode ter colunas vazias no começo).
    Dim ultLin As Long, ultCol As Long
    ultLin = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    ultCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If ultLin < 3 Or ultCol < 2 Then
        Aviso "Nenhum dado encontrado na aba ativa.", vbExclamation, "Formatar Banco de Horas"
        Exit Sub
    End If
    If ultCol < 14 Then ultCol = 14
    Dim dados As Variant
    dados = ws.Range(ws.Cells(1, 1), ws.Cells(ultLin, ultCol)).Value2

    ' ------------------------------------------------------------
    ' 1) Colunas das horas, pela linha de títulos "Locais | Positivo | Negativo |
    '    Credito | Débito | Hr. Anist. | Hr. Pagas | Positivo | Negativo".
    '    O 1º Positivo/Negativo é o saldo anterior; o 2º, o saldo final.
    '    Reserva: as posições do relatório de 03/11 (E a L).
    ' ------------------------------------------------------------
    Dim cAntPos As Long, cAntNeg As Long, cCred As Long, cDeb As Long
    Dim cAnist As Long, cPagas As Long, cFinPos As Long, cFinNeg As Long
    cAntPos = 5: cAntNeg = 6: cCred = 7: cDeb = 8: cAnist = 9: cPagas = 10: cFinPos = 11: cFinNeg = 12
    Dim r As Long, c As Long, t As String, achouTitulos As Boolean
    For r = 1 To ultLin
        If TextoSimples(Celula(dados, r, 1)) = "LOCAIS" Then
            Dim nPos As Long, nNeg As Long
            nPos = 0: nNeg = 0
            For c = 2 To ultCol
                t = TextoSimples(Celula(dados, r, c))
                Select Case t
                    Case "POSITIVO"
                        nPos = nPos + 1
                        If nPos = 1 Then cAntPos = c Else cFinPos = c
                    Case "NEGATIVO"
                        nNeg = nNeg + 1
                        If nNeg = 1 Then cAntNeg = c Else cFinNeg = c
                    Case "CREDITO": cCred = c
                    Case "DEBITO": cDeb = c
                    Case "HR. ANIST.", "HR ANIST", "HR. ANIST": cAnist = c
                    Case "HR. PAGAS", "HR PAGAS": cPagas = c
                End Select
            Next c
            achouTitulos = True
            Exit For
        End If
    Next r
    If Not achouTitulos Then
        Aviso "Não encontrei a linha de títulos ""Locais | Positivo | Negativo | Credito | Débito..."" " & _
              "do relatório de banco de horas. A aba ativa é mesmo o relatório HRES105?", vbCritical, "Formatar Banco de Horas"
        Exit Sub
    End If

    ' ------------------------------------------------------------
    ' 2) Percorre o relatório: guarda o período, a filial e o local correntes e
    '    monta uma linha por pessoa (matrícula numérica na coluna A + nome na B).
    ' ------------------------------------------------------------
    Dim saida() As Variant, nSaida As Long
    ReDim saida(1 To ultLin, 1 To N_COLS)
    Dim iniPer As Double, fimPer As Double, filial As String, localAtual As String
    Dim v As Variant, k As Long
    For r = 1 To ultLin
        ' "Período:" e "Filial:" (cabeçalho de cada página)
        For c = 1 To ultCol
            t = TextoSimples(Celula(dados, r, c))
            If t = "PERIODO:" Then
                For k = c + 1 To ultCol
                    v = Celula(dados, r, k)
                    If EhNumero(v) Then
                        If v > 30000 And v < 80000 Then
                            If iniPer = 0 Or v < iniPer Then iniPer = v
                            If v > fimPer Then fimPer = v
                        End If
                    End If
                Next k
            ElseIf t = "FILIAL:" Then
                For k = c + 1 To ultCol
                    If TextoCel(Celula(dados, r, k)) <> "" Then filial = TextoCel(Celula(dados, r, k)): Exit For
                Next k
            End If
        Next c

        v = Celula(dados, r, 1)
        t = TextoCel(v)
        If Left$(TextoSimples(v), 7) = "AURORA," Then
            ' linha do local: fica a última parte ("Ativ. Faturamento")
            localAtual = Trim$(Mid$(t, InStrRev(t, ",") + 1))
        ElseIf UCase$(t) = "AURORA" Or Left$(UCase$(t), 7) = "AURORA " Then
            localAtual = t
        ElseIf PareceNumero(v) And TextoCel(Celula(dados, r, 2)) <> "" Then
            nSaida = nSaida + 1
            saida(nSaida, 1) = CStr(CLng(CDbl(t)))
            saida(nSaida, 2) = TextoCel(Celula(dados, r, 2))
            saida(nSaida, 3) = filial
            saida(nSaida, 4) = localAtual
            saida(nSaida, 9) = HoraBH(Celula(dados, r, cAntPos))
            saida(nSaida, 10) = HoraBH(Celula(dados, r, cAntNeg))
            saida(nSaida, 11) = HoraBH(Celula(dados, r, cCred))
            saida(nSaida, 12) = HoraBH(Celula(dados, r, cDeb))
            saida(nSaida, 13) = HoraBH(Celula(dados, r, cAnist))
            saida(nSaida, 14) = HoraBH(Celula(dados, r, cPagas))
            saida(nSaida, 15) = HoraBH(Celula(dados, r, cFinPos))
            saida(nSaida, 16) = HoraBH(Celula(dados, r, cFinNeg))
        End If
    Next r

    If nSaida = 0 Then
        Aviso "Não encontrei nenhuma pessoa no relatório de banco de horas (linhas com matrícula na " & _
              "coluna A e nome na coluna B).", vbCritical, "Formatar Banco de Horas"
        Exit Sub
    End If
    If iniPer = 0 Or fimPer = 0 Then
        Aviso "Não encontrei o período (linha ""Período: ... a ..."") no relatório de banco de horas.", _
              vbCritical, "Formatar Banco de Horas"
        Exit Sub
    End If

    ' Período: início, fim, quantos meses (ciclos de 28 a 27) e se o banco fechou.
    ' (05/11) Na Aurora o banco tem 3 períodos fixos de 4 meses: novembro a fevereiro,
    ' março a junho e julho a outubro; o fechamento é no último mês de cada um
    ' (fevereiro, junho e outubro). O relatório pode ter 1 mês, 2, o período todo ou o
    ' ano: está "fechado" quando inclui o mês de um fechamento.
    Dim meses As Long, fechado As String
    meses = MesesDoPeriodo(CDate(iniPer), CDate(fimPer))
    fechado = IIf(FechamentosNoPeriodo(CDate(iniPer), CDate(fimPer)) > 0, "Sim", "Não")
    For r = 1 To nSaida
        saida(r, 5) = CDate(iniPer)
        saida(r, 6) = CDate(fimPer)
        saida(r, 7) = meses
        saida(r, 8) = fechado
    Next r

    ' ------------------------------------------------------------
    ' 3) Reescreve a aba
    ' ------------------------------------------------------------
    Application.ScreenUpdating = False
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.UnMerge
    On Error GoTo 0
    ws.Cells.Clear

    ' Estes nomes são o que o GerarCSVPonto procura: não mude sem mudar lá.
    Dim cab As Variant
    cab = Array("Matrícula", "Colaborador", "Filial", "Local", "Início do período", "Fim do período", _
                "Meses do período", "Banco fechado", "Saldo anterior positivo", "Saldo anterior negativo", _
                "Crédito", "Débito", "Horas anistiadas", "Horas pagas", "Saldo final positivo", "Saldo final negativo")
    For c = 0 To N_COLS - 1
        ws.Cells(1, c + 1).Value = cab(c)
    Next c
    ws.Range(ws.Cells(2, 1), ws.Cells(nSaida + 1, 1)).NumberFormat = "@"
    ws.Range(ws.Cells(2, 5), ws.Cells(nSaida + 1, 6)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(2, 9), ws.Cells(nSaida + 1, 16)).NumberFormat = "[h]:mm"

    Dim bloco() As Variant
    ReDim bloco(1 To nSaida, 1 To N_COLS)
    For r = 1 To nSaida
        For c = 1 To N_COLS
            bloco(r, c) = saida(r, c)
        Next c
    Next r
    ws.Range(ws.Cells(2, 1), ws.Cells(nSaida + 1, N_COLS)).Value = bloco

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, N_COLS))
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(nSaida + 1, N_COLS))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .AutoFilter
    End With
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 36
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:P").ColumnWidth = 11
    ws.Range("E:P").HorizontalAlignment = xlCenter

    Dim renomeada As Boolean
    On Error Resume Next
    ws.Name = NOME_ABA_SAIDA
    renomeada = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select
    Application.ScreenUpdating = True

    Aviso "Formatação concluída." & vbCrLf & _
          nSaida & " pessoa(s) no banco de horas, período de " & Format(CDate(iniPer), "dd/mm/yyyy") & _
          " a " & Format(CDate(fimPer), "dd/mm/yyyy") & " (" & meses & " mês(es): " & _
          IIf(fechado = "Sim", "inclui o fechamento do banco", "sem fechamento: o saldo é parcial") & ")." & vbCrLf & _
          IIf(renomeada, "Aba renomeada para """ & NOME_ABA_SAIDA & """.", _
              "A aba não foi renomeada (já existe uma aba """ & NOME_ABA_SAIDA & """)."), _
          vbInformation, "Formatar Banco de Horas"
End Sub

' ----------------------------------------------------------------------------
' FUNÇÕES DE APOIO (só deste módulo)
' ----------------------------------------------------------------------------
' Hora do relatório -> hora do Excel (fração de dia), sempre positiva.
' Número = já é hora do Excel; texto "25:40-" ou "-25:40" = 25h40 (o "-" só diz
' que é negativo, e o lado da coluna já diz isso); vazio = 0.
Private Function HoraBH(v As Variant) As Double
    Dim t As String, p As Long, h As Double, m As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) <> vbString Then
        If IsNumeric(v) Then HoraBH = Abs(CDbl(v))
        Exit Function
    End If
    t = Replace(Replace(Trim$(CStr(v)), "-", ""), " ", "")
    If t = "" Then Exit Function
    p = InStr(t, ":")
    If p = 0 Then
        If IsNumeric(t) Then HoraBH = Abs(CDbl(t))
        Exit Function
    End If
    If IsNumeric(Left$(t, p - 1)) Then h = CDbl(Left$(t, p - 1))
    If IsNumeric(Mid$(t, p + 1)) Then m = CDbl(Mid$(t, p + 1))
    HoraBH = (h * 60 + m) / 1440
End Function

' Quantos meses (ciclos de 28 a 27) o período cobre: 28/06 a 27/10 = 4
' (julho, agosto, setembro e outubro). Dia 28 em diante já é do ciclo seguinte.
Private Function MesesDoPeriodo(ByVal ini As Date, ByVal fim As Date) As Long
    Dim a As Long, b As Long
    a = Year(ini) * 12 + Month(ini) + IIf(Day(ini) >= 28, 1, 0)
    b = Year(fim) * 12 + Month(fim) + IIf(Day(fim) >= 28, 1, 0)
    MesesDoPeriodo = b - a + 1
End Function

' (05/11) Quantos fechamentos de banco (meses de outubro, fevereiro e junho, pelo
' ciclo de 28 a 27) o período inclui.
Private Function FechamentosNoPeriodo(ByVal ini As Date, ByVal fim As Date) As Long
    Dim a As Long, b As Long, k As Long, mes As Long
    a = Year(ini) * 12 + Month(ini) + IIf(Day(ini) >= 28, 1, 0)
    b = Year(fim) * 12 + Month(fim) + IIf(Day(fim) >= 28, 1, 0)
    For k = a To b
        mes = ((k - 1) Mod 12) + 1
        If mes = 10 Or mes = 2 Or mes = 6 Then FechamentosNoPeriodo = FechamentosNoPeriodo + 1
    Next k
End Function

Private Function Celula(dados As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If c < 1 Or c > UBound(dados, 2) Or r < 1 Or r > UBound(dados, 1) Then
        Celula = Empty
    Else
        Celula = dados(r, c)
    End If
End Function

Private Function EhNumero(v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then
        EhNumero = False
    ElseIf VarType(v) = vbString Then
        EhNumero = False
    Else
        EhNumero = IsNumeric(v)
    End If
End Function

Private Function PareceNumero(v As Variant) As Boolean
    Dim t As String
    t = TextoCel(v)
    PareceNumero = (t <> "" And IsNumeric(t))
End Function

Private Function TextoCel(v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then
        TextoCel = ""
    Else
        TextoCel = Trim$(Replace(Replace(Replace(CStr(v), vbCr, " "), vbLf, " "), Chr(160), " "))
    End If
End Function

Private Function TextoSimples(v As Variant) As String
    Dim t As String, com As String, sem As String, i As Long
    t = UCase$(TextoCel(v))
    com = "ÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇ"
    sem = "AAAAAEEEEIIIIOOOOOUUUUC"
    For i = 1 To Len(com)
        t = Replace(t, Mid$(com, i, 1), Mid$(sem, i, 1))
    Next i
    TextoSimples = t
End Function

' ----------------------------------------------------------------------------
' Modo automático. FormatarBancoHoras_Auto é chamada pela PrepararAbasDoMes /
' GerarCSVsDaPasta com a aba do relatório ativa: roda sem caixa de mensagem e
' devolve o texto (com "ERRO: " na frente quando o aviso é de erro).
' ----------------------------------------------------------------------------
Public Function FormatarBancoHoras_Auto() As String
    modoAuto = True: msgAuto = ""
    FormatarBancoHoras
    modoAuto = False
    FormatarBancoHoras_Auto = msgAuto
End Function

Private Function Aviso(ByVal texto As String, Optional ByVal botoes As Long = vbInformation, Optional ByVal titulo As String = "") As VbMsgBoxResult
    If modoAuto Then
        If (botoes And 16) = 16 Then texto = "ERRO: " & texto
        msgAuto = msgAuto & IIf(msgAuto = "", "", " ") & Replace(Replace(texto, vbCrLf, " "), vbLf, " ")
        Aviso = vbYes
    Else
        Aviso = MsgBox(texto, botoes, titulo)
    End If
End Function
