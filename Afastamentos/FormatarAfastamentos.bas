Attribute VB_Name = "modFormatarAfastamentos"
Option Explicit

' ============================================================================
' FormatarAfastamentos
'
' DE ONDE VEM: a aba ATIVA no momento em que a macro é executada. Espera-se
'   que seja o relatório "Histórico de Afastamentos" (HRCL006) do sistema,
'   exportado como "dump" de impressão:
'     - cabeçalho de página (empresa, "Histórico de Afastamentos", "Empresa",
'       "Local", rodapé "HRCL006.GER - data - hora") repetido a cada página;
'     - por colaborador, um cabeçalho "Tipo | Colaborador | Admissão" e a
'       linha da pessoa: tipo, matrícula (coluna de "Colaborador"), nome (uma
'       coluna à direita) e data de admissão;
'     - depois, o cabeçalho "Afastamento | Situação | Dem. | Dias | Faltas |
'       Dt.Término | Prev Término | Exame" e uma linha por afastamento, com os
'       valores DESLOCADOS em relação ao cabeçalho:
'         data de início  = coluna de "Afastamento";  hora de início = +1
'         código          = coluna de "Situação" + 1; descrição = +2
'         data de término = coluna de "Dt.Término";   hora de término = +1
'         previsão de término = coluna de "Prev Término" + 1
'         exame           = coluna de "Exame" + 1
'       ("Dem.", "Dias" e "Faltas" vêm sempre "00"/0 e não são usados.)
'
' PARA ONDE VAI: reescreve essa MESMA aba como uma tabela simples, cabeçalho
'   na linha 1, uma linha por afastamento, e renomeia a aba para
'   "Afastamentos" (se o nome estiver livre). Colunas:
'     Matrícula · Colaborador · Admissão · Cód. Situação · Situação · Início ·
'     Hora início · Término · Hora término · Dias · Horas · Prev. Término · Exame
'   Dias  = dias corridos do início ao término (término - início + 1).
'   Horas = só para afastamento de poucas horas no mesmo dia (ex.: "Saída
'           Médico/Empresa" das 08:00 às 11:18 = 3:18); vazio nos de dia inteiro.
'   O GerarCSVPonto acha esta aba pelo cabeçalho (Matrícula + Situação +
'   Início + Término), não pelo nome.
'
' Rode numa CÓPIA do arquivo exportado: a aba ativa é reescrita por completo.
' ============================================================================

Private Const NOME_ABA_SAIDA As String = "Afastamentos"

Sub FormatarAfastamentos()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim usado As Range
    Set usado = ws.UsedRange
    If usado.Rows.Count < 2 Then
        MsgBox "Nenhum dado encontrado na aba ativa.", vbExclamation, "Formatar Afastamentos"
        Exit Sub
    End If

    ' Value2: datas e horas vêm como número (com Value, data vira Date e
    ' IsNumeric dá False).
    Dim dados As Variant
    dados = usado.Value2
    Dim deslCol As Long, nLin As Long, nCol As Long
    deslCol = usado.Column - 1
    nLin = UBound(dados, 1)
    nCol = UBound(dados, 2)

    ' ------------------------------------------------------------
    ' 1) Colunas, achadas nos cabeçalhos do relatório. Sem cabeçalho, usa as
    '    posições do relatório de 07/10/2026.
    ' ------------------------------------------------------------
    Dim cTipo As Long, cMat As Long, cNome As Long, cAdm As Long
    Dim cIni As Long, cHoraIni As Long, cCod As Long, cDesc As Long
    Dim cFim As Long, cHoraFim As Long, cPrev As Long, cExame As Long
    cTipo = 1: cMat = 2: cNome = 3: cAdm = 8
    cIni = 1: cHoraIni = 2: cCod = 3: cDesc = 4: cFim = 9: cHoraFim = 10: cPrev = 11: cExame = 12

    Dim r As Long, c As Long, t As String
    Dim achouPessoa As Boolean, achouAfast As Boolean
    For r = 1 To nLin
        If Not achouPessoa And TextoSimples(dados(r, 1)) = "TIPO" Then
            For c = 1 To nCol
                t = TextoSimples(dados(r, c))
                If t = "TIPO" Then cTipo = c
                If t = "COLABORADOR" Then cMat = c: cNome = c + 1
                If t = "ADMISSAO" Then cAdm = c
            Next c
            achouPessoa = True
        End If
        If Not achouAfast And TextoSimples(dados(r, 1)) = "AFASTAMENTO" Then
            For c = 1 To nCol
                t = TextoSimples(dados(r, c))
                If t = "AFASTAMENTO" Then cIni = c: cHoraIni = c + 1
                If t = "SITUACAO" Then cCod = c + 1: cDesc = c + 2
                If t = "DT.TERMINO" Then cFim = c: cHoraFim = c + 1
                If t = "PREV TERMINO" Then cPrev = c + 1
                If t = "EXAME" Then cExame = c + 1
            Next c
            achouAfast = True
        End If
        If achouPessoa And achouAfast Then Exit For
    Next r

    ' ------------------------------------------------------------
    ' 2) Uma linha de saída por afastamento, com a pessoa do último
    '    cabeçalho de colaborador visto.
    '      linha de pessoa      = tipo pequeno (< 1000) + matrícula numérica
    '                             + nome em texto;
    '      linha de afastamento = data de início (número > 1000) + descrição.
    ' ------------------------------------------------------------
    Const N_COLS As Long = 13
    Dim saida() As Variant
    ReDim saida(1 To nLin, 1 To N_COLS)
    Dim nSaida As Long
    Dim matAtual As Variant, nomeAtual As String, admAtual As Variant
    Dim vIni As Variant, vFim As Variant, vHi As Variant, vHf As Variant, vPrev As Variant
    matAtual = Empty

    For r = 1 To nLin
        vIni = Celula(dados, r, cIni - deslCol)
        ' o tipo vem como texto ("1") no relatório: aceita número ou texto numérico
        If PareceNumero(Celula(dados, r, cTipo - deslCol)) And PareceNumero(Celula(dados, r, cMat - deslCol)) _
           And TextoCel(Celula(dados, r, cNome - deslCol)) <> "" Then
            If CDbl(TextoCel(Celula(dados, r, cTipo - deslCol))) < 1000 And CDbl(TextoCel(Celula(dados, r, cMat - deslCol))) > 0 _
               And Not PareceNumero(Celula(dados, r, cNome - deslCol)) Then
                matAtual = CLng(TextoCel(Celula(dados, r, cMat - deslCol)))
                nomeAtual = TextoCel(Celula(dados, r, cNome - deslCol))
                admAtual = Celula(dados, r, cAdm - deslCol)
                GoTo ProximaLinha
            End If
        End If

        If EhNumero(vIni) And Not IsEmpty(matAtual) Then
            If CDbl(vIni) > 1000 And TextoCel(Celula(dados, r, cDesc - deslCol)) <> "" Then
                nSaida = nSaida + 1
                saida(nSaida, 1) = matAtual
                saida(nSaida, 2) = nomeAtual
                If EhNumero(admAtual) Then saida(nSaida, 3) = CDate(Int(CDbl(admAtual)))
                saida(nSaida, 4) = CodigoTexto(Celula(dados, r, cCod - deslCol), 3)
                saida(nSaida, 5) = TextoCel(Celula(dados, r, cDesc - deslCol))
                saida(nSaida, 6) = CDate(Int(CDbl(vIni)))
                vHi = Celula(dados, r, cHoraIni - deslCol)
                If EhNumero(vHi) Then saida(nSaida, 7) = CDbl(vHi) - Int(CDbl(vHi))
                vFim = Celula(dados, r, cFim - deslCol)
                If EhNumero(vFim) Then
                    If CDbl(vFim) > 1000 Then
                        saida(nSaida, 8) = CDate(Int(CDbl(vFim)))
                        saida(nSaida, 10) = Int(CDbl(vFim)) - Int(CDbl(vIni)) + 1
                    End If
                End If
                vHf = Celula(dados, r, cHoraFim - deslCol)
                If EhNumero(vHf) Then saida(nSaida, 9) = CDbl(vHf) - Int(CDbl(vHf))
                ' horas: afastamento de poucas horas no mesmo dia
                If saida(nSaida, 10) = 1 And EhNumero(vHi) And EhNumero(vHf) Then
                    If CDbl(vHf) - Int(CDbl(vHf)) > CDbl(vHi) - Int(CDbl(vHi)) Then
                        saida(nSaida, 11) = (CDbl(vHf) - Int(CDbl(vHf))) - (CDbl(vHi) - Int(CDbl(vHi)))
                    End If
                End If
                vPrev = Celula(dados, r, cPrev - deslCol)
                If EhNumero(vPrev) Then
                    If CDbl(vPrev) > 1000 Then saida(nSaida, 12) = CDate(Int(CDbl(vPrev)))
                End If
                saida(nSaida, 13) = TextoCel(Celula(dados, r, cExame - deslCol))
            End If
        End If
ProximaLinha:
    Next r

    If nSaida = 0 Then
        MsgBox "Não encontrei nenhum afastamento na aba ativa." & vbCrLf & _
               "Confira se a aba ativa é o relatório Histórico de Afastamentos.", vbExclamation, "Formatar Afastamentos"
        Exit Sub
    End If

    ' ------------------------------------------------------------
    ' 3) Reescreve a aba
    ' ------------------------------------------------------------
    Application.ScreenUpdating = False
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.UnMerge
    On Error GoTo 0
    ws.Cells.Clear

    Dim cab As Variant
    cab = Array("Matrícula", "Colaborador", "Admissão", "Cód. Situação", "Situação", "Início", "Hora início", _
                "Término", "Hora término", "Dias", "Horas", "Prev. Término", "Exame")
    For c = 0 To N_COLS - 1
        ws.Cells(1, c + 1).Value = cab(c)
    Next c

    ws.Range(ws.Cells(2, 4), ws.Cells(nSaida + 1, 4)).NumberFormat = "@"
    ws.Range(ws.Cells(2, 3), ws.Cells(nSaida + 1, 3)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(2, 6), ws.Cells(nSaida + 1, 6)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(2, 8), ws.Cells(nSaida + 1, 8)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(2, 12), ws.Cells(nSaida + 1, 12)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(2, 7), ws.Cells(nSaida + 1, 7)).NumberFormat = "hh:mm"
    ws.Range(ws.Cells(2, 9), ws.Cells(nSaida + 1, 9)).NumberFormat = "hh:mm"
    ws.Range(ws.Cells(2, 11), ws.Cells(nSaida + 1, 11)).NumberFormat = "[h]:mm"

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
    ws.Columns("A").ColumnWidth = 11
    ws.Columns("B").ColumnWidth = 36
    ws.Columns("C").ColumnWidth = 11
    ws.Columns("D").ColumnWidth = 9
    ws.Columns("E").ColumnWidth = 30
    ws.Columns("F:L").ColumnWidth = 11
    ws.Columns("M").ColumnWidth = 7
    ws.Range("C:M").HorizontalAlignment = xlCenter

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

    MsgBox "Formatação concluída." & vbCrLf & _
           nSaida & " afastamento(s) na tabela." & vbCrLf & _
           IIf(renomeada, "Aba renomeada para """ & NOME_ABA_SAIDA & """.", _
               "A aba não foi renomeada (já existe uma aba """ & NOME_ABA_SAIDA & """)."), _
           vbInformation, "Formatar Afastamentos"
End Sub

' Valor do array lido do UsedRange; Empty se a coluna estiver fora dele.
Private Function Celula(dados As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If c < 1 Or c > UBound(dados, 2) Then
        Celula = Empty
    Else
        Celula = dados(r, c)
    End If
End Function

' Número de verdade (não vazio, não texto, não erro).
Private Function EhNumero(v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then
        EhNumero = False
    ElseIf VarType(v) = vbString Then
        EhNumero = False
    Else
        EhNumero = IsNumeric(v)
    End If
End Function

' Número ou texto que é um número ("1", "55277").
Private Function PareceNumero(v As Variant) As Boolean
    Dim t As String
    t = TextoCel(v)
    PareceNumero = (t <> "" And IsNumeric(t))
End Function

' Texto sem espaços nas pontas; erro do Excel (#N/A) vira "".
Private Function TextoCel(v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then
        TextoCel = ""
    Else
        TextoCel = Trim$(Replace(Replace(CStr(v), vbCr, " "), vbLf, " "))
    End If
End Function

' Maiúsculas e sem acento, para comparar cabeçalhos ("Situação" = "SITUACAO").
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

' Código com zeros à esquerda ("002", "014").
Private Function CodigoTexto(v As Variant, ByVal digitos As Long) As String
    If EhNumero(v) Then
        CodigoTexto = Format$(CLng(v), String$(digitos, "0"))
    Else
        CodigoTexto = TextoCel(v)
    End If
End Function
