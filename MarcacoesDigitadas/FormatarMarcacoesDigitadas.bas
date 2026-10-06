Attribute VB_Name = "modFormatarMarcacoesDigitadas"
Option Explicit

' ============================================================================
' FormatarMarcacoesDigitadas
'
' DE ONDE VEM: a aba ATIVA no momento em que a macro é executada. Espera-se
'   que seja o relatório "Marcações digitadas" do sistema de ponto, exportado
'   como "dump" de impressão:
'     - cabeçalho de página repetido ("Crachá | Colaborador | Cargo | Local |
'       Ori. | Data | Hora | Coletor | Função | Justificativa");
'     - uma linha por marcação digitada, com os valores DESLOCADOS em relação
'       ao cabeçalho: o nome fica uma coluna à direita de "Colaborador"; o
'       cargo vem em código (coluna de "Cargo") + descrição (coluna seguinte);
'       a justificativa vem em duas partes, o MOTIVO (coluna de
'       "Justificativa", ex.: "Inclusão por esquecimento") e o TEXTO livre
'       digitado (duas colunas à direita, sem cabeçalho);
'     - quando o motivo é longo, o relatório quebra o texto e o resto vai
'       para a LINHA DE BAIXO, sozinho (ex.: "Problemas no dispositivo de
'       registro de" + linha seguinte "ponto").
'
' PARA ONDE VAI: reescreve essa MESMA aba como uma tabela simples, com o
'   cabeçalho na linha 1 e uma linha por marcação, e renomeia a aba para
'   "Marcações Digitadas" (se o nome estiver livre). Colunas:
'     Matrícula · Colaborador · Cód. Cargo · Cargo · Cód. Local · Local ·
'     Origem · Data · Hora · Dia da semana · Coletor · Função · Motivo ·
'     Justificativa
'   O GerarCSVPonto acha esta aba pelo cabeçalho (Matrícula + Motivo +
'   Justificativa + Data + Hora), não pelo nome.
'
' O QUE NÃO FAZ: não recalcula nem interpreta nada — só reorganiza. As linhas
'   de cabeçalho de página são descartadas; a continuação do motivo é colada
'   no fim do motivo da marcação anterior.
'
' Rode numa CÓPIA do arquivo exportado: a aba ativa é reescrita por completo.
' ============================================================================

Private Const NOME_ABA_SAIDA As String = "Marcações Digitadas"

Sub FormatarMarcacoesDigitadas()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim usado As Range
    Set usado = ws.UsedRange
    If usado.Rows.Count < 2 Then
        MsgBox "Nenhum dado encontrado na aba ativa.", vbExclamation, "Formatar Marcações Digitadas"
        Exit Sub
    End If

    ' Lê tudo de uma vez (bem mais rápido que célula a célula). Value2 (e não
    ' Value) traz data e hora como número: com Value, uma célula formatada
    ' como data vem como Date e IsNumeric(Date) dá False no VBA. O índice do
    ' array começa na 1ª linha/coluna do UsedRange; desl* convertem para o
    ' número real de linha/coluna da aba.
    Dim dados As Variant
    dados = usado.Value2
    Dim deslLin As Long, deslCol As Long
    deslLin = usado.Row - 1
    deslCol = usado.Column - 1
    Dim nLin As Long, nCol As Long
    nLin = UBound(dados, 1)
    nCol = UBound(dados, 2)

    ' ------------------------------------------------------------
    ' 1) Acha a linha de cabeçalho de página ("Crachá" na 1ª coluna com
    '    texto) e, a partir dela, a coluna de cada campo. Se não achar, usa as
    '    posições do relatório de 28/08 a 27/09 (A, C, F, G, I, J, K, L, M, N,
    '    O, Q).
    ' ------------------------------------------------------------
    Dim cCracha As Long, cNome As Long, cCargoCod As Long, cCargo As Long, cLocal As Long
    Dim cOrigem As Long, cData As Long, cHora As Long, cColetor As Long, cFuncao As Long
    Dim cMotivo As Long, cTexto As Long
    cCracha = 1: cNome = 3: cCargoCod = 6: cCargo = 7: cLocal = 9: cOrigem = 10
    cData = 11: cHora = 12: cColetor = 13: cFuncao = 14: cMotivo = 15: cTexto = 17

    Dim r As Long, c As Long, achouCabecalho As Boolean
    For r = 1 To IIf(nLin < 30, nLin, 30)
        For c = 1 To nCol
            If TextoSimples(dados(r, c)) = "CRACHA" Then
                achouCabecalho = True
                Exit For
            End If
        Next c
        If achouCabecalho Then Exit For
    Next r
    If achouCabecalho Then
        Dim t As String
        For c = 1 To nCol
            t = TextoSimples(dados(r, c))
            Select Case t
                Case "CRACHA": cCracha = c + deslCol
                Case "COLABORADOR": cNome = c + deslCol + 1      ' o nome vem uma coluna à direita
                Case "CARGO": cCargoCod = c + deslCol: cCargo = c + deslCol + 1
                Case "LOCAL": cLocal = c + deslCol
                Case "ORI.", "ORI", "ORIGEM": cOrigem = c + deslCol
                Case "DATA": cData = c + deslCol
                Case "HORA": cHora = c + deslCol
                Case "COLETOR": cColetor = c + deslCol
                Case "FUNCAO": cFuncao = c + deslCol
                Case "JUSTIFICATIVA": cMotivo = c + deslCol: cTexto = c + deslCol + 2
            End Select
        Next c
    End If

    ' ------------------------------------------------------------
    ' 2) Monta as linhas de saída. Linha de dado = crachá numérico + data
    '    válida. Linha só com texto na coluna do motivo = continuação do
    '    motivo da marcação anterior.
    ' ------------------------------------------------------------
    Const N_COLS As Long = 14
    Dim saida() As Variant
    ReDim saida(1 To nLin, 1 To N_COLS)
    Dim nSaida As Long, nContinuacoes As Long
    nSaida = 0

    Dim vCracha As Variant, vData As Variant, vHora As Variant, localTxt As String, p As Long
    For r = 1 To nLin
        vCracha = Celula(dados, r, cCracha - deslCol)
        vData = Celula(dados, r, cData - deslCol)

        If IsNumeric(vCracha) And Not IsEmpty(vCracha) And IsNumeric(vData) And Not IsEmpty(vData) Then
            If CDbl(vData) > 0 Then
                nSaida = nSaida + 1
                saida(nSaida, 1) = CLng(vCracha)
                saida(nSaida, 2) = TextoCel(Celula(dados, r, cNome - deslCol))
                saida(nSaida, 3) = TextoCel(Celula(dados, r, cCargoCod - deslCol))
                saida(nSaida, 4) = TextoCel(Celula(dados, r, cCargo - deslCol))
                ' "2.028.30.304.3041 - Ativ. Operacional 1ºT" -> código | descrição
                localTxt = TextoCel(Celula(dados, r, cLocal - deslCol))
                p = InStr(1, localTxt, " - ")
                If p > 0 Then
                    saida(nSaida, 5) = Trim$(Left$(localTxt, p - 1))
                    saida(nSaida, 6) = Trim$(Mid$(localTxt, p + 3))
                Else
                    saida(nSaida, 5) = ""
                    saida(nSaida, 6) = localTxt
                End If
                saida(nSaida, 7) = TextoCel(Celula(dados, r, cOrigem - deslCol))
                saida(nSaida, 8) = CDate(Int(CDbl(vData)))
                vHora = Celula(dados, r, cHora - deslCol)
                If IsNumeric(vHora) And Not IsEmpty(vHora) Then
                    saida(nSaida, 9) = CDbl(vHora) - Int(CDbl(vHora))
                Else
                    saida(nSaida, 9) = Empty
                End If
                saida(nSaida, 10) = DiaDaSemana(CDate(saida(nSaida, 8)))
                saida(nSaida, 11) = CodigoTexto(Celula(dados, r, cColetor - deslCol), 3)
                saida(nSaida, 12) = CodigoTexto(Celula(dados, r, cFuncao - deslCol), 2)
                saida(nSaida, 13) = TextoCel(Celula(dados, r, cMotivo - deslCol))
                saida(nSaida, 14) = TextoCel(Celula(dados, r, cTexto - deslCol))
            End If
        ElseIf nSaida > 0 And TextoCel(vCracha) = "" Then
            ' continuação do texto quebrado (motivo e, se houver, justificativa)
            Dim contMotivo As String, contTexto As String
            contMotivo = TextoCel(Celula(dados, r, cMotivo - deslCol))
            contTexto = TextoCel(Celula(dados, r, cTexto - deslCol))
            If contMotivo <> "" And TextoSimples(contMotivo) <> "JUSTIFICATIVA" Then
                saida(nSaida, 13) = Trim$(saida(nSaida, 13) & " " & contMotivo)
                nContinuacoes = nContinuacoes + 1
            End If
            If contTexto <> "" Then saida(nSaida, 14) = Trim$(saida(nSaida, 14) & " " & contTexto)
        End If
    Next r

    If nSaida = 0 Then
        MsgBox "Não encontrei nenhuma marcação (linha com crachá numérico e data) na aba ativa." & vbCrLf & _
               "Confira se a aba ativa é o relatório de Marcações digitadas.", vbExclamation, "Formatar Marcações Digitadas"
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
    cab = Array("Matrícula", "Colaborador", "Cód. Cargo", "Cargo", "Cód. Local", "Local", "Origem", _
                "Data", "Hora", "Dia da semana", "Coletor", "Função", "Motivo", "Justificativa")
    For c = 0 To N_COLS - 1
        ws.Cells(1, c + 1).Value = cab(c)
    Next c

    ' códigos com zero à esquerda ficam como texto
    ws.Range(ws.Cells(2, 3), ws.Cells(nSaida + 1, 3)).NumberFormat = "@"
    ws.Range(ws.Cells(2, 5), ws.Cells(nSaida + 1, 5)).NumberFormat = "@"
    ws.Range(ws.Cells(2, 11), ws.Cells(nSaida + 1, 12)).NumberFormat = "@"
    ws.Range(ws.Cells(2, 8), ws.Cells(nSaida + 1, 8)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(2, 9), ws.Cells(nSaida + 1, 9)).NumberFormat = "hh:mm"

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
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 9
    ws.Columns("D").ColumnWidth = 26
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 28
    ws.Columns("G").ColumnWidth = 8
    ws.Columns("H").ColumnWidth = 11
    ws.Columns("I").ColumnWidth = 8
    ws.Columns("J").ColumnWidth = 9
    ws.Columns("K").ColumnWidth = 8
    ws.Columns("L").ColumnWidth = 8
    ws.Columns("M").ColumnWidth = 36
    ws.Columns("N").ColumnWidth = 60
    ws.Range("G:L").HorizontalAlignment = xlCenter

    ' renomeia a aba, se o nome estiver livre
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
           nSaida & " marcação(ões) digitada(s) na tabela." & vbCrLf & _
           nContinuacoes & " motivo(s) que estavam quebrados em duas linhas foram juntados." & vbCrLf & _
           IIf(renomeada, "Aba renomeada para """ & NOME_ABA_SAIDA & """.", _
               "A aba não foi renomeada (já existe uma aba """ & NOME_ABA_SAIDA & """)."), _
           vbInformation, "Formatar Marcações Digitadas"
End Sub

' Valor do array lido do UsedRange; Empty se a coluna estiver fora dele.
Private Function Celula(dados As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If c < 1 Or c > UBound(dados, 2) Then
        Celula = Empty
    Else
        Celula = dados(r, c)
    End If
End Function

' Texto sem espaços nas pontas; erro do Excel (#N/A) vira "".
Private Function TextoCel(v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then
        TextoCel = ""
    Else
        TextoCel = Trim$(Replace(Replace(CStr(v), vbCr, " "), vbLf, " "))
    End If
End Function

' Texto em maiúsculas e sem acento, para comparar cabeçalhos ("Crachá" =
' "CRACHA", "Função" = "FUNCAO").
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

' Código numérico com zeros à esquerda (o relatório grava o coletor "000"
' como o número 0).
Private Function CodigoTexto(v As Variant, ByVal digitos As Long) As String
    If IsNumeric(v) And Not IsEmpty(v) Then
        CodigoTexto = Format$(CLng(v), String$(digitos, "0"))
    Else
        CodigoTexto = TextoCel(v)
    End If
End Function

Private Function DiaDaSemana(d As Date) As String
    DiaDaSemana = Choose(Weekday(d, vbSunday), "DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SAB")
End Function
