Attribute VB_Name = "modFormatarPausas"
Option Explicit

' ============================================================================
' FormatarPausasTermicas
'
' DE ONDE VEM: a aba ATIVA, com o "Relatório Analítico de Pausas Térmicas" (HRES114) bruto: um
'   bloco por colaborador (matrícula, nome, cargo), uma linha por dia com as batidas (hora + tipo:
'   P = entrada/saída da jornada, F = pausa térmica, G = refeição), o bloco "Total do Colaborador"
'   e, no fim, o "Total Geral" e a legenda.
'
' PARA ONDE VAI: a mesma aba, reescrita com três tabelas, uma embaixo da outra:
'   1) POR CICLO (cabeçalho na linha 1): uma linha por colaborador e por ciclo de apuração (28 a 27),
'      com os 7 números CALCULADOS por esta macro a partir das linhas de cada dia:
'      Matrícula | Colaborador | Cargo | Fim do ciclo | Pausa corretas | Pausa menor 0:20 |
'      Pausa maior 0:20 | Trabalho correto 1:40 | Trabalho maior 1:40 | Trabalho menor 1:40 |
'      Marcações Ímpares. A linha em branco depois dela marca o fim para o GerarCSVPonto.
'   2) CONFERÊNCIA: o "Total do Colaborador" do próprio relatório (período inteiro) ao lado da soma
'      dos ciclos calculados, com "Confere?" = Sim/Não.
'   3) TOTAL GERAL do relatório.
'
' POR QUE CALCULAR (pedido do Dimitri, 03/11): o relatório vem com 3 meses e o total que ele traz
'   é dos 3 meses juntos; o painel precisa dos números de cada mês.
'
' A REGRA (descoberta comparando com o "Total do Colaborador" de 52 colaboradores, 28/07 a 27/10):
'   - as batidas do dia se alternam saída/volta pela POSIÇÃO; o trecho da 1ª para a 2ª batida, da
'     3ª para a 4ª... é TRABALHO; os outros são intervalo;
'   - intervalo entre duas batidas "F" é PAUSA TÉRMICA: 20 min ou mais = correta, menos = menor
'     (o relatório nunca conta "pausa maior"; a coluna fica, sempre 0);
'   - trabalho de exatamente 1:40 = correto, mais = maior, menos = menor;
'   - dia com número ímpar de batidas = 1 "marcação ímpar", e o último trecho do dia conta como
'     trabalho;
'   - batidas depois da meia-noite pertencem ao dia em que a jornada começou (como no relatório).
'   Resultado: pausas (corretas e menores) e marcações ímpares iguais ao sistema nos 52; os três
'   números de trabalho iguais em 36 dos 52 (nos outros, diferença de 1 a 13 trechos em dias com
'   batidas fora do padrão). A tabela de conferência mostra isso a cada mês.
' ============================================================================

Sub FormatarPausasTermicas()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ultLin As Long, ultCol As Long
    ultLin = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    ultCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If ultLin < 2 Then
        MsgBox "Nenhum dado encontrado na aba ativa.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Este macro vai REESCREVER completamente a aba """ & ws.Name & _
              """, deixando os totais por colaborador e por ciclo (28 a 27), a conferência com o relatório " & _
              "e o total geral." & vbCrLf & vbCrLf & "Continuar?", vbYesNo + vbQuestion, "Formatar Pausas Térmicas") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Lendo o relatório de pausas..."

    Dim dados As Variant
    dados = ws.Range(ws.Cells(1, 1), ws.Cells(ultLin, ultCol)).Value2

    ' por ciclo: chave "matrícula|fim do ciclo" -> Array(mat, nome, cargo, fimCiclo, 7 números)
    Dim dictCiclo As Object, ordemCiclo As Collection
    Set dictCiclo = CreateObject("Scripting.Dictionary")
    Set ordemCiclo = New Collection
    ' conferência: uma linha por bloco "Total do Colaborador"
    Dim confLinhas As Collection
    Set confLinhas = New Collection
    Dim geral(1 To 7) As Variant, temGeral As Boolean

    Dim curMat As String, curNome As String, curCargo As String
    Dim somaPessoa(1 To 7) As Long
    Dim r As Long, k As Long, i As Long, qtdBat As Long
    Dim horas() As Double, tipos() As String
    ReDim horas(1 To ultCol)
    ReDim tipos(1 To ultCol)
    Dim t(1 To 7) As Long, chave As String, fimCiclo As Date, item As Variant
    Dim a1 As Variant, b1 As Variant, c1 As Variant

    For r = 1 To UBound(dados, 1)
        a1 = PCel(dados, r, 1)
        b1 = PCel(dados, r, 2)
        c1 = PCel(dados, r, 3)

        If PNum(a1) And VarType(b1) = vbString And Not PEhTipo(c1) Then
            If Trim$(b1) <> "" And Trim$(b1) <> "-" Then
                ' identificação do colaborador: A = matrícula, B = nome, F = cargo
                curMat = CStr(CLng(a1))
                curNome = Trim$(b1)
                curCargo = PTxt(PCel(dados, r, 6))
                For i = 1 To 7: somaPessoa(i) = 0: Next i
            End If

        ElseIf PNum(a1) And PEhTipo(c1) And curMat <> "" Then
            ' linha do dia: A = data; depois pares (hora, tipo) a partir da coluna B
            qtdBat = 0
            k = 2
            Do While k + 1 <= UBound(dados, 2)
                If Not (PNum(dados(r, k)) And PEhTipo(dados(r, k + 1))) Then Exit Do
                qtdBat = qtdBat + 1
                horas(qtdBat) = CDbl(dados(r, k))
                tipos(qtdBat) = UCase$(Trim$(dados(r, k + 1)))
                k = k + 2
            Loop
            ClassificarDiaPausas horas, tipos, qtdBat, t

            fimCiclo = FimDoCiclo(CDate(Int(CDbl(a1))))
            chave = curMat & "|" & CLng(fimCiclo)
            If Not dictCiclo.Exists(chave) Then
                dictCiclo.Add chave, Array(curMat, curNome, curCargo, fimCiclo, 0, 0, 0, 0, 0, 0, 0)
                ordemCiclo.Add chave
            End If
            item = dictCiclo(chave)
            For i = 1 To 7
                item(3 + i) = item(3 + i) + t(i)
                somaPessoa(i) = somaPessoa(i) + t(i)
            Next i
            dictCiclo(chave) = item

        ElseIf Trim$(PTxt(a1)) = "Total do Colaborador" And curMat <> "" Then
            ' 7 números do sistema 2 linhas abaixo, nas colunas A, B, D, F, H, J, K
            Dim sist As Variant, confere As Boolean, cols As Variant
            cols = Array(1, 2, 4, 6, 8, 10, 11)
            sist = Array(0, 0, 0, 0, 0, 0, 0)
            confere = True
            For i = 0 To 6
                sist(i) = CLng(PNumOuZero(PCel(dados, r + 2, cols(i))))
                If sist(i) <> somaPessoa(i + 1) Then confere = False
            Next i
            confLinhas.Add Array(curMat, curNome, curCargo, sist, _
                Array(somaPessoa(1), somaPessoa(2), somaPessoa(3), somaPessoa(4), somaPessoa(5), somaPessoa(6), somaPessoa(7)), confere)

        ElseIf Trim$(PTxt(a1)) = "Total" And Trim$(PTxt(b1)) = "Geral" Then
            cols = Array(1, 2, 4, 6, 8, 10, 11)
            For i = 0 To 6
                geral(i + 1) = PCel(dados, r + 2, cols(i))
            Next i
            temGeral = True
        End If
    Next r

    If dictCiclo.Count = 0 Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "Não encontrei nenhuma linha de dia (data + batidas P/F/G) na aba ativa.", vbExclamation, _
            "Formatar Pausas Térmicas"
        Exit Sub
    End If

    Application.StatusBar = "Reescrevendo a aba..."
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear

    Dim nomes7 As Variant
    nomes7 = Array("Pausa corretas", "Pausa menor 0:20", "Pausa maior 0:20", "Trabalho correto 1:40", _
                   "Trabalho maior 1:40", "Trabalho menor 1:40", "Marcações Ímpares")

    ' ---------------- 1) por ciclo
    Dim lin As Long, hcol As Long
    ws.Cells(1, 1).Value = "Matrícula"
    ws.Cells(1, 2).Value = "Colaborador"
    ws.Cells(1, 3).Value = "Cargo"
    ws.Cells(1, 4).Value = "Fim do ciclo"
    For hcol = 0 To 6
        ws.Cells(1, 5 + hcol).Value = nomes7(hcol)
    Next hcol
    FormatarCabecalho ws.Range("A1:K1")

    Dim saida() As Variant, n As Long, ch As Variant
    ReDim saida(1 To dictCiclo.Count, 1 To 11)
    n = 0
    For Each ch In ordemCiclo
        item = dictCiclo(ch)
        n = n + 1
        saida(n, 1) = CLng(item(0))
        saida(n, 2) = item(1)
        saida(n, 3) = item(2)
        saida(n, 4) = item(3)
        For i = 1 To 7
            saida(n, 4 + i) = item(3 + i)
        Next i
    Next ch
    ws.Range("A2").Resize(n, 11).Value = saida
    ws.Range("D2").Resize(n, 1).NumberFormat = "dd/mm/yyyy"
    ws.Range("A1").Resize(n + 1, 11).Borders.LineStyle = xlContinuous
    lin = n + 4   ' linha n+2 fica EM BRANCO: é ela que marca o fim da tabela 1 para o GerarCSVPonto

    ' ---------------- 2) conferência com o relatório (período inteiro)
    ws.Cells(lin - 1, 1).Value = "Conferência: total do relatório (período inteiro) x soma dos ciclos calculados"
    ws.Cells(lin - 1, 1).Font.Bold = True
    ws.Cells(lin, 1).Value = "Matrícula"
    ws.Cells(lin, 2).Value = "Colaborador"
    ws.Cells(lin, 3).Value = "Cargo"
    For hcol = 0 To 6
        ws.Cells(lin, 4 + hcol).Value = "Relatório: " & nomes7(hcol)
    Next hcol
    ws.Cells(lin, 11).Value = "Confere?"
    ws.Cells(lin, 12).Value = "Diferenças (relatório x cálculo)"
    FormatarCabecalho ws.Range(ws.Cells(lin, 1), ws.Cells(lin, 12))
    Dim inicioConf As Long, qtdConfere As Long, dif As String, cf As Variant
    inicioConf = lin
    For Each cf In confLinhas
        lin = lin + 1
        ws.Cells(lin, 1).Value = CLng(cf(0))
        ws.Cells(lin, 2).Value = cf(1)
        ws.Cells(lin, 3).Value = cf(2)
        dif = ""
        For i = 0 To 6
            ws.Cells(lin, 4 + i).Value = cf(3)(i)
            If cf(3)(i) <> cf(4)(i) Then dif = dif & IIf(dif = "", "", "; ") & nomes7(i) & " " & cf(3)(i) & " x " & cf(4)(i)
        Next i
        ws.Cells(lin, 11).Value = IIf(cf(5), "Sim", "Não")
        ws.Cells(lin, 12).Value = dif
        If cf(5) Then qtdConfere = qtdConfere + 1
    Next cf
    ws.Range(ws.Cells(inicioConf, 1), ws.Cells(lin, 12)).Borders.LineStyle = xlContinuous

    ' ---------------- 3) total geral do relatório
    If temGeral Then
        lin = lin + 2
        With ws.Range(ws.Cells(lin, 1), ws.Cells(lin, 3))
            .Merge
            .Value = "Total Geral (relatório)"
        End With
        For hcol = 0 To 6
            ws.Cells(lin, 4 + hcol).Value = nomes7(hcol)
            ws.Cells(lin + 1, 4 + hcol).Value = geral(hcol + 1)
        Next hcol
        FormatarCabecalho ws.Range(ws.Cells(lin, 1), ws.Cells(lin, 10))
        With ws.Range(ws.Cells(lin + 1, 1), ws.Cells(lin + 1, 3))
            .Merge
            .Value = "Todos os colaboradores"
        End With
        ws.Range(ws.Cells(lin + 1, 1), ws.Cells(lin + 1, 10)).Font.Bold = True
        ws.Range(ws.Cells(lin, 1), ws.Cells(lin + 1, 10)).Borders.LineStyle = xlContinuous
    End If

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 32
    ws.Columns("C").ColumnWidth = 28
    ws.Columns("D:K").ColumnWidth = 14
    ws.Columns("L").ColumnWidth = 60
    ws.Range("D:K").HorizontalAlignment = xlCenter
    ws.Activate
    ws.Cells(1, 1).Select

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Formatação concluída." & vbCrLf & _
           n & " linha(s) por colaborador e ciclo, de " & confLinhas.Count & " colaborador(es)." & vbCrLf & _
           "Conferência com o total do relatório: " & qtdConfere & " de " & confLinhas.Count & _
           " colaborador(es) com os 7 números iguais (pausas e marcações ímpares costumam bater em todos;" & _
           " veja a coluna ""Diferenças"").", vbInformation, "Formatar Pausas Térmicas"
End Sub

' Classifica as batidas de UM dia (ver a regra no cabeçalho). t(1..7) = pausa correta, pausa menor,
' pausa maior (sempre 0), trabalho correto, trabalho maior, trabalho menor, marcação ímpar.
Private Sub ClassificarDiaPausas(horas() As Double, tipos() As String, ByVal qtd As Long, t() As Long)
    Dim i As Long, base As Long, ant As Long, x As Long, dur As Long, impar As Boolean
    Dim mins() As Long
    For i = 1 To 7: t(i) = 0: Next i
    If qtd = 0 Then Exit Sub
    ReDim mins(1 To qtd)
    base = 0
    For i = 1 To qtd
        x = CLng(Round(horas(i) * 1440, 0))
        If i > 1 Then
            If x + base < ant Then base = base + 1440
        End If
        mins(i) = x + base
        ant = mins(i)
    Next i
    impar = (qtd Mod 2 = 1)
    If impar Then t(7) = 1
    For i = 1 To qtd - 1
        dur = mins(i + 1) - mins(i)
        If (i Mod 2 = 1) Or (impar And i = qtd - 1) Then
            ' trecho de trabalho (1ª->2ª, 3ª->4ª...; no dia ímpar, também o último)
            If dur = 100 Then
                t(4) = t(4) + 1
            ElseIf dur > 100 Then
                t(5) = t(5) + 1
            Else
                t(6) = t(6) + 1
            End If
        ElseIf tipos(i) = "F" And tipos(i + 1) = "F" Then
            If dur >= 20 Then t(1) = t(1) + 1 Else t(2) = t(2) + 1
        End If
    Next i
End Sub

' Último dia do ciclo de apuração (28 a 27) a que a data pertence.
Private Function FimDoCiclo(ByVal d As Date) As Date
    If Day(d) >= 28 Then
        FimDoCiclo = DateSerial(Year(d), Month(d) + 1, 27)
    Else
        FimDoCiclo = DateSerial(Year(d), Month(d), 27)
    End If
End Function

Private Function PCel(dados As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If r < 1 Or r > UBound(dados, 1) Or c < 1 Or c > UBound(dados, 2) Then
        PCel = Empty
    Else
        PCel = dados(r, c)
    End If
End Function

Private Function PNum(v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then
        PNum = False
    ElseIf VarType(v) = vbString Then
        PNum = False
    Else
        PNum = IsNumeric(v)
    End If
End Function

Private Function PNumOuZero(v As Variant) As Double
    If PNum(v) Then PNumOuZero = CDbl(v) Else PNumOuZero = 0
End Function

Private Function PTxt(v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then PTxt = "" Else PTxt = Trim$(CStr(v))
End Function

' Tipo de batida do relatório: P, F ou G.
Private Function PEhTipo(v As Variant) As Boolean
    Dim s As String
    If VarType(v) <> vbString Then Exit Function
    s = UCase$(Trim$(v))
    PEhTipo = (s = "P" Or s = "F" Or s = "G")
End Function

' Visual padrão das linhas de cabeçalho das duas tabelas (negrito, fundo cinza, borda).
Private Sub FormatarCabecalho(rng As Range)
    With rng
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub
