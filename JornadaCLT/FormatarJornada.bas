Attribute VB_Name = "modFormatarJornada"
Option Explicit

' ============================================================================
' FormatarJornada: os 4 relatórios de jornada do sistema (Senior), cada um na sua aba
'
'   FormatarExcedentes           "Excedentes"            (mais de 10h no dia)
'   FormatarIntrajornada         "Intrajornada" HRES108  (intervalo curto de almoço)
'   FormatarInterjornada         "Interjornada"          (menos de 11h de descanso)
'   FormatarInterjornadaSemanal  "Interjornada semanal"  (trabalho na folga ou feriado)
'
' DE ONDE VEM: a aba ATIVA, com o relatório como o sistema exporta (dump de impressão:
'   cabeçalho de página repetido e valores deslocados em relação aos títulos).
' PARA ONDE VAI: a mesma aba, reescrita como UMA tabela (cabeçalho na linha 1):
'   Matrícula | Colaborador | Cargo | Local | Data | Ocorrência | Minutos | Detalhe
'   Ocorrência = nome técnico; Minutos depende do relatório:
'     Horas excedentes     -> jornada do dia (horas trabalhadas + adicional noturno, igual à
'                             coluna "Total" do relatório; o painel mostra quanto passou de 10h)
'     Intrajornada         -> soma dos intervalos do dia (0 = sem intervalo)
'     Interjornada         -> horas descansadas entre uma jornada e a outra
'     Interjornada semanal -> 0 (o relatório só diz o dia trabalhado no DSR/feriado)
'   Detalhe = texto com as marcações e os números do relatório, para o painel mostrar.
' O GerarCSVPonto acha essas abas pelo cabeçalho (Matrícula, Data, Ocorrência, Minutos, Detalhe)
' e grava uma linha "Jornada" por registro.
'
' Os títulos são procurados pelo texto; a posição dos valores em relação aos títulos foi tirada dos
' relatórios de 28/07 a 27/10/2026:
'   Excedentes:   matrícula = col. "Colaborador/Nome", nome +1; carga = "Carga Hor." +1; marcações a
'                 partir de "Marcações" +1; os 4 últimos números da linha = extras, horas, adicional
'                 noturno e total.
'   Intrajornada: marcações = "Registros de Ponto" até +7; jornada = +8; intervalos = +9 a +12.
'   Interjornada: data = "Data Apuração" +1; marcação anterior/atual = título +1 (data) e +2 (hora);
'                 horas descansadas e ocorrência = título +1; local e cargo = na coluna do título.
'   Semanal:      matrícula/nome só na 1ª linha da pessoa; o dia vem no texto "(dd/mm/aaaa)".
' ============================================================================

Sub FormatarExcedentes()
    FormatarJornadaAba "EXCEDENTES", "Formatar Horas Excedentes"
End Sub

Sub FormatarIntrajornada()
    FormatarJornadaAba "INTRA", "Formatar Intrajornada"
End Sub

Sub FormatarInterjornada()
    FormatarJornadaAba "INTER", "Formatar Interjornada"
End Sub

Sub FormatarInterjornadaSemanal()
    FormatarJornadaAba "SEMANAL", "Formatar Interjornada Semanal"
End Sub

Private Sub FormatarJornadaAba(ByVal tipoRel As String, ByVal tituloMsg As String)
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ultLin As Long, ultCol As Long
    ultLin = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    ultCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If ultLin < 2 Or ultCol < 2 Then
        MsgBox "Nenhum dado encontrado na aba ativa.", vbExclamation, tituloMsg
        Exit Sub
    End If

    Dim dados As Variant
    dados = ws.Range(ws.Cells(1, 1), ws.Cells(ultLin, ultCol)).Value2

    Dim saida() As Variant, n As Long
    ReDim saida(1 To ultLin, 1 To 8)
    Select Case tipoRel
        Case "EXCEDENTES": n = JLerExcedentes(dados, saida)
        Case "INTRA": n = JLerIntrajornada(dados, saida)
        Case "INTER": n = JLerInterjornada(dados, saida)
        Case Else: n = JLerSemanal(dados, saida)
    End Select

    If n < 0 Then
        MsgBox "Não encontrei os títulos do relatório na aba ativa (confira se é o relatório certo).", _
            vbExclamation, tituloMsg
        Exit Sub
    End If

    Application.ScreenUpdating = False
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear

    Dim cab As Variant, i As Long
    cab = Array("Matrícula", "Colaborador", "Cargo", "Local", "Data", "Ocorrência", "Minutos", "Detalhe")
    For i = 0 To 7
        ws.Cells(1, i + 1).Value = cab(i)
    Next i
    With ws.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(26, 26, 26)
        .Font.Color = RGB(255, 255, 255)
    End With
    If n > 0 Then
        ws.Range("A2").Resize(n, 8).Value = saida
        ws.Range("E2").Resize(n, 1).NumberFormat = "dd/mm/yyyy"
        ws.Range("G2").Resize(n, 1).NumberFormat = "0"
        ws.Range("A1").Resize(n + 1, 8).AutoFilter
    End If
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Columns("A:G").AutoFit
    ws.Columns("H").ColumnWidth = 90
    Application.ScreenUpdating = True

    MsgBox "Formatação concluída: " & n & " registro(s) na tabela.", vbInformation, tituloMsg
End Sub

' ---------------------------------------------------------------------------- Excedentes
Private Function JLerExcedentes(dados As Variant, saida() As Variant) As Long
    Dim rh As Long
    rh = JLinhaCab(dados, Array("Data", "Colaborador/Nome", "Marcações"))
    If rh = 0 Then
        JLerExcedentes = -1
        Exit Function
    End If
    Dim cData As Long, cMat As Long, cCargo As Long, cLocal As Long, cCarga As Long, cMarc As Long
    cData = JAchar(dados, rh, "Data")
    cMat = JAchar(dados, rh, "Colaborador/Nome")
    cCargo = JAchar(dados, rh, "Cargo")
    cLocal = JAchar(dados, rh, "Local")
    cCarga = JAchar(dados, rh, "Carga Hor.") + 1
    cMarc = JAchar(dados, rh, "Marcações") + 1

    Dim r As Long, c As Long, n As Long, k As Long
    Dim nums() As Double, qtd As Long, marcTxt As String
    ReDim nums(1 To UBound(dados, 2))
    For r = 1 To UBound(dados, 1)
        If JNum(JCel(dados, r, cData)) And JNum(JCel(dados, r, cMat)) And _
           VarType(JCel(dados, r, cMat + 1)) = vbString Then
            qtd = 0
            For c = cMarc To UBound(dados, 2)
                If JNum(dados(r, c)) Then
                    qtd = qtd + 1
                    nums(qtd) = CDbl(dados(r, c))
                End If
            Next c
            If qtd >= 4 Then
                marcTxt = ""
                For k = 1 To qtd - 4
                    marcTxt = marcTxt & IIf(marcTxt = "", "", " ") & JHora(nums(k))
                Next k
                n = n + 1
                saida(n, 1) = CLng(dados(r, cMat))
                saida(n, 2) = JTxt(dados(r, cMat + 1))
                saida(n, 3) = JTxt(JCel(dados, r, cCargo))
                saida(n, 4) = JTxt(JCel(dados, r, cLocal))
                saida(n, 5) = CDate(Int(CDbl(dados(r, cData))))
                saida(n, 6) = "Horas excedentes"
                saida(n, 7) = CLng(Round(nums(qtd) * 1440, 0))
                saida(n, 8) = "Jornada " & JDur(nums(qtd)) & " (" & JDur(nums(qtd - 2)) & " trabalhadas + " & _
                    JDur(nums(qtd - 1)) & " de adicional noturno) · carga " & JDur(JNumOuZero(JCel(dados, r, cCarga))) & _
                    " · extras " & JDur(nums(qtd - 3)) & " · marcações " & marcTxt
            End If
        End If
    Next r
    JLerExcedentes = n
End Function

' ---------------------------------------------------------------------------- Intrajornada
Private Function JLerIntrajornada(dados As Variant, saida() As Variant) As Long
    Dim rh As Long
    rh = JLinhaCab(dados, Array("Cadastro", "Registros de Ponto"))
    If rh = 0 Then
        JLerIntrajornada = -1
        Exit Function
    End If
    Dim cMat As Long, cNome As Long, cLocal As Long, cCargo As Long, cData As Long, cDia As Long, cRP As Long
    cMat = JAchar(dados, rh, "Cadastro")
    cNome = JAchar(dados, rh, "Nome")
    cLocal = JAchar(dados, rh, "Local")
    cCargo = JAchar(dados, rh, "Cargo")
    cData = JAchar(dados, rh, "Data")
    cDia = JAchar(dados, rh, "Dia")
    cRP = JAchar(dados, rh, "Registros de Ponto")

    Dim r As Long, c As Long, n As Long
    Dim marcTxt As String, intTxt As String, somaInt As Double, trab As Double
    For r = 1 To UBound(dados, 1)
        If JNum(JCel(dados, r, cMat)) And VarType(JCel(dados, r, cNome)) = vbString And JNum(JCel(dados, r, cData)) Then
            marcTxt = ""
            For c = cRP To cRP + 7
                If JNum(JCel(dados, r, c)) Then marcTxt = marcTxt & IIf(marcTxt = "", "", " ") & JHora(CDbl(dados(r, c)))
            Next c
            trab = JNumOuZero(JCel(dados, r, cRP + 8))
            intTxt = ""
            somaInt = 0
            For c = cRP + 9 To cRP + 12
                If JNum(JCel(dados, r, c)) Then
                    somaInt = somaInt + CDbl(dados(r, c))
                    intTxt = intTxt & IIf(intTxt = "", "", " + ") & JDur(CDbl(dados(r, c)))
                End If
            Next c
            n = n + 1
            saida(n, 1) = CLng(dados(r, cMat))
            saida(n, 2) = JTxt(dados(r, cNome))
            saida(n, 3) = JTxt(JCel(dados, r, cCargo))
            saida(n, 4) = JTxt(JCel(dados, r, cLocal))
            saida(n, 5) = CDate(Int(CDbl(dados(r, cData))))
            saida(n, 6) = "Intrajornada"
            saida(n, 7) = CLng(Round(somaInt * 1440, 0))
            saida(n, 8) = "Trabalhou " & JDur(trab) & " " & _
                IIf(intTxt = "", "sem intervalo", "com intervalo(s) de " & intTxt) & _
                " · marcações " & marcTxt & " · horário " & JTxt(JCel(dados, r, cDia))
        End If
    Next r
    JLerIntrajornada = n
End Function

' ---------------------------------------------------------------------------- Interjornada
Private Function JLerInterjornada(dados As Variant, saida() As Variant) As Long
    Dim rh As Long
    rh = JLinhaCab(dados, Array("Data Apuração", "Horas Descansadas"))
    If rh = 0 Then
        JLerInterjornada = -1
        Exit Function
    End If
    Dim cApur As Long, cAnt As Long, cAtu As Long, cDesc As Long, cOc As Long, cLocal As Long, cCargo As Long
    cApur = JAchar(dados, rh, "Data Apuração") + 1
    cAnt = JAchar(dados, rh, "Marcação Anterior") + 1
    cAtu = JAchar(dados, rh, "Marcação Atual") + 1
    cDesc = JAchar(dados, rh, "Horas Descansadas") + 1
    cOc = JAchar(dados, rh, "Ocorrência") + 1
    cLocal = JAchar(dados, rh, "Local")
    cCargo = JAchar(dados, rh, "Cargo")

    Dim r As Long, n As Long, desc As Double
    For r = rh + 1 To UBound(dados, 1)
        If JNum(JCel(dados, r, 1)) And VarType(JCel(dados, r, 2)) = vbString And JNum(JCel(dados, r, cApur)) Then
            desc = JNumOuZero(JCel(dados, r, cDesc))
            n = n + 1
            saida(n, 1) = CLng(dados(r, 1))
            saida(n, 2) = JTxt(dados(r, 2))
            saida(n, 3) = JTxt(JCel(dados, r, cCargo))
            saida(n, 4) = JTxt(JCel(dados, r, cLocal))
            saida(n, 5) = CDate(Int(CDbl(dados(r, cApur))))
            saida(n, 6) = "Interjornada"
            saida(n, 7) = CLng(Round(desc * 1440, 0))
            saida(n, 8) = "Descansou " & JDur(desc) & " entre " & JDiaMes(JCel(dados, r, cAnt)) & " " & _
                JHora(JNumOuZero(JCel(dados, r, cAnt + 1))) & " e " & JDiaMes(JCel(dados, r, cAtu)) & " " & _
                JHora(JNumOuZero(JCel(dados, r, cAtu + 1))) & " · " & JTxt(JCel(dados, r, cOc))
        End If
    Next r
    JLerInterjornada = n
End Function

' ---------------------------------------------------------------------------- Interjornada semanal
Private Function JLerSemanal(dados As Variant, saida() As Variant) As Long
    Dim rh As Long
    rh = JLinhaCab(dados, Array("Colaborador", "Horas DSR"))
    If rh = 0 Then
        JLerSemanal = -1
        Exit Function
    End If
    Dim cLocal As Long, cTxt As Long
    cLocal = JAchar(dados, rh, "Local")
    cTxt = JAchar(dados, rh, "Horas DSR")

    Dim r As Long, c As Long, n As Long, t As String, extra As String, p As Long
    Dim matAtual As Variant, nomeAtual As String, dia As Date
    For r = rh + 1 To UBound(dados, 1)
        t = JTxt(JCel(dados, r, cTxt))
        p = InStrRev(t, "(")
        If p > 0 And Len(t) >= p + 11 Then
            If Mid$(t, p + 3, 1) = "/" And Mid$(t, p + 6, 1) = "/" And IsNumeric(Mid$(t, p + 1, 2)) And _
               IsNumeric(Mid$(t, p + 4, 2)) And IsNumeric(Mid$(t, p + 7, 4)) Then
                dia = DateSerial(CInt(Mid$(t, p + 7, 4)), CInt(Mid$(t, p + 4, 2)), CInt(Mid$(t, p + 1, 2)))
                If JNum(JCel(dados, r, 1)) Then
                    matAtual = CLng(dados(r, 1))
                    nomeAtual = JTxt(JCel(dados, r, 2))
                End If
                extra = ""
                For c = cTxt + 1 To UBound(dados, 2)
                    If JNum(dados(r, c)) Then
                        extra = extra & " · " & JHora(CDbl(dados(r, c)))
                    ElseIf JTxt(dados(r, c)) <> "" Then
                        extra = extra & " · " & JTxt(dados(r, c))
                    End If
                Next c
                n = n + 1
                saida(n, 1) = matAtual
                saida(n, 2) = nomeAtual
                saida(n, 3) = ""
                saida(n, 4) = JSemCodigo(JTxt(JCel(dados, r, cLocal)))
                saida(n, 5) = dia
                saida(n, 6) = "Interjornada semanal"
                saida(n, 7) = 0
                saida(n, 8) = t & extra
            End If
        End If
    Next r
    JLerSemanal = n
End Function

' ---------------------------------------------------------------------------- auxiliares
Private Function JCel(dados As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If c < 1 Or c > UBound(dados, 2) Or r < 1 Or r > UBound(dados, 1) Then
        JCel = Empty
    Else
        JCel = dados(r, c)
    End If
End Function

' Número de verdade (não vazio, não texto, não erro).
Private Function JNum(v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then
        JNum = False
    ElseIf VarType(v) = vbString Then
        JNum = False
    Else
        JNum = IsNumeric(v)
    End If
End Function

Private Function JNumOuZero(v As Variant) As Double
    If JNum(v) Then JNumOuZero = CDbl(v) Else JNumOuZero = 0
End Function

Private Function JTxt(v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then
        JTxt = ""
    Else
        JTxt = Trim$(Replace(Replace(CStr(v), vbCr, " "), vbLf, " "))
    End If
End Function

' Coluna do título (comparação sem diferenciar maiúsculas) na linha r; 0 = não achou.
Private Function JAchar(dados As Variant, ByVal r As Long, ByVal rotulo As String) As Long
    Dim c As Long
    For c = 1 To UBound(dados, 2)
        If VarType(dados(r, c)) = vbString Then
            If LCase$(Trim$(dados(r, c))) = LCase$(rotulo) Then
                JAchar = c
                Exit Function
            End If
        End If
    Next c
    JAchar = 0
End Function

' Primeira linha que tem todos os títulos; 0 = não achou.
Private Function JLinhaCab(dados As Variant, rotulos As Variant) As Long
    Dim r As Long, i As Long, tem As Boolean
    For r = 1 To UBound(dados, 1)
        tem = True
        For i = LBound(rotulos) To UBound(rotulos)
            If JAchar(dados, r, CStr(rotulos(i))) = 0 Then
                tem = False
                Exit For
            End If
        Next i
        If tem Then
            JLinhaCab = r
            Exit Function
        End If
    Next r
    JLinhaCab = 0
End Function

' Fração do dia -> "hh:mm" (hora do relógio).
Private Function JHora(ByVal f As Double) As String
    Dim m As Long
    m = CLng(Round(f * 1440, 0)) Mod 1440
    JHora = Format$(m \ 60, "00") & ":" & Format$(m Mod 60, "00")
End Function

' Fração do dia -> "h:mm" (duração).
Private Function JDur(ByVal f As Double) As String
    Dim m As Long
    m = CLng(Round(f * 1440, 0))
    JDur = (m \ 60) & ":" & Format$(m Mod 60, "00")
End Function

Private Function JDiaMes(v As Variant) As String
    If JNum(v) Then JDiaMes = Format$(CDate(Int(CDbl(v))), "dd/mm") Else JDiaMes = ""
End Function

' "2.028.30.304.3041 - Ativ. Operacional 1ºT" -> "Ativ. Operacional 1ºT"
Private Function JSemCodigo(ByVal t As String) As String
    Dim p As Long, prefixo As String, i As Long, ch As String
    p = InStr(t, " - ")
    If p > 1 Then
        prefixo = Left$(t, p - 1)
        For i = 1 To Len(prefixo)
            ch = Mid$(prefixo, i, 1)
            If Not (ch Like "#" Or ch = ".") Then
                JSemCodigo = t
                Exit Function
            End If
        Next i
        JSemCodigo = Trim$(Mid$(t, p + 3))
    Else
        JSemCodigo = t
    End If
End Function
