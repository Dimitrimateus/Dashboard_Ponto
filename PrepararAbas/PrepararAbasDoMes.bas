Attribute VB_Name = "modPrepararAbasDoMes"
Option Explicit

' ================================================================================
' MACRO: PrepararAbasDoMes  (31/10)
' --------------------------------------------------------------------------------
' OBJETIVO
'   Formatar, de uma vez, TODOS os relatórios brutos do sistema que foram colados
'   como abas na planilha de Tratamento do mês, e deixar uma aba "Relatório da
'   preparação" dizendo, aba por aba, o que foi feito, se deu certo e o que deu
'   errado. Depois dela, é só rodar a GerarAbaCSV / ExportarCSVPorGestor.
'
' COMO USAR
'   1) Cole na planilha os módulos de todas as macros (cada .txt num módulo):
'      ConsolidarCartaoPonto, FormatarPausasTermicas, FormatarMarcacoesDigitadas,
'      FormatarAfastamentos, FormatarJornada e este. (O GerarCSVPonto também, para
'      o passo seguinte.)
'   2) Para cada relatório baixado do sistema, copie a aba dele para dentro da
'      planilha (botão direito na aba > Mover ou copiar > planilha de Tratamento,
'      marcando "Criar uma cópia"). O nome da aba não importa. O PontoNet (.csv)
'      também: abra o .csv no Excel e copie a aba.
'   3) Alt+F8 > PrepararAbasDoMes > Executar.
'
' COMO ELA RECONHECE CADA RELATÓRIO (pelo conteúdo, não pelo nome da aba)
'   Cartão Ponto (HRCP102) ......... "Cartão Ponto" + linhas "Empregado:"
'   Pausas térmicas (HRES114) ...... título com "Pausas Térmicas" (sem a tabela
'                                    já formatada "Pausa corretas" na linha 1)
'   Marcações digitadas ............ cabeçalhos "Crachá" e "Coletor"
'   Afastamentos (HRCL006) ......... "Histórico de Afastamentos" ou "Prev Término"
'   Jornada (intra, inter, inter
'   semanal, horas excedentes) ..... "Registros de Ponto", "Horas Descansadas",
'                                    "Horas DSR"/"Marcação Antes DSR" ou "Carga Hor."
'   PontoNet (Ausência de marcação)  A1 começa com "Data falta"
'   Aba que já está formatada, ou que não é relatório (RE, ORG, Tratamento...),
'   não é mexida.
'
' O QUE ACONTECE COM CADA UM
'   Cada relatório é formatado pela macro dele, no "modo automático" (as funções
'   <Macro>_Auto, que fazem o mesmo trabalho sem mostrar mensagens e devolvem o
'   texto da mensagem). Particularidades:
'   - Cartão Ponto: a macro dele cria a aba "Cartao_Consolidado"; aqui ela é
'     renomeada para "Cartão <Mês>" (pelo último dia do período) e a aba bruta é
'     apagada. Se já existe "Cartão <Mês>", a nova fica como "Cartão <Mês> (novo)"
'     e o relatório avisa.
'   - Pausas térmicas: com mais de um mês no mesmo relatório, sai uma aba por mês
'     (ver FormatarPausasTermicas).
'   - PontoNet: se o .csv veio todo na coluna A (separado por ";"), é separado em
'     colunas; a aba vira "Pontonet".
'   Um erro numa aba não para as outras: fica anotado no relatório.
'
' O RELATÓRIO (aba "Relatório da preparação", refeita a cada execução)
'   Aba original | O que era | Resultado | Aba(s) gerada(s) | Detalhes
'   e, embaixo, a lista do que se espera num mês completo, com o que FALTOU
'   (ex.: "Afastamentos: não encontrei nenhum relatório").
'
' Nunca rodou no Excel: rode numa CÓPIA da planilha na primeira vez.
' ================================================================================

Private Const NOME_RELATORIO As String = "Relatório da preparação"

Sub PrepararAbasDoMes()
    Dim wb As Workbook, abas As Collection, ws As Worksheet, i As Long
    Dim tipo As String, resultado As String, detalhe As String, geradas As String
    Dim linhas As Collection, vistos As Object
    Set wb = ThisWorkbook

    If MsgBox("Esta macro vai FORMATAR todas as abas de relatório do sistema que estão nesta planilha " & _
              "(Cartão Ponto, Pausas Térmicas, Marcações digitadas, Afastamentos, Jornada e PontoNet)." & vbCrLf & vbCrLf & _
              "As abas que já estão formatadas e as outras (RE, ORG, Tratamento...) não são mexidas." & vbCrLf & _
              "No fim, a aba """ & NOME_RELATORIO & """ mostra o que foi feito em cada uma." & vbCrLf & vbCrLf & _
              "Recomendado: rode numa CÓPIA da planilha na primeira vez. Continuar?", _
              vbYesNo + vbQuestion, "Preparar as abas do mês") <> vbYes Then Exit Sub

    ' guarda a lista de abas antes de começar (as macros criam, renomeiam e apagam abas)
    Set abas = New Collection
    For Each ws In wb.Worksheets
        If ws.Name <> NOME_RELATORIO Then abas.Add ws
    Next ws

    Set linhas = New Collection
    Set vistos = CreateObject("Scripting.Dictionary")
    Application.ScreenUpdating = False

    For i = 1 To abas.Count
        Set ws = abas(i)
        tipo = TipoDoRelatorioBruto(ws)
        If tipo <> "" Then
            Application.StatusBar = "Preparando a aba " & ws.Name & " (" & NomeDoTipo(tipo) & ")..."
            vistos(tipo) = True
            Dim nomeOriginal As String
            nomeOriginal = ws.Name
            geradas = "": detalhe = ""
            Dim antes As Collection
            Set antes = FotoDasAbas(wb)
            On Error Resume Next
            Err.Clear
            Select Case tipo
                Case "cartao": detalhe = PrepararCartao(wb, ws, geradas)
                Case "pontonet": detalhe = PrepararPontoNet(wb, ws, geradas)
                Case Else
                    ws.Activate
                    detalhe = CStr(Application.Run(MacroDoTipo(tipo)))
                    If Err.Number = 0 Then geradas = AbasGeradas(wb, ws, antes)
            End Select
            If Err.Number <> 0 Then
                If Err.Number = 1004 And InStr(1, Err.Description, "macro", vbTextCompare) > 0 Then
                    detalhe = "ERRO: a macro " & MacroDoTipo(tipo) & " não está nesta planilha. Cole o módulo dela (o .txt) e rode de novo."
                Else
                    detalhe = "ERRO " & Err.Number & ": " & Err.Description & IIf(detalhe <> "", " (" & detalhe & ")", "")
                End If
            End If
            On Error GoTo 0
            If InStr(1, detalhe, "ERRO", vbBinaryCompare) > 0 Then resultado = "Deu erro" Else resultado = "OK"
            linhas.Add Array(nomeOriginal, NomeDoTipo(tipo), resultado, geradas, detalhe)
        End If
    Next i

    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    EscreverRelatorio wb, linhas, vistos
    Application.ScreenUpdating = True

    Dim nErros As Long, l As Variant
    For Each l In linhas
        If l(2) <> "OK" Then nErros = nErros + 1
    Next l
    MsgBox linhas.Count & " relatório(s) encontrado(s) e preparado(s)" & _
           IIf(nErros > 0, ", " & nErros & " com erro", ", todos OK") & "." & vbCrLf & vbCrLf & _
           "Veja os detalhes na aba """ & NOME_RELATORIO & """." & vbCrLf & _
           "Se estiver tudo certo, o próximo passo é a GerarAbaCSV.", _
           IIf(nErros > 0, vbExclamation, vbInformation), "Preparar as abas do mês"
End Sub

' --------------------------------------------------------------------------------
' Que relatório bruto é esta aba ("" = não é relatório bruto, ou já está formatado).
' Olha o texto das primeiras 40 linhas x 34 colunas.
' --------------------------------------------------------------------------------
Private Function TipoDoRelatorioBruto(ws As Worksheet) As String
    Dim v As Variant, r As Long, c As Long, t As String, tudo As String, colA As String
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Function
    v = ws.Range("A1:AH40").Value2
    For r = 1 To 40
        For c = 1 To 34
            If Not IsError(v(r, c)) Then
                t = UCase$(Trim$(Replace(CStr(v(r, c)), Chr(160), " ")))
                If t <> "" Then
                    tudo = tudo & "|" & t
                    If c = 1 Then colA = colA & "|" & t
                End If
            End If
        Next c
    Next r
    tudo = tudo & "|": colA = colA & "|"

    ' já formatadas (cabeçalho das macros na linha 1 ou 2): não mexe
    Dim l1 As String, l2 As String
    For c = 1 To 34
        If Not IsError(v(1, c)) Then l1 = l1 & "|" & UCase$(Trim$(CStr(v(1, c))))
        If Not IsError(v(2, c)) Then l2 = l2 & "|" & UCase$(Trim$(CStr(v(2, c))))
    Next c
    l1 = l1 & "|": l2 = l2 & "|"
    If InStr(l2, "|MATRÍCULA|") > 0 And InStr(l2, "|DT|") > 0 Then Exit Function          ' Cartão já consolidado
    If InStr(l1, "|PAUSA CORRETAS|") > 0 Then Exit Function                               ' pausas já formatadas
    If InStr(l1, "|MATRÍCULA|") > 0 And (InStr(l1, "|MOTIVO|") > 0 Or InStr(l1, "|SITUAÇÃO|") > 0 _
        Or InStr(l1, "|OCORRÊNCIA|") > 0 Or InStr(l1, "|INTERVALO TOTAL|") > 0 Or InStr(l1, "|TOTAL|") > 0) Then Exit Function

    If Left$(UCase$(Trim$(CStr(v(1, 1)))), 10) = "DATA FALTA" Then
        ' a aba "Pontonet" já separada em colunas é a formatada: não mexe
        If LCase$(ws.Name) = "pontonet" And InStr(CStr(v(1, 1)), ";") = 0 Then Exit Function
        TipoDoRelatorioBruto = "pontonet": Exit Function
    End If
    If InStr(tudo, "|CARTÃO PONTO|") > 0 And InStr(colA, "|EMPREGADO:|") > 0 Then TipoDoRelatorioBruto = "cartao": Exit Function
    If InStr(tudo, "PAUSAS TÉRMICAS") > 0 Then TipoDoRelatorioBruto = "pausas": Exit Function
    If InStr(tudo, "|CRACHÁ|") > 0 And InStr(tudo, "|COLETOR|") > 0 Then TipoDoRelatorioBruto = "digitadas": Exit Function
    If InStr(tudo, "HISTÓRICO DE AFASTAMENTOS") > 0 Or (InStr(tudo, "|PREV TÉRMINO|") > 0 And InStr(tudo, "|DT.TÉRMINO|") > 0) Then
        TipoDoRelatorioBruto = "afastamentos": Exit Function
    End If
    If InStr(tudo, "|REGISTROS DE PONTO|") > 0 Or InStr(tudo, "|HORAS DESCANSADAS|") > 0 Or InStr(tudo, "|HORAS DSR|") > 0 _
        Or InStr(tudo, "|MARCAÇÃO ANTES DSR|") > 0 Or InStr(tudo, "|CARGA HOR.|") > 0 Then
        TipoDoRelatorioBruto = "jornada": Exit Function
    End If
End Function

Private Function NomeDoTipo(ByVal tipo As String) As String
    Select Case tipo
        Case "cartao": NomeDoTipo = "Cartão Ponto"
        Case "pausas": NomeDoTipo = "Pausas Térmicas"
        Case "digitadas": NomeDoTipo = "Marcações digitadas"
        Case "afastamentos": NomeDoTipo = "Afastamentos"
        Case "jornada": NomeDoTipo = "Jornada (intra, inter, semanal ou horas excedentes)"
        Case "pontonet": NomeDoTipo = "PontoNet (ausência de marcação)"
    End Select
End Function

Private Function MacroDoTipo(ByVal tipo As String) As String
    Select Case tipo
        Case "cartao": MacroDoTipo = "ConsolidarCartaoPontoDefinitivo_Auto"
        Case "pausas": MacroDoTipo = "FormatarPausasTermicas_Auto"
        Case "digitadas": MacroDoTipo = "FormatarMarcacoesDigitadas_Auto"
        Case "afastamentos": MacroDoTipo = "FormatarAfastamentos_Auto"
        Case "jornada": MacroDoTipo = "FormatarJornada_Auto"
        Case "pontonet": MacroDoTipo = "(nenhuma: feito por esta macro)"
    End Select
End Function

' As abas que existem agora (para saber, depois de uma formatação, quais são novas).
Private Function FotoDasAbas(wb As Workbook) As Collection
    Dim sh As Worksheet
    Set FotoDasAbas = New Collection
    For Each sh In wb.Worksheets
        FotoDasAbas.Add sh
    Next sh
End Function

' Nomes das abas que saíram da formatação: a própria aba (que pode ter sido
' renomeada) e as abas novas que não existiam antes (ex.: pausas de vários meses).
Private Function AbasGeradas(wb As Workbook, ws As Worksheet, abasAntes As Collection) As String
    Dim sh As Worksheet, a As Variant, nova As Boolean, t As String
    t = ws.Name
    For Each sh In wb.Worksheets
        If sh.Name <> NOME_RELATORIO And sh.Name <> ws.Name Then
            nova = True
            For Each a In abasAntes
                If a Is sh Then nova = False: Exit For
            Next a
            If nova And InStr(1, "|" & t & "|", "|" & sh.Name & "|") = 0 Then t = t & ", " & sh.Name
        End If
    Next sh
    AbasGeradas = t
End Function

' --------------------------------------------------------------------------------
' Cartão Ponto: roda a ConsolidarCartaoPonto (que grava em "Cartao_Consolidado"),
' renomeia o resultado para "Cartão <Mês>" e apaga a aba bruta.
' --------------------------------------------------------------------------------
Private Function PrepararCartao(wb As Workbook, ws As Worksheet, ByRef geradas As String) As String
    Dim msg As String, wsC As Worksheet, ult As Long, r As Long, maxDt As Date, nome As String
    ' se sobrou um "Cartao_Consolidado" de antes, tira do caminho (a macro o apagaria)
    If AbaExisteAqui(wb, "Cartao_Consolidado") Then wb.Worksheets("Cartao_Consolidado").Name = "Cartao_Consolidado " & Format(Now, "hhnnss")
    ws.Activate
    msg = CStr(Application.Run("ConsolidarCartaoPontoDefinitivo_Auto"))
    If Not AbaExisteAqui(wb, "Cartao_Consolidado") Then
        PrepararCartao = "ERRO: a consolidação não gerou a aba. " & msg
        Exit Function
    End If
    Set wsC = wb.Worksheets("Cartao_Consolidado")
    If InStr(msg, "ERRO") > 0 Then
        Application.DisplayAlerts = False
        wsC.Delete
        Application.DisplayAlerts = True
        PrepararCartao = msg
        Exit Function
    End If
    ult = wsC.Cells(wsC.Rows.Count, 5).End(xlUp).Row
    For r = 3 To ult
        If IsDate(wsC.Cells(r, 5).Value) Then
            If wsC.Cells(r, 5).Value > maxDt Then maxDt = wsC.Cells(r, 5).Value
        End If
    Next r
    If maxDt = 0 Then
        PrepararCartao = "ERRO: não achei as datas na aba consolidada. " & msg
        Exit Function
    End If
    nome = "Cartão " & NomeDoMesAqui(Month(maxDt))
    If AbaExisteAqui(wb, nome) Then
        msg = msg & " ATENÇÃO: já existia uma aba """ & nome & """; a nova ficou como """ & nome & " (novo)"". " & _
              "Apague a que estiver errada (o GerarCSV usa as 3 abas de Cartão mais recentes)."
        nome = nome & " (novo)"
    End If
    wsC.Name = nome
    geradas = nome
    Application.DisplayAlerts = False
    ws.Delete     ' a aba bruta: o consolidado já tem tudo
    Application.DisplayAlerts = True
    PrepararCartao = msg & " A aba bruta foi apagada."
End Function

' --------------------------------------------------------------------------------
' PontoNet ("Ausência de marcação", .csv): se veio tudo na coluna A, separa pelo
' ";" (datas no formato dia/mês/ano; "Código terceiro" como texto) e renomeia a
' aba para "Pontonet" (o nome que o GerarCSV procura).
' --------------------------------------------------------------------------------
Private Function PrepararPontoNet(wb As Workbook, ws As Worksheet, ByRef geradas As String) As String
    Dim msg As String, ult As Long
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If InStr(CStr(ws.Cells(1, 1).Value), ";") > 0 Then
        ws.Range("A1:A" & ult).TextToColumns Destination:=ws.Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, _
            Comma:=False, Space:=False, Other:=False, _
            FieldInfo:=Array(Array(1, xlDMYFormat), Array(11, xlTextFormat), Array(15, xlDMYFormat), _
                             Array(16, xlDMYFormat), Array(17, xlDMYFormat), Array(18, xlDMYFormat))
        msg = "O arquivo veio numa coluna só: separado em colunas pelo "";"". "
    End If
    If ws.Cells(1, 1).Value <> "Data falta" Then
        PrepararPontoNet = "ERRO: depois de separar, a coluna A não começa com ""Data falta"". Confira o arquivo."
        Exit Function
    End If
    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit
    If LCase$(ws.Name) <> "pontonet" Then
        If AbaExisteAqui(wb, "Pontonet") Then
            msg = msg & "ATENÇÃO: já existe uma aba ""Pontonet"" e o GerarCSV usa ela; esta ficou como """ & ws.Name & """. Apague a antiga e renomeie esta."
        Else
            ws.Name = "Pontonet"
        End If
    End If
    geradas = ws.Name
    PrepararPontoNet = msg & (ult - 1) & " pedido(s) de justificativa."
End Function

' --------------------------------------------------------------------------------
' Aba "Relatório da preparação": uma linha por relatório encontrado e, embaixo, o
' que se espera num mês completo e se veio.
' --------------------------------------------------------------------------------
Private Sub EscreverRelatorio(wb As Workbook, linhas As Collection, vistos As Object)
    Dim ws As Worksheet, r As Long, l As Variant, esperados As Variant, i As Long
    If AbaExisteAqui(wb, NOME_RELATORIO) Then
        Set ws = wb.Worksheets(NOME_RELATORIO)
        ws.Cells.Clear
    Else
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = NOME_RELATORIO
    End If
    ws.Range("A1").Value = "Preparação das abas do mês - " & Format(Now, "dd/mm/yyyy hh:nn")
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 13
    ws.Range("A3:E3").Value = Array("Aba original", "O que era", "Resultado", "Aba(s) gerada(s)", "Detalhes")
    ws.Range("A3:E3").Font.Bold = True
    ws.Range("A3:E3").Interior.Color = RGB(217, 217, 217)
    r = 4
    If linhas.Count = 0 Then
        ws.Cells(r, 1).Value = "Nenhum relatório bruto encontrado nas abas desta planilha."
        r = r + 1
    End If
    For Each l In linhas
        ws.Cells(r, 1).Value = l(0)
        ws.Cells(r, 2).Value = l(1)
        ws.Cells(r, 3).Value = l(2)
        ws.Cells(r, 4).Value = l(3)
        ws.Cells(r, 5).Value = l(4)
        ws.Cells(r, 3).Font.Bold = True
        ws.Cells(r, 3).Font.Color = IIf(l(2) = "OK", RGB(0, 120, 60), RGB(200, 0, 0))
        r = r + 1
    Next l

    r = r + 1
    ws.Cells(r, 1).Value = "O que se espera num mês completo"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    esperados = Array("cartao", "pausas", "digitadas", "afastamentos", "jornada", "pontonet")
    For i = 0 To UBound(esperados)
        ws.Cells(r, 1).Value = NomeDoTipo(CStr(esperados(i)))
        If vistos.Exists(esperados(i)) Then
            ws.Cells(r, 3).Value = "Veio"
            ws.Cells(r, 3).Font.Color = RGB(0, 120, 60)
        ElseIf ExisteAbaFormatada(wb, CStr(esperados(i))) Then
            ws.Cells(r, 3).Value = "Já estava formatada"
            ws.Cells(r, 3).Font.Color = RGB(0, 120, 60)
        Else
            ws.Cells(r, 3).Value = "FALTOU: não encontrei nenhum relatório"
            ws.Cells(r, 3).Font.Color = RGB(200, 0, 0)
        End If
        r = r + 1
    Next i
    ws.Cells(r, 1).Value = "Também precisam estar na planilha: a RE, o ORG e as abas Tratamento - Internos/Externos (essas não são mexidas)."

    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 34
    ws.Columns("E").ColumnWidth = 90
    ws.Range("E:E").WrapText = True
    ws.Range("A:E").VerticalAlignment = xlTop
    ws.Activate
    ws.Range("A1").Select
End Sub

' Para a lista do que se espera: já existe uma aba formatada daquele tipo?
Private Function ExisteAbaFormatada(wb As Workbook, ByVal tipo As String) As Boolean
    Dim sh As Worksheet, n As String
    For Each sh In wb.Worksheets
        n = LCase$(sh.Name)
        Select Case tipo
            Case "cartao": If Left$(n, 6) = "cartão" Or Left$(n, 6) = "cartao" Then ExisteAbaFormatada = True
            Case "pausas": If Left$(n, 15) = "pausas térmicas" Then ExisteAbaFormatada = True
            Case "digitadas": If n = "marcações digitadas" Then ExisteAbaFormatada = True
            Case "afastamentos": If n = "afastamentos" Then ExisteAbaFormatada = True
            Case "jornada": If n = "intrajornada" Or n = "interjornada" Or n = "interjornada semanal" Or n = "horas excedentes" Then ExisteAbaFormatada = True
            Case "pontonet": If n = "pontonet" Then ExisteAbaFormatada = True
        End Select
    Next sh
End Function

Private Function AbaExisteAqui(wb As Workbook, ByVal nome As String) As Boolean
    Dim sh As Object
    For Each sh In wb.Sheets
        If LCase$(sh.Name) = LCase$(nome) Then AbaExisteAqui = True: Exit Function
    Next sh
End Function

Private Function NomeDoMesAqui(ByVal m As Long) As String
    NomeDoMesAqui = Choose(m, "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", _
                           "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro")
End Function
