' ==============================================================================================================
' ATUALIZAR PAINEL DE PONTO - TUDO EM UM
'
' Este é o ÚNICO módulo que precisa ser colado na planilha de Tratamento. Ele contém, além desta
' parte nova, todas as macros do fluxo (ConsolidarCartaoPonto, FormatarPausasTermicas,
' FormatarMarcacoesDigitadas, FormatarAfastamentos e GerarCSVPonto), todas como Private: na lista de
' macros (Alt+F8) aparece só "AtualizarPainelPonto".
'
' O QUE A MACRO FAZ (nesta ordem)
'   1. Pede a pasta com os relatórios do sistema.
'   2. Confere se a pasta tem os 11 arquivos com os nomes combinados (lista em ArquivosEsperados).
'      O nome é comparado sem diferenciar maiúsculas, acentos e espaços repetidos; a extensão pode
'      ser .xlsx, .xls, .xlsm ou .xlsb. Se faltar algum (ou houver dois com o mesmo nome), avisa
'      quais e para, sem mexer em nada.
'   3. Pergunta em qual pasta salvar os CSVs.
'   4. Copia cada relatório para uma aba desta planilha (se a aba já existe, ela é LIMPA e
'      reaproveitada, para não quebrar fórmulas que apontem para ela) e roda a formatação:
'        Cartão ponto 1/2/3        -> ConsolidarCartaoPonto  -> abas "Cartão ponto 1", "2", "3"
'        Pausas térmicas 3 meses   -> FormatarPausasTermicas -> aba "Pausas Térmicas"
'        Digitadas 3 meses         -> FormatarMarcacoesDigitadas -> aba "Marcações Digitadas"
'        Afastamentos 3 meses      -> FormatarAfastamentos   -> aba "Afastamentos"
'        Pontonet 3 meses          -> (já vem como tabela)   -> aba "PontoNet"
'        Excedentes / Interjornada / Interjornada semanal / Intrajornada 3 meses
'                                  -> copiados como vieram    -> abas com o mesmo nome
'          (ainda NÃO entram no CSV: falta conhecer o formato desses relatórios)
'   5. Roda o GerarAbaCSV (usando SÓ as abas que acabou de importar para Cartão, Pausas, Digitadas e
'      Afastamentos; Tratamento, RE e ORG continuam vindo das abas da planilha) e o
'      ExportarCSVPorGestor na pasta escolhida.
'   6. Mostra um resumo único no fim (as mensagens de cada etapa não aparecem uma a uma).
'
' Durante a execução, as macros internas não mostram MsgBox: as mensagens vão para um registro
' (mLogLote) que aparece no resumo final ou, se algo der errado, na mensagem de erro.
' ==============================================================================================================

Private mLote As Boolean                    ' True enquanto AtualizarPainelPonto está rodando
Private mLogLote As String                  ' mensagens das macros internas durante a execução
Private mPastaSaida As String               ' pasta dos CSVs (o ExportarCSVPorGestor não pergunta de novo)
Private mAbasCartaoLote As Collection       ' abas de Cartão importadas agora (as antigas da planilha ficam de fora)
Private mAbaAfastLote As Worksheet
Private mAbaDigitLote As Worksheet
Private mAbaPausasLote As Worksheet

Private Const TITULO_LOTE As String = "Atualizar Painel de Ponto"
Private Const ABA_CONSOLIDADO As String = "Cartao_Consolidado"
Private Const ABA_CONSOLIDADO_TEMP As String = "Cartao_Consolidado (guardada)"

' Os 11 arquivos: nome do arquivo (sem extensão) | aba de destino | o que fazer com ele.
Private Function ArquivosEsperados() As Variant
    ArquivosEsperados = Array( _
        Array("Cartão ponto 1", "Cartão ponto 1", "CARTAO"), _
        Array("Cartão ponto 2", "Cartão ponto 2", "CARTAO"), _
        Array("Cartão ponto 3", "Cartão ponto 3", "CARTAO"), _
        Array("Pausas térmicas 3 meses", "Pausas Térmicas", "PAUSAS"), _
        Array("Digitadas 3 meses", "Marcações Digitadas", "DIGITADAS"), _
        Array("Afastamentos 3 meses", "Afastamentos", "AFASTAMENTOS"), _
        Array("Pontonet 3 meses", "PontoNet", "PONTONET"), _
        Array("Excedentes 3 meses", "Excedentes", "BRUTO"), _
        Array("Interjornada 3 meses", "Interjornada", "BRUTO"), _
        Array("Interjornada semanal 3 meses", "Interjornada Semanal", "BRUTO"), _
        Array("Intrajornada 3 meses", "Intrajornada", "BRUTO"))
End Function

Public Sub AtualizarPainelPonto()
    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' ---------------------------------------------------------------- 1. pasta dos relatórios
    Dim pastaEntrada As String
    pastaEntrada = EscolherPastaComTitulo("Selecione a pasta com os relatórios (Cartão ponto 1, 2, 3, Pausas, Digitadas...)")
    If pastaEntrada = "" Then Exit Sub

    ' ---------------------------------------------------------------- 2. confere os arquivos
    Dim lista As Variant, caminhos() As String, i As Long
    lista = ArquivosEsperados()
    ReDim caminhos(LBound(lista) To UBound(lista))

    Dim problemas As String
    problemas = LocalizarArquivos(pastaEntrada, lista, caminhos)
    If problemas <> "" Then
        MsgBox "A pasta escolhida não tem todos os arquivos necessários. Nada foi alterado." & vbCrLf & vbCrLf & _
            "Pasta: " & pastaEntrada & vbCrLf & vbCrLf & problemas, vbExclamation, TITULO_LOTE
        Exit Sub
    End If

    ' ---------------------------------------------------------------- 3. pasta dos CSVs
    Dim pastaSaida As String
    pastaSaida = EscolherPastaComTitulo("Todos os arquivos foram encontrados. Onde salvar os arquivos CSV?")
    If pastaSaida = "" Then
        MsgBox "Nenhuma pasta escolhida para os CSVs. Nada foi alterado.", vbInformation, TITULO_LOTE
        Exit Sub
    End If

    ' Avisa antes de sobrescrever abas que já existem na planilha.
    Dim abasExistentes As String
    For i = LBound(lista) To UBound(lista)
        If ExisteAba(wb, CStr(lista(i)(1))) Then abasExistentes = abasExistentes & vbCrLf & "  - " & lista(i)(1)
    Next i
    If abasExistentes <> "" Then
        If MsgBox("Os 11 arquivos foram encontrados." & vbCrLf & vbCrLf & _
                  "Estas abas já existem nesta planilha e o conteúdo delas será SUBSTITUÍDO pelos relatórios novos:" & _
                  abasExistentes & vbCrLf & vbCrLf & "Continuar?", vbYesNo + vbQuestion, TITULO_LOTE) <> vbYes Then Exit Sub
    End If

    ' ---------------------------------------------------------------- 4. importa e formata
    Dim calcAntes As XlCalculation
    calcAntes = Application.Calculation
    mLote = True
    mLogLote = ""
    mPastaSaida = pastaSaida
    Set mAbasCartaoLote = New Collection
    Set mAbaAfastLote = Nothing
    Set mAbaDigitLote = Nothing
    Set mAbaPausasLote = Nothing

    Dim etapa As String, avisos As String
    Dim consolidadoGuardado As Boolean
    Dim abasTemporarias As New Collection

    On Error GoTo Falha
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    ' O ConsolidarCartaoPonto escreve sempre na aba "Cartao_Consolidado". Se ela já existe na planilha,
    ' é guardada com outro nome durante a execução e volta ao nome original no fim.
    If ExisteAba(wb, ABA_CONSOLIDADO) Then
        If ExisteAba(wb, ABA_CONSOLIDADO_TEMP) Then wb.Worksheets(ABA_CONSOLIDADO_TEMP).Delete
        wb.Worksheets(ABA_CONSOLIDADO).Name = ABA_CONSOLIDADO_TEMP
        consolidadoGuardado = True
    End If

    Dim nomeArq As String, nomeAba As String, tipo As String
    Dim wsAba As Worksheet, wsBruto As Worksheet, wsCons As Worksheet
    For i = LBound(lista) To UBound(lista)
        nomeArq = CStr(lista(i)(0))
        nomeAba = CStr(lista(i)(1))
        tipo = CStr(lista(i)(2))
        etapa = "importar """ & nomeArq & """"
        Application.StatusBar = "Atualizar painel: importando " & nomeArq & "..."

        Select Case tipo
        Case "CARTAO"
            ' Bruto numa aba temporária -> Consolidar -> resultado copiado para "Cartão ponto N".
            Set wsBruto = AbaLimpa(wb, nomeAba & " (bruto)")
            abasTemporarias.Add wsBruto.Name
            CopiarRelatorio caminhos(i), wsBruto
            etapa = "consolidar """ & nomeArq & """ (ConsolidarCartaoPonto)"
            wsBruto.Activate
            ConsolidarCartaoPontoDefinitivo
            Set wsCons = Nothing
            If ExisteAba(wb, ABA_CONSOLIDADO) Then Set wsCons = wb.Worksheets(ABA_CONSOLIDADO)
            If wsCons Is Nothing Then Err.Raise vbObjectError + 1, , "A aba Cartao_Consolidado não foi criada."
            If ColunaPorCabecalho(wsCons, "Matrícula", 2) = 0 Or UltimaLinhaPreenchida(wsCons, 1) < 3 Then
                Err.Raise vbObjectError + 2, , "O arquivo não parece um Cartão Ponto (HRCP102): nenhum dia foi consolidado."
            End If
            Application.ScreenUpdating = False
            Set wsAba = AbaLimpa(wb, nomeAba)
            wsCons.UsedRange.Copy Destination:=wsAba.Range(wsCons.UsedRange.Address)
            wsAba.Columns.AutoFit
            wsCons.Delete
            mAbasCartaoLote.Add wsAba

        Case "PAUSAS"
            Set wsAba = AbaLimpa(wb, nomeAba)
            CopiarRelatorio caminhos(i), wsAba
            etapa = "formatar """ & nomeArq & """ (FormatarPausasTermicas)"
            wsAba.Activate
            FormatarPausasTermicas
            If AcharCabecalhoPausas(wsAba) = 0 Then
                Err.Raise vbObjectError + 3, , "A tabela de pausas (Matrícula / Pausa corretas) não foi gerada."
            End If
            Set mAbaPausasLote = wsAba

        Case "DIGITADAS"
            Set wsAba = AbaLimpa(wb, nomeAba)
            CopiarRelatorio caminhos(i), wsAba
            etapa = "formatar """ & nomeArq & """ (FormatarMarcacoesDigitadas)"
            wsAba.Activate
            FormatarMarcacoesDigitadas
            If ColunaPorCabecalho(wsAba, "Matrícula") = 0 Or ColunaPorCabecalho(wsAba, "Motivo") = 0 Then
                Err.Raise vbObjectError + 4, , "A tabela de marcações digitadas não foi gerada."
            End If
            Set mAbaDigitLote = wsAba

        Case "AFASTAMENTOS"
            Set wsAba = AbaLimpa(wb, nomeAba)
            CopiarRelatorio caminhos(i), wsAba
            etapa = "formatar """ & nomeArq & """ (FormatarAfastamentos)"
            wsAba.Activate
            FormatarAfastamentos
            If ColunaPorCabecalho(wsAba, "Matrícula") = 0 Or ColunaPorCabecalho(wsAba, "Início") = 0 Then
                Err.Raise vbObjectError + 5, , "A tabela de afastamentos não foi gerada."
            End If
            Set mAbaAfastLote = wsAba

        Case "PONTONET"
            Set wsAba = AbaLimpa(wb, nomeAba)
            CopiarRelatorio caminhos(i), wsAba
            If ColunaPorCabecalho(wsAba, "Matrícula") = 0 Or ColunaPorCabecalho(wsAba, "Data falta") = 0 Then
                avisos = avisos & vbCrLf & "- " & nomeArq & ": não achei as colunas 'Matrícula' e 'Data falta' na linha 1; " & _
                    "a demora no PontoNet vai sair vazia no painel."
            End If

        Case Else ' BRUTO
            Set wsAba = AbaLimpa(wb, nomeAba)
            CopiarRelatorio caminhos(i), wsAba
        End Select
        Application.ScreenUpdating = False
    Next i

    ' Apaga as abas temporárias dos Cartões brutos.
    Dim nomeTemp As Variant
    For Each nomeTemp In abasTemporarias
        If ExisteAba(wb, CStr(nomeTemp)) Then wb.Worksheets(CStr(nomeTemp)).Delete
    Next nomeTemp

    ' ---------------------------------------------------------------- 5. CSV
    etapa = "gerar a aba CSV (GerarAbaCSV)"
    Application.StatusBar = "Atualizar painel: gerando a aba CSV..."
    ' Limpa a aba CSV antiga: se o GerarAbaCSV parar no meio (ex.: sem aba RE), não exporta dados velhos.
    If ExisteAba(wb, "CSV") Then wb.Worksheets("CSV").Cells.Clear
    mLogLote = ""
    GerarAbaCSV
    Dim resumoCSV As String
    resumoCSV = mLogLote
    Dim wsCSV As Worksheet
    If ExisteAba(wb, "CSV") Then Set wsCSV = wb.Worksheets("CSV")
    If wsCSV Is Nothing Then Err.Raise vbObjectError + 6, , "A aba CSV não foi gerada."
    If UltimaLinhaPreenchida(wsCSV, 1) < 2 Then Err.Raise vbObjectError + 7, , "A aba CSV ficou vazia."

    etapa = "salvar os arquivos CSV (ExportarCSVPorGestor)"
    Application.StatusBar = "Atualizar painel: salvando os CSVs..."
    mLogLote = ""
    ExportarCSVPorGestor
    Dim resumoExport As String
    resumoExport = mLogLote

    LimparEstadoLote wb, consolidadoGuardado, calcAntes
    wb.Worksheets("CSV").Activate

    MsgBox "Pronto! Relatórios importados e CSVs salvos." & vbCrLf & vbCrLf & _
        resumoExport & vbCrLf & vbCrLf & _
        "---- Resumo do CSV ----" & vbCrLf & resumoCSV & vbCrLf & vbCrLf & _
        "Excedentes, Interjornada, Interjornada semanal e Intrajornada foram copiados para abas com " & _
        "esses nomes, mas ainda não entram no CSV/painel." & _
        IIf(avisos <> "", vbCrLf & vbCrLf & "ATENÇÃO:" & avisos, ""), vbInformation, TITULO_LOTE
    Exit Sub

Falha:
    Dim msgErro As String
    msgErro = "Erro ao " & etapa & ":" & vbCrLf & Err.Description & " (" & Err.Number & ")"
    On Error Resume Next
    For Each nomeTemp In abasTemporarias
        If ExisteAba(wb, CStr(nomeTemp)) Then wb.Worksheets(CStr(nomeTemp)).Delete
    Next nomeTemp
    LimparEstadoLote wb, consolidadoGuardado, calcAntes
    On Error GoTo 0
    MsgBox msgErro & vbCrLf & vbCrLf & "Os CSVs NÃO foram gerados." & _
        IIf(mLogLote <> "", vbCrLf & vbCrLf & "Mensagens das macros:" & vbCrLf & mLogLote, ""), _
        vbCritical, TITULO_LOTE
End Sub

' Volta o Excel ao normal e desfaz o modo "lote".
Private Sub LimparEstadoLote(wb As Workbook, ByVal consolidadoGuardado As Boolean, ByVal calcAntes As XlCalculation)
    On Error Resume Next
    If consolidadoGuardado Then
        If ExisteAba(wb, ABA_CONSOLIDADO) Then wb.Worksheets(ABA_CONSOLIDADO).Delete
        wb.Worksheets(ABA_CONSOLIDADO_TEMP).Name = ABA_CONSOLIDADO
    End If
    mLote = False
    mPastaSaida = ""
    Set mAbasCartaoLote = Nothing
    Set mAbaAfastLote = Nothing
    Set mAbaDigitLote = Nothing
    Set mAbaPausasLote = Nothing
    Application.CutCopyMode = False
    Application.Calculation = calcAntes
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' Procura cada arquivo esperado na pasta. Preenche "caminhos" e devolve "" se achou todos, ou o texto
' do aviso (arquivos que faltam, nomes repetidos e os arquivos Excel da pasta que não bateram).
Private Function LocalizarArquivos(ByVal pasta As String, lista As Variant, ByRef caminhos() As String) As String
    Dim fso As Object, arq As Object, ext As String, chave As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim porNome As Object, repetidos As Object
    Set porNome = CreateObject("Scripting.Dictionary")
    Set repetidos = CreateObject("Scripting.Dictionary")
    For Each arq In fso.GetFolder(pasta).Files
        ext = LCase$(fso.GetExtensionName(arq.Name))
        If Left$(arq.Name, 2) <> "~$" And (ext = "xlsx" Or ext = "xls" Or ext = "xlsm" Or ext = "xlsb") Then
            chave = NomeComparavel(fso.GetBaseName(arq.Name))
            If porNome.Exists(chave) Then
                repetidos(chave) = repetidos(chave) & ", " & arq.Name
            Else
                porNome.Add chave, arq.Path
                repetidos.Add chave, arq.Name
            End If
        End If
    Next arq

    Dim i As Long, faltando As String, duplicados As String, usados As Object
    Set usados = CreateObject("Scripting.Dictionary")
    For i = LBound(lista) To UBound(lista)
        chave = NomeComparavel(CStr(lista(i)(0)))
        If porNome.Exists(chave) Then
            caminhos(i) = porNome(chave)
            usados(chave) = True
            If InStr(repetidos(chave), ", ") > 0 Then
                duplicados = duplicados & vbCrLf & "  - " & lista(i)(0) & ": " & repetidos(chave)
            End If
        Else
            faltando = faltando & vbCrLf & "  - " & lista(i)(0)
        End If
    Next i

    Dim texto As String, sobrando As String, k As Variant
    If faltando <> "" Then texto = "FALTANDO:" & faltando
    If duplicados <> "" Then
        If texto <> "" Then texto = texto & vbCrLf & vbCrLf
        texto = texto & "NOME REPETIDO (deixe só um arquivo de cada):" & duplicados
    End If
    If texto <> "" Then
        For Each k In porNome.Keys
            If Not usados.Exists(k) Then sobrando = sobrando & vbCrLf & "  - " & fso.GetFileName(porNome(k))
        Next k
        If sobrando <> "" Then
            texto = texto & vbCrLf & vbCrLf & "Arquivos Excel da pasta que não bateram com nenhum nome " & _
                "(confira se não é um deles com o nome diferente):" & sobrando
        End If
    End If
    LocalizarArquivos = texto
End Function

' Nome para comparação: minúsculas, sem acento e com espaços simples ("Cartão  Ponto 1" = "cartao ponto 1").
Private Function NomeComparavel(ByVal t As String) As String
    t = LCase$(SemAcento(Trim$(t)))
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NomeComparavel = t
End Function

' Abre o relatório só para leitura e copia a primeira aba para "destino" (mesmas posições de células,
' com formatos, para as macros de formatação lerem como se fosse o arquivo original).
Private Sub CopiarRelatorio(ByVal caminho As String, destino As Worksheet)
    Dim wbOrigem As Workbook, jaAberto As Boolean, w As Workbook
    For Each w In Application.Workbooks
        If LCase$(w.FullName) = LCase$(caminho) Then
            Set wbOrigem = w
            jaAberto = True
            Exit For
        End If
    Next w
    If wbOrigem Is Nothing Then
        Set wbOrigem = Workbooks.Open(Filename:=caminho, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    End If

    Dim wsOrigem As Worksheet
    Set wsOrigem = wbOrigem.Worksheets(1)
    If Application.WorksheetFunction.CountA(wsOrigem.UsedRange) > 0 Then
        wsOrigem.UsedRange.Copy Destination:=destino.Range(wsOrigem.UsedRange.Address)
    End If
    Application.CutCopyMode = False
    If Not jaAberto Then wbOrigem.Close SaveChanges:=False
    ThisWorkbook.Activate
End Sub

' Devolve a aba com esse nome, LIMPA (conteúdo e formatos). Se não existe, cria no fim da pasta.
' A aba existente é reaproveitada (e não apagada) para não quebrar fórmulas que apontem para ela.
Private Function AbaLimpa(wb As Workbook, ByVal nome As String) As Worksheet
    Dim ws As Worksheet
    If ExisteAba(wb, nome) Then
        Set ws = wb.Worksheets(nome)
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        ws.Cells.Clear
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = nome
    End If
    ws.Visible = xlSheetVisible
    Set AbaLimpa = ws
End Function

Private Function ExisteAba(wb As Workbook, ByVal nome As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Sheets(nome)
    On Error GoTo 0
    ExisteAba = Not ws Is Nothing
End Function

Private Function EscolherPastaComTitulo(ByVal titulo As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(4) ' 4 = msoFileDialogFolderPicker
    fd.Title = titulo
    fd.AllowMultiSelect = False
    If ThisWorkbook.Path <> "" Then fd.InitialFileName = ThisWorkbook.Path & "\"
    If fd.Show = -1 Then
        EscolherPastaComTitulo = fd.SelectedItems(1)
    Else
        EscolherPastaComTitulo = ""
    End If
End Function

' Linha do cabeçalho da tabela de pausas (Matrícula + Pausa corretas) nas 10 primeiras linhas; 0 = não achou.
Private Function AcharCabecalhoPausas(ws As Worksheet) As Long
    Dim r As Long
    For r = 1 To 10
        If ColunaPorCabecalho(ws, "Matrícula", r) > 0 And ColunaPorCabecalho(ws, "Pausa corretas", r) > 0 Then
            LinhaCabecalhoPausas = r
            Exit Function
        End If
    Next r
    LinhaCabecalhoPausas = 0
End Function

' ---------------------------------------------------------------------------------------------
' Pontos de troca usados pelo GerarAbaCSV: durante o AtualizarPainelPonto, ele usa as abas que
' acabaram de ser importadas (e não uma aba antiga com o mesmo cabeçalho que esteja na planilha).
' Fora do lote, procuram como antes.
' ---------------------------------------------------------------------------------------------
Private Function LoteAbasCartao(wb As Workbook) As Collection
    If mLote And Not mAbasCartaoLote Is Nothing Then
        If mAbasCartaoLote.Count > 0 Then
            Set LoteAbasCartao = mAbasCartaoLote
            Exit Function
        End If
    End If
    Set LoteAbasCartao = EncontrarAbasCartao(wb)
End Function

Private Function LoteAbaAfastamentos(wb As Workbook) As Worksheet
    If mLote And Not mAbaAfastLote Is Nothing Then
        Set LoteAbaAfastamentos = mAbaAfastLote
    Else
        Set LoteAbaAfastamentos = EncontrarAbaAfastamentos(wb)
    End If
End Function

Private Function LoteAbaDigitadas(wb As Workbook) As Worksheet
    If mLote And Not mAbaDigitLote Is Nothing Then
        Set LoteAbaDigitadas = mAbaDigitLote
    Else
        Set LoteAbaDigitadas = EncontrarAbaDigitadas(wb)
    End If
End Function

Private Function LoteTabelaPausas(wb As Workbook, ByRef linhaCabecalho As Long) As Worksheet
    If mLote And Not mAbaPausasLote Is Nothing Then
        linhaCabecalho = AcharCabecalhoPausas(mAbaPausasLote)
        If linhaCabecalho > 0 Then
            Set LoteTabelaPausas = mAbaPausasLote
            Exit Function
        End If
    End If
    Set LoteTabelaPausas = EncontrarTabelaPausas(wb, linhaCabecalho)
End Function

' MsgBox das macros internas: durante o lote vai para o registro; fora dele, aparece normalmente.
Private Sub Avisar(ByVal texto As String, Optional ByVal botoes As VbMsgBoxStyle = vbOKOnly, Optional ByVal titulo As String = "")
    If mLote Then
        If mLogLote <> "" Then mLogLote = mLogLote & vbCrLf
        If titulo <> "" Then
            mLogLote = mLogLote & "[" & titulo & "] " & texto
        Else
            mLogLote = mLogLote & texto
        End If
    Else
        If titulo = "" Then titulo = TITULO_LOTE
        MsgBox texto, botoes, titulo
    End If
End Sub

' Pergunta das macros internas (ex.: "vai reescrever a aba, continuar?"): no lote, a resposta é Sim.
Private Function Perguntar(ByVal texto As String, Optional ByVal botoes As VbMsgBoxStyle = vbOKOnly, _
                           Optional ByVal titulo As String = "") As VbMsgBoxResult
    If mLote Then
        Perguntar = vbYes
    Else
        If titulo = "" Then titulo = TITULO_LOTE
        Perguntar = MsgBox(texto, botoes, titulo)
    End If
End Function
