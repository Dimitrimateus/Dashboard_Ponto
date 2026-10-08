Option Explicit

' ================================================================================
' MACRO: PURO_Para_EDITADO
' --------------------------------------------------------------------------------
' OBJETIVO
'   Ler os dados brutos da aba "PURO" (exportação do sistema de ponto) e gerar
'   uma nova aba, no formato "EDITADO_MESANO" (ex.: EDITADO_AGO26), já no layout
'   final usado para conferência/gestão, incluindo:
'     - Reordenação e renomeação de colunas
'     - Conversão de flags "X" para "SIM" nas colunas de marcação
'     - Inclusão das colunas manuais em branco (Gestor, Situação, Tratamento,
'       Pontonet), que não existem na PURO e são preenchidas à mão depois
'     - Descarte das colunas que não são usadas na EDITADO (Empr, T, Local,
'       Descrição, Column1)
'     - Remoção dos espaços sobrando no início/fim de TODOS os textos (nomes,
'       marcações, flags e cabeçalhos). A exportação do sistema preenche os
'       campos com espaços até uma largura fixa ("FULANO DE TAL      "), o que
'       atrapalha filtros, PROCV/XLOOKUP e comparações. Ver LimparTexto.
'
' POR QUE É "À PROVA DE ERROS"
'   1) Cada linha da PURO é processada de forma INDEPENDENTE (não existe
'      "amarração" linha a linha com uma EDITADO antiga). Isso significa que a
'      macro funciona igual mesmo que no mês seguinte tenham entrado ou saído
'      funcionários: ela simplesmente gera uma linha na EDITADO para cada linha
'      que existir na PURO naquele mês, seja 90, 111 ou 150 linhas.
'   2) As colunas são localizadas na PURO PELO NOME DO CABEÇALHO (e não pela
'      letra/posição). Se o sistema mudar a ordem das colunas na exportação,
'      a macro continua funcionando.
'   3) Se algum cabeçalho esperado não for encontrado na PURO, a macro avisa
'      o usuário (lista o que faltou) e segue em frente deixando a coluna
'      correspondente em branco na nova aba, em vez de travar com erro.
'   4) O nome da nova aba é gerado automaticamente a partir da data de apuração
'      mais recente encontrada na PURO (ex.: EDITADO_AGO26). Se já existir uma
'      aba com esse nome (rodou 2x no mesmo mês), a macro pergunta se deve
'      sobrescrever ou criar uma cópia numerada (EDITADO_AGO26 (2)).
'   5) Copia o formato de número (data/hora) da PURO para a nova aba, célula a
'      célula, então os valores continuam aparecendo como hora (hh:mm) ou data,
'      e não como número serial do Excel.
'   6) Usa tratamento de erro (On Error) em blocos críticos e restaura o estado
'      do Excel (cálculo automático, alertas etc.) mesmo se algo falhar no meio
'      do processo.
'
' COMO USAR
'   1) Abra o Editor VBA (Alt+F11) na planilha que contém a aba PURO.
'   2) Menu Arquivo > Importar Arquivo... e selecione este .bas
'      (ou copie/cole o conteúdo em um Módulo novo).
'   3) Volte à planilha, pressione Alt+F8, selecione "PURO_Para_EDITADO"
'      e clique em Executar.
'   ATENÇÃO: prefira COPIAR/COLAR. "Importar Arquivo" lê o .bas como ANSI e
'   estraga os acentos (o .bas está em UTF-8), e aí a busca de cabeçalhos com
'   acento ("Situação", "Dt.Apuraç.", "Marcações") deixa de achar a coluna.
' --------------------------------------------------------------------------------
' COMO SE LIGA AO RESTO DO PROJETO
'   aba PURO (exportação de ocorrências do sistema) -> esta macro -> aba
'   EDITADO_<MES><AA> -> o RH trata as ocorrências nessa aba (preenche Gestor,
'   Situação, Tratamento, Pontonet) e ela vira a aba "Tratamento" (ou
'   "Tratamento - Internos"/"Tratamento - Externos") -> o GerarAbaCSV
'   (GerarCSVPonto) lê essa aba pelos NOMES de cabeçalho e gera as ocorrências do
'   CSV -> painel (index.html). Por isso os nomes em targetHeaders (passo 3)
'   não devem mudar sem conferir o GerarCSVPonto.
'   Não depende de nenhum outro módulo (LimparTexto está no fim deste arquivo).
' ================================================================================

' Macro principal. Fluxo: 0) prepara o Excel; 1) acha a aba PURO; 2) mapeia os
' cabeçalhos; 3) define o layout de saída; 4) avisa cabeçalhos faltando;
' 5) descobre o mês para o nome da aba; 6) cria a aba; 7) cabeçalho;
' 8) copia linha a linha; 9) aparência. Os rótulos Finalizar/TratarErro, no fim,
' garantem que o Excel volte ao normal mesmo se der erro.
Sub PURO_Para_EDITADO()

    Dim wbAtual As Workbook
    Dim wsPuro As Worksheet
    Dim wsEditado As Worksheet
    Dim headerMap As Object          ' Dictionary: nome do cabeçalho (Trim) -> nº da coluna na PURO
    Dim colunasFaltando As String    ' Acumula avisos de cabeçalhos não encontrados

    Dim targetHeaders() As Variant   ' Cabeçalhos finais, na ordem exata da EDITADO
    Dim sourceKeys() As Variant      ' Nome do cabeçalho correspondente na PURO ("" = coluna manual/em branco)
    Dim convertFlag() As Variant     ' True = converter "X" -> "SIM" nesta coluna

    Dim i As Long, r As Long, lastColPuro As Long, lastRowPuro As Long
    Dim srcCol As Long
    Dim valorOriginal As String
    Dim statusBarAntes As Boolean

    Dim colDataApuracao As Long
    Dim dataReferencia As Date
    Dim temDataValida As Boolean
    Dim mesAno As String
    Dim nomeAbaFinal As String
    Dim sufixo As Long
    Dim monthAbbrev As Variant

    ' Qualquer erro não previsto pula para TratarErro (fim da rotina), que mostra a
    ' mensagem e restaura o Excel.
    On Error GoTo TratarErro

    ' ----------------------------------------------------------------------
    ' 0) Preparação: desliga atualizações de tela/cálculo para performance
    '    e para evitar "piscar" a tela durante a execução.
    ' ----------------------------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    statusBarAntes = Application.DisplayStatusBar
    Application.DisplayStatusBar = True

    Set wbAtual = ThisWorkbook

    ' ----------------------------------------------------------------------
    ' 1) Confirma que a aba PURO existe
    ' ----------------------------------------------------------------------
    On Error Resume Next
    Set wsPuro = wbAtual.Sheets("PURO")
    On Error GoTo TratarErro
    If wsPuro Is Nothing Then
        MsgBox "Não encontrei uma aba chamada 'PURO' nesta pasta de trabalho." & vbCrLf & _
               "Verifique o nome da aba e tente novamente.", vbCritical, "Aba não encontrada"
        GoTo Finalizar
    End If

    ' ----------------------------------------------------------------------
    ' 2) Mapeia os cabeçalhos da PURO (linha 1) pelo NOME, ignorando espaços
    '    extras no início/fim (a exportação do sistema costuma vir com
    '    espaços de preenchimento nos nomes das colunas).
    ' ----------------------------------------------------------------------
    Set headerMap = CreateObject("Scripting.Dictionary")
    lastColPuro = wsPuro.Cells(1, wsPuro.Columns.Count).End(xlToLeft).Column

    Dim chave As String
    For i = 1 To lastColPuro
        chave = Trim(CStr(wsPuro.Cells(1, i).Value))
        If chave <> "" Then
            If Not headerMap.Exists(chave) Then
                headerMap.Add chave, i
            End If
        End If
    Next i

    If Not headerMap.Exists("Nome") Then
        MsgBox "Não encontrei a coluna 'Nome' no cabeçalho da PURO. Não é possível continuar.", _
               vbCritical, "Cabeçalho obrigatório ausente"
        GoTo Finalizar
    End If

    lastRowPuro = wsPuro.Cells(wsPuro.Rows.Count, headerMap("Nome")).End(xlUp).Row
    If lastRowPuro < 2 Then
        MsgBox "A aba PURO não tem linhas de dados (apenas cabeçalho).", vbExclamation, "Nada a processar"
        GoTo Finalizar
    End If

    ' ----------------------------------------------------------------------
    ' 3) Define o layout final: cabeçalho exato, de qual cabeçalho da PURO
    '    ele vem, e se precisa converter X -> SIM.
    '    (Ordem = ordem das colunas A, B, C... na aba final)
    '
    '    Estes 3 arrays trabalham em conjunto, posição a posição (targetHeaders(i)
    '    corresponde a sourceKeys(i) corresponde a convertFlag(i)):
    '      targetHeaders(i) = nome da coluna na aba EDITADO (o que vai ser ESCRITO)
    '      sourceKeys(i)    = nome do cabeçalho a procurar na aba PURO (de ONDE vem o
    '                         valor); string vazia "" = não existe na PURO, é uma coluna
    '                         de preenchimento manual (Gestor/Situação/Tratamento/
    '                         Pontonet) que fica em branco para o RH preencher depois
    '      convertFlag(i)   = True nas colunas que guardam "marcação de intervalo"
    '                         (Intra/Inter/Marcações Impares/Horas Excedentes/Extras e
    '                         Faltas): a PURO grava "X" quando a condição ocorreu, e
    '                         aqui isso é convertido para "SIM" (mais legível para quem
    '                         só olha a planilha, sem precisar saber o código do sistema)
    ' ----------------------------------------------------------------------
    targetHeaders = Array( _
        "Matricula", _
        "Nome", _
        "Gestor", _
        "Situação", _
        "Tratamento", _
        "Pontonet", _
        "Dt.Apuraç.", _
        "Marcações", _
        "Extras", _
        "Faltas", _
        "Extra 100%", _
        "Qtd. Horas Pagas", _
        "Intra", _
        "Inter", _
        "Marcações Impares", _
        "Horas Excedentes", _
        "Extras e Faltas", _
        "Refeição fora de Escala")

    ' As 4 strings vazias abaixo correspondem, nesta ordem, às colunas
    ' Gestor, Situação, Tratamento e Pontonet - preenchimento manual,
    ' sem coluna de origem na PURO.
    sourceKeys = Array( _
        "Cadastro", _
        "Nome", _
        "", _
        "", _
        "", _
        "", _
        "Dt.Apuraç.", _
        "Marcações", _
        "Extras", _
        "Faltas", _
        "Extra 100%", _
        "Qtd. Horas Pagas", _
        "Intrajornada", _
        "Interjornada", _
        "Marcações Impares", _
        "Horas Excedentes", _
        "Extras e Faltas", _
        "Refeição fora de Escala")

    convertFlag = Array(False, False, False, False, False, False, False, False, _
                         False, False, False, False, True, True, True, True, True, False)

    ' ----------------------------------------------------------------------
    ' 4) Confere se todos os cabeçalhos esperados existem na PURO.
    '    Se algum não existir, avisa (mas não trava) - a coluna ficará
    '    em branco na aba final.
    ' ----------------------------------------------------------------------
    For i = 0 To UBound(sourceKeys)
        If sourceKeys(i) <> "" Then
            If Not headerMap.Exists(sourceKeys(i)) Then
                colunasFaltando = colunasFaltando & "- """ & sourceKeys(i) & """ (esperado para a coluna """ & targetHeaders(i) & """)" & vbCrLf
            End If
        End If
    Next i

    If colunasFaltando <> "" Then
        MsgBox "Atenção: os cabeçalhos abaixo NÃO foram encontrados na aba PURO." & vbCrLf & _
               "As colunas correspondentes ficarão em branco na aba gerada:" & vbCrLf & vbCrLf & _
               colunasFaltando, vbExclamation, "Cabeçalhos não encontrados"
    End If

    ' ----------------------------------------------------------------------
    ' 5) Descobre a competência (mês/ano) para nomear a aba automaticamente,
    '    usando a data de apuração MAIS RECENTE encontrada na PURO.
    '    Isso evita depender da data do sistema (relógio do PC) e reflete
    '    o período real dos dados importados.
    ' ----------------------------------------------------------------------
    monthAbbrev = Array("JAN", "FEV", "MAR", "ABR", "MAI", "JUN", _
                         "JUL", "AGO", "SET", "OUT", "NOV", "DEZ")

    temDataValida = False
    If headerMap.Exists("Dt.Apuraç.") Then
        colDataApuracao = headerMap("Dt.Apuraç.")
        For r = 2 To lastRowPuro
            If IsDate(wsPuro.Cells(r, colDataApuracao).Value) Then
                If Not temDataValida Then
                    dataReferencia = CDate(wsPuro.Cells(r, colDataApuracao).Value)
                    temDataValida = True
                ElseIf CDate(wsPuro.Cells(r, colDataApuracao).Value) > dataReferencia Then
                    dataReferencia = CDate(wsPuro.Cells(r, colDataApuracao).Value)
                End If
            End If
        Next r
    End If

    If temDataValida Then
        mesAno = monthAbbrev(Month(dataReferencia) - 1) & Format(dataReferencia, "yy")
    Else
        ' Nenhuma data válida encontrada na PURO -> usa a data/hora atual do
        ' sistema como último recurso, para garantir que a macro NUNCA trave
        ' por falta de nome de aba.
        mesAno = monthAbbrev(Month(Date) - 1) & Format(Date, "yy") & "_" & Format(Now, "hhmmss")
    End If

    nomeAbaFinal = "EDITADO_" & mesAno

    ' ----------------------------------------------------------------------
    ' 6) Cria a aba final. Se já existir uma aba com esse nome (segunda
    '    execução no mesmo mês), pergunta o que fazer.
    ' ----------------------------------------------------------------------
    Set wsEditado = Nothing
    On Error Resume Next
    Set wsEditado = wbAtual.Sheets(nomeAbaFinal)
    On Error GoTo TratarErro

    If Not wsEditado Is Nothing Then
        Dim resposta As VbMsgBoxResult
        resposta = MsgBox("Já existe uma aba chamada '" & nomeAbaFinal & "'." & vbCrLf & vbCrLf & _
                           "Sim = apagar o conteúdo dela e gerar novamente" & vbCrLf & _
                           "Não = criar uma cópia com outro nome (ex.: " & nomeAbaFinal & " (2))" & vbCrLf & _
                           "Cancelar = interromper sem alterar nada", _
                           vbYesNoCancel + vbQuestion, "Aba já existe")

        Select Case resposta
            Case vbYes
                wsEditado.Cells.Clear
            Case vbNo
                GoTo CriarComSufixo
            Case vbCancel
                MsgBox "Operação cancelada pelo usuário. Nada foi alterado.", vbInformation, "Cancelado"
                GoTo Finalizar
        End Select
    Else
        Set wsEditado = wbAtual.Sheets.Add(After:=wsPuro)
        wsEditado.Name = nomeAbaFinal
    End If

    ' Se chegou aqui, a aba já está pronta: pula o bloco CriarComSufixo, que só é
    ' usado quando o usuário escolheu "Não" (criar cópia numerada).
    GoTo PularCriacaoComSufixo

CriarComSufixo:
    ' Ponto auxiliar: cria a aba com sufixo numérico (2), (3)... quando o
    ' usuário optou por não sobrescrever a aba já existente.
    On Error Resume Next
    Set wsEditado = Nothing
    sufixo = 2
    Do
        Set wsEditado = wbAtual.Sheets(nomeAbaFinal & " (" & sufixo & ")")
        If wsEditado Is Nothing Then Exit Do
        sufixo = sufixo + 1
        Set wsEditado = Nothing
    Loop
    On Error GoTo TratarErro
    Set wsEditado = wbAtual.Sheets.Add(After:=wsPuro)
    wsEditado.Name = nomeAbaFinal & " (" & sufixo & ")"

PularCriacaoComSufixo:

    ' ----------------------------------------------------------------------
    ' 7) Escreve o cabeçalho
    ' ----------------------------------------------------------------------
    For i = 0 To UBound(targetHeaders)
        wsEditado.Cells(1, i + 1).Value = targetHeaders(i)
    Next i
    wsEditado.Rows(1).Font.Bold = True

    ' ----------------------------------------------------------------------
    ' 8) Copia e transforma linha a linha.
    '    Como cada linha da PURO é tratada de forma independente (sem
    '    depender de uma linha equivalente em execuções anteriores), a
    '    macro já funciona corretamente mesmo com admissões/desligamentos
    '    entre um mês e outro: simplesmente gera uma linha nova para cada
    '    linha existente na PURO, sejam quantas forem.
    ' ----------------------------------------------------------------------
    Dim linhaDestino As Long
    linhaDestino = 2

    For r = 2 To lastRowPuro

        Application.StatusBar = "Processando linha " & r & " de " & lastRowPuro & "..."

        ' Para cada linha de dado da PURO (r), percorre a lista de colunas de destino (i) e, para
        ' cada uma, escreve na aba EDITADO (linha "linhaDestino", coluna "i + 1") o valor
        ' correspondente da PURO (linha r, coluna "srcCol" — descoberta via headerMap, o dicionário
        ' "nome do cabeçalho -> nº da coluna" montado no passo 2). Ou seja: para cada colaborador
        ' (linha), refaz a mesma correspondência coluna-a-coluna definida no passo 3 acima.
        For i = 0 To UBound(sourceKeys)

            If sourceKeys(i) = "" Then
                ' Coluna manual (Gestor, Situação, Tratamento, Pontonet) -> fica em branco
                ' (não faz nada aqui de propósito)
            Else
                If headerMap.Exists(sourceKeys(i)) Then
                    srcCol = headerMap(sourceKeys(i))

                    If convertFlag(i) Then
                        ' -----------------------------------------------------
                        ' Colunas de flag: valor é "" (vazio/espaços) ou
                        ' começa com "X". Convertemos "X" -> "SIM" (sem os
                        ' espaços de preenchimento que vinham depois do "X");
                        ' qualquer outro valor sai sem espaços nas pontas, e
                        ' célula só com espaços sai vazia de verdade.
                        ' -----------------------------------------------------
                        valorOriginal = CStr(LimparTexto(wsPuro.Cells(r, srcCol).Value))
                        If Left(valorOriginal, 1) = "X" Then
                            wsEditado.Cells(linhaDestino, i + 1).Value = "SIM"
                        ElseIf valorOriginal <> "" Then
                            wsEditado.Cells(linhaDestino, i + 1).Value = valorOriginal
                        End If
                    Else
                        ' Cópia direta do valor (mantém tipo: data, hora e
                        ' número continuam como estão; só TEXTO passa por
                        ' LimparTexto, que tira os espaços do fim do nome,
                        ' das marcações etc.)
                        wsEditado.Cells(linhaDestino, i + 1).Value = LimparTexto(wsPuro.Cells(r, srcCol).Value)
                        ' Copia também o formato de número (datas e horas)
                        wsEditado.Cells(linhaDestino, i + 1).NumberFormat = _
                            wsPuro.Cells(r, srcCol).NumberFormat
                    End If
                End If
            End If

        Next i

        linhaDestino = linhaDestino + 1
    Next r

    ' ----------------------------------------------------------------------
    ' 9) Ajustes finais de aparência
    ' ----------------------------------------------------------------------
    wsEditado.Columns.AutoFit
    wsEditado.Rows(1).AutoFilter

    Application.StatusBar = False
    MsgBox "Concluído! " & (linhaDestino - 2) & " linha(s) transferida(s) da PURO para a aba '" & _
           wsEditado.Name & "'.", vbInformation, "Processo finalizado"

Finalizar:
    ' Restaura o Excel ao estado normal, mesmo se algo tiver falhado
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = statusBarAntes
    Application.StatusBar = False
    Exit Sub

' Rótulo de erro: só é alcançado via "On Error GoTo TratarErro". Mostra o erro e
' a linha da PURO em que estava, e volta para Finalizar.
TratarErro:
    MsgBox "Ocorreu um erro inesperado durante o processamento." & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Linha aproximada: " & r, vbCritical, "Erro na macro"
    Resume Finalizar

End Sub

' ================================================================================
' LimparTexto
' --------------------------------------------------------------------------------
' Tira os espaços das pontas de um valor de TEXTO (inclusive o espaço
' "não-quebrável", Chr(160), que exportações de sistema às vezes usam no
' lugar do espaço normal). Valores que não são texto (datas, horas, números)
' voltam intactos, para não perder o tipo nem o formato.
' Texto que só tinha espaços vira Empty, para a célula ficar vazia de verdade
' (senão CONT.VALORES/COUNTA e filtros de "(Vazias)" a contariam como
' preenchida).
' ================================================================================
Private Function LimparTexto(ByVal valor As Variant) As Variant
    Dim txt As String

    If VarType(valor) <> vbString Then
        LimparTexto = valor
        Exit Function
    End If

    txt = Trim$(Replace(valor, Chr(160), " "))
    If txt = "" Then
        LimparTexto = Empty
    Else
        LimparTexto = txt
    End If
End Function
