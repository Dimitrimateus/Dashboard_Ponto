Attribute VB_Name = "modFormatarJornada"
Option Explicit

' ============================================================================
' FormatarJornada
'
' UMA macro para QUATRO relatórios do sistema (Senior), todos sobre limites de
' jornada. A macro olha os títulos da aba ativa, descobre qual é o relatório e
' o transforma numa tabela simples (cabeçalho na linha 1, uma linha por
' ocorrência), renomeando a aba:
'
'   Relatório do sistema          Regra (o que ele aponta)            Aba gerada
'   ----------------------------  ----------------------------------  ---------------------
'   Intrajornada (HRES108)        intervalo de almoço/descanso dentro  "Intrajornada"
'                                 do dia menor que o mínimo
'   Interjornada                  menos de 11h de descanso entre o     "Interjornada"
'                                 fim de um dia e o início do outro
'   Interjornada semanal          trabalhou no dia do DSR/feriado      "Interjornada Semanal"
'                                 (descanso semanal não cumprido)
'   Horas excedentes              mais de 10h de trabalho no dia       "Horas Excedentes"
'
' DE ONDE VEM: o .xlsx baixado do sistema, aberto no Excel, com a aba do
'   relatório ativa. Os relatórios vêm como "dump" de impressão: títulos numa
'   linha, valores DESLOCADOS em relação aos títulos, cabeçalho de página
'   repetido e (na semanal) a pessoa só na primeira linha dela.
'
' PARA ONDE VAI: a aba reescrita é copiada para a planilha de Tratamento. O
'   GerarAbaCSV (módulo GerarCSVPonto) acha cada uma pelo cabeçalho (não pelo
'   nome) e gera uma linha no CSV por ocorrência, com tipo_ocorrencia
'   "Intrajornada", "Interjornada", "Interjornada Semanal" ou "Horas
'   Excedentes". Por isso os NOMES das colunas abaixo (nos Array(...) de cada
'   rotina Montar*) não devem mudar sem mudar o GerarCSVPonto junto.
'
' Não depende de nenhum outro módulo: as funções de apoio estão no fim.
' A versão em Python desta macro (para conferir sem Excel) é
' ferramentas/formatar_jornada.py: as duas seguem as mesmas regras.
'
' Rode numa CÓPIA do arquivo exportado: a aba ativa é reescrita por completo.
' ============================================================================

Sub FormatarJornada()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim usado As Range
    Set usado = ws.UsedRange
    If usado.Rows.Count < 2 Then
        MsgBox "Nenhum dado encontrado na aba ativa.", vbExclamation, "Formatar Jornada"
        Exit Sub
    End If

    ' Lê tudo de uma vez. Value2: datas e horas vêm como número (com Value,
    ' data vira Date e IsNumeric dá False).
    ' deslCol: o array começa na 1ª coluna USADA; as posições de reserva
    ' (relatórios de outubro/2026) são da planilha (A = 1) e são convertidas
    ' com "- deslCol" dentro de cada rotina Montar*.
    Dim dados As Variant, deslCol As Long
    dados = usado.Value2
    deslCol = usado.Column - 1

    ' 1) Qual relatório é? (pelos títulos das 12 primeiras linhas)
    Dim tipo As String
    tipo = TipoDoRelatorio(dados)
    If tipo = "" Then
        MsgBox "Não reconheci o relatório da aba ativa." & vbCrLf & _
               "Esta macro formata: Intrajornada, Interjornada, Interjornada semanal e Horas excedentes.", _
               vbExclamation, "Formatar Jornada"
        Exit Sub
    End If

    ' 2) Monta a tabela na memória. Cada rotina devolve o cabeçalho, as linhas
    '    (saida, nSaida), o nome da aba e o formato de cada coluna.
    Dim cab As Variant, saida() As Variant, nSaida As Long, nomeAba As String, formatos As Variant
    Select Case tipo
        Case "intrajornada": MontarIntrajornada dados, deslCol, cab, saida, nSaida, nomeAba, formatos
        Case "interjornada": MontarInterjornada dados, deslCol, cab, saida, nSaida, nomeAba, formatos
        Case "intersemanal": MontarInterSemanal dados, deslCol, cab, saida, nSaida, nomeAba, formatos
        Case "excedentes": MontarExcedentes dados, deslCol, cab, saida, nSaida, nomeAba, formatos
    End Select

    If nSaida = 0 Then
        MsgBox "Reconheci o relatório (" & nomeAba & "), mas não encontrei nenhuma ocorrência nele.", _
               vbExclamation, "Formatar Jornada"
        Exit Sub
    End If

    ' 3) Apaga a aba e escreve a tabela formatada.
    EscreverTabela ws, cab, saida, nSaida, formatos

    ' 4) Renomeia (se já existir uma aba com esse nome, mantém o nome atual).
    Dim renomeada As Boolean
    On Error Resume Next
    ws.Name = nomeAba
    renomeada = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    MsgBox "Relatório reconhecido: " & nomeAba & vbCrLf & _
           nSaida & " ocorrência(s) na tabela." & vbCrLf & _
           IIf(renomeada, "Aba renomeada para """ & nomeAba & """.", _
               "A aba não foi renomeada (já existe uma aba """ & nomeAba & """)."), _
           vbInformation, "Formatar Jornada"
End Sub

' ----------------------------------------------------------------------------
' Descobre o relatório pelos títulos (maiúsculas e sem acento):
'   "REGISTROS DE PONTO" + "INTERVALO"   -> intrajornada
'   "HORAS DESCANSADAS"                  -> interjornada
'   "HORAS DSR" ou "MARCACAO ANTES DSR"  -> intersemanal
'   "CARGA HOR" + "EXTRAS"               -> excedentes
' Devolve "" se não reconhecer.
' ----------------------------------------------------------------------------
Private Function TipoDoRelatorio(dados As Variant) As String
    Dim r As Long, c As Long, tudo As String, ultima As Long
    ultima = UBound(dados, 1)
    If ultima > 12 Then ultima = 12
    For r = 1 To ultima
        For c = 1 To UBound(dados, 2)
            If VarType(dados(r, c)) = vbString Then tudo = tudo & " | " & TextoSimples(dados(r, c))
        Next c
    Next r
    If InStr(tudo, "REGISTROS DE PONTO") > 0 And InStr(tudo, "INTERVALO") > 0 Then
        TipoDoRelatorio = "intrajornada"
    ElseIf InStr(tudo, "HORAS DESCANSADAS") > 0 Then
        TipoDoRelatorio = "interjornada"
    ElseIf InStr(tudo, "HORAS DSR") > 0 Or InStr(tudo, "MARCACAO ANTES DSR") > 0 Then
        TipoDoRelatorio = "intersemanal"
    ElseIf InStr(tudo, "CARGA HOR") > 0 And InStr(tudo, "EXTRAS") > 0 Then
        TipoDoRelatorio = "excedentes"
    Else
        TipoDoRelatorio = ""
    End If
End Function

' ----------------------------------------------------------------------------
' INTRAJORNADA (HRES108 "Relação de Ocorrências de Intrajornada")
' Títulos: Cadastro | Nome | Local | Admissão | Cargo | Data | Dia |
'          Registros de Ponto | Car. Horaria | Intervalo(s)
' Valores: cadastro = matrícula; "Dia" vem como "Sex (985)" = dia da semana +
'   código do horário; as batidas começam NA coluna de "Registros de Ponto" e
'   ocupam 8 colunas; logo depois vêm a carga horária trabalhada e até 3
'   intervalos (descansos feitos no dia).
' Uma linha de dado = matrícula numérica + data.
' ----------------------------------------------------------------------------
Private Sub MontarIntrajornada(dados As Variant, ByVal deslCol As Long, ByRef cab As Variant, _
    ByRef saida() As Variant, ByRef nSaida As Long, ByRef nomeAba As String, ByRef formatos As Variant)

    nomeAba = "Intrajornada"
    cab = Array("Matrícula", "Colaborador", "Local", "Admissão", "Cargo", "Data", "Dia", "Cód. Horário", _
                "Marcações", "Carga horária", "Intervalo 1", "Intervalo 2", "Intervalo 3", "Intervalo total")
    formatos = Array("0", "@", "@", "dd/mm/yyyy", "@", "dd/mm/yyyy", "@", "@", "@", "[h]:mm", "[h]:mm", "[h]:mm", "[h]:mm", "[h]:mm")

    ' colunas (índice do array); reserva = posições do relatório de 09/10/2026
    Dim cMat As Long, cNome As Long, cLocal As Long, cAdm As Long, cCargo As Long
    Dim cData As Long, cDia As Long, cReg As Long, cCarga As Long
    cMat = 1 - deslCol: cNome = 2 - deslCol: cLocal = 4 - deslCol: cAdm = 7 - deslCol: cCargo = 8 - deslCol
    cData = 10 - deslCol: cDia = 11 - deslCol: cReg = 12 - deslCol
    Dim r As Long, c As Long, t As String
    For r = 1 To UBound(dados, 1)
        If TextoSimples(dados(r, 1)) = "CADASTRO" Then
            For c = 1 To UBound(dados, 2)
                t = TextoSimples(dados(r, c))
                If t = "CADASTRO" Then cMat = c
                If t = "NOME" Then cNome = c
                If t = "LOCAL" Then cLocal = c
                If t = "ADMISSAO" Then cAdm = c
                If t = "CARGO" Then cCargo = c
                If t = "DATA" Then cData = c
                If t = "DIA" Then cDia = c
                If t = "REGISTROS DE PONTO" Then cReg = c
            Next c
            Exit For
        End If
    Next r
    cCarga = cReg + 8

    ReDim saida(1 To UBound(dados, 1), 1 To 14)
    nSaida = 0
    Dim k As Long, marcas As String, diaTxt As String, p As Long, totalInt As Double, vInt As Variant
    For r = 1 To UBound(dados, 1)
        If EhNumero(Celula(dados, r, cMat)) And EhData(Celula(dados, r, cData)) Then
            nSaida = nSaida + 1
            saida(nSaida, 1) = CLng(Celula(dados, r, cMat))
            saida(nSaida, 2) = TextoCel(Celula(dados, r, cNome))
            saida(nSaida, 3) = TextoCel(Celula(dados, r, cLocal))
            If EhData(Celula(dados, r, cAdm)) Then saida(nSaida, 4) = CDate(Int(CDbl(Celula(dados, r, cAdm))))
            saida(nSaida, 5) = TextoCel(Celula(dados, r, cCargo))
            saida(nSaida, 6) = CDate(Int(CDbl(Celula(dados, r, cData))))
            ' "Sex (985)" -> dia "Sex" e código do horário "985"
            diaTxt = TextoCel(Celula(dados, r, cDia))
            p = InStr(diaTxt, "(")
            If p > 0 Then
                saida(nSaida, 7) = Trim$(Left$(diaTxt, p - 1))
                saida(nSaida, 8) = Replace(Trim$(Mid$(diaTxt, p + 1)), ")", "")
            Else
                saida(nSaida, 7) = diaTxt
            End If
            ' batidas: até 8, separadas por espaço ("07:01 11:00 11:00 12:00")
            marcas = ""
            For k = cReg To cReg + 7
                If EhNumero(Celula(dados, r, k)) Then marcas = marcas & IIf(marcas = "", "", " ") & HoraMinuto(Celula(dados, r, k))
            Next k
            saida(nSaida, 9) = marcas
            If EhNumero(Celula(dados, r, cCarga)) Then saida(nSaida, 10) = FracaoDia(Celula(dados, r, cCarga))
            ' intervalos feitos no dia e o total deles
            totalInt = 0
            For k = 1 To 3
                vInt = Celula(dados, r, cCarga + k)
                If EhNumero(vInt) Then
                    saida(nSaida, 10 + k) = FracaoDia(vInt)
                    totalInt = totalInt + FracaoDia(vInt)
                End If
            Next k
            saida(nSaida, 14) = totalInt
        End If
    Next r
End Sub

' ----------------------------------------------------------------------------
' INTERJORNADA (descanso entre um dia e o outro, mínimo de 11h)
' Títulos (linha com "Data Apuração"): Data Apuração | Marcação Anterior |
'   Marcação Atual | Horas Descansadas | Ocorrência | Local | Cargo | C.C.
' Valores deslocados: data de apuração = título+1; cada marcação = data no
'   título+1 e hora no título+2; horas descansadas e ocorrência = título+1;
'   cargo e C.C. na própria coluna do título.
' A pessoa (matrícula na col. A, nome na B, local "código - descrição" na E)
'   vem na mesma linha; uma linha sem matrícula é outra ocorrência da mesma
'   pessoa.
' ----------------------------------------------------------------------------
Private Sub MontarInterjornada(dados As Variant, ByVal deslCol As Long, ByRef cab As Variant, _
    ByRef saida() As Variant, ByRef nSaida As Long, ByRef nomeAba As String, ByRef formatos As Variant)

    nomeAba = "Interjornada"
    cab = Array("Matrícula", "Colaborador", "Cód. Local", "Local", "Cargo", "C.C.", "Data apuração", _
                "Marcação anterior", "Marcação atual", "Horas descansadas", "Ocorrência")
    formatos = Array("0", "@", "@", "@", "@", "@", "dd/mm/yyyy", "dd/mm/yyyy hh:mm", "dd/mm/yyyy hh:mm", "[h]:mm", "@")

    Dim cAp As Long, cAnt As Long, cAtu As Long, cDesc As Long, cOc As Long, cCargo As Long, cCC As Long
    Dim cMat As Long, cNome As Long, cLoc As Long
    cMat = 1 - deslCol: cNome = 2 - deslCol: cLoc = 5 - deslCol
    cAp = 6 - deslCol: cAnt = 8 - deslCol: cAtu = 10 - deslCol: cDesc = 12 - deslCol: cOc = 13 - deslCol
    cCargo = 17 - deslCol: cCC = 20 - deslCol
    Dim r As Long, c As Long, t As String, achou As Boolean
    For r = 1 To UBound(dados, 1)
        For c = 1 To UBound(dados, 2)
            If TextoSimples(dados(r, c)) = "HORAS DESCANSADAS" Then achou = True: Exit For
        Next c
        If achou Then
            For c = 1 To UBound(dados, 2)
                t = TextoSimples(dados(r, c))
                If t = "DATA APURACAO" Then cAp = c + 1
                If t = "MARCACAO ANTERIOR" Then cAnt = c + 1
                If t = "MARCACAO ATUAL" Then cAtu = c + 1
                If t = "HORAS DESCANSADAS" Then cDesc = c + 1
                If t = "OCORRENCIA" Then cOc = c + 1
                If t = "CARGO" Then cCargo = c
                If t = "C.C." Then cCC = c
            Next c
            Exit For
        End If
    Next r

    ReDim saida(1 To UBound(dados, 1), 1 To 11)
    nSaida = 0
    Dim matAtual As Variant, nomeAtual As String, locAtual As String
    matAtual = Empty
    For r = 1 To UBound(dados, 1)
        ' linha com matrícula numérica e nome = começa (ou repete) a pessoa
        If EhNumero(Celula(dados, r, cMat)) And TextoCel(Celula(dados, r, cNome)) <> "" Then
            matAtual = CLng(Celula(dados, r, cMat))
            nomeAtual = TextoCel(Celula(dados, r, cNome))
            locAtual = TextoCel(Celula(dados, r, cLoc))
        End If
        If Not IsEmpty(matAtual) And EhData(Celula(dados, r, cAp)) And EhNumero(Celula(dados, r, cDesc)) Then
            nSaida = nSaida + 1
            saida(nSaida, 1) = matAtual
            saida(nSaida, 2) = nomeAtual
            saida(nSaida, 3) = CodigoDoLocal(locAtual)
            saida(nSaida, 4) = DescricaoDoLocal(locAtual)
            saida(nSaida, 5) = TextoCel(Celula(dados, r, cCargo))
            saida(nSaida, 6) = TextoCel(Celula(dados, r, cCC))
            saida(nSaida, 7) = CDate(Int(CDbl(Celula(dados, r, cAp))))
            saida(nSaida, 8) = DataComHora(Celula(dados, r, cAnt), Celula(dados, r, cAnt + 1))
            saida(nSaida, 9) = DataComHora(Celula(dados, r, cAtu), Celula(dados, r, cAtu + 1))
            saida(nSaida, 10) = FracaoDia(Celula(dados, r, cDesc))
            saida(nSaida, 11) = TextoCel(Celula(dados, r, cOc))
        End If
    Next r
End Sub

' ----------------------------------------------------------------------------
' INTERJORNADA SEMANAL (descanso semanal: trabalhou no DSR/feriado)
' Títulos: Colaborador | Local | Horas DSR | Marcação Antes DSR | Marcação Após DSR
' Cada linha de dado tem, numa célula mesclada (coluna H ou I), um texto como
'   "Trabalhou no Dia do DSR/Feriado (16/08/2026)": a macro separa a
'   ocorrência e a data. Matrícula na col. A, nome na B, local na col. de
'   "Local"; uma linha sem matrícula é outra data da mesma pessoa.
' ----------------------------------------------------------------------------
Private Sub MontarInterSemanal(dados As Variant, ByVal deslCol As Long, ByRef cab As Variant, _
    ByRef saida() As Variant, ByRef nSaida As Long, ByRef nomeAba As String, ByRef formatos As Variant)

    nomeAba = "Interjornada Semanal"
    cab = Array("Matrícula", "Colaborador", "Cód. Local", "Local", "Data DSR/Feriado", "Ocorrência")
    formatos = Array("0", "@", "@", "@", "dd/mm/yyyy", "@")

    Dim cMat As Long, cNome As Long, cLoc As Long, r As Long, c As Long
    cMat = 1 - deslCol: cNome = 2 - deslCol: cLoc = 7 - deslCol
    For r = 1 To UBound(dados, 1)
        If TextoSimples(dados(r, 1)) = "COLABORADOR" Then
            For c = 1 To UBound(dados, 2)
                If TextoSimples(dados(r, c)) = "LOCAL" Then cLoc = c: Exit For
            Next c
            Exit For
        End If
    Next r

    ' "texto (dd/mm/aaaa)" no fim da célula
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^(.*?)\s*\((\d{2})/(\d{2})/(\d{4})\)\s*$"

    ReDim saida(1 To UBound(dados, 1), 1 To 6)
    nSaida = 0
    Dim matAtual As Variant, nomeAtual As String, locTxt As String
    matAtual = Empty
    For r = 1 To UBound(dados, 1)
        If EhNumero(Celula(dados, r, cMat)) And TextoCel(Celula(dados, r, cNome)) <> "" Then
            matAtual = CLng(Celula(dados, r, cMat))
            nomeAtual = TextoCel(Celula(dados, r, cNome))
        End If
        If Not IsEmpty(matAtual) Then
            For c = 1 To UBound(dados, 2)
                If VarType(dados(r, c)) = vbString Then
                    If re.Test(TextoCel(dados(r, c))) Then
                        Set m = re.Execute(TextoCel(dados(r, c))).Item(0)
                        locTxt = TextoCel(Celula(dados, r, cLoc))
                        nSaida = nSaida + 1
                        saida(nSaida, 1) = matAtual
                        saida(nSaida, 2) = nomeAtual
                        saida(nSaida, 3) = CodigoDoLocal(locTxt)
                        saida(nSaida, 4) = DescricaoDoLocal(locTxt)
                        saida(nSaida, 5) = DateSerial(CInt(m.SubMatches(3)), CInt(m.SubMatches(2)), CInt(m.SubMatches(1)))
                        saida(nSaida, 6) = Trim$(m.SubMatches(0))
                        Exit For
                    End If
                End If
            Next c
        End If
    Next r
End Sub

' ----------------------------------------------------------------------------
' HORAS EXCEDENTES (mais de 10h trabalhadas no dia)
' Títulos: Data | Colaborador/Nome | Cargo | C. C. | Fil | Local | Carga Hor. |
'          Local | Marcações | Extras | Horas + Cv.Not. = Total
' Valores: matrícula na coluna de "Colaborador/Nome" e nome uma à direita;
'   carga horária prevista = título+1; local completo ("código - descrição")
'   = 2º título "Local"+1; batidas = 8 colunas a partir de "Marcações"+1;
'   logo depois delas: extras, horas trabalhadas, conversão noturna e total
'   (horas + conversão noturna). O total é o que passa de 10h.
' Uma linha de dado = data + matrícula numérica (o cabeçalho de página se repete).
' ----------------------------------------------------------------------------
Private Sub MontarExcedentes(dados As Variant, ByVal deslCol As Long, ByRef cab As Variant, _
    ByRef saida() As Variant, ByRef nSaida As Long, ByRef nomeAba As String, ByRef formatos As Variant)

    nomeAba = "Horas Excedentes"
    cab = Array("Matrícula", "Colaborador", "Cargo", "C.C.", "Filial", "Cód. Local", "Local", "Data", _
                "Carga horária", "Marcações", "Extras", "Horas trabalhadas", "Conv. noturna", "Total")
    formatos = Array("0", "@", "@", "@", "@", "@", "@", "dd/mm/yyyy", "[h]:mm", "@", "[h]:mm", "[h]:mm", "[h]:mm", "[h]:mm")

    Dim cData As Long, cMat As Long, cCargo As Long, cCC As Long, cFil As Long
    Dim cLocal1 As Long, cLocal2 As Long, cCarga As Long, cMarc As Long, cExtra As Long
    cData = 1 - deslCol: cMat = 2 - deslCol: cCargo = 5 - deslCol: cCC = 7 - deslCol: cFil = 8 - deslCol
    cLocal1 = 9 - deslCol: cCarga = 11 - deslCol: cLocal2 = 12 - deslCol: cMarc = 13 - deslCol
    Dim r As Long, c As Long, t As String, nLocais As Long
    For r = 1 To UBound(dados, 1)
        If TextoSimples(dados(r, 1)) = "DATA" Then
            For c = 1 To UBound(dados, 2)
                t = TextoSimples(dados(r, c))
                If t = "COLABORADOR/NOME" Then cMat = c
                If t = "CARGO" Then cCargo = c
                If t = "C. C." Then cCC = c
                If t = "FIL" Then cFil = c
                If t = "CARGA HOR." Then cCarga = c + 1
                If t = "MARCACOES" Then cMarc = c + 1
                If t = "LOCAL" Then
                    nLocais = nLocais + 1
                    If nLocais = 1 Then cLocal1 = c Else cLocal2 = c + 1
                End If
            Next c
            Exit For
        End If
    Next r
    cExtra = cMarc + 8

    ReDim saida(1 To UBound(dados, 1), 1 To 14)
    nSaida = 0
    Dim k As Long, marcas As String, locTxt As String, vFil As Variant
    For r = 1 To UBound(dados, 1)
        If EhData(Celula(dados, r, cData)) And EhNumero(Celula(dados, r, cMat)) Then
            nSaida = nSaida + 1
            saida(nSaida, 1) = CLng(Celula(dados, r, cMat))
            saida(nSaida, 2) = TextoCel(Celula(dados, r, cMat + 1))
            saida(nSaida, 3) = TextoCel(Celula(dados, r, cCargo))
            saida(nSaida, 4) = TextoCel(Celula(dados, r, cCC))
            vFil = Celula(dados, r, cFil)
            If EhNumero(vFil) Then saida(nSaida, 5) = Format$(CLng(vFil), "000") Else saida(nSaida, 5) = TextoCel(vFil)
            locTxt = TextoCel(Celula(dados, r, cLocal2))
            saida(nSaida, 6) = CodigoDoLocal(locTxt)
            saida(nSaida, 7) = DescricaoDoLocal(locTxt)
            If saida(nSaida, 7) = "" Then saida(nSaida, 7) = TextoCel(Celula(dados, r, cLocal1))
            saida(nSaida, 8) = CDate(Int(CDbl(Celula(dados, r, cData))))
            If EhNumero(Celula(dados, r, cCarga)) Then saida(nSaida, 9) = FracaoDia(Celula(dados, r, cCarga))
            marcas = ""
            For k = cMarc To cMarc + 7
                If EhNumero(Celula(dados, r, k)) Then marcas = marcas & IIf(marcas = "", "", " ") & HoraMinuto(Celula(dados, r, k))
            Next k
            saida(nSaida, 10) = marcas
            For k = 0 To 3
                If EhNumero(Celula(dados, r, cExtra + k)) Then saida(nSaida, 11 + k) = FracaoDia(Celula(dados, r, cExtra + k))
            Next k
        End If
    Next r
End Sub

' ----------------------------------------------------------------------------
' Apaga a aba e escreve: cabeçalho na linha 1, formatos por coluna ANTES dos
' valores (texto "@" mantém zeros à esquerda; "[h]:mm" soma mais de 24h),
' linhas de dado, bordas, filtro, fonte, larguras e linha 1 congelada.
' ----------------------------------------------------------------------------
Private Sub EscreverTabela(ws As Worksheet, cab As Variant, saida() As Variant, ByVal nSaida As Long, formatos As Variant)
    Dim nCols As Long, r As Long, c As Long
    nCols = UBound(cab) + 1

    Application.ScreenUpdating = False
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.UnMerge
    On Error GoTo 0
    ws.Cells.Clear

    For c = 1 To nCols
        ws.Cells(1, c).Value = cab(c - 1)
        ws.Range(ws.Cells(2, c), ws.Cells(nSaida + 1, c)).NumberFormat = formatos(c - 1)
    Next c

    Dim bloco() As Variant
    ReDim bloco(1 To nSaida, 1 To nCols)
    For r = 1 To nSaida
        For c = 1 To nCols
            bloco(r, c) = saida(r, c)
        Next c
    Next r
    ws.Range(ws.Cells(2, 1), ws.Cells(nSaida + 1, nCols)).Value = bloco

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols))
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(nSaida + 1, nCols))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .AutoFilter
    End With
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Columns.AutoFit
    ws.Columns(2).ColumnWidth = 36

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select
    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------------------------------------
' FUNÇÕES DE APOIO (Private: usadas só neste módulo, não aparecem como macro)
' ----------------------------------------------------------------------------

' Valor do array; Empty se a coluna estiver fora dele.
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

' Data do Excel (número de série maior que 1000; hora sozinha é menor que 1).
Private Function EhData(v As Variant) As Boolean
    EhData = False
    If EhNumero(v) Then EhData = (CDbl(v) > 1000)
End Function

' Só a parte da hora, arredondada ao minuto, como fração do dia (0,5 = 12:00).
' O sistema grava horas com resíduo de milissegundos (10:05 vira 10:05:00,96).
Private Function FracaoDia(v As Variant) As Double
    FracaoDia = Round((CDbl(v) - Int(CDbl(v))) * 1440, 0) / 1440
End Function

' Hora como texto "hh:mm" (para a coluna Marcações).
Private Function HoraMinuto(v As Variant) As String
    Dim m As Long
    m = CLng(Round((CDbl(v) - Int(CDbl(v))) * 1440, 0))
    HoraMinuto = Format$((m \ 60) Mod 24, "00") & ":" & Format$(m Mod 60, "00")
End Function

' Data (vData) + hora (vHora) num só valor; Empty se não houver data.
Private Function DataComHora(vData As Variant, vHora As Variant) As Variant
    If Not EhData(vData) Then
        DataComHora = Empty
    ElseIf EhNumero(vHora) Then
        DataComHora = CDate(Int(CDbl(vData)) + FracaoDia(vHora))
    Else
        DataComHora = CDate(Int(CDbl(vData)))
    End If
End Function

' "2.028.30.306 - Setor Operacional 3ºT" -> "2.028.30.306" (sem " - ": "").
Private Function CodigoDoLocal(ByVal t As String) As String
    Dim p As Long
    p = InStr(t, " - ")
    If p > 0 Then CodigoDoLocal = Trim$(Left$(t, p - 1)) Else CodigoDoLocal = ""
End Function

' "2.028.30.306 - Setor Operacional 3ºT" -> "Setor Operacional 3ºT" (sem " - ": o texto todo).
Private Function DescricaoDoLocal(ByVal t As String) As String
    Dim p As Long
    p = InStr(t, " - ")
    If p > 0 Then DescricaoDoLocal = Trim$(Mid$(t, p + 3)) Else DescricaoDoLocal = Trim$(t)
End Function

' Texto sem espaços nas pontas; erro do Excel (#N/A) vira "".
Private Function TextoCel(v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then
        TextoCel = ""
    Else
        TextoCel = Trim$(Replace(Replace(CStr(v), vbCr, " "), vbLf, " "))
    End If
End Function

' Maiúsculas e sem acento, para comparar títulos ("Situação" = "SITUACAO").
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
