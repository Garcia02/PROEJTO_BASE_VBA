'===============================================================================
' MÓDULO: modUtils
' RESPONSABILIDADE: Biblioteca de funções universais transversais: conversões
'                   seguras de tipos, detecção de dimensões de relatório,
'                   lookups na aba TAB, agregação/separação de valores e
'                   extração de veículo de C.E. Agnóstico a regras de negócio —
'                   apenas rotinas consumidas por dois ou mais módulos.
'===============================================================================
Option Explicit

' Índices das tabelas TAB mantidos durante o processamento atual.
Private mCacheIndicesTAB As Object

' ------------------------------------------------------------------
' CONVERSÕES SEGURAS DE TIPO
' ------------------------------------------------------------------

' Converte qualquer valor para String sem notação científica, preservando
' zeros à esquerda (identificadores: DPS, fatura, C.E., modalidade, chaves).
Public Function ParaString(valor As Variant) As String
    If IsMissing(valor) Or IsNull(valor) Or IsEmpty(valor) Then
        ParaString = ""
        Exit Function
    End If
    If VarType(valor) = vbError Then
        ParaString = ""
        Exit Function
    End If
    If IsNumeric(valor) Then
        ' Evita notação científica (ex.: 1,2345E+14)
        If InStr(1, CStr(valor), "E", vbTextCompare) > 0 Then
            ParaString = Format(valor, "0")
        Else
            ParaString = CStr(valor)
        End If
    Else
        ParaString = CStr(valor)
    End If
    ParaString = Trim(ParaString)
End Function

' Converte qualquer valor para Double, neutralizando vazio/erro/texto como 0,00.
Public Function ParaDouble(valor As Variant) As Double
    On Error GoTo Tratar
    If IsMissing(valor) Or IsNull(valor) Or IsEmpty(valor) Then
        ParaDouble = 0#
        Exit Function
    End If
    If VarType(valor) = vbError Then
        ParaDouble = 0#
        Exit Function
    End If
    If IsNumeric(valor) Then
        ParaDouble = CDbl(valor)
    Else
        ParaDouble = val(CStr(valor))
    End If
    Exit Function
Tratar:
    ParaDouble = 0#
End Function

' ------------------------------------------------------------------
' DIMENSÕES DE RELATÓRIO E LOCALIZAÇÃO DE COLUNAS
' ------------------------------------------------------------------

' Normaliza nomes de coluna para comparação tolerante a acentos, espaços e
' pequenas variações de escrita (ex.: "PONTO DE OPERACAO" x "PONTO DE OPERAÇÃO").
Private Function NormalizarCabecalho(valor As Variant) As String
    Dim s As String
    s = Trim(CStr(valor))
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    s = UCase(s)
    s = Replace(s, "Ã", "A")
    s = Replace(s, "Á", "A")
    s = Replace(s, "Â", "A")
    s = Replace(s, "À", "A")
    s = Replace(s, "É", "E")
    s = Replace(s, "Ê", "E")
    s = Replace(s, "Í", "I")
    s = Replace(s, "Ó", "O")
    s = Replace(s, "Ô", "O")
    s = Replace(s, "Ú", "U")
    s = Replace(s, "Ç", "C")
    NormalizarCabecalho = s
End Function

' Localiza o índice (1-based) de uma coluna pelo nome, ignorando maiúsculas/
' minúsculas, espaços extras e variações de acentos/acentuação.
Public Function EncontrarIndiceColuna(nomesColunas As Variant, nomeColuna As String) As Long
    Dim i As Long
    Dim nomeNormalizado As String
    nomeNormalizado = NormalizarCabecalho(nomeColuna)
    EncontrarIndiceColuna = 0
    If Not IsArray(nomesColunas) Then Exit Function
    For i = LBound(nomesColunas) To UBound(nomesColunas)
        If StrComp(NormalizarCabecalho(nomesColunas(i)), nomeNormalizado, vbBinaryCompare) = 0 Then
            EncontrarIndiceColuna = i
            Exit For
        End If
    Next i
End Function

' Verifica se uma linha é um cabeçalho plausível: pelo menos duas células
' preenchidas em sequência (evita títulos gerais ou linhas com apenas um valor).
Private Function LinhaEhCabecalhoPossivel(ws As Worksheet, linha As Long, _
                                          primeiraCol As Long, ultimaCol As Long) As Boolean
    Dim i As Long
    Dim contagem As Long
    Dim valor As String

    LinhaEhCabecalhoPossivel = False
    If linha < 1 Or ws Is Nothing Then Exit Function

    If ultimaCol < 2 Then Exit Function

    contagem = 0
    For i = primeiraCol To ultimaCol
        valor = Trim(CStr(ws.Cells(linha, i).value))
        If Len(valor) > 0 Then
            contagem = contagem + 1
        End If
    Next i

    If contagem < 2 Then Exit Function
    LinhaEhCabecalhoPossivel = True
End Function

' Detecta a estrutura de um relatório: linha de cabeçalho, limites de dados e
' nomes de colunas. Retorna Dictionary com as chaves:
'   temDados (Boolean), linhaHeader (Long), primeiraColuna (Long),
'   ultimaColuna (Long), ultimaLinha (Long), colunas (Long),
'   totalRegistros (Long), nomesColunas (array 1-based).
' Obs.: "colunas" e "ultimaColuna" são equivalentes; "totalRegistros" e
' "ultimaLinha" também — mantidas ambas para compatibilidade de consumo.
Public Function CapturarDimensoesRelatorio(ws As Worksheet) As Object
    Dim dict As Object
    Dim i As Long, ultCol As Long, ultLinha As Long, linhaPossivel As Long
    Dim primeiraColunaUsada As Long, primeiraLinhaUsada As Long
    Dim nomes() As String
    Dim areaUsada As Range

    Set dict = CreateObject("Scripting.Dictionary")
    dict("temDados") = False
    dict("linhaHeader") = 0
    dict("primeiraColuna") = 1
    dict("ultimaColuna") = 0
    dict("ultimaLinha") = 0
    dict("colunas") = 0
    dict("totalColunas") = 0
    dict("totalRegistros") = 0

    If ws Is Nothing Then
        Set CapturarDimensoesRelatorio = dict
        Exit Function
    End If

    Set areaUsada = ws.UsedRange
    If areaUsada Is Nothing Then
        Set CapturarDimensoesRelatorio = dict
        Exit Function
    End If

    primeiraColunaUsada = areaUsada.Column
    primeiraLinhaUsada = areaUsada.Row
    ultCol = areaUsada.Column + areaUsada.Columns.Count - 1
    ultLinha = areaUsada.Row + areaUsada.Rows.Count - 1

    linhaPossivel = 0
    For i = primeiraLinhaUsada To ultLinha
        If LinhaEhCabecalhoPossivel(ws, i, primeiraColunaUsada, ultCol) Then
            linhaPossivel = i
            Exit For
        End If
    Next i

    If linhaPossivel = 0 Then
        ' Fallback conservador: caso não encontre um cabeçalho plausível, usa a
        ' primeira linha não vazia, preservando compatibilidade com relatórios
        ' que já estejam em padrão antigo.
        For i = primeiraLinhaUsada To ultLinha
            If Len(Trim(CStr(ws.Cells(i, primeiraColunaUsada).value))) > 0 Then
                linhaPossivel = i
                Exit For
            End If
        Next i
    End If

    If linhaPossivel = 0 Then
        Set CapturarDimensoesRelatorio = dict
        Exit Function
    End If

    dict("linhaHeader") = linhaPossivel
    dict("primeiraColuna") = primeiraColunaUsada
    dict("ultimaColuna") = ultCol
    dict("ultimaLinha") = ultLinha
    dict("colunas") = ultCol - primeiraColunaUsada + 1
    dict("totalColunas") = ultCol - primeiraColunaUsada + 1
    dict("totalRegistros") = IIf(ultLinha > dict("linhaHeader"), ultLinha - dict("linhaHeader"), 0)
    dict("temDados") = (ultLinha > dict("linhaHeader"))

    ReDim nomes(1 To dict("totalColunas"))
    For i = 1 To dict("totalColunas")
        nomes(i) = CStr(ws.Cells(dict("linhaHeader"), primeiraColunaUsada + i - 1).value)
    Next i
    dict("nomesColunas") = nomes

    Set CapturarDimensoesRelatorio = dict
End Function

' ------------------------------------------------------------------
' LOOKUP EM TABELAS TAB (aba TAB / ListObjects)
' ------------------------------------------------------------------

' Limpa o cache quando as tabelas TAB forem alteradas durante o processamento.
Public Sub LimparCacheConsultasTAB()
    Set mCacheIndicesTAB = Nothing
End Sub

' Retorna o cache de índices inicializado para o processamento atual.
Private Function ObterCacheIndicesTAB() As Object
    If mCacheIndicesTAB Is Nothing Then
        Set mCacheIndicesTAB = CreateObject("Scripting.Dictionary")
    End If
    Set ObterCacheIndicesTAB = mCacheIndicesTAB
End Function

' Constrói um índice normalizado para uma combinação tabela/colunas.
Private Function ObterIndiceTabelaTAB(ByVal nomeTabela As String, _
                                      ByVal colunaReferencia As String, _
                                      ByVal colunaRetorno As String) As Object
    Dim indices As Object
    Dim indice As Object
    Dim wsTAB As Worksheet
    Dim tbl As ListObject
    Dim i As Long, idxRef As Long, idxRet As Long
    Dim chaveIndice As String
    Dim cabecalho As String
    Dim chave As String
    
    Set indices = ObterCacheIndicesTAB()
    chaveIndice = UCase(Trim(nomeTabela)) & "|" & _
                  NormalizarCabecalho(colunaReferencia) & "|" & _
                  NormalizarCabecalho(colunaRetorno)
    If indices.Exists(chaveIndice) Then
        Set ObterIndiceTabelaTAB = indices(chaveIndice)
        Exit Function
    End If
    
    Set indice = CreateObject("Scripting.Dictionary")
    On Error GoTo ArmazenarIndice
    
    Set wsTAB = ThisWorkbook.Worksheets(SHEET_TAB)
    Set tbl = wsTAB.ListObjects(nomeTabela)
    If tbl Is Nothing Then GoTo ArmazenarIndice
    
    For i = 1 To tbl.ListColumns.Count
        cabecalho = NormalizarCabecalho(tbl.ListColumns(i).Name)
        If StrComp(cabecalho, NormalizarCabecalho(colunaReferencia), vbBinaryCompare) = 0 Then idxRef = i
        If StrComp(cabecalho, NormalizarCabecalho(colunaRetorno), vbBinaryCompare) = 0 Then idxRet = i
    Next i
    If idxRef = 0 Or idxRet = 0 Then GoTo ArmazenarIndice
    If tbl.DataBodyRange Is Nothing Then GoTo ArmazenarIndice
    
    For i = 1 To tbl.ListRows.Count
        chave = NormalizarCabecalho(tbl.DataBodyRange(i, idxRef).value)
        If Len(chave) > 0 Then
            If Not indice.Exists(chave) Then
                indice.Add chave, ParaString(tbl.DataBodyRange(i, idxRet).value)
            End If
        End If
    Next i

ArmazenarIndice:
    If Not indices.Exists(chaveIndice) Then indices.Add chaveIndice, indice
    Set ObterIndiceTabelaTAB = indice
End Function

' Consulta estilo PROCV em um ListObject da aba TAB. Retorna o valor da
' coluna de retorno da primeira linha que casar com o valor de referência.
' Retorna "" se não encontrar ou se a tabela não existir.
Public Function ConsultarTabelaTAB(nomeTabela As String, colunaReferencia As String, _
                                   colunaRetorno As String, valorReferencia As String) As String
    Dim indice As Object
    ConsultarTabelaTAB = ""
    Set indice = ObterIndiceTabelaTAB(nomeTabela, colunaReferencia, colunaRetorno)
    If indice.Exists(NormalizarCabecalho(valorReferencia)) Then
        ConsultarTabelaTAB = CStr(indice(NormalizarCabecalho(valorReferencia)))
    End If
End Function

' ------------------------------------------------------------------
' AGREGAÇÃO E SEPARAÇÃO DE VALORES
' ------------------------------------------------------------------

' Agrega um novo valor a uma string de valores distintos, deduplicando por
' UCase e preservando o casing do primeiro valor adicionado.
Public Function AgregarValores(valoresAtuais As String, novoValor As String, _
                               delimitador As String) As String
    Dim v As String
    Dim partes() As String
    Dim i As Long

    v = ParaString(novoValor)
    If Len(v) = 0 Then
        AgregarValores = valoresAtuais
        Exit Function
    End If
    If Len(valoresAtuais) = 0 Then
        AgregarValores = v
        Exit Function
    End If

    partes = Split(valoresAtuais, delimitador)
    For i = LBound(partes) To UBound(partes)
        If StrComp(partes(i), v, vbTextCompare) = 0 Then
            AgregarValores = valoresAtuais
            Exit Function
        End If
    Next i

    AgregarValores = valoresAtuais & delimitador & v
End Function

' Separa uma string de valores pelo delimitador em um array 0-based.
' String vazia retorna array com um único item "".
Public Function SepararValores(valores As String, delimitador As String) As String()
    Dim arr() As String
    If Len(Trim(valores)) = 0 Then
        ReDim arr(0 To 0)
        arr(0) = ""
    Else
        arr = Split(valores, delimitador)
    End If
    SepararValores = arr
End Function

' Conta ocorrências separadas pelo delimitador (mínimo 1, mesmo se vazio).
Public Function ContarValores(valores As String, delimitador As String) As Long
    If Len(Trim(valores)) = 0 Then
        ContarValores = 1
    Else
        ContarValores = UBound(Split(valores, delimitador)) + 1
    End If
End Function

' Retorna o tempo decorrido em segundos, inclusive quando o processamento
' atravessa a meia-noite.
Public Function TempoDecorrido(ByVal inicio As Double) As Double
    Dim decorrido As Double
    decorrido = Timer - inicio
    If decorrido < 0 Then decorrido = decorrido + 86400#
    TempoDecorrido = decorrido
End Function

' Soma valores numéricos concatenados com o delimitador (reserva futura).
Public Function AgregarValorNumerico(valores As String, delimitador As String) As Double
    Dim partes() As String
    Dim i As Long
    Dim total As Double

    total = 0#
    If Len(Trim(valores)) > 0 Then
        partes = Split(valores, delimitador)
        For i = LBound(partes) To UBound(partes)
            total = total + ParaDouble(partes(i))
        Next i
    End If
    AgregarValorNumerico = total
End Function

' ------------------------------------------------------------------
' COLETA DE VALORES ÚNICOS
' ------------------------------------------------------------------

' Coleta valores únicos (chave UCase, valor original) de uma coluna do array.
Public Function ColetarValoresUnicos(arrDados As Variant, colIndex As Long, _
                                     linhaInicio As Long, linhaFim As Long) As Object
    Dim dict As Object
    Dim i As Long
    Dim v As String

    Set dict = CreateObject("Scripting.Dictionary")
    If Not IsArray(arrDados) Then
        Set ColetarValoresUnicos = dict
        Exit Function
    End If
    If linhaFim > UBound(arrDados, 1) Then linhaFim = UBound(arrDados, 1)
    If linhaInicio < LBound(arrDados, 1) Then linhaInicio = LBound(arrDados, 1)

    For i = linhaInicio To linhaFim
        v = ParaString(arrDados(i, colIndex))
        If Len(v) > 0 Then
            If Not dict.Exists(UCase(v)) Then dict.Add UCase(v), v
        End If
    Next i

    Set ColetarValoresUnicos = dict
End Function

' ------------------------------------------------------------------
' VALORES ÚNICOS EM COLUNA POR NOME
' ------------------------------------------------------------------

' Coleta valores únicos de uma coluna pelo nome da coluna e concatena em uma
' string separada por ", ". Usa a aba ativa do workbook, sem exigir validação de
' nome do relatório. Útil para gerar a lista de SHIP SELL para o relatório OTM.
Public Function ObterValoresUnicosColunaPorNome(ws As Worksheet, nomeColuna As String) As String
    Dim arrDados As Variant
    Dim nomesColunas() As String
    Dim totalLinhas As Long, totalColunas As Long
    Dim indiceColuna As Long, i As Long, n As Long
    Dim valor As String
    Dim dic As Object
    Dim valores() As String
    Dim chave As Variant

    ObterValoresUnicosColunaPorNome = ""
    If ws Is Nothing Then Exit Function

    totalLinhas = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If totalLinhas < 2 Then Exit Function

    totalColunas = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If totalColunas < 1 Then Exit Function

    arrDados = ws.Range(ws.Cells(1, 1), ws.Cells(totalLinhas, totalColunas)).value
    ReDim nomesColunas(1 To totalColunas)
    For i = 1 To totalColunas
        nomesColunas(i) = CStr(arrDados(1, i))
    Next i

    indiceColuna = EncontrarIndiceColuna(nomesColunas, nomeColuna)
    If indiceColuna = 0 Then Exit Function

    Set dic = CreateObject("Scripting.Dictionary")
    For i = 2 To totalLinhas
        valor = ParaString(arrDados(i, indiceColuna))
        If Len(Trim(valor)) > 0 Then
            If Not dic.Exists(UCase(valor)) Then dic.Add UCase(valor), valor
        End If
    Next i

    If dic.Count = 0 Then Exit Function

    ReDim valores(0 To dic.Count - 1)
    n = 0
    For Each chave In dic.Keys
        valores(n) = CStr(dic(chave))
        n = n + 1
    Next chave

    ObterValoresUnicosColunaPorNome = Join(valores, ",")
End Function

Public Function ProcurarIndiceColunaPorSinonimos(nomesColunas As Variant, ParamArray aliases() As Variant) As Long
    Dim i As Long, j As Long
    Dim nomeColuna As String
    Dim aliasAtual As String
    Dim melhorIndice As Long
    Dim melhorScore As Long
    Dim scoreAtual As Long

    ProcurarIndiceColunaPorSinonimos = 0
    melhorIndice = 0
    melhorScore = -1
    If Not IsArray(nomesColunas) Then Exit Function

    For i = LBound(nomesColunas) To UBound(nomesColunas)
        nomeColuna = NormalizarCabecalho(nomesColunas(i))
        If Len(nomeColuna) = 0 Then GoTo ProximoIndice

        For j = LBound(aliases) To UBound(aliases)
            aliasAtual = NormalizarCabecalho(aliases(j))
            If Len(aliasAtual) = 0 Then GoTo ProximoAlias

            If StrComp(nomeColuna, aliasAtual, vbBinaryCompare) = 0 Then
                scoreAtual = 1000 + Len(aliasAtual)
            ElseIf InStr(1, nomeColuna, aliasAtual, vbTextCompare) > 0 Then
                scoreAtual = 200 + Len(aliasAtual)
            ElseIf InStr(1, aliasAtual, nomeColuna, vbTextCompare) > 0 Then
                scoreAtual = 150 + Len(aliasAtual)
            Else
                scoreAtual = 0
            End If

            If scoreAtual > melhorScore Then
                melhorScore = scoreAtual
                melhorIndice = i
            End If
ProximoAlias:
        Next j
ProximoIndice:
    Next i

    If melhorScore > 0 Then
        ProcurarIndiceColunaPorSinonimos = melhorIndice
    End If
End Function

Public Function GerarChaveProcessoNF(ByVal numeroDPS As Variant, ByVal notasFiscais As Variant) As String
    Dim dps As String
    Dim nf As String
    Dim partes() As String
    Dim i As Long
    Dim chave As String

    dps = UCase(Trim(ParaString(numeroDPS)))
    nf = UCase(Trim(ParaString(notasFiscais)))
    nf = Replace(nf, " ", "")
    nf = Replace(nf, vbCr, "")
    nf = Replace(nf, vbLf, "")
    nf = Replace(nf, vbTab, "")

    If Len(nf) > 0 Then
        If InStr(1, nf, "/", vbTextCompare) > 0 Then
            partes = Split(nf, "/")
            For i = LBound(partes) To UBound(partes)
                chave = CStr(Trim(partes(i)))
                If Len(chave) > 0 Then
                    If Len(dps) > 0 Then
                        GerarChaveProcessoNF = dps & "|" & chave
                    Else
                        GerarChaveProcessoNF = chave
                    End If
                    Exit Function
                End If
            Next i
        End If
    End If

    If Len(dps) > 0 And Len(nf) > 0 Then
        GerarChaveProcessoNF = dps & "|" & nf
    ElseIf Len(dps) > 0 Then
        GerarChaveProcessoNF = dps
    Else
        GerarChaveProcessoNF = nf
    End If
End Function

' Extrai campos específicos da coluna PK do relatório OTM.
' A série fica entre o primeiro e o segundo "_" e a nota entre o quinto
' e o sexto "_".
Public Function ExtrairCampoPK(ByVal valorPK As Variant, ByVal campo As String) As String
    Dim partes() As String
    Dim chaveCampo As String

    ExtrairCampoPK = ""
    chaveCampo = UCase(Trim(campo))
    partes = Split(ParaString(valorPK), "_")

    If chaveCampo = "SERIE" Then
        If UBound(partes) >= 1 Then ExtrairCampoPK = Trim(partes(1))
    ElseIf chaveCampo = "NOTA" Then
        If UBound(partes) >= 5 Then ExtrairCampoPK = Trim(partes(5))
    End If
End Function

' Mantém apenas data e hora do valor OTM, descartando o fuso horário.
' Ex.: "07/08/2026 08:28 America/Sao_Paulo" -> "07/08/2026 08:28".
Public Function ExtrairDataHora(ByVal valor As Variant) As String
    Dim partes() As String
    Dim texto As String

    ExtrairDataHora = ""
    texto = Trim(ParaString(valor))
    If Len(texto) = 0 Then Exit Function

    partes = Split(texto, " ")
    If UBound(partes) >= 1 Then
        ExtrairDataHora = Trim(partes(0)) & " " & Trim(partes(1))
    Else
        ExtrairDataHora = texto
    End If
End Function

' Conta quantos valores únicos existem em uma coluna pelo nome da coluna.
' Utilizado para exibir a quantidade de processos na tela do relatório OTM.
Public Function ContarValoresUnicosColunaPorNome(ws As Worksheet, nomeColuna As String) As Long
    Dim arrDados As Variant
    Dim nomesColunas() As String
    Dim totalLinhas As Long, totalColunas As Long
    Dim indiceColuna As Long, i As Long
    Dim valor As String
    Dim dic As Object

    ContarValoresUnicosColunaPorNome = 0
    If ws Is Nothing Then Exit Function

    totalLinhas = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If totalLinhas < 2 Then Exit Function

    totalColunas = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If totalColunas < 1 Then Exit Function

    arrDados = ws.Range(ws.Cells(1, 1), ws.Cells(totalLinhas, totalColunas)).value
    ReDim nomesColunas(1 To totalColunas)
    For i = 1 To totalColunas
        nomesColunas(i) = CStr(arrDados(1, i))
    Next i

    indiceColuna = EncontrarIndiceColuna(nomesColunas, nomeColuna)
    If indiceColuna = 0 Then Exit Function

    Set dic = CreateObject("Scripting.Dictionary")
    For i = 2 To totalLinhas
        valor = ParaString(arrDados(i, indiceColuna))
        If Len(Trim(valor)) > 0 Then
            If Not dic.Exists(UCase(valor)) Then dic.Add UCase(valor), valor
        End If
    Next i

    ContarValoresUnicosColunaPorNome = dic.Count
End Function

' Conta a quantidade total de valores preenchidos em uma coluna pelo nome.
' Diferente da contagem de valores únicos: mede o total de processos/linhas.
Public Function ContarValoresColunaPorNome(ws As Worksheet, nomeColuna As String) As Long
    Dim arrDados As Variant
    Dim nomesColunas() As String
    Dim totalLinhas As Long, totalColunas As Long
    Dim indiceColuna As Long, i As Long
    Dim valor As String

    ContarValoresColunaPorNome = 0
    If ws Is Nothing Then Exit Function

    totalLinhas = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If totalLinhas < 2 Then Exit Function

    totalColunas = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If totalColunas < 1 Then Exit Function

    arrDados = ws.Range(ws.Cells(1, 1), ws.Cells(totalLinhas, totalColunas)).value
    ReDim nomesColunas(1 To totalColunas)
    For i = 1 To totalColunas
        nomesColunas(i) = CStr(arrDados(1, i))
    Next i

    indiceColuna = EncontrarIndiceColuna(nomesColunas, nomeColuna)
    If indiceColuna = 0 Then Exit Function

    For i = 2 To totalLinhas
        valor = ParaString(arrDados(i, indiceColuna))
        If Len(Trim(valor)) > 0 Then
            ContarValoresColunaPorNome = ContarValoresColunaPorNome + 1
        End If
    Next i
End Function

' ------------------------------------------------------------------
' VEÍCULO A PARTIR DE C.E. E MODALIDADE
' ------------------------------------------------------------------

' Extrai o veículo subtraindo a MODALIDADE do prefixo do C.E.
' Ex.: C.E.="FTL_DIST_TRANSF_6000_FIORINO", MODALIDADE="FTL_DIST_TRANSF_6000"
'      -> "FIORINO". Usa o 1º bloco se houver "/". Se o C.E. não tiver a
'      MODALIDADE como prefixo, retorna "" (evita poluição cadastral).
Public Function ExtrairVeiculoDeCE(ce As String, modalidade As String) As String
    Dim bloco As String

    ExtrairVeiculoDeCE = ""
    If Len(Trim(ce)) = 0 Then Exit Function

    ' Primeiro bloco quando há múltiplos valores separados por "/"
    If InStr(1, ce, "/") > 0 Then
        bloco = Trim(Split(ce, "/")(0))
    Else
        bloco = Trim(ce)
    End If

    ' Somente extrai se a MODALIDADE for prefixo do C.E.
    If Len(Trim(modalidade)) > 0 And Left(bloco, Len(modalidade)) = modalidade Then
        ExtrairVeiculoDeCE = Mid(bloco, Len(modalidade) + 1)
    End If
End Function

' Coleta veículos únicos extraídos de C.E. + MODALIDADE (chave UCase).
Public Function ColetarVeiculosDeCE(arrDados As Variant, colCE As Long, _
                                    colModalidade As Long, linhaInicio As Long, _
                                    linhaFim As Long) As Object
    Dim dict As Object
    Dim i As Long
    Dim v As String

    Set dict = CreateObject("Scripting.Dictionary")
    If Not IsArray(arrDados) Then
        Set ColetarVeiculosDeCE = dict
        Exit Function
    End If
    If linhaFim > UBound(arrDados, 1) Then linhaFim = UBound(arrDados, 1)
    If linhaInicio < LBound(arrDados, 1) Then linhaInicio = LBound(arrDados, 1)

    For i = linhaInicio To linhaFim
        v = ExtrairVeiculoDeCE(ParaString(arrDados(i, colCE)), ParaString(arrDados(i, colModalidade)))
        If Len(v) > 0 Then
            If Not dict.Exists(UCase(v)) Then dict.Add UCase(v), v
        End If
    Next i

    Set ColetarVeiculosDeCE = dict
End Function

' ------------------------------------------------------------------
' ARQUIVOS
' ------------------------------------------------------------------

' Extrai o nome do arquivo (com extensão) de um caminho completo.
Public Function ExtrairNomeArquivo(caminho As String) As String
    Dim s As String
    s = caminho
    If InStrRev(s, "\") > 0 Then s = Mid(s, InStrRev(s, "\") + 1)
    If InStrRev(s, "/") > 0 Then s = Mid(s, InStrRev(s, "/") + 1)
    ExtrairNomeArquivo = s
End Function

' ------------------------------------------------------------------
' CONTRATO DE MAPEAMENTO (modLayout)
' ------------------------------------------------------------------

' Localiza a posição (1-based) do par de mapeamento cuja coluna de DESTINO
' corresponde ao nome informado. Retorna 0 se não encontrar. Usado pelo
' modTransferencia (NOTAS FISCAIS, Indice) e modTratamento (colunas de lookup).
Public Function LocalizarColunaDestino(mapeamento As Collection, nomeDestino As String) As Long
    Dim i As Long
    Dim par As Variant

    LocalizarColunaDestino = 0
    For i = 1 To mapeamento.Count
        par = mapeamento(i)
        If StrComp(CStr(par(1)), nomeDestino, vbTextCompare) = 0 Then
            LocalizarColunaDestino = i
            Exit For
        End If
    Next i
End Function

' Captura o estado global do Excel antes de uma operação que o altere.
Public Sub CapturarEstadoExcel(ByRef screenUpdating As Boolean, _
                               ByRef enableEvents As Boolean, _
                               ByRef calculation As XlCalculation)
    screenUpdating = Application.ScreenUpdating
    enableEvents = Application.EnableEvents
    calculation = Application.Calculation
End Sub

' Restaura o estado global do Excel mesmo quando uma etapa falha.
Public Sub RestaurarEstadoExcel(ByVal screenUpdating As Boolean, _
                                ByVal enableEvents As Boolean, _
                                ByVal calculation As XlCalculation)
    On Error Resume Next
    Application.ScreenUpdating = screenUpdating
    Application.EnableEvents = enableEvents
    Application.Calculation = calculation
    On Error GoTo 0
End Sub