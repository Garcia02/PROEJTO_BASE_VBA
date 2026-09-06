'===============================================================================
' MÓDULO: modMesclagem
' RESPONSABILIDADE: Mesclagem de dados — processar e mesclar relatórios de
'                   despesa com os dados já importados, somando valores por
'                   chave, separando por modalidade e trazendo todas as
'                   colunas do relatório de despesa (valores únicos concatenados
'                   com "|"). Colunas que já existem na receita recebem prefixo.
'                   Ao final, registra na aba DESPESA NE os processos da receita
'                   sem correspondência na despesa (via modDespesaNE).
'===============================================================================
'---------------------------------------------------------------------------
' Mescla o relatório de despesa com a receita já alocada na aba TRATAMENTO.
' Adiciona colunas CHAVE, VALOR FRETE, VALOR COMPLEMENTAR, VALOR OUTROS,
' IMPOSTOS DESPESA, DETALHAMENTO VALORES e todas as demais colunas do
' relatório de despesa (valores únicos concatenados com "|").
' Colunas da despesa que já existem na TRATAMENTO recebem prefixo "DESP_".
'---------------------------------------------------------------------------
Public Function MesclarDespesa(ByVal caminhoDespesa As String) As clsResultado
    Dim resultado As New clsResultado
    Dim tempoInicio As Double
    Dim estadoScreenUpdating As Boolean
    Dim estadoEnableEvents As Boolean
    Dim estadoCalculation As XlCalculation
    Dim descricaoErro As String
    tempoInicio = Timer
    On Error GoTo TrataErro
    CapturarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    Dim wbDespesa As Workbook
    Dim wsDespesa As Worksheet
    Dim wsTratamento As Worksheet
    Dim dimDespesa As Object
    Dim dimReceita As Object
    Dim arrDespesa As Variant
    Dim arrReceita As Variant
    ' === 1. ABRIR ARQUIVO DE DESPESA ===
    RegistrarInfo "modMesclagem", "Abrindo arquivo de despesa: " & caminhoDespesa
    Application.ScreenUpdating = False
    Set wbDespesa = Workbooks.Open(caminhoDespesa, ReadOnly:=True)
    Set wsDespesa = wbDespesa.Sheets(1)
    ' === 2. CAPTURAR DIMENSÕES DA DESPESA ===
    Set dimDespesa = CapturarDimensoesRelatorio(wsDespesa)
    If Not dimDespesa("temDados") Then
        resultado.Sucesso = True
        resultado.RegistrosProcessados = 0
        resultado.Mensagem = "Arquivo de despesa sem dados."
        resultado.TempoExecucao = Timer - tempoInicio
        wbDespesa.Close SaveChanges:=False
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        Set MesclarDespesa = resultado
        Exit Function
    End If
    ' === 3. ENCONTRAR COLUNAS NO RELATÓRIO DE DESPESA ===
    Dim colPontoOp As Long, colNumeroDPS As Long
    Dim colValorShipment As Long, colImpostos As Long, colModalidade As Long
    colPontoOp = EncontrarIndiceColuna(dimDespesa("nomesColunas"), COL_PONTO_OPERACAO)
    colNumeroDPS = EncontrarIndiceColuna(dimDespesa("nomesColunas"), COL_NUMERO_DPS)
    colValorShipment = EncontrarIndiceColuna(dimDespesa("nomesColunas"), COL_VALOR_SHIPMENT_BUY_TOTAL)
    colImpostos = EncontrarIndiceColuna(dimDespesa("nomesColunas"), COL_IMPOSTOS)
    colModalidade = EncontrarIndiceColuna(dimDespesa("nomesColunas"), COL_MODALIDADE)
    ' Verificar colunas obrigatórias
    If colPontoOp = 0 Or colNumeroDPS = 0 Or colValorShipment = 0 Or colModalidade = 0 Then
        Dim colFaltando As String
        colFaltando = ""
        If colPontoOp = 0 Then colFaltando = colFaltando & COL_PONTO_OPERACAO & ", "
        If colNumeroDPS = 0 Then colFaltando = colFaltando & COL_NUMERO_DPS & ", "
        If colValorShipment = 0 Then colFaltando = colFaltando & COL_VALOR_SHIPMENT_BUY_TOTAL & ", "
        If colModalidade = 0 Then colFaltando = colFaltando & COL_MODALIDADE & ", "
        resultado.Sucesso = False
        resultado.Mensagem = "Colunas obrigatórias não encontradas na despesa: " & Left(colFaltando, Len(colFaltando) - 2)
        resultado.TempoExecucao = Timer - tempoInicio
        wbDespesa.Close SaveChanges:=False
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        Set MesclarDespesa = resultado
        Exit Function
    End If
    RegistrarInfo "modMesclagem", "Colunas mapeadas: PO=" & colPontoOp & _
                 ", DPS=" & colNumeroDPS & ", Shipment=" & colValorShipment & _
                 ", Impostos=" & colImpostos & ", Modalidade=" & colModalidade
    ' === 4. CARREGAR DADOS DA DESPESA EM ARRAY ===
    arrDespesa = wsDespesa.Range( _
                    wsDespesa.Cells(dimDespesa("linhaHeader"), dimDespesa("primeiraColuna")), _
                    wsDespesa.Cells(dimDespesa("ultimaLinha"), dimDespesa("ultimaColuna"))).value
    ' === 5. FECHAR ARQUIVO DE DESPESA ===
    wbDespesa.Close SaveChanges:=False
    ' === 6. CONSTRUIR DICTIONÁRIO DE SOMAS + VALORES ÚNICOS POR CHAVE ===
    Dim dictSomas As Object
    Set dictSomas = ConstruirSomasDespesa(arrDespesa, dimDespesa("totalRegistros"), _
                                           colPontoOp, colNumeroDPS, _
                                           colValorShipment, colImpostos, colModalidade, _
                                           dimDespesa("totalColunas"))
    RegistrarInfo "modMesclagem", "Dicionário construído: " & dictSomas.Count & " chave(s) única(s)"
    ' === 7. LER RECEITA DA ABA TRATAMENTO ===
    Set wsTratamento = ThisWorkbook.Sheets(SHEET_TRATAMENTO)
    Set dimReceita = CapturarDimensoesRelatorio(wsTratamento)
    If Not dimReceita("temDados") Then
        resultado.Sucesso = False
        resultado.Mensagem = "Aba TRATAMENTO vazia. Importe a receita primeiro."
        resultado.TempoExecucao = Timer - tempoInicio
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        Set MesclarDespesa = resultado
        Exit Function
    End If
    arrReceita = wsTratamento.Range( _
                    wsTratamento.Cells(dimReceita("linhaHeader"), dimReceita("primeiraColuna")), _
                    wsTratamento.Cells(dimReceita("ultimaLinha"), dimReceita("ultimaColuna"))).value
    ' === 8. ENCONTRAR COLUNAS NA RECEITA ===
    Dim colRecPontoOp As Long, colRecNumeroDPS As Long
    colRecPontoOp = EncontrarIndiceColuna(dimReceita("nomesColunas"), COL_PONTO_OPERACAO)
    colRecNumeroDPS = EncontrarIndiceColuna(dimReceita("nomesColunas"), COL_NUMERO_DPS)
    If colRecPontoOp = 0 Or colRecNumeroDPS = 0 Then
        resultado.Sucesso = False
        resultado.Mensagem = "Colunas PONTO DE OPERAÇÃO ou NÚMERO DPS não encontradas na aba TRATAMENTO."
        resultado.TempoExecucao = Timer - tempoInicio
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        Set MesclarDespesa = resultado
        Exit Function
    End If
    ' === 9. ADICIONAR COLUNAS DE RESULTADO NA ABA TRATAMENTO ===
    ' 9a. Colunas padrão de resultado (somadas + detalhamento)
    Dim ultimaColunaAtual As Long
    ultimaColunaAtual = wsTratamento.Cells(dimReceita("linhaHeader"), _
                                            wsTratamento.Columns.Count).End(xlToLeft).Column
    Dim colChave As Long, colValFrete As Long, colValCompl As Long
    Dim colValOutros As Long, colImpostosDesp As Long, colDetalhamento As Long
    colChave = EncontrarOuCriarColuna(wsTratamento, dimReceita("linhaHeader"), COL_CHAVE, ultimaColunaAtual)
    If colChave > ultimaColunaAtual Then ultimaColunaAtual = colChave
    wsTratamento.Cells(dimReceita("linhaHeader"), colChave).Interior.Color = COR_CABECALHO_DESPESA
    colValFrete = EncontrarOuCriarColuna(wsTratamento, dimReceita("linhaHeader"), COL_VALOR_FRETE, ultimaColunaAtual)
    If colValFrete > ultimaColunaAtual Then ultimaColunaAtual = colValFrete
    wsTratamento.Cells(dimReceita("linhaHeader"), colValFrete).Interior.Color = COR_CABECALHO_DESPESA
    colValCompl = EncontrarOuCriarColuna(wsTratamento, dimReceita("linhaHeader"), COL_VALOR_COMPLEMENTAR, ultimaColunaAtual)
    If colValCompl > ultimaColunaAtual Then ultimaColunaAtual = colValCompl
    wsTratamento.Cells(dimReceita("linhaHeader"), colValCompl).Interior.Color = COR_CABECALHO_DESPESA
    colValOutros = EncontrarOuCriarColuna(wsTratamento, dimReceita("linhaHeader"), COL_VALOR_OUTROS, ultimaColunaAtual)
    If colValOutros > ultimaColunaAtual Then ultimaColunaAtual = colValOutros
    wsTratamento.Cells(dimReceita("linhaHeader"), colValOutros).Interior.Color = COR_CABECALHO_DESPESA
    colImpostosDesp = EncontrarOuCriarColuna(wsTratamento, dimReceita("linhaHeader"), COL_IMPOSTOS_DESPESA, ultimaColunaAtual)
    If colImpostosDesp > ultimaColunaAtual Then ultimaColunaAtual = colImpostosDesp
    wsTratamento.Cells(dimReceita("linhaHeader"), colImpostosDesp).Interior.Color = COR_CABECALHO_DESPESA
    colDetalhamento = EncontrarOuCriarColuna(wsTratamento, dimReceita("linhaHeader"), COL_DETALHAMENTO, ultimaColunaAtual)
    If colDetalhamento > ultimaColunaAtual Then ultimaColunaAtual = colDetalhamento
    wsTratamento.Cells(dimReceita("linhaHeader"), colDetalhamento).Interior.Color = COR_CABECALHO_DESPESA
    ' 9b. Mapear TODAS as colunas da despesa (não-especiais) para colunas na TRATAMENTO
    '     Se o nome já existe na TRATAMENTO, cria com prefixo "DESP_"
    Dim colunasEspeciais As Object
    Set colunasEspeciais = CreateObject("Scripting.Dictionary")
    colunasEspeciais.Add colPontoOp, True
    colunasEspeciais.Add colNumeroDPS, True
    colunasEspeciais.Add colValorShipment, True
    If colImpostos > 0 Then colunasEspeciais.Add colImpostos, True
    colunasEspeciais.Add colModalidade, True
    Dim colMapDespesa As Object
    Set colMapDespesa = CreateObject("Scripting.Dictionary")
    Dim nomesColunasDespesa As Variant
    nomesColunasDespesa = dimDespesa("nomesColunas")
    Dim j As Long
    For j = 1 To dimDespesa("totalColunas")
        If Not colunasEspeciais.Exists(j) Then
            Dim nomeColDespesa As String
            nomeColDespesa = Trim(CStr(nomesColunasDespesa(j)))
            ' Ignorar colunas com nome vazio
            If Len(nomeColDespesa) > 0 Then
                ' Verificar se a coluna já existe na TRATAMENTO
                Dim colExistente As Long
                colExistente = 0
                Dim k As Long
                For k = 1 To ultimaColunaAtual
                    If UCase(Trim(CStr(wsTratamento.Cells(dimReceita("linhaHeader"), k).value))) = UCase(nomeColDespesa) Then
                        colExistente = k
                        Exit For
                    End If
                Next k
                ' Se já existe, criar com prefixo
                Dim nomeFinalColuna As String
                If colExistente > 0 Then
                    nomeFinalColuna = PREFIX_DESPESA & nomeColDespesa
                    ' Verificar se a coluna com prefixo já existe (reexecução)
                    colExistente = 0
                    For k = 1 To ultimaColunaAtual
                        If UCase(Trim(CStr(wsTratamento.Cells(dimReceita("linhaHeader"), k).value))) = UCase(nomeFinalColuna) Then
                            colExistente = k
                            Exit For
                        End If
                    Next k
                Else
                    nomeFinalColuna = nomeColDespesa
                End If
                ' Se ainda não existe (nem original nem com prefixo), criar
                If colExistente = 0 Then
                    ultimaColunaAtual = ultimaColunaAtual + 1
                    wsTratamento.Cells(dimReceita("linhaHeader"), ultimaColunaAtual).value = nomeFinalColuna
                    wsTratamento.Cells(dimReceita("linhaHeader"), ultimaColunaAtual).Interior.Color = COR_CABECALHO_DESPESA
                    colMapDespesa.Add j, ultimaColunaAtual
                Else
                    ' Coluna com prefixo já existe (reexecução) — reutilizar
                    wsTratamento.Cells(dimReceita("linhaHeader"), colExistente).Interior.Color = COR_CABECALHO_DESPESA
                    colMapDespesa.Add j, colExistente
                End If
            End If
        End If
    Next j
    RegistrarInfo "modMesclagem", "Colunas de despesa mapeadas: " & colMapDespesa.Count
    ' === 10. MESCLAR: para cada linha de receita, gerar chave, buscar somas e valores ===
    Dim i As Long
    Dim pontoOpOtm As String, pontoOpBase As String
    Dim numeroDPS As String, chave As String
    Dim dictInterno As Object
    Dim totalMesclado As Long
    Dim chavesNaoEncontradas As Long
    totalMesclado = 0
    chavesNaoEncontradas = 0
    For i = 2 To dimReceita("totalRegistros") + 1
        ' Ler valores da receita
        pontoOpOtm = Trim(CStr(arrReceita(i, colRecPontoOp)))
        numeroDPS = Trim(CStr(arrReceita(i, colRecNumeroDPS)))
        ' Consultar TabPontoOperação para obter o ponto de operação base
        pontoOpBase = ConsultarTabelaTAB(TBL_PONTO_OPERACAO, TBL_PO_COL_REF, _
                                          TBL_PO_COL_RETORNO, pontoOpOtm)
        ' Gerar chave composta
        chave = GerarChave(pontoOpBase, numeroDPS)
        ' Gravar chave na aba TRATAMENTO
        wsTratamento.Cells(dimReceita("linhaHeader") + i - 1, colChave).value = chave
        ' Buscar somas no dicionário
        If dictSomas.Exists(chave) Then
            Set dictInterno = dictSomas(chave)
            ' Gravar valores somados
            wsTratamento.Cells(dimReceita("linhaHeader") + i - 1, colValFrete).value = dictInterno("frete")
            wsTratamento.Cells(dimReceita("linhaHeader") + i - 1, colValCompl).value = dictInterno("complementar")
            wsTratamento.Cells(dimReceita("linhaHeader") + i - 1, colValOutros).value = dictInterno("outros")
            wsTratamento.Cells(dimReceita("linhaHeader") + i - 1, colImpostosDesp).value = dictInterno("impostos")
            ' Gravar detalhamento concatenado: FRETE-1000,00|COMPLEMENTAR-500,00|OUTROS-20,00
            wsTratamento.Cells(dimReceita("linhaHeader") + i - 1, colDetalhamento).value = _
                ConstruirDetalhamento(dictInterno("frete"), dictInterno("complementar"), dictInterno("outros"))
            ' Gravar valores concatenados das demais colunas de despesa
            Dim chaveCol As Variant
            For Each chaveCol In colMapDespesa.Keys
                Dim colDespesaIdx As Long
                colDespesaIdx = CLng(chaveCol)
                Dim strColKey As String
                strColKey = "col_" & colDespesaIdx
                If dictInterno.Exists(strColKey) Then
                    ' Tem valores únicos — concatenar com "|"
                    Dim dictUnicos As Object
                    Set dictUnicos = dictInterno(strColKey)
                    wsTratamento.Cells(dimReceita("linhaHeader") + i - 1, _
                                       colMapDespesa(colDespesaIdx)).value = Join(dictUnicos.Items, "|")
                Else
                    ' Sem valores para esta coluna nesta chave
                    wsTratamento.Cells(dimReceita("linhaHeader") + i - 1, _
                                       colMapDespesa(colDespesaIdx)).value = ""
                End If
            Next chaveCol
            totalMesclado = totalMesclado + 1
        Else
            ' Chave da receita não encontrada na despesa — zeros e vazios
            wsTratamento.Cells(dimReceita("linhaHeader") + i - 1, colValFrete).value = 0
            wsTratamento.Cells(dimReceita("linhaHeader") + i - 1, colValCompl).value = 0
            wsTratamento.Cells(dimReceita("linhaHeader") + i - 1, colValOutros).value = 0
            wsTratamento.Cells(dimReceita("linhaHeader") + i - 1, colImpostosDesp).value = 0
            wsTratamento.Cells(dimReceita("linhaHeader") + i - 1, colDetalhamento).value = ""
            chavesNaoEncontradas = chavesNaoEncontradas + 1
            LogSemCorrespondencia "modMesclagem", chave, "Despesa não encontrada para esta chave de receita"
        End If
    Next i
    ' === 10b. REGISTRAR PROCESSOS DA RECEITA SEM CORRESPONDÊNCIA NA DESPESA (DESPESA NE) ===
    ' Carrega DADOS TRATADOS (base final) para resolver o DPS origem dos complementares
    Dim wsDT As Worksheet, dimDT As Object, arrDT As Variant
    Set wsDT = ThisWorkbook.Sheets(SHEET_DADOS_TRATADOS)
    Set dimDT = CapturarDimensoesRelatorio(wsDT)
    If Not dimDT("temDados") Then
        RegistrarAviso "modMesclagem", _
            "DADOS TRATADOS sem dados — complementares não poderão ser resolvidos e irão para DESPESA NE."
    End If
    arrDT = wsDT.Range( _
        wsDT.Cells(dimDT("linhaHeader"), dimDT("primeiraColuna")), _
        wsDT.Cells(dimDT("ultimaLinha"), dimDT("ultimaColuna"))).value

    Dim resNE As clsResultado
    Set resNE = modDespesaNE.RegistrarProcessosSemDespesa( _
        arrReceita, dimReceita("nomesColunas"), _
        arrDespesa, dimDespesa("nomesColunas"), _
        arrDT, dimDT("nomesColunas"), _
        ConstruirDictPontoBase())
    If Not resNE.Sucesso Then
        RegistrarEvento nlErro, catMesclagem, "modMesclagem", _
                        "Falha ao registrar processos sem despesa: " & resNE.Mensagem
    ElseIf resNE.RegistrosProcessados > 0 Then
        RegistrarAviso "modMesclagem", _
            resNE.RegistrosProcessados & " processo(s) da receita sem correspondência na despesa (DESPESA NE)."
    End If
    ' === 11. RESULTADO ===
    RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    resultado.Sucesso = True
    resultado.RegistrosProcessados = totalMesclado
    resultado.Mensagem = "Mesclagem concluída: " & totalMesclado & " registro(s) mesclado(s)." & _
                         IIf(chavesNaoEncontradas > 0, " " & chavesNaoEncontradas & " sem correspondência.", "")
    resultado.TempoExecucao = Timer - tempoInicio
    RegistrarEvento nlInfo, catMesclagem, "modMesclagem", _
                    "Mesclagem concluída", "", totalMesclado, _
                    "Sem correspondência: " & chavesNaoEncontradas & _
                    " | Colunas extras: " & colMapDespesa.Count & _
                    " | Tempo: " & Format(resultado.TempoExecucao, "0.00") & "s"
    Set MesclarDespesa = resultado
    Exit Function
TrataErro:
    descricaoErro = Err.Description
    On Error Resume Next
    If Not wbDespesa Is Nothing Then wbDespesa.Close SaveChanges:=False
    On Error GoTo 0
    RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    resultado.Sucesso = False
    resultado.RegistrosProcessados = 0
    resultado.Mensagem = "Erro ao mesclar despesa: " & descricaoErro
    resultado.TempoExecucao = Timer - tempoInicio
    RegistrarEvento nlErro, catMesclagem, "modMesclagem", _
                    "Erro ao mesclar despesa: " & descricaoErro, caminhoDespesa
    Set MesclarDespesa = resultado
End Function

'---------------------------------------------------------------------------
' Mescla o relatório HP - Manifesto Carga na aba TRATAMENTO.
' Reaproveita a mesma lógica da despesa: identifica o cabeçalho real,
' normaliza o ponto de operação via TabPontoOperação e usa NUMERO CTRC para
' gerar a chave de merge. Os valores únicos de cada coluna são concatenados
' com "|" para evitar duplicidade e manter o histórico do relatório.
'---------------------------------------------------------------------------
Public Function MesclarHPManifesto(ByVal caminhoHP As String) As clsResultado
    Dim resultado As New clsResultado
    Dim tempoInicio As Double
    Dim wbHP As Workbook
    Dim wsHP As Worksheet
    Dim wsTratamento As Worksheet
    Dim dimHP As Object
    Dim dimTratamento As Object
    Dim arrHP As Variant
    Dim arrTratamento As Variant
    Dim colPontoOp As Long, colNumeroCTRC As Long
    Dim colRecPontoOp As Long, colRecNumeroDPS As Long
    Dim i As Long, j As Long, k As Long
    Dim linhaHeaderHP As Long, ultimaColunaAtual As Long
    Dim chave As String, valorColuna As String
    Dim pontoOperacaoOTM As String, numeroCTRC As String
    Dim nomeColuna As String, nomeFinalColuna As String
    Dim colExistente As Long
    Dim dictValores As Object, dictInterno As Object, dictUnicos As Object
    Dim colMapHP As Object
    Dim chaveCol As Variant, colIdx As Long
    Dim totalMesclado As Long
    Dim totalColunasHP As Long
    Dim linhasSemCorresp As Long
    Dim estadoScreenUpdating As Boolean
    Dim estadoEnableEvents As Boolean
    Dim estadoCalculation As XlCalculation
    Dim descricaoErro As String

    tempoInicio = Timer
    totalMesclado = 0
    linhasSemCorresp = 0
    On Error GoTo TrataErro
    CapturarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation

    RegistrarInfo "modMesclagem", "Abrindo relatório HP - Manifesto Carga: " & caminhoHP
    Application.ScreenUpdating = False
    Set wbHP = Workbooks.Open(caminhoHP, ReadOnly:=True)
    Set wsHP = wbHP.Sheets(1)
    Set dimHP = CapturarDimensoesRelatorio(wsHP)

    If Not dimHP("temDados") Then
        resultado.Sucesso = True
        resultado.RegistrosProcessados = 0
        resultado.Mensagem = "Arquivo HP - Manifesto Carga sem dados."
        resultado.TempoExecucao = Timer - tempoInicio
        wbHP.Close SaveChanges:=False
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        Set MesclarHPManifesto = resultado
        Exit Function
    End If

    colPontoOp = EncontrarIndiceColuna(dimHP("nomesColunas"), COL_PONTO_OPERACAO)
    colNumeroCTRC = EncontrarIndiceColuna(dimHP("nomesColunas"), COL_NUMERO_CTRC)
    If colNumeroCTRC = 0 Then colNumeroCTRC = EncontrarIndiceColuna(dimHP("nomesColunas"), COL_NUMERO_CTRC_ALT)

    If colPontoOp = 0 Or colNumeroCTRC = 0 Then
        resultado.Sucesso = False
        resultado.Mensagem = "Colunas obrigatórias não encontradas no HP Manifesto: PONTO DE OPERAÇÃO e NUMERO CTRC."
        resultado.TempoExecucao = Timer - tempoInicio
        wbHP.Close SaveChanges:=False
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        Set MesclarHPManifesto = resultado
        Exit Function
    End If

    RegistrarInfo "modMesclagem", "Colunas HP mapeadas: PO=" & colPontoOp & ", CTRC=" & colNumeroCTRC
    arrHP = wsHP.Range( _
            wsHP.Cells(dimHP("linhaHeader"), dimHP("primeiraColuna")), _
            wsHP.Cells(dimHP("ultimaLinha"), dimHP("ultimaColuna"))).value
    wbHP.Close SaveChanges:=False

    Set dictValores = CreateObject("Scripting.Dictionary")
    Set colMapHP = CreateObject("Scripting.Dictionary")
    totalColunasHP = dimHP("totalColunas")

    For i = 2 To dimHP("totalRegistros") + 1
        pontoOperacaoOTM = Trim(CStr(arrHP(i, colPontoOp)))
        numeroCTRC = Trim(CStr(arrHP(i, colNumeroCTRC)))
        chave = GerarChavePontoOperacaoNumero(pontoOperacaoOTM, numeroCTRC)
        If Len(chave) = 0 Or chave = "|" Then GoTo ProximoHP

        If Not dictValores.Exists(chave) Then
            Set dictInterno = CreateObject("Scripting.Dictionary")
            dictValores.Add chave, dictInterno
        Else
            Set dictInterno = dictValores(chave)
        End If

        For j = 1 To totalColunasHP
            If j <> colPontoOp And j <> colNumeroCTRC Then
                valorColuna = Trim(CStr(arrHP(i, j)))
                If Len(valorColuna) > 0 Then
                    Dim chaveColHP As String
                    chaveColHP = "col_" & j
                    If dictInterno.Exists(chaveColHP) Then
                        Set dictUnicos = dictInterno(chaveColHP)
                        If Not dictUnicos.Exists(UCase(valorColuna)) Then
                            dictUnicos.Add UCase(valorColuna), valorColuna
                        End If
                    Else
                        Set dictUnicos = CreateObject("Scripting.Dictionary")
                        dictUnicos.Add UCase(valorColuna), valorColuna
                        dictInterno.Add chaveColHP, dictUnicos
                    End If
                End If
            End If
        Next j
ProximoHP:
    Next i

    Set wsTratamento = ThisWorkbook.Sheets(SHEET_TRATAMENTO)
    Set dimTratamento = CapturarDimensoesRelatorio(wsTratamento)
    If Not dimTratamento("temDados") Then
        resultado.Sucesso = False
        resultado.Mensagem = "Aba TRATAMENTO vazia. Importe a receita antes do HP Manifesto."
        resultado.TempoExecucao = Timer - tempoInicio
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        Set MesclarHPManifesto = resultado
        Exit Function
    End If

    colRecPontoOp = EncontrarIndiceColuna(dimTratamento("nomesColunas"), COL_PONTO_OPERACAO)
    colRecNumeroDPS = EncontrarIndiceColuna(dimTratamento("nomesColunas"), COL_NUMERO_DPS)
    If colRecNumeroDPS = 0 Then colRecNumeroDPS = EncontrarIndiceColuna(dimTratamento("nomesColunas"), COL_NUMERO_CTRC)
    If colRecNumeroDPS = 0 Then colRecNumeroDPS = EncontrarIndiceColuna(dimTratamento("nomesColunas"), COL_NUMERO_CTRC_ALT)

    If colRecPontoOp = 0 Or colRecNumeroDPS = 0 Then
        resultado.Sucesso = False
        resultado.Mensagem = "Colunas obrigatórias não encontradas na aba TRATAMENTO para o merge do HP Manifesto."
        resultado.TempoExecucao = Timer - tempoInicio
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        Set MesclarHPManifesto = resultado
        Exit Function
    End If

    ultimaColunaAtual = wsTratamento.Cells(dimTratamento("linhaHeader"), wsTratamento.Columns.Count).End(xlToLeft).Column
    For j = 1 To totalColunasHP
        If j <> colPontoOp And j <> colNumeroCTRC Then
            nomeColuna = Trim(CStr(dimHP("nomesColunas")(j)))
            If Len(nomeColuna) > 0 Then
                colExistente = 0
                For k = 1 To ultimaColunaAtual
                    If UCase(Trim(CStr(wsTratamento.Cells(dimTratamento("linhaHeader"), k).value))) = UCase(nomeColuna) Then
                        colExistente = k
                        Exit For
                    End If
                Next k

                If UCase(Left(nomeColuna, Len(PREFIX_MANIFESTO))) = UCase(PREFIX_MANIFESTO) Then
                    nomeFinalColuna = nomeColuna
                Else
                    nomeFinalColuna = PREFIX_MANIFESTO & nomeColuna
                End If

                If colExistente > 0 Then
                    colExistente = 0
                    For k = 1 To ultimaColunaAtual
                        If UCase(Trim(CStr(wsTratamento.Cells(dimTratamento("linhaHeader"), k).value))) = UCase(nomeFinalColuna) Then
                            colExistente = k
                            Exit For
                        End If
                    Next k
                End If

                If colExistente = 0 Then
                    ultimaColunaAtual = ultimaColunaAtual + 1
                    wsTratamento.Cells(dimTratamento("linhaHeader"), ultimaColunaAtual).value = nomeFinalColuna
                    wsTratamento.Cells(dimTratamento("linhaHeader"), ultimaColunaAtual).Interior.Color = COR_CABECALHO_MANIFESTO
                    colMapHP.Add j, ultimaColunaAtual
                Else
                    wsTratamento.Cells(dimTratamento("linhaHeader"), colExistente).Interior.Color = COR_CABECALHO_MANIFESTO
                    colMapHP.Add j, colExistente
                End If
            End If
        End If
    Next j

    ' Garantir que o prefixo do manifesto seja preservado, mesmo quando já houver
    ' colunas com esse padrão em execuções repetidas do mesmo arquivo.
    RegistrarInfo "modMesclagem", "Mapeamento manifesto inicializado com prefixo '" & PREFIX_MANIFESTO & "'"

    arrTratamento = wsTratamento.Range( _
                    wsTratamento.Cells(dimTratamento("linhaHeader"), dimTratamento("primeiraColuna")), _
                    wsTratamento.Cells(dimTratamento("ultimaLinha"), dimTratamento("ultimaColuna"))).value

    For i = 2 To UBound(arrTratamento, 1)
        pontoOperacaoOTM = Trim(CStr(arrTratamento(i, colRecPontoOp)))
        numeroCTRC = Trim(CStr(arrTratamento(i, colRecNumeroDPS)))
        chave = GerarChavePontoOperacaoNumero(pontoOperacaoOTM, numeroCTRC)

        If Len(chave) > 0 And chave <> "|" Then
            If dictValores.Exists(chave) Then
                Set dictInterno = dictValores(chave)
                For Each chaveCol In colMapHP.Keys
                    colIdx = CLng(chaveCol)
                    If dictInterno.Exists("col_" & colIdx) Then
                        Set dictUnicos = dictInterno("col_" & colIdx)
                        wsTratamento.Cells(dimTratamento("linhaHeader") + i - 1, colMapHP(colIdx)).value = Join(dictUnicos.Items, "|")
                    Else
                        wsTratamento.Cells(dimTratamento("linhaHeader") + i - 1, colMapHP(colIdx)).value = ""
                    End If
                Next chaveCol
                totalMesclado = totalMesclado + 1
            Else
                linhasSemCorresp = linhasSemCorresp + 1
                RegistrarAviso "modMesclagem", "HP Manifesto sem correspondência na TRATAMENTO", chave, "PONTO/CTRC: " & pontoOperacaoOTM & "/" & numeroCTRC
            End If
        End If
    Next i

    resultado.Sucesso = True
    resultado.RegistrosProcessados = totalMesclado
    resultado.Mensagem = "Mesclagem HP Manifesto concluída: " & totalMesclado & " registro(s) ajustado(s)." & _
                         IIf(linhasSemCorresp > 0, " " & linhasSemCorresp & " sem correspondência.", "")
    resultado.TempoExecucao = Timer - tempoInicio
    RegistrarEvento nlInfo, catMesclagem, "modMesclagem", _
                    "Mesclagem do HP Manifesto concluída", "", totalMesclado, _
                    "Correspondências: " & totalMesclado & " | Sem match: " & linhasSemCorresp & _
                    " | Tempo: " & Format(resultado.TempoExecucao, "0.00") & "s"
    RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    Set MesclarHPManifesto = resultado
    Exit Function

TrataErro:
    descricaoErro = Err.Description
    On Error Resume Next
    If Not wbHP Is Nothing Then wbHP.Close SaveChanges:=False
    On Error GoTo 0
    RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    resultado.Sucesso = False
    resultado.RegistrosProcessados = 0
    resultado.Mensagem = "Erro ao mesclar HP Manifesto: " & descricaoErro
    resultado.TempoExecucao = Timer - tempoInicio
    RegistrarEvento nlErro, catMesclagem, "modMesclagem", _
                    "Erro ao mesclar HP Manifesto: " & descricaoErro, caminhoHP
    Set MesclarHPManifesto = resultado
End Function

'---------------------------------------------------------------------------
' FUNÇÕES PRIVADAS (auxiliares internos de modMesclagem)
'---------------------------------------------------------------------------
' Constrói dicionário de somas por chave a partir do array de despesa.
' Cada chave aponta para um dicionário interno com:
'   frete, complementar, outros, impostos ? Double (somados)
'   col_X ? Dictionary de valores únicos (para concatenação)
Private Function ConstruirSomasDespesa(arrDespesa As Variant, _
                                         totalRegistros As Long, _
                                         colPontoOp As Long, _
                                         colNumeroDPS As Long, _
                                         colValorShipment As Long, _
                                         colImpostos As Long, _
                                         colModalidade As Long, _
                                         totalColunas As Long) As Object
    Dim dictSomas As Object
    Set dictSomas = CreateObject("Scripting.Dictionary")
    ' Colunas especiais (não são concatenadas)
    Dim colunasEspeciais As Object
    Set colunasEspeciais = CreateObject("Scripting.Dictionary")
    colunasEspeciais.Add colPontoOp, True
    colunasEspeciais.Add colNumeroDPS, True
    colunasEspeciais.Add colValorShipment, True
    If colImpostos > 0 Then colunasEspeciais.Add colImpostos, True
    colunasEspeciais.Add colModalidade, True
    Dim i As Long, j As Long
    Dim pontoOpOtm As String, pontoOpBase As String
    Dim numeroDPS As String, chave As String
    Dim valorShipment As Double, valorImpostos As Double
    Dim modalidade As String, classificacao As String
    Dim dictInterno As Object
    Dim valorColuna As String
    Dim chaveCol As String
    Dim dictUnicos As Object
    For i = 2 To totalRegistros + 1
        ' Ler ponto de operação e converter para base
        pontoOpOtm = Trim(CStr(arrDespesa(i, colPontoOp)))
        pontoOpBase = ConsultarTabelaTAB(TBL_PONTO_OPERACAO, TBL_PO_COL_REF, _
                                          TBL_PO_COL_RETORNO, pontoOpOtm)
        numeroDPS = Trim(CStr(arrDespesa(i, colNumeroDPS)))
        chave = GerarChave(pontoOpBase, numeroDPS)
        ' Ler valores monetários
        valorShipment = ParaDouble(arrDespesa(i, colValorShipment))
        valorImpostos = ParaDouble(IIf(colImpostos > 0, arrDespesa(i, colImpostos), 0))
        ' Classificar modalidade
        modalidade = Trim(CStr(arrDespesa(i, colModalidade)))
        classificacao = modUtils.ConsultarTabelaTAB( _
            TBL_TIPO_PROCESSO, TBL_TP_COL_REF, TBL_TP_COL_CLASSIFICACAO, _
            modUtils.ParaString(modalidade))
        ' Determinar categoria da soma
        Dim somaFrete As Double, somaCompl As Double, somaOutros As Double
        somaFrete = 0
        somaCompl = 0
        somaOutros = 0
        If InStr(classificacao, UCase(CLASS_COMPLEMENTAR)) > 0 Then
            somaCompl = valorShipment
        ElseIf InStr(classificacao, UCase(CLASS_TAXA)) > 0 Then
            somaOutros = valorShipment
        ElseIf InStr(classificacao, UCase(CLASS_ADIANTAMENTO_PEDAGIO)) > 0 Then
            somaOutros = valorShipment
        Else
            somaFrete = valorShipment
        End If
        ' Criar ou atualizar entrada no dicionário
        If dictSomas.Exists(chave) Then
            Set dictInterno = dictSomas(chave)
            dictInterno("frete") = dictInterno("frete") + somaFrete
            dictInterno("complementar") = dictInterno("complementar") + somaCompl
            dictInterno("outros") = dictInterno("outros") + somaOutros
            dictInterno("impostos") = dictInterno("impostos") + valorImpostos
        Else
            Set dictInterno = CreateObject("Scripting.Dictionary")
            dictInterno("frete") = somaFrete
            dictInterno("complementar") = somaCompl
            dictInterno("outros") = somaOutros
            dictInterno("impostos") = valorImpostos
            dictSomas.Add chave, dictInterno
        End If
        ' Coletar valores únicos das colunas não-especiais
        For j = 1 To totalColunas
            If Not colunasEspeciais.Exists(j) Then
                chaveCol = "col_" & j
                valorColuna = Trim(CStr(arrDespesa(i, j)))
                If Len(valorColuna) > 0 Then
                    If dictInterno.Exists(chaveCol) Then
                        Set dictUnicos = dictInterno(chaveCol)
                        If Not dictUnicos.Exists(UCase(valorColuna)) Then
                            dictUnicos.Add UCase(valorColuna), valorColuna
                        End If
                    Else
                        Set dictUnicos = CreateObject("Scripting.Dictionary")
                        dictUnicos.Add UCase(valorColuna), valorColuna
                        dictInterno.Add chaveCol, dictUnicos
                    End If
                End If
            End If
        Next j
    Next i
    Set ConstruirSomasDespesa = dictSomas
End Function
' Constrói string de detalhamento no formato:
' FRETE-1000,00|COMPLEMENTAR-500,00|OUTROS-20,00
Private Function ConstruirDetalhamento(ByVal valorFrete As Double, _
                                        ByVal valorComplementar As Double, _
                                        ByVal valorOutros As Double) As String
    Dim partes(1 To 3) As String
    partes(1) = "FRETE-" & Format(valorFrete, "#,##0.00")
    partes(2) = "COMPLEMENTAR-" & Format(valorComplementar, "#,##0.00")
    partes(3) = "OUTROS-" & Format(valorOutros, "#,##0.00")
    ConstruirDetalhamento = Join(partes, "|")
End Function
' Encontra uma coluna pelo nome do cabeçalho; se não existe, cria no final
Private Function EncontrarOuCriarColuna(ws As Worksheet, linhaHeader As Long, _
                                         nomeColuna As String, _
                                         ultimaColunaAtual As Long) As Long
    Dim j As Long
    For j = 1 To ultimaColunaAtual
        If UCase(Trim(CStr(ws.Cells(linhaHeader, j).value))) = UCase(Trim(nomeColuna)) Then
            EncontrarOuCriarColuna = j
            Exit Function
        End If
    Next j
    ' Criar nova coluna
    Dim novaColuna As Long
    novaColuna = ultimaColunaAtual + 1
    ws.Cells(linhaHeader, novaColuna).value = nomeColuna
    EncontrarOuCriarColuna = novaColuna
End Function
' Constrói um dicionário PONTO OTM (UCase) -> PONTO base a partir da TabPontoOperação.
' Usado pelo modDespesaNE para recompor a chave da receita sem consulta linha a linha.
Private Function ConstruirDictPontoBase() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim wsTAB As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim colRef As Long, colRet As Long
    Dim ref As String, ret As String
    Set wsTAB = ThisWorkbook.Sheets(SHEET_TAB)
    On Error Resume Next
    Set tbl = wsTAB.ListObjects(TBL_PONTO_OPERACAO)
    On Error GoTo 0
    If Not tbl Is Nothing Then
        colRef = 0
        colRet = 0
        For i = 1 To tbl.ListColumns.Count
            If UCase(Trim(tbl.ListColumns(i).Name)) = UCase(TBL_PO_COL_REF) Then colRef = i
            If UCase(Trim(tbl.ListColumns(i).Name)) = UCase(TBL_PO_COL_RETORNO) Then colRet = i
        Next i
        If colRef > 0 And colRet > 0 Then
            If Not tbl.DataBodyRange Is Nothing Then
                For i = 1 To tbl.ListRows.Count
                    ref = UCase(Trim(CStr(tbl.DataBodyRange(i, colRef).value)))
                    ret = Trim(CStr(tbl.DataBodyRange(i, colRet).value))
                    If Len(ref) > 0 Then
                        If Not dict.Exists(ref) Then dict.Add ref, ret
                    End If
                Next i
            End If
        End If
    End If
    Set ConstruirDictPontoBase = dict
End Function