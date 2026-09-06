'===============================================================================
' MÓDULO: modTratamento
' RESPONSABILIDADE: Aplicar tratamentos pós-transferência na aba TRANSFERENCIA:
'                   abastecer colunas analíticas via tratamentos, ajustes e
'                   consultas na aba TAB. NÃO depende do modLayout — localiza
'                   as colunas pelo cabeçalho da própria TRANSFERENCIA e usa a
'                   aba TRATAMENTO como fonte auxiliar (SHIPMENT, DESP_SHIPMENT,
'                   RECEBEDOR, EXPEDIDOR) correlacionada por DPS + NOTA.
'                   Complementares: a SHIP BUY é preenchida consultando a base
'                   final (DADOS TRATADOS) pelo SHIP SELL do frete origem.
'===============================================================================
Option Explicit

Public Function TratarDadosTransferencia(Optional ByVal preencherRegiaoSemNFDiario As Boolean = False) As clsResultado
    Dim resultado As New clsResultado
    Dim wsX As Worksheet          ' TRANSFERENCIA (alvo do tratamento)
    Dim wsT As Worksheet          ' TRATAMENTO (fonte auxiliar)
    Dim arrX As Variant           ' matriz da TRANSFERENCIA (linha 1 = cabeçalho)
    Dim nomesCol() As String      ' nomes das colunas da TRANSFERENCIA
    Dim ultLinha As Long, ultCol As Long
    Dim i As Long, j As Long, k As Long, nNotas As Long
    Dim veiculo As String, operacao As String
    Dim inicio As Double
    ' --- Colunas da TRANSFERENCIA ---
    Dim colTipoProcesso As Long, colTipoOperacao As Long, colOBS As Long
    Dim colTransportadora As Long, colPontoOperacao As Long, colDestinatario As Long
    Dim colCE As Long, colModalidade As Long, colModal As Long, colContrato As Long
    Dim colChave As Long, colNumeroDPS As Long, colNotasFiscais As Long
    Dim colValorDPS As Long, colValorImposto As Long, colValorCTESemICMS As Long
    Dim colShipSell As Long, colShipBuy As Long, colTipoDPS As Long
    Dim colCidadeDestino As Long, colUFDestino As Long
    Dim colRegiaoDestino As Long, colRegiao As Long
    Dim nLocalidadesAtualizadas As Long, nRegioesAtualizadas As Long
    Dim nLocalidadesCriticas As Long
    ' --- Fonte auxiliar TRATAMENTO ---
    Dim arrT As Variant
    Dim nomesColT() As String
    Dim ultLinhaT As Long, ultColT As Long
    Dim colTDPS As Long, colTNotas As Long
    Dim colTShipment As Long, colTDespShipment As Long
    Dim colTRecebedor As Long, colTExpedidor As Long
    Dim dictTrat As Object
    Dim notasT() As String
    Dim chaveLinha As String
    Dim v As Variant
    Dim temDict As Boolean
    ' --- Fonte base final DADOS TRATADOS (SHIP BUY de complementares) ---
    Dim wsDT As Worksheet, arrDT As Variant
    Dim nomesColDT() As String
    Dim ultLinhaDT As Long, ultColDT As Long
    Dim colDTShipSell As Long, colDTShipBuy As Long, colDTTipoDPS As Long
    Dim dictShipBuy As Object
    Dim shipSellDT As String, shipBuyDT As String
    Dim nCompPreenchidos As Long, nCompSemOrigem As Long
    ' --- Fonte DADOS AUXILIARES (NF Diário + OTM) ---
    Dim wsAux As Worksheet, arrAux As Variant
    Dim ultimaCelulaAux As Range
    Dim nomesColAux() As String
    Dim ultLinhaAux As Long, ultColAux As Long
    Dim colAuxCte As Long, colAuxNF As Long
    Dim dictAuxChave As Object, dictAuxNF As Object
    Dim mapeamentoAux As Collection
    Dim colAuxOrigem() As Long, colAuxDestino() As Long
    Dim auxEhDataRecebimento() As Boolean
    Dim idxAux As Long, idxDestino As Long, chaveAux As String, notaAux As String
    Dim campoAux As String, campoTransferencia As String

    inicio = Timer
    On Error GoTo Falha

    ' --- 1. Abas alvo e fonte ---
    Set wsX = ThisWorkbook.Worksheets(SHEET_TRANSFERENCIA)
    Set wsT = ThisWorkbook.Worksheets(SHEET_TRATAMENTO)

    ' --- 2. Dimensões da TRANSFERENCIA ---
    ultLinha = wsX.Cells(wsX.Rows.Count, 8).End(xlUp).Row
    If ultLinha < 2 Then
        resultado.Sucesso = False
        resultado.Mensagem = "Aba TRANSFERENCIA sem dados para tratar."
        resultado.TempoExecucao = Timer - inicio
        Set TratarDadosTransferencia = resultado
        Exit Function
    End If
    ultCol = wsX.Cells(1, wsX.Columns.Count).End(xlToLeft).Column

    ' --- 3. Carrega matriz e nomes de colunas da TRANSFERENCIA ---
    arrX = wsX.Range(wsX.Cells(1, 1), wsX.Cells(ultLinha, ultCol)).value
    ReDim nomesCol(1 To ultCol)
    For i = 1 To ultCol
        nomesCol(i) = CStr(arrX(1, i))
    Next i

    ' --- 4. Localiza colunas pelo cabeçalho da TRANSFERENCIA ---
    colChave = modUtils.EncontrarIndiceColuna(nomesCol, "CHAVE")
    colTipoProcesso = modUtils.EncontrarIndiceColuna(nomesCol, "Tipo de Processo")
    colTipoOperacao = modUtils.EncontrarIndiceColuna(nomesCol, "Tipo de Operação")
    colOBS = modUtils.EncontrarIndiceColuna(nomesCol, "OBS")
    colTransportadora = modUtils.EncontrarIndiceColuna(nomesCol, "TRANSPORTADORA")
    colPontoOperacao = modUtils.EncontrarIndiceColuna(nomesCol, "PONTO DE OPERAÇÃO")
    colDestinatario = modUtils.EncontrarIndiceColuna(nomesCol, "DESTINATÁRIO")
    colCE = modUtils.EncontrarIndiceColuna(nomesCol, "C.E.")
    colModalidade = modUtils.EncontrarIndiceColuna(nomesCol, "MODALIDADE")
    colModal = modUtils.EncontrarIndiceColuna(nomesCol, "Modal")
    colContrato = modUtils.EncontrarIndiceColuna(nomesCol, "CONTRATO")
    colNumeroDPS = modUtils.EncontrarIndiceColuna(nomesCol, "NÚMERO DPS")
    colNotasFiscais = modUtils.EncontrarIndiceColuna(nomesCol, "NOTAS FISCAIS")
    colValorDPS = modUtils.EncontrarIndiceColuna(nomesCol, "VALOR DPS")
    colValorImposto = modUtils.EncontrarIndiceColuna(nomesCol, "VALOR IMPOSTO")
    colValorCTESemICMS = modUtils.EncontrarIndiceColuna(nomesCol, "VALOR CTE SEM ICMS")
    colShipSell = modUtils.EncontrarIndiceColuna(nomesCol, "SHIP SELL")
    colShipBuy = modUtils.EncontrarIndiceColuna(nomesCol, "SHIP BUY")
    colTipoDPS = modUtils.EncontrarIndiceColuna(nomesCol, "TIPO DO DPS")
    colCidadeDestino = modUtils.EncontrarIndiceColuna(nomesCol, COL_CIDADE_DESTINO)
    colUFDestino = modUtils.EncontrarIndiceColuna(nomesCol, COL_UF_DESTINO)
    colRegiaoDestino = modUtils.EncontrarIndiceColuna(nomesCol, COL_REGIAO_DESTINO)
    colRegiao = modUtils.EncontrarIndiceColuna(nomesCol, COL_REGIAO)

    ' --- 4b. Carrega TRATAMENTO e monta dicionário de correlação DPS + NOTA ---
    Set dictTrat = Nothing
    temDict = False
    ultLinhaT = wsT.Cells(wsT.Rows.Count, 1).End(xlUp).Row
    If ultLinhaT >= 2 Then
        ultColT = wsT.Cells(1, wsT.Columns.Count).End(xlToLeft).Column
        arrT = wsT.Range(wsT.Cells(1, 1), wsT.Cells(ultLinhaT, ultColT)).value
        ReDim nomesColT(1 To ultColT)
        For j = 1 To ultColT
            nomesColT(j) = CStr(arrT(1, j))
        Next j
        colTDPS = modUtils.EncontrarIndiceColuna(nomesColT, "NÚMERO DPS")
        colTNotas = modUtils.EncontrarIndiceColuna(nomesColT, "NOTAS FISCAIS")
        colTShipment = modUtils.EncontrarIndiceColuna(nomesColT, "SHIPMENT")
        colTDespShipment = modUtils.EncontrarIndiceColuna(nomesColT, "DESP_SHIPMENT")
        colTRecebedor = modUtils.EncontrarIndiceColuna(nomesColT, "RECEBEDOR")
        colTExpedidor = modUtils.EncontrarIndiceColuna(nomesColT, "EXPEDIDOR")
        Set dictTrat = CreateObject("Scripting.Dictionary")
        For j = 2 To UBound(arrT, 1)
            If colTNotas > 0 Then
                notasT = Split(LerColuna(arrT, j, colTNotas), SEP_NOTAS_FISCAIS)
                nNotas = UBound(notasT) + 1
            Else
                ReDim notasT(0 To 0)
                notasT(0) = ""
                nNotas = 1
            End If
            For k = 1 To nNotas
                chaveLinha = UCase(LerColuna(arrT, j, colTDPS) & "|" & Trim(modUtils.ParaString(notasT(k - 1))))
                If Len(chaveLinha) > 0 Then
                    If Not dictTrat.Exists(chaveLinha) Then
                        dictTrat.Add chaveLinha, Array( _
                            LerColuna(arrT, j, colTShipment), _
                            LerColuna(arrT, j, colTDespShipment), _
                            LerColuna(arrT, j, colTRecebedor), _
                            LerColuna(arrT, j, colTExpedidor))
                    End If
                End If
            Next k
        Next j
        temDict = True
    End If

    ' --- 4c. Carrega DADOS TRATADOS e monta dicionário SHIP SELL -> SHIP BUY (fretes origem) ---
    Set dictShipBuy = CreateObject("Scripting.Dictionary")
    Set wsDT = ThisWorkbook.Worksheets(SHEET_DADOS_TRATADOS)
    ultLinhaDT = wsDT.Cells(wsDT.Rows.Count, 1).End(xlUp).Row
    If ultLinhaDT >= 2 Then
        ultColDT = wsDT.Cells(1, wsDT.Columns.Count).End(xlToLeft).Column
        arrDT = wsDT.Range(wsDT.Cells(1, 1), wsDT.Cells(ultLinhaDT, ultColDT)).value
        ReDim nomesColDT(1 To ultColDT)
        For j = 1 To ultColDT
            nomesColDT(j) = CStr(arrDT(1, j))
        Next j
        colDTShipSell = modUtils.EncontrarIndiceColuna(nomesColDT, "SHIP SELL")
        colDTShipBuy = modUtils.EncontrarIndiceColuna(nomesColDT, "SHIP BUY")
        colDTTipoDPS = modUtils.EncontrarIndiceColuna(nomesColDT, "TIPO DO DPS")
        For j = 2 To UBound(arrDT, 1)
            ' Apenas linhas NÃO complementares (fretes origem)
            If InStr(1, modUtils.ParaString(arrDT(j, colDTTipoDPS)), KW_COMPLEMENTAR, vbTextCompare) = 0 Then
                shipSellDT = UCase(Trim(modUtils.ParaString(arrDT(j, colDTShipSell))))
                shipBuyDT = modUtils.ParaString(arrDT(j, colDTShipBuy))
                If Len(shipSellDT) > 0 And Len(shipBuyDT) > 0 Then
                    If Not dictShipBuy.Exists(shipSellDT) Then dictShipBuy.Add shipSellDT, shipBuyDT
                End If
            End If
        Next j
    End If

    ' --- 4d. Carrega DADOS AUXILIARES e indexa pelas duas chaves ---
    Set dictAuxChave = CreateObject("Scripting.Dictionary")
    Set dictAuxNF = CreateObject("Scripting.Dictionary")
    Set wsAux = ThisWorkbook.Worksheets(SHEET_DADOS_AUXILIARES)
    Set ultimaCelulaAux = wsAux.Cells.Find(What:="*", _
                                           After:=wsAux.Cells(1, 1), _
                                           LookIn:=xlFormulas, _
                                           SearchOrder:=xlByRows, _
                                           SearchDirection:=xlPrevious)
    If ultimaCelulaAux Is Nothing Then
        ultLinhaAux = 1
        ultColAux = 1
    Else
        ultLinhaAux = ultimaCelulaAux.Row
        ultColAux = wsAux.Cells.Find(What:="*", _
                                     After:=wsAux.Cells(1, 1), _
                                     LookIn:=xlFormulas, _
                                     SearchOrder:=xlByColumns, _
                                     SearchDirection:=xlPrevious).Column
    End If
    ReDim nomesColAux(1 To ultColAux)
    For j = 1 To ultColAux
        nomesColAux(j) = CStr(wsAux.Cells(1, j).value)
    Next j
    If ultLinhaAux >= 2 And ultColAux >= 1 Then
        arrAux = wsAux.Range(wsAux.Cells(1, 1), wsAux.Cells(ultLinhaAux, ultColAux)).value
        colAuxCte = EncontrarColunaDocumentoCTE(nomesColAux)
        colAuxNF = modUtils.EncontrarIndiceColuna(nomesColAux, "NF NOTA FISCAL")
        If colAuxNF > 0 Then
            For j = 2 To UBound(arrAux, 1)
                notaAux = UCase(Trim(modUtils.ParaString(arrAux(j, colAuxNF))))
                If colAuxCte > 0 Then
                    chaveAux = UCase(Trim(modUtils.ParaString(arrAux(j, colAuxCte))) & _
                                    Trim(modUtils.ParaString(arrAux(j, colAuxNF))))
                    If Len(Trim(modUtils.ParaString(arrAux(j, colAuxCte)))) > 0 And _
                       Len(notaAux) > 0 Then
                        If Not dictAuxChave.Exists(chaveAux) Then dictAuxChave.Add chaveAux, j
                    End If
                End If
                ' OTM usa somente a nota. Em uma base mista, o fallback
                ' também cobre linhas OTM sem CTE.
                If (colAuxCte = 0 Or _
                    Len(Trim(modUtils.ParaString(arrAux(j, colAuxCte)))) = 0) And _
                   Len(notaAux) > 0 Then
                    If Not dictAuxNF.Exists(notaAux) Then dictAuxNF.Add notaAux, j
                End If
            Next j
        End If
    End If

    Set mapeamentoAux = modLayout.MapeamentoDadosAuxiliaresTransferencia()
    ReDim colAuxOrigem(1 To mapeamentoAux.Count)
    ReDim colAuxDestino(1 To mapeamentoAux.Count)
    ReDim auxEhDataRecebimento(1 To mapeamentoAux.Count)
    For j = 1 To mapeamentoAux.Count
        campoAux = CStr(mapeamentoAux(j)(0))
        campoTransferencia = CStr(mapeamentoAux(j)(1))
        colAuxOrigem(j) = modUtils.EncontrarIndiceColuna(nomesColAux, campoAux)
        colAuxDestino(j) = modUtils.EncontrarIndiceColuna(nomesCol, campoTransferencia)
        auxEhDataRecebimento(j) = (UCase(Trim(campoAux)) = UCase("NF DATA RECEBIMENTO"))
    Next j

    ' --- 5. Aplica os tratamentos/lookups linha a linha ---
    nCompPreenchidos = 0
    nCompSemOrigem = 0
    For i = 2 To UBound(arrX, 1)
        ' Chave de correlação com a TRATAMENTO (DPS + NOTA)
        chaveLinha = UCase(LerColuna(arrX, i, colNumeroDPS) & "|" & Trim(LerColuna(arrX, i, colNotasFiscais)))
        If temDict And dictTrat.Exists(chaveLinha) Then v = dictTrat(chaveLinha)

        ' Enriquece com DADOS AUXILIARES: chave completa primeiro; NF como fallback.
        idxAux = 0
        chaveAux = UCase(Trim(LerColuna(arrX, i, colChave)))
        notaAux = UCase(Trim(LerColuna(arrX, i, colNotasFiscais)))
        If Len(chaveAux) > 0 And dictAuxChave.Exists(chaveAux) Then
            idxAux = CLng(dictAuxChave(chaveAux))
        ElseIf Len(notaAux) > 0 And dictAuxNF.Exists(notaAux) Then
            idxAux = CLng(dictAuxNF(notaAux))
        End If
        If idxAux > 0 Then
            For j = 1 To mapeamentoAux.Count
                k = colAuxOrigem(j)
                idxDestino = colAuxDestino(j)
                If k > 0 And idxDestino > 0 Then
                        If auxEhDataRecebimento(j) Then
                            arrX(i, idxDestino) = modUtils.ExtrairDataHora(arrAux(idxAux, k))
                        Else
                            arrX(i, idxDestino) = arrAux(idxAux, k)
                        End If
                End If
            Next j
        End If

        ' Tipo de Processo: TabTipoProcesso (ref MODALIDADE -> CLASSIFICAÇÃO)
        If colTipoProcesso > 0 And colModalidade > 0 Then
            arrX(i, colTipoProcesso) = modUtils.ConsultarTabelaTAB( _
                TBL_TIPO_PROCESSO, TBL_TP_COL_REF, TBL_TP_COL_CLASSIFICACAO, _
                modUtils.ParaString(arrX(i, colModalidade)))
        End If

        ' Tipo de Operação: TabTipoProcesso (ref MODALIDADE -> OPERAÇÃO); se "Inbound", usa Tipo
        If colTipoOperacao > 0 And colModalidade > 0 Then
            operacao = modUtils.ConsultarTabelaTAB( _
                TBL_TIPO_PROCESSO, TBL_TP_COL_REF, TBL_TP_COL_OPERACAO, _
                modUtils.ParaString(arrX(i, colModalidade)))
            If StrComp(operacao, "Inbound", vbTextCompare) = 0 Then
                arrX(i, colTipoOperacao) = modUtils.ConsultarTabelaTAB( _
                    TBL_TIPO_PROCESSO, TBL_TP_COL_REF, TBL_TP_COL_TIPO, _
                    modUtils.ParaString(arrX(i, colModalidade)))
            Else
                arrX(i, colTipoOperacao) = operacao
            End If
        End If

        ' PONTO DE OPERAÇÃO: TabPontoOperação (Ponto de operação OTM -> base)
        If colPontoOperacao > 0 Then
            arrX(i, colPontoOperacao) = modUtils.ConsultarTabelaTAB( _
                TBL_PONTO_OPERACAO, TBL_PO_COL_REF, TBL_PO_COL_RETORNO, _
                modUtils.ParaString(arrX(i, colPontoOperacao)))
        End If

        ' OBS: TabTipoProcesso (ref MODALIDADE -> OBS)
        If colOBS > 0 And colModalidade > 0 Then
            arrX(i, colOBS) = modUtils.ConsultarTabelaTAB( _
                TBL_TIPO_PROCESSO, TBL_TP_COL_REF, TBL_TP_COL_OBS, _
                modUtils.ParaString(arrX(i, colModalidade)))
        End If

        ' Transportadora: TabTransportadorOTM (ref Transportador OTM -> base)
        If colTransportadora > 0 Then
            arrX(i, colTransportadora) = modUtils.ConsultarTabelaTAB( _
                TBL_TRANSPORTADOR_OTM, TBL_TR_COL_REF, TBL_TR_COL_RETORNO, _
                modUtils.ParaString(arrX(i, colTransportadora)))
        End If

        ' DESTINATÁRIO: se RECEBEDOR contém "HP" -> EXPEDIDOR, senão RECEBEDOR
        If colDestinatario > 0 And temDict And dictTrat.Exists(chaveLinha) Then
            v = dictTrat(chaveLinha)
            If InStr(1, CStr(v(2)), "HP", vbTextCompare) > 0 Then
                arrX(i, colDestinatario) = CStr(v(3))
            Else
                arrX(i, colDestinatario) = CStr(v(2))
            End If
        End If

        ' CONTRATO: TabTipoProcesso (ref MODALIDADE -> CONTRATO)
        If colContrato > 0 And colModalidade > 0 Then
            arrX(i, colContrato) = modUtils.ConsultarTabelaTAB( _
                TBL_TIPO_PROCESSO, TBL_TP_COL_REF, TBL_TP_COL_CONTRATO, _
                modUtils.ParaString(arrX(i, colModalidade)))
        End If

        ' Modal: extrai veículo de C.E. - MODALIDADE e consulta TabVeiculoOTM
        If colModal > 0 And colCE > 0 And colModalidade > 0 Then
            veiculo = modUtils.ExtrairVeiculoDeCE( _
                modUtils.ParaString(arrX(i, colCE)), _
                modUtils.ParaString(arrX(i, colModalidade)))
            arrX(i, colModal) = modUtils.ConsultarTabelaTAB( _
                TBL_VEICULO_OTM, TBL_VE_COL_REF, TBL_VE_COL_RETORNO, veiculo)
        End If

        ' VALOR CTE SEM ICMS = VALOR DPS - VALOR IMPOSTO
        If colValorCTESemICMS > 0 And colValorDPS > 0 And colValorImposto > 0 Then
            arrX(i, colValorCTESemICMS) = modUtils.ParaDouble(arrX(i, colValorDPS)) - _
                                          modUtils.ParaDouble(arrX(i, colValorImposto))
        End If

        ' SHIP SELL: valor após "." na coluna SHIPMENT (TRATAMENTO)
        If colShipSell > 0 And temDict And dictTrat.Exists(chaveLinha) Then
            v = dictTrat(chaveLinha)
            If InStr(1, CStr(v(0)), ".") > 0 Then
                arrX(i, colShipSell) = Mid(CStr(v(0)), InStr(1, CStr(v(0)), ".") + 1)
            End If
        End If

        ' SHIP BUY: valor após "." na coluna DESP_SHIPMENT (TRATAMENTO)
        If colShipBuy > 0 And temDict And dictTrat.Exists(chaveLinha) Then
            v = dictTrat(chaveLinha)
            If InStr(1, CStr(v(1)), ".") > 0 Then
                arrX(i, colShipBuy) = Mid(CStr(v(1)), InStr(1, CStr(v(1)), ".") + 1)
            End If
        End If

        ' SHIP BUY de COMPLEMENTARES: consulta na base final (DADOS TRATADOS)
        ' A dinâmica do relatório de despesa não traz DESP_SHIPMENT para o
        ' complementar; o SHIP BUY é herdado do frete origem via SHIP SELL.
        If colShipBuy > 0 And colTipoDPS > 0 Then
            If InStr(1, modUtils.ParaString(arrX(i, colTipoDPS)), KW_COMPLEMENTAR, vbTextCompare) > 0 Then
                If Len(Trim(modUtils.ParaString(arrX(i, colShipBuy)))) = 0 Then
                    Dim shipSellComp As String
                    shipSellComp = UCase(Trim(modUtils.ParaString(arrX(i, colShipSell))))
                    If Len(shipSellComp) > 0 And dictShipBuy.Exists(shipSellComp) Then
                        arrX(i, colShipBuy) = CStr(dictShipBuy(shipSellComp))
                        nCompPreenchidos = nCompPreenchidos + 1
                    Else
                        nCompSemOrigem = nCompSemOrigem + 1
                    End If
                End If
            End If
        End If

        ' Sem NF Diário, a Praça do cadastro é a fonte da REGIAO DESTINO
        ' para todos os tipos. Com NF Diário, apenas OST é sobrescrito;
        ' os demais tipos são comparados com o valor já existente.
        If colCidadeDestino > 0 And colUFDestino > 0 And _
           colRegiaoDestino > 0 And colRegiao > 0 Then
            Dim chaveLocalidade As String
            Dim pracaLocalidade As String
            Dim regiaoLocalidade As String
            Dim tipoDPSAtual As String
            chaveLocalidade = UCase(Trim(modUtils.ParaString(arrX(i, colCidadeDestino)))) & "|" & _
                              UCase(Trim(modUtils.ParaString(arrX(i, colUFDestino))))
            pracaLocalidade = modUtils.ConsultarTabelaTAB( _
                TBL_LOCALIDADES_CORTES, TBL_LC_COL_REF, TBL_LC_COL_PRACA, chaveLocalidade)
            regiaoLocalidade = modUtils.ConsultarTabelaTAB( _
                TBL_LOCALIDADES_CORTES, TBL_LC_COL_REF, TBL_LC_COL_REGIAO, chaveLocalidade)
            tipoDPSAtual = modUtils.ParaString(arrX(i, colTipoDPS))

            If Len(Trim(pracaLocalidade)) = 0 Or Len(Trim(regiaoLocalidade)) = 0 Then
                nLocalidadesCriticas = nLocalidadesCriticas + 1
                RegistrarEvento nlCritico, catSistema, "modTratamento", _
                    "Localidade sem cadastro completo na TabLocalidadesCortes", _
                    chaveLocalidade, , "Praça/Região não encontrada ou vazia"
            ElseIf preencherRegiaoSemNFDiario Or _
                   InStr(1, tipoDPSAtual, "OST", vbTextCompare) > 0 Then
                arrX(i, colRegiaoDestino) = pracaLocalidade
                nLocalidadesAtualizadas = nLocalidadesAtualizadas + 1
                If preencherRegiaoSemNFDiario Then
                    RegistrarInfo "modTratamento", _
                        "REGIAO DESTINO preenchida via cadastro sem NF Diário", chaveLocalidade
                Else
                    RegistrarInfo "modTratamento", _
                        "REGIAO DESTINO preenchida diretamente para OST", chaveLocalidade
                End If
            ElseIf StrComp(Trim(modUtils.ParaString(arrX(i, colRegiaoDestino))), _
                           Trim(pracaLocalidade), vbTextCompare) <> 0 Then
                nLocalidadesCriticas = nLocalidadesCriticas + 1
                RegistrarEvento nlCritico, catSistema, "modTratamento", _
                    "REGIAO DESTINO divergente do cadastro", chaveLocalidade, _
                    arrX(i, colRegiaoDestino), "Esperado: " & pracaLocalidade
            End If

            arrX(i, colRegiao) = regiaoLocalidade
            nRegioesAtualizadas = nRegioesAtualizadas + 1
        End If
    Next i

    ' --- 6. Grava de volta na TRANSFERENCIA (CHAVE e Modal como texto) ---
    If colChave > 0 Then wsX.Columns(colChave).NumberFormat = "@"
    If colModal > 0 Then wsX.Columns(colModal).NumberFormat = "@"
    wsX.Range(wsX.Cells(1, 1), wsX.Cells(ultLinha, ultCol)).value = arrX

    ' --- 6b. Log do preenchimento de SHIP BUY de complementares ---
    If nCompPreenchidos > 0 Then
        RegistrarInfo "modTratamento", _
            "SHIP BUY de complementares preenchido via DADOS TRATADOS: " & nCompPreenchidos
    End If
    If nCompSemOrigem > 0 Then
        RegistrarAviso "modTratamento", _
            nCompSemOrigem & " complementar(es) sem SHIP SELL correspondente em DADOS TRATADOS — SHIP BUY não preenchido."
    End If
    If nLocalidadesAtualizadas > 0 Then
        RegistrarInfo "modTratamento", "REGIAO DESTINO preenchida para OST: " & nLocalidadesAtualizadas
    End If
    If nRegioesAtualizadas > 0 Then
        RegistrarInfo "modTratamento", "Região preenchida via TabLocalidadesCortes: " & nRegioesAtualizadas
    End If
    If nLocalidadesCriticas > 0 Then
        RegistrarEvento nlCritico, catSistema, "modTratamento", _
                        "Divergências críticas de localidades: " & nLocalidadesCriticas
    End If

    ' --- 7. Resultado ---
    resultado.Sucesso = True
    resultado.RegistrosProcessados = ultLinha - 1
    resultado.Mensagem = "Tratamento concluído com " & resultado.RegistrosProcessados & " linha(s)."
    resultado.TempoExecucao = Timer - inicio
    Set TratarDadosTransferencia = resultado
    Exit Function

Falha:
    resultado.Sucesso = False
    resultado.Mensagem = "Erro em TratarDadosTransferencia: " & Err.Description
    resultado.TempoExecucao = Timer - inicio
    Set TratarDadosTransferencia = resultado
End Function

' ------------------------------------------------------------------
' HELPERS
' ------------------------------------------------------------------
' Lê com segurança uma célula da matriz; col <= 0 retorna "".
Private Function LerColuna(arr As Variant, linha As Long, col As Long) As String
    If col <= 0 Then
        LerColuna = ""
    Else
        LerColuna = modUtils.ParaString(arr(linha, col))
    End If
End Function

Private Function EncontrarColunaDocumentoCTE(ByVal nomesColunas As Variant) As Long
    Dim aliases As Variant
    Dim i As Long, j As Long
    Dim nomeAtual As String

    aliases = Array("CTE NUM DOCUMENTO", "NUM DOCUMENTO CTE", _
                    "DOCUMENTO CTE", "NUMERO CTE")
    For i = LBound(nomesColunas) To UBound(nomesColunas)
        nomeAtual = UCase(Trim(CStr(nomesColunas(i))))
        For j = LBound(aliases) To UBound(aliases)
            If nomeAtual = UCase(CStr(aliases(j))) Then
                EncontrarColunaDocumentoCTE = i
                Exit Function
            End If
        Next j
    Next i
End Function