'===============================================================================
' MÓDULO: modCorrelacaoDespesa
' RESPONSABILIDADE: Regras de negócio de correlação entre RECEITA e DESPESA.
'   Decide, para cada processo da receita, se ele está presente na despesa
'   (fretes por chave PONTO_BASE|DPS; complementares por DPS origem + modalidade
'   + soma). Produz a classificação que o modDespesaNE consome para manter a aba
'   DESPESA NE (adicionar/ignorar/remover). NÃO toca na aba DESPESA NE.
'===============================================================================
Option Explicit

Public Function ClassificarProcessos( _
    ByVal arrReceita As Variant, _
    ByVal nomesColReceita As Variant, _
    ByVal arrDespesa As Variant, _
    ByVal nomesColDespesa As Variant, _
    ByVal arrDadosTratados As Variant, _
    ByVal nomesColDadosTratados As Variant, _
    ByVal dictPontoBase As Object) As Object

    ' Retorna um dicionário container com:
    '   "novos"     -> dict: chave -> índice da linha na receita (sem correspondência)
    '   "presentes" -> dict: chave -> True (processo presente na despesa)
    '   "origemNaoResolvida" -> Long (complementares sem origem resolvida)
    Dim container As Object
    Set container = CreateObject("Scripting.Dictionary")

    Dim dictFreteDesp As Object, dictCompDesp As Object, dictShipSell As Object
    Dim dictGruposRec As Object, dictPresentesGrupo As Object
    Dim dictNovos As Object, dictPresentesLinha As Object
    Set dictNovos = CreateObject("Scripting.Dictionary")
    Set dictPresentesLinha = CreateObject("Scripting.Dictionary")

    Dim colRecPontoOp As Long, colRecDPS As Long, colRecTipoDPS As Long
    Dim colRecShipment As Long, colRecModalidade As Long
    Dim colRecValorDPS As Long, colRecValorImposto As Long
    Dim i As Long
    Dim nOrigemNaoResolvida As Long
    Dim shipNum As String, dpsOrigem As String, chaveKey As String
    Dim pontoBaseRec As String, modalidadeRec As String
    Dim vG As Variant, vG2 As Variant, vD As Variant, kG As Variant

    ' Índices na receita
    colRecPontoOp = modUtils.EncontrarIndiceColuna(nomesColReceita, COL_PONTO_OPERACAO)
    colRecDPS = modUtils.EncontrarIndiceColuna(nomesColReceita, COL_NUMERO_DPS)
    colRecTipoDPS = modUtils.EncontrarIndiceColuna(nomesColReceita, COL_TIPO_DPS)
    colRecShipment = modUtils.EncontrarIndiceColuna(nomesColReceita, COL_SHIPMENT)
    colRecModalidade = modUtils.EncontrarIndiceColuna(nomesColReceita, COL_MODALIDADE)
    colRecValorDPS = modUtils.EncontrarIndiceColuna(nomesColReceita, COL_VALOR_DPS)
    colRecValorImposto = modUtils.EncontrarIndiceColuna(nomesColReceita, COL_VALOR_IMPOSTO)

    ' Dicionários auxiliares
    Set dictFreteDesp = ConstruirDictFreteDespesa(arrDespesa, nomesColDespesa, dictPontoBase)
    Set dictCompDesp = ConstruirDictCompDespesa(arrDespesa, nomesColDespesa, dictPontoBase)
    Set dictShipSell = ConstruirDictShipSell(arrDadosTratados, nomesColDadosTratados)
    Set dictGruposRec = CreateObject("Scripting.Dictionary")
    Set dictPresentesGrupo = CreateObject("Scripting.Dictionary")

    ' --- PRÉ-PASSAGEM: agrupar complementares por grupo (PONTO_BASE|DPS_ORIGEM|MODALIDADE) ---
    nOrigemNaoResolvida = 0
    For i = 2 To UBound(arrReceita, 1)
        If IsComplementar(arrReceita, i, colRecTipoDPS) Then
            shipNum = ExtrairNumeroAposPonto(modUtils.ParaString(arrReceita(i, colRecShipment)))
            dpsOrigem = ResolverDPSOrigem(shipNum, dictShipSell)
            If Len(dpsOrigem) = 0 Then
                nOrigemNaoResolvida = nOrigemNaoResolvida + 1
            Else
                pontoBaseRec = ObterPontoBase(modUtils.ParaString(arrReceita(i, colRecPontoOp)), dictPontoBase)
                modalidadeRec = UCase(Trim(modUtils.ParaString(arrReceita(i, colRecModalidade))))
                chaveKey = UCase(pontoBaseRec & "|" & dpsOrigem) & "|" & modalidadeRec
                If dictGruposRec.Exists(chaveKey) Then
                    vG = dictGruposRec(chaveKey)
                    vG(0) = vG(0) + 1
                    vG(1) = vG(1) + ValorLiquidoReceita(arrReceita, i, colRecValorDPS, colRecValorImposto)
                    dictGruposRec(chaveKey) = vG
                Else
                    dictGruposRec.Add chaveKey, _
                        Array(1, ValorLiquidoReceita(arrReceita, i, colRecValorDPS, colRecValorImposto))
                End If
            End If
        End If
    Next i

    ' --- AVALIAR GRUPOS: quais complementares estão presentes na despesa ---
    For Each kG In dictGruposRec.Keys
        vG2 = dictGruposRec(kG)   ' Array(n, somaLiquida receita)
        If dictCompDesp.Exists(kG) Then
            vD = dictCompDesp(kG) ' Array(n, somaShipment despesa)
            If vD(0) = 1 Then
                dictPresentesGrupo.Add kG, True
            ElseIf Abs(vG2(1) - vD(1)) <= TOLERANCIA_SOMA Then
                dictPresentesGrupo.Add kG, True
            End If
        End If
    Next kG

    ' --- PASSAGEM PRINCIPAL: classificar cada linha da receita ---
    For i = 2 To UBound(arrReceita, 1)
        pontoBaseRec = ObterPontoBase(modUtils.ParaString(arrReceita(i, colRecPontoOp)), dictPontoBase)
        If IsComplementar(arrReceita, i, colRecTipoDPS) Then
            shipNum = ExtrairNumeroAposPonto(modUtils.ParaString(arrReceita(i, colRecShipment)))
            dpsOrigem = ResolverDPSOrigem(shipNum, dictShipSell)
            If Len(dpsOrigem) = 0 Then
                ' Origem não resolvida -> sem correspondência
                chaveKey = UCase(pontoBaseRec & "|" & modUtils.ParaString(arrReceita(i, colRecDPS)))
                If Len(chaveKey) > 0 Then
                    If Not dictNovos.Exists(chaveKey) Then dictNovos.Add chaveKey, i
                End If
            Else
                modalidadeRec = UCase(Trim(modUtils.ParaString(arrReceita(i, colRecModalidade))))
                chaveKey = UCase(pontoBaseRec & "|" & dpsOrigem) & "|" & modalidadeRec
                If dictPresentesGrupo.Exists(chaveKey) Then
                    If Not dictPresentesLinha.Exists(chaveKey) Then dictPresentesLinha.Add chaveKey, True
                Else
                    If Not dictNovos.Exists(chaveKey) Then dictNovos.Add chaveKey, i
                End If
            End If
        Else
            ' Frete: chave direta PONTO_BASE|DPS
            chaveKey = UCase(pontoBaseRec & "|" & modUtils.ParaString(arrReceita(i, colRecDPS)))
            If Len(chaveKey) > 0 Then
                If dictFreteDesp.Exists(chaveKey) Then
                    If Not dictPresentesLinha.Exists(chaveKey) Then dictPresentesLinha.Add chaveKey, True
                Else
                    If Not dictNovos.Exists(chaveKey) Then dictNovos.Add chaveKey, i
                End If
            End If
        End If
    Next i

    container.Add "novos", dictNovos
    container.Add "presentes", dictPresentesLinha
    container.Add "origemNaoResolvida", nOrigemNaoResolvida
    Set ClassificarProcessos = container
End Function

' ------------------------------------------------------------------
' HELPERS PRIVADOS
' ------------------------------------------------------------------
' Verifica se a linha da receita é um complementar (TIPO DO DPS contém "Complementar")
Private Function IsComplementar(arr As Variant, linha As Long, colTipoDPS As Long) As Boolean
    If colTipoDPS <= 0 Then
        IsComplementar = False
    Else
        IsComplementar = InStr(1, modUtils.ParaString(arr(linha, colTipoDPS)), _
                               KW_COMPLEMENTAR, vbTextCompare) > 0
    End If
End Function

' Valor líquido do complementar na receita: VALOR DPS - VALOR IMPOSTO
Private Function ValorLiquidoReceita(arr As Variant, linha As Long, _
                                     colValorDPS As Long, colValorImposto As Long) As Double
    ValorLiquidoReceita = modUtils.ParaDouble(arr(linha, colValorDPS)) - _
                          modUtils.ParaDouble(arr(linha, colValorImposto))
End Function

' Extrai o número após o último "." (ex.: "219.123456" -> "123456")
Private Function ExtrairNumeroAposPonto(ByVal valor As String) As String
    Dim p As Long
    valor = Trim(valor)
    p = InStrRev(valor, ".")
    If p > 0 And p < Len(valor) Then
        ExtrairNumeroAposPonto = Trim(Mid(valor, p + 1))
    Else
        ExtrairNumeroAposPonto = ""
    End If
End Function

' Resolve o DPS origem a partir do número extraído do SHIPMENT (lookup em DADOS TRATADOS)
Private Function ResolverDPSOrigem(ByVal shipNum As String, ByVal dictShipSell As Object) As String
    If Len(shipNum) > 0 Then
        If dictShipSell.Exists(UCase(shipNum)) Then
            ResolverDPSOrigem = CStr(dictShipSell(UCase(shipNum)))
            Exit Function
        End If
    End If
    ResolverDPSOrigem = ""
End Function

' Chaves de despesa (fretes): UCase(PONTO_BASE|DPS) -> True (sem filtrar complementar)
Private Function ConstruirDictFreteDespesa(arrDespesa As Variant, nomesColDespesa As Variant, _
                                           dictPontoBase As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim colPontoOp As Long, colDPS As Long, i As Long
    Dim pontoBase As String, chave As String
    colPontoOp = modUtils.EncontrarIndiceColuna(nomesColDespesa, COL_PONTO_OPERACAO)
    colDPS = modUtils.EncontrarIndiceColuna(nomesColDespesa, COL_NUMERO_DPS)
    If colPontoOp = 0 Or colDPS = 0 Then
        Set ConstruirDictFreteDespesa = dict
        Exit Function
    End If
    For i = 2 To UBound(arrDespesa, 1)
        pontoBase = ObterPontoBase(modUtils.ParaString(arrDespesa(i, colPontoOp)), dictPontoBase)
        chave = UCase(pontoBase & "|" & modUtils.ParaString(arrDespesa(i, colDPS)))
        If Len(chave) > 0 And Not dict.Exists(chave) Then dict.Add chave, True
    Next i
    Set ConstruirDictFreteDespesa = dict
End Function

' Complementares da despesa: UCase(PONTO_BASE|DPS|MODALIDADE) -> Array(n, somaShipmentBuyTotal)
' A despesa NÃO distingue complementar, então agrega todas as linhas por chave+modalidade.
Private Function ConstruirDictCompDespesa(arrDespesa As Variant, nomesColDespesa As Variant, _
                                          dictPontoBase As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim colPontoOp As Long, colDPS As Long, colModalidade As Long, colShipment As Long
    Dim i As Long, pontoBase As String, chave As String, modalidade As String, v As Variant
    colPontoOp = modUtils.EncontrarIndiceColuna(nomesColDespesa, COL_PONTO_OPERACAO)
    colDPS = modUtils.EncontrarIndiceColuna(nomesColDespesa, COL_NUMERO_DPS)
    colModalidade = modUtils.EncontrarIndiceColuna(nomesColDespesa, COL_MODALIDADE)
    colShipment = modUtils.EncontrarIndiceColuna(nomesColDespesa, COL_VALOR_SHIPMENT_BUY_TOTAL)
    If colPontoOp = 0 Or colDPS = 0 Or colModalidade = 0 Or colShipment = 0 Then
        Set ConstruirDictCompDespesa = dict
        Exit Function
    End If
    For i = 2 To UBound(arrDespesa, 1)
        pontoBase = ObterPontoBase(modUtils.ParaString(arrDespesa(i, colPontoOp)), dictPontoBase)
        modalidade = UCase(Trim(modUtils.ParaString(arrDespesa(i, colModalidade))))
        chave = UCase(pontoBase & "|" & modUtils.ParaString(arrDespesa(i, colDPS))) & "|" & modalidade
        If dict.Exists(chave) Then
            v = dict(chave)
            v(0) = v(0) + 1
            v(1) = v(1) + modUtils.ParaDouble(arrDespesa(i, colShipment))
            dict(chave) = v
        Else
            dict.Add chave, Array(1, modUtils.ParaDouble(arrDespesa(i, colShipment)))
        End If
    Next i
    Set ConstruirDictCompDespesa = dict
End Function

' SHIP SELL -> NÚMERO DPS origem (apenas linhas não-complementares de DADOS TRATADOS)
Private Function ConstruirDictShipSell(arrDT As Variant, nomesColDT As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim colShipSell As Long, colTipoDPS As Long, colDPS As Long, i As Long
    Dim ship As String, tipo As String
    colShipSell = modUtils.EncontrarIndiceColuna(nomesColDT, COL_SHIP_SELL)
    colTipoDPS = modUtils.EncontrarIndiceColuna(nomesColDT, COL_TIPO_DPS)
    colDPS = modUtils.EncontrarIndiceColuna(nomesColDT, COL_NUMERO_DPS)
    If colShipSell = 0 Or colTipoDPS = 0 Or colDPS = 0 Then
        Set ConstruirDictShipSell = dict
        Exit Function
    End If
    For i = 2 To UBound(arrDT, 1)
        tipo = modUtils.ParaString(arrDT(i, colTipoDPS))
        If InStr(1, tipo, KW_COMPLEMENTAR, vbTextCompare) = 0 Then
            ship = UCase(Trim(modUtils.ParaString(arrDT(i, colShipSell))))
            If Len(ship) > 0 And Not dict.Exists(ship) Then
                dict.Add ship, modUtils.ParaString(arrDT(i, colDPS))
            End If
        End If
    Next i
    Set ConstruirDictShipSell = dict
End Function

' Converte o ponto de operação OTM para base via dicionário; se ausente, usa o próprio valor
Private Function ObterPontoBase(ByVal pontoOTM As String, ByVal dictPontoBase As Object) As String
    Dim chave As String
    chave = UCase(Trim(pontoOTM))
    If Len(chave) > 0 And Not dictPontoBase Is Nothing Then
        If dictPontoBase.Exists(chave) Then
            ObterPontoBase = CStr(dictPontoBase(chave))
            Exit Function
        End If
    End If
    ObterPontoBase = pontoOTM
End Function