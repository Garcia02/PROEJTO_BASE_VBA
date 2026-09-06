'===============================================================================
' MÓDULO: modDespesaNE
' RESPONSABILIDADE: Orquestrar a manutenção da aba DESPESA NE usando o resultado
'   da correlação (modCorrelacaoDespesa.ClassificarProcessos). Funcionamento
'   ACUMULATIVO (a aba NÃO é limpa):
'     1. Processos já presentes na aba são IGNORADOS (log);
'     2. Processos novos (sem correspondência na despesa) são ADICIONADOS (log);
'     3. Processos da aba que agora aparecem na despesa são REMOVIDOS (log).
'   Cabeçalho na linha 8, dados a partir de C9. Logs + aviso ao usuário.
'   Chave de identidade na coluna CHAVE:
'     - FRETE: PONTO_BASE|DPS
'     - COMPLEMENTAR: PONTO_BASE|DPS_ORIGEM|MODALIDADE
'===============================================================================
Option Explicit

' Variável pública para o alerta ser exibido ao final do pipeline (modPrincipal)
Public UltimaMensagemNE As String

Public Function RegistrarProcessosSemDespesa( _
    ByVal arrReceita As Variant, _
    ByVal nomesColReceita As Variant, _
    ByVal arrDespesa As Variant, _
    ByVal nomesColDespesa As Variant, _
    ByVal arrDadosTratados As Variant, _
    ByVal nomesColDadosTratados As Variant, _
    ByVal dictPontoBase As Object) As clsResultado

    Dim resultado As New clsResultado
    Dim wsNE As Worksheet
    Dim colunasNE As Variant
    Dim colIdx() As Long
    Dim j As Long, nCol As Long
    Dim nAdicionados As Long, nIgnorados As Long, nRemovidos As Long
    Dim nOrigemNaoResolvida As Long
    Dim inicio As Double

    Dim clsCorr As Object
    Dim dictNovos As Object, dictPresentesLinha As Object
    Dim dictAba As Object
    Dim kNovo As Variant
    Dim jNE As Long, colNotasNE As Long

    inicio = Timer
    On Error GoTo Falha

    ' --- 1. Aba alvo ---
    Set wsNE = ThisWorkbook.Worksheets(SHEET_DESPESA_NE)

    ' --- 2. Colunas da DESPESA NE (modConfig) ---
    colunasNE = ColunasDespesaNE()
    nCol = UBound(colunasNE) + 1

    ' --- 3. Índices das colunas NE na receita ---
    ReDim colIdx(0 To nCol - 1)
    For j = 0 To nCol - 1
        colIdx(j) = modUtils.EncontrarIndiceColuna(nomesColReceita, CStr(colunasNE(j)))
    Next j

    ' --- 4. Cabeçalho e formatos (SEM limpar dados da aba) ---
    EscreverCabecalho wsNE, colunasNE
    colNotasNE = 0
    For jNE = 0 To UBound(colunasNE)
        If StrComp(CStr(colunasNE(jNE)), NE_COL_NOTAS_FISCAIS, vbTextCompare) = 0 Then
            colNotasNE = NE_COLUNA_INICIO + jNE
            Exit For
        End If
    Next jNE
    If colNotasNE > 0 Then wsNE.Columns(colNotasNE).NumberFormat = "@"

    ' --- 5. Carregar CHAVEs já existentes na aba (linha 9 em diante) ---
    Set dictAba = CreateObject("Scripting.Dictionary")
    CarregarChavesAba wsNE, dictAba

    ' --- 6. Correlação (modCorrelacaoDespesa) ---
    Set clsCorr = modCorrelacaoDespesa.ClassificarProcessos( _
        arrReceita, nomesColReceita, _
        arrDespesa, nomesColDespesa, _
        arrDadosTratados, nomesColDadosTratados, _
        dictPontoBase)
    Set dictNovos = clsCorr("novos")
    Set dictPresentesLinha = clsCorr("presentes")
    nOrigemNaoResolvida = CLng(clsCorr("origemNaoResolvida"))

    ' --- 7. FASE ADIÇÃO: processos novos não existentes na aba ---
    nAdicionados = 0
    nIgnorados = 0
    For Each kNovo In dictNovos.Keys
        If dictAba.Exists(kNovo) Then
            nIgnorados = nIgnorados + 1
            RegistrarInfo "modDespesaNE", "Processo já presente na DESPESA NE, ignorado: " & kNovo
        Else
            GravarLinhaNE wsNE, arrReceita, CLng(dictNovos(kNovo)), colIdx, CStr(kNovo)
            dictAba.Add kNovo, 0   ' 0 = recém-adicionada (não é alvo de remoção nesta execução)
            nAdicionados = nAdicionados + 1
            RegistrarInfo "modDespesaNE", "Processo adicionado à DESPESA NE: " & kNovo
        End If
    Next kNovo

    ' --- 8. FASE REMOÇÃO: processos da aba que agora estão presentes na despesa ---
    nRemovidos = 0
    RemoverPresentesNaDespesa wsNE, dictAba, dictPresentesLinha, nRemovidos

    ' --- 9. Aviso de origem não resolvida (relatório comprometido) ---
    If nOrigemNaoResolvida > 0 Then
        RegistrarAviso "modDespesaNE", _
            nOrigemNaoResolvida & " complementar(es) com origem não resolvida (SHIPMENT sem '.' ou sem match em DADOS TRATADOS). " & _
            "O relatório DESPESA NE está comprometido para esses registros."
    End If

    ' --- 10. Log de resumo + guardar mensagem para alerta ao final ---
    ' O popup final só deve aparecer quando houver alterações relevantes para o usuário
    ' (adicionados ou removidos). Ignorados não são considerados gatilho.
    If nAdicionados > 0 Or nRemovidos > 0 Then
        RegistrarInfo "modDespesaNE", _
            "DESPESA NE - adicionados: " & nAdicionados & _
            " | ignorados (já existentes): " & nIgnorados & _
            " | removidos (agora na despesa): " & nRemovidos
        UltimaMensagemNE = "DESPESA NE atualizada:" & vbCrLf & _
                           "  Adicionados: " & nAdicionados & vbCrLf & _
                           "  Ignorados (já existentes): " & nIgnorados & vbCrLf & _
                           "  Removidos (agora presentes na despesa): " & nRemovidos
    Else
        UltimaMensagemNE = ""
    End If

    ' --- 11. Resultado ---
    resultado.Sucesso = True
    resultado.RegistrosProcessados = nAdicionados
    resultado.Mensagem = "DESPESA NE atualizada: " & nAdicionados & " adicionado(s), " & _
                         nIgnorados & " já existente(s), " & nRemovidos & " removido(s)."
    resultado.TempoExecucao = Timer - inicio
    Set RegistrarProcessosSemDespesa = resultado
    Exit Function

Falha:
    resultado.Sucesso = False
    resultado.Mensagem = "Erro em RegistrarProcessosSemDespesa: " & Err.Description
    resultado.TempoExecucao = Timer - inicio
    RegistrarEvento nlErro, catMesclagem, "modDespesaNE", _
                    "Erro em RegistrarProcessosSemDespesa: " & Err.Description
    Set RegistrarProcessosSemDespesa = resultado
End Function

' ------------------------------------------------------------------
' HELPERS PRIVADOS
' ------------------------------------------------------------------
' Carrega as CHAVEs já gravadas na aba (linha 9 em diante) -> dictAba(chave) = linha
Private Sub CarregarChavesAba(wsNE As Worksheet, dictAba As Object)
    Dim ultLinha As Long, r As Long
    Dim chave As String
    ultLinha = wsNE.Cells(wsNE.Rows.Count, NE_COLUNA_INICIO).End(xlUp).Row
    If ultLinha < NE_LINHA_INICIO Then Exit Sub
    For r = NE_LINHA_INICIO To ultLinha
        chave = UCase(Trim(CStr(wsNE.Cells(r, NE_COLUNA_INICIO).value)))
        If Len(chave) > 0 Then
            If Not dictAba.Exists(chave) Then dictAba.Add chave, r
        End If
    Next r
End Sub

' Grava uma linha da receita na DESPESA NE (apenas as colunas configuradas)
Private Sub GravarLinhaNE(wsNE As Worksheet, arrOrigem As Variant, linhaOrigem As Long, _
                          colIdx() As Long, ByVal chave As String)
    Dim ultLinha As Long, nova As Long, j As Long
    ultLinha = wsNE.Cells(wsNE.Rows.Count, NE_COLUNA_INICIO).End(xlUp).Row
    If ultLinha < NE_LINHA_INICIO Then ultLinha = NE_LINHA_INICIO - 1
    nova = ultLinha + 1
    ' Grava as colunas de dados (a partir da 2ª posição, pois a 1ª é a CHAVE)
    For j = 1 To UBound(colIdx)
        If colIdx(j) > 0 Then
            wsNE.Cells(nova, NE_COLUNA_INICIO + j).value = arrOrigem(linhaOrigem, colIdx(j))
        Else
            wsNE.Cells(nova, NE_COLUNA_INICIO + j).value = ""
        End If
    Next j
    ' Grava a CHAVE por último (não é sobrescrita)
    wsNE.Cells(nova, NE_COLUNA_INICIO).value = chave
End Sub

' Remove da aba as linhas cuja CHAVE está presente na despesa (dictPresentesLinha)
Private Sub RemoverPresentesNaDespesa(wsNE As Worksheet, dictAba As Object, _
                                       dictPresentesLinha As Object, ByRef nRemovidos As Long)
    Dim linhasParaDeletar As Object
    Dim k As Variant, r As Long
    Set linhasParaDeletar = CreateObject("Scripting.Dictionary")
    For Each k In dictAba.Keys
        If dictPresentesLinha.Exists(k) Then
            r = CLng(dictAba(k))
            If r > 0 Then
                If Not linhasParaDeletar.Exists(CStr(r)) Then linhasParaDeletar.Add CStr(r), r
            End If
        End If
    Next k
    ' Deletar de baixo para cima para não deslocar as linhas
    Dim chavesOrd As Variant, idx As Long, rDel As Long
    chavesOrd = linhasParaDeletar.Keys
    For idx = UBound(chavesOrd) To LBound(chavesOrd) Step -1
        rDel = CLng(chavesOrd(idx))
        wsNE.Rows(rDel).Delete
        RegistrarInfo "modDespesaNE", "Processo removido da DESPESA NE (agora na despesa): " & k
        nRemovidos = nRemovidos + 1
    Next idx
End Sub

' Escreve o cabeçalho na linha 8 com as colunas configuradas
Private Sub EscreverCabecalho(wsNE As Worksheet, colunasNE As Variant)
    Dim j As Long
    For j = 0 To UBound(colunasNE)
        wsNE.Cells(NE_LINHA_CABECALHO, NE_COLUNA_INICIO + j).value = CStr(colunasNE(j))
    Next j
End Sub