'===============================================================================
' MÓDULO: modImportacao
' RESPONSABILIDADE: Importação de dados — ler e importar relatórios de
'                   receita para o sistema, alocando os dados na aba
'                   TRATAMENTO (planilha livre, sem tabela estruturada).
'===============================================================================

Public UltimaListaShipSellFaltantes As String
Public UltimaQtdShipSellFaltantes As Long

' Importa o relatório de receita e aloca os dados na aba TRATAMENTO
Public Function ImportarReceita(ByVal caminhoReceita As String) As clsResultado
    Dim resultado As New clsResultado
    Dim tempoInicio As Double
    Dim estadoScreenUpdating As Boolean
    Dim estadoEnableEvents As Boolean
    Dim estadoCalculation As XlCalculation
    Dim descricaoErro As String
    tempoInicio = Timer
    On Error GoTo TrataErro
    CapturarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation

    Dim wbReceita As Workbook
    Dim wsReceita As Worksheet
    Dim wsTratamento As Worksheet
    Dim dimensoes As Object
    Dim arrOrigem As Variant
    Dim totalRegistros As Long

    ' === 1. ABRIR ARQUIVO DE RECEITA ===
    RegistrarInfo "modImportacao", "Abrindo arquivo: " & caminhoReceita, caminhoReceita
    Application.ScreenUpdating = False
    Set wbReceita = Workbooks.Open(caminhoReceita, ReadOnly:=True)
    Set wsReceita = wbReceita.Sheets(1)

    ' === 2. REFERENCIAR DESTINO (ABA TRATAMENTO) ===
    Set wsTratamento = ThisWorkbook.Sheets(SHEET_TRATAMENTO)

    ' === 3. CAPTURAR DIMENSÕES DO RELATÓRIO (delegado para modUtils) ===
    Set dimensoes = CapturarDimensoesRelatorio(wsReceita)

    If Not dimensoes("temDados") Then
        RegistrarAviso "modImportacao", "Arquivo sem dados para importar", caminhoReceita
        resultado.Sucesso = True
        resultado.RegistrosProcessados = 0
        resultado.Mensagem = "Arquivo sem dados para importar."
        resultado.TempoExecucao = Timer - tempoInicio
        wbReceita.Close SaveChanges:=False
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        Set ImportarReceita = resultado
        Exit Function
    End If

    RegistrarInfo "modImportacao", "Dimensões detectadas: " & _
                 dimensoes("totalRegistros") & " registro(s), " & _
                 dimensoes("totalColunas") & " coluna(s)"

    ' === 4. CARREGAR DADOS EM ARRAY (cabeçalho + dados) ===
    arrOrigem = wsReceita.Range( _
                    wsReceita.Cells(dimensoes("linhaHeader"), dimensoes("primeiraColuna")), _
                    wsReceita.Cells(dimensoes("ultimaLinha"), dimensoes("ultimaColuna"))).value

    ' === 5. FECHAR ARQUIVO DE RECEITA ===
    wbReceita.Close SaveChanges:=False

    ' === 6. LIMPAR ABA TRATAMENTO ===
    LimparAba wsTratamento

    ' === 7. GRAVAR DADOS NA ABA TRATAMENTO (escrita em lote, a partir de A1) ===
    totalRegistros = dimensoes("totalRegistros")
    wsTratamento.Range(wsTratamento.Cells(1, 1), _
                       wsTratamento.Cells(dimensoes("totalRegistros") + 1, _
                                          dimensoes("totalColunas"))).value = arrOrigem

    ' === 7b. COLORIR CABEÇALHOS DA RECEITA ===
    Dim idxCol As Long
    For idxCol = 1 To dimensoes("totalColunas")
        wsTratamento.Cells(1, idxCol).Interior.Color = COR_CABECALHO_RECEITA
    Next idxCol

    ' === 8. RESULTADO ===
    RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    resultado.Sucesso = True
    resultado.RegistrosProcessados = totalRegistros
    resultado.Mensagem = "Importação concluída: " & totalRegistros & " registro(s)."
    resultado.TempoExecucao = Timer - tempoInicio

    RegistrarEvento nlInfo, catImportacao, "modImportacao", _
                    "Importação concluída com sucesso", "", totalRegistros, _
                    "Tempo: " & Format(resultado.TempoExecucao, "0.00") & "s"

    Set ImportarReceita = resultado
    Exit Function

TrataErro:
    descricaoErro = Err.Description
    On Error Resume Next
    If Not wbReceita Is Nothing Then wbReceita.Close SaveChanges:=False
    On Error GoTo 0
    RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation

    resultado.Sucesso = False
    resultado.RegistrosProcessados = 0
    resultado.Mensagem = "Erro ao importar receita: " & descricaoErro
    resultado.TempoExecucao = Timer - tempoInicio

    RegistrarEvento nlErro, catImportacao, "modImportacao", _
                    "Erro ao importar receita: " & descricaoErro, caminhoReceita

    Set ImportarReceita = resultado
End Function

' Importa qualquer relatório auxiliar opcional para a aba DADOS AUXILIARES.
' A validação por nome é responsabilidade do chamador; esta rotina apenas
' garante que o arquivo seja lido e gravado com a mesma estrutura do restante
' do pipeline (uso de CapturarDimensoesRelatorio e logs integrados).
Public Function ImportarRelatorioAuxiliar(ByVal caminhoArquivo As String, _
                                         Optional ByVal nomeRelatorio As String = "Relatório auxiliar") As clsResultado
    Dim resultado As New clsResultado
    Dim tempoInicio As Double
    Dim wbArquivo As Workbook
    Dim wsArquivo As Worksheet
    Dim wsAux As Worksheet
    Dim dimensoes As Object
    Dim arrOrigem As Variant
    Dim totalRegistros As Long
    Dim idxCol As Long
    Dim estadoScreenUpdating As Boolean
    Dim estadoEnableEvents As Boolean
    Dim estadoCalculation As XlCalculation
    Dim descricaoErro As String

    tempoInicio = Timer
    On Error GoTo TrataErro
    CapturarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation

    RegistrarInfo "modImportacao", "Abrindo relatório auxiliar: " & caminhoArquivo, caminhoArquivo
    Application.ScreenUpdating = False

    Set wbArquivo = Workbooks.Open(caminhoArquivo, ReadOnly:=True)
    Set wsArquivo = wbArquivo.Sheets(1)
    Set wsAux = ObterOuCriarAbaDadosAuxiliares

    If UCase(Trim(nomeRelatorio)) = UCase("Relatório OTM") Then
        wbArquivo.Close SaveChanges:=False
        Set ImportarRelatorioAuxiliar = ImportarRelatorioOTMApended(caminhoArquivo, nomeRelatorio)
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        Exit Function
    End If

    Set dimensoes = CapturarDimensoesRelatorio(wsArquivo)
    If Not dimensoes("temDados") Then
        RegistrarAviso "modImportacao", "Arquivo auxiliar sem dados para importar", caminhoArquivo
        resultado.Sucesso = True
        resultado.RegistrosProcessados = 0
        resultado.Mensagem = nomeRelatorio & " sem dados para importar."
        resultado.TempoExecucao = Timer - tempoInicio
        wbArquivo.Close SaveChanges:=False
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        Set ImportarRelatorioAuxiliar = resultado
        Exit Function
    End If

    RegistrarInfo "modImportacao", "Dimensões do relatório " & nomeRelatorio & " detectadas: " & _
                  dimensoes("totalRegistros") & " registro(s), " & _
                  dimensoes("totalColunas") & " coluna(s)"

    arrOrigem = wsArquivo.Range( _
                    wsArquivo.Cells(dimensoes("linhaHeader"), dimensoes("primeiraColuna")), _
                    wsArquivo.Cells(dimensoes("ultimaLinha"), dimensoes("ultimaColuna"))).value

    wbArquivo.Close SaveChanges:=False

    LimparAba wsAux
    totalRegistros = dimensoes("totalRegistros")
    wsAux.Range(wsAux.Cells(1, 1), wsAux.Cells(totalRegistros + 1, dimensoes("totalColunas"))).value = arrOrigem

    For idxCol = 1 To dimensoes("totalColunas")
        wsAux.Cells(1, idxCol).Interior.Color = COR_CABECALHO_MANIFESTO
    Next idxCol

    RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    resultado.Sucesso = True
    resultado.RegistrosProcessados = totalRegistros
    resultado.Mensagem = nomeRelatorio & " importado para DADOS AUXILIARES: " & totalRegistros & " registro(s)."
    resultado.TempoExecucao = Timer - tempoInicio

    RegistrarEvento nlInfo, catImportacao, "modImportacao", _
                    nomeRelatorio & " importado com sucesso para DADOS AUXILIARES", caminhoArquivo, totalRegistros, _
                    "Tempo: " & Format(resultado.TempoExecucao, "0.00") & "s"

    Set ImportarRelatorioAuxiliar = resultado
    Exit Function

TrataErro:
    descricaoErro = Err.Description
    On Error Resume Next
    If Not wbArquivo Is Nothing Then wbArquivo.Close SaveChanges:=False
    On Error GoTo 0
    RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation

    resultado.Sucesso = False
    resultado.RegistrosProcessados = 0
    resultado.Mensagem = "Erro ao importar " & nomeRelatorio & ": " & descricaoErro
    resultado.TempoExecucao = Timer - tempoInicio

    RegistrarEvento nlErro, catImportacao, "modImportacao", _
                    "Erro ao importar " & nomeRelatorio & ": " & descricaoErro, caminhoArquivo

    Set ImportarRelatorioAuxiliar = resultado
End Function

' Importa um relatório auxiliar opcional para a aba DADOS AUXILIARES.
' A validação do tipo e a inspeção de dimensões acontecem antes da cópia,
' garantindo consistência com o restante do pipeline.
Public Function ImportarNFDiario(ByVal caminhoNFDiario As String) As clsResultado
    Set ImportarNFDiario = ImportarRelatorioAuxiliar(caminhoNFDiario, "HP - NF Diário")
End Function

Private Function ImportarRelatorioOTMApended(ByVal caminhoArquivo As String, ByVal nomeRelatorio As String) As clsResultado
    Dim resultado As New clsResultado
    Dim tempoInicio As Double
    Dim wbArquivo As Workbook
    Dim wsArquivo As Worksheet
    Dim wsAux As Worksheet
    Dim dimensoes As Object
    Dim arrOrigem As Variant
    Dim arrAux As Variant
    Dim nomesOrigem() As String
    Dim nomesDestino() As String
    Dim mapeamento As Collection
    Dim linhaDestino As Long, ultimaLinhaAux As Long
    Dim totalLinhasOrigem As Long, totalColunasDestino As Long
    Dim i As Long, j As Long, k As Long, idxOrigem As Long
    Dim valorDestino As Variant
    Dim nomeDestino As String
    Dim colOrigemDestino() As Long
    Dim origemEhPK() As Boolean
    Dim estadoScreenUpdating As Boolean
    Dim estadoEnableEvents As Boolean
    Dim estadoCalculation As XlCalculation
    Dim descricaoErro As String

    tempoInicio = Timer
    On Error GoTo TrataErro
    CapturarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation

    Set wbArquivo = Workbooks.Open(caminhoArquivo, ReadOnly:=True)
    Set wsArquivo = wbArquivo.Sheets(1)
    Set wsAux = ObterOuCriarAbaDadosAuxiliares

    Set dimensoes = CapturarDimensoesRelatorio(wsArquivo)
    If Not dimensoes("temDados") Then
        resultado.Sucesso = True
        resultado.RegistrosProcessados = 0
        resultado.Mensagem = nomeRelatorio & " sem dados para importar."
        resultado.TempoExecucao = Timer - tempoInicio
        wbArquivo.Close SaveChanges:=False
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        Set ImportarRelatorioOTMApended = resultado
        Exit Function
    End If

    arrOrigem = wsArquivo.Range(wsArquivo.Cells(dimensoes("linhaHeader"), dimensoes("primeiraColuna")), wsArquivo.Cells(dimensoes("ultimaLinha"), dimensoes("ultimaColuna"))).value
    totalLinhasOrigem = UBound(arrOrigem, 1)
    ReDim nomesOrigem(1 To UBound(arrOrigem, 2))
    For j = 1 To UBound(arrOrigem, 2)
        nomesOrigem(j) = CStr(arrOrigem(1, j))
    Next j

    Set mapeamento = modLayout.MapeamentoDadosAuxiliaresOTM()
    totalColunasDestino = wsAux.Cells(1, wsAux.Columns.Count).End(xlToLeft).Column
    If Len(Trim(CStr(wsAux.Cells(1, 1).value))) = 0 Then
        totalColunasDestino = mapeamento.Count
        For j = 1 To totalColunasDestino
            wsAux.Cells(1, j).value = CStr(mapeamento(j)(1))
        Next j
    End If

    ReDim nomesDestino(1 To totalColunasDestino)
    For j = 1 To totalColunasDestino
        nomesDestino(j) = CStr(wsAux.Cells(1, j).value)
    Next j

    ultimaLinhaAux = wsAux.Cells(wsAux.Rows.Count, 1).End(xlUp).Row
    If ultimaLinhaAux < 1 Then ultimaLinhaAux = 1
    ReDim colOrigemDestino(1 To totalColunasDestino)
    ReDim origemEhPK(1 To totalColunasDestino)
    For j = 1 To totalColunasDestino
        For k = 1 To mapeamento.Count
            If StrComp(UCase(Trim(CStr(nomesDestino(j)))), _
                       UCase(Trim(CStr(mapeamento(k)(1)))), vbTextCompare) = 0 Then
                colOrigemDestino(j) = modUtils.EncontrarIndiceColuna( _
                    nomesOrigem, CStr(mapeamento(k)(0)))
                origemEhPK(j) = (UCase(Trim(CStr(mapeamento(k)(0)))) = "PK")
                Exit For
            End If
        Next k
    Next j
    linhaDestino = ultimaLinhaAux + 1
    For i = 2 To totalLinhasOrigem
        For j = 1 To totalColunasDestino
            If Len(Trim(CStr(nomesDestino(j)))) > 0 Then
                nomeDestino = UCase(Trim(CStr(nomesDestino(j))))
                idxOrigem = colOrigemDestino(j)

                If idxOrigem > 0 Then
                    valorDestino = arrOrigem(i, idxOrigem)
                    If origemEhPK(j) Then
                        If nomeDestino = "NF NOTA FISCAL" Then
                            valorDestino = modUtils.ExtrairCampoPK(valorDestino, "NOTA")
                        ElseIf nomeDestino = "NF SERIE" Then
                            valorDestino = modUtils.ExtrairCampoPK(valorDestino, "SERIE")
                        End If
                    End If
                    wsAux.Cells(linhaDestino, j).value = valorDestino
                Else
                    wsAux.Cells(linhaDestino, j).value = ""
                End If
            End If
        Next j
        linhaDestino = linhaDestino + 1
    Next i

    wbArquivo.Close SaveChanges:=False
    RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    resultado.Sucesso = True
    resultado.RegistrosProcessados = totalLinhasOrigem - 1
    resultado.Mensagem = nomeRelatorio & " anexado à aba DADOS AUXILIARES: " & (totalLinhasOrigem - 1) & " linha(s)."
    resultado.TempoExecucao = Timer - tempoInicio
    Set ImportarRelatorioOTMApended = resultado
    Exit Function

TrataErro:
    descricaoErro = Err.Description
    On Error Resume Next
    If Not wbArquivo Is Nothing Then wbArquivo.Close SaveChanges:=False
    On Error GoTo 0
    RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    resultado.Sucesso = False
    resultado.RegistrosProcessados = 0
    resultado.Mensagem = "Erro ao anexar " & nomeRelatorio & ": " & descricaoErro
    resultado.TempoExecucao = Timer - tempoInicio
    Set ImportarRelatorioOTMApended = resultado
End Function

Public Function ValidarNFDiarioContraTransferencia() As clsResultado
    Dim resultado As New clsResultado
    Dim wsTransferencia As Worksheet
    Dim wsAux As Worksheet
    Dim arrTransferencia As Variant
    Dim arrAux As Variant
    Dim nomesTransferencia() As String
    Dim nomesAux() As String
    Dim totalLinhasT As Long
    Dim totalLinhasA As Long
    Dim i As Long, j As Long
    Dim idxDpsT As Long, idxNotasT As Long, idxShipSellT As Long
    Dim idxCteA As Long, idxNfA As Long
    Dim chave As String
    Dim chaveUpper As String
    Dim dictAux As Object
    Dim dictFaltantes As Object
    Dim listaShipSell() As String
    Dim qFaltantes As Long
    Dim shipSellValor As String
    Dim shipSellLista As String

    resultado.Sucesso = True
    resultado.RegistrosProcessados = 0
    resultado.Mensagem = "NF Diário validado com sucesso."

    Set wsTransferencia = ThisWorkbook.Worksheets(SHEET_TRANSFERENCIA)
    Set wsAux = ObterOuCriarAbaDadosAuxiliares
    If wsTransferencia Is Nothing Or wsAux Is Nothing Then
        resultado.Sucesso = False
        resultado.Mensagem = "Aba TRANSFERENCIA ou DADOS AUXILIARES indisponível para validação."
        Set ValidarNFDiarioContraTransferencia = resultado
        Exit Function
    End If

    totalLinhasT = wsTransferencia.Cells(wsTransferencia.Rows.Count, 1).End(xlUp).Row
    totalLinhasA = wsAux.Cells(wsAux.Rows.Count, 2).End(xlUp).Row
    If totalLinhasT < 2 Or totalLinhasA < 2 Then
        resultado.Sucesso = False
        resultado.Mensagem = "DADOS AUXILIARES ou TRANSFERENCIA sem dados para validar."
        Set ValidarNFDiarioContraTransferencia = resultado
        Exit Function
    End If

    arrTransferencia = wsTransferencia.Range(wsTransferencia.Cells(1, 1), wsTransferencia.Cells(totalLinhasT, wsTransferencia.Cells(1, wsTransferencia.Columns.Count).End(xlToLeft).Column)).value
    arrAux = wsAux.Range(wsAux.Cells(1, 1), wsAux.Cells(totalLinhasA, wsAux.Cells(1, wsAux.Columns.Count).End(xlToLeft).Column)).value

    ReDim nomesTransferencia(1 To UBound(arrTransferencia, 2))
    ReDim nomesAux(1 To UBound(arrAux, 2))
    For j = 1 To UBound(arrTransferencia, 2)
        nomesTransferencia(j) = CStr(arrTransferencia(1, j))
    Next j
    For j = 1 To UBound(arrAux, 2)
        nomesAux(j) = CStr(arrAux(1, j))
    Next j

    idxDpsT = modUtils.ProcurarIndiceColunaPorSinonimos(nomesTransferencia, "NÚMERO DPS", "NUMERO DPS", "DPS")
    idxNotasT = modUtils.ProcurarIndiceColunaPorSinonimos(nomesTransferencia, "NOTAS FISCAIS", "NOTA FISCAL", "NF")
    idxShipSellT = modUtils.ProcurarIndiceColunaPorSinonimos(nomesTransferencia, "SHIP SELL", "SHIPSELL")
    idxCteA = modUtils.ProcurarIndiceColunaPorSinonimos(nomesAux, "CTE NUM DOCUMENTO", "NUM DOCUMENTO CTE", "DOCUMENTO CTE", "CTE", "NUMERO CTE")
    idxNfA = modUtils.ProcurarIndiceColunaPorSinonimos(nomesAux, "NF NOTA FISCAL", "NOTA FISCAL", "NF", "NOTA")

    If idxDpsT = 0 Or idxNotasT = 0 Or idxCteA = 0 Or idxNfA = 0 Then
        resultado.Sucesso = False
        resultado.Mensagem = "Não foi possível localizar as colunas-chave para validar o NF Diário contra a TRANSFERENCIA."
        Set ValidarNFDiarioContraTransferencia = resultado
        Exit Function
    End If

    Set dictAux = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(arrAux, 1)
        chave = modUtils.GerarChaveProcessoNF(arrAux(i, idxCteA), arrAux(i, idxNfA))
        If Len(Trim(chave)) > 0 Then
            chaveUpper = UCase(chave)
            If Not dictAux.Exists(chaveUpper) Then dictAux.Add chaveUpper, chave
        End If
    Next i

    Set dictFaltantes = CreateObject("Scripting.Dictionary")
    ReDim listaShipSell(0 To 0)
    shipSellLista = ""

    For i = 2 To UBound(arrTransferencia, 1)
        chave = modUtils.GerarChaveProcessoNF(arrTransferencia(i, idxDpsT), arrTransferencia(i, idxNotasT))
        If Len(Trim(chave)) > 0 Then
            chaveUpper = UCase(chave)
            If Not dictAux.Exists(chaveUpper) Then
                If Not dictFaltantes.Exists(chaveUpper) Then
                    dictFaltantes.Add chaveUpper, chave
                    If idxShipSellT > 0 Then
                        shipSellValor = modUtils.ParaString(arrTransferencia(i, idxShipSellT))
                        If Len(Trim(shipSellValor)) > 0 Then
                            If InStr(1, "," & shipSellLista & ",", "," & shipSellValor & ",", vbTextCompare) = 0 Then
                                If Len(shipSellLista) > 0 Then shipSellLista = shipSellLista & ","
                                shipSellLista = shipSellLista & shipSellValor
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    qFaltantes = dictFaltantes.Count
    UltimaListaShipSellFaltantes = shipSellLista
    UltimaQtdShipSellFaltantes = qFaltantes

    If qFaltantes > 0 Then
        resultado.Sucesso = False
        resultado.RegistrosProcessados = qFaltantes
        resultado.Mensagem = "Faltam " & qFaltantes & " processo(s) no NF Diário para completar a base."
        Set ValidarNFDiarioContraTransferencia = resultado
        Exit Function
    End If

    resultado.Mensagem = "NF Diário cobriu 100% dos processos da TRANSFERENCIA."
    Set ValidarNFDiarioContraTransferencia = resultado
End Function

Private Function ObterOuCriarAbaDadosAuxiliares() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_DADOS_AUXILIARES)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_DADOS_AUXILIARES
    End If

    Set ObterOuCriarAbaDadosAuxiliares = ws
End Function

'---------------------------------------------------------------------------
' FUNÇÕES PRIVADAS (auxiliares internos de modImportacao)
'---------------------------------------------------------------------------

' Limpa completamente a aba: remove conteúdo, formatação e dimensões
Private Sub LimparAba(ws As Worksheet)
ws.UsedRange.Clear
End Sub