'===============================================================================
' MÓDULO: modTransferencia
' RESPONSABILIDADE: Transferir os dados da aba TRATAMENTO para a aba
'                   TRANSFERENCIA, aplicando o layout da base final (contrato
'                   do modLayout) e separando as notas fiscais em linhas
'                   individuais com replicação dimensional de contexto.
'                   NOTA: esta etapa NÃO grava na base final (DADOS TRATADOS) —
'                   isso ocorre depois, no modAtualizacao.
'===============================================================================
Option Explicit

' ------------------------------------------------------------------
' ROTINA PRINCIPAL
' ------------------------------------------------------------------
Public Function TransferirParaTransferencia() As clsResultado
    Dim resultado As New clsResultado
    Dim wsT As Worksheet
    Dim wsX As Worksheet
    Dim arrT As Variant
    Dim nomesCol() As String
    Dim mapeamento As Collection
    Dim colOrigem() As Long
    Dim ultLinha As Long, ultCol As Long
    Dim i As Long, j As Long, k As Long, linhaDest As Long
    Dim notas() As String, nNotas As Long
    Dim inicio As Double
    Dim idxNotas As Long, idxChave As Long, idxNumeroDPS As Long
    Dim chaveTransferencia As String

    inicio = Timer
    On Error GoTo Falha

    Set wsT = ThisWorkbook.Worksheets(SHEET_TRATAMENTO)
    ultLinha = wsT.Cells(wsT.Rows.Count, 1).End(xlUp).Row
    If ultLinha < 2 Then
        resultado.Sucesso = False
        resultado.Mensagem = "Aba TRATAMENTO sem dados para transferir."
        resultado.TempoExecucao = Timer - inicio
        Set TransferirParaTransferencia = resultado
        Exit Function
    End If

    ultCol = wsT.Cells(1, wsT.Columns.Count).End(xlToLeft).Column
    arrT = wsT.Range(wsT.Cells(1, 1), wsT.Cells(ultLinha, ultCol)).value
    ReDim nomesCol(1 To UBound(arrT, 2))
    For j = 1 To UBound(arrT, 2)
        nomesCol(j) = CStr(arrT(1, j))
    Next j

    Set mapeamento = modLayout.MapeamentoTransferencia()

    ReDim colOrigem(1 To mapeamento.Count)
    idxNotas = 0
    idxChave = 0
    idxNumeroDPS = 0
    For i = 1 To mapeamento.Count
        colOrigem(i) = modUtils.EncontrarIndiceColuna(nomesCol, CStr(mapeamento(i)(0)))
        If UCase(CStr(mapeamento(i)(1))) = "NOTAS FISCAIS" Then idxNotas = i
        If UCase(CStr(mapeamento(i)(1))) = "CHAVE" Then idxChave = i
        If UCase(CStr(mapeamento(i)(1))) = "NÚMERO DPS" Then idxNumeroDPS = i
    Next i

    Set wsX = ObterOuCriarAbaTransferencia()
    LimparAba wsX
    For i = 1 To mapeamento.Count
        wsX.Cells(1, i).value = CStr(mapeamento(i)(1))
    Next i
    If idxChave > 0 Then wsX.Columns(idxChave).NumberFormat = "@"

    linhaDest = 2
    For i = 2 To UBound(arrT, 1)
        If idxNotas > 0 And colOrigem(idxNotas) > 0 Then
            notas = Split(CStr(arrT(i, colOrigem(idxNotas))), SEP_NOTAS_FISCAIS)
            nNotas = UBound(notas) + 1
        Else
            ReDim notas(0 To 0)
            notas(0) = ""
            nNotas = 1
        End If

            For k = 1 To nNotas
            For j = 1 To mapeamento.Count
                    If j = idxNotas Then
                        wsX.Cells(linhaDest, j).value = Trim(notas(k - 1))
                    ElseIf colOrigem(j) > 0 Then
                        wsX.Cells(linhaDest, j).value = arrT(i, colOrigem(j))
                    End If
                Next j

                ' A nota já está individualizada nesta linha. Portanto, a
                ' chave pode ser criada imediatamente, sem aguardar o tratamento.
                If idxChave > 0 And idxNumeroDPS > 0 And _
                   colOrigem(idxNumeroDPS) > 0 Then
                    chaveTransferencia = modUtils.ParaString( _
                        arrT(i, colOrigem(idxNumeroDPS))) & _
                        Trim(modUtils.ParaString(notas(k - 1)))
                    wsX.Cells(linhaDest, idxChave).value = chaveTransferencia
                End If

                linhaDest = linhaDest + 1
            Next k
    Next i

    resultado.Sucesso = True
    resultado.RegistrosProcessados = linhaDest - 2
    resultado.Mensagem = "Transferência concluída com " & resultado.RegistrosProcessados & " linha(s)."
    resultado.TempoExecucao = Timer - inicio
    Set TransferirParaTransferencia = resultado
    Exit Function

Falha:
    resultado.Sucesso = False
    resultado.Mensagem = "Erro em TransferirParaTransferencia: " & Err.Description
    resultado.TempoExecucao = Timer - inicio
    Set TransferirParaTransferencia = resultado
End Function

' ------------------------------------------------------------------
' HELPERS
' ------------------------------------------------------------------
Private Function ObterOuCriarAbaTransferencia() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_TRANSFERENCIA)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_TRANSFERENCIA
    End If
    Set ObterOuCriarAbaTransferencia = ws
End Function

Private Sub LimparAba(ws As Worksheet)
    ws.Cells.Clear
End Sub