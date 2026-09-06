'===============================================================================
' MÓDULO: modChaves
' RESPONSABILIDADE: Geração e gestão de chaves — gerar chaves compostas,
'                   construir dicionários de chaves, buscar por chave.
'                   Não faz validação de regras de negócio.
'===============================================================================
Option Explicit

'---------------------------------------------------------------------------
' Gera uma chave composta a partir do Ponto de Operação base + Número DPS
' Formato: "PONTOOPBASE|NUMERODPS" (sempre em maiúsculas)
'---------------------------------------------------------------------------
Public Function GerarChave(ByVal pontoOperacaoBase As String, _
                          ByVal numeroDPS As String) As String
    pontoOperacaoBase = Trim(pontoOperacaoBase)
    numeroDPS = Trim(numeroDPS)
    If Len(pontoOperacaoBase) = 0 Or Len(numeroDPS) = 0 Then
        GerarChave = ""
        Exit Function
    End If
    GerarChave = UCase(pontoOperacaoBase) & "|" & UCase(numeroDPS)
End Function

' Gera a chave usando a lógica de negócio do ponto de operação base e o número
' do documento, consultando TabPontoOperação para normalizar o Ponto OTM para
' a base equivalente antes de concatenar com o número do documento.
Public Function GerarChavePontoOperacaoNumero(ByVal pontoOperacaoOTM As String, _
                                            ByVal numeroDocumento As String) As String
    Dim pontoOperacaoBase As String
    pontoOperacaoBase = Trim(modUtils.ConsultarTabelaTAB(TBL_PONTO_OPERACAO, _
                                   TBL_PO_COL_REF, TBL_PO_COL_RETORNO, _
                                   Trim(pontoOperacaoOTM)))
    If Len(pontoOperacaoBase) = 0 Then
        pontoOperacaoBase = Trim(pontoOperacaoOTM)
    End If
    GerarChavePontoOperacaoNumero = GerarChave(pontoOperacaoBase, numeroDocumento)
End Function

'---------------------------------------------------------------------------
' Constrói um dicionário de chaves a partir de um array de dados.
' Cada chave aponta para a linha correspondente no array.
'
' Parâmetros:
'   arrDados      = array 2D com os dados (1-based)
'   colPontoOp    = índice da coluna com o Ponto de Operação (OTM)
'   colNumeroDPS  = índice da coluna com o Número DPS
'   linhaInicio   = primeira linha de dados (no array)
'   linhaFim      = última linha de dados (no array)
'
' Retorna: Dictionary {chave: numeroLinha}
'---------------------------------------------------------------------------
Public Function ConstruirDicionarioChaves(ByVal arrDados As Variant, _
                                          ByVal colPontoOp As Long, _
                                          ByVal colNumeroDPS As Long, _
                                          ByVal linhaInicio As Long, _
                                          ByVal linhaFim As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim pontoOp As String, numeroDPS As String, chave As String

    For i = linhaInicio To linhaFim
        pontoOp = Trim(CStr(arrDados(i, colPontoOp)))
        numeroDPS = Trim(CStr(arrDados(i, colNumeroDPS)))
        chave = GerarChavePontoOperacaoNumero(pontoOp, numeroDPS)

        If Len(chave) > 0 And Not dict.Exists(chave) Then
            dict.Add chave, i
        End If
    Next i

    Set ConstruirDicionarioChaves = dict
End Function

'---------------------------------------------------------------------------
' Verifica se uma chave existe no dicionário
'---------------------------------------------------------------------------
Public Function ChaveExiste(ByVal dict As Object, ByVal chave As String) As Boolean
    ChaveExiste = False
    If dict Is Nothing Then Exit Function
    chave = UCase(Trim(chave))
    If Len(chave) = 0 Then Exit Function
    ChaveExiste = dict.Exists(chave)
End Function