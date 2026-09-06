'===============================================================================
' MÓDULO: modLayout
' RESPONSABILIDADE: Contrato de mapeamento colunar entre a camada de tratamento
'                   (TRATAMENTO) e a base final (TRANSFERENCIA - 59 colunas).
'                   Centraliza o layout da base final em um único ponto de
'                   manutenção, isolando-o do modConfig (que permanece
'                   estritamente declarativo).
'===============================================================================
Option Explicit

' --- Aba DESPESA NE: colunas a trazer (processos da receita sem despesa) ---
Public Const NE_COL_CHAVE As String = "CHAVE"
Public Const NE_COL_PONTO_OPERACAO As String = "PONTO DE OPERAÇÃO"
Public Const NE_COL_TIPO_DPS As String = "TIPO DO DPS"
Public Const NE_COL_NUMERO_DPS As String = "NÚMERO DPS"
Public Const NE_COL_EMISSAO_DPS As String = "EMISSÃO DPS"
Public Const NE_COL_NOTAS_FISCAIS As String = "NOTAS FISCAIS"
Public Const NE_COL_VALOR_DPS As String = "VALOR DPS"
Public Const NE_COL_SHIPMENT As String = "SHIPMENT"
Public Const NE_COL_TRANSPORTADORA As String = "TRANSPORTADORA"
Public Const NE_COL_MODALIDADE As String = "MODALIDADE"

' Ordem das colunas na aba DESPESA NE (cabeçalho na linha 8, dados a partir de C9)
Public Function ColunasDespesaNE() As Variant
    ColunasDespesaNE = Array(NE_COL_CHAVE, NE_COL_PONTO_OPERACAO, NE_COL_TIPO_DPS, _
        NE_COL_NUMERO_DPS, NE_COL_EMISSAO_DPS, NE_COL_NOTAS_FISCAIS, _
        NE_COL_VALOR_DPS, NE_COL_SHIPMENT, NE_COL_TRANSPORTADORA, NE_COL_MODALIDADE)
End Function

' ------------------------------------------------------------------
' CONTRATO DE MAPEAMENTO TRATAMENTO -> TRANSFERENCIA (59 pares)
' Cada par: Array("COLUNA EM TRATAMENTO", "COLUNA NA TRANSFERENCIA").
' Colunas alimentadas por lookup (Tipo de Processo, OBS, PONTO DE OPERAÇÃO,
' TRANSPORTADORA, PLACA) usam a coluna-chave da TRATAMENTO como âncora de
' cópia e são sobrescritas pelo modTratamento com o valor da tabela TAB.
' ------------------------------------------------------------------
Public Function MapeamentoTransferencia() As Collection
    Dim c As New Collection

    ' --- A. Chaves e identificação ---
    c.Add Array("CHAVE", "CHAVE")
    c.Add Array("CLL_CENTRO_CUSTO", "Tipo de Processo")      ' lookup TabTipoProcesso -> Tipo
    c.Add Array("TIPO OPERAÇÃO", "Tipo de Operação")
    c.Add Array(COL_PONTO_OPERACAO, "PONTO DE OPERAÇÃO")    ' lookup TabPontoOperação -> base
    c.Add Array("TIPO DO DPS", "TIPO DO DPS")
    c.Add Array("CLL_CENTRO_CUSTO", "OBS")                   ' lookup TabTipoProcesso -> OBS
    c.Add Array(COL_TRANSPORTADORA, "TRANSPORTADORA")        ' lookup TabTransportadorOTM -> base
    c.Add Array(COL_NUMERO_DPS, "NÚMERO DPS")

    ' --- B. Documental / notas ---
    'TRATAMENTO-TRANSFERENCIA
    c.Add Array("EMISSÃO DPS", "EMISSÃO DPS")
    c.Add Array("NOTAS FISCAIS", "NOTAS FISCAIS")            ' split "/" -> 1 linha por NF
    c.Add Array("NF SERIE", "NF SERIE")
    c.Add Array("EXPEDIDOR", "EXPEDIDOR")
    c.Add Array("UF ORIGEM", "UF ORIGEM")
    c.Add Array("CIDADE ORIGEM", "CIDADE ORIGEM")
    c.Add Array("DESTINATÁRIO", "DESTINATÁRIO")
    c.Add Array("UF DESTINO", "UF DESTINO")
    c.Add Array("CIDADE DESTINO", "CIDADE DESTINO")
    c.Add Array("REGIAO DESTINO", "REGIAO DESTINO")
    c.Add Array("REGIÃO", "Região")
    c.Add Array("NF DATA EMISSÃO", "NF DATA EMISSÃO")
    c.Add Array("CONTRATO", "Contrato")
    c.Add Array("DATA RECEBIMENTO", "Data Recebimento")
    c.Add Array("HORA RECEBIMENTO", "Hora Recebimento")
    c.Add Array("NF PEDIDO", "NF PEDIDO")
    c.Add Array("NF PESO REAL", "NF PESO REAL")
    c.Add Array("NF VAL DECLARADO", "NF VAL DECLARADO")
    c.Add Array("VALOR TOTAL MERCADORIA", "VALOR TOTAL MERCADORIA")
    c.Add Array("NF VOLUME", "NF VOLUME")
    c.Add Array("QTDE PALLET", "QTDE PALLET")
    c.Add Array("MODAL", "Modal")

    ' --- C. Financeiro ---
    c.Add Array("FRETE PESO", "FRETE PESO")
    c.Add Array("GRIS", "GRIS")
    c.Add Array("AD VALOREM", "AD VALOREM")
    c.Add Array("ICMS REDESPACHO", "ICMS REDESPACHO")
    c.Add Array("PEDAGIO", "PEDAGIO")
    c.Add Array("ESTACIONAMENTO", "ESTACIONAMENTO")
    c.Add Array("DESPACHO |CTE |CONTAINER", "DESPACHO |CTE |CONTAINER")
    c.Add Array("VALOR CTE SEM ICMS", "VALOR CTE SEM ICMS")
    c.Add Array("VALOR IMPOSTO", "VALOR IMPOSTO")
    c.Add Array("VALOR DPS", "VALOR DPS")
    c.Add Array("SHIP SELL", "SHIP SELL")
    c.Add Array("IMPOSTOS DESPESA", "IMPOSTO BUY") 'ajustado
    c.Add Array("VALOR FRETE", "VALOR SHIPMENT BUY TOTAL") 'ajustado
    c.Add Array("VALOR COMPLEMENTAR", "VALOR BUY TOTAL (COMPLEMENTAR)") 'ajustado
    c.Add Array("VALOR OUTROS", "OUTROS CUSTOS BUY") 'ajustado
    c.Add Array("DETALHAMENTO VALORES", "MODALIDADE BUY") 'ajustado
    c.Add Array("NÚMERO FATURA", "NUMERO DA FATURA")
    c.Add Array("SHIP BUY", "SHIP BUY")
    c.Add Array(COL_MODALIDADE, "MODALIDADE")
    c.Add Array("NR PARA CONTABILIDADE", "NR PARA CONTABILIDADE")

    ' --- D. Operacional / enriquecimento ---
    c.Add Array(COL_CE, "C.E.")
    c.Add Array("AVERBAÇÃO", "AVERBAÇÃO")
    c.Add Array("OBSERVAÇÃO", "OBSERVAÇÃO")
    c.Add Array("TABELA CLIENTE", "TABELA CLIENTE")
    c.Add Array("MANIFESTO_PLACA DO VEICULO", "PLACA")                             ' lookup TabVeiculoOTM (C.E. - MODALIDADE)
    c.Add Array("MANIFESTO_NOME MOTORISTA", "MOTORISTA")
    c.Add Array("MANIFESTO_NUMERO DO MANIFESTO", "MANIFESTO")
    c.Add Array("OBSERVAÇÃO | COMENTARIOS", "OBSERVAÇÃO | COMENTARIOS")

    ' --- E. Controle ---
    c.Add Array("INDICE", "Indice")                          ' sequencial gerado na transferência

    Set MapeamentoTransferencia = c
End Function

' ------------------------------------------------------------------
' MAPEAMENTO PARA RECONCILIAR RELATÓRIO OTM COM DADOS AUXILIARES
' ------------------------------------------------------------------
' Estrutura: Array("NOME DA COLUNA NO RELATÓRIO OTM", "NOME DA COLUNA EM DADOS AUXILIARES")
' Cada par relaciona diretamente a origem OTM ao destino já existente.
' Ajuste os nomes reais dos campos quando os relatórios forem validados com o
' cliente, mantendo a lógica centralizada neste ponto.
Public Function MapeamentoDadosAuxiliaresOTM() As Collection
    Dim c As New Collection

    c.Add Array("PK", "NF NOTA FISCAL")
    c.Add Array("PK", "NF SERIE")
    c.Add Array("Nome do Destinatario", "DESTINATÁRIO")
    c.Add Array("Cidade do Destinatario", "CTE DESTINO")
    c.Add Array("Data de Emissao", "NF DATA EMISSÃO")
    c.Add Array("Data recebimento romaneio", "NF DATA RECEBIMENTO")
    c.Add Array("Agrupador do Cliente", "NF PEDIDO")
    c.Add Array("Peso Bruto Kg", "NF PESO REAL")
    c.Add Array("Volume", "NF VOLUME")
    c.Add Array("Valor Total", "NF VAL DECLARADO")
    c.Add Array("Transportadora", "NOME TRANSPORTADORA")
    c.Add Array("Tipo de Veiculo", "DCT VEICULO FORNECEDOR ID")

    Set MapeamentoDadosAuxiliaresOTM = c
End Function

' Estrutura: Array("NOME DA COLUNA EM DADOS AUXILIARES", "NOME DA COLUNA EM TRANSFERENCIA")
Public Function MapeamentoDadosAuxiliaresTransferencia() As Collection
    Dim c As New Collection

    c.Add Array("NF SERIE", "NF SERIE")
    c.Add Array("REGIAO DESTINO", "REGIAO DESTINO")
    c.Add Array("REGIÃO", "Região")
    c.Add Array("NF DATA EMISSÃO", "NF DATA EMISSÃO")
    c.Add Array("NF DATA RECEBIMENTO", "Data Recebimento")
    c.Add Array("DATA RECEBIMENTO", "Data Recebimento")
    c.Add Array("HORA RECEBIMENTO", "Hora Recebimento")
    c.Add Array("NF PEDIDO", "NF PEDIDO")
    c.Add Array("NF PESO REAL", "NF PESO REAL")
    c.Add Array("NF VAL DECLARADO", "NF VAL DECLARADO")
    c.Add Array("NF VOLUME", "NF VOLUME")
    c.Add Array("QTDE PALLET", "QTDE PALLET")
    
    c.Add Array("FRETE PESO", "FRETE PESO")
    c.Add Array("GRIS", "GRIS")
    c.Add Array("AD VALOREM", "AD VALOREM")
    
    c.Add Array("OBSERVAÇÃO", "OBSERVAÇÃO")

    Set MapeamentoDadosAuxiliaresTransferencia = c
End Function

' Valida o contrato mínimo de qualquer mapeamento de colunas.
Public Function ValidarMapeamentoColunas(ByVal mapeamento As Collection, _
                                         ByVal nomeMapeamento As String) As clsResultado
    Dim resultado As New clsResultado
    Dim i As Long
    Dim par As Variant
    Dim origem As String
    Dim destino As String

    On Error GoTo TrataErro
    If mapeamento Is Nothing Then
        resultado.Sucesso = False
        resultado.Mensagem = "Mapeamento vazio: " & nomeMapeamento
        Set ValidarMapeamentoColunas = resultado
        Exit Function
    End If
    If mapeamento.Count = 0 Then
        resultado.Sucesso = False
        resultado.Mensagem = "Mapeamento vazio: " & nomeMapeamento
        Set ValidarMapeamentoColunas = resultado
        Exit Function
    End If

    For i = 1 To mapeamento.Count
        par = mapeamento(i)
        If Not IsArray(par) Then
            resultado.Sucesso = False
            resultado.Mensagem = "Entrada inválida no mapeamento '" & nomeMapeamento & "' (item " & i & ")."
            Set ValidarMapeamentoColunas = resultado
            Exit Function
        End If
        origem = Trim(CStr(par(0)))
        destino = Trim(CStr(par(1)))
        If Len(origem) = 0 Or Len(destino) = 0 Then
            resultado.Sucesso = False
            resultado.Mensagem = "Origem ou destino vazio no mapeamento '" & nomeMapeamento & "' (item " & i & ")."
            Set ValidarMapeamentoColunas = resultado
            Exit Function
        End If
    Next i

    resultado.Sucesso = True
    resultado.Mensagem = "Mapeamento validado: " & nomeMapeamento & " (" & mapeamento.Count & " item(ns))."
    Set ValidarMapeamentoColunas = resultado
    Exit Function

TrataErro:
    resultado.Sucesso = False
    resultado.Mensagem = "Erro ao validar o mapeamento '" & nomeMapeamento & "': " & Err.Description
    Set ValidarMapeamentoColunas = resultado
End Function

Public Function ValidarContratosLayout() As clsResultado
    Dim resultado As New clsResultado
    Dim validacao As clsResultado
    Dim mapeamento As Collection
    Dim nomes As Variant
    Dim i As Long

    On Error GoTo TrataErro
    nomes = Array("MapeamentoTransferencia", "MapeamentoDadosAuxiliaresOTM", _
                  "MapeamentoDadosAuxiliaresTransferencia")
    For i = LBound(nomes) To UBound(nomes)
        Select Case CStr(nomes(i))
            Case "MapeamentoTransferencia"
                Set mapeamento = MapeamentoTransferencia()
            Case "MapeamentoDadosAuxiliaresOTM"
                Set mapeamento = MapeamentoDadosAuxiliaresOTM()
            Case "MapeamentoDadosAuxiliaresTransferencia"
                Set mapeamento = MapeamentoDadosAuxiliaresTransferencia()
        End Select
        Set validacao = ValidarMapeamentoColunas(mapeamento, CStr(nomes(i)))
        If Not validacao.Sucesso Then
            resultado.Sucesso = False
            resultado.Mensagem = validacao.Mensagem
            Set ValidarContratosLayout = resultado
            Exit Function
        End If
    Next i

    resultado.Sucesso = True
    resultado.Mensagem = "Contratos de layout validados com sucesso."
    Set ValidarContratosLayout = resultado
    Exit Function

TrataErro:
    resultado.Sucesso = False
    resultado.Mensagem = "Erro ao validar contratos de layout: " & Err.Description
    Set ValidarContratosLayout = resultado
End Function