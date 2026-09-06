'===============================================================================
' CLASSE: clsContextoProcessamento
' RESPONSABILIDADE: Concentrar o estado produzido durante uma execução.
'===============================================================================

Option Explicit

Public Arquivos As clsArquivos
Public ResultadoImportacao As clsResultado
Public ResultadoValidacao As clsResultado
Public ResultadoMesclagem As clsResultado
Public ResultadoTransferencia As clsResultado
Public ResultadoTratamento As clsResultado
Public ResultadoFinal As clsResultado
Public ListaShipSellFaltantes As String
Public QuantidadeShipSellFaltantes As Long
Public MensagemDespesaNE As String
Public EtapaAtual As String

Public Sub RegistrarEtapa(ByVal nomeEtapa As String)
    EtapaAtual = nomeEtapa
End Sub
