'===============================================================================
' MÓDULO: modConfig
' RESPONSABILIDADE: Centralizar constantes de configuração (nomes de sheets,
'                   tabelas e parâmetros globais). Enums de domínio específico
'                   ficam em seus respectivos módulos. Puramente declarativo —
'                   nenhuma rotina executável ou cálculo reside aqui.
'===============================================================================
Option Explicit

' --- Sheets ---
Public Const SHEET_TRATAMENTO As String = "TRATAMENTO"
Public Const SHEET_LOGS As String = "LOGS"
Public Const SHEET_TAB As String = "TAB"
Public Const SHEET_TRANSFERENCIA As String = "TRANSFERENCIA"
Public Const SHEET_DADOS_AUXILIARES As String = "DADOS AUXILIARES"

' --- Tabelas (ListObjects na aba TAB) ---
Public Const TBL_PONTO_OPERACAO As String = "TabPontoOperação"
Public Const TBL_TIPO_PROCESSO As String = "TabTipoProcesso"
Public Const TBL_TRANSPORTADOR_OTM As String = "TabTransportadorOTM"
Public Const TBL_VEICULO_OTM As String = "TabVeiculoOTM"
Public Const TBL_LOCALIDADES_CORTES As String = "TabLocalidadesCortes"

' --- Palavras-chave de tipo de relatório ---
Public Const KW_RECEITA As String = "receita"
Public Const KW_DESPESA As String = "despesa"
Public Const KW_HP_MANIFESTO As String = "manifesto"
Public Const KW_HP_NF_DIARIO As String = "nf_diario"

' --- Colunas dos relatórios (usadas por importação/mesclagem/validação) ---
Public Const COL_PONTO_OPERACAO As String = "PONTO DE OPERAÇÃO"
Public Const COL_NUMERO_DPS As String = "NÚMERO DPS"
Public Const COL_NUMERO_CTRC As String = "NUMERO CTRC"
Public Const COL_NUMERO_CTRC_ALT As String = "NÚMERO CTRC"
Public Const COL_VALOR_SHIPMENT_BUY_TOTAL As String = "VALOR SHIPMENT BUY TOTAL"
Public Const COL_IMPOSTOS As String = "IMPOSTOS"
Public Const COL_MODALIDADE As String = "MODALIDADE"
Public Const COL_TRANSPORTADORA As String = "TRANSPORTADORA"
Public Const COL_CE As String = "C.E."
Public Const COL_CIDADE_DESTINO As String = "CIDADE DESTINO"
Public Const COL_UF_DESTINO As String = "UF DESTINO"
Public Const COL_REGIAO_DESTINO As String = "REGIAO DESTINO"
Public Const COL_REGIAO As String = "Região"

' --- Colunas de resultado (adicionadas na aba TRATAMENTO pela mesclagem) ---
Public Const COL_CHAVE As String = "CHAVE"
Public Const COL_VALOR_FRETE As String = "VALOR FRETE"
Public Const COL_VALOR_COMPLEMENTAR As String = "VALOR COMPLEMENTAR"
Public Const COL_VALOR_OUTROS As String = "VALOR OUTROS"
Public Const COL_IMPOSTOS_DESPESA As String = "IMPOSTOS DESPESA"
Public Const COL_DETALHAMENTO As String = "DETALHAMENTO VALORES"

' --- Colunas da TabPontoOperação ---
Public Const TBL_PO_COL_REF As String = "Ponto de operação OTM"
Public Const TBL_PO_COL_RETORNO As String = "Ponto de operação base"

' --- TabTipoProcesso: referência e colunas de retorno ---
Public Const TBL_TP_COL_REF As String = "CLL_CENTRO_CUSTO"
Public Const TBL_TP_COL_CLASSIFICACAO As String = "CLASSIFICAÇÃO"
Public Const TBL_TP_COL_OBS As String = "OBS"
Public Const TBL_TP_COL_TIPO As String = "Tipo"
Public Const TBL_TP_COL_OPERACAO As String = "OPERAÇÃO"   ' NOVA
Public Const TBL_TP_COL_CONTRATO As String = "CONTRATO"

' --- Colunas da TabTransportadorOTM ---
Public Const TBL_TR_COL_REF As String = "Transportador OTM"
Public Const TBL_TR_COL_RETORNO As String = "Transportador base"

' --- Colunas da TabVeiculoOTM ---
Public Const TBL_VE_COL_REF As String = "Veículo OTM"
Public Const TBL_VE_COL_RETORNO As String = "Veículo BASE"

' --- TabLocalidadesCortes ---
Public Const TBL_LC_COL_REF As String = "chave"
Public Const TBL_LC_COL_PRACA As String = "Praça"
Public Const TBL_LC_COL_REGIAO As String = "Região"
Public Const TBL_LC_COL_CIDADE As String = "Cidade"
Public Const TBL_LC_COL_UF As String = "UF"

' --- Palavras-chave de classificação de modalidade ---
Public Const CLASS_COMPLEMENTAR As String = "COMPLEMENTAR"
Public Const CLASS_TAXA As String = "TAXA"
Public Const CLASS_ADIANTAMENTO_PEDAGIO As String = "ADIANTAMENTO PEDAGIO"

' --- Prefixo para colunas de despesa que já existem na receita ---
Public Const PREFIX_DESPESA As String = "DESP_"
Public Const PREFIX_MANIFESTO As String = "MANIFESTO_"

' --- Cores dos cabeçalhos por origem do dado na aba TRATAMENTO ---
Public Const COR_CABECALHO_RECEITA As Long = 13421823   ' azul muito claro
Public Const COR_CABECALHO_DESPESA As Long = 10092543   ' amarelo muito claro
Public Const COR_CABECALHO_MANIFESTO As Long = 13434828  ' verde muito claro

' --- Separador de notas fiscais ---
Public Const SEP_NOTAS_FISCAIS As String = "/"

' --- Aba DESPESA NE (processos da receita sem correspondência na despesa) ---
Public Const SHEET_DESPESA_NE As String = "DESPESA NE"
Public Const NE_LINHA_CABECALHO As Long = 8
Public Const NE_LINHA_INICIO As Long = 9
Public Const NE_COLUNA_INICIO As Long = 3          ' coluna C
Public Const NE_COL_CHAVE As String = "CHAVE"

' --- DADOS TRATADOS (base final) e colunas de apoio ---
Public Const SHEET_DADOS_TRATADOS As String = "DADOS TRATADOS"
Public Const COL_SHIP_SELL As String = "SHIP SELL"
Public Const COL_TIPO_DPS As String = "TIPO DO DPS"
Public Const COL_SHIPMENT As String = "SHIPMENT"
Public Const COL_VALOR_DPS As String = "VALOR DPS"        ' adicionar se ainda não existir
Public Const COL_VALOR_IMPOSTO As String = "VALOR IMPOSTO"
Public Const KW_COMPLEMENTAR As String = "Complementar"
Public Const TOLERANCIA_SOMA As Double = 0.03