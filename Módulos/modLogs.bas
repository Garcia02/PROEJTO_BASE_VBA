'===============================================================================
' MÓDULO: modLogs
' RESPONSABILIDADE: Registro de logs — inicializar, registrar eventos
'                   (info/aviso/erro/crítico), registrar divergências, limpar
'                   logs. Possui seus próprios enums e constantes internas.
'===============================================================================

' --- Enums de domínio do módulo ---
Public Enum NivelLog
    nlInfo = 0
    nlAviso = 1
    nlErro = 2
    nlCritico = 3
End Enum

Public Enum CategoriaLog
    catArquivo = 0
    catImportacao = 1
    catMesclagem = 2
    catDivergencia = 3
    catUpsert = 4
    catChaveDuplicada = 5
    catSemCorrespondencia = 6
    catCentroCusto = 7
    catSistema = 8
End Enum

' --- Constantes internas (colunas da aba de logs) ---
Private Const LOG_COL_DATA As Long = 1
Private Const LOG_COL_NIVEL As Long = 2
Private Const LOG_COL_CATEGORIA As Long = 3
Private Const LOG_COL_MODULO As Long = 4
Private Const LOG_COL_CHAVE As Long = 5
Private Const LOG_COL_DESCRICAO As Long = 6
Private Const LOG_COL_VALOR As Long = 7
Private Const LOG_COL_DETALHE As Long = 8
Private Const LOG_TOTAL_COLUNAS As Long = 8

' --- Cores por nível (preenchimento + fonte) ---
Private Const COR_INFO_FILL As Long = 15921906      ' verde-azulado muito claro
Private Const COR_INFO_FONT As Long = 32768          ' verde escuro

Private Const COR_AVISO_FILL As Long = 13431551     ' amarelo muito claro
Private Const COR_AVISO_FONT As Long = 32832        ' amarelo escuro/marrom

Private Const COR_ERRO_FILL As Long = 13816545      ' vermelho-rosado claro
Private Const COR_ERRO_FONT As Long = 192           ' vermelho escuro

Private Const COR_CRITICO_FILL As Long = 255        ' vermelho puro (RGB 255,0,0)
Private Const COR_CRITICO_FONT As Long = 16777215   ' branco

'---------------------------------------------------------------------------
' FUNÇÕES PÚBLICAS
'---------------------------------------------------------------------------

' Inicializa a aba de logs: cria se não existir, escreve cabeçalhos se vazio
Public Function InicializarSistemaLog() As Boolean
    On Error GoTo TrataErro

    Dim wsLog As Worksheet
    Set wsLog = ObterOuCriarAbaLog

    If wsLog Is Nothing Then
        InicializarSistemaLog = False
        Exit Function
    End If

    ' Escrever cabeçalhos se a aba estiver vazia
    If Len(Trim(CStr(wsLog.Cells(1, LOG_COL_DATA).value))) = 0 Then
        EscreverCabecalhos wsLog
    End If

    InicializarSistemaLog = True
    Exit Function

TrataErro:
    InicializarSistemaLog = False
End Function

' Função núcleo — registra qualquer evento com todos os parâmetros disponíveis
Public Sub RegistrarEvento(ByVal nivel As NivelLog, ByVal categoria As CategoriaLog, _
                           ByVal moduloOrigem As String, ByVal descricao As String, _
                           Optional ByVal chaveAfetada As String = "", _
                           Optional ByVal valorAfetado As Variant, _
                           Optional ByVal detalhe As String = "")
    On Error Resume Next

    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SHEET_LOGS)

    ' Se a aba não existe, inicializar automaticamente
    If wsLog Is Nothing Then
        InicializarSistemaLog
        Set wsLog = ThisWorkbook.Sheets(SHEET_LOGS)
    End If

    If wsLog Is Nothing Then Exit Sub

    ' Garantir que os cabeçalhos existam
    If Len(Trim(CStr(wsLog.Cells(1, LOG_COL_DATA).value))) = 0 Then
        EscreverCabecalhos wsLog
    End If

    ' Encontrar próxima linha vazia
    Dim proximaLinha As Long
    proximaLinha = wsLog.Cells(wsLog.Rows.Count, LOG_COL_DATA).End(xlUp).Row + 1
    If proximaLinha < 2 Then proximaLinha = 2

    ' Gravar registro
    wsLog.Cells(proximaLinha, LOG_COL_DATA).value = Now
    wsLog.Cells(proximaLinha, LOG_COL_NIVEL).value = ObterTextoNivel(nivel)
    wsLog.Cells(proximaLinha, LOG_COL_CATEGORIA).value = ObterTextoCategoria(categoria)
    wsLog.Cells(proximaLinha, LOG_COL_MODULO).value = moduloOrigem
    wsLog.Cells(proximaLinha, LOG_COL_CHAVE).value = chaveAfetada
    wsLog.Cells(proximaLinha, LOG_COL_DESCRICAO).value = descricao

    If Not IsMissing(valorAfetado) Then
        wsLog.Cells(proximaLinha, LOG_COL_VALOR).value = valorAfetado
    End If

    If Len(detalhe) > 0 Then
        wsLog.Cells(proximaLinha, LOG_COL_DETALHE).value = detalhe
    End If

    ' Aplicar cor na célula de NÍVEL
    AplicarCorNivel wsLog, proximaLinha, nivel
End Sub

' --- Funções de conveniência (pré-definem nível e/ou categoria) ---

Public Sub RegistrarInfo(ByVal modulo As String, ByVal descricao As String, _
                         Optional ByVal chave As String = "")
    RegistrarEvento nlInfo, catSistema, modulo, descricao, chave
End Sub

Public Sub RegistrarAviso(ByVal modulo As String, ByVal descricao As String, _
                          Optional ByVal chave As String = "", _
                          Optional ByVal detalhe As String = "")
    RegistrarEvento nlAviso, catSistema, modulo, descricao, chave, , detalhe
End Sub

Public Sub RegistrarErro(ByVal modulo As String, ByVal descricao As String, _
                         Optional ByVal chave As String = "")
    RegistrarEvento nlErro, catSistema, modulo, descricao, chave
End Sub

Public Sub RegistrarInicioEtapa(ByVal modulo As String, ByVal etapa As String)
    RegistrarInfo modulo, "Início da etapa: " & etapa
End Sub

Public Sub RegistrarFimEtapa(ByVal modulo As String, ByVal etapa As String, _
                            ByVal tempoInicio As Double, _
                            Optional ByVal registros As Long = -1)
    Dim mensagem As String
    mensagem = "Fim da etapa: " & etapa & " | Tempo: " & _
               Format(modUtils.TempoDecorrido(tempoInicio), "0.00") & "s"
    If registros >= 0 Then mensagem = mensagem & " | Registros: " & registros
    RegistrarInfo modulo, mensagem
End Sub

Public Sub LogDivergencia(ByVal modulo As String, ByVal chave As String, _
                          ByVal descricao As String, Optional ByVal valor As Variant)
    RegistrarEvento nlAviso, catDivergencia, modulo, descricao, chave, valor
End Sub

Public Sub LogChaveDuplicada(ByVal modulo As String, ByVal chave As String)
    RegistrarEvento nlAviso, catChaveDuplicada, modulo, "Chave duplicada encontrada", chave
End Sub

Public Sub LogSemCorrespondencia(ByVal modulo As String, ByVal chave As String, _
                                 Optional ByVal detalhe As String = "")
    RegistrarEvento nlAviso, catSemCorrespondencia, modulo, _
                    "Sem correspondência encontrada", chave, , detalhe
End Sub

' Remove todos os registros mantendo apenas os cabeçalhos
Public Sub LimparLogs()
    On Error Resume Next

    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SHEET_LOGS)

    If wsLog Is Nothing Then Exit Sub

    Dim ultimaLinha As Long
    ultimaLinha = wsLog.Cells(wsLog.Rows.Count, LOG_COL_DATA).End(xlUp).Row

    If ultimaLinha > 1 Then
        wsLog.Range(wsLog.Cells(2, 1), _
                    wsLog.Cells(ultimaLinha, LOG_TOTAL_COLUNAS)).Clear
    End If
End Sub

Public Function ObterUltimaLinhaLog() As Long
    Dim wsLog As Worksheet
    On Error GoTo Sair
    Set wsLog = ThisWorkbook.Sheets(SHEET_LOGS)
    ObterUltimaLinhaLog = wsLog.Cells(wsLog.Rows.Count, LOG_COL_DATA).End(xlUp).Row
    If ObterUltimaLinhaLog < 1 Then ObterUltimaLinhaLog = 1
    Exit Function
Sair:
    ObterUltimaLinhaLog = 1
End Function

Public Function ObterResumoLogsCriticos(ByVal linhaInicial As Long) As String
    Dim wsLog As Worksheet
    Dim ultimaLinha As Long, i As Long
    Dim resumo As String
    On Error GoTo Sair
    Set wsLog = ThisWorkbook.Sheets(SHEET_LOGS)
    ultimaLinha = wsLog.Cells(wsLog.Rows.Count, LOG_COL_DATA).End(xlUp).Row
    If linhaInicial < 2 Then linhaInicial = 2
    For i = linhaInicial To ultimaLinha
        If UCase(Trim(CStr(wsLog.Cells(i, LOG_COL_NIVEL).value))) = "CRÍTICO" Then
            resumo = resumo & "- " & CStr(wsLog.Cells(i, LOG_COL_MODULO).value) & ": " & _
                     CStr(wsLog.Cells(i, LOG_COL_DESCRICAO).value)
            If Len(Trim(CStr(wsLog.Cells(i, LOG_COL_DETALHE).value))) > 0 Then
                resumo = resumo & " (" & CStr(wsLog.Cells(i, LOG_COL_DETALHE).value) & ")"
            End If
            resumo = resumo & vbCrLf
        End If
    Next i
Sair:
    ObterResumoLogsCriticos = resumo
End Function

'---------------------------------------------------------------------------
' FUNÇÕES PRIVADAS (auxiliares internos)
'---------------------------------------------------------------------------

' Verifica se a aba de logs existe; cria se necessário
Private Function ObterOuCriarAbaLog() As Worksheet
    On Error Resume Next

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = SHEET_LOGS Then
            Set ObterOuCriarAbaLog = ws
            Exit Function
        End If
    Next ws

    ' Criar aba no final
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SHEET_LOGS
    Set ObterOuCriarAbaLog = ws
End Function

' Escreve os cabeçalhos e formata a linha 1
Private Sub EscreverCabecalhos(ByVal ws As Worksheet)
    ws.Cells(1, LOG_COL_DATA).value = "DATA"
    ws.Cells(1, LOG_COL_NIVEL).value = "NÍVEL"
    ws.Cells(1, LOG_COL_CATEGORIA).value = "CATEGORIA"
    ws.Cells(1, LOG_COL_MODULO).value = "MÓDULO"
    ws.Cells(1, LOG_COL_CHAVE).value = "CHAVE"
    ws.Cells(1, LOG_COL_DESCRICAO).value = "DESCRIÇÃO"
    ws.Cells(1, LOG_COL_VALOR).value = "VALOR"
    ws.Cells(1, LOG_COL_DETALHE).value = "DETALHE"

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, LOG_TOTAL_COLUNAS))
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
    End With
End Sub

' Aplica cor de preenchimento e fonte na LINHA INTEIRA conforme o nível
Private Sub AplicarCorNivel(ByVal ws As Worksheet, ByVal linha As Long, _
                            ByVal nivel As NivelLog)
    Dim fill As Long, fontColor As Long

    Select Case nivel
        Case nlInfo
            fill = COR_INFO_FILL
            fontColor = COR_INFO_FONT
        Case nlAviso
            fill = COR_AVISO_FILL
            fontColor = COR_AVISO_FONT
        Case nlErro
            fill = COR_ERRO_FILL
            fontColor = COR_ERRO_FONT
        Case nlCritico
            fill = COR_CRITICO_FILL
            fontColor = COR_CRITICO_FONT
        Case Else
            Exit Sub
    End Select

    ' Colorir a linha inteira (da coluna DATA até DETALHE)
    With ws.Range(ws.Cells(linha, LOG_COL_DATA), _
                  ws.Cells(linha, LOG_TOTAL_COLUNAS))
        .Interior.Color = fill
        .Font.Color = fontColor
        .Font.Bold = (nivel >= nlErro)
    End With
End Sub

' Converte enum NivelLog em texto legível
Private Function ObterTextoNivel(ByVal nivel As NivelLog) As String
    Select Case nivel
        Case nlInfo: ObterTextoNivel = "INFO"
        Case nlAviso: ObterTextoNivel = "AVISO"
        Case nlErro: ObterTextoNivel = "ERRO"
        Case nlCritico: ObterTextoNivel = "CRÍTICO"
        Case Else: ObterTextoNivel = "DESCONHECIDO"
    End Select
End Function

' Converte enum CategoriaLog em texto legível
Private Function ObterTextoCategoria(ByVal categoria As CategoriaLog) As String
    Select Case categoria
        Case catArquivo: ObterTextoCategoria = "Arquivo"
        Case catImportacao: ObterTextoCategoria = "Importação"
        Case catMesclagem: ObterTextoCategoria = "Mesclagem"
        Case catDivergencia: ObterTextoCategoria = "Divergência"
        Case catUpsert: ObterTextoCategoria = "Upsert"
        Case catChaveDuplicada: ObterTextoCategoria = "Chave Duplicada"
        Case catSemCorrespondencia: ObterTextoCategoria = "Sem Correspondência"
        Case catCentroCusto: ObterTextoCategoria = "Centro de Custo"
        Case catSistema: ObterTextoCategoria = "Sistema"
        Case Else: ObterTextoCategoria = "Desconhecido"
    End Select
End Function