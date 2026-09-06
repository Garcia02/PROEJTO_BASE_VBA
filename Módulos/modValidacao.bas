'===============================================================================
' MÓDULO: modValidacao
' RESPONSABILIDADE: Validação de dados e regras de negócio — validar arquivos,
'                   validar chaves, validar layouts, validar tabelas. Contém
'                   validadores universais parametrizáveis.
'===============================================================================
'---------------------------------------------------------------------------
' Valida a existência de uma aba obrigatória.
Public Function ValidarAbaObrigatoria(ByVal nomeAba As String) As clsResultado
    Dim resultado As New clsResultado
    Dim ws As Worksheet

    On Error GoTo TrataErro
    Set ws = ThisWorkbook.Worksheets(nomeAba)
    resultado.Sucesso = True
    If ws Is Nothing Then
        resultado.Sucesso = False
    End If
    If resultado.Sucesso Then
        resultado.Mensagem = "Aba validada: " & nomeAba
    Else
        resultado.Mensagem = "A aba obrigatória '" & nomeAba & "' não foi encontrada."
    End If
    Set ValidarAbaObrigatoria = resultado
    Exit Function

TrataErro:
    resultado.Sucesso = False
    resultado.Mensagem = "A aba obrigatória '" & nomeAba & "' não foi encontrada."
    RegistrarEvento nlCritico, catSistema, "modValidacao", resultado.Mensagem, nomeAba
    Set ValidarAbaObrigatoria = resultado
End Function

' Valida várias colunas em um único ponto, usando nomes separados por pipe.
Public Function ValidarColunasObrigatorias(ByVal nomesColunas As Variant, _
                                           ByVal colunasObrigatorias As String) As clsResultado
    Dim resultado As New clsResultado
    Dim nomes As Variant
    Dim i As Long
    Dim indice As Long
    Dim ausentes As String

    On Error GoTo TrataErro
    If Not IsArray(nomesColunas) Then
        resultado.Mensagem = "Não foi possível validar as colunas: cabeçalhos inválidos."
        Set ValidarColunasObrigatorias = resultado
        Exit Function
    End If

    nomes = Split(colunasObrigatorias, "|")
    For i = LBound(nomes) To UBound(nomes)
        If Len(Trim(CStr(nomes(i)))) > 0 Then
            indice = modUtils.EncontrarIndiceColuna(nomesColunas, Trim(CStr(nomes(i))))
            If indice = 0 Then
                If Len(ausentes) > 0 Then ausentes = ausentes & ", "
                ausentes = ausentes & Trim(CStr(nomes(i)))
            End If
        End If
    Next i

    resultado.Sucesso = (Len(ausentes) = 0)
    If resultado.Sucesso Then
        resultado.Mensagem = "Colunas obrigatórias validadas."
    Else
        resultado.Mensagem = "Coluna(s) obrigatória(s) ausente(s): " & ausentes
        RegistrarEvento nlCritico, catArquivo, "modValidacao", resultado.Mensagem
    End If
    Set ValidarColunasObrigatorias = resultado
    Exit Function

TrataErro:
    resultado.Sucesso = False
    resultado.Mensagem = "Erro ao validar colunas: " & Err.Description
    RegistrarEvento nlErro, catArquivo, "modValidacao", resultado.Mensagem
    Set ValidarColunasObrigatorias = resultado
End Function

' Valida se o nome do arquivo contém a palavra-chave esperada (case-insensitive)
' Função universal: funciona para receita, despesa ou qualquer futuro relatório
'
' Parâmetros:
'   caminho      = caminho completo do arquivo
'   palavraChave = palavra que deve estar presente no nome do arquivo
'
' Retorna: clsResultado
'   Sucesso = True se o nome do arquivo contém a palavra-chave
'---------------------------------------------------------------------------
Public Function ValidarTipoRelatorio(ByVal caminho As String, _
                                     ByVal palavraChave As String) As clsResultado
    Dim resultado As New clsResultado
    On Error GoTo TrataErro
    Dim nomeArquivo As String
    nomeArquivo = ExtrairNomeArquivo(caminho)
    If Len(nomeArquivo) = 0 Then
        resultado.Sucesso = False
        resultado.Mensagem = "Não foi possível identificar o nome do arquivo."
        Set ValidarTipoRelatorio = resultado
        Exit Function
    End If
    If InStr(1, LCase(nomeArquivo), LCase(palavraChave)) > 0 Then
        resultado.Sucesso = True
        resultado.Mensagem = "Arquivo validado: " & nomeArquivo
        RegistrarInfo "modValidacao", "Tipo de relatório validado: " & palavraChave, nomeArquivo
    Else
        resultado.Sucesso = False
        resultado.Mensagem = "Arquivo inválido. O nome deve conter '" & palavraChave & "'. Arquivo: " & nomeArquivo
        RegistrarEvento nlCritico, catArquivo, "modValidacao", _
                        "Tipo de relatório rejeitado — script interrompido", nomeArquivo, , _
                        "Esperado: '" & palavraChave & "'"
    End If
    Set ValidarTipoRelatorio = resultado
    Exit Function
TrataErro:
    resultado.Sucesso = False
    resultado.Mensagem = "Erro ao validar tipo de relatório: " & Err.Description
    RegistrarEvento nlErro, catArquivo, "modValidacao", _
                    "Erro ao validar: " & Err.Description, caminho
    Set ValidarTipoRelatorio = resultado
End Function
'---------------------------------------------------------------------------
' FUNÇÃO UNIVERSAL: ValidarTabelaTAB
' Valida uma tabela da aba TAB de forma universal:
'   1. Verifica se a tabela existe
'   2. Verifica se as colunas de referência e retorno existem
'   3. Verifica se todos os valores de referência informados existem cadastrados
'   4. Se não existirem, adiciona automaticamente (retornos ficam vazios)
'   5. Verifica se alguma coluna de retorno está vazia (incluindo recém-adicionados)
'
' Parâmetros:
'   nomeTabela        = nome do ListObject na aba TAB
'   colunaReferencia  = nome da coluna de referência (onde buscar)
'   colunaRetorno     = coluna(s) de retorno (pipe-separated para múltiplas:
'                       "CLASSIFICAÇÃO|OBS|Tipo|CONTRATO" ou apenas "CLASSIFICAÇÃO")
'   valoresReferencia = Dictionary com os valores a verificar
'
' Retorna: clsResultado
'   Sucesso = True se tudo OK
'   Sucesso = False se há pendências
'   Mensagem = descrição detalhada de todas as pendências
'   RegistrosProcessados = quantidade de pendências
'---------------------------------------------------------------------------
Public Function ValidarTabelaTAB(ByVal nomeTabela As String, _
                                 ByVal colunaReferencia As String, _
                                 ByVal colunaRetorno As String, _
                                 ByVal valoresReferencia As Object) As clsResultado
    Dim resultado As New clsResultado
    Dim pendencias As String
    Dim totalPendencias As Long
    pendencias = ""
    totalPendencias = 0
    On Error GoTo TrataErro
    Dim wsTAB As Worksheet
    Dim tbl As ListObject
    Set wsTAB = ThisWorkbook.Sheets(SHEET_TAB)
    ' === 1. VERIFICAR SE A TABELA EXISTE ===
    Set tbl = ObterTabela(wsTAB, nomeTabela)
    If tbl Is Nothing Then
        resultado.Sucesso = False
        resultado.Mensagem = "Tabela '" & nomeTabela & "' não encontrada na aba TAB."
        resultado.RegistrosProcessados = 1
        RegistrarEvento nlCritico, catSistema, "modValidacao", _
                        "Tabela não encontrada — script interrompido", nomeTabela
        Set ValidarTabelaTAB = resultado
        Exit Function
    End If
    ' === 2. VERIFICAR COLUNA DE REFERÊNCIA ===
    Dim colRefIdx As Long
    colRefIdx = EncontrarColunaTabela(tbl, colunaReferencia)
    If colRefIdx = 0 Then
        pendencias = pendencias & "- Coluna de referência '" & colunaReferencia & "' não existe na tabela '" & nomeTabela & "'" & vbCrLf
        totalPendencias = totalPendencias + 1
        resultado.Sucesso = False
        resultado.Mensagem = pendencias
        resultado.RegistrosProcessados = totalPendencias
        Set ValidarTabelaTAB = resultado
        Exit Function
    End If
    ' === 3. VERIFICAR COLUNAS DE RETORNO (suporta múltiplas via pipe) ===
    Dim colunasRetorno() As String
    Dim colRetIdxs() As Long
    Dim numColRetorno As Long
    Dim k As Long
    colunasRetorno = Split(colunaRetorno, "|")
    numColRetorno = UBound(colunasRetorno) + 1
    ReDim colRetIdxs(0 To numColRetorno - 1)
    Dim todasColRetExistem As Boolean
    todasColRetExistem = True
    For k = 0 To numColRetorno - 1
        colRetIdxs(k) = EncontrarColunaTabela(tbl, Trim(colunasRetorno(k)))
        If colRetIdxs(k) = 0 Then
            pendencias = pendencias & "- Coluna de retorno '" & Trim(colunasRetorno(k)) & "' não existe na tabela '" & nomeTabela & "'" & vbCrLf
            totalPendencias = totalPendencias + 1
            todasColRetExistem = False
        End If
    Next k
    If Not todasColRetExistem Then
        resultado.Sucesso = False
        resultado.Mensagem = pendencias
        resultado.RegistrosProcessados = totalPendencias
        Set ValidarTabelaTAB = resultado
        Exit Function
    End If
    ' === 4. VERIFICAR REFERÊNCIAS AUSENTES E ADICIONAR ===
    Dim chave As Variant
    Dim valorRef As String
    Dim i As Long
    ' Construir set de referências já cadastradas
    Dim refsExistentes As Object
    Set refsExistentes = CreateObject("Scripting.Dictionary")
    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            valorRef = UCase(Trim(CStr(tbl.DataBodyRange(i, colRefIdx).value)))
            If Len(valorRef) > 0 Then
                If Not refsExistentes.Exists(valorRef) Then
                    refsExistentes.Add valorRef, True
                End If
            End If
        Next i
    End If
    ' Verificar cada valor informado
    Dim refsAdicionadas As String
    refsAdicionadas = ""
    For Each chave In valoresReferencia.Keys
        valorRef = UCase(Trim(CStr(chave)))
        If Len(valorRef) > 0 Then
            If Not refsExistentes.Exists(valorRef) Then
                ' Adicionar nova linha com a referência
                Dim novaLinha As ListRow
                Set novaLinha = tbl.ListRows.Add
                novaLinha.Range(1, colRefIdx).value = valoresReferencia(chave)
                refsAdicionadas = refsAdicionadas & "  - " & valoresReferencia(chave) & vbCrLf
                totalPendencias = totalPendencias + 1
                RegistrarAviso "modValidacao", "Referência adicionada: " & valoresReferencia(chave), nomeTabela
            End If
        End If
    Next chave
    If Len(refsAdicionadas) > 0 Then
        pendencias = pendencias & "- Referências adicionadas automaticamente em '" & nomeTabela & "' (preencha os retornos):" & vbCrLf & refsAdicionadas
    End If
    ' === 5. VERIFICAR RETORNOS VAZIOS (verifica todas as colunas de retorno) ===
    Dim retornosVazios As String
    retornosVazios = ""
    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            valorRef = Trim(CStr(tbl.DataBodyRange(i, colRefIdx).value))
            If Len(valorRef) > 0 Then
                Dim colsVazias As String
                Dim temVazia As Boolean
                colsVazias = ""
                temVazia = False
                For k = 0 To numColRetorno - 1
                    Dim valorRetorno As String
                    valorRetorno = Trim(CStr(tbl.DataBodyRange(i, colRetIdxs(k)).value))
                    If Len(valorRetorno) = 0 Then
                        If Len(colsVazias) > 0 Then colsVazias = colsVazias & ", "
                        colsVazias = colsVazias & Trim(colunasRetorno(k))
                        temVazia = True
                    End If
                Next k
                If temVazia Then
                    retornosVazios = retornosVazios & "  - " & valorRef & " (coluna(s): " & colsVazias & ")" & vbCrLf
                    totalPendencias = totalPendencias + 1
                End If
            End If
        Next i
    End If
    If Len(retornosVazios) > 0 Then
        pendencias = pendencias & "- Retornos vazios em '" & nomeTabela & "':" & vbCrLf & retornosVazios
    End If
    ' === 6. RESULTADO ===
    modUtils.LimparCacheConsultasTAB
    If totalPendencias = 0 Then
        resultado.Sucesso = True
        resultado.Mensagem = "Tabela '" & nomeTabela & "' validada com sucesso."
        RegistrarInfo "modValidacao", "Tabela validada: " & nomeTabela
    Else
        resultado.Sucesso = False
        resultado.Mensagem = pendencias
        RegistrarAviso "modValidacao", "Pendências na tabela " & nomeTabela & ": " & totalPendencias
    End If
    resultado.RegistrosProcessados = totalPendencias
    Set ValidarTabelaTAB = resultado
    Exit Function
TrataErro:
    resultado.Sucesso = False
    resultado.Mensagem = "Erro ao validar tabela '" & nomeTabela & "': " & Err.Description
    resultado.RegistrosProcessados = 1
    RegistrarEvento nlErro, catSistema, "modValidacao", _
                    "Erro ao validar tabela: " & Err.Description, nomeTabela
    Set ValidarTabelaTAB = resultado
End Function
'---------------------------------------------------------------------------
' Valida todas as 5 tabelas TAB necessárias antes da mesclagem.
' Coleta valores de referência da receita (TRATAMENTO) e despesa,
' valida cada tabela em modo batch e retorna o resultado consolidado.
'
' Tabelas validadas:
'   1. TabPontoOperação    — refs combinadas (receita + despesa) — 1 retorno
'   2. TabTipoProcesso     — refs combinadas (receita + despesa) — 4 retornos
'   3. TabTransportadorOTM — refs apenas da receita              — 1 retorno
'   4. TabVeiculoOTM       — refs extraídas de C.E. + MODALIDADE  — 1 retorno
'   5. TabLocalidadesCortes — refs de CIDADE DESTINO + UF DESTINO — 2 retornos
'---------------------------------------------------------------------------
Public Function ValidarTabelasAntesMesclagem(ByVal caminhoDespesa As String) As clsResultado
    Dim resultado As New clsResultado
    Dim pendenciasTotal As String
    Dim totalPendencias As Long
    Dim estadoScreenUpdating As Boolean
    Dim estadoEnableEvents As Boolean
    Dim estadoCalculation As XlCalculation
    Dim descricaoErro As String
    pendenciasTotal = ""
    totalPendencias = 0
    On Error GoTo TrataErro
    CapturarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    Dim wsTratamento As Worksheet
    Dim wsDespesa As Worksheet
    Dim wbDespesa As Workbook
    Dim dimReceita As Object
    Dim dimDespesa As Object
    Dim validacaoColunas As clsResultado
    Dim resLocalidades As clsResultado
    ' === 1. COLETAR REFERÊNCIAS DA RECEITA (TRATAMENTO) ===
    Set wsTratamento = ThisWorkbook.Sheets(SHEET_TRATAMENTO)
    Set dimReceita = CapturarDimensoesRelatorio(wsTratamento)
    If Not dimReceita("temDados") Then
        resultado.Sucesso = False
        resultado.Mensagem = "Aba TRATAMENTO vazia. Importe a receita primeiro."
        Set ValidarTabelasAntesMesclagem = resultado
        Exit Function
    End If
    Set validacaoColunas = ValidarColunasObrigatorias( _
        dimReceita("nomesColunas"), COL_PONTO_OPERACAO & "|" & COL_MODALIDADE)
    If Not validacaoColunas.Sucesso Then
        resultado.Sucesso = False
        resultado.Mensagem = "Layout da receita inválido: " & validacaoColunas.Mensagem
        Set ValidarTabelasAntesMesclagem = resultado
        Exit Function
    End If
    Dim arrReceita As Variant
    arrReceita = wsTratamento.Range( _
                    wsTratamento.Cells(dimReceita("linhaHeader"), dimReceita("primeiraColuna")), _
                    wsTratamento.Cells(dimReceita("ultimaLinha"), dimReceita("ultimaColuna"))).value
    ' Mapear colunas da receita
    Dim colPontoOpRec As Long, colModalidadeRec As Long
    Dim colTransportadoraRec As Long, colCERec As Long
    colPontoOpRec = EncontrarIndiceColuna(dimReceita("nomesColunas"), COL_PONTO_OPERACAO)
    colModalidadeRec = EncontrarIndiceColuna(dimReceita("nomesColunas"), COL_MODALIDADE)
    colTransportadoraRec = EncontrarIndiceColuna(dimReceita("nomesColunas"), COL_TRANSPORTADORA)
    colCERec = EncontrarIndiceColuna(dimReceita("nomesColunas"), COL_CE)
    ' Coletar valores únicos da receita
    Dim refsPontoOpRec As Object, refsModalidadeRec As Object
    Dim refsTransportadoraRec As Object, refsCERec As Object
    Set refsPontoOpRec = ColetarSeExistir(arrReceita, colPontoOpRec, dimReceita("totalRegistros"))
    Set refsModalidadeRec = ColetarSeExistir(arrReceita, colModalidadeRec, dimReceita("totalRegistros"))
    Set refsTransportadoraRec = ColetarSeExistir(arrReceita, colTransportadoraRec, dimReceita("totalRegistros"))
    ' Coletar veículos extraídos de C.E. + MODALIDADE (não usa valor bruto de C.E.)
    If colCERec > 0 And colModalidadeRec > 0 Then
        Set refsCERec = ColetarVeiculosDeCE(arrReceita, colCERec, colModalidadeRec, _
                                            2, dimReceita("totalRegistros") + 1)
    Else
        Set refsCERec = CreateObject("Scripting.Dictionary")
    End If

    Set resLocalidades = ValidarTabLocalidadesCortes(arrReceita, _
                                                     dimReceita("nomesColunas"), _
                                                     dimReceita("totalRegistros"))
    If Not resLocalidades.Sucesso Then
        pendenciasTotal = pendenciasTotal & "=== " & TBL_LOCALIDADES_CORTES & " ===" & vbCrLf & _
                          resLocalidades.Mensagem & vbCrLf
        totalPendencias = totalPendencias + resLocalidades.RegistrosProcessados
    End If
    ' === 2. COLETAR REFERÊNCIAS DA DESPESA ===
    Application.ScreenUpdating = False
    Set wbDespesa = Workbooks.Open(caminhoDespesa, ReadOnly:=True)
    Set wsDespesa = wbDespesa.Sheets(1)
    Set dimDespesa = CapturarDimensoesRelatorio(wsDespesa)
    If Not dimDespesa("temDados") Then
        wbDespesa.Close SaveChanges:=False
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        resultado.Sucesso = False
        resultado.Mensagem = "Arquivo de despesa sem dados."
        Set ValidarTabelasAntesMesclagem = resultado
        Exit Function
    End If
    Set validacaoColunas = ValidarColunasObrigatorias( _
        dimDespesa("nomesColunas"), COL_PONTO_OPERACAO & "|" & COL_MODALIDADE)
    If Not validacaoColunas.Sucesso Then
        wbDespesa.Close SaveChanges:=False
        RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
        resultado.Sucesso = False
        resultado.Mensagem = "Layout da despesa inválido: " & validacaoColunas.Mensagem
        Set ValidarTabelasAntesMesclagem = resultado
        Exit Function
    End If
    Dim arrDespesa As Variant
    arrDespesa = wsDespesa.Range( _
                    wsDespesa.Cells(dimDespesa("linhaHeader"), dimDespesa("primeiraColuna")), _
                    wsDespesa.Cells(dimDespesa("ultimaLinha"), dimDespesa("ultimaColuna"))).value
    wbDespesa.Close SaveChanges:=False
    RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    ' Mapear colunas da despesa (apenas PO e MODALIDADE — transportadora e C.E. são só da receita)
    Dim colPontoOpDesp As Long, colModalidadeDesp As Long
    colPontoOpDesp = EncontrarIndiceColuna(dimDespesa("nomesColunas"), COL_PONTO_OPERACAO)
    colModalidadeDesp = EncontrarIndiceColuna(dimDespesa("nomesColunas"), COL_MODALIDADE)
    ' Coletar valores únicos da despesa
    Dim refsPontoOpDesp As Object, refsModalidadeDesp As Object
    Set refsPontoOpDesp = ColetarSeExistir(arrDespesa, colPontoOpDesp, dimDespesa("totalRegistros"))
    Set refsModalidadeDesp = ColetarSeExistir(arrDespesa, colModalidadeDesp, dimDespesa("totalRegistros"))
    ' === 3. COMBINAR REFERÊNCIAS (receita + despesa) ===
    Dim refsPontoOp As Object, refsModalidade As Object
    Set refsPontoOp = CombinarDicionarios(refsPontoOpRec, refsPontoOpDesp)
    Set refsModalidade = CombinarDicionarios(refsModalidadeRec, refsModalidadeDesp)
    ' === 4. VALIDAR TABPONTOOPERACAO ===
    RegistrarInfo "modValidacao", "Validando " & refsPontoOp.Count & " referência(s) em TabPontoOperação"
    Dim resPO As clsResultado
    Set resPO = ValidarTabelaTAB(TBL_PONTO_OPERACAO, TBL_PO_COL_REF, TBL_PO_COL_RETORNO, refsPontoOp)
    If Not resPO.Sucesso Then
        pendenciasTotal = pendenciasTotal & "=== TabPontoOperação ===" & vbCrLf & resPO.Mensagem & vbCrLf
        totalPendencias = totalPendencias + resPO.RegistrosProcessados
    End If
    ' === 5. VALIDAR TABTIPOPROCESSO (4 colunas de retorno: CLASSIFICAÇÃO|OBS|Tipo|CONTRATO) ===
    RegistrarInfo "modValidacao", "Validando " & refsModalidade.Count & " referência(s) em TabTipoProcesso"
    Dim resTP As clsResultado
    Set resTP = ValidarTabelaTAB(TBL_TIPO_PROCESSO, TBL_TP_COL_REF, _
                TBL_TP_COL_CLASSIFICACAO & "|" & TBL_TP_COL_OBS & "|" & TBL_TP_COL_TIPO & "|" & TBL_TP_COL_CONTRATO, refsModalidade)
    If Not resTP.Sucesso Then
        pendenciasTotal = pendenciasTotal & "=== TabTipoProcesso ===" & vbCrLf & resTP.Mensagem & vbCrLf
        totalPendencias = totalPendencias + resTP.RegistrosProcessados
    End If
    ' === 6. VALIDAR TABTRANSPORTADOROTM ===
    RegistrarInfo "modValidacao", "Validando " & refsTransportadoraRec.Count & " referência(s) em TabTransportadorOTM"
    Dim resTR As clsResultado
    Set resTR = ValidarTabelaTAB(TBL_TRANSPORTADOR_OTM, TBL_TR_COL_REF, TBL_TR_COL_RETORNO, refsTransportadoraRec)
    If Not resTR.Sucesso Then
        pendenciasTotal = pendenciasTotal & "=== TabTransportadorOTM ===" & vbCrLf & resTR.Mensagem & vbCrLf
        totalPendencias = totalPendencias + resTR.RegistrosProcessados
    End If
    ' === 7. VALIDAR TABVEICULOOTM ===
    RegistrarInfo "modValidacao", "Validando " & refsCERec.Count & " referência(s) em TabVeiculoOTM"
    Dim resVE As clsResultado
    Set resVE = ValidarTabelaTAB(TBL_VEICULO_OTM, TBL_VE_COL_REF, TBL_VE_COL_RETORNO, refsCERec)
    If Not resVE.Sucesso Then
        pendenciasTotal = pendenciasTotal & "=== TabVeiculoOTM ===" & vbCrLf & resVE.Mensagem & vbCrLf
        totalPendencias = totalPendencias + resVE.RegistrosProcessados
    End If
    ' === 8. RESULTADO CONSOLIDADO ===
    If totalPendencias = 0 Then
        resultado.Sucesso = True
        resultado.Mensagem = "Todas as tabelas TAB validadas com sucesso."
        RegistrarInfo "modValidacao", "Validação de tabelas TAB concluída sem pendências"
    Else
        resultado.Sucesso = False
        resultado.Mensagem = "Pendências encontradas (" & totalPendencias & "). Corrija e reexecute:" & vbCrLf & vbCrLf & pendenciasTotal
        RegistrarEvento nlCritico, catSistema, "modValidacao", _
                        "Validação de tabelas TAB bloqueou o script", "", totalPendencias, _
                        totalPendencias & " pendência(s) encontrada(s)"
    End If
    resultado.RegistrosProcessados = totalPendencias
    Set ValidarTabelasAntesMesclagem = resultado
    Exit Function
TrataErro:
    descricaoErro = Err.Description
    On Error Resume Next
    If Not wbDespesa Is Nothing Then wbDespesa.Close SaveChanges:=False
    On Error GoTo 0
    RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    resultado.Sucesso = False
    resultado.Mensagem = "Erro ao validar tabelas TAB: " & descricaoErro
    resultado.RegistrosProcessados = 1
    RegistrarEvento nlErro, catSistema, "modValidacao", _
                    "Erro ao validar tabelas TAB: " & descricaoErro
    Set ValidarTabelasAntesMesclagem = resultado
End Function

' Valida localidades e cadastra Cidade/UF ausentes sem tocar na coluna CHAVE,
' que é calculada pela própria tabela.
Private Function ValidarTabLocalidadesCortes(ByVal arrReceita As Variant, _
                                             ByVal nomesColunas As Variant, _
                                             ByVal totalRegistros As Long) As clsResultado
    Dim resultado As New clsResultado
    Dim wsTAB As Worksheet
    Dim tbl As ListObject
    Dim colRef As Long, colPraca As Long, colRegiao As Long
    Dim colCidade As Long, colUF As Long
    Dim colCidadeReceita As Long, colUFReceita As Long
    Dim existentes As Object, referencias As Object
    Dim i As Long
    Dim idx As Variant
    Dim cidade As String, uf As String, chave As String
    Dim pendencias As String
    Dim novaLinha As ListRow

    On Error GoTo TrataErro
    Set wsTAB = ThisWorkbook.Worksheets(SHEET_TAB)
    Set tbl = ObterTabela(wsTAB, TBL_LOCALIDADES_CORTES)
    If tbl Is Nothing Then
        resultado.Mensagem = "Tabela '" & TBL_LOCALIDADES_CORTES & "' não encontrada na aba TAB."
        resultado.RegistrosProcessados = 1
        RegistrarEvento nlCritico, catSistema, "modValidacao", resultado.Mensagem
        Set ValidarTabLocalidadesCortes = resultado
        Exit Function
    End If

    colRef = EncontrarColunaTabela(tbl, TBL_LC_COL_REF)
    colPraca = EncontrarColunaTabela(tbl, TBL_LC_COL_PRACA)
    colRegiao = EncontrarColunaTabela(tbl, TBL_LC_COL_REGIAO)
    colCidade = EncontrarColunaTabela(tbl, TBL_LC_COL_CIDADE)
    colUF = EncontrarColunaTabela(tbl, TBL_LC_COL_UF)
    If colRef = 0 Or colPraca = 0 Or colRegiao = 0 Or colCidade = 0 Or colUF = 0 Then
        resultado.Mensagem = "A tabela '" & TBL_LOCALIDADES_CORTES & _
                             "' deve conter as colunas: chave, Praça, Região, Cidade e UF."
        resultado.RegistrosProcessados = 1
        RegistrarEvento nlCritico, catSistema, "modValidacao", resultado.Mensagem
        Set ValidarTabLocalidadesCortes = resultado
        Exit Function
    End If

    colCidadeReceita = EncontrarIndiceColuna(nomesColunas, "CIDADE DESTINO")
    colUFReceita = EncontrarIndiceColuna(nomesColunas, "UF DESTINO")
    If colCidadeReceita = 0 Or colUFReceita = 0 Then
        resultado.Mensagem = "O relatório de receita deve conter CIDADE DESTINO e UF DESTINO."
        resultado.RegistrosProcessados = 1
        RegistrarEvento nlCritico, catArquivo, "modValidacao", resultado.Mensagem
        Set ValidarTabLocalidadesCortes = resultado
        Exit Function
    End If

    Set existentes = CreateObject("Scripting.Dictionary")
    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            cidade = Trim(ParaString(tbl.DataBodyRange(i, colCidade).value))
            uf = Trim(ParaString(tbl.DataBodyRange(i, colUF).value))
            chave = GerarChaveLocalidade(cidade, uf)
            If Len(chave) > 0 Then
                If Not existentes.Exists(chave) Then existentes.Add chave, True
            End If
        Next i
    End If

    Set referencias = CreateObject("Scripting.Dictionary")
    For i = 2 To totalRegistros + 1
        cidade = Trim(ParaString(arrReceita(i, colCidadeReceita)))
        uf = Trim(ParaString(arrReceita(i, colUFReceita)))
        chave = GerarChaveLocalidade(cidade, uf)
        If Len(chave) > 0 Then
            If Not referencias.Exists(chave) Then referencias.Add chave, Array(cidade, uf)
        End If
    Next i

    For Each idx In referencias.Keys
        If Not existentes.Exists(CStr(idx)) Then
            Set novaLinha = tbl.ListRows.Add
            novaLinha.Range(1, colCidade).value = referencias(idx)(0)
            novaLinha.Range(1, colUF).value = referencias(idx)(1)
            existentes.Add CStr(idx), True
            pendencias = pendencias & "  - " & referencias(idx)(0) & " / " & referencias(idx)(1) & vbCrLf
        End If
    Next idx

    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            cidade = Trim(ParaString(tbl.DataBodyRange(i, colCidade).value))
            uf = Trim(ParaString(tbl.DataBodyRange(i, colUF).value))
            chave = GerarChaveLocalidade(cidade, uf)
            If Len(chave) > 0 Then
                If Len(Trim(ParaString(tbl.DataBodyRange(i, colPraca).value))) = 0 Or _
                   Len(Trim(ParaString(tbl.DataBodyRange(i, colRegiao).value))) = 0 Then
                    pendencias = pendencias & "  - Cadastro incompleto: " & cidade & " / " & uf & vbCrLf
                    resultado.RegistrosProcessados = resultado.RegistrosProcessados + 1
                End If
            End If
        Next i
    End If

    ' A coluna chave é mantida sob responsabilidade da fórmula da tabela.
    If Len(pendencias) > 0 Then
        resultado.Mensagem = "Pendências na tabela '" & TBL_LOCALIDADES_CORTES & _
                             "'. Preencha Praça e Região:" & vbCrLf & pendencias
        If resultado.RegistrosProcessados = 0 Then
            resultado.RegistrosProcessados = referencias.Count
        End If
        RegistrarAviso "modValidacao", "Pendências de classificação na tabela: " & TBL_LOCALIDADES_CORTES
    Else
        resultado.Sucesso = True
        resultado.Mensagem = "Tabela '" & TBL_LOCALIDADES_CORTES & "' validada com sucesso."
        RegistrarInfo "modValidacao", resultado.Mensagem
    End If
    modUtils.LimparCacheConsultasTAB
    Set ValidarTabLocalidadesCortes = resultado
    Exit Function

TrataErro:
    resultado.Mensagem = "Erro ao validar '" & TBL_LOCALIDADES_CORTES & "': " & Err.Description
    resultado.RegistrosProcessados = 1
    RegistrarEvento nlErro, catSistema, "modValidacao", resultado.Mensagem
    Set ValidarTabLocalidadesCortes = resultado
End Function

Private Function GerarChaveLocalidade(ByVal cidade As String, ByVal uf As String) As String
    cidade = UCase(Trim(cidade))
    uf = UCase(Trim(uf))
    If Len(cidade) = 0 Or Len(uf) = 0 Then Exit Function
    GerarChaveLocalidade = cidade & "|" & uf
End Function
'---------------------------------------------------------------------------
' FUNÇÕES PRIVADAS (auxiliares internos)
'---------------------------------------------------------------------------
' Obtém um ListObject pelo nome em uma worksheet
Private Function ObterTabela(ByVal ws As Worksheet, ByVal nomeTabela As String) As ListObject
    On Error Resume Next
    Set ObterTabela = ws.ListObjects(nomeTabela)
    On Error GoTo 0
End Function
' Encontra o índice de uma coluna em um ListObject (case-insensitive)
Private Function EncontrarColunaTabela(ByVal tbl As ListObject, ByVal nomeColuna As String) As Long
    Dim i As Long
    For i = 1 To tbl.ListColumns.Count
        If UCase(Trim(tbl.ListColumns(i).Name)) = UCase(Trim(nomeColuna)) Then
            EncontrarColunaTabela = i
            Exit Function
        End If
    Next i
    EncontrarColunaTabela = 0
End Function
' Combina dois dicionários em um único (chaves únicas, case-insensitive)
Private Function CombinarDicionarios(ByVal dict1 As Object, ByVal dict2 As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim chave As Variant
    For Each chave In dict1.Keys
        If Not dict.Exists(UCase(CStr(chave))) Then
            dict.Add UCase(CStr(chave)), dict1(chave)
        End If
    Next chave
    For Each chave In dict2.Keys
        If Not dict.Exists(UCase(CStr(chave))) Then
            dict.Add UCase(CStr(chave)), dict2(chave)
        End If
    Next chave
    Set CombinarDicionarios = dict
End Function
' Coleta valores únicos de uma coluna se ela existir; retorna dict vazio caso contrário
Private Function ColetarSeExistir(ByVal arrDados As Variant, ByVal colIndex As Long, _
                                   ByVal totalRegistros As Long) As Object
    If colIndex > 0 Then
        Set ColetarSeExistir = ColetarValoresUnicos(arrDados, colIndex, 2, totalRegistros + 1)
    Else
        Set ColetarSeExistir = CreateObject("Scripting.Dictionary")
    End If
End Function