'===============================================================================
' MÓDULO: modPrincipal
' RESPONSABILIDADE: Orquestração principal — controlar o fluxo de execução,
'                   chamar módulos na ordem correta, tratar erros globais.
'===============================================================================

' Fluxo principal: captura arquivos, valida, importa receita, valida tabelas
' TAB, mescla despesa e transfere para aba TRANSFERENCIA
Public Sub ProcessarReceita()
    On Error GoTo TrataErroGlobal

    Dim arq As New clsArquivos
    Dim resultado As clsResultado
    Dim resultadoFinal As clsResultado
    Dim validacao As clsResultado
    Dim contexto As New clsContextoProcessamento
    Dim inicioEtapa As Double
    Dim estadoScreenUpdating As Boolean
    Dim estadoEnableEvents As Boolean
    Dim estadoCalculation As XlCalculation
    Dim inicioProcessamento As Double
    Dim tempoTotal As Double
    Dim linhaInicialLogCritico As Long
    Dim resumoCriticos As String
    Dim otmImportadoAntesTratamento As Boolean

    inicioProcessamento = Timer
    CapturarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    Set contexto.Arquivos = arq

    ' 0. Inicializar sistema de logs
    InicializarSistemaLog
    linhaInicialLogCritico = ObterUltimaLinhaLog() + 1
    modUtils.LimparCacheConsultasTAB
    RegistrarInfo "modPrincipal", "=== Início do processamento ==="

    Set validacao = DiagnosticarProjeto()
    If Not validacao.Sucesso Then
        MsgBox validacao.Mensagem, vbCritical, "Diagnóstico estrutural inválido"
        RegistrarEvento nlCritico, catSistema, "modPrincipal", _
                        "Processamento interrompido no diagnóstico: " & validacao.Mensagem
        Exit Sub
    End If

    ' 1. Capturar caminhos dos arquivos (receita + despesa num único diálogo)
    contexto.RegistrarEtapa "Seleção de arquivos"
    inicioEtapa = Timer
    RegistrarInicioEtapa "modPrincipal", contexto.EtapaAtual
    If Not CapturarCaminhosArquivos(arq) Then
        MsgBox "Nenhum arquivo selecionado.", vbExclamation, "Aviso"
        RegistrarAviso "modPrincipal", "Processamento cancelado: nenhum arquivo selecionado"
        Exit Sub
    End If
    RegistrarFimEtapa "modPrincipal", contexto.EtapaAtual, inicioEtapa

    ' 2. Solicitar relatório auxiliar opcional logo após selecionar os arquivos
    '    obrigatórios para que o fluxo continue sem interrupções futuras.
    SelecionarRelatorioAuxiliarOpcional arq

    ' 3. Validar tipo do arquivo de receita
    contexto.RegistrarEtapa "Validação dos arquivos"
    inicioEtapa = Timer
    RegistrarInicioEtapa "modPrincipal", contexto.EtapaAtual
    Set validacao = ValidarTipoRelatorio(arq.caminhoReceita, KW_RECEITA)
    If Not validacao.Sucesso Then
        MsgBox validacao.Mensagem, vbExclamation, "Arquivo de Receita Inválido"
        RegistrarEvento nlCritico, catArquivo, "modPrincipal", _
                        "Script interrompido: receita inválida"
        Exit Sub
    End If
    Set contexto.ResultadoValidacao = validacao

    ' 4. Validar tipo do arquivo de despesa
    Set validacao = ValidarTipoRelatorio(arq.caminhoDespesa, KW_DESPESA)
    If Not validacao.Sucesso Then
        MsgBox validacao.Mensagem, vbExclamation, "Arquivo de Despesa Inválido"
        RegistrarEvento nlCritico, catArquivo, "modPrincipal", _
                        "Script interrompido: despesa inválida"
        Exit Sub
    End If
    Set contexto.ResultadoValidacao = validacao
    RegistrarFimEtapa "modPrincipal", contexto.EtapaAtual, inicioEtapa

    ' 5. Importar receita e alocar na aba TRATAMENTO
    contexto.RegistrarEtapa "Importação da receita"
    inicioEtapa = Timer
    RegistrarInicioEtapa "modPrincipal", contexto.EtapaAtual
    Set resultado = ImportarReceita(arq.caminhoReceita)
    Set contexto.ResultadoImportacao = resultado
    If Not resultado.Sucesso Then
        MsgBox resultado.Mensagem, vbCritical, "Erro na Importação"
        RegistrarEvento nlErro, catImportacao, "modPrincipal", _
                        "Script interrompido: falha na importação da receita"
        Exit Sub
    End If
    RegistrarFimEtapa "modPrincipal", contexto.EtapaAtual, inicioEtapa, resultado.RegistrosProcessados

    ' 5. Validar tabelas TAB antes da mesclagem (batch — coleta refs de ambos)
    contexto.RegistrarEtapa "Validação das tabelas TAB"
    inicioEtapa = Timer
    RegistrarInicioEtapa "modPrincipal", contexto.EtapaAtual
    Set validacao = ValidarTabelasAntesMesclagem(arq.caminhoDespesa)
    Set contexto.ResultadoValidacao = validacao
    If Not validacao.Sucesso Then
        MsgBox validacao.Mensagem, vbExclamation, "Pendências nas Tabelas TAB"
        RegistrarEvento nlCritico, catSistema, "modPrincipal", _
                        "Script interrompido: pendências nas tabelas TAB"
        Exit Sub
    End If
    RegistrarFimEtapa "modPrincipal", contexto.EtapaAtual, inicioEtapa, validacao.RegistrosProcessados

    ' 6. Mesclar despesa na aba TRATAMENTO
    contexto.RegistrarEtapa "Mesclagem da despesa"
    inicioEtapa = Timer
    RegistrarInicioEtapa "modPrincipal", contexto.EtapaAtual
    Set resultado = MesclarDespesa(arq.caminhoDespesa)
    Set contexto.ResultadoMesclagem = resultado
    If Not resultado.Sucesso Then
        MsgBox resultado.Mensagem, vbCritical, "Erro na Mesclagem"
        RegistrarEvento nlErro, catMesclagem, "modPrincipal", _
                        "Script interrompido: falha na mesclagem"
        Exit Sub
    End If
    RegistrarFimEtapa "modPrincipal", contexto.EtapaAtual, inicioEtapa, resultado.RegistrosProcessados

    ' 6b. Manifesto opcional: avisar imediatamente se não foi selecionado.
    If Len(Trim(arq.CaminhoHPManifesto)) = 0 Then
        Dim confirmarSemManifesto As VbMsgBoxResult
        confirmarSemManifesto = MsgBox( _
            "Sem o relatório HP - Manifesto Carga, ficarão ausentes os dados de motorista, placa e número de manifesto na base." & vbCrLf & _
            "Deseja continuar sem esse relatório?", _
            vbYesNo + vbExclamation, "Manifesto opcional")
        If confirmarSemManifesto = vbNo Then
            Dim respostaManifesto As Boolean
            respostaManifesto = False
            If Not CapturarCaminhosArquivos(arq) Then
                MsgBox "Operação cancelada. Nenhum arquivo válido foi selecionado.", vbExclamation, "Cancelado"
                RegistrarEvento nlCritico, catArquivo, "modPrincipal", _
                                "Script interrompido: seleção de arquivos cancelada pelo usuário"
                Exit Sub
            End If
            If Len(Trim(arq.CaminhoHPManifesto)) > 0 Then
                respostaManifesto = True
            End If
            If Not respostaManifesto Then
                RegistrarAviso "modPrincipal", "Processamento continuará sem HP Manifesto Carga. Dados de manifesto ficarão ausentes."
            End If
        Else
            RegistrarAviso "modPrincipal", "Processamento continuará sem HP Manifesto Carga. Dados de manifesto ficarão ausentes."
        End If
    Else
        Set validacao = ValidarTipoRelatorio(arq.CaminhoHPManifesto, KW_HP_MANIFESTO)
        If Not validacao.Sucesso Then
            MsgBox validacao.Mensagem, vbExclamation, "Arquivo HP - Manifesto Carga Inválido"
            RegistrarEvento nlCritico, catArquivo, "modPrincipal", _
                            "Script interrompido: HP - Manifesto Carga inválido"
            Exit Sub
        End If

        Set resultado = MesclarHPManifesto(arq.CaminhoHPManifesto)
        If Not resultado.Sucesso Then
            MsgBox resultado.Mensagem, vbCritical, "Erro na Mesclagem do HP Manifesto"
            RegistrarEvento nlErro, catMesclagem, "modPrincipal", _
                            "Script interrompido: falha na mesclagem do HP Manifesto"
            Exit Sub
        End If

        RegistrarInfo "modPrincipal", "HP Manifesto mesclado: " & resultado.RegistrosProcessados & " linha(s) ajustada(s)"
    End If

    ' 7. Transferir dados tratados para aba TRANSFERENCIA
    contexto.RegistrarEtapa "Transferência dos dados"
    inicioEtapa = Timer
    RegistrarInicioEtapa "modPrincipal", contexto.EtapaAtual
    Set resultado = TransferirParaTransferencia()
    Set contexto.ResultadoTransferencia = resultado
    If Not resultado.Sucesso Then
        MsgBox resultado.Mensagem, vbCritical, "Erro na Transferência"
        RegistrarEvento nlErro, catSistema, "modPrincipal", _
                        "Script interrompido: falha na transferência"
        Exit Sub
    End If
    RegistrarFimEtapa "modPrincipal", contexto.EtapaAtual, inicioEtapa, resultado.RegistrosProcessados

    ' O OTM selecionado como auxiliar principal precisa estar disponível antes
    ' do tratamento, pois a etapa 4d consulta DADOS AUXILIARES.
    If Len(Trim(arq.CaminhoRelatorioOTM)) > 0 Then
        Set resultado = ImportarRelatorioAuxiliar(arq.CaminhoRelatorioOTM, "Relatório OTM")
        If Not resultado.Sucesso Then
            MsgBox resultado.Mensagem, vbCritical, "Erro na importação do relatório OTM"
            RegistrarEvento nlErro, catImportacao, "modPrincipal", _
                            "Falha na importação do relatório OTM antes do tratamento"
            Exit Sub
        End If
        otmImportadoAntesTratamento = True
        RegistrarInfo "modPrincipal", _
                     "Relatório OTM carregado em DADOS AUXILIARES antes do tratamento: " & _
                     resultado.RegistrosProcessados & " linha(s)"
    End If

    ' 8. Aplicar tratamentos na aba TRANSFERENCIA
    contexto.RegistrarEtapa "Tratamento dos dados"
    inicioEtapa = Timer
    RegistrarInicioEtapa "modPrincipal", contexto.EtapaAtual
    Set resultado = TratarDadosTransferencia( _
        Len(Trim(arq.CaminhoHPNFDiario)) = 0)
    Set contexto.ResultadoTratamento = resultado
    If Not resultado.Sucesso Then
        MsgBox resultado.Mensagem, vbCritical, "Erro no Tratamento"
        RegistrarEvento nlErro, catSistema, "modPrincipal", _
                        "Script interrompido: falha no tratamento"
        Exit Sub
    End If
    RegistrarFimEtapa "modPrincipal", contexto.EtapaAtual, inicioEtapa, resultado.RegistrosProcessados
    Set resultadoFinal = resultado
    Set contexto.ResultadoFinal = resultadoFinal
    contexto.ListaShipSellFaltantes = UltimaListaShipSellFaltantes
    contexto.QuantidadeShipSellFaltantes = UltimaQtdShipSellFaltantes
    contexto.MensagemDespesaNE = modDespesaNE.UltimaMensagemNE

    ' 8b. Relatório auxiliar opcional já foi definido no início do fluxo.
    '     Valida também se o NF Diário cobre todos os processos da TRANSFERENCIA.
    If Len(Trim(arq.CaminhoHPNFDiario)) > 0 Then
        Set validacao = ValidarTipoRelatorio(arq.CaminhoHPNFDiario, KW_HP_NF_DIARIO)
        If Not validacao.Sucesso Then
            MsgBox validacao.Mensagem, vbExclamation, "Arquivo HP - NF Diário Inválido"
            RegistrarEvento nlCritico, catArquivo, "modPrincipal", _
                            "Arquivo HP - NF Diário inválido; processamento continua sem relatório auxiliar"
        Else
            Set resultado = ImportarNFDiario(arq.CaminhoHPNFDiario)
            If Not resultado.Sucesso Then
                MsgBox resultado.Mensagem, vbCritical, "Erro na importação do HP - NF Diário"
                RegistrarEvento nlErro, catImportacao, "modPrincipal", _
                                "Falha na importação do HP - NF Diário; processamento continua sem relatório auxiliar"
            Else
                Set validacao = ValidarNFDiarioContraTransferencia()
                contexto.ListaShipSellFaltantes = UltimaListaShipSellFaltantes
                contexto.QuantidadeShipSellFaltantes = UltimaQtdShipSellFaltantes
                If Not validacao.Sucesso Then
                    ExibirUserFormNotasConcat contexto.ListaShipSellFaltantes, contexto.QuantidadeShipSellFaltantes
                    arq.CaminhoRelatorioOTM = ""
                    If Not SelecionarRelatorioOTM(arq) Then
                        RegistrarAviso "modPrincipal", "Usuário não selecionou relatório OTM após ausência no NF Diário."
                    Else
                        Set resultado = ImportarRelatorioAuxiliar(arq.CaminhoRelatorioOTM, "Relatório OTM")
                        If Not resultado.Sucesso Then
                            MsgBox resultado.Mensagem, vbCritical, "Erro na importação do relatório OTM"
                            RegistrarEvento nlErro, catImportacao, "modPrincipal", _
                                            "Falha na importação do relatório OTM após ausência no NF Diário"
                        Else
                            RegistrarInfo "modPrincipal", "Relatório OTM importado para DADOS AUXILIARES: " & resultado.RegistrosProcessados & " linha(s)"
                        End If
                    End If
                Else
                    RegistrarInfo "modPrincipal", "HP - NF Diário importado para DADOS AUXILIARES: " & resultado.RegistrosProcessados & " linha(s)"
                End If
            End If
        End If
    ElseIf Len(Trim(arq.CaminhoRelatorioOTM)) > 0 Then
        ' SHIP SELL só é calculado no tratamento, a partir de SHIPMENT.
        Dim shipSellOTM As String
        shipSellOTM = modUtils.ObterValoresUnicosColunaPorNome( _
            ThisWorkbook.Worksheets(SHEET_TRANSFERENCIA), "SHIP SELL")
        If Len(Trim(shipSellOTM)) = 0 Then
            RegistrarAviso "modPrincipal", _
                           "Nenhum SHIP SELL foi encontrado em TRANSFERENCIA após o tratamento; relatório OTM não foi orientado."
        Else
            ExibirUserFormNotasConcat shipSellOTM, resultadoFinal.RegistrosProcessados
        End If
        If otmImportadoAntesTratamento Then
            RegistrarInfo "modPrincipal", "Relatório OTM já estava disponível para o tratamento."
        End If
    End If

    ' 9. Exibir resultado final
    tempoTotal = modUtils.TempoDecorrido(inicioProcessamento)
    MsgBox resultadoFinal.Mensagem & vbCrLf & _
           "Tempo total: " & Format(tempoTotal, "0.00") & "s", _
           vbInformation, "Processamento Concluído"
    RegistrarInfo "modPrincipal", "Processamento concluído com sucesso | Tempo total: " & _
                  Format(tempoTotal, "0.00") & "s"

    RegistrarInfo "modPrincipal", "=== Fim do processamento ==="

    resumoCriticos = ObterResumoLogsCriticos(linhaInicialLogCritico)
    If Len(Trim(resumoCriticos)) > 0 Then
        MsgBox "Foram registrados eventos críticos durante o processamento:" & vbCrLf & _
               vbCrLf & resumoCriticos, vbExclamation, "Atenção - Eventos Críticos"
    End If
    
    ' === ALERTA FINAL: DESPESA NE (exibido só quando o script termina) ===
    contexto.MensagemDespesaNE = modDespesaNE.UltimaMensagemNE
    If Len(contexto.MensagemDespesaNE) > 0 Then
        MsgBox contexto.MensagemDespesaNE, vbInformation, "DESPESA NE"
    End If
    
    Exit Sub

TrataErroGlobal:
    RestaurarEstadoExcel estadoScreenUpdating, estadoEnableEvents, estadoCalculation
    RegistrarEvento nlErro, catSistema, "modPrincipal", _
                    "Erro global: " & Err.Description
    MsgBox "Erro inesperado: " & Err.Description, vbCritical, "Erro"
End Sub