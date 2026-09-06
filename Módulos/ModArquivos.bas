'===============================================================================
' MÓDULO: modArquivos
' RESPONSABILIDADE: Captura e gestão de arquivos — abrir, fechar, obter
'                   caminhos, contar arquivos. Nenhuma validação aqui.
'===============================================================================
Option Explicit

' Exibe diálogo único para seleção dos arquivos de receita, despesa e manifesto.
' O manifesto é opcional, mas deve ser solicitado no mesmo popup da receita e
' despesa, avisando o usuário sobre as colunas que ficarão ausentes caso não
' selecione o relatório.
Public Function CapturarCaminhosArquivos(ByRef arq As clsArquivos) As Boolean
    On Error GoTo TrataErro

    Dim fd As FileDialog
    Dim totalArquivos As Long
    Dim arquivos() As String
    Dim i As Long
    Dim nomeArquivo As String

    arq.caminhoReceita = ""
    arq.caminhoDespesa = ""
    arq.CaminhoHPManifesto = ""
    arq.CaminhoHPNFDiario = ""
    arq.CaminhoRelatorioOTM = ""

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Selecione RECEITA, DESPESA e, opcionalmente, HP - Manifesto Carga"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Arquivos Excel", "*.xlsx;*.xls;*.xlsm;*.xlsb"

        If .Show <> -1 Then
            RegistrarAviso "modArquivos", "Usuário cancelou a seleção dos arquivos"
            CapturarCaminhosArquivos = False
            Exit Function
        End If

        totalArquivos = .SelectedItems.Count
        If totalArquivos < 2 Or totalArquivos > 3 Then
            MsgBox "Selecione 2 arquivos obrigatórios:" & vbCrLf & _
                    "- RECEITA" & vbCrLf & _
                    "- DESPESA" & vbCrLf & vbCrLf & _
                    "e, opcionalmente, o relatório HP - Manifesto Carga.", _
                    vbExclamation, "Seleção inválida"
            RegistrarAviso "modArquivos", "Seleção inválida: " & totalArquivos & " arquivo(s) selecionado(s)"
            CapturarCaminhosArquivos = False
            Exit Function
        End If

        ReDim arquivos(1 To totalArquivos)
        For i = 1 To totalArquivos
            arquivos(i) = .SelectedItems(i)
        Next i

        For i = 1 To totalArquivos
            nomeArquivo = LCase(ExtrairNomeArquivo(arquivos(i)))
            If InStr(nomeArquivo, KW_RECEITA) > 0 Then
                arq.caminhoReceita = arquivos(i)
            ElseIf InStr(nomeArquivo, KW_DESPESA) > 0 Then
                arq.caminhoDespesa = arquivos(i)
            ElseIf InStr(nomeArquivo, KW_HP_MANIFESTO) > 0 Then
                arq.CaminhoHPManifesto = arquivos(i)
            End If
        Next i

        If Len(arq.caminhoReceita) = 0 Or Len(arq.caminhoDespesa) = 0 Then
            MsgBox "Não foi possível identificar corretamente os arquivos de RECEITA e DESPESA." & vbCrLf & _
                    "Use nomes contendo 'receita' e 'despesa' no arquivo." & vbCrLf & _
                    "O relatório HP - Manifesto Carga é opcional.", _
                    vbExclamation, "Identificação falhou"
            RegistrarAviso "modArquivos", "Arquivos obrigatórios não identificados pelo nome"
            CapturarCaminhosArquivos = False
            Exit Function
        End If

        If Len(arq.CaminhoHPManifesto) = 0 Then
            MsgBox "Atenção: sem o relatório HP - Manifesto Carga, os dados de motorista, placa e número de manifesto ficarão ausentes." & vbCrLf & _
                    "Você pode selecionar esse relatório agora em uma nova tentativa.", _
                    vbExclamation, "Manifesto opcional ausente"
            RegistrarAviso "modArquivos", "Manifesto opcional não selecionado. Solicitação imediata de nova chance.", arq.caminhoReceita

            If MsgBox("Deseja selecionar o relatório HP - Manifesto Carga agora?", vbYesNo + vbQuestion, "Selecionar manifesto") = vbYes Then
                If Not SelecionarManifestoOpcional(arq) Then
                    RegistrarAviso "modArquivos", "Usuário optou por continuar sem HP Manifesto Carga."
                End If
            End If
        End If

        RegistrarInfo "modArquivos", "Receita: " & arq.caminhoReceita
        RegistrarInfo "modArquivos", "Despesa: " & arq.caminhoDespesa
        If Len(arq.CaminhoHPManifesto) > 0 Then
            RegistrarInfo "modArquivos", "Manifesto: " & arq.CaminhoHPManifesto
        Else
            RegistrarAviso "modArquivos", "HP Manifesto Carga não selecionado. Dados de motorista/placa/manifesto ficarão ausentes."
        End If
        CapturarCaminhosArquivos = True
    End With

    Exit Function

TrataErro:
    RegistrarErro "modArquivos", "Erro ao capturar caminhos: " & Err.Description
    CapturarCaminhosArquivos = False
End Function

Private Function SelecionarManifestoOpcional(ByRef arq As clsArquivos) As Boolean
    Dim fd As FileDialog
    Dim caminho As String
    Dim nomeArquivo As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecione o relatório HP - Manifesto Carga (opcional)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Arquivos Excel", "*.xlsx;*.xls;*.xlsm;*.xlsb"

        If .Show <> -1 Then
            SelecionarManifestoOpcional = False
            Exit Function
        End If

        If .SelectedItems.Count = 0 Then
            SelecionarManifestoOpcional = False
            Exit Function
        End If

        caminho = .SelectedItems(1)
        nomeArquivo = LCase(ExtrairNomeArquivo(caminho))
        If InStr(nomeArquivo, KW_HP_MANIFESTO) > 0 Then
            arq.CaminhoHPManifesto = caminho
            RegistrarInfo "modArquivos", "Manifesto opcional selecionado na segunda chance: " & caminho
            SelecionarManifestoOpcional = True
        Else
            MsgBox "O arquivo selecionado não parece ser o relatório HP - Manifesto Carga.", vbExclamation, "Arquivo inválido"
            SelecionarManifestoOpcional = False
        End If
    End With
End Function

Public Function SelecionarNFDiarioOpcional(ByRef arq As clsArquivos) As Boolean
    Dim fd As FileDialog
    Dim caminho As String
    Dim nomeArquivo As String
    Dim tentarNovamente As VbMsgBoxResult

    Do
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .Title = "Selecione o relatório HP - NF Diário (opcional)"
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Arquivos Excel", "*.xlsx;*.xls;*.xlsm;*.xlsb"

            If .Show <> -1 Then
                SelecionarNFDiarioOpcional = False
                Exit Function
            End If

            If .SelectedItems.Count = 0 Then
                SelecionarNFDiarioOpcional = False
                Exit Function
            End If

            caminho = .SelectedItems(1)
            nomeArquivo = LCase(ExtrairNomeArquivo(caminho))
            If InStr(nomeArquivo, KW_HP_NF_DIARIO) > 0 Then
                arq.CaminhoHPNFDiario = caminho
                RegistrarInfo "modArquivos", "NF Diário opcional selecionado: " & caminho
                SelecionarNFDiarioOpcional = True
                Exit Function
            End If
        End With

        MsgBox "O arquivo selecionado não parece ser o relatório HP - NF Diário." & vbCrLf & _
               "Tente novamente para selecionar o arquivo correto.", vbExclamation, "Arquivo inválido"
        tentarNovamente = MsgBox("Deseja tentar selecionar o HP - NF Diário novamente?", vbYesNo + vbQuestion, "Nova tentativa")
        If tentarNovamente = vbNo Then
            arq.CaminhoHPNFDiario = ""
            SelecionarNFDiarioOpcional = False
            Exit Function
        End If
    Loop
End Function

Public Function ExibirUserFormNotasConcat(ByVal shipSellLista As String, Optional ByVal qtdProcessos As Long = 0) As Boolean
    Dim frm As UserFormNotasConcat
    Dim qtdNotas As Long
    Dim wsTransferencia As Worksheet

    Set wsTransferencia = ThisWorkbook.Worksheets(SHEET_TRANSFERENCIA)
    If qtdProcessos > 0 Then
        qtdNotas = qtdProcessos
    Else
        qtdNotas = modUtils.ContarValoresColunaPorNome(wsTransferencia, "SHIP SELL")
    End If

    Set frm = New UserFormNotasConcat
    frm.TextBoxQTDNotas.Text = CStr(qtdNotas)
    frm.TextBoxNotasConcat.Text = shipSellLista
    frm.Show vbModal
    ExibirUserFormNotasConcat = True
End Function

Public Function SelecionarRelatorioOTM(ByRef arq As clsArquivos) As Boolean
    Dim fd As FileDialog
    Dim caminho As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecione o relatório extraído do OTM (opcional)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Arquivos Excel", "*.xlsx;*.xls;*.xlsm;*.xlsb"

        If .Show <> -1 Then
            SelecionarRelatorioOTM = False
            Exit Function
        End If

        If .SelectedItems.Count = 0 Then
            SelecionarRelatorioOTM = False
            Exit Function
        End If

        caminho = .SelectedItems(1)
        arq.CaminhoRelatorioOTM = caminho
        RegistrarInfo "modArquivos", "Relatório OTM selecionado: " & caminho
        SelecionarRelatorioOTM = True
    End With
End Function

Public Function SelecionarRelatorioAuxiliarOpcional(ByRef arq As clsArquivos) As Boolean
    Dim opcao As VbMsgBoxResult
    opcao = MsgBox("Deseja importar um relatório auxiliar?" & vbCrLf & _
                   "Sim = HP - NF Diário" & vbCrLf & _
                   "Não = Relatório extraído do OTM" & vbCrLf & _
                   "Cancelar = ignorar", vbYesNoCancel + vbQuestion, "Relatório auxiliar opcional")

    Select Case opcao
        Case vbYes
            arq.CaminhoHPNFDiario = ""
            arq.CaminhoRelatorioOTM = ""
            If Not SelecionarNFDiarioOpcional(arq) Then
                RegistrarAviso "modArquivos", "Usuário optou por não selecionar NF Diário."
                SelecionarRelatorioAuxiliarOpcional = False
                Exit Function
            End If
            SelecionarRelatorioAuxiliarOpcional = True

        Case vbNo
            arq.CaminhoHPNFDiario = ""
            arq.CaminhoRelatorioOTM = ""
            If Not SelecionarRelatorioOTM(arq) Then
                RegistrarAviso "modArquivos", "Usuário optou por não selecionar relatório OTM."
                SelecionarRelatorioAuxiliarOpcional = False
                Exit Function
            End If
            SelecionarRelatorioAuxiliarOpcional = True

        Case vbCancel
            RegistrarAviso "modArquivos", "Usuário ignorou relatório auxiliar opcional."
            SelecionarRelatorioAuxiliarOpcional = False
    End Select
End Function

' Mantido para compatibilidade; não é usado no fluxo principal deliberação do
' manifesto, que agora é realizado junto com receita + despesa em um único popup.
Public Function CapturarCaminhoHPManifesto(ByRef arq As clsArquivos) As Boolean
    CapturarCaminhoHPManifesto = Len(arq.CaminhoHPManifesto) > 0
End Function