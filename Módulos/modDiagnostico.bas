'===============================================================================
' MÓDULO: modDiagnostico
' RESPONSABILIDADE: Executar verificações estruturais sem alterar dados.
'===============================================================================
Option Explicit

Public Function DiagnosticarProjeto() As clsResultado
    Dim resultado As New clsResultado
    Dim validacao As clsResultado
    Dim abas As Variant
    Dim i As Long

    On Error GoTo TrataErro
    abas = Array(SHEET_TRATAMENTO, SHEET_TAB)
    For i = LBound(abas) To UBound(abas)
        Set validacao = ValidarAbaObrigatoria(CStr(abas(i)))
        If Not validacao.Sucesso Then
            resultado.Sucesso = False
            resultado.Mensagem = validacao.Mensagem
            RegistrarEvento nlCritico, catSistema, "modDiagnostico", resultado.Mensagem
            Set DiagnosticarProjeto = resultado
            Exit Function
        End If
    Next i

    Set validacao = modLayout.ValidarContratosLayout()
    If Not validacao.Sucesso Then
        resultado.Sucesso = False
        resultado.Mensagem = validacao.Mensagem
        RegistrarEvento nlCritico, catSistema, "modDiagnostico", resultado.Mensagem
        Set DiagnosticarProjeto = resultado
        Exit Function
    End If

    resultado.Sucesso = True
    resultado.Mensagem = "Diagnóstico estrutural concluído com sucesso."
    RegistrarInfo "modDiagnostico", resultado.Mensagem
    Set DiagnosticarProjeto = resultado
    Exit Function

TrataErro:
    resultado.Sucesso = False
    resultado.Mensagem = "Erro no diagnóstico estrutural: " & Err.Description
    RegistrarEvento nlErro, catSistema, "modDiagnostico", resultado.Mensagem
    Set DiagnosticarProjeto = resultado
End Function

Public Sub ExecutarDiagnosticoProjeto()
    Dim resultado As clsResultado
    Set resultado = DiagnosticarProjeto()
    If resultado.Sucesso Then
        MsgBox resultado.Mensagem, vbInformation, "Diagnóstico do Projeto"
    Else
        MsgBox resultado.Mensagem, vbCritical, "Diagnóstico do Projeto"
    End If
End Sub
