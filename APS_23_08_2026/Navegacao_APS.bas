Option Explicit

' ============================================================
' NAVEGACAO_APS
' FASE 2 — INTERFACE PROFISSIONAL
' ============================================================
'
' Responsável por:
'
'   • Navegação entre abas visíveis
'   • Chamadas às macros existentes
'   • Atualização de interface
'   • Ocultação de abas técnicas
'
' NÃO altera regras de negócio.
' Apenas conecta a interface aos módulos existentes.
'
' ============================================================


' ============================================================
' NAVEGAÇÃO PRINCIPAL
' ============================================================

Public Sub IrParaInicio()
    IrParaAba "INICIO"
End Sub


Public Sub IrParaOPs()
    IrParaAba "OPERACOES"
End Sub


Public Sub IrParaMaquinas()
    IrParaAba "MAQUINAS"
End Sub


Public Sub IrParaPlanejamento()
    IrParaAba "PLANEJAMENTO"
End Sub


Public Sub IrParaConfig()
    IrParaAba "CONFIG"
End Sub


' ============================================================
' IR PARA ABA
' ============================================================

Private Sub IrParaAba( _
    ByVal nomeAba As String)

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets(nomeAba)

    On Error GoTo 0

    If ws Is Nothing Then

        MsgBox _
            "Aba '" & nomeAba & "' não encontrada.", _
            vbCritical, _
            "APS"

        Exit Sub

    End If


    ws.Activate

End Sub


' ============================================================
' ATUALIZAR DASHBOARD OPERACIONAL
' ============================================================

Public Sub AtualizarDashboardOperacional()

    Dim wsInicio As Worksheet

    On Error Resume Next

    Set wsInicio = _
        ThisWorkbook.Worksheets("INICIO")

    On Error GoTo 0

    If wsInicio Is Nothing Then Exit Sub


    ' --------------------------------------------------------
    ' Atualizar data/hora
    ' --------------------------------------------------------

    wsInicio.Range("A2").Value = _
        "Sistema de Planejamento e Controle da Produção  |  " & _
        Format(Now, "dd/mm/yyyy hh:mm")


    ' --------------------------------------------------------
    ' Chamar dashboard técnico existente
    ' --------------------------------------------------------

    On Error Resume Next

    AtualizarDashboardAPS

    On Error GoTo 0

End Sub


' ============================================================
' ABRIR CADASTRO DE OP
' ============================================================

Public Sub AbrirCadastroOP()

    Dim wsOps As Worksheet

    On Error Resume Next

    Set wsOps = _
        ThisWorkbook.Worksheets("OPs")

    On Error GoTo 0

    If wsOps Is Nothing Then

        MsgBox _
            "Aba OPs não encontrada.", _
            vbCritical, _
            "APS"

        Exit Sub

    End If


    wsOps.Activate

End Sub


' ============================================================
' ABRIR CADASTRO DE MÁQUINA
' ============================================================

Public Sub AbrirCadastroMaquina()

    Dim wsMaq As Worksheet

    On Error Resume Next

    Set wsMaq = _
        ThisWorkbook.Worksheets("MAQUINAS")

    On Error GoTo 0

    If wsMaq Is Nothing Then

        MsgBox _
            "Aba MAQUINAS não encontrada.", _
            vbCritical, _
            "APS"

        Exit Sub

    End If


    wsMaq.Activate

End Sub


' ============================================================
' RECALCULAR APS (FLUXO PRINCIPAL)
' ============================================================

Public Sub RecalcularAPSOperacional()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    On Error Resume Next

    AtualizarAPS

    On Error GoTo 0


    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    AtualizarDashboardOperacional

End Sub


' ============================================================
' APLICAR ATRASOS (FLUXO PRINCIPAL)
' ============================================================

Public Sub AplicarAtrasosOperacional()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    On Error Resume Next

    AtualizarAtrasosAPS

    On Error GoTo 0


    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    AtualizarDashboardOperacional

End Sub


' ============================================================
' APLICAR EVENTOS (FLUXO PRINCIPAL)
' ============================================================

Public Sub AplicarEventosOperacional()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    On Error Resume Next

    AplicarEventosAPS

    On Error GoTo 0


    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    AtualizarDashboardOperacional

End Sub


' ============================================================
' DESENHAR REFEIÇÕES (FLUXO PRINCIPAL)
' ============================================================

Public Sub DesenharRefeicoesOperacional()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    On Error Resume Next

    DesenharRefeicoes

    On Error GoTo 0


    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    AtualizarDashboardOperacional

End Sub


' ============================================================
' CONFIGURAR APS (FLUXO PRINCIPAL)
' ============================================================

Public Sub ConfigurarAPSOperacional()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    On Error Resume Next

    ConfigurarAPS

    On Error GoTo 0


    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    MsgBox _
        "APS configurado com sucesso!" & vbCrLf & vbCrLf & _
        "Estrutura criada." & vbCrLf & _
        "Abas preparadas.", _
        vbInformation, _
        "APS"

End Sub


' ============================================================
' APLICAR ALTERAÇÕES DE CARDS (DRAG & DROP)
' ============================================================

Public Sub AplicarAlteracoesCardsOperacional()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    On Error Resume Next

    AplicarAlteracoesCards

    On Error GoTo 0


    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    AtualizarDashboardOperacional

End Sub


' ============================================================
' FIM DO MÓDULO
' ============================================================
