Option Explicit

' ============================================================
' MÓDULO 11 - FINAL APS
' PADRONIZAÇÃO + VALIDAÇÃO + INICIALIZAÇÃO
' ============================================================

Private Const ABA_RECURSOS As String = "RECURSOS"
Private Const ABA_PLAN As String = "PLANEJAMENTO"
Private Const ABA_RESUMO As String = "RESUMO"


' ============================================================
' EXECUÇÃO PRINCIPAL
' ============================================================

Public Sub IniciarSistemaAPS()

    On Error GoTo TrataErro

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    ' --------------------------------------------------------
    ' 1. CRIAR E PADRONIZAR ESTRUTURA
    ' --------------------------------------------------------

    PrepararAbas


    PrepararDados


    PrepararPlanejamento


    PrepararResumo


    ' --------------------------------------------------------
    ' 2. VALIDAR ESTRUTURA
    ' --------------------------------------------------------

    If Not ValidarSistemaAPS Then

        GoTo SaidaSegura

    End If


    ' --------------------------------------------------------
    ' 3. CALCULAR
    ' --------------------------------------------------------

    Application.Calculate


    ' --------------------------------------------------------
    ' 3.5 MOTOR DE CÁLCULO
    ' --------------------------------------------------------

    ExecutarMotorAPS


    ' --------------------------------------------------------
    ' 3.6 SEQUENCIAMENTO
    ' --------------------------------------------------------

    RecalcularSequenciamentoAPS


    ' --------------------------------------------------------
    ' 4. TIMELINE
    ' --------------------------------------------------------

    ConstruirTimelineAPS


    ' --------------------------------------------------------
    ' 5. CARDS
    ' --------------------------------------------------------

    ApagarCardsAPS

    CriarCardsAPS

    ApagarCardsOrfaos


    ' --------------------------------------------------------
    ' 6. DASHBOARD
    ' --------------------------------------------------------

    CriarDashboardAPS


    ' --------------------------------------------------------
    ' 7. BOTÕES
    ' --------------------------------------------------------

    CriarBotoesAPS


    Application.Calculate


SaidaSegura:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    MsgBox _
        "Sistema APS inicializado." & vbCrLf & vbCrLf & _
        "Estrutura validada." & vbCrLf & _
        "Timeline criada." & vbCrLf & _
        "Cards criados." & vbCrLf & _
        "Dashboard criado." & vbCrLf & _
        "Botões preparados.", _
        vbInformation, _
        "APS"


    Exit Sub


TrataErro:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True


    MsgBox _
        "Ocorreu um erro na inicialização:" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "APS"

End Sub


' ============================================================
' PREPARAR ABAS
' ============================================================

Private Sub PrepararAbas()

    Dim ws As Worksheet


    Set ws = ObterOuCriarAba(ABA_DADOS)

    Set ws = ObterOuCriarAba(ABA_RECURSOS)

    Set ws = ObterOuCriarAba(ABA_PLAN)

    Set ws = ObterOuCriarAba(ABA_RESUMO)

End Sub


' ============================================================
' PREPARAR DADOS
' ============================================================

Private Sub PrepararDados()

    Dim ws As Worksheet


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS)


    ' --------------------------------------------------------
    ' NÃO APAGA DADOS EXISTENTES.
    ' Apenas garante os cabeçalhos.
    ' --------------------------------------------------------

    CriarCabecalhoSeNaoExiste _
        ws, "OP"


    CriarCabecalhoSeNaoExiste _
        ws, "Produto"


    CriarCabecalhoSeNaoExiste _
        ws, "Dosagem"


    CriarCabecalhoSeNaoExiste _
        ws, "Quantidade"


    CriarCabecalhoSeNaoExiste _
        ws, "Unidade"


    CriarCabecalhoSeNaoExiste _
        ws, "Caixas"


    CriarCabecalhoSeNaoExiste _
        ws, "Máquina"


    CriarCabecalhoSeNaoExiste _
        ws, "Sequência"


    CriarCabecalhoSeNaoExiste _
        ws, "Data Planejada"


    CriarCabecalhoSeNaoExiste _
        ws, "Velocidade"


    CriarCabecalhoSeNaoExiste _
        ws, "OEE"


    CriarCabecalhoSeNaoExiste _
        ws, "Capacidade_h"


    CriarCabecalhoSeNaoExiste _
        ws, "Produção_h"


    CriarCabecalhoSeNaoExiste _
        ws, "Setup_h"


    CriarCabecalhoSeNaoExiste _
        ws, "Duração Total (h)"


    CriarCabecalhoSeNaoExiste _
        ws, "Início"


    CriarCabecalhoSeNaoExiste _
        ws, "Fim"


    CriarCabecalhoSeNaoExiste _
        ws, "Status"


    CriarCabecalhoSeNaoExiste _
        ws, "Atraso_h"


    CriarCabecalhoSeNaoExiste _
        ws, "Refeição_h"


    CriarCabecalhoSeNaoExiste _
        ws, "Observação"


    ' --------------------------------------------------------
    ' FORMATAÇÃO
    ' --------------------------------------------------------

    ws.Rows(1).Font.Bold = True


    ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(1, UltimaColuna(ws)) _
    ).AutoFilter


    ws.Columns.AutoFit

End Sub


' ============================================================
' PREPARAR PLANEJAMENTO
' ============================================================

Private Sub PrepararPlanejamento()

    Dim ws As Worksheet


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_PLAN)


    ws.Cells(1, 1).Value = _
        "PLANEJAMENTO APS"


    ws.Cells(2, 1).Value = _
        "LINHA DO TEMPO"


    ws.Cells(4, 1).Value = _
        "MÁQUINA"


    ws.Columns(1).ColumnWidth = 20

End Sub


' ============================================================
' PREPARAR RESUMO
' ============================================================

Private Sub PrepararResumo()

    Dim ws As Worksheet


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_RESUMO)


    ws.Range("A1").Value = _
        "APS - CONTROLE DE PRODUÇÃO"


    ws.Range("A1").Font.Bold = True

End Sub


' ============================================================
' VALIDAR SISTEMA
' ============================================================

Public Function ValidarSistemaAPS() As Boolean

    Dim mensagens As String

    Dim ws As Worksheet


    ValidarSistemaAPS = False


    ' --------------------------------------------------------
    ' DADOS
    ' --------------------------------------------------------

    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS)


    VerificarCampo _
        ws, "OP", mensagens


    VerificarCampo _
        ws, "Produto", mensagens


    VerificarCampo _
        ws, "Máquina", mensagens


    VerificarCampo _
        ws, "Sequência", mensagens


    VerificarCampo _
        ws, "Início", mensagens


    VerificarCampo _
        ws, "Fim", mensagens


    VerificarCampo _
        ws, "Status", mensagens


    VerificarCampo _
        ws, "Atraso_h", mensagens


    VerificarCampo _
        ws, "Refeição_h", mensagens


    ' --------------------------------------------------------
    ' RESULTADO
    ' --------------------------------------------------------

    If mensagens <> "" Then


        MsgBox _
            "A estrutura ainda possui problemas:" & _
            vbCrLf & vbCrLf & _
            mensagens, _
            vbExclamation, _
            "Validação APS"


        Exit Function


    End If


    ValidarSistemaAPS = True

End Function


' ============================================================
' VERIFICAR CAMPO
' ============================================================

Private Sub VerificarCampo( _
    ByVal ws As Worksheet, _
    ByVal campo As String, _
    ByRef mensagens As String)

    If EncontrarColunaFinal( _
        ws, _
        campo) = 0 Then


        mensagens = _
            mensagens & _
            "• Campo ausente: " & _
            campo & vbCrLf


    End If

End Sub


' ============================================================
' CRIAR CABEÇALHO
' ============================================================

Private Sub CriarCabecalhoSeNaoExiste( _
    ByVal ws As Worksheet, _
    ByVal nome As String)

    Dim coluna As Long


    coluna = _
        EncontrarColunaFinal( _
            ws, _
            nome)


    If coluna = 0 Then


        coluna = _
            UltimaColuna(ws) + 1


        ws.Cells(1, coluna).Value = _
            nome


    End If

End Sub


' ============================================================
' OBTER / CRIAR ABA
' ============================================================

Private Function ObterOuCriarAba( _
    ByVal nome As String) As Worksheet

    On Error Resume Next

    Set ObterOuCriarAba = _
        ThisWorkbook.Worksheets(nome)

    On Error GoTo 0


    If ObterOuCriarAba Is Nothing Then


        Set ObterOuCriarAba = _
            ThisWorkbook.Worksheets.Add( _
                After:= _
                ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))


        ObterOuCriarAba.Name = nome


    End If

End Function


' ============================================================
' ENCONTRAR COLUNA
' ============================================================

Private Function EncontrarColunaFinal( _
    ByVal ws As Worksheet, _
    ByVal nome As String) As Long

    Dim coluna As Long

    Dim ultima As Long


    ultima = UltimaColuna(ws)


    For coluna = 1 To ultima


        If StrComp( _
            Trim(CStr( _
                ws.Cells(1, coluna).Value)), _
            Trim(nome), _
            vbTextCompare) = 0 Then


            EncontrarColunaFinal = _
                coluna


            Exit Function


        End If


    Next coluna

End Function


' ============================================================
' ÚLTIMA COLUNA
' ============================================================

Private Function UltimaColuna( _
    ByVal ws As Worksheet) As Long

    UltimaColuna = _
        ws.Cells( _
            1, _
            ws.Columns.Count _
        ).End(xlToLeft).Column


    If UltimaColuna < 1 Then

        UltimaColuna = 1

    End If

End Function


' ============================================================
' FIM DO MÓDULO 11
' ============================================================