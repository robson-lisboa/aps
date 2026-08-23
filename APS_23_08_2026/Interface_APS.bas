Option Explicit

' ============================================================
' INTERFACE_APS
' FASE 3 — VINCULAÇÃO PROGRAMÁTICA DOS BOTÕES
' ============================================================
'
' Responsável por:
'
'   • Configurar botões da aba INICIO
'   • Vincular shapes a macros existentes
'   • Criar interface de cadastros
'   • Integrar navegação com VBA
'
' NÃO altera regras de negócio.
' Apenas configura a interface.
'
' ============================================================


' ============================================================
' CONFIGURAR BOTÕES DA ABA INICIO
' ============================================================

Public Sub ConfigurarBotoesInicio()

    Dim wsInicio As Worksheet
    Dim wsPlan As Worksheet

    On Error Resume Next

    Set wsInicio = _
        ThisWorkbook.Worksheets("INICIO")

    On Error GoTo 0

    If wsInicio Is Nothing Then

        MsgBox _
            "Aba INICIO não encontrada.", _
            vbCritical, _
            "APS"

        Exit Sub

    End If


    ' --------------------------------------------------------
    ' Apagar botões antigos
    ' --------------------------------------------------------

    Dim i As Long

    For i = wsInicio.Shapes.Count To 1 Step -1

        If Left( _
            wsInicio.Shapes(i).Name, _
            9) = "APS_INI_" Then

            wsInicio.Shapes(i).Delete

        End If

    Next i


    ' --------------------------------------------------------
    ' Criar botões do menu principal
    ' --------------------------------------------------------

    ' Botão 1: Ordens de Produção

    CriarBotaoInicio _
        wsInicio, _
        "📋 Ordens de Produção", _
        "IrParaOPs", _
        20, _
        20, _
        140, _
        35


    ' Botão 2: Máquinas

    CriarBotaoInicio _
        wsInicio, _
        "🏭 Máquinas", _
        "IrParaMaquinas", _
        180, _
        20, _
        140, _
        35


    ' Botão 3: Planejamento

    CriarBotaoInicio _
        wsInicio, _
        "📅 Planejamento", _
        "IrParaPlanejamento", _
        340, _
        20, _
        140, _
        35


    ' Botão 4: Atrasos

    CriarBotaoInicio _
        wsInicio, _
        "⚠ Atrasos", _
        "AplicarAtrasosOperacional", _
        500, _
        20, _
        140, _
        35


    ' Botão 5: Eventos

    CriarBotaoInicio _
        wsInicio, _
        "🔧 Eventos/Manutenção", _
        "AplicarEventosOperacional", _
        660, _
        20, _
        160, _
        35


    ' Botão 6: Refeições

    CriarBotaoInicio _
        wsInicio, _
        "🍽 Refeições", _
        "DesenharRefeicoesOperacional", _
        20, _
        65, _
        140, _
        35


    ' Botão 7: Recalcular APS

    CriarBotaoInicio _
        wsInicio, _
        "🔄 Recalcular APS", _
        "RecalcularAPSOperacional", _
        180, _
        65, _
        140, _
        35


    ' Botão 8: Configurações

    CriarBotaoInicio _
        wsInicio, _
        "⚙ Configurações", _
        "IrParaConfig", _
        340, _
        65, _
        140, _
        35


    ' --------------------------------------------------------
    ' Botões na aba PLANEJAMENTO
    ' --------------------------------------------------------

    On Error Resume Next

    Set wsPlan = _
        ThisWorkbook.Worksheets("PLANEJAMENTO")

    On Error GoTo 0

    If Not wsPlan Is Nothing Then

        ' Apagar botões antigos do planejamento

        For i = wsPlan.Shapes.Count To 1 Step -1

            If Left( _
                wsPlan.Shapes(i).Name, _
                10) = "APS_PLAN_" Then

                wsPlan.Shapes(i).Delete

            End If

        Next i


        ' Botão: Recalcular

        CriarBotaoPlanejamento _
            wsPlan, _
            "🔄 Recalcular", _
            "RecalcularAPSOperacional", _
            10, _
            10, _
            110, _
            25


        ' Botão: Atrasos

        CriarBotaoPlanejamento _
            wsPlan, _
            "⚠ Atrasos", _
            "AplicarAtrasosOperacional", _
            130, _
            10, _
            100, _
            25


        ' Botão: Eventos

        CriarBotaoPlanejamento _
            wsPlan, _
            "🔧 Eventos", _
            "AplicarEventosOperacional", _
            240, _
            10, _
            100, _
            25


        ' Botão: Refeições

        CriarBotaoPlanejamento _
            wsPlan, _
            "🍽 Refeições", _
            "DesenharRefeicoesOperacional", _
            350, _
            10, _
            100, _
            25


        ' Botão: Dashboard

        CriarBotaoPlanejamento _
            wsPlan, _
            "🏠 Dashboard", _
            "IrParaInicio", _
            460, _
            10, _
            100, _
            25

    End If


    MsgBox _
        "Botões configurados com sucesso!" & vbCrLf & vbCrLf & _
        "✓ Botões do menu principal vinculados" & vbCrLf & _
        "✓ Botões do planejamento vinculados", _
        vbInformation, _
        "APS"

End Sub


' ============================================================
' CRIAR BOTÃO DA ABA INICIO
' ============================================================

Private Sub CriarBotaoInicio( _
    ByVal ws As Worksheet, _
    ByVal texto As String, _
    ByVal macro As String, _
    ByVal posLeft As Double, _
    ByVal posTop As Double, _
    ByVal largura As Double, _
    ByVal altura As Double)

    Dim shp As Shape

    Dim nomeSeguro As String


    nomeSeguro = _
        "APS_INI_" & _
        Replace(Replace(texto, " ", "_"), "/", "_")


    Set shp = _
        ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            posLeft, _
            posTop, _
            largura, _
            altura)


    shp.Name = nomeSeguro


    shp.TextFrame2.TextRange.Text = texto


    shp.OnAction = macro


    shp.Placement = xlFreeFloating


    With shp.TextFrame2

        .VerticalAnchor = msoAnchorMiddle

        .TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

        .MarginLeft = 3
        .MarginRight = 3
        .MarginTop = 2
        .MarginBottom = 2

    End With


    With shp.TextFrame2.TextRange.Font

        .Size = 10
        .Bold = msoTrue

    End With


    shp.Fill.ForeColor.RGB = RGB(46, 117, 182)


    shp.Line.Visible = msoFalse

End Sub


' ============================================================
' CRIAR BOTÃO DA ABA PLANEJAMENTO
' ============================================================

Private Sub CriarBotaoPlanejamento( _
    ByVal ws As Worksheet, _
    ByVal texto As String, _
    ByVal macro As String, _
    ByVal posLeft As Double, _
    ByVal posTop As Double, _
    ByVal largura As Double, _
    ByVal altura As Double)

    Dim shp As Shape

    Dim nomeSeguro As String


    nomeSeguro = _
        "APS_PLAN_" & _
        Replace(Replace(texto, " ", "_"), "/", "_")


    Set shp = _
        ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            posLeft, _
            posTop, _
            largura, _
            altura)


    shp.Name = nomeSeguro


    shp.TextFrame2.TextRange.Text = texto


    shp.OnAction = macro


    shp.Placement = xlFreeFloating


    With shp.TextFrame2

        .VerticalAnchor = msoAnchorMiddle

        .TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

        .MarginLeft = 3
        .MarginRight = 3
        .MarginTop = 2
        .MarginBottom = 2

    End With


    With shp.TextFrame2.TextRange.Font

        .Size = 9
        .Bold = msoTrue

    End With


    shp.Fill.ForeColor.RGB = RGB(46, 117, 182)


    shp.Line.Visible = msoFalse

End Sub


' ============================================================
' FIM DO MÓDULO
' ============================================================
