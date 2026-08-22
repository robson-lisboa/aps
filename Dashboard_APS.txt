Option Explicit

' ============================================================
' MÓDULO 9 - DASHBOARD / KPIs APS
' ============================================================
'
' RESPONSABILIDADE:
'
'   • Criar Dashboard
'   • Atualizar KPIs
'   • Consolidar dados por máquina
'   • Mostrar status das OPs
'   • Mostrar horas planejadas
'   • Mostrar horas de atraso
'   • Mostrar utilização
'   • Criar gráficos
'
' FONTE:
'
'   Aba DADOS
'
' DESTINO:
'
'   Aba RESUMO
'
' ============================================================


Private Const ABA_DADOS_DASH As String = "DADOS"
Private Const ABA_RESUMO_DASH As String = "RESUMO"


' ============================================================
' CRIAR DASHBOARD
' ============================================================

Public Sub CriarDashboardAPS()

    On Error GoTo TrataErro

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim ws As Worksheet

    Set ws = ObterOuCriarResumoDashboard()


    ' --------------------------------------------------------
    ' Limpar dashboard existente
    ' --------------------------------------------------------

    LimparDashboard ws


    ' --------------------------------------------------------
    ' Criar estrutura
    ' --------------------------------------------------------

    CriarTituloDashboard ws

    CriarCardsKPI ws

    CriarTabelaMaquinas ws

    CriarTabelaStatus ws

    CriarGraficoStatus ws

    CriarGraficoMaquinas ws

    FormatarDashboard ws


    Application.EnableEvents = True
    Application.ScreenUpdating = True


    MsgBox _
        "Dashboard criado com sucesso.", _
        vbInformation, _
        "APS - Dashboard"

    Exit Sub


TrataErro:

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "Erro ao criar o Dashboard:" & vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "APS - Dashboard"

End Sub


' ============================================================
' ATUALIZAR DASHBOARD
' ============================================================

Public Sub AtualizarDashboardAPS()

    On Error GoTo TrataErro

    Application.ScreenUpdating = False
    Application.EnableEvents = False


    Dim ws As Worksheet

    Set ws = ObterOuCriarResumoDashboard()


    AtualizarKPIsDashboard ws

    AtualizarTabelaMaquinasDashboard ws

    AtualizarTabelaStatusDashboard ws

    AtualizarGraficosDashboard ws

    ws.Calculate


    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Exit Sub


TrataErro:

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "Erro ao atualizar o Dashboard:" & vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "APS - Dashboard"

End Sub


' ============================================================
' OBTER / CRIAR ABA RESUMO
' ============================================================

Private Function ObterOuCriarResumoDashboard() As Worksheet

    Dim ws As Worksheet


    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_RESUMO_DASH)

    On Error GoTo 0


    If ws Is Nothing Then

        Set ws = _
            ThisWorkbook.Worksheets.Add( _
                After:= _
                ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        ws.Name = ABA_RESUMO_DASH

    End If


    Set ObterOuCriarResumoDashboard = ws

End Function


' ============================================================
' LIMPAR DASHBOARD
' ============================================================

Private Sub LimparDashboard( _
    ByVal ws As Worksheet)

    Dim i As Long


    ws.Range("A1:O13").ClearContents


    ws.Range("A16:C1000").ClearContents


    ws.Range("E16:F100").ClearContents


    For i = ws.Shapes.Count To 1 Step -1

        ws.Shapes(i).Delete

    Next i

End Sub


' ============================================================
' TÍTULO
' ============================================================

Private Sub CriarTituloDashboard( _
    ByVal ws As Worksheet)

    With ws.Range("A1:O2")

        .Merge

        .Value = _
            "APS — PLANEJAMENTO E CONTROLE DA PRODUÇÃO"

        .HorizontalAlignment = xlCenter

        .VerticalAlignment = xlCenter

        .Font.Bold = True

        .Font.Size = 18

    End With


    ws.Range("A3").Value = _
        "Última atualização:"


    ws.Range("B3").Value = Now


    ws.Range("B3").NumberFormat = _
        "dd/mm/yyyy hh:mm"

End Sub


' ============================================================
' CARDS KPI
' ============================================================

Private Sub CriarCardsKPI( _
    ByVal ws As Worksheet)

    CriarKPI _
        ws, _
        "A5:C8", _
        "TOTAL DE OPs", _
        CalcularTotalOPs()


    CriarKPI _
        ws, _
        "D5:F8", _
        "CONCLUÍDAS", _
        CalcularOPsPorStatus("CONCLUÍDO")


    CriarKPI _
        ws, _
        "G5:I8", _
        "EM PRODUÇÃO", _
        CalcularOPsPorStatus("EM PRODUÇÃO")


    CriarKPI _
        ws, _
        "J5:L8", _
        "PLANEJADAS", _
        CalcularOPsPorStatus("PLANEJADO")


    CriarKPI _
        ws, _
        "M5:O8", _
        "ATRASADAS", _
        CalcularOPsPorStatus("ATRASADO")


    CriarKPI _
        ws, _
        "A10:C13", _
        "TOTAL DE CAIXAS", _
        CalcularTotalCaixas()


    CriarKPI _
        ws, _
        "D10:F13", _
        "HORAS PLANEJADAS", _
        CalcularTotalHoras()


    CriarKPI _
        ws, _
        "G10:I13", _
        "HORAS DE ATRASO", _
        CalcularTotalAtrasos()


    CriarKPI _
        ws, _
        "J10:L13", _
        "UTILIZAÇÃO", _
        CalcularUtilizacao()


    CriarKPI _
        ws, _
        "M10:O13", _
        "OPs COM ATRASO", _
        CalcularOPsComAtraso()

End Sub


' ============================================================
' CRIAR KPI
' ============================================================

Private Sub CriarKPI( _
    ByVal ws As Worksheet, _
    ByVal endereco As String, _
    ByVal titulo As String, _
    ByVal valor As Variant)

    Dim r As Range


    Set r = ws.Range(endereco)


    r.Merge


    r.Value = _
        titulo & _
        vbCrLf & _
        FormatValorKPI(valor)


    r.HorizontalAlignment = xlCenter

    r.VerticalAlignment = xlCenter

    r.WrapText = True

    r.Font.Bold = True

    r.Font.Size = 12


    ' --------------------------------------------------------
    ' Borda
    ' --------------------------------------------------------

    With r.Borders

        .LineStyle = xlContinuous

        .Weight = xlThin

    End With

End Sub


' ============================================================
' FORMATAR VALOR KPI
' ============================================================

Private Function FormatValorKPI( _
    ByVal valor As Variant) As String

    If IsNumeric(valor) Then

        FormatValorKPI = _
            Format(CDbl(valor), "#,##0.00")

    Else

        FormatValorKPI = CStr(valor)

    End If

End Function


' ============================================================
' ATUALIZAR KPIs
' ============================================================

Private Sub AtualizarKPIsDashboard( _
    ByVal ws As Worksheet)

    ws.Range("B3").Value = Now

    ws.Range("B3").NumberFormat = _
        "dd/mm/yyyy hh:mm"


    AtualizarValorKPI _
        ws.Range("A5:C8"), _
        "TOTAL DE OPs", _
        CalcularTotalOPs()


    AtualizarValorKPI _
        ws.Range("D5:F8"), _
        "CONCLUÍDAS", _
        CalcularOPsPorStatus("CONCLUÍDO")


    AtualizarValorKPI _
        ws.Range("G5:I8"), _
        "EM PRODUÇÃO", _
        CalcularOPsPorStatus("EM PRODUÇÃO")


    AtualizarValorKPI _
        ws.Range("J5:L8"), _
        "PLANEJADAS", _
        CalcularOPsPorStatus("PLANEJADO")


    AtualizarValorKPI _
        ws.Range("M5:O8"), _
        "ATRASADAS", _
        CalcularOPsPorStatus("ATRASADO")


    AtualizarValorKPI _
        ws.Range("A10:C13"), _
        "TOTAL DE CAIXAS", _
        CalcularTotalCaixas()


    AtualizarValorKPI _
        ws.Range("D10:F13"), _
        "HORAS PLANEJADAS", _
        CalcularTotalHoras()


    AtualizarValorKPI _
        ws.Range("G10:I13"), _
        "HORAS DE ATRASO", _
        CalcularTotalAtrasos()


    AtualizarValorKPI _
        ws.Range("J10:L13"), _
        "UTILIZAÇÃO", _
        CalcularUtilizacao(), _
        True


    AtualizarValorKPI _
        ws.Range("M10:O13"), _
        "OPs COM ATRASO", _
        CalcularOPsComAtraso()

End Sub


' ============================================================
' ATUALIZAR UM KPI
' ============================================================

Private Sub AtualizarValorKPI( _
    ByVal r As Range, _
    ByVal titulo As String, _
    ByVal valor As Variant, _
    Optional ByVal percentual As Boolean = False)

    If percentual Then

        r.Value = _
            titulo & _
            vbCrLf & _
            Format(CDbl(valor), "0.0%")

    Else

        r.Value = _
            titulo & _
            vbCrLf & _
            FormatValorKPI(valor)

    End If

End Sub


' ============================================================
' TOTAL DE OPs
' ============================================================

Private Function CalcularTotalOPs() As Long

    Dim ws As Worksheet

    Dim cOP As Long

    Dim ultimaLinha As Long


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS_DASH)


    cOP = _
        EncontrarColunaDashboard( _
            ws, _
            "OP")


    If cOP = 0 Then Exit Function


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            cOP).End(xlUp).Row


    If ultimaLinha < 2 Then

        CalcularTotalOPs = 0

    Else

        CalcularTotalOPs = _
            Application.WorksheetFunction.CountA( _
                ws.Range( _
                    ws.Cells(2, cOP), _
                    ws.Cells(ultimaLinha, cOP)))

    End If

End Function


' ============================================================
' OPs POR STATUS
' ============================================================

Private Function CalcularOPsPorStatus( _
    ByVal statusProcurado As String) As Long

    Dim ws As Worksheet

    Dim cStatus As Long

    Dim ultimaLinha As Long

    Dim linha As Long

    Dim statusAtual As String


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS_DASH)


    cStatus = _
        EncontrarColunaDashboard( _
            ws, _
            "Status")


    If cStatus = 0 Then Exit Function


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            cStatus).End(xlUp).Row


    For linha = 2 To ultimaLinha

        statusAtual = _
            UCase(Trim(CStr( _
                ws.Cells(linha, cStatus).Value)))


        If statusAtual = _
           UCase(statusProcurado) Then

            CalcularOPsPorStatus = _
                CalcularOPsPorStatus + 1

        End If

    Next linha

End Function


' ============================================================
' TOTAL DE CAIXAS
' ============================================================

Private Function CalcularTotalCaixas() As Double

    Dim ws As Worksheet

    Dim c As Long

    Dim ultimaLinha As Long


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS_DASH)


    c = _
        EncontrarColunaDashboard( _
            ws, _
            "Caixas")


    If c = 0 Then Exit Function


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            c).End(xlUp).Row


    If ultimaLinha < 2 Then Exit Function


    CalcularTotalCaixas = _
        Application.WorksheetFunction.Sum( _
            ws.Range( _
                ws.Cells(2, c), _
                ws.Cells(ultimaLinha, c)))

End Function


' ============================================================
' TOTAL DE HORAS
' ============================================================

Private Function CalcularTotalHoras() As Double

    Dim ws As Worksheet

    Dim c As Long

    Dim ultimaLinha As Long


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS_DASH)


    c = _
        EncontrarColunaDashboard( _
            ws, _
            "Duração Total (h)")


    If c = 0 Then Exit Function


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            c).End(xlUp).Row


    If ultimaLinha < 2 Then Exit Function


    CalcularTotalHoras = _
        Application.WorksheetFunction.Sum( _
            ws.Range( _
                ws.Cells(2, c), _
                ws.Cells(ultimaLinha, c)))

End Function


' ============================================================
' TOTAL DE ATRASOS
' ============================================================

Private Function CalcularTotalAtrasos() As Double

    Dim ws As Worksheet

    Dim c As Long

    Dim ultimaLinha As Long


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS_DASH)


    c = _
        EncontrarColunaDashboard( _
            ws, _
            "Atraso_h")


    If c = 0 Then Exit Function


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            c).End(xlUp).Row


    If ultimaLinha < 2 Then Exit Function


    CalcularTotalAtrasos = _
        Application.WorksheetFunction.Sum( _
            ws.Range( _
                ws.Cells(2, c), _
                ws.Cells(ultimaLinha, c)))

End Function


' ============================================================
' OPs COM ATRASO
' ============================================================

Private Function CalcularOPsComAtraso() As Long

    Dim ws As Worksheet

    Dim c As Long

    Dim ultimaLinha As Long

    Dim linha As Long


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS_DASH)


    c = _
        EncontrarColunaDashboard( _
            ws, _
            "Atraso_h")


    If c = 0 Then Exit Function


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            c).End(xlUp).Row


    For linha = 2 To ultimaLinha

        If IsNumeric(ws.Cells(linha, c).Value) Then

            If CDbl(ws.Cells(linha, c).Value) > 0 Then

                CalcularOPsComAtraso = _
                    CalcularOPsComAtraso + 1

            End If

        End If

    Next linha

End Function


' ============================================================
' UTILIZAÇÃO
'
' Utilização = horas ocupadas / horas disponíveis
'
' Se existir "Capacidade_h", utiliza essa coluna.
'
' Caso não exista, utiliza "Horas Disponíveis".
'
' ============================================================

Private Function CalcularUtilizacao() As Double

    Dim ws As Worksheet

    Dim cHoras As Long

    Dim cCapacidade As Long

    Dim cDisponivel As Long

    Dim ultimaLinha As Long

    Dim linha As Long

    Dim horas As Double

    Dim capacidade As Double


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS_DASH)


    cHoras = _
        EncontrarColunaDashboard( _
            ws, _
            "Duração Total (h)")


    cCapacidade = _
        EncontrarColunaDashboard( _
            ws, _
            "Capacidade_h")


    cDisponivel = _
        EncontrarColunaDashboard( _
            ws, _
            "Horas Disponíveis")


    If cHoras = 0 Then Exit Function


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            cHoras).End(xlUp).Row


    For linha = 2 To ultimaLinha

        If IsNumeric(ws.Cells(linha, cHoras).Value) Then

            horas = _
                horas + _
                CDbl(ws.Cells(linha, cHoras).Value)

        End If


        If cCapacidade > 0 Then

            If IsNumeric( _
                ws.Cells(linha, cCapacidade).Value) Then

                capacidade = _
                    capacidade + _
                    CDbl(ws.Cells(linha, cCapacidade).Value)

            End If


        ElseIf cDisponivel > 0 Then

            If IsNumeric( _
                ws.Cells(linha, cDisponivel).Value) Then

                capacidade = _
                    capacidade + _
                    CDbl(ws.Cells(linha, cDisponivel).Value)

            End If

        End If

    Next linha


    If capacidade > 0 Then

        CalcularUtilizacao = _
            horas / capacidade

    Else

        CalcularUtilizacao = 0

    End If

End Function


' ============================================================
' TABELA DE MÁQUINAS
' ============================================================

Private Sub CriarTabelaMaquinas( _
    ByVal ws As Worksheet)

    Dim dados As Worksheet

    Dim cMaquina As Long
    Dim cOP As Long
    Dim cHoras As Long

    Dim ultimaLinha As Long
    Dim linha As Long
    Dim destino As Long

    Dim maquina As String

    Dim dictOP As Object
    Dim dictHoras As Object

    Dim chave As Variant


    Set dados = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS_DASH)


    cMaquina = _
        EncontrarColunaDashboard( _
            dados, _
            "Máquina")


    cOP = _
        EncontrarColunaDashboard( _
            dados, _
            "OP")


    cHoras = _
        EncontrarColunaDashboard( _
            dados, _
            "Duração Total (h)")


    If cMaquina = 0 Then Exit Sub


    Set dictOP = _
        CreateObject("Scripting.Dictionary")


    Set dictHoras = _
        CreateObject("Scripting.Dictionary")


    ultimaLinha = _
        dados.Cells( _
            dados.Rows.Count, _
            cMaquina).End(xlUp).Row


    For linha = 2 To ultimaLinha

        maquina = _
            Trim(CStr( _
                dados.Cells( _
                    linha, _
                    cMaquina).Value))


        If maquina <> "" Then


            If Not dictOP.Exists(maquina) Then

                dictOP.Add maquina, 0

            End If


            dictOP(maquina) = _
                dictOP(maquina) + 1


            If Not dictHoras.Exists(maquina) Then

                dictHoras.Add maquina, 0

            End If


            If cHoras > 0 Then

                If IsNumeric( _
                    dados.Cells( _
                        linha, _
                        cHoras).Value) Then

                    dictHoras(maquina) = _
                        dictHoras(maquina) + _
                        CDbl( _
                            dados.Cells( _
                                linha, _
                                cHoras).Value)

                End If

            End If

        End If

    Next linha


    ws.Range("A16:C1000").ClearContents


    ws.Range("A16").Value = _
        "DESEMPENHO POR MÁQUINA"


    ws.Range("A17").Value = _
        "MÁQUINA"


    ws.Range("B17").Value = _
        "Nº DE OPs"


    ws.Range("C17").Value = _
        "HORAS"


    destino = 18


    For Each chave In dictOP.Keys

        ws.Cells(destino, 1).Value = chave

        ws.Cells(destino, 2).Value = _
            dictOP(chave)

        ws.Cells(destino, 3).Value = _
            dictHoras(chave)

        destino = destino + 1

    Next chave


    If destino > 18 Then

        ws.Range( _
            "C18:C" & destino - 1).NumberFormat = _
            "0.00"

    End If


    ws.Range( _
        "A17:C" & destino - 1).Borders.LineStyle = _
        xlContinuous

End Sub


' ============================================================
' ATUALIZAR TABELA DE MÁQUINAS
' ============================================================

Private Sub AtualizarTabelaMaquinasDashboard( _
    ByVal ws As Worksheet)

    CriarTabelaMaquinas ws

End Sub


' ============================================================
' TABELA DE STATUS
' ============================================================

Private Sub CriarTabelaStatus( _
    ByVal ws As Worksheet)

    ws.Range("E16:F100").ClearContents


    ws.Range("E16").Value = _
        "STATUS DAS OPs"


    ws.Range("E17").Value = _
        "STATUS"


    ws.Range("F17").Value = _
        "QUANTIDADE"


    ws.Range("E18").Value = _
        "CONCLUÍDO"


    ws.Range("E19").Value = _
        "EM PRODUÇÃO"


    ws.Range("E20").Value = _
        "PLANEJADO"


    ws.Range("E21").Value = _
        "ATRASADO"


    ws.Range("F18").Value = _
        CalcularOPsPorStatus("CONCLUÍDO")


    ws.Range("F19").Value = _
        CalcularOPsPorStatus("EM PRODUÇÃO")


    ws.Range("F20").Value = _
        CalcularOPsPorStatus("PLANEJADO")


    ws.Range("F21").Value = _
        CalcularOPsPorStatus("ATRASADO")


    ws.Range("E17:F21").Borders.LineStyle = _
        xlContinuous

End Sub


' ============================================================
' ATUALIZAR STATUS
' ============================================================

Private Sub AtualizarTabelaStatusDashboard( _
    ByVal ws As Worksheet)

    CriarTabelaStatus ws

End Sub


' ============================================================
' GRÁFICO DE STATUS
' ============================================================

Private Sub CriarGraficoStatus( _
    ByVal ws As Worksheet)

    Dim chartObj As ChartObject


    ApagarGrafico _
        ws, _
        "APS_GRAFICO_STATUS"


    Set chartObj = _
        ws.ChartObjects.Add( _
            Left:=ws.Range("H16").Left, _
            Top:=ws.Range("H16").Top, _
            Width:=330, _
            Height:=220)


    chartObj.Name = _
        "APS_GRAFICO_STATUS"


    With chartObj.Chart

        .SetSourceData _
            Source:= _
            ws.Range("E17:F21")

        .ChartType = _
            xlColumnClustered

        .HasTitle = True

        .ChartTitle.Text = _
            "Status das OPs"

        .HasLegend = False

    End With

End Sub


' ============================================================
' GRÁFICO POR MÁQUINA
' ============================================================

Private Sub CriarGraficoMaquinas( _
    ByVal ws As Worksheet)

    Dim chartObj As ChartObject

    Dim ultimaLinha As Long


    ApagarGrafico _
        ws, _
        "APS_GRAFICO_MAQUINAS"


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).Row


    If ultimaLinha < 18 Then Exit Sub


    Set chartObj = _
        ws.ChartObjects.Add( _
            Left:=ws.Range("H30").Left, _
            Top:=ws.Range("H30").Top, _
            Width:=450, _
            Height:=240)


    chartObj.Name = _
        "APS_GRAFICO_MAQUINAS"


    With chartObj.Chart

        .SetSourceData _
            Source:= _
            ws.Range( _
                "A17:C" & ultimaLinha)

        .ChartType = _
            xlColumnClustered

        .HasTitle = True

        .ChartTitle.Text = _
            "Carga por Máquina"

        .HasLegend = True

    End With

End Sub


' ============================================================
' APAGAR GRÁFICO
' ============================================================

Private Sub ApagarGrafico( _
    ByVal ws As Worksheet, _
    ByVal nome As String)

    On Error Resume Next

    ws.ChartObjects(nome).Delete

    On Error GoTo 0

End Sub


' ============================================================
' ATUALIZAR GRÁFICOS
' ============================================================

Private Sub AtualizarGraficosDashboard( _
    ByVal ws As Worksheet)

    ' --------------------------------------------------------
    ' Recria os gráficos.
    '
    ' Isso garante que mudanças na quantidade de máquinas
    ' e nos dados sejam refletidas corretamente.
    ' --------------------------------------------------------

    CriarGraficoStatus ws

    CriarGraficoMaquinas ws

End Sub


' ============================================================
' FORMATAÇÃO
' ============================================================

Private Sub FormatarDashboard( _
    ByVal ws As Worksheet)

    Dim ultimaLinha As Long


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).Row


    ws.Columns("A:O").ColumnWidth = 12


    ws.Columns("A").ColumnWidth = 20


    ws.Columns("E").ColumnWidth = 18


    ws.Rows("1:50").RowHeight = 20


    ws.Range("A1:O50").VerticalAlignment = _
        xlCenter


    ws.Range("A1:O50").HorizontalAlignment = _
        xlCenter


    ws.Range("A16:C16").Font.Bold = True

    ws.Range("E16:F16").Font.Bold = True

    ws.Range("A17:C17").Font.Bold = True

    ws.Range("E17:F17").Font.Bold = True


    ws.Range( _
        "A17:C" & ultimaLinha).Borders.LineStyle = _
        xlContinuous


    ws.Range("E17:F21").Borders.LineStyle = _
        xlContinuous


    ws.Range("C18:C" & ultimaLinha).NumberFormat = _
        "0.00"


    ActiveWindow.Zoom = 85

End Sub


' ============================================================
' ENCONTRAR COLUNA
' ============================================================

Private Function EncontrarColunaDashboard( _
    ByVal ws As Worksheet, _
    ByVal nome As String) As Long

    Dim coluna As Long

    Dim ultimaColuna As Long


    ultimaColuna = _
        ws.Cells( _
            1, _
            ws.Columns.Count).End(xlToLeft).Column


    For coluna = 1 To ultimaColuna

        If StrComp( _
            Trim(CStr( _
                ws.Cells(1, coluna).Value)), _
            Trim(nome), _
            vbTextCompare) = 0 Then

            EncontrarColunaDashboard = coluna

            Exit Function

        End If

    Next coluna

End Function


' ============================================================
' NÚMERO DA COLUNA → LETRA
' ============================================================

Private Function ColunaLetraDashboard( _
    ByVal ws As Worksheet, _
    ByVal numero As Long) As String

    If numero <= 0 Then Exit Function


    ColunaLetraDashboard = _
        Split( _
            ws.Cells(1, numero).Address, _
            "$")(1)

End Function


' ============================================================
' FIM DO MÓDULO 9
' ============================================================