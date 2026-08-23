Option Explicit

' ============================================================
' MÓDULO 5 - ATRASOS / IMPACTO VISUAL
' ============================================================
'
' OBJETIVO:
'
'   • Mostrar o atraso dentro do card.
'   • Mostrar quanto tempo foi acrescentado.
'   • Aumentar visualmente o card.
'   • Manter o horário original identificável.
'   • Mostrar o novo horário.
'   • Não mexer na refeição.
'
' Exemplo:
'
'   OP001
'   PURAN - 25 mcg
'
'   ORIGINAL 05:30 → 10:00
'   ATRASO +02:00
'   NOVO 07:30 → 12:00
'
' ============================================================


Public Const COLUNA_ATRASO_H As String = "Atraso_h"
Public Const COLUNA_ATRASO As String = "Atraso"
Public Const COLUNA_FIM_ORIGINAL As String = "Fim Original"
Public Const COLUNA_INICIO_ORIGINAL As String = "Início Original"
Public Const COLUNA_DURACAO_BASE As String = "Duração Base_h"


' ============================================================
' MENU PRINCIPAL
' ============================================================

Public Sub AtualizarAtrasosAPS()

    On Error GoTo TrataErro

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    PrepararColunasAtraso

    RegistrarHorariosOriginais

    AplicarAtrasos

    RecalcularComAtrasos

    AtualizarCardsAtraso

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "Atrasos atualizados." & vbCrLf & vbCrLf & _
        "Os cards foram recalculados e os tempos adicionais foram incorporados.", _
        vbInformation, _
        "APS - Atrasos"

    Exit Sub

TrataErro:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "Erro no módulo de atrasos:" & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "APS - Atrasos"

End Sub


' ============================================================
' PREPARAR COLUNAS
' ============================================================

Private Sub PrepararColunasAtraso()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("DADOS")

    GarantirColunaAtraso ws, COLUNA_ATRASO_H
    GarantirColunaAtraso ws, COLUNA_ATRASO
    GarantirColunaAtraso ws, COLUNA_FIM_ORIGINAL
    GarantirColunaAtraso ws, COLUNA_INICIO_ORIGINAL
    GarantirColunaAtraso ws, COLUNA_DURACAO_BASE

End Sub


' ============================================================
' REGISTRAR HORÁRIO ORIGINAL
' ============================================================

Private Sub RegistrarHorariosOriginais()

    Dim ws As Worksheet

    Dim ultimaLinha As Long
    Dim linha As Long

    Dim cInicio As Long
    Dim cFim As Long

    Dim cInicioOriginal As Long
    Dim cFimOriginal As Long

    Dim cDuracao As Long
    Dim cDuracaoBase As Long

    Set ws = ThisWorkbook.Worksheets("DADOS")

    cInicio = EncontrarColunaAtraso(ws, "Início")
    cFim = EncontrarColunaAtraso(ws, "Fim")

    cInicioOriginal = _
        EncontrarColunaAtraso(ws, COLUNA_INICIO_ORIGINAL)

    cFimOriginal = _
        EncontrarColunaAtraso(ws, COLUNA_FIM_ORIGINAL)

    cDuracao = _
        EncontrarColunaAtraso(ws, "Duração Total (h)")

    cDuracaoBase = _
        EncontrarColunaAtraso(ws, COLUNA_DURACAO_BASE)

    If cInicio = 0 Or cFim = 0 Then Exit Sub

    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1 _
        ).End(xlUp).Row

    For linha = 2 To ultimaLinha

        If ws.Cells(linha, cInicioOriginal).Value = "" Then

            If IsDate(ws.Cells(linha, cInicio).Value) Then

                ws.Cells(linha, cInicioOriginal).Value = _
                    ws.Cells(linha, cInicio).Value

            End If

            If IsDate(ws.Cells(linha, cFim).Value) Then

                ws.Cells(linha, cFimOriginal).Value = _
                    ws.Cells(linha, cFim).Value

            End If

        End If

        ' ----------------------------------------------------
        ' REGISTRAR DURAÇÃO BASE
        ' ----------------------------------------------------

        If cDuracao > 0 And cDuracaoBase > 0 Then

            If Not IsNumeric(ws.Cells(linha, cDuracaoBase).Value) _
               Or ws.Cells(linha, cDuracaoBase).Value = 0 Then

                If IsNumeric(ws.Cells(linha, cDuracao).Value) Then

                    ws.Cells(linha, cDuracaoBase).Value = _
                        ws.Cells(linha, cDuracao).Value

                End If

            End If

        End If

    Next linha

    ws.Columns(cInicioOriginal).NumberFormat = _
        "dd/mm/yyyy hh:mm"

    ws.Columns(cFimOriginal).NumberFormat = _
        "dd/mm/yyyy hh:mm"

End Sub


' ============================================================
' APLICAR ATRASOS
' ============================================================

Private Sub AplicarAtrasos()

    Dim ws As Worksheet

    Dim ultimaLinha As Long
    Dim linha As Long

    Dim cAtraso As Long
    Dim cDuracao As Long
    Dim cDuracaoBase As Long

    Set ws = ThisWorkbook.Worksheets("DADOS")

    cAtraso = _
        EncontrarColunaAtraso( _
            ws, _
            COLUNA_ATRASO_H)

    cDuracao = _
        EncontrarColunaAtraso( _
            ws, _
            "Duração Total (h)")

    cDuracaoBase = _
        EncontrarColunaAtraso( _
            ws, _
            COLUNA_DURACAO_BASE)

    If cAtraso = 0 _
       Or cDuracao = 0 _
       Or cDuracaoBase = 0 Then Exit Sub

    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1 _
        ).End(xlUp).Row

    For linha = 2 To ultimaLinha

        If IsNumeric( _
            ws.Cells(linha, cAtraso).Value) Then

            ' ------------------------------------------------
            ' A DURAÇÃO É RECALCULADA PELO MOTOR CENTRAL
            '
            ' Duração Total =
            ' Duração Base + Atraso + Eventos
            ' ------------------------------------------------

            RecalcularDuracaoOP ws, linha

        End If

    Next linha

End Sub


' ============================================================
' RECALCULAR COM ATRASOS
' ============================================================

Private Sub RecalcularComAtrasos()

    Dim ws As Worksheet

    Dim ultimaLinha As Long
    Dim linha As Long

    Dim cOP As Long
    Dim cMaquina As Long
    Dim cSequencia As Long

    Dim cInicioOriginal As Long
    Dim cInicio As Long
    Dim cFim As Long
    Dim cDuracao As Long
    Dim cAtraso As Long

    Dim maquinaAnterior As String
    Dim fimAnterior As Date

    Dim maquina As String

    Dim inicio As Date
    Dim fim As Date

    Dim duracao As Double

    Set ws = ThisWorkbook.Worksheets("DADOS")

    cOP = EncontrarColunaAtraso(ws, "OP")
    cMaquina = EncontrarColunaAtraso(ws, "Máquina")
    cSequencia = EncontrarColunaAtraso(ws, "Sequência")

    cInicioOriginal = _
        EncontrarColunaAtraso( _
            ws, _
            COLUNA_INICIO_ORIGINAL)

    cInicio = _
        EncontrarColunaAtraso( _
            ws, _
            "Início")

    cFim = _
        EncontrarColunaAtraso( _
            ws, _
            "Fim")

    cDuracao = _
        EncontrarColunaAtraso( _
            ws, _
            "Duração Total (h)")

    cAtraso = _
        EncontrarColunaAtraso( _
            ws, _
            COLUNA_ATRASO_H)

    If cOP = 0 _
       Or cMaquina = 0 _
       Or cInicio = 0 _
       Or cFim = 0 _
       Or cDuracao = 0 _
       Or cInicioOriginal = 0 Then Exit Sub

    OrdenarAtrasos ws

    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            cOP _
        ).End(xlUp).Row

    maquinaAnterior = ""
    fimAnterior = 0

    For linha = 2 To ultimaLinha

        If Trim(CStr( _
            ws.Cells(linha, cOP).Value)) <> "" Then

            maquina = _
                Trim(CStr( _
                    ws.Cells( _
                        linha, _
                        cMaquina _
                    ).Value))

            duracao = _
                NzAtraso( _
                    ws.Cells( _
                        linha, _
                        cDuracao _
                    ).Value)

            ' ------------------------------------------------
            ' PRIMEIRA OP DA MÁQUINA
            ' ------------------------------------------------

            If StrComp( _
                maquina, _
                maquinaAnterior, _
                vbTextCompare) <> 0 Then

                If IsDate( _
                    ws.Cells( _
                        linha, _
                        cInicioOriginal _
                    ).Value) Then

                    inicio = _
                        CDate( _
                            ws.Cells( _
                                linha, _
                                cInicioOriginal _
                            ).Value)

                Else

                    inicio = _
                        NzDataInicioAtraso( _
                            ws, _
                            linha)

                End If

            Else

                ' ------------------------------------------------
                ' OP SEGUINTE
                ' ------------------------------------------------

                inicio = fimAnterior

            End If

            ' ------------------------------------------------
            ' FIM
            ' ------------------------------------------------

            fim = _
                inicio + _
                duracao / 24

            ' ------------------------------------------------
            ' ESCREVER NOVO HORÁRIO
            ' ------------------------------------------------

            ws.Cells( _
                linha, _
                cInicio _
            ).Value = inicio

            ws.Cells( _
                linha, _
                cFim _
            ).Value = fim

            ws.Cells( _
                linha, _
                cInicio _
            ).NumberFormat = _
                "dd/mm/yyyy hh:mm"

            ws.Cells( _
                linha, _
                cFim _
            ).NumberFormat = _
                "dd/mm/yyyy hh:mm"

            ' ------------------------------------------------
            ' STATUS
            ' ------------------------------------------------

            AtualizarStatusAtraso _
                ws, _
                linha, _
                inicio, _
                fim

            maquinaAnterior = maquina
            fimAnterior = fim

        End If

    Next linha

End Sub


' ============================================================
' DATA DE INÍCIO SEGURA
' ============================================================

Private Function NzDataInicioAtraso( _
    ByVal ws As Worksheet, _
    ByVal linha As Long) As Date

    Dim c As Long

    c = EncontrarColunaAtraso(ws, "Início")

    If c > 0 Then

        If IsDate(ws.Cells(linha, c).Value) Then

            NzDataInicioAtraso = _
                CDate(ws.Cells(linha, c).Value)

        End If

    End If

End Function


' ============================================================
' ATUALIZAR CARDS
' ============================================================

Public Sub AtualizarCardsAtraso()

    Dim ws As Worksheet
    Dim wsDados As Worksheet
    Dim i As Long

    Dim shp As Shape

    Dim op As String
    Dim partes As Variant

    Set ws = _
        ThisWorkbook.Worksheets( _
            "PLANEJAMENTO")

    Set wsDados = _
        ThisWorkbook.Worksheets("DADOS")

    For i = ws.Shapes.Count To 1 Step -1

        If Left( _
            ws.Shapes(i).Name, _
            Len("APS_CARD_")) = "APS_CARD_" Then

            ' ------------------------------------------------
            ' CORREÇÃO:
            '
            ' O Módulo 6 grava:
            '
            ' APS|OP|NUMERO_DA_OP|MAQUINA
            '
            ' Portanto:
            '
            ' partes(0) = APS
            ' partes(1) = OP
            ' partes(2) = número da OP
            ' partes(3) = máquina
            ' ------------------------------------------------

            op = ""

            On Error Resume Next

            partes = _
                Split( _
                    ws.Shapes(i).AlternativeText, _
                    "|")

            On Error GoTo 0

            If IsArray(partes) Then

                If UBound(partes) >= 2 Then

                    op = _
                        Trim(CStr(partes(2)))

                End If

            End If

            If op <> "" Then

                If EncontrarLinhaOPAtraso( _
                        wsDados, _
                        op) = 0 Then

                    ws.Shapes(i).Delete

                Else

                    AtualizarCardIndividual op

                End If

            Else

                ws.Shapes(i).Delete

            End If

        End If

    Next i

End Sub


' ============================================================
' ATUALIZAR CARD INDIVIDUAL
' ============================================================

Public Sub AtualizarCardIndividual( _
    ByVal numeroOP As String)

    Dim wsDados As Worksheet
    Dim wsPlan As Worksheet

    Dim linha As Long

    Dim inicio As Date
    Dim fim As Date

    Dim inicioOriginal As Date
    Dim fimOriginal As Date

    Dim maquina As String
    Dim produto As String
    Dim dosagem As String

    Dim atraso As Double

    Dim shp As Shape

    Dim x As Double
    Dim y As Double
    Dim largura As Double

    Dim cAtraso As Long

    Dim texto As String

    Set wsDados = _
        ThisWorkbook.Worksheets("DADOS")

    Set wsPlan = _
        ThisWorkbook.Worksheets("PLANEJAMENTO")

    linha = _
        EncontrarLinhaOPAtraso( _
            wsDados, _
            numeroOP)

    If linha = 0 Then Exit Sub

    maquina = _
        ValorAtraso( _
            wsDados, _
            linha, _
            "Máquina")

    produto = _
        ValorAtraso( _
            wsDados, _
        linha, _
        "Produto")

    dosagem = _
        ValorAtraso( _
            wsDados, _
            linha, _
            "Dosagem")

    inicio = _
        DataAtraso( _
            wsDados, _
            linha, _
            "Início")

    fim = _
        DataAtraso( _
            wsDados, _
            linha, _
            "Fim")

    inicioOriginal = _
        DataAtraso( _
            wsDados, _
            linha, _
            COLUNA_INICIO_ORIGINAL)

    fimOriginal = _
        DataAtraso( _
            wsDados, _
            linha, _
            COLUNA_FIM_ORIGINAL)

    cAtraso = _
        EncontrarColunaAtraso( _
            wsDados, _
            COLUNA_ATRASO_H)

    If cAtraso > 0 Then

        atraso = _
            NzAtraso( _
                wsDados.Cells( _
                    linha, _
                    cAtraso).Value)

    Else

        atraso = 0

    End If

    ' --------------------------------------------------------
    ' ENCONTRAR CARD
    ' --------------------------------------------------------

    Set shp = _
        EncontrarShapeOP( _
            wsPlan, _
            numeroOP)

    If shp Is Nothing Then Exit Sub

    ' --------------------------------------------------------
    ' NOVA POSIÇÃO
    ' --------------------------------------------------------

    x = _
        CalcularXAtraso( _
            wsPlan, _
            inicio)

    y = _
        CalcularYAtraso( _
            wsPlan, _
            maquina)

    largura = _
        CalcularLarguraAtraso( _
            wsPlan, _
            inicio, _
            fim)

    If largura < 35 Then

        largura = 35

    End If

    shp.Left = x
    shp.Top = y
    shp.Width = largura

    ' --------------------------------------------------------
    ' TEXTO
    ' --------------------------------------------------------

    texto = _
        numeroOP & _
        vbCrLf & _
        produto

    If dosagem <> "" Then

        texto = _
            texto & _
            " - " & _
            dosagem

    End If

    If atraso > 0 Then

        texto = _
            texto & _
            vbCrLf & _
            "ORIGINAL " & _
            Format(inicioOriginal, "hh:mm") & _
            " → " & _
            Format(fimOriginal, "hh:mm")

        texto = _
            texto & _
            vbCrLf & _
            "ATRASO +" & _
            FormatHorasAtraso(atraso)

        texto = _
            texto & _
            vbCrLf & _
            "NOVO " & _
            Format(inicio, "hh:mm") & _
            " → " & _
            Format(fim, "hh:mm")

    Else

        texto = _
            texto & _
            vbCrLf & _
            Format(inicio, "hh:mm") & _
            " → " & _
            Format(fim, "hh:mm")

    End If

    shp.TextFrame2.TextRange.Text = texto

    ConfigurarVisualAtraso _
        shp, _
        atraso

End Sub


' ============================================================
' VISUAL DO ATRASO
' ============================================================

Private Sub ConfigurarVisualAtraso( _
    ByVal shp As Shape, _
    ByVal atraso As Double)

    If atraso > 0 Then

        shp.Fill.ForeColor.RGB = _
            RGB(198, 40, 40)

        shp.Line.ForeColor.RGB = _
            RGB(127, 0, 0)

        shp.Line.Weight = 2

        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(255, 255, 255)

    Else

        shp.Fill.ForeColor.RGB = _
            RGB(33, 150, 243)

        shp.Line.ForeColor.RGB = _
            RGB(13, 71, 161)

        shp.Line.Weight = 1.25

        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(255, 255, 255)

    End If

    With shp.TextFrame2

        .VerticalAnchor = _
            msoAnchorMiddle

        .TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

        .MarginLeft = 5
        .MarginRight = 5
        .MarginTop = 3
        .MarginBottom = 3

    End With

    With shp.TextFrame2.TextRange.Font

        .Size = 8
        .Bold = msoTrue

    End With

End Sub


' ============================================================
' FORMATAR HORAS
' ============================================================

Private Function FormatHorasAtraso( _
    ByVal horas As Double) As String

    Dim h As Long
    Dim minutos As Long

    h = Int(horas)

    minutos = _
        Round( _
            (horas - h) * 60, _
            0)

    If minutos >= 60 Then

        h = h + 1
        minutos = 0

    End If

    FormatHorasAtraso = _
        Format(h, "00") & ":" & _
        Format(minutos, "00")

End Function


' ============================================================
' ENCONTRAR SHAPE DA OP
' ============================================================

Private Function EncontrarShapeOP( _
    ByVal ws As Worksheet, _
    ByVal numeroOP As String) As Shape

    Dim i As Long

    Dim info As String
    Dim partes As Variant

    For i = 1 To ws.Shapes.Count

        If Left( _
            ws.Shapes(i).Name, _
            Len("APS_CARD_")) = "APS_CARD_" Then

            info = _
                ws.Shapes(i).AlternativeText

            partes = _
                Split(info, "|")

            ' ------------------------------------------------
            ' Módulo 6:
            '
            ' APS|OP|NUMERO_OP|MAQUINA
            '
            ' Número da OP = partes(2)
            ' ------------------------------------------------

            If UBound(partes) >= 2 Then

                If StrComp( _
                    Trim(CStr(partes(2))), _
                    Trim(numeroOP), _
                    vbTextCompare) = 0 Then

                    Set EncontrarShapeOP = _
                        ws.Shapes(i)

                    Exit Function

                End If

            End If

        End If

    Next i

End Function


' ============================================================
' POSIÇÃO X
' ============================================================

Private Function CalcularXAtraso( _
    ByVal ws As Worksheet, _
    ByVal horario As Date) As Double

    Dim c As Long
    Dim ultimaColuna As Long

    Dim h1 As Date
    Dim h2 As Date

    Dim fracao As Double

    ultimaColuna = _
        ws.Cells( _
            4, _
            ws.Columns.Count _
        ).End(xlToLeft).Column

    For c = 2 To ultimaColuna - 1

        If IsDate(ws.Cells(4, c).Value) _
           And IsDate(ws.Cells(4, c + 1).Value) Then

            h1 = _
                CDate(ws.Cells(4, c).Value)

            h2 = _
                CDate(ws.Cells(4, c + 1).Value)

            If horario >= h1 _
               And horario < h2 Then

                fracao = _
                    (horario - h1) _
                    / _
                    (h2 - h1)

                CalcularXAtraso = _
                    ws.Cells(4, c).Left _
                    + _
                    ws.Columns(c).Width * fracao

                Exit Function

            End If

        End If

    Next c

    CalcularXAtraso = _
        ws.Cells(4, 2).Left

End Function


' ============================================================
' POSIÇÃO Y
' ============================================================

Private Function CalcularYAtraso( _
    ByVal ws As Worksheet, _
    ByVal maquina As String) As Double

    Dim linha As Long

    linha = _
        EncontrarLinhaMaquinaAtraso( _
            ws, _
            maquina)

    If linha = 0 Then

        linha = 5

    End If

    CalcularYAtraso = _
        ws.Rows(linha).Top + 3

End Function


' ============================================================
' LARGURA
' ============================================================

Private Function CalcularLarguraAtraso( _
    ByVal ws As Worksheet, _
    ByVal inicio As Date, _
    ByVal fim As Date) As Double

    Dim x1 As Double
    Dim x2 As Double

    x1 = _
        CalcularXAtraso( _
            ws, _
            inicio)

    x2 = _
        CalcularXAtraso( _
            ws, _
            fim)

    CalcularLarguraAtraso = _
        x2 - x1

End Function


' ============================================================
' ENCONTRAR LINHA DA MÁQUINA
' ============================================================

Private Function EncontrarLinhaMaquinaAtraso( _
    ByVal ws As Worksheet, _
    ByVal maquina As String) As Long

    Dim linha As Long
    Dim ultimaLinha As Long

    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1 _
        ).End(xlUp).Row

    For linha = 5 To ultimaLinha

        If StrComp( _
            Trim(CStr( _
                ws.Cells(linha, 1).Value)), _
            Trim(maquina), _
            vbTextCompare) = 0 Then

            EncontrarLinhaMaquinaAtraso = _
                linha

            Exit Function

        End If

    Next linha

End Function


' ============================================================
' ORDENAR OPs
' ============================================================

Private Sub OrdenarAtrasos( _
    ByVal ws As Worksheet)

    Dim ultimaLinha As Long
    Dim ultimaColuna As Long

    Dim cMaquina As Long
    Dim cSequencia As Long

    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1 _
        ).End(xlUp).Row

    ultimaColuna = _
        ws.Cells( _
            1, _
            ws.Columns.Count _
        ).End(xlToLeft).Column

    cMaquina = _
        EncontrarColunaAtraso( _
            ws, _
            "Máquina")

    cSequencia = _
        EncontrarColunaAtraso( _
            ws, _
            "Sequência")

    If cMaquina = 0 _
       Or cSequencia = 0 Then Exit Sub

    With ws.Sort

        .SortFields.Clear

        .SortFields.Add _
            Key:=ws.Range( _
                ws.Cells(2, cMaquina), _
                ws.Cells(ultimaLinha, cMaquina)), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SortFields.Add _
            Key:=ws.Range( _
                ws.Cells(2, cSequencia), _
                ws.Cells(ultimaLinha, cSequencia)), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SetRange _
            ws.Range( _
                ws.Cells(1, 1), _
                ws.Cells(ultimaLinha, ultimaColuna))

        .Header = xlYes

        .Apply

    End With

End Sub


' ============================================================
' STATUS
' ============================================================

Private Sub AtualizarStatusAtraso( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal inicio As Date, _
    ByVal fim As Date)

    Dim c As Long

    c = _
        EncontrarColunaAtraso( _
            ws, _
            "Status")

    If c = 0 Then Exit Sub

    If fim < Now Then

        ws.Cells(linha, c).Value = _
            "CONCLUÍDO"

    ElseIf inicio <= Now _
       And fim >= Now Then

        ws.Cells(linha, c).Value = _
            "EM PRODUÇÃO"

    Else

        ws.Cells(linha, c).Value = _
            "PLANEJADO"

    End If

End Sub


' ============================================================
' ENCONTRAR COLUNA
' ============================================================

Private Function EncontrarColunaAtraso( _
    ByVal ws As Worksheet, _
    ByVal nome As String) As Long

    Dim c As Long
    Dim ultimaColuna As Long

    ultimaColuna = _
        ws.Cells( _
            1, _
            ws.Columns.Count _
        ).End(xlToLeft).Column

    For c = 1 To ultimaColuna

        If StrComp( _
            Trim(CStr( _
                ws.Cells(1, c).Value)), _
            Trim(nome), _
            vbTextCompare) = 0 Then

            EncontrarColunaAtraso = c

            Exit Function

        End If

    Next c

End Function


' ============================================================
' GARANTIR COLUNA
' ============================================================

Private Sub GarantirColunaAtraso( _
    ByVal ws As Worksheet, _
    ByVal nome As String)

    Dim c As Long

    c = _
        EncontrarColunaAtraso( _
            ws, _
            nome)

    If c = 0 Then

        c = _
            ws.Cells( _
                1, _
                ws.Columns.Count _
            ).End(xlToLeft).Column + 1

        ws.Cells(1, c).Value = nome

    End If

End Sub


' ============================================================
' ENCONTRAR LINHA OP
' ============================================================

Private Function EncontrarLinhaOPAtraso( _
    ByVal ws As Worksheet, _
    ByVal numeroOP As String) As Long

    Dim c As Long
    Dim linha As Long
    Dim ultimaLinha As Long

    c = _
        EncontrarColunaAtraso( _
            ws, _
            "OP")

    If c = 0 Then Exit Function

    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            c _
        ).End(xlUp).Row

    For linha = 2 To ultimaLinha

        If StrComp( _
            Trim(CStr( _
                ws.Cells(linha, c).Value)), _
            Trim(numeroOP), _
            vbTextCompare) = 0 Then

            EncontrarLinhaOPAtraso = _
                linha

            Exit Function

        End If

    Next linha

End Function


' ============================================================
' VALOR TEXTO
' ============================================================

Private Function ValorAtraso( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal campo As String) As String

    Dim c As Long

    c = _
        EncontrarColunaAtraso( _
            ws, _
            campo)

    If c > 0 Then

        ValorAtraso = _
            Trim(CStr( _
                ws.Cells(linha, c).Value))

    End If

End Function


' ============================================================
' VALOR DATA
' ============================================================

Private Function DataAtraso( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal campo As String) As Date

    Dim c As Long

    c = _
        EncontrarColunaAtraso( _
            ws, _
            campo)

    If c > 0 Then

        If IsDate( _
            ws.Cells(linha, c).Value) Then

            DataAtraso = _
                CDate( _
                    ws.Cells(linha, c).Value)

        End If

    End If

End Function


' ============================================================
' NÚMERO SEGURO
' ============================================================

Private Function NzAtraso( _
    ByVal valor As Variant) As Double

    If IsNumeric(valor) Then

        NzAtraso = _
            CDbl(valor)

    Else

        NzAtraso = 0

    End If

End Function


' ============================================================
' FIM DO MÓDULO 5
' ============================================================