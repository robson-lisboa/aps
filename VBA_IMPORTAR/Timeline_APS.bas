Option Explicit

' ============================================================
' MÓDULO 7 - LINHA DO TEMPO HORIZONTAL APS
' ============================================================
'
' ESTRUTURA:
'
'       MÁQUINA | 05:00 | 06:00 | 07:00 | 08:00 | 09:00 ...
'       -----------------------------------------------------
'       FETTE   |       [       OP001       ]
'       FETTE 2 |                 [ OP002 ]
'       MEDISEAL| [ OP003 ]
'
' HORÁRIO:
'   esquerda -> direita
'
' MÁQUINA:
'   cima -> baixo
'
' CARDS:
'   Shapes flutuantes
'
' ============================================================


Private Const LINHA_TITULO As Long = 1
Private Const LINHA_SUBTITULO As Long = 2
Private Const LINHA_HORAS As Long = 4
Private Const PRIMEIRA_LINHA_MAQUINA As Long = 5

Private Const COLUNA_MAQUINA As Long = 1

Private Const HORA_INICIAL_PADRAO As Double = 5.5
Private Const HORA_FINAL_PADRAO As Double = 22

Private Const INTERVALO_MINUTOS As Long = 30

Private Const LARGURA_HORA As Double = 10

Private Const ALTURA_LINHA As Double = 65


' ============================================================
' CONSTRUIR TIMELINE
' ============================================================

Public Sub ConstruirTimelineAPS()

    On Error GoTo TrataErro

    Application.ScreenUpdating = False

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO)


    LimparTimeline

    CriarCabecalhoTimeline

    CriarEscalaHorario

    CriarLinhasMaquinas

    FormatarTimeline

    CongelarTimeline

    Application.ScreenUpdating = True


    MsgBox _
        "Linha do tempo criada com sucesso.", _
        vbInformation, _
        "APS - Timeline"

    Exit Sub


TrataErro:

    Application.ScreenUpdating = True

    MsgBox _
        "Erro ao construir a linha do tempo:" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "APS - Timeline"

End Sub


' ============================================================
' LIMPAR TIMELINE
'
' NÃO APAGA OS CARDS.
' ============================================================

Private Sub LimparTimeline()

    Dim ws As Worksheet
    Dim ultimaColuna As Long
    Dim ultimaLinha As Long


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO)


    ultimaColuna = _
        ws.Cells( _
            LINHA_HORAS, _
            ws.Columns.Count _
        ).End(xlToLeft).Column


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1 _
        ).End(xlUp).Row


    If ultimaColuna >= 2 Then

        ws.Range( _
            ws.Cells(1, 2), _
            ws.Cells(ultimaLinha, ultimaColuna) _
        ).Clear

    End If


    ws.Columns(1).ColumnWidth = 20

End Sub


' ============================================================
' CABEÇALHO
' ============================================================

Private Sub CriarCabecalhoTimeline()

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO)


    ws.Cells(1, 1).Value = _
        "PLANEJAMENTO APS"


    ws.Cells(2, 1).Value = _
        "LINHA DO TEMPO"


    ws.Cells(3, 1).Value = _
        "MÁQUINA"


    ws.Cells(4, 1).Value = _
        "MÁQUINA"


    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(2, 1).Font.Bold = True
    ws.Cells(3, 1).Font.Bold = True
    ws.Cells(4, 1).Font.Bold = True


    ws.Rows(1).RowHeight = 25
    ws.Rows(2).RowHeight = 20
    ws.Rows(3).RowHeight = 5
    ws.Rows(4).RowHeight = 25

End Sub


' ============================================================
' CRIAR ESCALA DE HORÁRIO
' ============================================================

Private Sub CriarEscalaHorario()

    Dim ws As Worksheet

    Dim coluna As Long

    Dim horario As Double

    Dim horaInicial As Double
    Dim horaFinal As Double

    Dim intervalo As Double


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO)


    horaInicial = _
        HORA_INICIAL_PADRAO


    horaFinal = _
        HORA_FINAL_PADRAO


    intervalo = _
        INTERVALO_MINUTOS / 1440


    coluna = 2

    horario = horaInicial


    Do While horario <= horaFinal


        ws.Cells(4, coluna).Value = _
            horario / 24


        ws.Cells(4, coluna).NumberFormat = _
            "hh:mm"


        ws.Columns(coluna).ColumnWidth = _
            LARGURA_HORA


        coluna = coluna + 1

        horario = _
            horario + intervalo


    Loop


    CriarMarcadoresDeHora ws

End Sub


' ============================================================
' MARCADORES DE HORA
'
' A cada hora fazemos uma divisão visual mais forte.
' ============================================================

Private Sub CriarMarcadoresDeHora( _
    ByVal ws As Worksheet)

    Dim coluna As Long
    Dim ultimaColuna As Long

    Dim horario As Date


    ultimaColuna = _
        ws.Cells( _
            LINHA_HORAS, _
            ws.Columns.Count _
        ).End(xlToLeft).Column


    For coluna = 2 To ultimaColuna


        If IsDate( _
            ws.Cells( _
                LINHA_HORAS, _
                coluna).Value) Then


            horario = _
                CDate( _
                    ws.Cells( _
                        LINHA_HORAS, _
                        coluna).Value)


            If Minute(horario) = 0 Then


                With ws.Range( _
                    ws.Cells(4, coluna), _
                    ws.Cells( _
                        ws.Rows.Count, _
                        coluna))


                    .Borders(xlEdgeLeft).Weight = _
                        xlMedium


                End With


            Else


                With ws.Range( _
                    ws.Cells(4, coluna), _
                    ws.Cells( _
                        ws.Rows.Count, _
                        coluna))


                    .Borders(xlEdgeLeft).Weight = _
                        xlThin


                End With


            End If

        End If

    Next coluna

End Sub


' ============================================================
' CRIAR LINHAS DAS MÁQUINAS
' ============================================================

Private Sub CriarLinhasMaquinas()

    Dim wsPlan As Worksheet
    Dim wsDados As Worksheet

    Dim maquinas As Collection

    Dim ultimaLinha As Long
    Dim linha As Long

    Dim maquina As String

    Dim destino As Long


    Set wsPlan = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO)


    Set wsDados = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS)


    Set maquinas = _
        New Collection


    On Error Resume Next


    ultimaLinha = _
        wsDados.Cells( _
            wsDados.Rows.Count, _
            EncontrarColuna( _
                wsDados, _
                "Máquina") _
        ).End(xlUp).Row


    For linha = 2 To ultimaLinha


        maquina = _
            Trim(CStr( _
                wsDados.Cells( _
                    linha, _
                    EncontrarColuna( _
                        wsDados, _
                        "Máquina")) _
                    .Value))


        If maquina <> "" Then


            maquinas.Add _
                maquina, _
                UCase(maquina)


        End If

    Next linha


    On Error GoTo 0


    destino = PRIMEIRA_LINHA_MAQUINA


    Dim item As Variant


    For Each item In maquinas


        wsPlan.Cells( _
            destino, _
            COLUNA_MAQUINA).Value = item


        destino = destino + 1


    Next item


    ' --------------------------------------------------------
    ' Se não houver máquinas nos dados,
    ' cria as máquinas da aba RECURSOS.
    ' --------------------------------------------------------

    If destino = PRIMEIRA_LINHA_MAQUINA Then

        CriarMaquinasDosRecursos _
            wsPlan, _
            destino

    End If

End Sub


' ============================================================
' MÁQUINAS DA ABA RECURSOS
' ============================================================

Private Sub CriarMaquinasDosRecursos( _
    ByVal wsPlan As Worksheet, _
    ByVal destinoInicial As Long)

    Dim wsRec As Worksheet

    Dim coluna As Long
    Dim ultimaLinha As Long

    Dim linha As Long
    Dim destino As Long

    Dim maquina As String


    On Error Resume Next

    Set wsRec = _
        ThisWorkbook.Worksheets( _
            "RECURSOS")

    On Error GoTo 0


    If wsRec Is Nothing Then Exit Sub


    coluna = _
        EncontrarColuna( _
            wsRec, _
            "Máquina")


    If coluna = 0 Then Exit Sub


    ultimaLinha = _
        wsRec.Cells( _
            wsRec.Rows.Count, _
            coluna).End(xlUp).Row


    destino = destinoInicial


    For linha = 2 To ultimaLinha


        maquina = _
            Trim(CStr( _
                wsRec.Cells( _
                    linha, _
                    coluna).Value))


        If maquina <> "" Then


            wsPlan.Cells( _
                destino, _
                1).Value = maquina


            destino = destino + 1


        End If

    Next linha

End Sub


' ============================================================
' FORMATAÇÃO
' ============================================================

Private Sub FormatarTimeline()

    Dim ws As Worksheet

    Dim ultimaColuna As Long
    Dim ultimaLinha As Long

    Dim coluna As Long
    Dim linha As Long


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO)


    ultimaColuna = _
        ws.Cells( _
            LINHA_HORAS, _
            ws.Columns.Count _
        ).End(xlToLeft).Column


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1 _
        ).End(xlUp).Row


    ' --------------------------------------------------------
    ' Cabeçalho de horários
    ' --------------------------------------------------------

    With ws.Range( _
        ws.Cells(4, 1), _
        ws.Cells(4, ultimaColuna))

        .Font.Bold = True

        .HorizontalAlignment = _
            xlCenter

        .VerticalAlignment = _
            xlCenter

        .Borders.LineStyle = _
            xlContinuous

    End With


    ' --------------------------------------------------------
    ' Máquinas
    ' --------------------------------------------------------

    With ws.Range( _
        ws.Cells(5, 1), _
        ws.Cells(ultimaLinha, 1))

        .Font.Bold = True

        .VerticalAlignment = _
            xlCenter

        .Borders.LineStyle = _
            xlContinuous

    End With


    ' --------------------------------------------------------
    ' Área da timeline
    ' --------------------------------------------------------

    With ws.Range( _
        ws.Cells(5, 2), _
        ws.Cells(ultimaLinha, ultimaColuna))

        .Borders.LineStyle = _
            xlContinuous

        .Borders.Weight = _
            xlHairline

        .VerticalAlignment = _
            xlCenter

    End With


    ' --------------------------------------------------------
    ' Altura das linhas
    ' --------------------------------------------------------

    For linha = 5 To ultimaLinha

        ws.Rows(linha).RowHeight = _
            ALTURA_LINHA

    Next linha


    ' --------------------------------------------------------
    ' Largura das colunas
    ' --------------------------------------------------------

    For coluna = 2 To ultimaColuna

        ws.Columns(coluna).ColumnWidth = _
            LARGURA_HORA

    Next coluna


    ws.Columns(1).ColumnWidth = 20


    ' --------------------------------------------------------
    ' Zoom
    ' --------------------------------------------------------

    ActiveWindow.Zoom = 85

End Sub


' ============================================================
' CONGELAR CABEÇALHO E MÁQUINAS
' ============================================================

Private Sub CongelarTimeline()

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO)


    ws.Activate


    ActiveWindow.FreezePanes = False


    ws.Range("B5").Select


    ActiveWindow.FreezePanes = True

End Sub


' ============================================================
' CALCULAR POSIÇÃO X
'
' USADO TAMBÉM PELOS CARDS.
'
' ============================================================

Private Function TimelineX( _
    ByVal horario As Date) As Double

    Dim ws As Worksheet

    Dim coluna As Long
    Dim ultimaColuna As Long

    Dim h1 As Date
    Dim h2 As Date

    Dim fracao As Double


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO)


    ultimaColuna = _
        ws.Cells( _
            LINHA_HORAS, _
            ws.Columns.Count _
        ).End(xlToLeft).Column


    For coluna = 2 To ultimaColuna - 1


        If IsDate( _
            ws.Cells(4, coluna).Value) _
           And _
           IsDate( _
            ws.Cells(4, coluna + 1).Value) Then


            h1 = _
                CDate( _
                    ws.Cells(4, coluna).Value)


            h2 = _
                CDate( _
                    ws.Cells(4, coluna + 1).Value)


            If horario >= h1 _
               And horario < h2 Then


                fracao = _
                    (horario - h1) _
                    / _
                    (h2 - h1)


                TimelineX = _
                    ws.Cells(4, coluna).Left _
                    + _
                    ws.Columns(coluna).Width _
                    * fracao


                Exit Function

            End If

        End If

    Next coluna


    ' --------------------------------------------------------
    ' Antes do início da timeline
    ' --------------------------------------------------------

    If horario < _
       CDate( _
           ws.Cells(4, 2).Value) Then


        TimelineX = _
            ws.Cells(4, 2).Left


        Exit Function

    End If


    ' --------------------------------------------------------
    ' Depois do fim da timeline
    ' --------------------------------------------------------

    TimelineX = _
        ws.Cells( _
            4, _
            ultimaColuna).Left _
        + _
        ws.Columns( _
            ultimaColuna).Width

End Function


' ============================================================
' CALCULAR LARGURA BASEADA NO TEMPO
' ============================================================

Private Function TimelineLargura( _
    ByVal inicio As Date, _
    ByVal fim As Date) As Double

    Dim x1 As Double
    Dim x2 As Double


    x1 = TimelineX(inicio)

    x2 = TimelineX(fim)


    TimelineLargura = _
        x2 - x1


    If TimelineLargura < 20 Then

        TimelineLargura = 20

    End If

End Function


' ============================================================
' CALCULAR Y DA MÁQUINA
' ============================================================

Private Function TimelineY( _
    ByVal maquina As String) As Double

    Dim ws As Worksheet

    Dim linha As Long

    Dim ultimaLinha As Long


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO)


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).Row


    For linha = 5 To ultimaLinha


        If StrComp( _
            Trim(CStr( _
                ws.Cells(linha, 1).Value)), _
            Trim(maquina), _
            vbTextCompare) = 0 Then


            TimelineY = _
                ws.Rows(linha).Top + 3


            Exit Function

        End If

    Next linha


    TimelineY = _
        ws.Rows(5).Top + 3

End Function


' ============================================================
' ENCONTRAR COLUNA
' ============================================================

Private Function EncontrarColuna( _
    ByVal ws As Worksheet, _
    ByVal nome As String) As Long

    Dim coluna As Long

    Dim ultimaColuna As Long


    ultimaColuna = _
        ws.Cells( _
            1, _
            ws.Columns.Count _
        ).End(xlToLeft).Column


    For coluna = 1 To ultimaColuna


        If StrComp( _
            Trim(CStr( _
                ws.Cells(1, coluna).Value)), _
            Trim(nome), _
            vbTextCompare) = 0 Then


            EncontrarColuna = coluna

            Exit Function

        End If

    Next coluna

End Function


' ============================================================
' FIM DO MÓDULO 7
' ============================================================