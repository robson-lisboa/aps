Option Explicit

' ============================================================
' MÓDULO 4 - REFEIÇÃO FIXA
' ============================================================
'
' REGRA:
'
' 1. A refeição é opcional.
' 2. Possui horário fixo.
' 3. Possui duração configurável.
' 4. Fica visualmente separada dos cards de produção.
' 5. NÃO altera Duração Total (h) da OP.
' 6. NÃO é empurrada por atraso.
' 7. NÃO é empurrada por eventos.
' 8. Permanece no horário configurado.
'
' IMPORTANTE:
'
' A refeição é uma janela fixa de parada visual.
' Ela NÃO participa do cálculo de duração da OP.
'
' ============================================================


Public Const ABA_REFEICAO As String = "REFEICAO"
Public Const ABA_PLANEJAMENTO_REFEICAO As String = "PLANEJAMENTO"

Private Const PREFIXO_REFEICAO As String = "APS_REFEICAO_"

Private Const LINHA_HORA As Long = 4
Private Const LINHA_MAQUINA As Long = 5

Private Const ALTURA_REFEICAO As Double = 18
Private Const LARGURA_MINIMA_REFEICAO As Double = 30


' ============================================================
' CRIAR ESTRUTURA DA ABA REFEIÇÃO
' ============================================================

Public Sub CriarAbaRefeicao()

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(ABA_REFEICAO)

    On Error GoTo 0


    ' --------------------------------------------------------
    ' Criar aba se não existir
    ' --------------------------------------------------------

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
                ThisWorkbook.Worksheets.Count))

        ws.Name = ABA_REFEICAO

    End If


    ' --------------------------------------------------------
    ' Cabeçalhos
    ' --------------------------------------------------------

    ws.Cells(1, 1).Value = "ID"
    ws.Cells(1, 2).Value = "Máquina"
    ws.Cells(1, 3).Value = "Início"
    ws.Cells(1, 4).Value = "Duração_h"
    ws.Cells(1, 5).Value = "Fim"
    ws.Cells(1, 6).Value = "Ativo"
    ws.Cells(1, 7).Value = "Observação"


    ' --------------------------------------------------------
    ' Formatação
    ' --------------------------------------------------------

    With ws.Rows(1)

        .Font.Bold = True

    End With


    ws.Columns("A:G").ColumnWidth = 18


    ws.Columns("C:C").NumberFormat = _
        "dd/mm/yyyy hh:mm"

    ws.Columns("E:E").NumberFormat = _
        "dd/mm/yyyy hh:mm"


    ' --------------------------------------------------------
    ' Exemplo inicial
    '
    ' Só cria se ainda não houver registro.
    ' --------------------------------------------------------

    If Trim(CStr(ws.Cells(2, 1).Value)) = "" Then

        ws.Cells(2, 1).Value = 1

        ws.Cells(2, 2).Value = _
            "FETTE 2090"

        ws.Cells(2, 3).Value = _
            DateSerial(2026, 8, 19) + _
            TimeSerial(12, 0, 0)

        ws.Cells(2, 4).Value = 1

        ws.Cells(2, 6).Value = True

        ws.Cells(2, 7).Value = _
            "Refeição"

    End If


    ' --------------------------------------------------------
    ' Calcular fim
    ' --------------------------------------------------------

    AtualizarFimRefeicoes


End Sub


' ============================================================
' ATUALIZAR FIM DAS REFEIÇÕES
' ============================================================

Public Sub AtualizarFimRefeicoes()

    Dim ws As Worksheet

    Dim ultimaLinha As Long
    Dim linha As Long

    Dim inicio As Date
    Dim duracao As Double


    On Error GoTo TrataErro


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_REFEICAO)


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1 _
        ).End(xlUp).Row


    If ultimaLinha < 2 Then Exit Sub


    For linha = 2 To ultimaLinha


        If IsDate(ws.Cells(linha, 3).Value) _
           And IsNumeric(ws.Cells(linha, 4).Value) Then


            inicio = _
                CDate( _
                    ws.Cells(linha, 3).Value)


            duracao = _
                CDbl( _
                    ws.Cells(linha, 4).Value)


            If duracao >= 0 Then

                ws.Cells(linha, 5).Value = _
                    inicio + _
                    duracao / 24

                ws.Cells(linha, 5).NumberFormat = _
                    "dd/mm/yyyy hh:mm"

            End If

        End If

    Next linha


    Exit Sub


TrataErro:

    MsgBox _
        "Erro ao atualizar fim das refeições:" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "APS - Refeição"

End Sub


' ============================================================
' DESENHAR REFEIÇÕES
' ============================================================

Public Sub DesenharRefeicoes()

    On Error GoTo TrataErro

    Application.ScreenUpdating = False
    Application.EnableEvents = False


    ' --------------------------------------------------------
    ' Garantir estrutura
    ' --------------------------------------------------------

    CriarAbaRefeicao


    ' --------------------------------------------------------
    ' Atualizar horários finais
    ' --------------------------------------------------------

    AtualizarFimRefeicoes


    ' --------------------------------------------------------
    ' Apagar somente cards de refeição
    ' --------------------------------------------------------

    ApagarCardsRefeicao


    ' --------------------------------------------------------
    ' Criar novamente
    ' --------------------------------------------------------

    CriarCardsRefeicao


Saida:

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Exit Sub


TrataErro:

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "Erro ao desenhar refeições:" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "APS - Refeição"

End Sub


' ============================================================
' APAGAR CARDS DE REFEIÇÃO
' ============================================================

Public Sub ApagarCardsRefeicao()

    Dim ws As Worksheet
    Dim i As Long


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO_REFEICAO)


    For i = ws.Shapes.Count To 1 Step -1


        If Left( _
            ws.Shapes(i).Name, _
            Len(PREFIXO_REFEICAO)) = _
            PREFIXO_REFEICAO Then


            ws.Shapes(i).Delete


        End If

    Next i

End Sub


' ============================================================
' CRIAR CARDS DE REFEIÇÃO
' ============================================================

Private Sub CriarCardsRefeicao()

    Dim wsPlan As Worksheet
    Dim wsRef As Worksheet

    Dim ultimaLinha As Long
    Dim linha As Long

    Dim maquina As String

    Dim inicio As Date
    Dim fim As Date

    Dim ativo As Boolean

    Dim x As Double
    Dim largura As Double
    Dim y As Double

    Dim shp As Shape

    Dim id As String


    Set wsPlan = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO_REFEICAO)

    Set wsRef = _
        ThisWorkbook.Worksheets( _
            ABA_REFEICAO)


    ultimaLinha = _
        wsRef.Cells( _
            wsRef.Rows.Count, _
            1 _
        ).End(xlUp).Row


    If ultimaLinha < 2 Then Exit Sub


    For linha = 2 To ultimaLinha


        id = Trim(CStr( _
            wsRef.Cells(linha, 1).Value))


        maquina = Trim(CStr( _
            wsRef.Cells(linha, 2).Value))


        ativo = _
            RefeicaoAtiva( _
                wsRef, _
                linha)


        If ativo _
           And id <> "" _
           And maquina <> "" _
           And IsDate(wsRef.Cells(linha, 3).Value) _
           And IsDate(wsRef.Cells(linha, 5).Value) Then


            inicio = _
                CDate( _
                    wsRef.Cells(linha, 3).Value)


            fim = _
                CDate( _
                    wsRef.Cells(linha, 5).Value)


            ' ------------------------------------------------
            ' Ignorar duração inválida
            ' ------------------------------------------------

            If fim <= inicio Then GoTo ProximaRefeicao


            ' ------------------------------------------------
            ' POSIÇÃO X
            ' ------------------------------------------------

            x = _
                PosicaoXRefeicao( _
                    wsPlan, _
                    inicio)


            ' ------------------------------------------------
            ' LARGURA
            ' ------------------------------------------------

            largura = _
                LarguraRefeicao( _
                    wsPlan, _
                    inicio, _
                    fim)


            If largura < LARGURA_MINIMA_REFEICAO Then

                largura = _
                    LARGURA_MINIMA_REFEICAO

            End If


            ' ------------------------------------------------
            ' POSIÇÃO Y
            ' ------------------------------------------------

            y = _
                PosicaoYRefeicao( _
                    wsPlan, _
                    maquina)


            ' ------------------------------------------------
            ' CRIAR SHAPE
            ' ------------------------------------------------

            Set shp = _
                wsPlan.Shapes.AddShape( _
                    msoShapeRoundedRectangle, _
                    x, _
                    y, _
                    largura, _
                    ALTURA_REFEICAO)


            shp.Name = _
                PREFIXO_REFEICAO & id


            ' ------------------------------------------------
            ' TEXTO
            ' ------------------------------------------------

            shp.TextFrame2.TextRange.Text = _
                "REFEIÇÃO  " & _
                Format(inicio, "hh:mm") & _
                " - " & _
                Format(fim, "hh:mm")


            With shp.TextFrame2

                .VerticalAnchor = _
                    msoAnchorMiddle

                .TextRange.ParagraphFormat.Alignment = _
                    msoAlignCenter

                .MarginLeft = 3
                .MarginRight = 3
                .MarginTop = 1
                .MarginBottom = 1

            End With


            With shp.TextFrame2.TextRange.Font

                .Size = 8
                .Bold = msoTrue

            End With


            ' ------------------------------------------------
            ' VISUAL
            ' ------------------------------------------------

            shp.Fill.ForeColor.RGB = _
                RGB(255, 152, 0)


            shp.Line.ForeColor.RGB = _
                RGB(230, 81, 0)


            shp.Line.Weight = 1


            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                RGB(255, 255, 255)


            ' ------------------------------------------------
            ' REFEIÇÃO É LIVRE
            '
            ' Não acompanha células.
            ' Não é deslocada pelos cards.
            ' ------------------------------------------------

            shp.Placement = xlFreeFloating


            ' ------------------------------------------------
            ' METADADOS
            '
            ' APS | REFEICAO | ID | MAQUINA
            ' ------------------------------------------------

            shp.AlternativeText = _
                "APS|REFEICAO|" & _
                id & "|" & _
                maquina


        End If


ProximaRefeicao:

        Set shp = Nothing

    Next linha

End Sub


' ============================================================
' POSIÇÃO X
' ============================================================

Private Function PosicaoXRefeicao( _
    ByVal ws As Worksheet, _
    ByVal horario As Date) As Double

    Dim ultimaColuna As Long
    Dim coluna As Long

    Dim horaAtual As Date
    Dim proximaHora As Date

    Dim fracao As Double


    If Not IsDate( _
        ws.Cells( _
            LINHA_HORA, _
            2).Value) Then


        PosicaoXRefeicao = _
            ws.Cells( _
                LINHA_HORA, _
                2).Left

        Exit Function

    End If


    ultimaColuna = _
        ws.Cells( _
            LINHA_HORA, _
            ws.Columns.Count _
        ).End(xlToLeft).Column


    If ultimaColuna < 2 Then

        PosicaoXRefeicao = 0

        Exit Function

    End If


    ' --------------------------------------------------------
    ' Antes da primeira hora
    ' --------------------------------------------------------

    If horario <= _
        CDate(ws.Cells(LINHA_HORA, 2).Value) Then


        PosicaoXRefeicao = _
            ws.Cells( _
                LINHA_HORA, _
                2).Left

        Exit Function

    End If


    ' --------------------------------------------------------
    ' Procurar intervalo
    ' --------------------------------------------------------

    For coluna = 2 To ultimaColuna - 1


        If IsDate(ws.Cells( _
            LINHA_HORA, _
            coluna).Value) _
           And IsDate(ws.Cells( _
            LINHA_HORA, _
            coluna + 1).Value) Then


            horaAtual = _
                CDate(ws.Cells( _
                    LINHA_HORA, _
                    coluna).Value)


            proximaHora = _
                CDate(ws.Cells( _
                    LINHA_HORA, _
                    coluna + 1).Value)


            If proximaHora > horaAtual Then


                If horario >= horaAtual _
                   And horario < proximaHora Then


                    fracao = _
                        (horario - horaAtual) _
                        / _
                        (proximaHora - horaAtual)


                    PosicaoXRefeicao = _
                        ws.Cells( _
                            LINHA_HORA, _
                            coluna).Left _
                        + _
                        ws.Columns(coluna).Width _
                        * fracao


                    Exit Function

                End If

            End If

        End If

    Next coluna


    ' --------------------------------------------------------
    ' Depois do final da escala
    ' --------------------------------------------------------

    PosicaoXRefeicao = _
        ws.Cells( _
            LINHA_HORA, _
            ultimaColuna).Left _
        + _
        ws.Columns(ultimaColuna).Width

End Function


' ============================================================
' LARGURA DA REFEIÇÃO
' ============================================================

Private Function LarguraRefeicao( _
    ByVal ws As Worksheet, _
    ByVal inicio As Date, _
    ByVal fim As Date) As Double

    Dim x1 As Double
    Dim x2 As Double


    x1 = _
        PosicaoXRefeicao( _
            ws, _
            inicio)


    x2 = _
        PosicaoXRefeicao( _
            ws, _
            fim)


    LarguraRefeicao = _
        x2 - x1


    If LarguraRefeicao < 0 Then

        LarguraRefeicao = 0

    End If

End Function


' ============================================================
' POSIÇÃO Y
' ============================================================

Private Function PosicaoYRefeicao( _
    ByVal ws As Worksheet, _
    ByVal maquina As String) As Double

    Dim linha As Long


    linha = _
        EncontrarLinhaMaquinaRefeicao( _
            ws, _
            maquina)


    If linha = 0 Then

        linha = LINHA_MAQUINA

    End If


    ' --------------------------------------------------------
    ' Refeição fica na parte inferior da faixa da máquina.
    ' --------------------------------------------------------

    PosicaoYRefeicao = _
        ws.Rows(linha).Top _
        + _
        ws.Rows(linha).Height _
        - _
        ALTURA_REFEICAO _
        - 2


End Function


' ============================================================
' ENCONTRAR LINHA DA MÁQUINA
' ============================================================

Private Function EncontrarLinhaMaquinaRefeicao( _
    ByVal ws As Worksheet, _
    ByVal maquina As String) As Long

    Dim ultimaLinha As Long
    Dim linha As Long


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1 _
        ).End(xlUp).Row


    If ultimaLinha < LINHA_MAQUINA Then Exit Function


    For linha = LINHA_MAQUINA To ultimaLinha


        If StrComp( _
            Trim(CStr( _
                ws.Cells(linha, 1).Value)), _
            Trim(maquina), _
            vbTextCompare) = 0 Then


            EncontrarLinhaMaquinaRefeicao = _
                linha


            Exit Function

        End If

    Next linha

End Function


' ============================================================
' REFEIÇÃO ATIVA
' ============================================================

Private Function RefeicaoAtiva( _
    ByVal ws As Worksheet, _
    ByVal linha As Long) As Boolean

    Dim valor As Variant


    valor = _
        ws.Cells(linha, 6).Value


    ' --------------------------------------------------------
    ' Vazio = ativo
    ' --------------------------------------------------------

    If IsEmpty(valor) _
       Or Trim(CStr(valor)) = "" Then


        RefeicaoAtiva = True

        Exit Function

    End If


    ' --------------------------------------------------------
    ' Booleano
    ' --------------------------------------------------------

    If VarType(valor) = vbBoolean Then

        RefeicaoAtiva = _
            CBool(valor)

        Exit Function

    End If


    ' --------------------------------------------------------
    ' Texto / número
    ' --------------------------------------------------------

    Select Case UCase( _
        Trim(CStr(valor)))


        Case "SIM", _
             "S", _
             "TRUE", _
             "VERDADEIRO", _
             "1"

            RefeicaoAtiva = True


        Case "NÃO", _
             "NAO", _
             "N", _
             "FALSE", _
             "FALSO", _
             "0"

            RefeicaoAtiva = False


        Case Else

            RefeicaoAtiva = False

    End Select

End Function


' ============================================================
' ATUALIZAR UMA REFEIÇÃO ESPECÍFICA
'
' Útil quando apenas uma linha da aba REFEICAO foi alterada.
' ============================================================

Public Sub AtualizarRefeicao(ByVal id As String)

    Dim ws As Worksheet
    Dim linha As Long
    Dim ultimaLinha As Long

    Dim encontrado As Boolean


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_REFEICAO)


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1 _
        ).End(xlUp).Row


    For linha = 2 To ultimaLinha


        If StrComp( _
            Trim(CStr(ws.Cells(linha, 1).Value)), _
            Trim(id), _
            vbTextCompare) = 0 Then


            encontrado = True

            Exit For

        End If

    Next linha


    If Not encontrado Then Exit Sub


    AtualizarFimRefeicoes

    DesenharRefeicoes

End Sub


' ============================================================
' FIM DO MÓDULO 4
' ============================================================