Option Explicit

' ============================================================
' DRAGCARDS_APS
' APS PURAN - ARRASTAR / REPROGRAMAR CARDS
' ============================================================
'
' FLUXO:
'
' 1. Usuário arrasta um card.
'
' 2. Executa:
'
'       AplicarPosicaoDosCards
'
' 3. O sistema lê:
'
'       X        = novo início
'       Top      = nova máquina
'       OP       = identificação da OP
'
' 4. Atualiza DADOS.
'
' 5. Recalcula Sequência por máquina.
'
' 6. Recalcula Início/Fim.
'
' 7. Atualiza Status.
'
' 8. Reconstrói os cards.
'
'
' FORMATO OFICIAL DO ALTERNATETEXT:
'
'       APS|OP|OP001|FETTE
'
'
'       partes(0) = APS
'       partes(1) = OP
'       partes(2) = Número da OP
'       partes(3) = Máquina
'
'
' REGRA DO ARRASTE:
'
'       X     -> novo horário
'       Y/Top -> nova máquina
'
'
' A ORDEM DAS OPs É DEFINIDA PELA POSIÇÃO X.
'
' Portanto, se uma OP intermediária for arrastada para antes
' de outra, a Sequência será alterada.
'
' ============================================================


Private Const LINHA_HORAS_DRAG As Long = 4
Private Const PRIMEIRA_LINHA_MAQUINA_DRAG As Long = 5


' ============================================================
' ESTRUTURA TEMPORÁRIA DOS CARDS
' ============================================================

Private Type TCardDrag

    NomeShape As String
    NumeroOP As String
    Maquina As String

    PosX As Double
    PosTop As Double

    NovoInicio As Date

    LinhaDados As Long
    Sequencia As Long

End Type


' ============================================================
' APLICAR POSIÇÃO DOS CARDS
' ============================================================

Public Sub AplicarPosicaoDosCards()

    Dim wsPlan As Worksheet
    Dim wsDados As Worksheet

    Dim cards() As TCardDrag
    Dim quantidadeCards As Long

    Dim i As Long

    Dim calcAnterior As XlCalculation
    Dim eventosAnterior As Boolean
    Dim telaAnterior As Boolean
    Dim statusAnterior As Variant

    Dim sucesso As Boolean

    On Error GoTo TrataErro


    ' --------------------------------------------------------
    ' PRESERVAR ESTADO DO EXCEL
    ' --------------------------------------------------------

    calcAnterior = Application.Calculation
    eventosAnterior = Application.EnableEvents
    telaAnterior = Application.ScreenUpdating
    statusAnterior = Application.StatusBar


    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    Set wsPlan = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO)

    Set wsDados = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS)


    ' --------------------------------------------------------
    ' LOCALIZAR TODOS OS CARDS
    ' --------------------------------------------------------

    quantidadeCards = _
        LerCardsDaPlanilha( _
            wsPlan, _
            wsDados, _
            cards)


    If quantidadeCards = 0 Then

        MsgBox _
            "Nenhum card APS válido foi encontrado." & _
            vbCrLf & vbCrLf & _
            "Verifique se os cards possuem:" & _
            vbCrLf & _
            "APS|OP|NúmeroOP|Máquina", _
            vbExclamation, _
            "APS PURAN - Reprogramação"

        GoTo SaidaSegura

    End If


    ' --------------------------------------------------------
    ' ATUALIZAR DADOS COM A POSIÇÃO DOS CARDS
    ' --------------------------------------------------------

    Application.StatusBar = _
        "APS PURAN - Aplicando posições dos cards..."


    For i = 1 To quantidadeCards

        If cards(i).LinhaDados > 0 Then

            AtualizarDadosPeloCard _
                wsDados, _
                cards(i)

        End If

    Next i


    ' --------------------------------------------------------
    ' DEFINIR NOVA SEQUÊNCIA
    '
    ' A ordem passa a ser:
    '
    ' Máquina
    ' X
    '
    ' --------------------------------------------------------

    OrdenarCardsPorMaquinaEPosicao _
        cards, _
        quantidadeCards


    AtribuirSequencias _
        wsDados, _
        cards, _
        quantidadeCards


    ' --------------------------------------------------------
    ' RECALCULAR PROGRAMAÇÃO
    ' --------------------------------------------------------

    Application.StatusBar = _
        "APS PURAN - Recalculando programação..."


    RecalcularTodasAsMaquinas


    ' --------------------------------------------------------
    ' RECRIAR CARDS
    ' --------------------------------------------------------

    Application.StatusBar = _
        "APS PURAN - Atualizando cards..."


    AtualizarCardsDepoisDoDrag


    sucesso = True


SaidaSegura:

    ' --------------------------------------------------------
    ' RESTAURAR EXCEL
    ' --------------------------------------------------------

    On Error Resume Next

    Application.Calculation = calcAnterior
    Application.EnableEvents = eventosAnterior
    Application.ScreenUpdating = telaAnterior
    Application.StatusBar = statusAnterior

    On Error GoTo 0


    If sucesso Then

        MsgBox _
            "Posições dos cards aplicadas com sucesso." & _
            vbCrLf & vbCrLf & _
            "A posição horizontal foi convertida em horário." & _
            vbCrLf & _
            "A posição vertical foi convertida em máquina." & _
            vbCrLf & _
            "A sequência das OPs foi recalculada." & _
            vbCrLf & _
            "A programação foi atualizada.", _
            vbInformation, _
            "APS PURAN - Reprogramação"

    End If


    Exit Sub


TrataErro:

    On Error Resume Next

    Application.Calculation = calcAnterior
    Application.EnableEvents = eventosAnterior
    Application.ScreenUpdating = telaAnterior
    Application.StatusBar = statusAnterior

    On Error GoTo 0


    MsgBox _
        "Erro ao aplicar a posição dos cards." & _
        vbCrLf & vbCrLf & _
        "Procedimento: " & Err.Source & _
        vbCrLf & _
        "Erro: " & CStr(Err.Number) & _
        vbCrLf & _
        "Descrição: " & Err.Description, _
        vbCritical, _
        "APS PURAN - Erro"

End Sub


' ============================================================
' LER CARDS DA PLANILHA
' ============================================================

Private Function LerCardsDaPlanilha( _
    ByVal wsPlan As Worksheet, _
    ByVal wsDados As Worksheet, _
    ByRef cards() As TCardDrag) As Long

    Dim i As Long
    Dim contador As Long

    Dim partes As Variant

    Dim op As String
    Dim maquina As String

    Dim horario As Date
    Dim maquinaDetectada As String

    Dim linhaDados As Long


    ReDim cards(1 To 1)


    For i = 1 To wsPlan.Shapes.Count

        If EhCardAPS(wsPlan.Shapes(i)) Then

            partes = _
                Split( _
                    CStr(wsPlan.Shapes(i).AlternativeText), _
                    "|")


            If UBound(partes) >= ALT_IDX_OP Then

                op = _
                    Trim$( _
                        CStr( _
                            partes(ALT_IDX_OP)))


                If op <> "" Then

                    ' ------------------------------------------------
                    ' HORÁRIO PELO X
                    ' ------------------------------------------------

                    horario = _
                        HorarioDaPosicaoX( _
                            wsPlan.Shapes(i).Left)


                    If horario > 0 Then

                        ' --------------------------------------------
                        ' MÁQUINA PELO Y
                        ' --------------------------------------------

                        maquinaDetectada = _
                            MaquinaDaPosicaoY( _
                                wsPlan.Shapes(i).Top, _
                                wsPlan)


                        ' --------------------------------------------
                        ' FALLBACK:
                        ' se não conseguir detectar pela posição,
                        ' usa a máquina existente no AlternativeText.
                        ' --------------------------------------------

                        If maquinaDetectada = "" Then

                            If UBound(partes) >= ALT_IDX_MAQUINA Then

                                maquinaDetectada = _
                                    Trim$( _
                                        CStr( _
                                            partes(ALT_IDX_MAQUINA)))

                            End If

                        End If


                        ' --------------------------------------------
                        ' LOCALIZAR OP EM DADOS
                        ' --------------------------------------------

                        linhaDados = _
                            EncontrarLinhaOPDrag( _
                                wsDados, _
                                op)


                        If linhaDados > 0 Then

                            contador = contador + 1

                            If contador > UBound(cards) Then

                                ReDim Preserve _
                                    cards( _
                                        1 To contador)

                            End If


                            cards(contador).NomeShape = _
                                wsPlan.Shapes(i).Name

                            cards(contador).NumeroOP = _
                                op

                            cards(contador).Maquina = _
                                maquinaDetectada

                            cards(contador).PosX = _
                                wsPlan.Shapes(i).Left

                            cards(contador).PosTop = _
                                wsPlan.Shapes(i).Top

                            cards(contador).NovoInicio = _
                                horario

                            cards(contador).LinhaDados = _
                                linhaDados

                        End If

                    End If

                End If

            End If

        End If

    Next i


    If contador = 0 Then

        Erase cards

    Else

        ReDim Preserve _
            cards( _
                1 To contador)

    End If


    LerCardsDaPlanilha = contador

End Function


' ============================================================
' VERIFICAR SE É CARD APS
' ============================================================

Private Function EhCardAPS( _
    ByVal shp As Shape) As Boolean

    If shp Is Nothing Then Exit Function


    EhCardAPS = _
        (Left$( _
            shp.Name, _
            Len(PREFIXO_CARD)) = _
            PREFIXO_CARD)

End Function


' ============================================================
' CONVERTER X EM HORÁRIO
'
' A timeline trabalha com DATA + HORA completas.
'
' ============================================================

Private Function HorarioDaPosicaoX( _
    ByVal posX As Double) As Date

    Dim ws As Worksheet

    Dim coluna As Long
    Dim ultimaColuna As Long

    Dim x1 As Double
    Dim x2 As Double

    Dim h1 As Date
    Dim h2 As Date

    Dim percentual As Double
    Dim resultado As Date


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO)


    ultimaColuna = _
        ws.Cells( _
            LINHA_HORAS_DRAG, _
            ws.Columns.Count).End(xlToLeft).Column


    If ultimaColuna < 2 Then Exit Function


    ' --------------------------------------------------------
    ' ANTES DA TIMELINE
    ' --------------------------------------------------------

    If posX <= _
       ws.Cells( _
           LINHA_HORAS_DRAG, _
           2).Left Then

        If IsDate( _
            ws.Cells( _
                LINHA_HORAS_DRAG, _
                2).Value) Then

            HorarioDaPosicaoX = _
                CDate( _
                    ws.Cells( _
                        LINHA_HORAS_DRAG, _
                        2).Value)

        End If

        Exit Function

    End If


    ' --------------------------------------------------------
    ' PROCURAR INTERVALO
    ' --------------------------------------------------------

    For coluna = 2 To ultimaColuna - 1

        If IsDate( _
            ws.Cells( _
                LINHA_HORAS_DRAG, _
                coluna).Value) _
           And _
           IsDate( _
            ws.Cells( _
                LINHA_HORAS_DRAG, _
                coluna + 1).Value) Then


            x1 = _
                ws.Cells( _
                    LINHA_HORAS_DRAG, _
                    coluna).Left


            x2 = _
                ws.Cells( _
                    LINHA_HORAS_DRAG, _
                    coluna + 1).Left


            If x2 > x1 Then

                If posX >= x1 _
                   And posX < x2 Then


                    h1 = _
                        CDate( _
                            ws.Cells( _
                                LINHA_HORAS_DRAG, _
                                coluna).Value)


                    h2 = _
                        CDate( _
                            ws.Cells( _
                                LINHA_HORAS_DRAG, _
                                coluna + 1).Value)


                    percentual = _
                        (posX - x1) _
                        / _
                        (x2 - x1)


                    resultado = _
                        h1 + _
                        (h2 - h1) * percentual


                    resultado = _
                        ArredondarHorario( _
                            resultado, _
                            5)


                    HorarioDaPosicaoX = _
                        resultado


                    Exit Function

                End If

            End If

        End If

    Next coluna


    ' --------------------------------------------------------
    ' DEPOIS DA TIMELINE
    ' --------------------------------------------------------

    If posX >= _
       ws.Cells( _
           LINHA_HORAS_DRAG, _
           ultimaColuna).Left Then


        If IsDate( _
            ws.Cells( _
                LINHA_HORAS_DRAG, _
                ultimaColuna).Value) Then

            HorarioDaPosicaoX = _
                CDate( _
                    ws.Cells( _
                        LINHA_HORAS_DRAG, _
                        ultimaColuna).Value)

        End If

    End If

End Function


' ============================================================
' CONVERTER POSIÇÃO Y EM MÁQUINA
'
' Cada máquina ocupa uma linha da aba PLANEJAMENTO.
'
' Coluna A:
'
' A5 = FETTE 2090
' A6 = FETTE 2
' A7 = MEDISEAL
' ...
'
' O centro vertical do card é utilizado para determinar a linha.
' ============================================================

Private Function MaquinaDaPosicaoY( _
    ByVal posTop As Double, _
    ByVal ws As Worksheet) As String

    Dim ultimaLinha As Long
    Dim linha As Long

    Dim topLinha As Double
    Dim bottomLinha As Double

    Dim centroCard As Double

    Dim maquina As String


    If ws Is Nothing Then Exit Function


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).Row


    If ultimaLinha < PRIMEIRA_LINHA_MAQUINA_DRAG Then Exit Function


    centroCard = _
        posTop + 10


    For linha = _
        PRIMEIRA_LINHA_MAQUINA_DRAG _
        To ultimaLinha


        maquina = _
            Trim$( _
                CStr( _
                    ws.Cells( _
                        linha, _
                        1).Value))


        If maquina <> "" Then


            topLinha = _
                ws.Rows(linha).Top


            bottomLinha = _
                topLinha + _
                ws.Rows(linha).Height


            If centroCard >= topLinha _
               And centroCard < bottomLinha Then


                MaquinaDaPosicaoY = maquina

                Exit Function

            End If

        End If

    Next linha

End Function


' ============================================================
' ARREDONDAR HORÁRIO
' ============================================================

Private Function ArredondarHorario( _
    ByVal horario As Date, _
    ByVal minutos As Long) As Date

    Dim totalMinutos As Double
    Dim arredondado As Double


    If minutos <= 0 Then

        ArredondarHorario = horario

        Exit Function

    End If


    ' --------------------------------------------------------
    ' Usa a parte do dia completa.
    '
    ' Isso evita perder a data quando a timeline atravessa
    ' meia-noite.
    ' --------------------------------------------------------

    totalMinutos = _
        CDbl(horario - _
        DateSerial( _
            Year(horario), _
            Month(horario), _
            Day(horario))) * _
        1440#


    arredondado = _
        Int( _
            totalMinutos / minutos + 0.5) * _
            minutos


    ArredondarHorario = _
        DateSerial( _
            Year(horario), _
            Month(horario), _
            Day(horario)) _
        + _
        arredondado / 1440#

End Function


' ============================================================
' ATUALIZAR DADOS PELO CARD
' ============================================================

Private Sub AtualizarDadosPeloCard( _
    ByVal wsDados As Worksheet, _
    ByRef card As TCardDrag)

    Dim cInicio As Long
    Dim cMaquina As Long
    Dim linha As Long


    linha = card.LinhaDados

    If linha <= 0 Then Exit Sub


    cInicio = _
        EncontrarColunaDrag( _
            wsDados, _
            "Início")


    cMaquina = _
        EncontrarColunaDrag( _
            wsDados, _
            "Máquina")


    If cInicio > 0 Then

        wsDados.Cells( _
            linha, _
            cInicio).Value = _
            card.NovoInicio

        wsDados.Cells( _
            linha, _
            cInicio).NumberFormat = _
            "dd/mm/yyyy hh:mm"

    End If


    ' --------------------------------------------------------
    ' Só altera máquina se a posição vertical encontrou
    ' uma máquina válida.
    ' --------------------------------------------------------

    If cMaquina > 0 Then

        If Trim$(card.Maquina) <> "" Then

            wsDados.Cells( _
                linha, _
                cMaquina).Value = _
                card.Maquina

        End If

    End If

End Sub


' ============================================================
' ORDENAR CARDS
'
' Critério:
'
' 1. Máquina
' 2. Posição X
'
' ============================================================

Private Sub OrdenarCardsPorMaquinaEPosicao( _
    ByRef cards() As TCardDrag, _
    ByVal quantidade As Long)

    Dim i As Long
    Dim j As Long

    Dim temp As TCardDrag


    If quantidade <= 1 Then Exit Sub


    For i = 1 To quantidade - 1

        For j = i + 1 To quantidade

            If CardDeveVirAntes( _
                cards(j), _
                cards(i)) Then


                temp = cards(i)

                cards(i) = cards(j)

                cards(j) = temp

            End If

        Next j

    Next i

End Sub


' ============================================================
' COMPARAR CARDS
' ============================================================

Private Function CardDeveVirAntes( _
    ByRef cardA As TCardDrag, _
    ByRef cardB As TCardDrag) As Boolean

    Dim comparacao As Long


    comparacao = _
        StrComp( _
            cardA.Maquina, _
            cardB.Maquina, _
            vbTextCompare)


    If comparacao < 0 Then

        CardDeveVirAntes = True

        Exit Function

    End If


    If comparacao > 0 Then Exit Function


    If cardA.PosX < cardB.PosX Then

        CardDeveVirAntes = True

    ElseIf cardA.PosX = cardB.PosX Then

        CardDeveVirAntes = _
            StrComp( _
                cardA.NumeroOP, _
                cardB.NumeroOP, _
                vbTextCompare) < 0

    End If

End Function


' ============================================================
' ATRIBUIR NOVAS SEQUÊNCIAS
' ============================================================

Private Sub AtribuirSequencias( _
    ByVal wsDados As Worksheet, _
    ByRef cards() As TCardDrag, _
    ByVal quantidade As Long)

    Dim i As Long

    Dim maquinaAtual As String
    Dim sequencia As Long

    Dim cSequencia As Long


    cSequencia = _
        EncontrarColunaDrag( _
            wsDados, _
            "Sequência")


    If cSequencia = 0 Then

        GarantirColunaDrag _
            wsDados, _
            "Sequência"


        cSequencia = _
            EncontrarColunaDrag( _
                wsDados, _
                "Sequência")

    End If


    maquinaAtual = ""
    sequencia = 0


    For i = 1 To quantidade


        If StrComp( _
            cards(i).Maquina, _
            maquinaAtual, _
            vbTextCompare) <> 0 Then


            maquinaAtual = _
                cards(i).Maquina

            sequencia = 1

        Else

            sequencia = _
                sequencia + 1

        End If


        cards(i).Sequencia = sequencia


        If cards(i).LinhaDados > 0 _
           And cSequencia > 0 Then


            wsDados.Cells( _
                cards(i).LinhaDados, _
                cSequencia).Value = _
                sequencia

        End If


    Next i

End Sub


' ============================================================
' RECALCULAR TODAS AS MÁQUINAS
'
' IMPORTANTE:
'
' A primeira OP de cada máquina mantém o início escolhido
' pelo usuário.
'
' As seguintes são encadeadas:
'
'       Início = Fim da OP anterior
'
' ============================================================

Private Sub RecalcularTodasAsMaquinas()

    Dim ws As Worksheet

    Dim ultimaLinha As Long

    Dim cOP As Long
    Dim cMaquina As Long
    Dim cInicio As Long
    Dim cFim As Long
    Dim cDuracao As Long
    Dim cSequencia As Long

    Dim linha As Long

    Dim maquinaAtual As String
    Dim maquina As String

    Dim inicio As Date
    Dim fim As Date
    Dim fimAnterior As Date

    Dim duracao As Double

    Dim valorInicio As Variant


    Set ws = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS)


    cOP = _
        EncontrarColunaDrag( _
            ws, _
            "OP")


    cMaquina = _
        EncontrarColunaDrag( _
            ws, _
            "Máquina")


    cInicio = _
        EncontrarColunaDrag( _
            ws, _
            "Início")


    cFim = _
        EncontrarColunaDrag( _
            ws, _
            "Fim")


    cDuracao = _
        EncontrarColunaDrag( _
            ws, _
            "Duração Total (h)")


    cSequencia = _
        EncontrarColunaDrag( _
            ws, _
            "Sequência")


    If cOP = 0 _
       Or cMaquina = 0 _
       Or cInicio = 0 _
       Or cFim = 0 _
       Or cDuracao = 0 Then


        Err.Raise _
            vbObjectError + 701, _
            "RecalcularTodasAsMaquinas", _
            "Estrutura de DADOS incompleta. " & _
            "São necessárias as colunas OP, Máquina, Início, Fim e Duração Total (h)."

    End If


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            cOP).End(xlUp).Row


    ' --------------------------------------------------------
    ' ORDENAR A PLANILHA PELO CONTRATO OFICIAL:
    '
    ' Máquina
    ' Sequência
    '
    ' --------------------------------------------------------

    OrdenarDadosDrag _
        ws, _
        cMaquina, _
        cSequencia, _
        ultimaLinha


    maquinaAtual = ""


    For linha = 2 To ultimaLinha


        If Trim$( _
            CStr( _
                ws.Cells( _
                    linha, _
                    cOP).Value)) <> "" Then


            maquina = _
                Trim$( _
                    CStr( _
                        ws.Cells( _
                            linha, _
                            cMaquina).Value))


            duracao = _
                ValorNumeroCelula( _
                    ws.Cells( _
                        linha, _
                        cDuracao).Value)


            If duracao < 0 Then

                duracao = 0

            End If


            ' ------------------------------------------------
            ' PRIMEIRA OP DA MÁQUINA
            ' ------------------------------------------------

            If StrComp( _
                maquina, _
                maquinaAtual, _
                vbTextCompare) <> 0 Then


                valorInicio = _
                    ws.Cells( _
                        linha, _
                        cInicio).Value


                If IsDate(valorInicio) Then

                    inicio = _
                        CDate(valorInicio)

                Else

                    Err.Raise _
                        vbObjectError + 702, _
                        "RecalcularTodasAsMaquinas", _
                        "Início inválido para a OP " & _
                        CStr( _
                            ws.Cells( _
                                linha, _
                                cOP).Value) & _
                        " da máquina " & _
                        maquina & "."

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
                duracao / 24#


            ws.Cells( _
                linha, _
                cInicio).Value = _
                inicio


            ws.Cells( _
                linha, _
                cFim).Value = _
                fim


            ws.Cells( _
                linha, _
                cInicio).NumberFormat = _
                "dd/mm/yyyy hh:mm"


            ws.Cells( _
                linha, _
                cFim).NumberFormat = _
                "dd/mm/yyyy hh:mm"


            ' ------------------------------------------------
            ' STATUS
            ' ------------------------------------------------

            AtualizarStatusDrag _
                ws, _
                linha, _
                inicio, _
                fim


            maquinaAtual = maquina

            fimAnterior = fim


        End If


    Next linha

End Sub


' ============================================================
' ORDENAR DADOS
' ============================================================

Private Sub OrdenarDadosDrag( _
    ByVal ws As Worksheet, _
    ByVal cMaquina As Long, _
    ByVal cSequencia As Long, _
    ByVal ultimaLinha As Long)

    Dim ultimaColuna As Long


    If ultimaLinha < 2 Then Exit Sub


    ultimaColuna = _
        ws.Cells( _
            1, _
            ws.Columns.Count).End(xlToLeft).Column


    With ws.Sort

        .SortFields.Clear


        .SortFields.Add _
            Key:=ws.Range( _
                ws.Cells(2, cMaquina), _
                ws.Cells(ultimaLinha, cMaquina)), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal


        If cSequencia > 0 Then

            .SortFields.Add _
                Key:=ws.Range( _
                    ws.Cells(2, cSequencia), _
                    ws.Cells(ultimaLinha, cSequencia)), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal

        End If


        .SetRange _
            ws.Range( _
                ws.Cells(1, 1), _
                ws.Cells(ultimaLinha, ultimaColuna))


        .Header = xlYes

        .MatchCase = False

        .Orientation = xlTopToBottom

        .Apply

    End With

End Sub


' ============================================================
' STATUS
' ============================================================
'
' PRIORIDADE:
'
' 1. CONCLUÍDO
' 2. ATRASADO
' 3. EM PRODUÇÃO
' 4. PLANEJADO
'
' ATRASADO:
'
'   Se a OP possui Atraso_h > 0 e ainda não terminou.
'
' ============================================================

Private Sub AtualizarStatusDrag( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal inicio As Date, _
    ByVal fim As Date)

    Dim colunaStatus As Long
    Dim atraso As Double


    colunaStatus = _
        EncontrarColunaDrag( _
            ws, _
            "Status")


    If colunaStatus = 0 Then

        GarantirColunaDrag _
            ws, _
            "Status"


        colunaStatus = _
            EncontrarColunaDrag( _
                ws, _
                "Status")

    End If


    atraso = _
        ValorNumeroCampoDrag( _
            ws, _
            linha, _
            "Atraso_h")


    If fim <= Now Then


        ws.Cells( _
            linha, _
            colunaStatus).Value = _
            STATUS_CONCLUIDO


    ElseIf atraso > 0 _
       And inicio > Now Then


        ws.Cells( _
            linha, _
            colunaStatus).Value = _
            STATUS_ATRASADO


    ElseIf inicio <= Now _
       And fim > Now Then


        ws.Cells( _
            linha, _
            colunaStatus).Value = _
            STATUS_EM_PRODUCAO


    Else


        ws.Cells( _
            linha, _
            colunaStatus).Value = _
            STATUS_PLANEJADO


    End If

End Sub


' ============================================================
' ATUALIZAR CARDS
' ============================================================

Private Sub AtualizarCardsDepoisDoDrag()

    Dim wsPlan As Worksheet


    Set wsPlan = _
        ThisWorkbook.Worksheets( _
            ABA_PLANEJAMENTO)


    ' --------------------------------------------------------
    ' Apagar somente cards de OP.
    ' --------------------------------------------------------

    ApagarSomenteCardsOP _
        wsPlan


    ' --------------------------------------------------------
    ' Recriar cards.
    '
    ' NÃO usar On Error Resume Next aqui.
    '
    ' Se CriarCardsAPS estiver ausente ou quebrado,
    ' queremos descobrir o erro.
    ' --------------------------------------------------------

    CriarCardsAPS

End Sub


' ============================================================
' APAGAR SOMENTE CARDS APS
'
' Refeições e outros Shapes são preservados.
' ============================================================

Private Sub ApagarSomenteCardsOP( _
    ByVal ws As Worksheet)

    Dim i As Long


    For i = _
        ws.Shapes.Count _
        To 1 _
        Step -1


        If Left$( _
            ws.Shapes(i).Name, _
            Len(PREFIXO_CARD)) = _
            PREFIXO_CARD Then


            ws.Shapes(i).Delete


        End If


    Next i

End Sub


' ============================================================
' ENCONTRAR LINHA DA OP
' ============================================================

Private Function EncontrarLinhaOPDrag( _
    ByVal ws As Worksheet, _
    ByVal numeroOP As String) As Long

    Dim coluna As Long
    Dim linha As Long
    Dim ultimaLinha As Long

    Dim valor As String


    coluna = _
        EncontrarColunaDrag( _
            ws, _
            "OP")


    If coluna = 0 Then Exit Function


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            coluna).End(xlUp).Row


    For linha = 2 To ultimaLinha


        valor = _
            Trim$( _
                CStr( _
                    ws.Cells( _
                        linha, _
                        coluna).Value))


        If StrComp( _
            valor, _
            Trim$(numeroOP), _
            vbTextCompare) = 0 Then


            EncontrarLinhaOPDrag = linha

            Exit Function

        End If


    Next linha

End Function


' ============================================================
' ENCONTRAR COLUNA
' ============================================================

Private Function EncontrarColunaDrag( _
    ByVal ws As Worksheet, _
    ByVal nome As String) As Long

    Dim coluna As Long
    Dim ultimaColuna As Long


    If ws Is Nothing Then Exit Function


    ultimaColuna = _
        ws.Cells( _
            1, _
            ws.Columns.Count).End(xlToLeft).Column


    For coluna = 1 To ultimaColuna


        If StrComp( _
            NormalizarCabecalhoDrag( _
                CStr( _
                    ws.Cells( _
                        1, _
                        coluna).Value)), _
            NormalizarCabecalhoDrag(nome), _
            vbTextCompare) = 0 Then


            EncontrarColunaDrag = coluna

            Exit Function

        End If


    Next coluna

End Function


' ============================================================
' NORMALIZAR CABEÇALHO
' ============================================================

Private Function NormalizarCabecalhoDrag( _
    ByVal texto As String) As String

    texto = Trim$(texto)

    texto = Replace( _
        texto, _
        vbCr, _
        "")

    texto = Replace( _
        texto, _
        vbLf, _
        "")


    NormalizarCabecalhoDrag = texto

End Function


' ============================================================
' GARANTIR COLUNA
' ============================================================

Private Sub GarantirColunaDrag( _
    ByVal ws As Worksheet, _
    ByVal nome As String)

    Dim coluna As Long
    Dim ultimaColuna As Long


    coluna = _
        EncontrarColunaDrag( _
            ws, _
            nome)


    If coluna > 0 Then Exit Sub


    ultimaColuna = _
        ws.Cells( _
            1, _
            ws.Columns.Count).End(xlToLeft).Column


    If Trim$( _
        CStr( _
            ws.Cells( _
                1, _
                1).Value)) = "" Then

        coluna = 1

    Else

        coluna = _
            ultimaColuna + 1

    End If


    ws.Cells( _
        1, _
        coluna).Value = _
        nome

End Sub


' ============================================================
' VALOR NUMÉRICO
' ============================================================

Private Function ValorNumeroCelula( _
    ByVal valor As Variant) As Double

    If IsError(valor) Then

        ValorNumeroCelula = 0

    ElseIf IsNumeric(valor) Then

        ValorNumeroCelula = _
            CDbl(valor)

    Else

        ValorNumeroCelula = 0

    End If

End Function


' ============================================================
' VALOR NUMÉRICO POR CAMPO
' ============================================================

Private Function ValorNumeroCampoDrag( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal campo As String) As Double

    Dim coluna As Long


    coluna = _
        EncontrarColunaDrag( _
            ws, _
            campo)


    If coluna = 0 Then Exit Function


    ValorNumeroCampoDrag = _
        ValorNumeroCelula( _
            ws.Cells( _
                linha, _
                coluna).Value)

End Function


' ============================================================
' FIM DO DRAGCARDS_APS
' ============================================================