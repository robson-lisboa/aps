Option Explicit

' ============================================================
' EVENTOS APS - MÓDULO 3
' ============================================================
'
' EVENTOS ADICIONAIS DA PRODUÇÃO
'
' Exemplos:
'
'   ATRASO
'   MANUTENÇÃO
'   SETUP ADICIONAL
'   LIMPEZA
'   AJUSTE
'   FALTA DE MATERIAL
'
'
' REGRA:
'
'   Evento aumenta Eventos_h da OP.
'
'   Duração Total (h) é calculada pelo MOTOR DE CÁLCULO:
'
'       Duração Base_h
'       + Atraso_h
'       + Eventos_h
'
'
'   A OP seguinte da mesma máquina é deslocada quando
'   necessário.
'
'
' REFEIÇÃO:
'
'   NÃO é tratada neste módulo.
'
' ============================================================


Public Const ABA_EVENTOS As String = "EVENTOS"


' ============================================================
' EXECUTAR EVENTOS
' ============================================================

Public Sub AplicarEventosAPS()

    On Error GoTo TrataErro

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    ' --------------------------------------------------------
    ' 1. Garantir estrutura da aba EVENTOS
    ' --------------------------------------------------------

    GarantirEstruturaEventosAPS


    ' --------------------------------------------------------
    ' 2. Aplicar eventos ainda não aplicados
    ' --------------------------------------------------------

    AplicarEventosNasOPsAPS


    ' --------------------------------------------------------
    ' 3. Recalcular sequenciamento
    ' --------------------------------------------------------

    RecalcularSequenciamentoComEventosAPS


    ' --------------------------------------------------------
    ' 4. Atualizar cards
    ' --------------------------------------------------------

    AtualizarCardsDepoisEventosAPS


SaidaNormal:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "Eventos aplicados com sucesso." & vbCrLf & vbCrLf & _
        "As OPs afetadas foram recalculadas e o planejamento atualizado.", _
        vbInformation, _
        "APS - Eventos"

    Exit Sub


TrataErro:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "Erro ao aplicar eventos:" & vbCrLf & vbCrLf & _
        "Erro " & Err.Number & ": " & Err.Description, _
        vbCritical, _
        "APS - Eventos"

End Sub


' ============================================================
' GARANTIR ESTRUTURA DA ABA EVENTOS
' ============================================================

Private Sub GarantirEstruturaEventosAPS()

    Dim ws As Worksheet

    Dim cabecalhos As Variant
    Dim i As Long


    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets(ABA_EVENTOS)

    On Error GoTo 0


    ' --------------------------------------------------------
    ' Criar aba caso não exista
    ' --------------------------------------------------------

    If ws Is Nothing Then

        Set ws = _
            ThisWorkbook.Worksheets.Add( _
                After:= _
                ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        ws.Name = ABA_EVENTOS

    End If


    ' --------------------------------------------------------
    ' Cabeçalhos
    ' --------------------------------------------------------

    cabecalhos = Array( _
        "OP", _
        "Tipo", _
        "Duração_h", _
        "Ativo", _
        "Aplicado")


    For i = LBound(cabecalhos) To UBound(cabecalhos)

        If Trim(CStr( _
            ws.Cells(1, i + 1).Value)) = "" Then

            ws.Cells(1, i + 1).Value = _
                cabecalhos(i)

        End If

    Next i


    ' --------------------------------------------------------
    ' Formatação
    ' --------------------------------------------------------

    With ws.Rows(1)

        .Font.Bold = True

    End With


    ws.Columns("A:E").AutoFit


    ' --------------------------------------------------------
    ' Se houver evento novo sem Ativo definido,
    ' considerar ativo.
    '
    ' Não altera eventos que já possuem SIM/NÃO.
    ' --------------------------------------------------------

    Dim ultimaLinha As Long
    Dim linha As Long
    Dim cAtivo As Long


    cAtivo = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Ativo")


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).Row


    If cAtivo > 0 Then

        For linha = 2 To ultimaLinha

            If Trim(CStr( _
                ws.Cells(linha, cAtivo).Value)) = "" Then

                ws.Cells(linha, cAtivo).Value = "SIM"

            End If

        Next linha

    End If

End Sub


' ============================================================
' APLICAR EVENTOS NAS OPs
' ============================================================

Private Sub AplicarEventosNasOPsAPS()

    Dim wsEventos As Worksheet
    Dim wsDados As Worksheet

    Dim ultimaLinha As Long
    Dim linha As Long

    Dim op As String
    Dim tipo As String

    Dim duracao As Double

    Dim linhaOP As Long

    Dim colAplicado As Long


    Set wsEventos = _
        ThisWorkbook.Worksheets( _
            ABA_EVENTOS)


    Set wsDados = _
        ThisWorkbook.Worksheets( _
            "DADOS")


    ' --------------------------------------------------------
    ' Garantir colunas
    ' --------------------------------------------------------

    GarantirColunaEventoAPS _
        wsEventos, _
        "Aplicado"


    colAplicado = _
        EncontrarColunaEventoAPS( _
            wsEventos, _
            "Aplicado")


    ultimaLinha = _
        wsEventos.Cells( _
            wsEventos.Rows.Count, _
            1).End(xlUp).Row


    ' --------------------------------------------------------
    ' Percorrer eventos
    ' --------------------------------------------------------

    For linha = 2 To ultimaLinha


        ' ----------------------------------------------------
        ' Evento precisa estar ativo
        ' ----------------------------------------------------

        If EventoAtivoAPS( _
            wsEventos, _
            linha) Then


            ' ------------------------------------------------
            ' Não reaplicar evento já processado
            ' ------------------------------------------------

            If Not EventoJaAplicadoAPS( _
                wsEventos, _
                linha, _
                colAplicado) Then


                op = _
                    ValorEventoAPS( _
                        wsEventos, _
                        linha, _
                        "OP")


                tipo = _
                    ValorEventoAPS( _
                        wsEventos, _
                        linha, _
                        "Tipo")


                duracao = _
                    ValorNumeroEventoAPS( _
                        wsEventos, _
                        linha, _
                        "Duração_h")


                ' ------------------------------------------------
                ' Validar evento
                ' ------------------------------------------------

                If op <> "" _
                   And duracao > 0 Then


                    linhaOP = _
                        EncontrarLinhaOPAPS( _
                            wsDados, _
                            op)


                    If linhaOP > 0 Then


                        ' ----------------------------------------
                        ' Adicionar evento à OP
                        ' ----------------------------------------

                        AdicionarEventoNaOPAPS _
                            wsDados, _
                            linhaOP, _
                            duracao, _
                            tipo


                        ' ----------------------------------------
                        ' Marcar como aplicado
                        ' ----------------------------------------

                        wsEventos.Cells( _
                            linha, _
                            colAplicado).Value = "SIM"


                        wsEventos.Cells( _
                            linha, _
                            colAplicado).NumberFormat = "@"


                    End If

                End If

            End If

        End If

    Next linha

End Sub


' ============================================================
' EVENTO JÁ FOI APLICADO?
' ============================================================

Private Function EventoJaAplicadoAPS( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal colAplicado As Long) As Boolean


    If colAplicado = 0 Then

        EventoJaAplicadoAPS = False

        Exit Function

    End If


    Select Case UCase( _
        Trim(CStr( _
            ws.Cells( _
                linha, _
                colAplicado).Value)))


        Case "SIM", "S", "TRUE", "1"

            EventoJaAplicadoAPS = True


        Case Else

            EventoJaAplicadoAPS = False


    End Select

End Function


' ============================================================
' ADICIONAR EVENTO À OP
' ============================================================

Private Sub AdicionarEventoNaOPAPS( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal duracaoHoras As Double, _
    ByVal tipo As String)


    Dim colEvento As Long
    Dim colTipo As Long


    ' --------------------------------------------------------
    ' Garantir Eventos_h
    ' --------------------------------------------------------

    GarantirColunaEventoAPS _
        ws, _
        "Eventos_h"


    colEvento = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Eventos_h")


    ' --------------------------------------------------------
    ' SOMAR somente o evento.
    '
    ' A Duração Total NÃO é alterada diretamente.
    '
    ' O Motor de Cálculo será responsável por calcular:
    '
    ' Duração Base_h
    ' + Atraso_h
    ' + Eventos_h
    '
    ' --------------------------------------------------------

    ws.Cells( _
        linha, _
        colEvento).Value = _
            NzNumeroEventoAPS( _
                ws.Cells( _
                    linha, _
                    colEvento).Value) _
            + duracaoHoras


    ' --------------------------------------------------------
    ' Recalcular duração da OP
    ' --------------------------------------------------------

    RecalcularDuracaoOP _
        ws, _
        linha


    ' --------------------------------------------------------
    ' Registrar tipo do evento
    ' --------------------------------------------------------

    GarantirColunaEventoAPS _
        ws, _
        "Eventos"


    colTipo = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Eventos")


    If Trim(CStr( _
        ws.Cells( _
            linha, _
            colTipo).Value)) = "" Then


        ws.Cells( _
            linha, _
            colTipo).Value = tipo


    Else


        If tipo <> "" Then

            ws.Cells( _
                linha, _
                colTipo).Value = _
                    ws.Cells( _
                        linha, _
                        colTipo).Value _
                    & " | " & tipo

        End If


    End If

End Sub


' ============================================================
' RECALCULAR SEQUENCIAMENTO COM EVENTOS
' ============================================================

Private Sub RecalcularSequenciamentoComEventosAPS()

    Dim ws As Worksheet

    Dim ultimaLinha As Long
    Dim i As Long

    Dim colOP As Long
    Dim colMaquina As Long
    Dim colSequencia As Long
    Dim colInicio As Long
    Dim colFim As Long
    Dim colDuracao As Long

    Dim maquinaAtual As String
    Dim maquina As String

    Dim inicioPlanejado As Date
    Dim inicioOriginal As Date

    Dim duracao As Double

    Dim novoInicio As Date
    Dim novoFim As Date

    Dim fimAnterior As Date


    Set ws = _
        ThisWorkbook.Worksheets( _
            "DADOS")


    ' --------------------------------------------------------
    ' Localizar colunas
    ' --------------------------------------------------------

    colOP = _
        EncontrarColunaEventoAPS( _
            ws, _
            "OP")


    colMaquina = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Máquina")


    colSequencia = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Sequência")


    colInicio = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Início")


    colFim = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Fim")


    colDuracao = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Duração Total (h)")


    If colOP = 0 _
       Or colMaquina = 0 _
       Or colInicio = 0 _
       Or colFim = 0 _
       Or colDuracao = 0 Then


        Err.Raise _
            vbObjectError + 3101, _
            "Módulo 3 - Eventos", _
            "A aba DADOS não possui todas as colunas necessárias."


    End If


    ' --------------------------------------------------------
    ' Garantir ordenação
    ' --------------------------------------------------------

    If colSequencia > 0 Then

        OrdenarOPsAPS

    End If


    ' --------------------------------------------------------
    ' Última linha
    ' --------------------------------------------------------

    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            colOP).End(xlUp).Row


    maquinaAtual = ""
    fimAnterior = 0


    ' ========================================================
    ' PERCORRER OPs
    ' ========================================================

    For i = 2 To ultimaLinha


        If Trim(CStr( _
            ws.Cells( _
                i, _
                colOP).Value)) <> "" Then


            maquina = _
                Trim(CStr( _
                    ws.Cells( _
                        i, _
                        colMaquina).Value))


            ' ------------------------------------------------
            ' DURAÇÃO TOTAL
            ' ------------------------------------------------

            duracao = _
                NzNumeroEventoAPS( _
                    ws.Cells( _
                        i, _
                        colDuracao).Value)


            ' ------------------------------------------------
            ' DATA BASE DA OP
            '
            ' Prioridade:
            '
            ' 1. Início Original
            ' 2. Data Planejada
            ' 3. Início atual
            '
            ' ------------------------------------------------

            inicioOriginal = _
                ObterInicioBaseAPS( _
                    ws, _
                    i)


            inicioPlanejado = _
                inicioOriginal


            ' ------------------------------------------------
            ' PRIMEIRA OP DA MÁQUINA
            ' ------------------------------------------------

            If maquinaAtual = "" _
               Or StrComp( _
                    maquina, _
                    maquinaAtual, _
                    vbTextCompare) <> 0 Then


                novoInicio = _
                    inicioPlanejado


            Else


                ' ------------------------------------------------
                ' OP seguinte da mesma máquina.
                '
                ' Nunca pode começar antes do término da OP
                ' anterior.
                ' ------------------------------------------------

                If fimAnterior > inicioPlanejado Then

                    novoInicio = _
                        fimAnterior

                Else

                    novoInicio = _
                        inicioPlanejado

                End If


            End If


            ' ------------------------------------------------
            ' CALCULAR NOVO FIM
            ' ------------------------------------------------

            novoFim = _
                novoInicio _
                + (duracao / 24)


            ' ------------------------------------------------
            ' ESCREVER INÍCIO
            ' ------------------------------------------------

            ws.Cells( _
                i, _
                colInicio).Value = _
                    novoInicio


            ws.Cells( _
                i, _
                colInicio).NumberFormat = _
                    "dd/mm/yyyy hh:mm"


            ' ------------------------------------------------
            ' ESCREVER FIM
            ' ------------------------------------------------

            ws.Cells( _
                i, _
                colFim).Value = _
                    novoFim


            ws.Cells( _
                i, _
                colFim).NumberFormat = _
                    "dd/mm/yyyy hh:mm"


            ' ------------------------------------------------
            ' STATUS
            ' ------------------------------------------------

            AtualizarStatusEventoAPS _
                ws, _
                i, _
                novoInicio, _
                novoFim


            ' ------------------------------------------------
            ' GUARDAR ESTADO
            ' ------------------------------------------------

            maquinaAtual = maquina
            fimAnterior = novoFim


        End If

    Next i

End Sub


' ============================================================
' OBTER INÍCIO BASE
' ============================================================

Private Function ObterInicioBaseAPS( _
    ByVal ws As Worksheet, _
    ByVal linha As Long) As Date

    Dim c As Long


    ' --------------------------------------------------------
    ' 1. Início Original
    ' --------------------------------------------------------

    c = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Início Original")


    If c > 0 Then

        If IsDate( _
            ws.Cells( _
                linha, _
                c).Value) Then


            ObterInicioBaseAPS = _
                CDate( _
                    ws.Cells( _
                        linha, _
                        c).Value)


            Exit Function

        End If

    End If


    ' --------------------------------------------------------
    ' 2. Data Planejada
    ' --------------------------------------------------------

    c = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Data Planejada")


    If c > 0 Then

        If IsDate( _
            ws.Cells( _
                linha, _
                c).Value) Then


            ObterInicioBaseAPS = _
                CDate( _
                    ws.Cells( _
                        linha, _
                        c).Value)


            Exit Function

        End If

    End If


    ' --------------------------------------------------------
    ' 3. Início atual
    ' --------------------------------------------------------

    c = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Início")


    If c > 0 Then

        If IsDate( _
            ws.Cells( _
                linha, _
                c).Value) Then


            ObterInicioBaseAPS = _
                CDate( _
                    ws.Cells( _
                        linha, _
                        c).Value)

        End If

    End If

End Function


' ============================================================
' ORDENAR OPs
' ============================================================

Private Sub OrdenarOPsAPS()

    Dim ws As Worksheet

    Dim ultimaLinha As Long
    Dim ultimaColuna As Long

    Dim colMaquina As Long
    Dim colSequencia As Long


    Set ws = _
        ThisWorkbook.Worksheets( _
            "DADOS")


    colMaquina = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Máquina")


    colSequencia = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Sequência")


    If colMaquina = 0 _
       Or colSequencia = 0 Then Exit Sub


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            colMaquina).End(xlUp).Row


    ultimaColuna = _
        ws.Cells( _
            1, _
            ws.Columns.Count).End(xlToLeft).Column


    If ultimaLinha < 2 Then Exit Sub


    With ws.Sort


        .SortFields.Clear


        .SortFields.Add _
            Key:=ws.Range( _
                ws.Cells(2, colMaquina), _
                ws.Cells(ultimaLinha, colMaquina)), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal


        .SortFields.Add _
            Key:=ws.Range( _
                ws.Cells(2, colSequencia), _
                ws.Cells(ultimaLinha, colSequencia)), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal


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
' ATUALIZAR CARDS APÓS EVENTOS
' ============================================================

Private Sub AtualizarCardsDepoisEventosAPS()

    On Error GoTo TrataErro

    ' --------------------------------------------------------
    ' O Módulo 6 reconstrói todos os cards.
    '
    ' Isso é mais seguro do que tentar movimentar somente
    ' alguns shapes, pois eventos podem deslocar várias OPs.
    ' --------------------------------------------------------

    CriarCardsAPS

    Exit Sub

TrataErro:

    MsgBox _
        "Erro ao atualizar cards após eventos:" & _
        vbCrLf & vbCrLf & _
        "Procedimento: " & Err.Source & _
        vbCrLf & _
        "Erro: " & CStr(Err.Number) & _
        vbCrLf & _
        "Descrição: " & Err.Description, _
        vbCritical, _
        "APS - Eventos"

End Sub


' ============================================================
' EVENTO ATIVO?
' ============================================================

Private Function EventoAtivoAPS( _
    ByVal ws As Worksheet, _
    ByVal linha As Long) As Boolean

    Dim c As Long
    Dim valor As Variant


    c = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Ativo")


    ' --------------------------------------------------------
    ' Se não houver coluna, considera ativo.
    ' --------------------------------------------------------

    If c = 0 Then

        EventoAtivoAPS = True

        Exit Function

    End If


    valor = _
        ws.Cells( _
            linha, _
            c).Value


    ' --------------------------------------------------------
    ' Vazio = ativo
    ' --------------------------------------------------------

    If Trim(CStr(valor)) = "" Then

        EventoAtivoAPS = True

        Exit Function

    End If


    Select Case VarType(valor)


        Case vbBoolean

            EventoAtivoAPS = _
                CBool(valor)


        Case vbByte, _
             vbInteger, _
             vbLong, _
             vbSingle, _
             vbDouble, _
             vbCurrency

            EventoAtivoAPS = _
                (CDbl(valor) <> 0)


        Case Else

            Select Case UCase( _
                Trim(CStr(valor)))


                Case "SIM", _
                     "S", _
                     "TRUE", _
                     "VERDADEIRO", _
                     "1", _
                     "ATIVO"

                    EventoAtivoAPS = True


                Case Else

                    EventoAtivoAPS = False


            End Select


    End Select

End Function


' ============================================================
' ENCONTRAR LINHA DA OP
' ============================================================

Private Function EncontrarLinhaOPAPS( _
    ByVal ws As Worksheet, _
    ByVal numeroOP As String) As Long

    Dim coluna As Long
    Dim ultimaLinha As Long
    Dim linha As Long


    coluna = _
        EncontrarColunaEventoAPS( _
            ws, _
            "OP")


    If coluna = 0 Then Exit Function


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            coluna).End(xlUp).Row


    For linha = 2 To ultimaLinha


        If StrComp( _
            Trim(CStr( _
                ws.Cells( _
                    linha, _
                    coluna).Value)), _
            Trim(numeroOP), _
            vbTextCompare) = 0 Then


            EncontrarLinhaOPAPS = linha

            Exit Function


        End If


    Next linha

End Function


' ============================================================
' GARANTIR COLUNA
' ============================================================

Private Sub GarantirColunaEventoAPS( _
    ByVal ws As Worksheet, _
    ByVal nome As String)

    Dim coluna As Long
    Dim ultimaColuna As Long


    coluna = _
        EncontrarColunaEventoAPS( _
            ws, _
            nome)


    If coluna > 0 Then Exit Sub


    ultimaColuna = _
        ws.Cells( _
            1, _
            ws.Columns.Count).End(xlToLeft).Column


    If Trim(CStr( _
        ws.Cells(1, ultimaColuna).Value)) = "" Then

        coluna = ultimaColuna

    Else

        coluna = ultimaColuna + 1

    End If


    ws.Cells(1, coluna).Value = nome

End Sub


' ============================================================
' ENCONTRAR COLUNA
' ============================================================

Private Function EncontrarColunaEventoAPS( _
    ByVal ws As Worksheet, _
    ByVal nome As String) As Long

    Dim ultimaColuna As Long
    Dim coluna As Long


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


            EncontrarColunaEventoAPS = coluna

            Exit Function


        End If


    Next coluna

End Function


' ============================================================
' VALOR TEXTO
' ============================================================

Private Function ValorEventoAPS( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal campo As String) As String

    Dim c As Long


    c = _
        EncontrarColunaEventoAPS( _
            ws, _
            campo)


    If c > 0 Then

        ValorEventoAPS = _
            Trim(CStr( _
                ws.Cells( _
                    linha, _
                    c).Value))

    End If

End Function


' ============================================================
' VALOR NUMÉRICO
' ============================================================

Private Function ValorNumeroEventoAPS( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal campo As String) As Double

    Dim c As Long


    c = _
        EncontrarColunaEventoAPS( _
            ws, _
            campo)


    If c > 0 Then

        ValorNumeroEventoAPS = _
            NzNumeroEventoAPS( _
                ws.Cells( _
                    linha, _
                    c).Value)

    End If

End Function


' ============================================================
' NÚMERO SEGURO
' ============================================================

Private Function NzNumeroEventoAPS( _
    ByVal valor As Variant) As Double

    If IsNumeric(valor) Then

        NzNumeroEventoAPS = _
            CDbl(valor)

    Else

        NzNumeroEventoAPS = 0

    End If

End Function


' ============================================================
' ATUALIZAR STATUS
' ============================================================

Private Sub AtualizarStatusEventoAPS( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal inicio As Date, _
    ByVal fim As Date)

    Dim c As Long


    c = _
        EncontrarColunaEventoAPS( _
            ws, _
            "Status")


    If c = 0 Then Exit Sub


    If fim < Now Then


        ws.Cells( _
            linha, _
            c).Value = _
                "CONCLUÍDO"


    ElseIf inicio <= Now _
       And fim >= Now Then


        ws.Cells( _
            linha, _
            c).Value = _
                "EM PRODUÇÃO"


    Else


        ws.Cells( _
            linha, _
            c).Value = _
                "PLANEJADO"


    End If

End Sub


' ============================================================
' FIM DO MÓDULO 3
' ============================================================