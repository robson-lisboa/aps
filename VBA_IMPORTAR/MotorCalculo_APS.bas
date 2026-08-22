Option Explicit

' ============================================================
' MOTORCALCULO_APS
' APS PURAN
' ============================================================
'
' RESPONSABILIDADE:
'
'   Calcular os tempos das OPs utilizando:
'
'       DADOS
'       RECURSOS
'
'   Resultado:
'
'       Caixas
'       Capacidade_h
'       Produção_h
'       Setup_h
'       Duração Base_h
'       Duração Total (h)
'
' REGRA:
'
'       Duração Total (h)
'       =
'       Duração Base_h
'       + Atraso_h
'       + Eventos_h
'
' IMPORTANTE:
'
'   A duração BASE representa somente:
'
'       Produção + Setup
'
'   Atrasos e Eventos são parcelas independentes.
'
'   A duração total nunca é incrementada diretamente.
'   Ela sempre é recalculada do zero.
'
' ============================================================


Private Const ABA_RECURSOS_MOTOR As String = "RECURSOS"

Public Const COLUNA_DURACAO_BASE_MOTOR As String = "Duração Base_h"
Public Const COLUNA_DURACAO_TOTAL_MOTOR As String = "Duração Total (h)"


' ============================================================
' EXECUTAR MOTOR APS
' ============================================================

Public Sub ExecutarMotorAPS()

    Dim wsDados As Worksheet
    Dim wsRecursos As Worksheet

    Dim ultimaLinha As Long
    Dim linha As Long
    Dim colunaOP As Long

    Dim opsSemMaquina As String
    Dim opsSemVelocidade As String
    Dim opsSemQuantidade As String
    Dim opsSemComprimidosPorCaixa As String

    Dim calculoOK As Boolean

    Dim calcAnterior As XlCalculation
    Dim eventosAnterior As Boolean
    Dim telaAnterior As Boolean
    Dim statusAnterior As Variant

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
    Application.StatusBar = "APS PURAN - Executando motor de cálculo..."

    ' --------------------------------------------------------
    ' VALIDAR DADOS
    ' --------------------------------------------------------

    Set wsDados = ObterPlanilhaMotor(ABA_DADOS)

    If wsDados Is Nothing Then

        Err.Raise _
            vbObjectError + 501, _
            "ExecutarMotorAPS", _
            "A aba '" & ABA_DADOS & "' não foi encontrada."

    End If

    Set wsRecursos = ObterPlanilhaMotor(ABA_RECURSOS_MOTOR)

    If wsRecursos Is Nothing Then

        Err.Raise _
            vbObjectError + 502, _
            "ExecutarMotorAPS", _
            "A aba '" & ABA_RECURSOS_MOTOR & "' não foi encontrada."

    End If

    ' --------------------------------------------------------
    ' GARANTIR ESTRUTURA
    ' --------------------------------------------------------

    GarantirColunaMotor wsDados, "OP"
    GarantirColunaMotor wsDados, "Máquina"
    GarantirColunaMotor wsDados, "Quantidade"
    GarantirColunaMotor wsDados, "Unidade"

    GarantirColunaMotor wsDados, "Caixas"
    GarantirColunaMotor wsDados, "Velocidade"
    GarantirColunaMotor wsDados, "OEE"
    GarantirColunaMotor wsDados, "Capacidade_h"
    GarantirColunaMotor wsDados, "Produção_h"
    GarantirColunaMotor wsDados, "Setup_h"
    GarantirColunaMotor wsDados, COLUNA_DURACAO_BASE_MOTOR
    GarantirColunaMotor wsDados, COLUNA_DURACAO_TOTAL_MOTOR
    GarantirColunaMotor wsDados, "Atraso_h"
    GarantirColunaMotor wsDados, "Eventos_h"

    ' --------------------------------------------------------
    ' LOCALIZAR OP
    ' --------------------------------------------------------

    colunaOP = EncontrarColunaMotor(wsDados, "OP")

    If colunaOP = 0 Then

        Err.Raise _
            vbObjectError + 503, _
            "ExecutarMotorAPS", _
            "A coluna 'OP' não foi encontrada."

    End If

    ultimaLinha = _
        wsDados.Cells( _
            wsDados.Rows.Count, _
            colunaOP).End(xlUp).Row

    If ultimaLinha < 2 Then

        Err.Raise _
            vbObjectError + 504, _
            "ExecutarMotorAPS", _
            "A aba DADOS não possui OPs para calcular."

    End If

    ' --------------------------------------------------------
    ' PROCESSAR OPs
    ' --------------------------------------------------------

    For linha = 2 To ultimaLinha

        Application.StatusBar = _
            "APS PURAN - Calculando OP " & _
            CStr(linha - 1) & " de " & _
            CStr(ultimaLinha - 1) & "..."

        If Trim$(ValorTextoMotor(wsDados, linha, "OP")) <> "" Then

            CalcularOP _
                wsDados, _
                wsRecursos, _
                linha, _
                opsSemMaquina, _
                opsSemVelocidade, _
                opsSemQuantidade, _
                opsSemComprimidosPorCaixa

        End If

    Next linha

    calculoOK = True

SaidaNormal:

    ' --------------------------------------------------------
    ' RESTAURAR EXCEL
    ' --------------------------------------------------------

    On Error Resume Next

    Application.Calculation = calcAnterior
    Application.EnableEvents = eventosAnterior
    Application.ScreenUpdating = telaAnterior
    Application.StatusBar = statusAnterior

    On Error GoTo 0

    If Not calculoOK Then Exit Sub

    ' --------------------------------------------------------
    ' RESUMO
    ' --------------------------------------------------------

    Dim resumo As String

    resumo = _
        "Motor de cálculo APS executado com sucesso." & _
        vbCrLf & vbCrLf & _
        "Foram recalculados:" & vbCrLf & _
        "• Caixas" & vbCrLf & _
        "• Capacidade" & vbCrLf & _
        "• Produção" & vbCrLf & _
        "• Setup" & vbCrLf & _
        "• Duração Base" & vbCrLf & _
        "• Duração Total"

    If opsSemMaquina <> "" Then

        resumo = _
            resumo & _
            vbCrLf & vbCrLf & _
            "ATENÇÃO - máquina não encontrada em RECURSOS:" & _
            vbCrLf & _
            opsSemMaquina

    End If

    If opsSemVelocidade <> "" Then

        resumo = _
            resumo & _
            vbCrLf & vbCrLf & _
            "ATENÇÃO - Velocidade/OEE inválidos:" & _
            vbCrLf & _
            opsSemVelocidade

    End If

    If opsSemQuantidade <> "" Then

        resumo = _
            resumo & _
            vbCrLf & vbCrLf & _
            "ATENÇÃO - Quantidade inválida:" & _
            vbCrLf & _
            opsSemQuantidade

    End If

    If opsSemComprimidosPorCaixa <> "" Then

        resumo = _
            resumo & _
            vbCrLf & vbCrLf & _
            "ATENÇÃO - Comprimidos por caixa não cadastrado:" & _
            vbCrLf & _
            opsSemComprimidosPorCaixa

    End If

    MsgBox _
        resumo, _
        vbInformation, _
        "APS PURAN - Motor de Cálculo"

    Exit Sub


TrataErro:

    calculoOK = False

    On Error Resume Next

    Application.Calculation = calcAnterior
    Application.EnableEvents = eventosAnterior
    Application.ScreenUpdating = telaAnterior
    Application.StatusBar = statusAnterior

    On Error GoTo 0

    MsgBox _
        "Erro no Motor APS." & _
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
' CALCULAR UMA OP
' ============================================================

Private Sub CalcularOP( _
    ByVal wsDados As Worksheet, _
    ByVal wsRecursos As Worksheet, _
    ByVal linha As Long, _
    ByRef opsSemMaquina As String, _
    ByRef opsSemVelocidade As String, _
    ByRef opsSemQuantidade As String, _
    ByRef opsSemComprimidosPorCaixa As String)

    Dim op As String
    Dim maquina As String
    Dim unidade As String

    Dim quantidade As Double
    Dim comprimidosPorCaixa As Double
    Dim caixas As Double

    Dim velocidade As Double
    Dim oee As Double
    Dim setupPadrao As Double

    Dim capacidadeEfetiva As Double
    Dim producaoH As Double
    Dim setupH As Double
    Dim duracaoBase As Double

    Dim linhaRecurso As Long

    Dim possuiQuantidade As Boolean
    Dim possuiVelocidade As Boolean
    Dim possuiOEE As Boolean


    ' --------------------------------------------------------
    ' IDENTIFICAÇÃO
    ' --------------------------------------------------------

    op = ValorTextoMotor(wsDados, linha, "OP")

    maquina = ValorTextoMotor(wsDados, linha, "Máquina")

    If op = "" Then Exit Sub


    ' --------------------------------------------------------
    ' VALIDAR MÁQUINA
    ' --------------------------------------------------------

    If maquina = "" Then

        opsSemMaquina = _
            opsSemMaquina & _
            "• " & op & " - máquina vazia" & vbCrLf

        LimparCalculosOP wsDados, linha

        Exit Sub

    End If


    linhaRecurso = _
        EncontrarLinhaRecurso( _
            wsRecursos, _
            maquina)


    If linhaRecurso = 0 Then

        opsSemMaquina = _
            opsSemMaquina & _
            "• " & op & _
            " (" & maquina & ")" & _
            vbCrLf

        LimparCalculosOP wsDados, linha

        Exit Sub

    End If


    ' --------------------------------------------------------
    ' QUANTIDADE
    ' --------------------------------------------------------

    quantidade = _
        LerNumeroSeguro( _
            wsDados, _
            linha, _
            "Quantidade", _
            possuiQuantidade)

    If Not possuiQuantidade Or quantidade < 0 Then

        opsSemQuantidade = _
            opsSemQuantidade & _
            "• " & op & _
            " - quantidade inválida" & _
            vbCrLf

        EscreverNumeroMotor _
            wsDados, _
            linha, _
            "Caixas", _
            0

        GoTo CalcularRecursos

    End If


    ' --------------------------------------------------------
    ' UNIDADE
    ' --------------------------------------------------------

    unidade = _
        UCase$( _
            Trim$( _
                ValorTextoMotor( _
                    wsDados, _
                    linha, _
                    "Unidade")))


    ' --------------------------------------------------------
    ' DADOS DO RECURSO
    ' --------------------------------------------------------

CalcularRecursos:

    velocidade = _
        LerNumeroSeguroRecurso( _
            wsRecursos, _
            linhaRecurso, _
            "Velocidade", _
            possuiVelocidade)

    oee = _
        LerNumeroSeguroRecurso( _
            wsRecursos, _
            linhaRecurso, _
            "OEE", _
            possuiOEE)

    setupPadrao = _
        ValorNumeroMotor( _
            wsRecursos, _
            linhaRecurso, _
            "Setup padrão")

    comprimidosPorCaixa = _
        ValorNumeroMotor( _
            wsRecursos, _
            linhaRecurso, _
            "Comprimidos por caixa")


    ' --------------------------------------------------------
    ' NORMALIZAR OEE
    ' --------------------------------------------------------

    oee = NormalizarOEE(oee)


    ' --------------------------------------------------------
    ' CONVERTER QUANTIDADE PARA CAIXAS
    ' --------------------------------------------------------

    If possuiQuantidade Then

        If UnidadeRepresentaCaixa(unidade) Then

            caixas = quantidade

        Else

            If comprimidosPorCaixa > 0 Then

                caixas = _
                    quantidade / comprimidosPorCaixa

            Else

                opsSemComprimidosPorCaixa = _
                    opsSemComprimidosPorCaixa & _
                    "• " & op & _
                    " (" & maquina & ")" & _
                    " - Comprimidos por caixa não cadastrado" & vbCrLf

                LimparCalculosOP wsDados, linha

                Exit Sub

            End If

        End If

    Else

        caixas = 0

    End If


    ' --------------------------------------------------------
    ' ARREDONDAMENTO DAS CAIXAS
    ' --------------------------------------------------------
    '
    ' OPs em caixas devem representar caixas inteiras.
    '
    ' Caso uma conversão de comprimidos gere 1333,2 caixas,
    ' serão necessárias 1334 caixas.
    '

    If caixas > 0 Then

        caixas = _
            Application.WorksheetFunction.RoundUp( _
                caixas, _
                0)

    End If


    EscreverNumeroMotor _
        wsDados, _
        linha, _
        "Caixas", _
        caixas


    ' --------------------------------------------------------
    ' GRAVAR PARÂMETROS DO RECURSO NA OP
    ' --------------------------------------------------------

    EscreverNumeroMotor _
        wsDados, _
        linha, _
        "Velocidade", _
        velocidade

    EscreverNumeroMotor _
        wsDados, _
        linha, _
        "OEE", _
        oee

    EscreverNumeroMotor _
        wsDados, _
        linha, _
        "Capacidade_h", _
        velocidade * oee


    ' --------------------------------------------------------
    ' CAPACIDADE EFETIVA
    ' --------------------------------------------------------

    capacidadeEfetiva = _
        velocidade * oee


    ' --------------------------------------------------------
    ' PRODUÇÃO
    ' --------------------------------------------------------

    If caixas <= 0 Then

        producaoH = 0

    ElseIf capacidadeEfetiva <= 0 Then

        producaoH = 0

        opsSemVelocidade = _
            opsSemVelocidade & _
            "• " & op & _
            " (" & maquina & ")" & _
            " - Velocidade/OEE inválidos" & _
            vbCrLf

    Else

        producaoH = _
            caixas / capacidadeEfetiva

    End If


    ' --------------------------------------------------------
    ' SETUP
    ' --------------------------------------------------------

    If setupPadrao < 0 Then
        setupPadrao = 0
    End If

    setupH = setupPadrao


    EscreverNumeroMotor _
        wsDados, _
        linha, _
        "Produção_h", _
        producaoH

    EscreverNumeroMotor _
        wsDados, _
        linha, _
        "Setup_h", _
        setupH


    ' --------------------------------------------------------
    ' DURAÇÃO BASE
    ' --------------------------------------------------------

    duracaoBase = _
        producaoH + setupH

    EscreverNumeroMotor _
        wsDados, _
        linha, _
        COLUNA_DURACAO_BASE_MOTOR, _
        duracaoBase


    ' --------------------------------------------------------
    ' DURAÇÃO TOTAL
    ' --------------------------------------------------------
    '
    ' IMPORTANTE:
    '
    ' Nunca incrementar a duração existente.
    '
    ' Sempre:
    '
    ' Base + Atraso + Eventos
    '

    RecalcularDuracaoOP _
        wsDados, _
        linha

End Sub


' ============================================================
' RECALCULAR DURAÇÃO TOTAL
' ============================================================
'
' REGRA CENTRAL DO APS:
'
' Duração Total =
'       Duração Base
'       + Atraso
'       + Eventos
'
' A rotina sempre parte do zero.
'
' Isso evita:
'
' 10h + 2h
' depois
' 12h + 2h
' depois
' 14h + 2h
'
' O resultado será sempre:
'
' 10h + 2h = 12h
'
' ============================================================

Public Sub RecalcularDuracaoOP( _
    ByVal wsDados As Worksheet, _
    ByVal linha As Long)

    Dim base As Double
    Dim atraso As Double
    Dim eventos As Double
    Dim total As Double


    If wsDados Is Nothing Then Exit Sub

    If linha < 2 Then Exit Sub


    base = _
        ValorNumeroMotor( _
            wsDados, _
            linha, _
            COLUNA_DURACAO_BASE_MOTOR)

    atraso = _
        ValorNumeroMotor( _
            wsDados, _
            linha, _
            "Atraso_h")

    eventos = _
        ValorNumeroMotor( _
            wsDados, _
            linha, _
            "Eventos_h")


    If base < 0 Then base = 0
    If atraso < 0 Then atraso = 0
    If eventos < 0 Then eventos = 0


    total = _
        base + _
        atraso + _
        eventos


    EscreverNumeroMotor _
        wsDados, _
        linha, _
        COLUNA_DURACAO_TOTAL_MOTOR, _
        total

End Sub


' ============================================================
' FUNÇÃO PÚBLICA PARA RECALCULAR TODAS AS OPs
' ============================================================

Public Sub RecalcularTodasDuracoesAPS()

    Dim ws As Worksheet

    Dim colunaOP As Long
    Dim ultimaLinha As Long
    Dim linha As Long

    On Error GoTo TrataErro

    Set ws = _
        ObterPlanilhaMotor(ABA_DADOS)

    If ws Is Nothing Then

        Err.Raise _
            vbObjectError + 505, _
            "RecalcularTodasDuracoesAPS", _
            "A aba DADOS não foi encontrada."

    End If


    colunaOP = _
        EncontrarColunaMotor(ws, "OP")


    If colunaOP = 0 Then

        Err.Raise _
            vbObjectError + 506, _
            "RecalcularTodasDuracoesAPS", _
            "A coluna OP não foi encontrada."

    End If


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            colunaOP).End(xlUp).Row


    For linha = 2 To ultimaLinha

        If Trim$(ValorTextoMotor(ws, linha, "OP")) <> "" Then

            RecalcularDuracaoOP _
                ws, _
                linha

        End If

    Next linha


    Exit Sub


TrataErro:

    MsgBox _
        "Erro ao recalcular durações:" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "APS PURAN"

End Sub


' ============================================================
' ENCONTRAR LINHA DO RECURSO
' ============================================================

Private Function EncontrarLinhaRecurso( _
    ByVal ws As Worksheet, _
    ByVal maquina As String) As Long

    Dim colunaMaquina As Long
    Dim ultimaLinha As Long
    Dim linha As Long
    Dim valor As String


    colunaMaquina = _
        EncontrarColunaMotor( _
            ws, _
            "Máquina")


    If colunaMaquina = 0 Then

        ' Compatibilidade com cadastros antigos:
        ' Máquina na coluna A.

        colunaMaquina = 1

    End If


    ultimaLinha = _
        ws.Cells( _
            ws.Rows.Count, _
            colunaMaquina).End(xlUp).Row


    For linha = 2 To ultimaLinha

        valor = _
            Trim$( _
                CStr( _
                    ws.Cells( _
                        linha, _
                        colunaMaquina).Value))


        If StrComp( _
            valor, _
            Trim$(maquina), _
            vbTextCompare) = 0 Then

            EncontrarLinhaRecurso = linha

            Exit Function

        End If

    Next linha

End Function


' ============================================================
' NORMALIZAR OEE
' ============================================================

Private Function NormalizarOEE( _
    ByVal valor As Double) As Double

    If valor <= 0 Then

        NormalizarOEE = 0

    ElseIf valor > 1 Then

        NormalizarOEE = valor / 100

    Else

        NormalizarOEE = valor

    End If

    ' Impedir OEE superior a 100%.

    If NormalizarOEE > 1 Then
        NormalizarOEE = 1
    End If

End Function


' ============================================================
' VERIFICAR UNIDADE
' ============================================================

Private Function UnidadeRepresentaCaixa( _
    ByVal unidade As String) As Boolean

    unidade = _
        UCase$( _
            Trim$(unidade))


    If unidade = "" Then

        ' Premissa padrão do APS:
        ' Quantidade já está em caixas.

        UnidadeRepresentaCaixa = True

        Exit Function

    End If


    If InStr(1, unidade, "CX", vbTextCompare) > 0 Then

        UnidadeRepresentaCaixa = True

        Exit Function

    End If


    If InStr(1, unidade, "CAIX", vbTextCompare) > 0 Then

        UnidadeRepresentaCaixa = True

        Exit Function

    End If


    If InStr(1, unidade, "BOX", vbTextCompare) > 0 Then

        UnidadeRepresentaCaixa = True

        Exit Function

    End If


    UnidadeRepresentaCaixa = False

End Function


' ============================================================
' OBTER PLANILHA
' ============================================================

Private Function ObterPlanilhaMotor( _
    ByVal nome As String) As Worksheet

    On Error Resume Next

    Set ObterPlanilhaMotor = _
        ThisWorkbook.Worksheets(nome)

    On Error GoTo 0

End Function


' ============================================================
' ENCONTRAR COLUNA POR CABEÇALHO
' ============================================================

Private Function EncontrarColunaMotor( _
    ByVal ws As Worksheet, _
    ByVal nome As String) As Long

    Dim ultimaColuna As Long
    Dim coluna As Long

    If ws Is Nothing Then Exit Function


    ultimaColuna = _
        ws.Cells( _
            1, _
            ws.Columns.Count).End(xlToLeft).Column


    For coluna = 1 To ultimaColuna

        If StrComp( _
            NormalizarCabecalho( _
                CStr(ws.Cells(1, coluna).Value)), _
            NormalizarCabecalho(nome), _
            vbTextCompare) = 0 Then

            EncontrarColunaMotor = coluna

            Exit Function

        End If

    Next coluna

End Function


' ============================================================
' NORMALIZAR CABEÇALHO
' ============================================================

Private Function NormalizarCabecalho( _
    ByVal texto As String) As String

    texto = Trim$(texto)

    texto = Replace(texto, vbCr, "")
    texto = Replace(texto, vbLf, "")

    NormalizarCabecalho = texto

End Function


' ============================================================
' GARANTIR COLUNA
' ============================================================

Private Sub GarantirColunaMotor( _
    ByVal ws As Worksheet, _
    ByVal nome As String)

    Dim coluna As Long
    Dim ultimaColuna As Long


    If ws Is Nothing Then Exit Sub


    coluna = _
        EncontrarColunaMotor( _
            ws, _
            nome)


    If coluna > 0 Then Exit Sub


    ultimaColuna = _
        ws.Cells( _
            1, _
            ws.Columns.Count).End(xlToLeft).Column


    If Trim$(CStr(ws.Cells(1, 1).Value)) = "" Then

        coluna = 1

    Else

        coluna = ultimaColuna + 1

    End If


    ws.Cells(1, coluna).Value = nome

End Sub


' ============================================================
' LER TEXTO
' ============================================================

Private Function ValorTextoMotor( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal campo As String) As String

    Dim coluna As Long


    If ws Is Nothing Then Exit Function

    If linha < 1 Then Exit Function


    coluna = _
        EncontrarColunaMotor( _
            ws, _
            campo)


    If coluna = 0 Then Exit Function


    If IsError(ws.Cells(linha, coluna).Value) Then Exit Function


    ValorTextoMotor = _
        Trim$( _
            CStr( _
                ws.Cells( _
                    linha, _
                    coluna).Value))

End Function


' ============================================================
' LER NÚMERO
' ============================================================

Private Function ValorNumeroMotor( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal campo As String) As Double

    Dim coluna As Long
    Dim valor As Variant


    If ws Is Nothing Then Exit Function

    If linha < 1 Then Exit Function


    coluna = _
        EncontrarColunaMotor( _
            ws, _
            campo)


    If coluna = 0 Then Exit Function


    valor = _
        ws.Cells( _
            linha, _
            coluna).Value


    If IsError(valor) Then Exit Function


    If IsNumeric(valor) Then

        ValorNumeroMotor = _
            CDbl(valor)

    Else

        ValorNumeroMotor = 0

    End If

End Function


' ============================================================
' LER NÚMERO COM INDICADOR DE VALIDADE
' ============================================================

Private Function LerNumeroSeguro( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal campo As String, _
    ByRef valido As Boolean) As Double

    Dim coluna As Long
    Dim valor As Variant


    valido = False


    If ws Is Nothing Then Exit Function


    coluna = _
        EncontrarColunaMotor( _
            ws, _
            campo)


    If coluna = 0 Then Exit Function


    valor = _
        ws.Cells( _
            linha, _
            coluna).Value


    If IsError(valor) Then Exit Function


    If IsNumeric(valor) Then

        LerNumeroSeguro = CDbl(valor)

        valido = True

    End If

End Function


' ============================================================
' LER NÚMERO DO RECURSO
' ============================================================

Private Function LerNumeroSeguroRecurso( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal campo As String, _
    ByRef valido As Boolean) As Double

    LerNumeroSeguroRecurso = _
        LerNumeroSeguro( _
            ws, _
            linha, _
            campo, _
            valido)

End Function


' ============================================================
' ESCREVER NÚMERO
' ============================================================

Private Sub EscreverNumeroMotor( _
    ByVal ws As Worksheet, _
    ByVal linha As Long, _
    ByVal campo As String, _
    ByVal valor As Double)

    Dim coluna As Long


    If ws Is Nothing Then Exit Sub

    If linha < 1 Then Exit Sub


    coluna = _
        EncontrarColunaMotor( _
            ws, _
            campo)


    If coluna = 0 Then

        GarantirColunaMotor _
            ws, _
            campo

        coluna = _
            EncontrarColunaMotor( _
                ws, _
                campo)

    End If


    If coluna > 0 Then

        ws.Cells( _
            linha, _
            coluna).Value = valor

    End If

End Sub


' ============================================================
' LIMPAR CÁLCULOS DA OP
' ============================================================

Private Sub LimparCalculosOP( _
    ByVal ws As Worksheet, _
    ByVal linha As Long)

    EscreverNumeroMotor _
        ws, linha, "Caixas", 0

    EscreverNumeroMotor _
        ws, linha, "Capacidade_h", 0

    EscreverNumeroMotor _
        ws, linha, "Produção_h", 0

    EscreverNumeroMotor _
        ws, linha, "Setup_h", 0

    EscreverNumeroMotor _
        ws, linha, COLUNA_DURACAO_BASE_MOTOR, 0

    RecalcularDuracaoOP _
        ws, _
        linha

End Sub


' ============================================================
' FIM DO MOTORCALCULO_APS
' ============================================================