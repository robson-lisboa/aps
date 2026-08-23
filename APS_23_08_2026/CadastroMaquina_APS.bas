Option Explicit

' ============================================================
' CADASTRO_MAQUINA_APS
' Cadastro inline de máquina na aba MAQUINAS
' ============================================================

Public Const ABA_MAQUINAS As String = "MAQUINAS"
Public Const ABA_RECURSOS_CADASTRO As String = "RECURSOS"

Private Const LINHA_INICIO_CADASTRO_MAQ As Long = 5
Private Const COLUNA_LABEL_MAQ As Long = 1
Private Const COLUNA_VALOR_MAQ As Long = 2


' ============================================================
' SALVAR NOVA MAQUINA
' ============================================================

Public Sub SalvarNovaMaquina()

    Dim wsMaq As Worksheet
    Dim wsRec As Worksheet

    Dim codigo As String
    Dim nome As String
    Dim velocidade As Double
    Dim oee As Double
    Dim setup As Double

    Dim ultimaLinha As Long
    Dim colunaMaq As Long
    Dim linha As Long

    Dim calcAnterior As XlCalculation
    Dim eventosAnterior As Boolean
    Dim telaAnterior As Boolean

    On Error GoTo TrataErro


    Set wsMaq = _
        ThisWorkbook.Worksheets( _
            ABA_MAQUINAS)


    Set wsRec = _
        ThisWorkbook.Worksheets( _
            ABA_RECURSOS_CADASTRO)


    calcAnterior = Application.Calculation
    eventosAnterior = Application.EnableEvents
    telaAnterior = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    codigo = Trim$(CStr( _
        wsMaq.Cells( _
            LINHA_INICIO_CADASTRO_MAQ, _
            COLUNA_VALOR_MAQ).Value))


    nome = Trim$(CStr( _
        wsMaq.Cells( _
            LINHA_INICIO_CADASTRO_MAQ + 1, _
            COLUNA_VALOR_MAQ).Value))


    velocidade = ValorNumeroSeguroMaq( _
        wsMaq.Cells( _
            LINHA_INICIO_CADASTRO_MAQ + 2, _
            COLUNA_VALOR_MAQ).Value)


    oee = ValorNumeroSeguroMaq( _
        wsMaq.Cells( _
            LINHA_INICIO_CADASTRO_MAQ + 3, _
            COLUNA_VALOR_MAQ).Value)


    setup = ValorNumeroSeguroMaq( _
        wsMaq.Cells( _
            LINHA_INICIO_CADASTRO_MAQ + 4, _
            COLUNA_VALOR_MAQ).Value)


    If codigo = "" Then

        MsgBox _
            "Informe o código da máquina.", _
            vbExclamation, _
            "APS - Cadastro"

        GoTo SaidaSegura

    End If


    If nome = "" Then

        nome = codigo

    End If


    If velocidade <= 0 Then

        MsgBox _
            "Informe uma velocidade válida.", _
            vbExclamation, _
            "APS - Cadastro"

        GoTo SaidaSegura

    End If


    If oee <= 0 Or oee > 1 Then

        If oee > 1 Then

            oee = oee / 100

        End If

    End If


    If setup < 0 Then

        setup = 0

    End If


    colunaMaq = _
        EncontrarColunaMaquinaCadastro( _
            wsRec, _
            "Máquina")


    If colunaMaq = 0 Then

        MsgBox _
            "A coluna 'Máquina' não foi encontrada na aba RECURSOS.", _
            vbCritical, _
            "APS - Cadastro"

        GoTo SaidaSegura

    End If


    ultimaLinha = _
        wsRec.Cells( _
            wsRec.Rows.Count, _
            colunaMaq).End(xlUp).Row + 1


    If ultimaLinha < 2 Then

        ultimaLinha = 2

    End If


    wsRec.Cells(ultimaLinha, colunaMaq).Value = _
        codigo


    GarantirColunaMaquinaCadastro wsRec, "Velocidade"
    wsRec.Cells(ultimaLinha, EncontrarColunaMaquinaCadastro(wsRec, "Velocidade")).Value = _
        velocidade


    GarantirColunaMaquinaCadastro wsRec, "OEE"
    wsRec.Cells(ultimaLinha, EncontrarColunaMaquinaCadastro(wsRec, "OEE")).Value = _
        oee


    GarantirColunaMaquinaCadastro wsRec, "Setup padrão"
    wsRec.Cells(ultimaLinha, EncontrarColunaMaquinaCadastro(wsRec, "Setup padrão")).Value = _
        setup


    GarantirColunaMaquinaCadastro wsRec, "Observação"
    wsRec.Cells(ultimaLinha, EncontrarColunaMaquinaCadastro(wsRec, "Observação")).Value = _
        nome


    wsRec.Cells(ultimaLinha, EncontrarColunaMaquinaCadastro(wsRec, "Velocidade")).NumberFormat = "0.00"
    wsRec.Cells(ultimaLinha, EncontrarColunaMaquinaCadastro(wsRec, "OEE")).NumberFormat = "0%"
    wsRec.Cells(ultimaLinha, EncontrarColunaMaquinaCadastro(wsRec, "Setup padrão")).NumberFormat = "0.00"


    LimparCamposCadastroMaquina wsMaq


    MsgBox _
        "Máquina cadastrada com sucesso!" & vbCrLf & vbCrLf & _
        "Código: " & codigo & vbCrLf & _
        "Nome: " & nome & vbCrLf & _
        "Velocidade: " & Format(velocidade, "#,##0.00") & " cx/h" & vbCrLf & _
        "OEE: " & Format(oee, "0%") & vbCrLf & _
        "Setup: " & Format(setup, "0.00") & " h", _
        vbInformation, _
        "APS - Cadastro"


SaidaSegura:

    Application.Calculation = calcAnterior
    Application.EnableEvents = eventosAnterior
    Application.ScreenUpdating = telaAnterior

    Exit Sub


TrataErro:

    Application.Calculation = calcAnterior
    Application.EnableEvents = eventosAnterior
    Application.ScreenUpdating = telaAnterior

    MsgBox _
        "Erro ao cadastrar máquina:" & vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "APS - Cadastro"

End Sub


' ============================================================
' LIMPAR CAMPOS
' ============================================================

Private Sub LimparCamposCadastroMaquina( _
    ByVal ws As Worksheet)

    Dim i As Long
    Dim r As Long

    For i = 0 To 4

        r = LINHA_INICIO_CADASTRO_MAQ + i

        ws.Cells(r, COLUNA_VALOR_MAQ).Value = ""
        ws.Cells(r, COLUNA_VALOR_MAQ).Interior.Color = _
            RGB(231, 230, 230)

    Next i

    ws.Cells(LINHA_INICIO_CADASTRO_MAQ, COLUNA_VALOR_MAQ).Select

End Sub


' ============================================================
' ENCONTRAR COLUNA
' ============================================================

Private Function EncontrarColunaMaquinaCadastro( _
    ByVal ws As Worksheet, _
    ByVal nome As String) As Long

    Dim coluna As Long
    Dim ultimaColuna As Long

    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For coluna = 1 To ultimaColuna

        If StrComp( _
            Trim$(CStr(ws.Cells(1, coluna).Value)), _
            Trim$(nome), _
            vbTextCompare) = 0 Then

            EncontrarColunaMaquinaCadastro = coluna
            Exit Function

        End If

    Next coluna

End Function


' ============================================================
' GARANTIR COLUNA
' ============================================================

Private Sub GarantirColunaMaquinaCadastro( _
    ByVal ws As Worksheet, _
    ByVal nome As String)

    Dim coluna As Long

    coluna = EncontrarColunaMaquinaCadastro(ws, nome)

    If coluna > 0 Then Exit Sub

    coluna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1

    ws.Cells(1, coluna).Value = nome
    ws.Cells(1, coluna).Font.Bold = True

End Sub


' ============================================================
' NÚMERO SEGURO
' ============================================================

Private Function ValorNumeroSeguroMaq( _
    ByVal valor As Variant) As Double

    If IsNumeric(valor) Then

        ValorNumeroSeguroMaq = CDbl(valor)

    Else

        ValorNumeroSeguroMaq = 0

    End If

End Function


' ============================================================
' FIM DO MÓDULO
' ============================================================
