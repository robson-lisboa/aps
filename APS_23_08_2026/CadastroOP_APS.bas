Option Explicit

' ============================================================
' CADASTRO_OP_APS
' Cadastro inline de OP na aba OPERACOES
' ============================================================

Public Const ABA_OPERACOES As String = "OPERACOES"
Public Const ABA_DADOS_CADASTRO As String = "DADOS"

Private Const LINHA_INICIO_CADASTRO As Long = 7
Private Const COLUNA_LABEL As Long = 1
Private Const COLUNA_VALOR As Long = 2


' ============================================================
' SALVAR NOVA OP
' ============================================================

Public Sub SalvarNovaOP()

    Dim wsOps As Worksheet
    Dim wsDados As Worksheet

    Dim numeroOP As String
    Dim produto As String
    Dim maquina As String
    Dim quantidade As Double
    Dim unidade As String

    Dim ultimaLinha As Long
    Dim colunaOP As Long
    Dim linha As Long

    Dim calcAnterior As XlCalculation
    Dim eventosAnterior As Boolean
    Dim telaAnterior As Boolean

    On Error GoTo TrataErro


    Set wsOps = _
        ThisWorkbook.Worksheets( _
            ABA_OPERACOES)


    Set wsDados = _
        ThisWorkbook.Worksheets( _
            ABA_DADOS_CADASTRO)


    calcAnterior = Application.Calculation
    eventosAnterior = Application.EnableEvents
    telaAnterior = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    numeroOP = Trim$(CStr( _
        wsOps.Cells( _
            LINHA_INICIO_CADASTRO, _
            COLUNA_VALOR).Value))


    produto = Trim$(CStr( _
        wsOps.Cells( _
            LINHA_INICIO_CADASTRO + 1, _
            COLUNA_VALOR).Value))


    maquina = Trim$(CStr( _
        wsOps.Cells( _
            LINHA_INICIO_CADASTRO + 2, _
            COLUNA_VALOR).Value))


    quantidade = ValorNumeroSeguro( _
        wsOps.Cells( _
            LINHA_INICIO_CADASTRO + 3, _
            COLUNA_VALOR).Value)


    unidade = Trim$(CStr( _
        wsOps.Cells( _
            LINHA_INICIO_CADASTRO + 4, _
            COLUNA_VALOR).Value))


    If numeroOP = "" Then

        MsgBox _
            "Informe o número da OP.", _
            vbExclamation, _
            "APS - Cadastro"

        GoTo SaidaSegura

    End If


    If produto = "" Then

        MsgBox _
            "Informe o produto.", _
            vbExclamation, _
            "APS - Cadastro"

        GoTo SaidaSegura

    End If


    If maquina = "" Then

        MsgBox _
            "Informe a máquina.", _
            vbExclamation, _
            "APS - Cadastro"

        GoTo SaidaSegura

    End If


    If quantidade <= 0 Then

        MsgBox _
            "Informe uma quantidade válida.", _
            vbExclamation, _
            "APS - Cadastro"

        GoTo SaidaSegura

    End If


    colunaOP = _
        EncontrarColunaCadastro( _
            wsDados, _
            "OP")


    If colunaOP = 0 Then

        MsgBox _
            "A coluna 'OP' não foi encontrada na aba DADOS.", _
            vbCritical, _
            "APS - Cadastro"

        GoTo SaidaSegura

    End If


    ultimaLinha = _
        wsDados.Cells( _
            wsDados.Rows.Count, _
            colunaOP).End(xlUp).Row + 1


    If ultimaLinha < 2 Then

        ultimaLinha = 2

    End If


    wsDados.Cells(ultimaLinha, colunaOP).Value = _
        numeroOP


    GarantirColunaCadastro wsDados, "Produto"
    wsDados.Cells(ultimaLinha, EncontrarColunaCadastro(wsDados, "Produto")).Value = _
        produto


    GarantirColunaCadastro wsDados, "Máquina"
    wsDados.Cells(ultimaLinha, EncontrarColunaCadastro(wsDados, "Máquina")).Value = _
        maquina


    GarantirColunaCadastro wsDados, "Quantidade"
    wsDados.Cells(ultimaLinha, EncontrarColunaCadastro(wsDados, "Quantidade")).Value = _
        quantidade


    GarantirColunaCadastro wsDados, "Unidade"
    wsDados.Cells(ultimaLinha, EncontrarColunaCadastro(wsDados, "Unidade")).Value = _
        unidade


    wsDados.Cells(ultimaLinha, EncontrarColunaCadastro(wsDados, "Quantidade")).NumberFormat = "0"


    LimparCamposCadastroOP wsOps


    MsgBox _
        "OP cadastrada com sucesso!" & vbCrLf & vbCrLf & _
        "OP: " & numeroOP & vbCrLf & _
        "Produto: " & produto & vbCrLf & _
        "Máquina: " & maquina & vbCrLf & _
        "Quantidade: " & Format(quantidade, "#,##0") & " " & unidade, _
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
        "Erro ao cadastrar OP:" & vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "APS - Cadastro"

End Sub


' ============================================================
' LIMPAR CAMPOS
' ============================================================

Private Sub LimparCamposCadastroOP( _
    ByVal ws As Worksheet)

    Dim i As Long
    Dim r As Long

    For i = 0 To 4

        r = LINHA_INICIO_CADASTRO + i

        ws.Cells(r, COLUNA_VALOR).Value = ""
        ws.Cells(r, COLUNA_VALOR).Interior.Color = _
            RGB(231, 230, 230)

    Next i

    ws.Cells(LINHA_INICIO_CADASTRO, COLUNA_VALOR).Select

End Sub


' ============================================================
' ENCONTRAR COLUNA
' ============================================================

Private Function EncontrarColunaCadastro( _
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

            EncontrarColunaCadastro = coluna
            Exit Function

        End If

    Next coluna

End Function


' ============================================================
' GARANTIR COLUNA
' ============================================================

Private Sub GarantirColunaCadastro( _
    ByVal ws As Worksheet, _
    ByVal nome As String)

    Dim coluna As Long

    coluna = EncontrarColunaCadastro(ws, nome)

    If coluna > 0 Then Exit Sub

    coluna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1

    ws.Cells(1, coluna).Value = nome
    ws.Cells(1, coluna).Font.Bold = True

End Sub


' ============================================================
' NÚMERO SEGURO
' ============================================================

Private Function ValorNumeroSeguro( _
    ByVal valor As Variant) As Double

    If IsNumeric(valor) Then

        ValorNumeroSeguro = CDbl(valor)

    Else

        ValorNumeroSeguro = 0

    End If

End Function


' ============================================================
' FIM DO MÓDULO
' ============================================================
