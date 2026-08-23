Option Explicit

' ============================================================
' MÓDULO 1 — CONFIGURAÇÃO DO APS (LEGADO)
' ============================================================
'
' ATENÇÃO: Este módulo é mantido para compatibilidade.
' O fluxo principal do APS PURAN não utiliza ConfigurarAPS.
' Use IniciarSistemaAPS ou AtualizarAPS como pontos de entrada.
'
' ============================================================

Public Sub ConfigurarAPS()

    Application.ScreenUpdating = False

    CriarAbaSeNaoExistir "RECURSOS"
    CriarAbaSeNaoExistir "DADOS"
    CriarAbaSeNaoExistir "PLANEJAMENTO"
    CriarAbaSeNaoExistir "RESUMO"

    PrepararRecursos
    PrepararDados
    PrepararPlanejamento

    Application.ScreenUpdating = True

    MsgBox "Estrutura do APS preparada com sucesso!", _
           vbInformation, "APS"

End Sub


' ============================================================
' CRIAR ABA
' ============================================================

Private Sub CriarAbaSeNaoExistir(ByVal NomeAba As String)

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NomeAba)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        ws.Name = NomeAba

    End If

End Sub


' ============================================================
' ABA RECURSOS
' ============================================================

Private Sub PrepararRecursos()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("RECURSOS")

    ws.Range("A1:H1").Value = Array( _
        "Máquina", _
        "Velocidade", _
        "Unidade Velocidade", _
        "OEE", _
        "Comprimidos por caixa", _
        "Setup padrão", _
        "Disponibilidade", _
        "Observação")

    ws.Rows(1).Font.Bold = True

    ws.Columns("A:H").AutoFit

End Sub


' ============================================================
' ABA DADOS
' ============================================================

Private Sub PrepararDados()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("DADOS")

    ws.Range("A1:U1").Value = Array( _
        "OP", _
        "Produto", _
        "Dosagem", _
        "Quantidade", _
        "Unidade", _
        "Caixas", _
        "Máquina", _
        "Sequência", _
        "Data Planejada", _
        "Velocidade", _
        "OEE", _
        "Capacidade_h", _
        "Produção_h", _
        "Setup_h", _
        "Duração Total (h)", _
        "Início", _
        "Fim", _
        "Status", _
        "Atraso_h", _
        "Refeição_h", _
        "Observação")

    ws.Rows(1).Font.Bold = True

    ws.Columns("A:U").AutoFit

End Sub


' ============================================================
' ABA PLANEJAMENTO
' ============================================================

Private Sub PrepararPlanejamento()

    Dim ws As Worksheet
    Dim i As Long
    Dim hora As Date
    Dim resposta As VbMsgBoxResult

    Set ws = ThisWorkbook.Worksheets("PLANEJAMENTO")

    If Trim$(CStr(ws.Range("A1").Value)) <> "" Then

        resposta = MsgBox( _
            "A aba PLANEJAMENTO já possui conteúdo." & _
            vbCrLf & vbCrLf & _
            "Deseja reconfigurar a aba? Todos os dados atuais serão perdidos.", _
            vbYesNo + vbExclamation, _
            "Configurar APS")

        If resposta <> vbYes Then

            Exit Sub

        End If

    End If

    ws.Cells.Clear

    ws.Range("A1").Value = _
        "APS — PLANEJAMENTO DA PRODUÇÃO"

    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    ws.Range("A4").Value = "MÁQUINA"

    ws.Range("A4").Font.Bold = True

    hora = TimeSerial(5, 0, 0)

    For i = 2 To 49

        ws.Cells(4, i).Value = hora
        ws.Cells(4, i).NumberFormat = "hh:mm"
        ws.Cells(4, i).Font.Bold = True

        hora = DateAdd("n", 30, hora)

    Next i

    ws.Range("A5").Value = "FETTE 2090"
    ws.Range("A6").Value = "FETTE 2"
    ws.Range("A7").Value = "MEDISEAL"
    ws.Range("A8").Value = "BLISTERFLEX"

    ws.Columns("A").ColumnWidth = 20

    For i = 2 To 49
        ws.Columns(i).ColumnWidth = 9
    Next i

    ws.Rows("5:8").RowHeight = 60
End Sub


' ============================================================
' FIM DO MÓDULO 1
' ============================================================