Option Explicit

' ============================================================
' TESTE_FASE3
' FASE 3 — DADOS DE TESTE E VALIDAÇÃO
' ============================================================
'
' Responsável por:
'
'   • Criar dados de teste automáticos
'   • Executar fluxo completo
'   • Validar cards, timeline, atrasos, eventos
'   • Gerar relatório de testes
'
' NÃO altera regras de negócio.
' Apenas valida o sistema.
'
' ============================================================


' ============================================================
' EXECUTAR TESTE COMPLETO DA FASE 3
' ============================================================

Public Sub ExecutarTesteFase3()

    Dim resultado As String

    Dim inicio As Double

    Dim fim As Double


    inicio = Timer


    resultado = _
        "=== TESTE FASE 3 — APS PURAN ===" & vbCrLf & _
        "Data: " & Format(Now, "dd/mm/yyyy hh:mm") & vbCrLf & vbCrLf


    ' --------------------------------------------------------
    ' 1. Inicializar sistema
    ' --------------------------------------------------------

    resultado = resultado & _
        "1. INICIALIZAÇÃO" & vbCrLf

    On Error Resume Next

    IniciarSistemaAPS

    If Err.Number <> 0 Then

        resultado = resultado & _
            "   FAIL: " & Err.Description & vbCrLf

    Else

        resultado = resultado & _
            "   PASS" & vbCrLf

    End If

    On Error GoTo 0


    ' --------------------------------------------------------
    ' 2. Configurar botões
    ' --------------------------------------------------------

    resultado = resultado & vbCrLf & _
        "2. BOTÕES" & vbCrLf

    On Error Resume Next

    ConfigurarBotoesInicio

    If Err.Number <> 0 Then

        resultado = resultado & _
            "   FAIL: " & Err.Description & vbCrLf

    Else

        resultado = resultado & _
            "   PASS" & vbCrLf

    End If

    On Error GoTo 0


    ' --------------------------------------------------------
    ' 3. Criar dados de teste
    ' --------------------------------------------------------

    resultado = resultado & vbCrLf & _
        "3. DADOS DE TESTE" & vbCrLf

    Call CriarDadosTeste


    ' --------------------------------------------------------
    ' 4. Recalcular APS
    ' --------------------------------------------------------

    resultado = resultado & vbCrLf & _
        "4. RECALCULAR APS" & vbCrLf

    On Error Resume Next

    RecalcularAPSOperacional

    If Err.Number <> 0 Then

        resultado = resultado & _
            "   FAIL: " & Err.Description & vbCrLf

    Else

        resultado = resultado & _
            "   PASS" & vbCrLf

    End If

    On Error GoTo 0


    ' --------------------------------------------------------
    ' 5. Verificar cards
    ' --------------------------------------------------------

    resultado = resultado & vbCrLf & _
        "5. CARDS" & vbCrLf

    Call VerificarCards


    ' --------------------------------------------------------
    ' 6. Verificar timeline
    ' --------------------------------------------------------

    resultado = resultado & vbCrLf & _
        "6. TIMELINE" & vbCrLf

    Call VerificarTimeline


    ' --------------------------------------------------------
    ' 7. Teste de atraso
    ' --------------------------------------------------------

    resultado = resultado & vbCrLf & _
        "7. ATRASO" & vbCrLf

    Call TestarAtraso


    ' --------------------------------------------------------
    ' 8. Teste de manutenção
    ' --------------------------------------------------------

    resultado = resultado & vbCrLf & _
        "8. MANUTENÇÃO" & vbCrLf

    Call TestarManutencao


    ' --------------------------------------------------------
    ' 9. Teste de cadastro de máquina
    ' --------------------------------------------------------

    resultado = resultado & vbCrLf & _
        "9. CADASTRO MÁQUINA" & vbCrLf

    Call TestarCadastroMaquina


    ' --------------------------------------------------------
    ' 10. Teste de cadastro de OP
    ' --------------------------------------------------------

    resultado = resultado & vbCrLf & _
        "10. CADASTRO OP" & vbCrLf

    Call TestarCadastroOP


    ' --------------------------------------------------------
    ' Finalizar
    ' --------------------------------------------------------

    fim = Timer

    resultado = resultado & vbCrLf & _
        "=== TEMPO TOTAL: " & Format((fim - inicio) / 86400, "hh:mm:ss") & " ==="


    ' Salvar em arquivo de log

    Dim arquivoLog As String

    Dim numeroArquivo As Integer


    arquivoLog = _
        ThisWorkbook.Path & "\TESTE_FASE3_" & _
        Format(Now, "yyyymmdd_hhmmss") & ".txt"


    numeroArquivo = FreeFile


    On Error Resume Next

    Open arquivoLog For Output As numeroArquivo

    Print numeroArquivo, resultado

    Close numeroArquivo

    On Error GoTo 0


    MsgBox _
        "Teste Fase 3 concluído!" & vbCrLf & vbCrLf & _
        "Log salvo em:" & vbCrLf & _
        arquivoLog & vbCrLf & vbCrLf & _
        "Tempo total: " & Format((fim - inicio) / 86400, "hh:mm:ss"), _
        vbInformation, _
        "APS - Teste Fase 3"

End Sub


' ============================================================
' CRIAR DADOS DE TESTE
' ============================================================

Private Sub CriarDadosTeste()

    Dim wsDados As Worksheet
    Dim wsRecursos As Worksheet

    Dim ultimaLinha As Long

    Dim i As Long


    On Error Resume Next

    Set wsDados = _
        ThisWorkbook.Worksheets("DADOS")

    Set wsRecursos = _
        ThisWorkbook.Worksheets("RECURSOS")

    On Error GoTo 0

    If wsDados Is Nothing Or wsRecursos Is Nothing Then

        MsgBox _
            "Abas DADOS ou RECURSOS não encontradas.", _
            vbCritical, _
            "APS"

        Exit Sub

    End If


    ' --------------------------------------------------------
    ' Limpar dados existentes
    ' --------------------------------------------------------

    wsDados.Cells.Clear
    wsRecursos.Cells.Clear


    ' --------------------------------------------------------
    ' Cabeçalhos DADOS
    ' --------------------------------------------------------

    Dim cabecalhosDados As Variant

    Dim j As Long


    cabecalhosDados = Array( _
        "OP", "Produto", "Dosagem", "Quantidade", "Unidade", _
        "Caixas", "Máquina", "Sequência", "Data Planejada", _
        "Velocidade", "OEE", "Capacidade_h", "Produção_h", "Setup_h", _
        "Duração Base_h", "Duração Total (h)", "Início", "Fim", _
        "Status", "Atraso_h", "Eventos_h", "Refeição_h", _
        "Início Original", "Fim Original", "Observação")


    For j = LBound(cabecalhosDados) To UBound(cabecalhosDados)

        wsDados.Cells(1, j + 1).Value = _
            cabecalhosDados(j)

    Next j


    ' --------------------------------------------------------
    ' Cabeçalhos RECURSOS
    ' --------------------------------------------------------

    Dim cabecalhosRecursos As Variant


    cabecalhosRecursos = Array( _
        "Máquina", "Velocidade", "Unidade Velocidade", "OEE", _
        "Comprimidos por caixa", "Setup padrão", "Disponibilidade", "Observação")


    For j = LBound(cabecalhosRecursos) To UBound(cabecalhosRecursos)

        wsRecursos.Cells(1, j + 1).Value = _
            cabecalhosRecursos(j)

    Next j


    ' --------------------------------------------------------
    ' Dados de RECURSOS (3 máquinas)
    ' --------------------------------------------------------

    wsRecursos.Cells(2, 1).Value = "FETTE 2090"
    wsRecursos.Cells(2, 2).Value = 500
    wsRecursos.Cells(2, 3).Value = "cx/h"
    wsRecursos.Cells(2, 4).Value = 0.85
    wsRecursos.Cells(2, 5).Value = 20
    wsRecursos.Cells(2, 6).Value = 0.5
    wsRecursos.Cells(2, 7).Value = 0.9
    wsRecursos.Cells(2, 8).Value = "Envelope"

    wsRecursos.Cells(3, 1).Value = "FETTE 2"
    wsRecursos.Cells(3, 2).Value = 450
    wsRecursos.Cells(3, 3).Value = "cx/h"
    wsRecursos.Cells(3, 4).Value = 0.8
    wsRecursos.Cells(3, 5).Value = 20
    wsRecursos.Cells(3, 6).Value = 0.5
    wsRecursos.Cells(3, 7).Value = 0.9
    wsRecursos.Cells(3, 8).Value = "Envelope"

    wsRecursos.Cells(4, 1).Value = "MEDISEAL"
    wsRecursos.Cells(4, 2).Value = 600
    wsRecursos.Cells(4, 3).Value = "cx/h"
    wsRecursos.Cells(4, 4).Value = 0.75
    wsRecursos.Cells(4, 5).Value = 20
    wsRecursos.Cells(4, 6).Value = 0.5
    wsRecursos.Cells(4, 7).Value = 0.85
    wsRecursos.Cells(4, 8).Value = "Blister"


    ' --------------------------------------------------------
    ' Dados de DADOS (6 OPs)
    ' --------------------------------------------------------

    ' OP1 — FETTE 2090 — 08:00-10:00

    ultimaLinha = 2

    wsDados.Cells(ultimaLinha, 1).Value = "OP001"
    wsDados.Cells(ultimaLinha, 2).Value = "PURAN"
    wsDados.Cells(ultimaLinha, 3).Value = "25 mcg"
    wsDados.Cells(ultimaLinha, 4).Value = 1000
    wsDados.Cells(ultimaLinha, 5).Value = "COMP"
    wsDados.Cells(ultimaLinha, 7).Value = "FETTE 2090"
    wsDados.Cells(ultimaLinha, 8).Value = 1
    wsDados.Cells(ultimaLinha, 9).Value = DateTime(2026, 8, 25, 8, 0)
    wsDados.Cells(ultimaLinha, 17).Value = DateTime(2026, 8, 25, 8, 0)
    wsDados.Cells(ultimaLinha, 18).Value = DateTime(2026, 8, 25, 10, 0)


    ' OP2 — FETTE 2090 — 10:00-12:00

    ultimaLinha = 3

    wsDados.Cells(ultimaLinha, 1).Value = "OP002"
    wsDados.Cells(ultimaLinha, 2).Value = "PURAN"
    wsDados.Cells(ultimaLinha, 3).Value = "50 mcg"
    wsDados.Cells(ultimaLinha, 4).Value = 800
    wsDados.Cells(ultimaLinha, 5).Value = "COMP"
    wsDados.Cells(ultimaLinha, 7).Value = "FETTE 2090"
    wsDados.Cells(ultimaLinha, 8).Value = 2
    wsDados.Cells(ultimaLinha, 9).Value = DateTime(2026, 8, 25, 10, 0)
    wsDados.Cells(ultimaLinha, 17).Value = DateTime(2026, 8, 25, 10, 0)
    wsDados.Cells(ultimaLinha, 18).Value = DateTime(2026, 8, 25, 12, 0)


    ' OP3 — FETTE 2 — 08:00-11:00

    ultimaLinha = 4

    wsDados.Cells(ultimaLinha, 1).Value = "OP003"
    wsDados.Cells(ultimaLinha, 2).Value = "DIPIRONA"
    wsDados.Cells(ultimaLinha, 3).Value = "1g"
    wsDados.Cells(ultimaLinha, 4).Value = 500
    wsDados.Cells(ultimaLinha, 5).Value = "CX"
    wsDados.Cells(ultimaLinha, 7).Value = "FETTE 2"
    wsDados.Cells(ultimaLinha, 8).Value = 1
    wsDados.Cells(ultimaLinha, 9).Value = DateTime(2026, 8, 25, 8, 0)
    wsDados.Cells(ultimaLinha, 17).Value = DateTime(2026, 8, 25, 8, 0)
    wsDados.Cells(ultimaLinha, 18).Value = DateTime(2026, 8, 25, 11, 0)


    ' OP4 — FETTE 2 — 11:00-13:00

    ultimaLinha = 5

    wsDados.Cells(ultimaLinha, 1).Value = "OP004"
    wsDados.Cells(ultimaLinha, 2).Value = "DIPIRONA"
    wsDados.Cells(ultimaLinha, 3).Value = "500mg"
    wsDados.Cells(ultimaLinha, 4).Value = 300
    wsDados.Cells(ultimaLinha, 5).Value = "CX"
    wsDados.Cells(ultimaLinha, 7).Value = "FETTE 2"
    wsDados.Cells(ultimaLinha, 8).Value = 2
    wsDados.Cells(ultimaLinha, 9).Value = DateTime(2026, 8, 25, 11, 0)
    wsDados.Cells(ultimaLinha, 17).Value = DateTime(2026, 8, 25, 11, 0)
    wsDados.Cells(ultimaLinha, 18).Value = DateTime(2026, 8, 25, 13, 0)


    ' OP5 — MEDISEAL — 08:00-10:30

    ultimaLinha = 6

    wsDados.Cells(ultimaLinha, 1).Value = "OP005"
    wsDados.Cells(ultimaLinha, 2).Value = "AMOXIL"
    wsDados.Cells(ultimaLinha, 3).Value = "250mg"
    wsDados.Cells(ultimaLinha, 4).Value = 400
    wsDados.Cells(ultimaLinha, 5).Value = "CX"
    wsDados.Cells(ultimaLinha, 7).Value = "MEDISEAL"
    wsDados.Cells(ultimaLinha, 8).Value = 1
    wsDados.Cells(ultimaLinha, 9).Value = DateTime(2026, 8, 25, 8, 0)
    wsDados.Cells(ultimaLinha, 17).Value = DateTime(2026, 8, 25, 8, 0)
    wsDados.Cells(ultimaLinha, 18).Value = DateTime(2026, 8, 25, 10, 30)


    ' OP6 — MEDISEAL — 10:30-13:00

    ultimaLinha = 7

    wsDados.Cells(ultimaLinha, 1).Value = "OP006"
    wsDados.Cells(ultimaLinha, 2).Value = "AMOXIL"
    wsDados.Cells(ultimaLinha, 3).Value = "500mg"
    wsDados.Cells(ultimaLinha, 4).Value = 350
    wsDados.Cells(ultimaLinha, 5).Value = "CX"
    wsDados.Cells(ultimaLinha, 7).Value = "MEDISEAL"
    wsDados.Cells(ultimaLinha, 8).Value = 2
    wsDados.Cells(ultimaLinha, 9).Value = DateTime(2026, 8, 25, 10, 30)
    wsDados.Cells(ultimaLinha, 17).Value = DateTime(2026, 8, 25, 10, 30)
    wsDados.Cells(ultimaLinha, 18).Value = DateTime(2026, 8, 25, 13, 0)


    ' Formatar datas

    wsDados.Range( _
        wsDados.Cells(2, 17), _
        wsDados.Cells(7, 18)).NumberFormat = _
            "dd/mm/yyyy hh:mm"


    ' Salvar arquivo

    ThisWorkbook.Save


    MsgBox _
        "Dados de teste criados!" & vbCrLf & vbCrLf & _
        "✓ 3 máquinas cadastradas" & vbCrLf & _
        "✓ 6 OPs criadas" & vbCrLf & _
        "✓ Durações variadas" & vbCrLf & _
        "✓ Prioridades diferentes", _
        vbInformation, _
        "APS - Teste"

End Sub


' ============================================================
' VERIFICAR CARDS
' ============================================================

Private Sub VerificarCards()

    Dim wsPlan As Worksheet

    Dim quantidadeCards As Long

    Dim i As Long

    Dim nome As String


    On Error Resume Next

    Set wsPlan = _
        ThisWorkbook.Worksheets("PLANEJAMENTO")

    On Error GoTo 0

    If wsPlan Is Nothing Then

        MsgBox _
            "Aba PLANEJAMENTO não encontrada.", _
            vbCritical, _
            "APS"

        Exit Sub

    End If


    quantidadeCards = 0

    For i = 1 To wsPlan.Shapes.Count

        nome = wsPlan.Shapes(i).Name

        If Left(nome, 9) = "APS_CARD_" Then

            quantidadeCards = quantidadeCards + 1

        End If

    Next i


    If quantidadeCards >= 6 Then

        MsgBox _
            "CARDS OK" & vbCrLf & vbCrLf & _
            "Cards encontrados: " & quantidadeCards & vbCrLf & _
            "Esperado: mínimo 6", _
            vbInformation, _
            "APS"

    Else

        MsgBox _
            "CARDS FAIL" & vbCrLf & vbCrLf & _
            "Cards encontrados: " & quantidadeCards & vbCrLf & _
            "Esperado: mínimo 6", _
            vbCritical, _
            "APS"

    End If

End Sub


' ============================================================
' VERIFICAR TIMELINE
' ============================================================

Private Sub VerificarTimeline()

    Dim wsPlan As Worksheet

    Dim linhaMaquina As Long

    Dim totalMaquinas As Long


    On Error Resume Next

    Set wsPlan = _
        ThisWorkbook.Worksheets("PLANEJAMENTO")

    On Error GoTo 0

    If wsPlan Is Nothing Then

        MsgBox _
            "Aba PLANEJAMENTO não encontrada.", _
            vbCritical, _
            "APS"

        Exit Sub

    End If


    ' Contar máquinas na timeline (linhas a partir da 7)

    totalMaquinas = 0

    For linhaMaquina = 7 To 20

        If Trim(CStr( _
            wsPlan.Cells( _
                linhaMaquina, _
                1).Value)) <> "" Then

            totalMaquinas = totalMaquinas + 1

        End If

    Next linhaMaquina


    If totalMaquinas >= 3 Then

        MsgBox _
            "TIMELINE OK" & vbCrLf & vbCrLf & _
            "Máquinas encontradas: " & totalMaquinas & vbCrLf & _
            "Esperado: mínimo 3", _
            vbInformation, _
            "APS"

    Else

        MsgBox _
            "TIMELINE FAIL" & vbCrLf & vbCrLf & _
            "Máquinas encontradas: " & totalMaquinas & vbCrLf & _
            "Esperado: mínimo 3", _
            vbCritical, _
            "APS"

    End If

End Sub


' ============================================================
' TESTAR ATRASO
' ============================================================

Private Sub TestarAtraso()

    Dim wsDados As Worksheet

    Dim linhaOP1 As Long

    Dim linhaOP2 As Long

    Dim inicioOriginal As Date

    Dim fimOriginal As Date

    Dim inicioNovo As Date

    Dim fimNovo As Date


    On Error Resume Next

    Set wsDados = _
        ThisWorkbook.Worksheets("DADOS")

    On Error GoTo 0

    If wsDados Is Nothing Then

        MsgBox _
            "Aba DADOS não encontrada.", _
            vbCritical, _
            "APS"

        Exit Sub

    End If


    ' Encontrar OP1 e OP2

    linhaOP1 = EncontrarLinhaOP(wsDados, "OP001")
    linhaOP2 = EncontrarLinhaOP(wsDados, "OP002")


    If linhaOP1 = 0 Or linhaOP2 = 0 Then

        MsgBox _
            "OP001 ou OP002 não encontradas.", _
            vbCritical, _
            "APS"

        Exit Sub

    End If


    ' Registrar horários originais

    inicioOriginal = wsDados.Cells(linhaOP1, 17).Value
    fimOriginal = wsDados.Cells(linhaOP1, 18).Value


    ' Aplicar atraso de 2 horas na OP1

    wsDados.Cells(linhaOP1, 20).Value = 2


    ' Recalcular

    AtualizarAtrasosAPS


    ' Verificar novos horários

    inicioNovo = wsDados.Cells(linhaOP1, 17).Value
    fimNovo = wsDados.Cells(linhaOP1, 18).Value


    If fimNovo >= fimOriginal + (2 / 24) Then

        MsgBox _
            "ATRASO OK" & vbCrLf & vbCrLf & _
            "OP1: " & Format(inicioOriginal, "hh:mm") & " - " & Format(fimOriginal, "hh:mm") & vbCrLf & _
            "OP1 + atraso: " & Format(inicioNovo, "hh:mm") & " - " & Format(fimNovo, "hh:mm") & vbCrLf & vbCrLf & _
            "Verifique se OP2 foi deslocada.", _
            vbInformation, _
            "APS"

    Else

        MsgBox _
            "ATRASO FAIL" & vbCrLf & vbCrLf & _
            "OP1 não foi atualizada corretamente.", _
            vbCritical, _
            "APS"

    End If

End Sub


' ============================================================
' TESTAR MANUTENÇÃO
' ============================================================

Private Sub TestarManutencao()

    Dim wsEventos As Worksheet

    Dim ultimaLinha As Long


    On Error Resume Next

    Set wsEventos = _
        ThisWorkbook.Worksheets("EVENTOS")

    On Error GoTo 0

    If wsEventos Is Nothing Then

        MsgBox _
            "Aba EVENTOS não encontrada.", _
            vbCritical, _
            "APS"

        Exit Sub

    End If


    ' Limpar eventos anteriores

    wsEventos.Cells.Clear


    ' Cabeçalhos

    wsEventos.Cells(1, 1).Value = "OP"
    wsEventos.Cells(1, 2).Value = "Tipo"
    wsEventos.Cells(1, 3).Value = "Duração_h"
    wsEventos.Cells(1, 4).Value = "Ativo"
    wsEventos.Cells(1, 5).Value = "Aplicado"


    ' Evento de manutenção para todas as OPs da máquina FETTE 2090

    ' Duração de 2 horas (14:00-16:00)

    wsEventos.Cells(2, 1).Value = "OP001"
    wsEventos.Cells(2, 2).Value = "MANUTENÇÃO"
    wsEventos.Cells(2, 3).Value = 2
    wsEventos.Cells(2, 4).Value = "SIM"

    wsEventos.Cells(3, 1).Value = "OP002"
    wsEventos.Cells(3, 2).Value = "MANUTENÇÃO"
    wsEventos.Cells(3, 3).Value = 2
    wsEventos.Cells(3, 4).Value = "SIM"


    ' Aplicar eventos

    AplicarEventosAPS


    MsgBox _
        "MANUTENÇÃO CRIADA" & vbCrLf & vbCrLf & _
        "Evento: MANUTENÇÃO 14:00-16:00" & vbCrLf & _
        "Máquina: FETTE 2090" & vbCrLf & vbCrLf & _
        "Verifique se as OPs foram reposicionadas.", _
        vbInformation, _
        "APS"

End Sub


' ============================================================
' TESTAR CADASTRO DE MÁQUINA
' ============================================================

Private Sub TestarCadastroMaquina()

    Dim wsRecursos As Worksheet

    Dim ultimaLinha As Long


    On Error Resume Next

    Set wsRecursos = _
        ThisWorkbook.Worksheets("RECURSOS")

    On Error GoTo 0

    If wsRecursos Is Nothing Then

        MsgBox _
            "Aba RECURSOS não encontrada.", _
            vbCritical, _
            "APS"

        Exit Sub

    End If


    ' Adicionar nova máquina ENC-03

    ultimaLinha = wsRecursos.Cells( _
        wsRecursos.Rows.Count, _
        1).End(xlUp).Row + 1


    wsRecursos.Cells(ultimaLinha, 1).Value = "ENC-03"
    wsRecursos.Cells(ultimaLinha, 2).Value = 300
    wsRecursos.Cells(ultimaLinha, 3).Value = "cx/h"
    wsRecursos.Cells(ultimaLinha, 4).Value = 0.8
    wsRecursos.Cells(ultimaLinha, 5).Value = 20
    wsRecursos.Cells(ultimaLinha, 6).Value = 0.5
    wsRecursos.Cells(ultimaLinha, 7).Value = 0.9
    wsRecursos.Cells(ultimaLinha, 8).Value = "Envelope"


    ' Recalcular para incluir nova máquina

    ExecutarMotorAPS


    MsgBox _
        "MÁQUINA CADASTRADA" & vbCrLf & vbCrLf & _
        "Máquina: ENC-03" & vbCrLf & _
        "Velocidade: 300 cx/h" & vbCrLf & _
        "OEE: 80%" & vbCrLf & vbCrLf & _
        "Verifique se aparece na timeline.", _
        vbInformation, _
        "APS"

End Sub


' ============================================================
' TESTAR CADASTRO DE OP
' ============================================================

Private Sub TestarCadastroOP()

    Dim wsDados As Worksheet

    Dim ultimaLinha As Long


    On Error Resume Next

    Set wsDados = _
        ThisWorkbook.Worksheets("DADOS")

    On Error GoTo 0

    If wsDados Is Nothing Then

        MsgBox _
            "Aba DADOS não encontrada.", _
            vbCritical, _
            "APS"

        Exit Sub

    End If


    ' Adicionar nova OP

    ultimaLinha = wsDados.Cells( _
        wsDados.Rows.Count, _
        1).End(xlUp).Row + 1


    wsDados.Cells(ultimaLinha, 1).Value = "OP007"
    wsDados.Cells(ultimaLinha, 2).Value = "PARACETAMOL"
    wsDados.Cells(ultimaLinha, 3).Value = "750mg"
    wsDados.Cells(ultimaLinha, 4).Value = 600
    wsDados.Cells(ultimaLinha, 5).Value = "CX"
    wsDados.Cells(ultimaLinha, 7).Value = "ENC-03"
    wsDados.Cells(ultimaLinha, 8).Value = 1
    wsDados.Cells(ultimaLinha, 9).Value = DateTime(2026, 8, 25, 8, 0)
    wsDados.Cells(ultimaLinha, 17).Value = DateTime(2026, 8, 25, 8, 0)
    wsDados.Cells(ultimaLinha, 18).Value = DateTime(2026, 8, 25, 9, 30)


    ' Recalcular

    ExecutarMotorAPS

    CriarCardsAPS

    ConstruirTimelineAPS


    MsgBox _
        "OP CADASTRADA" & vbCrLf & vbCrLf & _
        "OP: OP007" & vbCrLf & _
        "Produto: PARACETAMOL 750mg" & vbCrLf & _
        "Máquina: ENC-03" & vbCrLf & vbCrLf & _
        "Verifique se aparece no planejamento.", _
        vbInformation, _
        "APS"

End Sub


' ============================================================
' FUNÇÕES AUXILIARES
' ============================================================

Private Function EncontrarLinhaOP( _
    ByVal ws As Worksheet, _
    ByVal numeroOP As String) As Long

    Dim colunaOP As Long

    Dim ultimaLinha As Long

    Dim linha As Long


    colunaOP = EncontrarColuna(ws, "OP")

    If colunaOP = 0 Then Exit Function


    ultimaLinha = ws.Cells( _
        ws.Rows.Count, _
        colunaOP).End(xlUp).Row


    For linha = 2 To ultimaLinha

        If StrComp( _
            Trim(CStr(ws.Cells(linha, colunaOP).Value)), _
            Trim(numeroOP), _
            vbTextCompare) = 0 Then

            EncontrarLinhaOP = linha

            Exit Function

        End If

    Next linha

End Function


Private Function EncontrarColuna( _
    ByVal ws As Worksheet, _
    ByVal nome As String) As Long

    Dim ultimaColuna As Long

    Dim coluna As Long


    ultimaColuna = ws.Cells( _
        1, _
        ws.Columns.Count).End(xlToLeft).Column


    For coluna = 1 To ultimaColuna

        If StrComp( _
            Trim(CStr(ws.Cells(1, coluna).Value)), _
            Trim(nome), _
            vbTextCompare) = 0 Then

            EncontrarColuna = coluna

            Exit Function

        End If

    Next coluna

End Function


' ============================================================
' FIM DO MÓDULO
' ============================================================
