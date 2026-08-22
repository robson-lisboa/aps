Option Explicit

' ============================================================
' GLOBAIS_APS
' APS PURAN - CONSTANTES GLOBAIS DO SISTEMA
' ============================================================

' ============================================================
' ABAS PRINCIPAIS
' ============================================================

Public Const ABA_DADOS As String = "DADOS"
Public Const ABA_PLANEJAMENTO As String = "PLANEJAMENTO"

' ============================================================
' IDENTIFICAÇÃO DOS CARDS
' ============================================================

Public Const PREFIXO_CARD As String = "APS_CARD_"

' ============================================================
' IDENTIFICAÇÃO DO ALTERNATETEXT
'
' Formato oficial:
'
' APS|OP|OP001|FETTE
'
' partes(0) = APS
' partes(1) = OP
' partes(2) = Número da OP
' partes(3) = Máquina
'
' ============================================================

Public Const ALT_IDX_OP As Long = 2
Public Const ALT_IDX_MAQUINA As Long = 3

' ============================================================
' STATUS OFICIAIS
' ============================================================

Public Const STATUS_PLANEJADO As String = "PLANEJADO"
Public Const STATUS_EM_PRODUCAO As String = "EM PRODUÇÃO"
Public Const STATUS_ATRASADO As String = "ATRASADO"
Public Const STATUS_CONCLUIDO As String = "CONCLUÍDO"

' ============================================================
' FIM DO MÓDULO
' ============================================================