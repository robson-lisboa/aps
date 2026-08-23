============================================================
APS PURAN — SISTEMA DE PLANEJAMENTO E CONTROLE DA PRODUCAO
Versao: 2.0
Data: 23/08/2026
============================================================

DESCRICAO
---------
O APS PURAN e um sistema de Planejamento e Escalonamento
Avançado de Producao (APS - Advanced Planning and Scheduling)
desenvolvido para industrias farmaceuticas.

O sistema transforma o Excel Desktop em uma plataforma de
planejamento visual com timeline horizontal, cards flutuantes,
drag & drop, propagacao automatica de atrasos e dashboard
operacional.

NAO é uma planilha comum. E um sistema APS dentro do Excel.

TELAS PRINCIPAIS
----------------
1. INICIO
   - Dashboard operacional
   - KPIs: Total OPs, Em Producao, Atrasadas, Concluidas
   - Horas planejadas e horas de atraso
   - Alertas e notificacoes
   - Menu de navegacao

2. OPERACOES
   - Cadastro inline de OP
   - Pesquisa por OP, produto, maquina, status
   - Visualizacao de status
   - Sem criacao de novas abas

3. MAQUINAS
   - Cadastro inline de maquina
   - Pesquisa por codigo ou nome
   - Visualizacao de capacidade e utilizacao
   - Sem criacao de novas abas

4. PLANEJAMENTO
   - Timeline horizontal (05:00 - 22:00)
   - Cards das OPs posicionados por horario e maquina
   - Drag & Drop funcional
   - Botoes: Recalcular, Atrasos, Eventos, Refeicoes, Dashboard

ABAS TECNICAS (ocultas)
-----------------------
- DADOS
- RECURSOS
- EVENTOS
- REFEICAO
- RESUMO
- CONFIG

MODULOS VBA (17)
----------------
1. Globais_APS.bas          - Constantes globais
2. motor_APS.bas            - Configuracao legado
3. MotorCalculo_APS.bas     - Motor de calculo + sequenciamento
4. Atrasos_APS.bas          - Atrasos e propagacao
5. Eventos_APS.bas          - Eventos/manutencao
6. Refeicao_fixa_.bas       - Refeicoes fixas
7. Cards_APS.bas            - Cards flutuantes + remocao orfaos
8. Timeline_APS.bas         - Timeline horizontal
9. Dashboard_APS.bas        - Dashboard e KPIs
10. DragCards_APS.bas       - Drag & Drop real
11. Integracao_APS.bas      - Integracao e fluxos principais
12. Final_APS.bas           - Inicializacao e validacao
13. Navegacao_APS.bas       - Navegacao entre telas
14. Interface_APS.bas       - Interface e botoes
15. Teste_Fase3.bas         - Testes automaticos
16. CadastroOP_APS.bas      - Cadastro inline de OP
17. CadastroMaquina_APS.bas - Cadastro inline de maquina

REQUISITOS
----------
- Excel Desktop (Windows)
- Macros habilitadas
- VBA habilitado

INSTALACAO
----------
1. Abra APS_PURAN_FINAL.xlsx no Excel Desktop
2. Habilite as macros
3. Pressione Alt+F11 para abrir o VBA Editor
4. Importe os modulos .bas na ordem do arquivo
   INSTRUCOES_EXCEL_DESKTOP.txt
5. Compile o projeto (Debug > Compile)
6. Salve como .xlsm
7. Execute InicializarAPS na janela Imediata

LIMITACOES
----------
- Validacao visual e execucao real no Excel Desktop ainda
  dependem do Excel Desktop.
- Este ambiente Linux nao permite execucao de VBA.
- Arquivo .xlsm nao pode ser gerado via Python/openpyxl.
- O usuario deve salvar como .xlsm apos importar o VBA.

SUPORTE
-------
Documentacao: MANUAL_INSTALACAO.md, MANUAL_USUARIO.md
Testes: Teste_Fase3.bas
Relatorio: RELATORIO_FINAL.md

============================================================
APS PURAN — Sistema de Planejamento e Controle da Producao
Versao: 2.0
Data: 23/08/2026
Desenvolvido para industria farmaceutica
============================================================
