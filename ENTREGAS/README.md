# ENTREGAS — APS PURAN

Esta pasta contém os arquivos finais do projeto APS PURAN.

## Arquivos

| Arquivo | Descrição |
|---------|-----------|
| `APS_PURAN_Fase2.xlsx` | Planilha base com estrutura de abas e interface visual |
| `VBA_IMPORTAR/` | Pasta com todos os módulos `.bas` prontos para importar |

## Como montar o arquivo `.xlsm`

1. Abra `APS_PURAN_Fase2.xlsx` no Excel Desktop.
2. Habilite as macros quando solicitado.
3. Pressione `Alt+F11` para abrir o VBA Editor.
4. Importe os módulos na ordem abaixo:
   1. `Globais_APS.bas`
   2. `motor_APS.bas`
   3. `MotorCalculo_APS.bas`
   4. `Atrasos_APS.bas`
   5. `Eventos_APS.bas`
   6. `Refeição_fixa_.bas`
   7. `Cards_APS.bas`
   8. `Timeline_APS.bas`
   9. `Dashboard_APS.bas`
   10. `DragCards_APS.bas`
   11. `Integracao_APS.bas`
   12. `Final_APS.bas`
   13. `Navegacao_APS.bas`
   14. `Interface_APS.bas`
   15. `Teste_Fase3.bas`
5. No menu **Depurar**, clique em **Compilar VBAProject**.
6. Salve o arquivo como **Pasta de trabalho habilitada para macro do Excel (*.xlsm)**.
7. Feche e reabra o arquivo.
8. Habilite macros novamente.
9. Na janela **Imediata** (`Ctrl+G`), execute:
   ```vba
   ExecutarTesteFase3
   ```

## Fluxo de teste recomendado

1. `ExecutarTesteFase3` — cria dados de teste e executa validação completa
2. Verifique os cards na aba **PLANEJAMENTO**
3. Teste o drag & drop dos cards
4. Teste atrasos e propagação
5. Teste eventos/manutenção
6. Teste cadastro de máquina e OP

## Estrutura de abas

| Aba visível | Função |
|-------------|--------|
| **INICIO** | Dashboard operacional e menu principal |
| **OPs** | Cadastro e consulta de ordens de produção |
| **MAQUINAS** | Cadastro de máquinas/recursos |
| **PLANEJAMENTO** | Timeline horizontal, cards, drag & drop |

Abas técnicas ocultas:
- `DADOS`
- `RECURSOS`
- `EVENTOS`
- `REFEICAO`
- `RESUMO`
- `CONFIG`

## Suporte

Consulte os arquivos de documentação do projeto para detalhes técnicos.
