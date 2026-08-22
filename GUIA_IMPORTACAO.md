# GUIA DE IMPORTAÇÃO E EXECUÇÃO — APS PURAN

**Arquivo base:** `APS_PURAN_Base.xlsx`  
**Módulos VBA:** pasta `/workspaces/aps/*.bas` e `*.txt`  
**Data:** 2026-08-22

---

## PASSO 1 — ABRIR O ARQUIVO BASE

1. Abra `APS_PURAN_Base.xlsx` no **Excel Desktop** (Windows/Mac).
2. Verifique se as 6 abas estão presentes:
   - `DADOS`
   - `PLANEJAMENTO`
   - `RESUMO`
   - `RECURSOS`
   - `EVENTOS`
   - `REFEICAO`
3. Se alguma aba estiver faltando, recrie manualmente com o mesmo nome.

---

## PASSO 2 — HABILITAR MACROS E ABRIR O VBA EDITOR

1. Clique em **Conteúdo** ou **Habilitar Edição** na barra amarela (se aparecer).
2. Pressione `Alt + F11` para abrir o **VBA Editor**.
3. No menu **Ferramentas** → **Referências**, verifique se não há referências quebradas.
4. No **Gerenciador de Projeto** (painel esquerdo), clique com o botão direito no nome do projeto (normalmente `VBAProject (APS_PURAN_Base.xlsx)`).
5. Escolha **Inserir** → **Módulo**.
6. Repita para cada módulo listado abaixo.

---

## PASSO 3 — IMPORTAR MÓDULOS VBA

Importe **exatamente** os arquivos abaixo, na ordem indicada:

| # | Arquivo | Tipo | Descrição |
|---|---------|------|-----------|
| 1 | `Globais_APS.bas` | `.bas` | Constantes globais |
| 2 | `motor_APS.txt` | `.txt` → `.bas` | Configuração legado |
| 3 | `MotorCalculo_APS.bas` | `.bas` | Motor de cálculo |
| 4 | `Atrasos_APS.txt` | `.txt` → `.bas` | Módulo de atrasos |
| 5 | `Eventos_APS.bas` | `.bas` | Módulo de eventos |
| 6 | `Cards_APS.txt` | `.txt` → `.bas` | Cards flutuantes |
| 7 | `Timeline_APS.txt` | `.txt` → `.bas` | Timeline horizontal |
| 8 | `Dashboard_APS.txt` | `.txt` → `.bas` | Dashboard |
| 9 | `Refeição_fixa_.bas` | `.bas` | Refeições fixas |
| 10 | `DragCards_APS.txt` | `.txt` → `.bas` | Drag & Drop |
| 11 | `Integracao_APS.txt` | `.txt` → `.bas` | Integração final |
| 12 | `Final_APS.txt` | `.txt` → `.bas` | Final/validação |

**Como importar:**
- Arquivos `.bas`: clique direito em módulo → **Importar arquivo** → selecione o `.bas`.
- Arquivos `.txt`: clique direito em módulo → **Importar arquivo** → selecione o `.txt`. O Excel converte automaticamente para `.bas`.

---

## PASSO 4 — COMPILAR O PROJETO

1. No VBA Editor, vá em **Depurar** → **Compilar VBAProject**.
2. Se aparecer alguma mensagem de erro, anote o procedimento e a linha.
3. Se compilar sem erros, prossiga.

---

## PASSO 5 — SALVAR COMO .XLSM

1. Feche o VBA Editor.
2. No Excel, clique em **Arquivo** → **Salvar Como**.
3. Em **Tipo**, selecione **Pasta de trabalho habilitada para macro do Excel (*.xlsm)**.
4. Salve como `APS_PURAN.xlsm`.
5. Feche o arquivo.

---

## PASSO 6 — EXECUTAR A MACRO DE INICIALIZAÇÃO

1. Reabra `APS_PURAN.xlsm`.
2. Habilite as macros quando solicitado.
3. Pressione `Alt + F11` para abrir o VBA Editor.
4. Pressione `Ctrl + G` para abrir a janela **Imediata**.
5. Digite:
   ```
   InicializarAPS
   ```
6. Pressione `Enter`.

**Resultado esperado:**
- Aba `PLANEJAMENTO` recebe 9 botões.
- Dashboard é atualizado.
- Botões aparecem na linha superior da aba PLANEJAMENTO.

---

## PASSO 7 — TESTAR FUNCIONALIDADES PRINCIPAIS

### 7.1 — IniciarSistemaAPS
```
IniciarSistemaAPS
```
Esperado: Estrutura validada, motor executado, timeline criada, cards criados, dashboard criado, botões preparados.

### 7.2 — AtualizarAPS
```
AtualizarAPS
```
Esperado: Dados recalculados, timeline atualizada, cards atualizados, dashboard atualizado.

### 7.3 — ExecutarMotorAPS
```
ExecutarMotorAPS
```
Esperado: Colunas Caixas, Capacidade_h, Produção_h, Setup_h, Duração Base_h, Duração Total (h) preenchidas.

### 7.4 — ConstruirTimelineAPS
```
ConstruirTimelineAPS
```
Esperado: Cabeçalho com horários de 05:00 a 22:00, máquinas listadas.

### 7.5 — CriarCardsAPS
```
CriarCardsAPS
```
Esperado: Cards flutuantes criados na aba PLANEJAMENTO para cada OP.

### 7.6 — AtualizarAtrasosAPS
```
AtualizarAtrasosAPS
```
Esperado: Atrasos aplicados, cards atualizados.

### 7.7 — AplicarEventosAPS
```
AplicarEventosAPS
```
Esperado: Eventos aplicados, sequência recalculada, cards atualizados.

### 7.8 — DesenharRefeicoes
```
DesenharRefeicoes
```
Esperado: Cards de refeição criados na aba PLANEJAMENTO.

### 7.9 — AplicarAlteracoesCards
```
AplicarAlteracoesCards
```
Esperado: Posição dos cards convertida em horário/máquina.

---

## PASSO 8 — VERIFICAR BOTÕES

Na aba `PLANEJEAMENTO`, verifique se existem 9 botões:

| Botão | Macro associada |
|-------|-----------------|
| ATUALIZAR APS | `AtualizarAPS` |
| APLICAR CARDS | `AplicarAlteracoesCards` |
| TIMELINE | `ConstruirTimelineAPS` |
| DASHBOARD | `AtualizarDashboardAPS` |
| LIMPAR CARDS | `ApagarCardsAPS` |
| ATRASOS | `AtualizarAtrasosAPS` |
| EVENTOS | `AplicarEventosAPS` |
| REFEIÇÃO | `DesenharRefeicoes` |
| MOTOR | `ExecutarMotorAPS` |

Clique em cada botão e verifique se a macro é executada sem erro.

---

## PASSO 9 — TESTE DE PERSISTÊNCIA

1. Execute `AtualizarAPS`.
2. Salve o arquivo.
3. Feche o Excel.
4. Reabra o arquivo.
5. Habilite macros.
6. Execute `AtualizarAPS` novamente.

Esperado: Dados mantidos, timeline reconstruída, cards recriados, dashboard atualizado.

---

## SOLUÇÃO DE PROBLEMAS COMUNS

| Problema | Solução |
|-----------|---------|
| "Macro desabilitada" | Arquivo → Opções → Central de Confiança → Configurações da Central de Confiança → Habilitar macros |
| Erro de compilação | Verifique se todos os 12 módulos foram importados |
| Botões não aparecem | Execute `CriarBotoesAPS` |
| Abas não existem | Recrie manualmente com os nomes corretos |
| Referência quebrada | Verifique referências em Ferramentas → Referências |

---

## ESTRUTURA DO ARQUIVO FINAL

```
APS_PURAN.xlsm
├── DADOS           (OPs, cálculos, horários)
├── PLANEJAMENTO    (Timeline + cards + botões)
├── RESUMO          (Dashboard + KPIs)
├── RECURSOS        (Máquinas, velocidade, OEE)
├── EVENTOS         (Eventos adicionais)
├── REFEICAO        (Refeições fixas)
└── Módulos VBA     (12 módulos importados)
```

---

## NOTAS IMPORTANTES

- **Não renomeie as abas** — o VBA depende dos nomes exatos.
- **Não altere a lógica do VBA** — a interface é apenas visual.
- **Sempre salve como .xlsm** — macros não funcionam em .xlsx.
- **Execute macros apenas no Excel Desktop** — Excel Online não suporta VBA.

---

## PRÓXIMOS PASSOS SUGERIDOS

1. Formatar DADOS como Tabela Excel (Ctrl+T) para facilitar manutenção.
2. Adicionar proteção nas abas de dados para evitar edição acidental.
3. Criar um manual do usuário com prints da interface.
