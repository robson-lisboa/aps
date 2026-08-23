# AUDITORIA DE ARQUITETURA — FASE 2 RECONSTRUÇÃO DA INTERFACE

**Data:** 2026-08-23  
**Branch:** main (commit 1508ec9)  
**Objetivo:** Mapear arquitetura atual para reconstrução da UX/UI sem alterar regras de negócio

---

## 1. MAPA DE MÓDULOS E RESPONSABILIDADES

| Módulo | Tipo | Responsabilidade principal | Dependências |
|--------|------|---------------------------|--------------|
| `Globais_APS.bas` | Módulo | Constantes globais (abas, status, prefixos) | Nenhuma |
| `motor_APS.bas` | Módulo | Configuração legado, criação de estrutura inicial | Nenhuma |
| `MotorCalculo_APS.bas` | Módulo | Motor de cálculo (Caixas, Capacidade_h, Produção_h, Setup_h, Duração Base_h, Duração Total) | Globais_APS, DADOS, RECURSOS |
| `Atrasos_APS.bas` | Módulo | Atrasos, cards de atraso, atualização visual | Globais_APS, MotorCalculo_APS |
| `Eventos_APS.bas` | Módulo | Eventos de produção, sequenciamento, status | Globais_APS, MotorCalculo_APS |
| `Refeição_fixa_.bas` | Módulo | Refeições fixas, cards de refeição | Globais_APS |
| `Cards_APS.bas` | Módulo | Criação/atualização de cards flutuantes | Globais_APS, DADOS, PLANEJAMENTO |
| `Timeline_APS.bas` | Módulo | Timeline horizontal, escala de horários | Globais_APS, PLANEJAMENTO |
| `Dashboard_APS.bas` | Módulo | Dashboard, KPIs, gráficos | Globais_APS, DADOS, RESUMO |
| `DragCards_APS.bas` | Módulo | Drag & Drop, conversão posição → horário/máquina | Globais_APS, Cards_APS |
| `Integracao_APS.bas` | Módulo | Integração, botões, fluxos principais | Todos |
| `Final_APS.bas` | Módulo | Inicialização, validação, estrutura | Globais_APS, Todos |

---

## 2. DEPENDÊNCIAS E ORDEM DE IMPORTAÇÃO

### Ordem correta de importação no VBA Editor:

1. `Globais_APS.bas` — primeiro (constantes globais)
2. `motor_APS.bas` — legado, não depende de Globais
3. `MotorCalculo_APS.bas` — depende de Globais
4. `Atrasos_APS.bas` — depende de Globais + MotorCalculo
5. `Eventos_APS.bas` — depende de Globais + MotorCalculo
6. `Refeição_fixa_.bas` — depende de Globais
7. `Cards_APS.bas` — depende de Globais
8. `Timeline_APS.bas` — depende de Globais
9. `Dashboard_APS.bas` — depende de Globais
10. `DragCards_APS.bas` — depende de Globais + Cards
11. `Integracao_APS.bas` — depende de todos
12. `Final_APS.bas` — depende de Globais + todos

### Dependências de tempo de execução:

```
IniciarSistemaAPS / AtualizarAPS
    ├── GarantirEstrutura (cria abas)
    ├── Application.Calculate
    ├── ExecutarMotorAPS (MotorCalculo_APS)
    │   ├── ObterPlanilhaMotor(ABA_DADOS)
    │   ├── ObterPlanilhaMotor(ABA_RECURSOS)
    │   ├── CalcularOP (para cada OP)
    │   └── RecalcularDuracaoOP
    ├── ConstruirTimelineAPS (Timeline_APS)
    │   ├── LimparTimeline
    │   ├── CriarCabecalhoTimeline
    │   ├── CriarEscalaHorario
    │   ├── CriarMarcadoresDeHora
    │   ├── CriarLinhasMaquinas
    │   ├── CriarMaquinasDosRecursos
    │   ├── FormatarTimeline
    │   └── CongelarTimeline
    ├── ApagarCardsAPS (Cards_APS)
    ├── CriarCardsAPS (Cards_APS)
    │   ├── PrepararPlanejamento
    │   ├── CriarCardsDasOPs
    │   │   ├── MontarTextoCard
    │   │   ├── ConfigurarCard
    │   │   └── AplicarVisualCards
    │   └── AtualizarCardsExistentes
    ├── CriarDashboardAPS / AtualizarDashboardAPS (Dashboard_APS)
    │   ├── LimparDashboard
    │   ├── CriarTituloDashboard
    │   ├── CriarCardsKPI
    │   ├── CriarKPI
    │   ├── AtualizarKPIsDashboard
    │   ├── CriarTabelaMaquinas
    │   ├── CriarTabelaStatus
    │   ├── AtualizarTabelaStatusDashboard
    │   ├── CriarGraficoStatus
    │   ├── CriarGraficoMaquinas
    │   ├── ApagarGrafico
    │   └── AtualizarGraficosDashboard
    └── CriarBotoesAPS (Integracao_APS)
        ├── ApagarBotoesAPS
        └── CriarBotao (×9)
```

---

## 3. FLUXOS DE EXECUÇÃO PRINCIPAIS

### Fluxo 1: Inicialização completa
```
IniciarSistemaAPS (Final_APS)
    ├── PrepararAbas (DADOS, RECURSOS, PLANEJAMENTO, RESUMO)
    ├── PrepararDados (cabeçalhos + autofiltro)
    ├── PrepararPlanejamento (cabeçalho mínimo)
    ├── PrepararResumo (título)
    ├── ValidarSistemaAPS
    ├── ExecutarMotorAPS
    ├── ConstruirTimelineAPS
    ├── ApagarCardsAPS + CriarCardsAPS
    ├── CriarDashboardAPS
    └── CriarBotoesAPS
```

### Fluxo 2: Atualização completa
```
AtualizarAPS (Integracao_APS)
    ├── GarantirEstrutura
    ├── Application.Calculate
    ├── ExecutarMotorAPS
    ├── ConstruirTimelineAPS
    ├── ApagarCardsAPS + CriarCardsAPS
    ├── AtualizarDashboardAPS
    └── Application.Calculate
```

### Fluxo 3: Drag & Drop
```
AplicarAlteracoesCards (Integracao_APS)
    └── AplicarPosicaoDosCards (DragCards_APS)
        ├── LerCardsDaPlanilha
        ├── HorarioDaPosicaoX → inicio
        ├── MaquinaDaPosicaoY → máquina
        ├── AtualizarDadosPeloCard
        ├── OrdenarDadosDrag
        ├── AtribuirSequencias
        ├── RecalcularTodasAsMaquinas
        ├── AtualizarStatusDrag
        └── AtualizarCardsDepoisDoDrag
            └── CriarCardsAPS
```

### Fluxo 4: Atrasos
```
AtualizarAtrasosAPS (Atrasos_APS)
    ├── PrepararColunasAtraso
    ├── RegistrarHorariosOriginais
    ├── AplicarAtrasos
    │   └── RecalcularDuracaoOP (MotorCalculo_APS)
    ├── RecalcularComAtrasos
    └── AtualizarCardsAtraso
```

### Fluxo 5: Eventos
```
AplicarEventosAPS (Eventos_APS)
    ├── GarantirEstruturaEventosAPS
    ├── AplicarEventosNasOPsAPS
    │   ├── EventoJaAplicadoAPS (checa coluna Aplicado)
    │   ├── AdicionarEventoNaOPAPS
    │   │   └── RecalcularDuracaoOP (MotorCalculo_APS)
    │   └── Marcar como aplicado (SIM)
    ├── RecalcularSequenciamentoComEventosAPS
    │   ├── OrdenarOPsAPS
    │   └── Para cada OP: calcular novo início/fim, atualizar status
    └── AtualizarCardsDepoisEventosAPS
        └── CriarCardsAPS
```

### Fluxo 6: Refeições
```
DesenharRefeicoes (Refeição_fixa_)
    ├── CriarAbaRefeicao
    ├── AtualizarFimRefeicoes
    ├── ApagarCardsRefeicao
    └── CriarCardsRefeicao
```

---

## 4. PROPOSTA DE ESTRUTURA DE ABAS

### Abas visíveis ao usuário (4)

| # | Nome | Visível | Função |
|---|------|---------|--------|
| 1 | **INICIO** | ✅ Visível | Dashboard operacional, menu principal, KPIs resumidos |
| 2 | **OPs** | ✅ Visível | Cadastro e consulta de ordens de produção |
| 3 | **MAQUINAS** | ✅ Visível | Cadastro de máquinas/recursos |
| 4 | **PLANEJAMENTO** | ✅ Visível | Timeline/Gantt, cards, drag & drop |

### Abas técnicas ocultas (3)

| # | Nome | Visível | Função |
|---|------|---------|--------|
| 5 | **DADOS** | ❌ Oculto | Dados brutos das OPs, cálculos do motor |
| 6 | **RECURSOS** | ❌ Oculto | Cadastro técnico de máquinas (compatibilidade) |
| 7 | **EVENTOS** | ❌ Oculto | Eventos de produção |
| 8 | **REFEICAO** | ❌ Oculto | Refeições fixas |
| 9 | **RESUMO** | ❌ Oculto | Dashboard técnico (compatibilidade) |
| 10 | **CONFIG** | ❌ Oculto | Configurações técnicas |

**Observação:** As abas ocultas são necessárias porque o VBA atual escreve/lê diretamente nelas. Elas serão ocultadas via VBA (`xlSheetHidden`) e acessadas apenas por código.

### Mapeamento de responsabilidades por aba

| Funcionalidade | Aba visível | Aba técnica usada pelo VBA |
|----------------|-------------|---------------------------|
| Dashboard operacional | INICIO | RESUMO (lógica atual) |
| Cadastro de OPs | OPs | DADOS |
| Cadastro de máquinas | MAQUINAS | RECURSOS |
| Planejamento/Gantt | PLANEJAMENTO | DADOS + PLANEJAMENTO |
| Eventos | Menu → Eventos | EVENTOS |
| Refeições | Menu → Refeições | REFEICAO |
| Atrasos | Menu → Atrasos | DADOS |

---

## 5. PROPOSTA DE LAYOUT DA INTERFACE PRINCIPAL

### 5.1 ABA INICIO (Dashboard Operacional)

**Estrutura:**
- Cabeçalho fixo (linhas 1-3)
  - Logo/nome: **APS PURAN**
  - Data/hora de atualização
  - Status do planejamento (colorido)
- Área de KPIs (linhas 5-10)
  - Cards grandes com indicadores:
    - Total de OPs
    - Em produção
    - Atrasadas
    - Concluídas
    - Total de caixas
    - Horas planejadas
    - Horas de atraso
    - OPs com atraso
- Área de alertas (linhas 12-20)
  - Lista de OPs atrasadas
  - Lista de eventos pendentes
  - Lista de máquinas com manutenção
- Menu principal (linhas 22-30)
  - Botões grandes e claros:
    - 📋 Ordens de Produção
    - 🏭 Máquinas
    - 📅 Planejamento
    - ⚠ Atrasos
    - 🔧 Eventos/Manutenção
    - ⚙ Configurações
    - 🔄 Recalcular APS

**Comportamento:**
- Ao abrir o arquivo, executar `InicializarAPS`
- Dashboard atualizado automaticamente após cada operação
- Botões redirecionam para as abas correspondentes ou executam macros

### 5.2 ABA OPs (Cadastro de Ordens)

**Estrutura:**
- Título: "ORDENS DE PRODUÇÃO"
- Área de filtros/busca (linha 4)
  - Campo de busca por OP/produto
  - Filtro por máquina
  - Filtro por status
- Tabela de OPs (a partir da linha 6)
  - Colunas visíveis:
    - OP
    - Produto
    - Dosagem
    - Quantidade
    - Máquina
    - Status
    - Início
    - Fim
    - Duração Total (h)
  - Botões de ação por linha:
    - Editar
    - Excluir
    - Ver detalhes
- Botões de ação geral:
  - + Nova OP
  - 📊 Ver Planejamento
  - 🔄 Recalcular

**Comportamento:**
- Formulário de cadastro via UserForm ou área dedicada
- Busca/filtro em tempo real
- Não exibir colunas técnicas (Caixas, Capacidade_h, etc.)

### 5.3 ABA MAQUINAS (Cadastro de Recursos)

**Estrutura:**
- Título: "CADASTRO DE MÁQUINAS"
- Tabela de máquinas
  - Colunas:
    - Máquina
    - Velocidade (cx/h)
    - OEE (%)
    - Setup (h)
    - CpC (Comprimidos por caixa)
    - Ativa (Sim/Não)
    - Ações
- Botão: + Nova Máquina

**Comportamento:**
- Cadastro simplificado
- Validação de dados
- Integração direta com aba RECURSOS (técnica)

### 5.4 ABA PLANEJAMENTO (Gantt/Timeline)

**Estrutura:**
- Cabeçalho fixo com botões de ação
  - 🔄 Recalcular
  - ⚠ Aplicar Atrasos
  - 🔧 Aplicar Eventos
  - 🍽 Refeições
  - 🏠 Dashboard
- Timeline horizontal
  - Eixo X: horários (05:00 - 22:00, intervalos de 30min)
  - Eixo Y: máquinas
- Cards das OPs
  - Mostram: OP, Produto, Início, Fim, Duração, Status
  - Cores por status
  - Drag & Drop funcional
- Área de legenda/status

**Comportamento:**
- Drag & Drop real (altera dados, não apenas visual)
- Propagação automática de atrasos
- Cards de refeição sobrepostos (não bloqueiam OPs)
- Zoom temporal (dia/semana)

---

## 6. MAPA DE REUTILIZAÇÃO DOS MÓDULOS EXISTENTES

| Módulo atual | Como será reutilizado | Adaptações necessárias |
|--------------|----------------------|------------------------|
| `Globais_APS.bas` | Manter como está | Nenhuma |
| `motor_APS.bas` | Manter para compatibilidade | Pode ser chamado por `ConfigurarAPS` |
| `MotorCalculo_APS.bas` | Motor central — **NÃO ALTERAR** | Nenhuma |
| `Atrasos_APS.bas` | Manter lógica — chamar via menu | Nenhuma |
| `Eventos_APS.bas` | Manter lógica — chamar via menu | Nenhuma |
| `Refeição_fixa_.bas` | Manter lógica — chamar via menu | Nenhuma |
| `Cards_APS.bas` | Manter criação de cards | Nenhuma |
| `Timeline_APS.bas` | Manter timeline horizontal | Nenhuma |
| `Dashboard_APS.bas` | Manter dashboard técnico | Pode ser chamado por dashboard operacional |
| `DragCards_APS.bas` | Manter drag & drop | Nenhuma |
| `Integracao_APS.bas` | Manter funções públicas | Adicionar novas funções de navegação |
| `Final_APS.bas` | Manter `IniciarSistemaAPS` | Nenhuma |

**Regra fundamental:** Nenhum módulo existente será alterado. A nova interface apenas chamará as funções públicas já existentes.

---

## 7. FUNÇÕES PÚBLICAS EXISTENTES (PONTO DE INTEGRAÇÃO)

### Globais_APS.bas
- Constantes: `ABA_DADOS`, `ABA_PLANEJAMENTO`, `PREFIXO_CARD`, `STATUS_*`, `ALT_IDX_*`

### motor_APS.bas
- `ConfigurarAPS()` — configuração inicial

### MotorCalculo_APS.bas
- `ExecutarMotorAPS()` — executa motor de cálculo
- `RecalcularDuracaoOP(ws, linha)` — recalcula duração de uma OP
- `RecalcularTodasDuracoesAPS()` — recalcula todas as OPs

### Atrasos_APS.bas
- `AtualizarAtrasosAPS()` — aplica atrasos
- `AtualizarCardsAtraso()` — atualiza cards após atraso
- `AtualizarCardIndividual(...)` — atualiza card individual

### Eventos_APS.bas
- `AplicarEventosAPS()` — aplica eventos
- `GarantirEstruturaEventosAPS()` — garante estrutura da aba EVENTOS

### Refeição_fixa_.bas
- `DesenharRefeicoes()` — desenha refeições
- `CriarAbaRefeicao()` — cria aba de refeições
- `AtualizarFimRefeicoes()` — atualiza horários
- `ApagarCardsRefeicao()` — remove cards de refeição

### Cards_APS.bas
- `CriarCardsAPS()` — cria todos os cards
- `ApagarCardsAPS()` — remove todos os cards
- `AtualizarCardsAPS()` — atualiza cards existentes
- `AtualizarCardOP(...)` — atualiza card de OP específica

### Timeline_APS.bas
- `ConstruirTimelineAPS()` — constrói timeline horizontal

### Dashboard_APS.bas
- `CriarDashboardAPS()` — cria dashboard
- `AtualizarDashboardAPS()` — atualiza dashboard

### DragCards_APS.bas
- `AplicarPosicaoDosCards()` — aplica posição dos cards após drag
- `AplicarAlteracoesCards()` — integra drag + dashboard

### Integracao_APS.bas
- `AtualizarAPS()` — atualização completa
- `AplicarAlteracoesCards()` — aplica alterações de cards
- `CriarBotoesAPS()` — cria botões
- `AtualizacaoRapidaAPS()` — atualização rápida
- `ReconstruirAPS()` — reconstrói sistema
- `InicializarAPS()` — inicialização ao abrir

### Final_APS.bas
- `IniciarSistemaAPS()` — inicialização completa
- `ValidarSistemaAPS()` — valida estrutura
- `PrepararAbas()` — cria abas
- `PrepararDados()` — prepara dados
- `PrepararPlanejamento()` — prepara planejamento
- `PrepararResumo()` — prepara resumo

---

## 8. PROBLEMAS IDENTIFICADOS E SOLUÇÕES

### Problema 1: Muitas abas visíveis
**Solução:** Ocultar abas técnicas (DADOS, RECURSOS, EVENTOS, REFEICAO, RESUMO, CONFIG) usando `xlSheetHidden` ou `xlSheetVeryHidden`.

### Problema 2: Botões espalhados
**Solução:** Centralizar botões em menu principal na aba INICIO e em barra de ferramentas na aba PLANEJAMENTO.

### Problema 3: Usuário não entende onde cadastrar
**Solução:** Criar abas dedicadas com interfaces simplificadas (OPs, MAQUINAS) que espelham para as abas técnicas.

### Problema 4: Cards não aparecem
**Solução:** Garantir que `CriarCardsAPS` seja chamado automaticamente após `ExecutarMotorAPS` e `ConstruirTimelineAPS`.

### Problema 5: Propagação de atrasos não é automática
**Solução:** Reutilizar `RecalcularSequenciamentoComEventosAPS` e `RecalcularComAtrasos` que já existem, apenas garantir que sejam chamados na ordem correta.

### Problema 6: Drag & Drop apenas visual
**Solução:** `AplicarPosicaoDosCards` já converte posição em dados. Garantir que `AplicarAlteracoesCards` seja chamado após o drag.

---

## 9. ORDEM DE IMPLEMENTAÇÃO RECOMENDADA

### Fase 2.1: Nova estrutura de abas
1. Criar aba INICIO
2. Ocultar abas técnicas
3. Migrar dashboard para INICIO
4. Criar menu principal

### Fase 2.2: Interface de cadastros
5. Criar aba OPs com interface simplificada
6. Criar aba MAQUINAS com interface simplificada
7. Conectar cadastros às abas técnicas via VBA

### Fase 2.3: Planejamento profissional
8. Aprimorar PLANEJAMENTO com Gantt
9. Garantir cards sempre visíveis
10. Implementar drag & drop com feedback visual

### Fase 2.4: Integração e testes
11. Conectar todos os botões
12. Testar fluxos completos
13. Gerar documentação

---

## 10. CÓDIGO QUE NÃO DEVE SER IMPORTADO COMO `.bas`

**Nenhum.** Todos os 12 arquivos são módulos VBA padrão.

**Observação:** Se o usuário quiser execução automática ao abrir, ele deve adicionar manualmente no objeto `ThisWorkbook`:

```vba
Private Sub Workbook_Open()
    InicializarAPS
End Sub
```

---

## 11. DECISÕES DE DESIGN

1. **Não criar novas abas por máquina** — Manter aba única MAQUINAS
2. **Não alterar motor de cálculo** — Reutilizar `MotorCalculo_APS.bas` intacto
3. **Não alterar regras de negócio** — Todos os módulos existentes são chamados como estão
4. **Interface separada de dados** — Abas visíveis são apenas interface; abas técnicas são backend
5. **Navegação por botões** — Não usar ribbons ou menus complexos
6. **Drag & Drop real** — Já existe, apenas garantir chamada correta
7. **Propagação automática** — Já existe, apenas garantir ordem de execução

---

## PRÓXIMOS PASSOS

1. Implementar nova estrutura de abas no `APS_PURAN_Base.xlsx`
2. Criar interface INICIO profissional
3. Criar interfaces OPs e MAQUINAS
4. Aprimorar PLANEJAMENTO
5. Conectar todos os botões
6. Testar fluxos
7. Gerar `RELATORIO_UX_FASE2.md`
8. Commit e push
