# MANUAL DO USUÁRIO — APS PURAN

**Versão:** 1.0  
**Data:** 2026-08-23  
**Sistema:** APS PURAN — Advanced Production Planning & Scheduling

---

## 1. BEM-VINDO AO APS PURAN

O APS PURAN é um sistema de planejamento e escalonamento de produção desenvolvido para indústrias farmacêuticas. Ele permite:

- Cadastrar ordens de produção (OPs)
- Cadastrar máquinas e recursos
- Planejar produção em timeline horizontal
- Visualizar OPs como cards flutuantes
- Arrastar cards para reprogramar (Drag & Drop)
- Aplicar atrasos com propagação automática
- Gerenciar eventos e manutenções
- Controlar refeições
- Acompanhar indicadores em tempo real

---

## 2. PRINCIPAIS TELAS

### 2.1 INÍCIO
- Dashboard com indicadores principais
- Acesso rápido às funcionalidades
- Status do sistema

### 2.2 OPERAÇÕES
- Cadastro de OPs
- Listagem de OPs
- Busca e filtros
- Ações: Nova OP, Atualizar, Voltar

### 2.3 MÁQUINAS
- Cadastro de máquinas
- Listagem de recursos
- Busca de máquinas
- Ações: Nova Máquina

### 2.4 PLANEJAMENTO
- Timeline horizontal
- Cards das OPs
- Drag & Drop
- Botões: Recalcular, Atrasos, Eventos, Refeições, Dashboard

---

## 3. FLUXO BÁSICO

```
ABRIR SISTEMA
    ↓
INÍCIO
    ↓
NOVA OP / NOVA MÁQUINA
    ↓
ATUALIZAR APS
    ↓
PLANEJAMENTO
    ↓
VER CARDS / ARRASTAR / APLICAR ATRASO
    ↓
DASHBOARD ATUALIZA
```

---

## 4. CADASTRO DE OP

1. Acesse **OPERAÇÕES**
2. Clique em **+ NOVA OP**
3. Preencha os campos:
   - Nº OP
   - Produto
   - Quantidade
   - Unidade
   - Máquina
   - Prioridade
   - Data desejada
4. Clique em **SALVAR**
5. Clique em **ATUALIZAR APS** para recalcular

---

## 5. CADASTRO DE MÁQUINA

1. Acesse **MÁQUINAS**
2. Clique em **+ NOVA MÁQUINA**
3. Preencha os campos:
   - Código
   - Nome
   - Velocidade
   - OEE
   - Setup padrão
   - Capacidade
   - Status
4. Clique em **SALVAR**
5. A máquina aparecerá automaticamente na timeline

---

## 6. PLANEJAMENTO

### 6.1 Timeline
- Horários no eixo horizontal (05:00 às 22:00)
- Máquinas no eixo vertical
- Cards representam OPs

### 6.2 Cards
Cada card mostra:
- Número da OP
- Produto
- Máquina
- Horário de início e fim
- Duração
- Status

### 6.3 Drag & Drop
1. Clique e arraste um card
2. Solte na nova posição
3. O sistema recalcula automaticamente:
   - Novo horário
   - Nova sequência
   - Conflitos
   - Cards posteriores

---

## 7. ATRASOS

1. Acesse **PLANEJAMENTO**
2. Clique em **⚠ ATRASOS**
3. Informe o atraso na OP
4. Clique em **APLICAR**
5. O sistema:
   - Recalcula a duração
   - Propaga o atraso para OPs seguintes
   - Atualiza os cards
   - Atualiza o dashboard

---

## 8. EVENTOS / MANUTENÇÃO

1. Acesse **PLANEJAMENTO**
2. Clique em **🔧 EVENTOS**
3. Selecione o tipo:
   - Manutenção
   - Parada
   - Limpeza
   - Setup adicional
   - Indisponibilidade
4. Informe máquina, horário e duração
5. Clique em **APLICAR**
6. O sistema bloqueia o período e reorganiza as OPs

---

## 9. REFEIÇÕES

- Configuradas automaticamente
- Bloqueiam períodos na timeline
- O cálculo considera os intervalos
- Não é necessário configurar manualmente

---

## 10. DASHBOARD

Indicadores disponíveis:
- Total de OPs
- OPs em produção
- OPs atrasadas
- OPs concluídas
- Total de caixas
- Horas planejadas
- Horas de atraso
- OPs com atraso
- Utilização por máquina

---

## 11. PESQUISA

1. Acesse **INÍCIO**
2. Clique em **🔎 PESQUISAR OP**
3. Digite o número da OP
4. O sistema localiza e destaca o card na timeline

---

## 12. ATALHOS DE NAVEGAÇÃO

- **INÍCIO** — Dashboard principal
- **OPERAÇÕES** — Cadastro de OPs
- **MÁQUINAS** — Cadastro de máquinas
- **PLANEJAMENTO** — Timeline e cards
- **EVENTOS** — Manutenções e paradas
- **ATUALIZAR APS** — Recalcula todo o sistema

---

## 13. DICAS

- Sempre clique em **ATUALIZAR APS** após alterações
- Use **Drag & Drop** para reprogramar rapidamente
- Verifique o **DASHBOARD** para identificar gargalos
- Aproprie-se dos **CARDS** na timeline para visão rápida
- Use **PESQUISA** para localizar OPs rapidamente

---

## 14. SUPORTE

Em caso de dúvidas ou problemas:
1. Verifique se todas as abas técnicas estão ocultas
2. Verifique se o VBA foi importado corretamente
3. Verifique se as macros estão habilitadas
4. Consulte a documentação técnica em `DOCUMENTACAO/`

---

**APS PURAN** — Sistema de Planejamento e Controle da Produção  
**Desenvolvido para indústria farmacêutica**
