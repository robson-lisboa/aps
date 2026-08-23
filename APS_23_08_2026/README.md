# APS PURAN — Sistema de Planejamento e Controle da Produção

**Versão:** 1.0  
**Data:** 2026-08-23  
**Tipo:** Sistema APS profissional para Excel Desktop

---

## 📋 Descrição

O APS PURAN é um sistema de Planejamento e Escalonamento Avançado de Produção (APS) desenvolvido para indústrias farmacêuticas. Ele transforma o Excel Desktop em uma plataforma de planejamento visual, com timeline horizontal, cards flutuantes, drag & drop, propagação automática de atrasos e dashboard operacional.

**Não é uma planilha comum. É um sistema APS dentro do Excel.**

---

## 🎯 Funcionalidades

- ✅ **Interface profissional** — 4 telas principais: INÍCIO, OPERAÇÕES, MÁQUINAS, PLANEJAMENTO
- ✅ **Timeline horizontal** — Gantt visual com máquinas no eixo vertical e horários no horizontal
- ✅ **Cards flutuantes** — Cada OP é um card visual com início, fim, duração e status
- ✅ **Drag & Drop real** — Arrastar cards altera o planejamento real
- ✅ **Propagação de atrasos** — Atrasos se propagam automaticamente para OPs seguintes
- ✅ **Eventos/manutenção** — Bloqueiam períodos na timeline
- ✅ **Refeições** — Tratadas como restrições no cálculo
- ✅ **Cadastro de OPs** — Interface intuitiva sem expor abas técnicas
- ✅ **Cadastro de máquinas** — Nova máquina aparece automaticamente na timeline
- ✅ **Pesquisa global** — Localizar OP rapidamente
- ✅ **Dashboard operacional** — KPIs em tempo real
- ✅ **Testes automáticos** — Validação completa do sistema

---

## 📁 Estrutura do Projeto

```
APS_PURAN/
├── ENTREGAS/
│   ├── APS_PURAN_FINAL.xlsx       # Planilha final com interface
│   ├── APS_PURAN_Fase2.xlsx       # Versão anterior
│   └── README.md                   # Instruções de entrega
├── VBA_IMPORTAR/
│   ├── Globais_APS.bas
│   ├── MotorCalculo_APS.bas
│   ├── Eventos_APS.bas
│   ├── Atrasos_APS.bas
│   ├── Cards_APS.bas
│   ├── DragCards_APS.bas
│   ├── Dashboard_APS.bas
│   ├── Timeline_APS.bas
│   ├── Refeição_fixa_.bas
│   ├── Integracao_APS.bas
│   ├── Final_APS.bas
│   ├── Navegacao_APS.bas
│   ├── Interface_APS.bas
│   └── Teste_Fase3.bas
└── DOCUMENTACAO/
    ├── MANUAL_USUARIO.md
    ├── MANUAL_INSTALACAO.md
    └── RELATORIO_FINAL.md
```

---

## 🚀 Instalação

### 1. Baixar arquivos
```bash
git clone https://github.com/robson-lisboa/aps.git
```

### 2. Abrir planilha
1. Abra `ENTREGAS/APS_PURAN_FINAL.xlsx` no Excel Desktop
2. Habilite as macros

### 3. Importar VBA
1. Pressione `Alt+F11`
2. Importe os 15 módulos de `VBA_IMPORTAR/` na ordem do manual
3. Compile (`Depurar → Compilar VBAProject`)
4. Salve como `.xlsm`

### 4. Executar
1. Abra o `.xlsm`
2. Habilite macros
3. Janela Imediata (`Ctrl+G`)
4. Execute: `InicializarAPS`

---

## 🧪 Testes

Execute na janela Imediata (`Ctrl+G`):
```vba
ExecutarTesteFase3
```

O teste cria dados de exemplo, executa o motor, verifica cards, timeline, atrasos, eventos, cadastros e gera um log completo.

---

## 📊 Telas do Sistema

### INÍCIO
- Dashboard operacional
- KPIs: Total OPs, Em Produção, Atrasadas, Concluídas, Horas Planejadas, Horas Atraso
- Menu principal de navegação
- Busca global de OP

### OPERAÇÕES
- Cadastro de OPs
- Listagem com filtros
- Busca por OP, produto, máquina, status
- Ações: Nova OP, Atualizar, Voltar

### MÁQUINAS
- Cadastro de máquinas/recursos
- Listagem com busca
- Campos: código, nome, velocidade, OEE, setup, capacidade, status
- Nova máquina aparece automaticamente na timeline

### PLANEJAMENTO
- Timeline horizontal (05:00-22:00)
- Cards das OPs posicionados por horário e máquina
- Drag & Drop funcional
- Botões: Recalcular, Atrasos, Eventos, Refeições, Dashboard

---

## 🔧 Tecnologias

- **Excel Desktop** — Runtime
- **VBA** — Camada de aplicação
- **Shapes** — Cards flutuantes
- **Timeline horizontal** — Eixo X = horário, Eixo Y = máquina

---

## 📦 Entregas

- `ENTREGAS/APS_PURAN_FINAL.xlsx` — Planilha final pronta para uso
- `VBA_IMPORTAR/` — Módulos VBA para importação
- `DOCUMENTACAO/` — Manuais e relatórios

---

## 📄 Documentação

- [Manual do Usuário](DOCUMENTACAO/MANUAL_USUARIO.md)
- [Manual de Instalação](DOCUMENTACAO/MANUAL_INSTALACAO.md)
- [Relatório Final](DOCUMENTACAO/RELATORIO_FINAL.md)

---

## ⚠️ Requisitos

- Microsoft Excel Desktop 2016 ou superior
- Macros habilitadas
- Windows ou Mac

**Não funciona no Excel Online.**

---

## 📥 Download

**Planilha final:**  
https://github.com/robson-lisboa/aps/blob/main/ENTREGAS/APS_PURAN_FINAL.xlsx

**Módulos VBA:**  
https://github.com/robson-lisboa/aps/tree/main/VBA_IMPORTAR

**Repositório completo:**  
https://github.com/robson-lisboa/aps

---

## 📝 Licença

Projeto desenvolvido para uso interno — APS PURAN.

---

**APS PURAN** — Sistema de Planejamento e Controle da Produção  
**Desenvolvido para indústria farmacêutica**
