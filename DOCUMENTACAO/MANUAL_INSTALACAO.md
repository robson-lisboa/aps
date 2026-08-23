# MANUAL DE INSTALAÇÃO — APS PURAN

**Versão:** 1.0  
**Data:** 2026-08-23  
**Requisitos:** Excel Desktop (Windows/Mac), macros habilitadas

---

## 1. REQUISITOS DO SISTEMA

- Microsoft Excel Desktop 2016 ou superior
- Macros habilitadas
- Espaço em disco: ~10 MB
- Memória RAM: 4 GB mínimo

---

## 2. ARQUIVOS DO PROJETO

```
APS_PURAN/
├── ENTREGAS/
│   ├── APS_PURAN_FINAL.xlsx
│   └── README.md
├── VBA_IMPORTAR/
│   ├── Globais_APS.bas
│   ├── motor_APS.bas
│   ├── MotorCalculo_APS.bas
│   ├── Atrasos_APS.bas
│   ├── Eventos_APS.bas
│   ├── Refeição_fixa_.bas
│   ├── Cards_APS.bas
│   ├── Timeline_APS.bas
│   ├── Dashboard_APS.bas
│   ├── DragCards_APS.bas
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

## 3. PROCEDIMENTO DE INSTALAÇÃO

### 3.1 Baixar arquivos
1. Baixe o repositório: `git clone https://github.com/robson-lisboa/aps.git`
2. Ou baixe diretamente:
   - `ENTREGAS/APS_PURAN_FINAL.xlsx`
   - Todos os arquivos de `VBA_IMPORTAR/`

### 3.2 Abrir planilha
1. Abra `ENTREGAS/APS_PURAN_FINAL.xlsx` no Excel Desktop
2. Habilite as macros quando solicitado
3. Se o Excel bloquear as macros:
   - Arquivo → Opções → Central de Confiança
   - Configurações da Central de Confiança
   - Habilitar macros

### 3.3 Importar módulos VBA
1. Pressione `Alt+F11` para abrir o VBA Editor
2. No menu **Arquivo** → **Importar arquivo**
3. Importe os módulos **na ordem abaixo**:

| Ordem | Arquivo |
|-------|---------|
| 1 | `Globais_APS.bas` |
| 2 | `motor_APS.bas` |
| 3 | `MotorCalculo_APS.bas` |
| 4 | `Atrasos_APS.bas` |
| 5 | `Eventos_APS.bas` |
| 6 | `Refeição_fixa_.bas` |
| 7 | `Cards_APS.bas` |
| 8 | `Timeline_APS.bas` |
| 9 | `Dashboard_APS.bas` |
| 10 | `DragCards_APS.bas` |
| 11 | `Integracao_APS.bas` |
| 12 | `Final_APS.bas` |
| 13 | `Navegacao_APS.bas` |
| 14 | `Interface_APS.bas` |
| 15 | `Teste_Fase3.bas` |

### 3.4 Compilar projeto
1. No VBA Editor, vá em **Depurar** → **Compilar VBAProject**
2. Se houver erros, verifique se todos os 15 módulos foram importados
3. Se compilar sem erros, prossiga

### 3.5 Salvar como .xlsm
1. Feche o VBA Editor
2. No Excel, clique em **Arquivo** → **Salvar Como**
3. Em **Tipo**, selecione **Pasta de trabalho habilitada para macro do Excel (*.xlsm)**
4. Salve como `APS_PURAN.xlsm`
5. Feche o arquivo

---

## 4. EXECUÇÃO DO SISTEMA

### 4.1 Primeira execução
1. Abra `APS_PURAN.xlsm`
2. Habilite as macros
3. Pressione `Alt+F11` para abrir o VBA Editor
4. Pressione `Ctrl+G` para abrir a janela **Imediata**
5. Digite:
   ```
   InicializarAPS
   ```
6. Pressione `Enter`

### 4.2 Teste automático
1. Abra a janela Imediata (`Ctrl+G`)
2. Digite:
   ```
   ExecutarTesteFase3
   ```
3. Pressione `Enter`
4. Aguarde a conclusão
5. Verifique o log gerado na pasta do arquivo

---

## 5. ESTRUTURA DE ABAS

### Abas visíveis
- **INICIO** — Dashboard operacional
- **OPs** — Cadastro de ordens de produção
- **MAQUINAS** — Cadastro de máquinas
- **PLANEJAMENTO** — Timeline e cards

### Abas técnicas ocultas
- DADOS
- RECURSOS
- EVENTOS
- REFEICAO
- RESUMO
- CONFIG

---

## 6. SOLUÇÃO DE PROBLEMAS

### Problema: Macros desabilitadas
**Solução:** 
- Arquivo → Opções → Central de Confiança → Habilitar macros

### Problema: Erro de compilação
**Solução:**
- Verifique se todos os 15 módulos foram importados
- Verifique se não há referências quebradas em Ferramentas → Referências

### Problema: Botões não aparecem
**Solução:**
- Execute `ConfigurarBotoesInicio` na janela Imediata

### Problema: Abas não existem
**Solução:**
- Execute `InicializarAPS` na janela Imediata

### Problema: Cards não aparecem
**Solução:**
- Execute `AtualizarAPS` na janela Imediata
- Verifique se há OPs cadastradas em DADOS
- Verifique se há máquinas cadastradas em RECURSOS

---

## 7. CONTATO

Em caso de dúvidas técnicas, consulte a documentação do projeto no GitHub.

---

**APS PURAN** — Sistema de Planejamento e Controle da Produção  
**Desenvolvido para indústria farmacêutica**
