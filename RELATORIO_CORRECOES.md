# APS — Relatório de correções

Analisei os 12 arquivos como um projeto VBA único (é assim que o Excel vai
compilá-los, já que módulos VBA compartilham o mesmo namespace). Achei 2 erros
que **impediam o projeto inteiro de compilar** e 2 bugs silenciosos que
faziam partes inteiras do sistema simplesmente não funcionarem, sem gerar
nenhum erro visível. Corrigi tudo. Detalhes abaixo.

---

## 🔴 1. Erro de compilação: nome duplicado `ApagarCardsAPS`

`Cards_APS.txt` e `Cards_Flutuantes.bas` tinham cada um o seu próprio:

```vb
Public Sub ApagarCardsAPS()
```

Em VBA isso **não compila em nenhuma planilha** — dá o erro
`Ambiguous name detected: ApagarCardsAPS` assim que você tenta rodar
qualquer macro do projeto.

Investigando, `Cards_Flutuantes.bas` era uma segunda implementação (mais
antiga) do sistema de cards, que nenhum outro módulo chama — quem realmente
é usado é o conjunto `Cards_APS.txt` + `DragCards_APS.txt` +
`Atrasos_APS.txt` (mais completo, já trata atraso e arraste).

**Problema:** `Cards_Flutuantes.bas` também era o único lugar do projeto que
declarava, como constantes *públicas* (globais):

```vb
Public Const ABA_DADOS As String = "DADOS"
Public Const ABA_PLANEJAMENTO As String = "PLANEJAMENTO"
```

E os módulos `Eventos_APS.bas` e `Refeição_fixa_.bas` usam essas duas
constantes sem declará-las localmente — ou seja, dependiam desse arquivo
para compilar, mesmo sem saber.

**Correção:** removi `Cards_Flutuantes.bas` e criei **`Globais_APS.bas`**,
só com essas duas constantes. Resolve a ambiguidade e mantém tudo que
realmente era usado.

➡️ **No Excel: apague o módulo `Cards_Flutuantes` do VBA Project e importe
o novo `Globais_APS.bas` no lugar.**

---

## 🔴 2. Erro de compilação: `ExecutarMotorAPS` não existe

Em `Eventos_APS.bas`, a macro `AplicarEventosAPS` começava chamando:

```vb
ExecutarMotorAPS
```

Essa Sub **não existe em nenhum dos 12 arquivos**. Com `Option Explicit`
ativo (está em todos os módulos, o que é ótimo), isso também impede a
compilação do projeto — erro `Sub or Function not defined`.

**Correção:** removi a chamada. Ela era desnecessária: logo depois já é
chamada `RecalcularSequenciamentoComEventos`, que sozinha já recalcula
Início/Fim/Status de todas as OPs a partir da "Duração Total (h)" — ou
seja, o próprio módulo já fazia o recálculo completo, só sobrava uma
chamada "fantasma" no início.

---

## 🟠 3. Bug silencioso: nome de coluna errado (`"Duração Total_h"`)

A coluna real criada em `DADOS` (definida em `motor_APS.txt` e
`Final_APS.txt`) se chama:

```
"Duração Total (h)"
```

Só que **3 módulos diferentes** procuravam por um nome ligeiramente
diferente:

```
"Duração Total_h"
```

Como a busca de coluna (`EncontrarColuna...`) não encontra essa string,
ela retorna `0`, e todo esse trecho tem uma checagem do tipo
`If cDuracao = 0 Then Exit Sub` — ou seja, **a macro roda, não dá erro
nenhum, mostra até a mensagem de sucesso, mas não faz nada**. É o tipo de
bug mais difícil de perceber, porque parece que funcionou.

Isso quebrava, silenciosamente:

| Arquivo | O que ficava sem efeito |
|---|---|
| `Atrasos_APS.txt` | Recalcular os horários das OPs depois de um atraso |
| `DragCards_APS.txt` | Recalcular a programação depois de arrastar um card |
| `Eventos_APS.bas` | Aplicar a duração de eventos (manutenção, setup extra, etc.) |

**Correção:** troquei `"Duração Total_h"` → `"Duração Total (h)"` nos 3
arquivos.

---

## 🟠 4. Bug de lógica: atraso nunca alterava a duração (e acumulava)

Mesmo com o nome de coluna corrigido, achei mais um problema em
`Atrasos_APS.txt`: a função `AplicarAtrasos` somava o valor de `Atraso_h`
apenas na coluna `Eventos_h` (histórico), mas **nunca somava esse tempo na
"Duração Total (h)"** — que é a coluna que o resto do sistema usa para
recalcular o novo horário das OPs. Ou seja: mesmo corrigindo o nome da
coluna, o atraso continuaria sem "empurrar" a programação.

Além disso, a forma como estava escrito (`Eventos_h = Eventos_h + atraso`)
tinha um segundo problema: cada vez que você rodasse "Atualizar Atrasos"
de novo (por exemplo, só para atualizar outro card), o mesmo atraso seria
somado de novo, inflando a duração a cada clique.

**Correção:** segui o mesmo padrão que o módulo já usa para "Início
Original" / "Fim Original" — criei uma coluna nova, **"Duração Base_h"**,
que guarda a duração real da OP (sem atraso) na primeira vez. A partir
daí, toda vez que o módulo roda:

```
Duração Total (h) = Duração Base_h + Atraso_h
```

Isso é recalculado do zero a cada execução (não somado em cima do valor
anterior), então rodar o módulo várias vezes é seguro — e agora o atraso
realmente empurra a programação, como a documentação do próprio módulo
descreve no exemplo do cabeçalho (`ORIGINAL 05:30→10:00 / ATRASO +02:00 /
NOVO 07:30→12:00`).

---

## 🟡 5. Melhoria: módulos sem nenhum botão

`Atrasos_APS.txt`, `Eventos_APS.bas` e `Refeição_fixa_.bas` são macros
completas e funcionais, mas **nenhum dos 5 botões** criados por
`CriarBotoesAPS` (em `Integracao_APS.txt`) aciona qualquer uma delas — só
dava para rodar manualmente pelo VBA (Alt+F8).

**Melhoria aplicada:** adicionei 3 novos botões — **ATRASOS**, **EVENTOS**
e **REFEIÇÃO** — no mesmo padrão visual dos outros 5, chamando
respectivamente `AtualizarAtrasosAPS`, `AplicarEventosAPS` e
`DesenharRefeicoes`.

---

## 🟢 6. Novo: motor de cálculo (`MotorCalculo_APS.bas`)

Criei o módulo que faltava — `MotorCalculo_APS.bas` — com a Sub
`ExecutarMotorAPS` (esse nome não foi escolha minha: era exatamente o que
`Eventos_APS.bas` já tentava chamar antes, no bug nº 2; então era esse o
nome que o próprio sistema já esperava).

Para cada OP com `Máquina` preenchida, ele busca os dados da máquina na
aba `RECURSOS` e calcula:

```
Caixas          = Quantidade (já está em caixas, confirmado com você;
                   se a coluna "Unidade" indicar comprimidos, ele
                   converte usando "Comprimidos por caixa")
Capacidade_h    = Velocidade (cx/h, cadastrada em RECURSOS)
Produção_h      = Caixas / (Velocidade × OEE)
Setup_h         = Setup padrão (cadastrado em RECURSOS)
Duração Base_h  = Produção_h + Setup_h
Duração Total(h)= Duração Base_h + Atraso_h + Eventos_h
```

Detalhes:

- Se **OEE** estiver cadastrado como `85` em vez de `0,85`, o motor
  corrige sozinho (divide por 100).
- Se uma OP tiver uma **Máquina que não existe em RECURSOS**, ou
  **Velocidade/OEE zerados**, o motor não trava — ele pula essa OP e
  te avisa no final, com a lista de quais OPs precisam de atenção.
- É seguro rodar quantas vezes quiser: os valores são sempre
  recalculados do zero (não somados em cima do anterior).

**Onde ele roda:** adicionei um botão **MOTOR**, e ele também passou a
rodar automaticamente dentro do fluxo principal (`AtualizarAPS`, botão
"ATUALIZAR APS") e do fluxo de inicialização completa (`IniciarSistemaAPS`
em `Final_APS.txt`), sempre logo antes da Timeline.

---

## 🟢 7. Novo: coluna "Aplicado" evita duplicidade em Eventos

Resolvido também o segundo ponto: `AplicarEventosNasOPs` agora verifica
uma nova coluna **"Aplicado"** na aba `EVENTOS` antes de somar o tempo de
um evento numa OP. Um evento só é somado a `Eventos_h` **uma única vez**
— clicar no botão "EVENTOS" de novo não soma tudo outra vez.

Se você editar a duração de um evento já aplicado e quiser que ele seja
recalculado, basta apagar o "SIM" da coluna "Aplicado" naquela linha e
rodar o botão de novo.

Também aproveitei para fazer a aba `EVENTOS` **se criar sozinha** (com os
cabeçalhos `OP`, `Tipo`, `Duração_h`, `Ativo`, `Aplicado`) na primeira vez
que você clicar em "EVENTOS" — antes, se essa aba não existisse, o botão
simplesmente dava erro.

---

## 🟢 8. Unificação: "Duração Total (h)" nunca mais é somada diretamente

Como consequência dos pontos 6 e 7, também corrigi um problema que
descobri ao montar o motor: `AplicarAtrasos` (Atrasos_APS.txt) e
`AdicionarEventoNaOP` (Eventos_APS.bas) escreviam, cada um do seu jeito,
direto em `Duração Total (h)` **e também disputavam a coluna `Eventos_h`**
entre si (o módulo de atrasos usava `Eventos_h` para guardar o atraso, o
que não faz sentido — `Eventos_h` é do módulo de eventos).

Agora existe uma única Sub responsável por isso —
`RecalcularDuracaoOP` (em `MotorCalculo_APS.bas`) — que sempre recalcula:

```
Duração Total (h) = Duração Base_h + Atraso_h + Eventos_h
```

Tanto `Atrasos_APS.txt` quanto `Eventos_APS.bas` chamam essa mesma Sub
depois de atualizar sua própria coluna (`Atraso_h` ou `Eventos_h`,
respectivamente). Isso significa que atraso e eventos agora **se somam
corretamente** na duração total, e a ordem em que você roda os botões
(MOTOR, ATRASOS, EVENTOS) não importa mais para o resultado final.

---

## 🔵 Ponto que ainda ficou como está (baixo risco, sem urgência)

`Refeição_fixa_.bas` declara `ABA_PLANEJAMENTO_REFEICAO` mas nunca usa
essa constante (usa `ABA_PLANEJAMENTO`, que vem do `Globais_APS.bas`) —
é código morto, inofensivo. Não mexi porque não afeta o funcionamento.

---

## ✅ Como aplicar no Excel

1. Abra o VBA Editor (Alt+F11).
2. **Remova o módulo `Cards_Flutuantes`** (clique direito → Remover).
3. Importe (ou substitua) todos os arquivos corrigidos, incluindo os dois
   módulos novos: `Globais_APS.bas` e `MotorCalculo_APS.bas`.
4. Menu Debug → **Compile VBA Project**. Deve compilar sem erro agora.
5. Rode `CriarBotoesAPS` (ou `ReconstruirAPS`) para recriar os botões —
   agora são 9: ATUALIZAR APS, APLICAR CARDS, TIMELINE, DASHBOARD,
   LIMPAR CARDS, ATRASOS, EVENTOS, REFEIÇÃO e **MOTOR**.
6. Cadastre (ou confira) a aba `RECURSOS` — Velocidade (cx/h), OEE,
   Setup padrão e Comprimidos por caixa de cada máquina — e clique em
   **MOTOR** para calcular a duração de todas as OPs pela primeira vez.

