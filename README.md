# meuAddin — Macros utilitárias para Excel

Suplemento (`.xlam`) com as macros que eu mais uso no dia a dia: layouts padronizados, bordas e fundos com a identidade visual verde, transposição de tabelas para formato banco de dados, snapshot de dados, cópia de seleção como imagem, alternância de abas de apoio e um **Sistema de Alertas** que audita erros de fórmula e comentários de qualquer arquivo.

O código fica versionado aqui como arquivos `.bas` (texto puro), e qualquer pessoa consegue montar o `.xlam` no próprio Excel em poucos minutos seguindo este guia.

---

## Sumário

1. [Conteúdo do repositório](#conteúdo-do-repositório)
2. [Requisitos](#requisitos)
3. [Instalação passo a passo](#instalação-passo-a-passo)
4. [Guia "Macros Lobão" na Faixa de Opções (opcional)](#guia-macros-lobão-na-faixa-de-opções-opcional)
5. [Como executar as macros sem a Faixa de Opções](#como-executar-as-macros-sem-a-faixa-de-opções)
6. [Referência das macros](#referência-das-macros)
7. [Sistema de Alertas](#sistema-de-alertas)
8. [Personalização](#personalização)
9. [Atualizando o repositório depois de mexer no código](#atualizando-o-repositório-depois-de-mexer-no-código)
10. [Solução de problemas](#solução-de-problemas)
11. [Licença](#licença)

---

## Conteúdo do repositório

```
meuAddin/
├── README.md                  ← este guia
├── CHANGELOG.md               ← histórico de versões
├── LICENSE                    ← licença MIT
├── .gitattributes             ← impede o Git de alterar encoding/quebras de linha dos .bas
├── .gitignore
├── src/
│   ├── modUtilitarios.bas     ← macros de layout, formatação, dados e abas
│   ├── modInstalador.bas      ← cria o Sistema de Alertas no arquivo ativo
│   ├── modAlertasModelo.bas   ← código-modelo copiado para o arquivo do usuário
│   └── modRibbon.bas          ← callbacks dos botões da Faixa de Opções
└── ribbon/
    └── customUI14.xml         ← definição da guia "Macros Lobão" (opcional)
```

| Módulo | Papel |
|---|---|
| `modUtilitarios` | Todas as macros de uso direto (Layout1, Layout2, FundoVerde, LinhasVerdes, BarraLateral, FormatoBD, Copia_Imagem, SnapshotDados, ExibirOcultarAbasComApoio). |
| `modInstalador` | Macro `CriarSistemaAlertas`, que monta a aba ALERTAS e injeta o código de auditoria no arquivo ativo. |
| `modAlertasModelo` | **Não é executado no add-in.** É o "molde" que o instalador copia para dentro do arquivo do usuário. O nome precisa ser exatamente esse. |
| `modRibbon` | Liga os botões da guia personalizada às macros. Só tem efeito se o XML da Faixa de Opções for embutido. |

---

## Requisitos

- **Excel para Windows**, versão 2010 ou superior (testado no Microsoft 365).
- **Comentários modernos (com respostas)** no Sistema de Alertas exigem Excel 365 ou 2021+. Em versões mais antigas, apenas as anotações (notas amarelas) são listadas — sem erro.
- **Sistema de Alertas** exige a opção *"Confiar no acesso ao modelo de objeto do projeto do VBA"* (explicado abaixo) e usa `Scripting.Dictionary`, que **não existe no Excel para Mac**. As demais macros funcionam no Mac, exceto onde houver diferença de API.
- Fonte **Simplon BP Regular** (opcional). Se ela não estiver instalada, o Excel usa uma fonte substituta. Para trocar, veja [Personalização](#personalização).

---

## Instalação passo a passo

### 1. Baixe o código

Na página do repositório, clique em **Code → Download ZIP** e extraia em qualquer pasta. Ou, se usa Git:

```bash
git clone https://github.com/<seu-usuario>/meuAddin.git
```

### 2. Crie um arquivo em branco e abra o editor VBA

1. Abra o Excel e crie uma **pasta de trabalho em branco**.
2. Pressione **Alt + F11** para abrir o Editor do Visual Basic (VBE).

### 3. Importe os módulos

1. No VBE, menu **Arquivo → Importar arquivo...** (ou **Ctrl + M**).
2. Importe, um por vez, os quatro arquivos da pasta `src/`:
   - `modUtilitarios.bas`
   - `modInstalador.bas`
   - `modAlertasModelo.bas`
   - `modRibbon.bas`
3. Na janela **Projeto** (à esquerda), confira se os quatro aparecem dentro de *Módulos* com esses nomes exatos. O nome vem da primeira linha de cada arquivo (`Attribute VB_Name = ...`), então não precisa renomear nada.
4. Menu **Depurar → Compilar VBAProject**. Se não aparecer nenhuma mensagem, está tudo certo.

> 💡 Se os acentos aparecerem como `Ã§`, `Ã£` etc., o arquivo foi salvo com o encoding errado. Veja [Solução de problemas](#solução-de-problemas).

### 4. Salve como suplemento (.xlam)

1. Volte ao Excel (**Alt + F11** de novo ou feche o VBE).
2. **Arquivo → Salvar como → Procurar**.
3. Em *Tipo*, escolha **Suplemento do Excel (\*.xlam)**.
4. O Excel muda automaticamente a pasta para `%APPDATA%\Microsoft\AddIns`. Mantenha essa pasta (recomendado) e salve como **`meuAddin.xlam`**.
5. Feche essa pasta de trabalho.

> ⚠️ **Não proteja o projeto VBA com senha.** O Sistema de Alertas precisa ler o código do `modAlertasModelo` para copiá-lo; com o projeto bloqueado, a cópia falha.

### 5. Ative o suplemento

1. **Arquivo → Opções → Suplementos**.
2. No rodapé, em *Gerenciar*, selecione **Suplementos do Excel** e clique em **Ir...**.
3. Marque **meuAddin**. Se não aparecer, clique em **Procurar...** e aponte para o `.xlam`.
4. **OK**. A partir de agora o add-in abre junto com o Excel, em qualquer arquivo.

### 6. (Só para o Sistema de Alertas) Libere o acesso ao projeto VBA

**Arquivo → Opções → Central de Confiabilidade → Configurações da Central de Confiabilidade → Configurações de Macro** → marque **"Confiar no acesso ao modelo de objeto do projeto do VBA"**.

Sem essa opção, todas as outras macros funcionam normalmente; só o `CriarSistemaAlertas` avisa e para.

### 7. Se receber o `.xlam` pronto (em vez de montar)

Arquivos baixados da internet ou recebidos por e-mail vêm com a "marca da Web" e o Excel bloqueia as macros. Para liberar:

1. Clique com o botão direito no `.xlam` → **Propriedades**.
2. Na aba *Geral*, marque **Desbloquear** → **OK**.
3. Ou coloque o arquivo numa **pasta confiável** (Central de Confiabilidade → Locais Confiáveis).

> Dica: em vez de versionar o `.xlam` no código, publique-o em **Releases** do GitHub. Assim o repositório guarda só texto (fácil de comparar versões) e quem quiser o binário pronto baixa da release.

---

## Guia "Macros Lobão" na Faixa de Opções (opcional)

Macros de suplementos **não aparecem** na lista do Alt + F8. O jeito mais prático de usá-las é com uma guia própria na Faixa de Opções:

| Grupo | Botões |
|---|---|
| Layout | Layout 1 · Layout 2 · Fundo verde · Linhas verdes · Barra lateral |
| Dados | Formato BD · Snapshot · Copiar imagem |
| Abas | Abas de apoio |
| Auditoria | Criar alertas |

O Excel não permite criar a guia pelo VBE: o XML precisa ser embutido no arquivo. Para isso:

1. **Feche o Excel** (o `.xlam` não pode estar aberto).
2. Baixe e instale o **Office RibbonX Editor** (gratuito): <https://github.com/fernandreu/office-ribbonx-editor/releases>
3. No RibbonX Editor, **File → Open** e abra o `meuAddin.xlam`.
4. Clique com o botão direito no nome do arquivo → **Insert Office 2010+ Custom UI Part**.
5. Copie todo o conteúdo de `ribbon/customUI14.xml` e cole na área de edição.
6. Clique em **Validate** (ícone de ✓). Deve aparecer "Custom UI XML is well formed".
7. **File → Save** e abra o Excel. A guia **Macros Lobão** aparece em qualquer arquivo.

> Se um ícone ficar em branco, é só trocar o atributo `imageMso` no XML por outro (há galerias de ícones `imageMso` disponíveis na internet). Para ver mensagens de erro do XML, ative **Arquivo → Opções → Avançado → Geral → "Mostrar erros da interface do usuário de suplementos"**.

---

## Como executar as macros sem a Faixa de Opções

Se preferir não mexer no XML, há três alternativas:

1. **Barra de Ferramentas de Acesso Rápido**: Arquivo → Opções → Barra de Ferramentas de Acesso Rápido → em *Escolher comandos em*, selecione **Macros** → adicione as macros desejadas e escolha um ícone em *Modificar...*.
2. **Alt + F8 digitando o nome**: a lista vem vazia para macros de add-in, mas se você digitar o nome (ex.: `FormatoBD`) e clicar em *Executar*, funciona.
3. **Atalho de teclado**: `Copia_Imagem` já vem com **Ctrl + I** configurado.

---

## Referência das macros

### Visão geral

| Macro | Onde age | Resumo |
|---|---|---|
| [`Layout1`](#layout1) | Aba ativa | Cabeçalho verde, sem grade, fonte padrão na aba toda |
| [`Layout2`](#layout2) | Aba ativa | Cabeçalho + painel escuro + título e caixa de gráfico |
| [`FundoVerde`](#fundoverde) | Seleção | Fundo verde, fonte branca 10, centralizado |
| [`LinhasVerdes`](#linhasverdes) | Seleção | Bordas verdes médias, sem borda esquerda |
| [`BarraLateral`](#barralateral) | Seleção | Bordas escuras médias, sem borda direita |
| [`FormatoBD`](#formatobd) | Aba ativa → aba nova | Despivota tabela (períodos em colunas → linhas) |
| [`SnapshotDados`](#snapshotdados) | Aba ativa → HISTORICO | Grava os valores da aba com data/hora |
| [`Copia_Imagem`](#copia_imagem) | Seleção | Copia como imagem (Ctrl + I) |
| [`ExibirOcultarAbasComApoio`](#exibirocultarabascomapoio) | Arquivo ativo | Mostra/oculta abas com "apoio" no nome |
| [`CriarSistemaAlertas`](#sistema-de-alertas) | Arquivo ativo | Cria a aba ALERTAS de auditoria |

> ⚠️ As macros alteram formatação e criam abas. **O Ctrl + Z não desfaz ações de macro.** Em arquivos importantes, salve antes de rodar.

---

### `Layout1`

**O que faz:** aplica o layout base de relatório na aba ativa.

- Oculta as linhas de grade.
- Pinta `A1:Z5` de verde (`RGB(26,114,0)`) com altura de linha 15,75.
- Transforma `A5:Z5` numa linha fina de acento verde-claro (`RGB(0,224,0)`, altura 5).
- Deixa a coluna A estreita (0,85) como margem visual.
- Aplica a fonte padrão, tamanho 12, cor de texto do tema, em **todas as células da aba**.

**Como usar:** ative a aba que deseja formatar e rode a macro. Ideal para abas novas, antes de colar dados.

**Atenção:** como muda a fonte da aba inteira, formatos de fonte personalizados que já existirem serão sobrescritos.

---

### `Layout2`

**O que faz:** layout de painel/dashboard na aba ativa.

- Mesmo cabeçalho e linha de acento do `Layout1`, mais coluna A estreita.
- Pinta `A6:Z39` com fundo escuro (`RGB(25,43,28)`).
- Cria três formas:
  - `txtHeader` — caixa de texto "Título Aqui" (tamanho 24) sobre o cabeçalho.
  - `txtGrafico` — caixa de texto "Título do Gráfico Aqui" (14, negrito, branco).
  - `shpGrafico` — retângulo arredondado que serve de moldura para um gráfico.

**Como usar:** rode na aba do painel, depois edite os textos clicando nas caixas e posicione seu gráfico sobre o `shpGrafico`.

**Observação:** pode rodar quantas vezes quiser; as formas com esses nomes são apagadas e recriadas, nunca empilhadas. Diferente do `Layout1`, não oculta as linhas de grade nem muda a fonte das células.

---

### `FundoVerde`

**O que faz:** formata a seleção como cabeçalho de tabela — fundo verde, fonte padrão tamanho 10 em branco, alinhamento horizontal e vertical centralizado, sem quebra de texto.

**Como usar:** selecione as células do cabeçalho da tabela e rode.

---

### `LinhasVerdes`

**O que faz:** aplica bordas verdes na seleção:

- Bordas externas superior, inferior e direita médias.
- **Sem** borda esquerda.
- Linhas horizontais internas finas (separa as linhas da tabela).
- Sem linhas verticais internas.

**Como usar:** selecione o corpo da tabela e rode. Combina bem com `FundoVerde` no cabeçalho.

---

### `BarraLateral`

**O que faz:** mesma lógica do `LinhasVerdes`, mas na cor escura do painel (`RGB(25,43,28)`), **com** borda esquerda e **sem** borda direita. Útil para blocos laterais ou colunas de destaque.

**Como usar:** selecione o intervalo e rode.

---

### `FormatoBD`

**O que faz:** "despivota" uma tabela em que os períodos estão nas colunas, gerando uma tabela em formato banco de dados (uma linha por registro × período), pronta para tabela dinâmica, Power BI ou Power Query.

**Exemplo** — aba original (2 colunas de identificação):

| Região | Produto | jan | fev | mar |
|---|---|---|---|---|
| Sul | A | 10 | 12 | 9 |
| Norte | B | 7 | 8 | 11 |

Resultado na aba nova `<NomeDaAba>_BD`:

| Região | Produto | DIA | VALOR |
|---|---|---|---|
| Sul | A | jan | 10 |
| Norte | B | jan | 7 |
| Sul | A | fev | 12 |
| Norte | B | fev | 8 |
| Sul | A | mar | 9 |
| Norte | B | mar | 11 |

**Como usar:**

1. Ative a aba com a tabela. Ela precisa começar em **A1**, com os **cabeçalhos na linha 1**.
2. Rode a macro e informe **quantas colunas de identificação** existem à esquerda (no exemplo, `2`).
3. A macro cria uma aba nova no fim do arquivo com o resultado. Se já existir uma com o mesmo nome, adiciona `_1`, `_2`...

**Regras importantes:**

- A última linha é detectada pela **coluna A** — não deixe a coluna A vazia em linhas com dados.
- A última coluna é detectada pela **linha 1** — todas as colunas de período precisam ter cabeçalho.
- O progresso aparece na barra de status, no rodapé do Excel.

---

### `SnapshotDados`

**O que faz:** tira uma "foto" da aba ativa — copia **somente os valores** do intervalo usado para o fim da aba `HISTORICO` do mesmo arquivo e grava a data/hora duas colunas à direita do bloco. Se a aba `HISTORICO` não existir, é criada no fim do arquivo.

**Como usar:** ative a aba que quer registrar (ex.: um resumo diário calculado por fórmulas) e rode. Cada execução empilha um novo bloco embaixo do anterior, formando um histórico.

**Observações:** não roda com a própria `HISTORICO` ativa. Não usa a área de transferência, então não interfere no que você copiou antes.

---

### `Copia_Imagem`

**O que faz:** copia a seleção (células, gráfico ou forma) **como imagem bitmap** para a área de transferência, exatamente como aparece na tela. Basta colar (Ctrl + V) no PowerPoint, Outlook, Teams, WhatsApp etc.

**Como usar:** selecione e pressione **Ctrl + I**.

**Atenção:** enquanto o add-in estiver ativo, **Ctrl + I deixa de ser o atalho de Itálico** no Excel. Para usar outro atalho, veja [Personalização](#personalização).

---

### `ExibirOcultarAbasComApoio`

**O que faz:** alterna a visibilidade de todas as abas que têm **"apoio"** no nome (sem diferenciar maiúsculas/minúsculas: `Apoio_Base`, `base apoio`, `APOIO` etc.).

- Se **todas** estão visíveis → oculta todas.
- Se **alguma** está oculta → exibe todas.

**Como usar:** rode no arquivo desejado. O resultado aparece na barra de status. Rode de novo para inverter.

**Observação:** se as únicas abas visíveis forem de apoio, a macro avisa e não oculta (o Excel exige pelo menos uma aba visível).

---

## Sistema de Alertas

Cria, no arquivo ativo, uma aba **ALERTAS** que lista, para cada aba do arquivo, quantas células têm **erro de fórmula** e quantas têm **comentários/anotações** — com os detalhes no balão ao passar o mouse.

### Como usar

1. Garanta que a opção *"Confiar no acesso ao modelo de objeto do projeto do VBA"* está ativa ([passo 6](#6-só-para-o-sistema-de-alertas-libere-o-acesso-ao-projeto-vba)).
2. Abra o arquivo a ser auditado e rode **`CriarSistemaAlertas`** (botão *Criar alertas* na guia, se instalada).
3. A macro:
   1. copia o código de `modAlertasModelo` para um módulo novo chamado **`modAlertas` dentro do arquivo**;
   2. cria a aba **ALERTAS** como primeira aba, com cabeçalho, formatação condicional e o botão **Atualizar alertas**;
   3. roda a primeira atualização;
   4. se o arquivo for `.xlsx` ou ainda não tiver sido salvo, pede para salvar como **`.xlsm`** (senão o código seria perdido).
4. Depois disso, basta clicar em **Atualizar alertas** sempre que quiser revisar. O botão chama a macro que está no próprio arquivo, então **funciona mesmo em computadores sem o add-in**.

### O que aparece na aba

| Coluna | Conteúdo |
|---|---|
| A — Aba | Nome de cada aba do arquivo |
| B — Erros de fórmula | Quantidade de células com `#DIV/0!`, `#N/D`, `#NOME?`, `#NULO!`, `#NÚM!`, `#REF!`, `#VALOR!` etc. Fica **vermelha** quando > 0. Passe o mouse para ver os endereços e o tipo de erro. |
| C — Comentários | Quantidade de células com comentários modernos (autor + nº de respostas) e anotações. Fica **amarela** quando > 0. Passe o mouse para ver endereço e texto (truncado em 120 caracteres). |

A célula **B3** guarda a data/hora da última atualização. Cada balão lista até **60 itens**; acima disso aparece "... e mais N".

### Por que o código é copiado para o arquivo?

Para que a aba ALERTAS seja autossuficiente: quem recebe o arquivo consegue atualizar os alertas sem ter o add-in instalado. O `modAlertasModelo` usa `ThisWorkbook` de propósito — dentro do arquivo do usuário, `ThisWorkbook` passa a ser o próprio arquivo auditado.

---

## Personalização

Todas as configurações ficam no topo de `modUtilitarios`:

```vba
Private Const COR_HEADER_VERDE   As Long = 29210      ' RGB(26,114,0)
Private Const COR_LINHA_ACENTO   As Long = 57344      ' RGB(0,224,0)
Private Const COR_FUNDO_PAINEL   As Long = 1846041    ' RGB(25,43,28)
Private Const COR_BORDA_VERDE    As Long = -16748006  ' RGB(26,114,0)
Private Const COR_BORDA_LATERAL  As Long = -14931175  ' RGB(25,43,28)
Private Const COR_SHAPE_GRAFICO  As Long = 1183744    ' RGB(0,16,18)
Private Const FONTE_PADRAO       As String = "Simplon BP Regular"
```

- **Cores:** o valor Long é `R + G*256 + B*65536`. Para converter, rode no VBE (Ctrl + G, janela Imediata): `? RGB(26,114,0)` → `29210`.
- **Fonte:** troque `FONTE_PADRAO` por `"Calibri"`, `"Aptos"` ou outra instalada.
- **Atalho do Copia_Imagem:** em `src/modUtilitarios.bas`, logo abaixo de `Sub Copia_Imagem()`, edite a linha `Attribute Copia_Imagem.VB_ProcData.VB_Invoke_Func = "i\n14"` trocando o `i` por outra letra (letra maiúscula = Ctrl + Shift; ex.: `"I\n14"` vira Ctrl + Shift + I, que não conflita com o Itálico). Depois remova o módulo no VBE e importe o `.bas` de novo. Essa linha fica invisível dentro do VBE, só aparece no arquivo exportado.
- **Limite de itens nos balões do Alerta:** constante `MAX_ITENS` em `modAlertasModelo`.

---

## Atualizando o repositório depois de mexer no código

O fluxo é: editar no `.xlam` → exportar os módulos → commitar.

1. No VBE, clique com o botão direito no módulo alterado → **Exportar arquivo...** → salve por cima do `.bas` correspondente em `src/`.
2. Commit e push:

```bash
git add src/
git commit -m "FormatoBD: aceita cabeçalho fora da linha 1"
git push
```

3. Atualize o `CHANGELOG.md` e, se quiser, gere uma nova release com o `.xlam`.

### Sobre encoding (importante)

O VBE importa e exporta `.bas` em **Windows-1252 (ANSI)** com quebras de linha **CRLF**. Se um editor (VS Code, por exemplo) regravar o arquivo em UTF-8, os acentos quebram na importação. Por isso:

- O `.gitattributes` deste repositório marca `*.bas` como `-text`, para o Git nunca converter o arquivo.
- Ao editar `.bas` fora do Excel, configure o editor para **Windows 1252** e **CRLF** (no VS Code, canto inferior direito da janela). No GitHub web, os acentos podem aparecer estranhos na visualização — isso é normal e não afeta a importação.

---

## Solução de problemas

| Sintoma | Causa provável | Solução |
|---|---|---|
| Acentos viram `Ã§`, `Ã£` após importar | `.bas` foi salvo em UTF-8 | Reabra o arquivo no editor como UTF-8 e salve como **Windows 1252**; ou baixe de novo pelo ZIP do GitHub |
| Todo o código aparece em uma linha só / erro de sintaxe na importação | Quebras de linha LF em vez de CRLF | Converta para CRLF no editor e importe de novo |
| "Erro de compilação: tipo definido pelo usuário não definido" em `IRibbonControl` | Falta a referência *Microsoft Office XX.0 Object Library* | VBE → Ferramentas → Referências → marque a biblioteca do Office |
| Macro não aparece no Alt + F8 | Comportamento normal de add-in | Digite o nome no Alt + F8, use a guia personalizada ou a Barra de Acesso Rápido |
| "Ative em: Arquivo > Opções > Central de Confiabilidade..." ao criar alertas | Acesso ao modelo de objeto do VBA desativado | Siga o [passo 6](#6-só-para-o-sistema-de-alertas-libere-o-acesso-ao-projeto-vba) |
| Erro ao criar alertas mesmo com o acesso liberado | Projeto VBA do add-in protegido por senha | Remova a senha (VBE → Ferramentas → Propriedades → Proteção) |
| Excel avisa "Macros bloqueadas — fonte não confiável" | `.xlam` baixado da internet | Propriedades do arquivo → **Desbloquear**, ou mova para uma pasta confiável |
| Ctrl + I não deixa mais o texto em itálico | `Copia_Imagem` usa esse atalho | Troque o atalho (veja [Personalização](#personalização)) |
| A guia "Macros Lobão" não aparece | XML não embutido ou com erro | Revalide no RibbonX Editor e ative "Mostrar erros da interface do usuário de suplementos" |
| Botão da guia dá "não é possível executar a macro" | Nome do callback diferente do XML | Confira se `modRibbon` e `modInstalador` foram importados e compilados |

---

## Licença

Distribuído sob a licença MIT. Veja [LICENSE](LICENSE).

Feito por **Lobão**.
