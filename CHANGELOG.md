# Changelog

## [2.1] - 2026-11-11

Primeira versão publicada no GitHub.

### Corrigido
- `ExibirOcultarAbasComApoio`: usava `ThisWorkbook`, que dentro de um add-in aponta para o próprio `.xlam`, e não para o arquivo aberto. Agora age sobre o arquivo ativo e avisa se a operação deixaria o arquivo sem nenhuma aba visível.
- `SnapshotDados`: procurava a aba `HISTORICO` no add-in (`ThisWorkbook`) mas criava no arquivo ativo, o que gerava erro de nome duplicado a partir da segunda execução. Agora tudo acontece no arquivo ativo, a `HISTORICO` é criada no fim, os valores são copiados sem usar a área de transferência e a data/hora recebe formato.
- `FormatoBD`: cancelar a caixa de entrada não mostra mais mensagem de "valor inválido".
- Comentário da constante `COR_SHAPE_GRAFICO` dizia `RGB(24,114,0)`, mas o valor `1183744` corresponde a `RGB(0,16,18)`. O valor foi mantido (é o que já era aplicado) e o comentário corrigido.

### Alterado
- `Módulo1` renomeado para `modUtilitarios`.
- Fonte centralizada na constante `FONTE_PADRAO`.
- Macros que agem sobre a aba ativa verificam se há arquivo aberto e se a aba é uma planilha (não um gráfico).
- Atalho Ctrl + I de `Copia_Imagem` gravado no próprio `.bas` (preservado na importação).
- `CriarSistemaAlertas` avisa quando executado com o próprio add-in ativo, em vez de sair em silêncio.

### Adicionado
- `modRibbon` + `ribbon/customUI14.xml`: guia "Macros Lobão" na Faixa de Opções (opcional).
