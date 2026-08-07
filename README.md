# Painel Executivo de Demandas — CDES

Painel web responsivo para consolidar as demandas exportadas do Microsoft Planner, com indicadores executivos, filtros, pesquisa e roadmaps expansíveis por time.

## Acessar o painel

O painel é uma página estática: publique a raiz deste repositório (por exemplo, com
GitHub Pages) e acesse `index.html`. Ao abrir, a página carrega automaticamente o
arquivo `produtos-e-times.xlsx` que está na mesma pasta, sem exigir instalação,
build ou servidor de desenvolvimento.

## Atualizar a planilha

A planilha `produtos-e-times.xlsx` é lida automaticamente. Se quiser conferir outro
arquivo sem alterar o publicado, use o botão **Importar Excel**. O painel procura:

- uma aba cujo nome contenha `bucket`, usada para obter o nome do time e sua quantidade de desenvolvedores;
- outra aba com as demandas, contendo colunas equivalentes a título/tarefa, bucket/time, início, conclusão/previsão, status, progresso e prioridade.

Os nomes das colunas podem estar em português ou seguir os nomes comuns da exportação
do Planner. Para atualizar os dados publicados, basta substituir
`produtos-e-times.xlsx` por uma nova exportação com o mesmo nome.
