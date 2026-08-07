# Painel Executivo de Demandas — CDES

Painel web responsivo para consolidar as demandas exportadas do Microsoft Planner, com indicadores executivos, filtros, pesquisa e roadmaps expansíveis por time.

## Acessar o painel

O painel é uma página estática: abra `index.html` diretamente ou publique a raiz
deste repositório (por exemplo, com GitHub Pages). O arquivo já inclui uma cópia dos
dados da planilha para funcionar mesmo sem servidor. Quando publicado, ele também
tenta carregar automaticamente `produtos-e-times.xlsx` para obter a versão mais recente.

## Atualizar a planilha

A planilha `produtos-e-times.xlsx` é lida automaticamente. Se quiser conferir outro
arquivo sem alterar o publicado, use o botão **Importar Excel**. O painel procura:

- uma aba cujo nome contenha `bucket`, usada para obter o nome do time e sua quantidade de desenvolvedores;
- outra aba com as demandas, contendo colunas equivalentes a título/tarefa, bucket/time, início, conclusão/previsão, status, progresso e prioridade.

Os nomes das colunas podem estar em português ou seguir os nomes comuns da exportação
do Planner. Para atualizar os dados publicados, basta substituir
`produtos-e-times.xlsx` por uma nova exportação com o mesmo nome.


## Desenvolvimento

Depois de alterar os arquivos em `src/`, execute `npm run build`. O comando gera e
copia `dashboard.js` e `dashboard.css` para a raiz, que são os arquivos carregados
pelo `index.html`.
