# Painel Executivo de Demandas — CDES

Painel web responsivo para consolidar as demandas exportadas do Microsoft Planner, com indicadores executivos, filtros, pesquisa e roadmaps expansíveis por time.

## Publicar o painel

O painel é gerado como um site estático pelo Vite. No Cloudflare Pages, configure:

- **Comando de build:** `npm run build`
- **Diretório de saída:** `dist`

O build inclui o `index.html`, os arquivos JavaScript e CSS compilados e uma cópia de
`produtos-e-times.xlsx`. O painel também contém dados padrão para continuar funcionando
caso a planilha publicada não possa ser carregada.

## Atualizar a planilha

A planilha `produtos-e-times.xlsx` é lida automaticamente. Se quiser conferir outro
arquivo sem alterar o publicado, use o botão **Importar Excel**. O painel procura:

- uma aba cujo nome contenha `bucket`, usada para obter o nome do time e sua quantidade de desenvolvedores;
- outra aba com as demandas, contendo colunas equivalentes a título/tarefa, bucket/time, início, conclusão/previsão, status, progresso e prioridade.

Os nomes das colunas podem estar em português ou seguir os nomes comuns da exportação
do Planner. Para atualizar os dados publicados, basta substituir
`produtos-e-times.xlsx` por uma nova exportação com o mesmo nome.


## Desenvolvimento

Depois de alterar os arquivos em `src/`, execute `npm run build`. O comando gera o
site pronto para publicação em `dist/`. Para testá-lo localmente, execute
`npm run preview` depois do build.
