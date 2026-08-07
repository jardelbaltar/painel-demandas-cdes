# Painel Executivo de Demandas — CDES

Painel web responsivo para consolidar as demandas exportadas do Microsoft Planner, com indicadores executivos, filtros, pesquisa e roadmaps expansíveis por time.

## Executar localmente

```bash
npm install
npm run dev
```

## Importar a planilha

Use o botão **Importar Excel** e selecione `produtos-e-times.xlsx`. O painel procura automaticamente:

- uma aba cujo nome contenha `bucket`, usada para obter o nome do time e sua quantidade de desenvolvedores;
- outra aba com as demandas, contendo colunas equivalentes a título/tarefa, bucket/time, início, conclusão/previsão, status, progresso e prioridade.

Os nomes das colunas podem estar em português ou seguir os nomes comuns da exportação do Planner. Enquanto nenhuma planilha é importada, o painel apresenta dados de demonstração para permitir a avaliação da interface.
