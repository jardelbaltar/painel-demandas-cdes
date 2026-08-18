# Painel Executivo de Demandas — CDES

Painel web responsivo para consolidar as demandas do Microsoft Planner, com autenticação delegada do usuário, fallback para Excel, indicadores executivos, filtros, pesquisa e roadmaps expansíveis por time.

## Integração com o Microsoft Planner

O painel tenta reutilizar silenciosamente a sessão Microsoft do usuário. Quando não há uma sessão válida, ele carrega `produtos-e-times.xlsx` em modo de contingência e oferece o botão **Conectar ao Planner**. Somente usuários que conseguirem entrar no tenant e tiverem acesso ao plano verão os dados sincronizados. A fonte ativa fica sempre identificada no topo do painel.

Crie um registro de aplicativo do tipo **Single-page application (SPA)** no Microsoft Entra ID, cadastre as URLs de produção e desenvolvimento como redirect URIs e conceda as permissões delegadas Microsoft Graph `User.Read`, `Tasks.Read` e `Group.Read.All`. Depois, configure no ambiente de build:

```env
VITE_MICROSOFT_CLIENT_ID=ee2b54f7-4e41-43b3-a4a3-a77280e9acc5
VITE_MICROSOFT_TENANT_ID=de23d5f0-ccac-4c84-81d6-2892a8c055aa
VITE_PLANNER_PLAN_ID=_IFjpmPlW02Q7eVsII-VQmQADmgL
```

Não configure client secret: a autenticação SPA usa as credenciais e permissões do próprio usuário com Authorization Code + PKCE. Variáveis `VITE_*` são públicas no bundle e devem conter apenas identificadores não secretos.

O Object ID (`d76731bb-4da7-4757-a0de-da60b3300ddc`) identifica o objeto do registro dentro do Entra ID e não é necessário na configuração OAuth do painel.

Buckets com o padrão `Nome do time - N`, por exemplo `Julgamento - 3`, são apresentados como time **Julgamento** com **3 desenvolvedores**.
Ao sincronizar, os buckets são exibidos conforme a ordem definida no Planner (campo `orderHint`), preservando a mesma sequência encontrada na exportação da planilha.

## Publicar o painel

O painel é gerado como um site estático pelo Vite. No Cloudflare Pages, configure:

- **Comando de build:** `npm run build`
- **Diretório de saída:** `dist`

O build inclui o `index.html`, os arquivos JavaScript e CSS compilados e uma cópia de
`produtos-e-times.xlsx`. O painel também contém dados padrão para continuar funcionando
caso a planilha publicada não possa ser carregada.

## Atualizar a planilha

A planilha `produtos-e-times.xlsx` é lida automaticamente. O painel procura:

- uma aba cujo nome contenha `bucket`, usada para obter o nome do time e sua quantidade de desenvolvedores;
- outra aba com as demandas, contendo colunas equivalentes a título/tarefa, bucket/time, início, conclusão/previsão, status, progresso e prioridade.

Os nomes das colunas podem estar em português ou seguir os nomes comuns da exportação
do Planner. Para atualizar os dados publicados, basta substituir
`produtos-e-times.xlsx` por uma nova exportação com o mesmo nome.


## Desenvolvimento

Depois de alterar os arquivos em `src/`, execute `npm run build`. O comando gera o
site pronto para publicação em `dist/`. Para testá-lo localmente, execute
`npm run preview` depois do build.
