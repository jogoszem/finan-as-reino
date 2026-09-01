# Reino Financeiro

Painel financeiro responsivo criado a partir da planilha `Libras.xlsx` e da base administrativa `finançassinaisdoreino.json`.

## Executar

```bash
npm install
npm run dev
```

Para gerar a versão de produção:

```bash
npm run build
```

O painel lê a primeira aba da planilha, usa a linha 25 como cabeçalho e considera como lançamento válido toda linha que possua participante e evento. Também é possível importar outra planilha com a mesma estrutura diretamente pela interface.

O arquivo `finançassinaisdoreino.json` alimenta a área “Ministério e despesas” e é complementado pelas entradas administrativas recorrentes cadastradas em `src/registered-ministry-entries.js`. As movimentações do Ministério permanecem separadas dos recebimentos dos alunos, evitando misturar bases com escopos diferentes.

O painel operacional está organizado em quatro áreas: Visão geral, Curso e alunos, Ministério e despesas e Fluxo e projeções. As antigas páginas de tabela e planejamento foram incorporadas às áreas correspondentes; as listagens filtradas são exibidas em sequência, sem paginação. A aba Apresentação permanece como modo próprio de exibição em tela cheia.

Antes de carregar os dados, a interface solicita as credenciais financeiras do Ministério. A autenticação atual é uma barreira local no navegador e mantém a sessão apenas enquanto a aba estiver aberta; para publicação em ambiente público, ela deve ser substituída por autenticação validada em servidor.

A apresentação financeira possui cinco páginas de auditoria: resumo executivo, receita líquida do curso, pendências, detalhamento das despesas e fluxo mensal consolidado. O resultado consolidado é identificado como resultado dos registros e não como saldo bancário enquanto não houver conciliação com extrato e saldo inicial.

Os lançamentos acrescentados manualmente ao painel ficam centralizados em `src/manual-records.js`. Antes de adicioná-los, a aplicação verifica nome, vencimento, parcela e valor para evitar duplicidade caso uma planilha atualizada já contenha o mesmo registro.

Participantes que não devem aparecer nem entrar nos cálculos também são declarados em `excludedParticipants`, no mesmo arquivo.
