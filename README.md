# Reino Financeiro

Painel financeiro responsivo criado a partir da planilha `Libras.xlsx` e da base administrativa `finançassinaisdoreino.csv`.

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

O arquivo `finançassinaisdoreino.csv` alimenta a aba “Finanças do Ministério”. Essa base permanece separada dos recebimentos dos alunos: somente o resumo de despesas é apresentado também na Visão geral, sem produzir um saldo combinado entre bases de escopos diferentes.

Os lançamentos acrescentados manualmente ao painel ficam centralizados em `src/manual-records.js`. Antes de adicioná-los, a aplicação verifica nome, vencimento, parcela e valor para evitar duplicidade caso uma planilha atualizada já contenha o mesmo registro.

Participantes que não devem aparecer nem entrar nos cálculos também são declarados em `excludedParticipants`, no mesmo arquivo.
