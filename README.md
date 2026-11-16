# Simulador de Financiamento

Página que calcula a parcela usando as **fórmulas da própria planilha** `simulador.xlsm`.
O navegador baixa a planilha e recalcula tudo localmente (sem servidor, sem Excel).

## Atualizar a planilha
Substitua o arquivo `simulador.xlsm` no repositório por uma nova versão **com o mesmo nome**.
O site passa a usar a nova versão automaticamente.

Se mudar a posição dos campos na aba "Simula", ajuste o bloco `ENTRADAS` / `SAIDAS` no `index.html`.

## Arquivos
- `index.html` – tela (campos verdes + resultado)
- `engine.js` – lê a planilha e recalcula as fórmulas (inclui a função XIRR)
- `simulador.xlsm` – a planilha
- `vendor/` – bibliotecas SheetJS (Apache-2.0) e HyperFormula (GPL-3.0)

## Testar no computador
Abrir o `index.html` direto não funciona (o navegador bloqueia a leitura do arquivo).
Use um servidor local, por exemplo: `python -m http.server` e acesse http://localhost:8000
