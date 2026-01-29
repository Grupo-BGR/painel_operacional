# Dashboard de Fluxo de Importação (HTML/CSS/JS)

Tela simples de dashboard para listar **vários processos de importação** e exibir a **fase atual** no fluxo:

A embarcar → Em Trânsito → A Registrar → A Desembaraçar → A Carregar

## Como usar

- Abra o arquivo `index.html` no navegador.
- Use **Buscar** e **Filtrar por fase** no topo.
- Os processos ficam **agrupados por cliente** e você pode **expandir/colapsar** cada grupo.

## Onde editar os dados

O dashboard tenta carregar automaticamente o arquivo `processos geral.xlsx` (na raiz).

- Se você abrir via **duplo clique** (modo `file://`), o navegador geralmente **bloqueia** a leitura automática do Excel. Nesse caso, selecione o arquivo no campo **Excel** no topo.
- Se você abrir por um **servidor local** (ex.: Live Server), o carregamento automático da raiz funciona.

## Como o Excel é interpretado

O `app.js` tenta reconhecer colunas pelo nome (ex.: `id`, `cliente`, `fornecedor`, `modal`, `fase atual/status`) e também datas por etapa (colunas contendo o nome da etapa).

Se os nomes das suas colunas forem diferentes, ajuste as “aliases” dentro do `app.js` (procure por `pickFirstHeader` / `mapRowsToProcesses`).

- `id`
- `cliente`
- `fornecedor`
- `modal`
- `faseAtual` (uma das fases do fluxo)
- `status` (`ok`, `warn`, `danger`)

