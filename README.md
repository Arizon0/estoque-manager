# Estoque Vivo

Controle de estoque da empresa, acessível pelo celular.

- **App (celular):** https://claude.ai/artifact/JiMT1jfH7Vc3MUYtYcDNG7
- **Código da página:** [`app/index.html`](app/index.html)

O app é uma página publicada com banco de dados próprio: o saldo fica salvo no
servidor, não no aparelho. Abrir pelo celular, pelo computador ou por outra
pessoa com acesso mostra sempre o mesmo estoque, atualizado na hora.

## O que dá para fazer pelo celular

| Ação | Onde |
| --- | --- |
| Ver saldo, valor em custo e alertas de reposição | topo da aba **Estoque** |
| Filtrar só o que está baixo ou zerado | chips abaixo da busca (ou tocar na placa **Alertas**) |
| Buscar por nome, código ou categoria | campo de busca |
| Dar baixa de uma venda | tocar no item → **Registrar venda** |
| Lançar uma compra / reposição | tocar no item → **Entrada** |
| Acertar o saldo depois de uma contagem | tocar no item → **Corrigir saldo** |
| Cadastrar ou editar um produto | **Novo item** / **Editar cadastro** |
| Ver o histórico dia a dia | aba **Movimentos** |
| Baixar um backup da posição atual | **Exportar estoque (CSV)** |

Todo lançamento grava data, hora, quantidade, saldo resultante e observação.
Nada é apagado quando um item é excluído — o histórico continua na aba
Movimentos.

## Rotina combinada

1. A planilha com as quantidades é enviada no chat; o estoque é carregado a
   partir dela (substituindo os itens de exemplo).
2. A cada 48h chega o relatório de vendas; cada linha vira uma baixa no item
   correspondente, marcada como `relatório 48h` no histórico.
3. O saldo novo aparece no celular assim que o lançamento é gravado.

Correções de última hora podem ser feitas direto no app, sem esperar o
próximo relatório.

## Formato da planilha

Qualquer planilha serve — o importante é ter o nome do produto e a
quantidade. Quanto mais colunas abaixo vierem preenchidas, mais completo fica
o controle. [`dados/modelo-importacao.csv`](dados/modelo-importacao.csv) traz
o formato ideal.

| Coluna | Obrigatória | Para que serve |
| --- | --- | --- |
| `nome` | sim | identifica o produto |
| `quantidade` | sim | saldo atual |
| `codigo` | não | SKU, usado para casar as linhas do relatório de vendas |
| `categoria` | não | agrupa e facilita a busca |
| `unidade` | não | `un`, `cx`, `kg`, `rolo`… (padrão: `un`) |
| `minimo` | não | dispara o alerta de reposição |
| `custo` | não | calcula o valor total do estoque |
| `preco` | não | preço de venda, só para referência |

Números em português (`1.250,75`) são aceitos.

## Como os dados são guardados

Duas coleções no banco do app:

- `itens/<id>` — `nome`, `sku`, `categoria`, `unidade`, `qtd`, `minimo`,
  `custo`, `preco`, `atualizadoEm`.
- `movimentos/<AAAA-MM-DD>` — um documento por dia, com `lancamentos[]`:
  `itemId`, `nome`, `sku`, `tipo` (`venda` | `entrada` | `ajuste`), `delta`,
  `saldo`, `obs`, `em`, `origem` (`app` | `relatorio`).

O histórico é agrupado por dia de propósito: o banco do app aceita até 5.000
documentos, e um documento por venda estouraria esse limite com o tempo.

## Editar o app

`app/index.html` é a página inteira (HTML, CSS e JavaScript em um arquivo só,
sem dependências). Para publicar uma alteração, o mesmo arquivo é republicado
na URL acima — quem estiver com a página aberta recebe a nova versão sozinho,
e os dados do estoque não são afetados.
