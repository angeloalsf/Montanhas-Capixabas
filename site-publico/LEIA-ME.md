# site-publico — o site da oferta turística, feito do card do Power BI

Site público de quem está escolhendo o que fazer em Alfredo Chaves. **O card é o visual
`powerbi/visuais/CardEmpreendimento` reescrito para a web**: mesma estrutura de DOM, mesmos
nomes de classe (`ce-*`), mesma paleta, os mesmos ícones em SVG desenhados em código e a
mesma ordem de montagem. Em volta do card há o que o Power BI resolveria com slicers —
busca, abas de segmento e um painel de filtros — e a ficha completa de cada lugar num modal.

## A página

- **Cabeçalho** no verde do site da Prefeitura de Alfredo Chaves (`#005946`, a
  `--cor-primaria` deles; o fundo da página é o `#F2F4F3` deles) com o título, a busca e o
  botão **Filtros**. Não é fixo: rola junto com a página.
- **Abas de segmento** (Todos, Onde comer, Onde dormir…) logo abaixo, com a contagem de cada
  uma já considerando os outros filtros. Uma aba por vez.
- **Painel de filtros** que abre pela direita (tela cheia no celular): distrito, zona
  urbana/rural, dia de funcionamento e o que o lugar oferece. Os filtros ativos aparecem como
  pílulas acima da grade, cada uma com um × para remover.
- **Card como resumo de altura fixa.** Endereço, dias e horário têm uma linha cada (o texto
  longo termina em "…"; vazio vira "não informado"), a tarja de via e os selos ocupam no
  máximo duas linhas e o que não cabe vira um "+N", e os contatos ficam num rodapé: um número,
  WhatsApp e ícones de Instagram e site. Assim os cards saem praticamente do mesmo tamanho
  sem espaço vazio. Tudo o que o resumo corta está na ficha.
- **Ficha do lugar**: clicar no card (ou no título) abre um modal com tudo organizado em
  "Como chegar", "Quando abre" e "Estrutura e atendimento", e os botões de contato no rodapé.
  A ficha muda o endereço da página para `index.html#E001`, então dá para mandar o link de
  um lugar específico, e o botão "voltar" do navegador fecha a ficha.

Esta pasta é independente: **não altera nem depende de `site-visitante/`, de `powerbi/`
ou do visual.** As duas entregas leem os mesmos CSVs de `dados/tratados/`.

## Arquivos

| Arquivo | O que é |
|---|---|
| `index.html` | O site. Abre com duplo clique, sem servidor e sem internet. |
| `gerar_site.py` | Injeta os dados de `dados/tratados/` dentro do `index.html`. |

## Como atualizar depois de mexer nos dados

```
python scripts/preparar_dados.py "dados/brutos/ALFREDO CHAVES - Diagnóstico Turístico - Formulário.xlsx" dados/tratados
python site-publico/gerar_site.py
```

O `gerar_site.py` usa os **nomes atuais** das tabelas (`dim_` / `ponte_`, convenção da §7-8
de `analise/diagnostico-modelo-dados.md`). Ele não tem a dívida de nomes antigos registrada
na seção "Pendências" do `README.md` — aquela continua valendo só para
`site-visitante/gerar_site.py` e `powerbi/mockup/v2/injetar_dados.py`.

O navegador bloqueia `fetch` de arquivo local quando a página é aberta do disco (`file://`).
Por isso os CSVs ficam embutidos em blocos `<script type="text/plain" id="csv-...">` e são
lidos por um parser de CSV em JavaScript na abertura — é o mesmo mecanismo do protótipo v2
do painel.

## O que vem de onde

| Bloco embutido | Tabela | Para quê |
|---|---|---|
| `csv-dim_empreendimento` | `dim_empreendimento` (31 colunas) | tudo o que aparece no card |
| `csv-dim_segmento` | `dim_segmento` | os chips, em linguagem de visitante e na ordem oficial |
| `csv-dim_regiao` | `dim_regiao` | a lista de distritos |
| `csv-dim_dia` | `dim_dia` | a lista de dias |
| `csv-dim_conveniencia` | `dim_conveniencia` | as 7 caixas de seleção |
| `csv-ponte_empreendimento_segmento` | ponte | de quais segmentos cada lugar é |
| `csv-ponte_empreendimento_dia` | ponte | o filtro "aberto em" |
| `csv-ponte_empreendimento_conveniencia` | ponte | o filtro de conveniências |

**O site não calcula nada.** Quem é "acessível", quem "aceita pet", o texto de cada selo, o
resumo de dias e a faixa de grupos vêm prontos das colunas do `preparar_dados.py`. Para mudar
uma dessas regras, mexa no Python e regere os CSVs — não no JavaScript, que só exibe. É a
mesma disciplina do `site-visitante/` e do `.pbix`.

## O card, campo a campo

A montagem segue `src/listaCards.ts` na mesma ordem, e os selos seguem a divisão de
`capabilities.json`:

| Parte do card | Origem |
|---|---|
| Título | `nome_empreendimento` |
| Segmentos · local | `segmentos_rotulo` · `distrito` · `localizacao_rotulo` |
| 📍 endereço | `endereco_completo` |
| 📅 dias | `dias_resumo` (+ "· abre em feriados" quando `abre_feriados`) |
| 🕐 horário | `horario_atendimento` |
| Referência | `ponto_referencia` |
| Tarja âmbar de via | `condicao_via` |
| Selos (Sim/Não) | `aceita_cartao`, `tem_estacionamento_proprio`, `espaco_infantil`, `atende_24h` |
| Selos (texto) | `pet_texto`, `grupos_capacidade` |
| Selos (texto pronto) | `badge_acessibilidade`, `badge_idiomas` |
| Botões | `telefones`, `whatsapp`, `instagram`, `site` |
| Ficha (modal) | o card inteiro + `dias_funcionamento`, `via_acesso`, `estacionamento`, `vagas_*`, `idiomas`, `sanitario_adaptado` |

A regra de telefone é a mesma de `src/contato.ts`: o número que também é WhatsApp aparece
uma vez, e o sufixo "· WhatsApp" só entra quando o card tem mais de um número — dos 93, três
têm dois.

## Diferenças assumidas em relação ao visual

As três primeiras existem porque o visual roda num iframe restrito do Power BI e o site não:

1. **"Ligar" disca.** No visual o clique copiava o número, porque o host só abre `http`/`https`
   e `tel:` não chegava a lugar nenhum. Aqui o botão é um link `tel:`.
2. **O ▼ virou a ficha.** No visual ele era só um marcador (a ficha vinha de outro grão).
   Aqui o card inteiro abre um modal montado com colunas que já estão em
   `dim_empreendimento` — nenhuma tabela nova.
3. **`localizacao_rotulo` e `badge_idiomas`.** As duas colunas que faltavam quando o visual
   foi escrito já existem no CSV; o site usa "zona rural"/"zona urbana" e "Atende em inglês"
   em vez do texto cru. São as duas pendências da seção "Limitações conhecidas" do
   `CardEmpreendimento/README.md`.

E duas decisões de conteúdo:

4. **A tarja âmbar só aparece quando é aviso.** O visual gerava selo para os três valores de
   `condicao_via`, inclusive "Pavimentada". Aqui só "Não pavimentada" e "Mista" viram tarja —
   é a regra do `site-visitante/LEIA-ME.md`: condição da via é aviso individual para quem vai
   chegar, nunca estatística.
5. **"Não recebe grupos" e "Não se aplica" saem sem prefixo**, porque "Recebe grupos: Não
   recebe grupos" não é frase.

Fora isso: a página rola normalmente (não há canvas fixo de 1280×720 como no `site-visitante/`),
a grade é responsiva, e os cards de uma mesma linha têm a mesma altura — que é o efeito do
`uniformizarAltura` do visual, aqui obtido só com CSS Grid.

### Tamanho da letra

O card herda tudo de `--base`, no `:root` do `index.html`. O valor entregue é **`13px`**,
que lê melhor num site; `12px` reproduz o visual do Power BI pixel a pixel. Mude só essa
linha, o resto do card acompanha.

## O que nunca entra neste HTML

Formalização (CNPJ/CPF), nº de funcionários, ano de início, público principal, canal de
divulgação, necessidades do município e do negócio, capacidade total de hospedagem, preço e
a autodeclaração `facilidade_pcd`/`capacitado_pcd`. Tudo isso é conteúdo do painel de gestão
(`analise/relatorio-analise.md`), com acesso restrito — a lista de colunas está fixada em
`COLS_EMP`, no topo do `gerar_site.py`.

Nome do proprietário, CPF/CNPJ, razão social e e-mail continuam fora dos CSVs — não estão nem
nos dados tratados nem aqui.

## Restrições técnicas

Nenhum recurso externo: sem CDN, sem web font, sem imagem remota, sem `fetch`. Os ícones são
SVG desenhado em código, portados de `src/icones.ts` — os de WhatsApp, Instagram e site são
genéricos (balão, câmera e globo) e não reproduzem as marcas dessas empresas; quem identifica
o serviço é o texto do botão.
