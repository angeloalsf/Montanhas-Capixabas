# site-publico — guia da oferta turística de Alfredo Chaves

Site público de quem está escolhendo o que fazer em Alfredo Chaves, publicado no GitHub Pages
(<https://angeloalsf.github.io/Montanhas-Capixabas/>) pelo workflow `.github/workflows/pages.yml`
a cada push na `main` que mexa nesta pasta.

É uma página só (`index.html`) em HTML, CSS e JavaScript puros, **sem build**, com os CSVs de
`dados/tratados/` embutidos. Abre com duplo clique, sem servidor e sem internet: a página não faz
nenhum pedido de rede.

Redesenho de 11/10/2026: nova identidade visual, página inicial em seções e explorador com
categorias, busca, filtros e ordenação. Os dados, a busca, os filtros, a ficha e o link
`index.html#E001` continuam os mesmos. Um mapa interativo chegou a ser feito nessa versão e foi
retirado a pedido; ver "Limitações conhecidas".

## A página, de cima para baixo

1. **Topo**: marca "Montanhas Capixabas" e navegação (Categorias, Sobre os dados, Explorar
   lugares). Um link "Pular para a lista de lugares" aparece no primeiro Tab.
2. **Abertura (hero)**: título, busca, três números calculados dos dados (lugares, distritos,
   lugares na zona rural) e a **carta dos distritos**: o contorno real dos 7 distritos
   (ilustração, não interativa). Embaixo, a serra em SVG.
3. **Categorias**: as 7 de `dim_segmento` (nome, ordem e contagem vêm dos dados) e "Todos os
   lugares". Clicar filtra e leva ao explorador.
4. **Explorador** (`#explorar`):
   - barra fixa no alto com busca, botão **Filtros** e **Ordenar** (nome ou distrito);
   - chips de categoria com contagem, pílulas removíveis para cada filtro ativo (categoria,
     busca, distrito, zona, dia, conveniências) e "Limpar tudo";
   - **grade de cards** em toda a largura, **9 por página** (`POR_PAGINA` no script), com
     "Mostrando 1–9 de 93", números de página e Anterior/Próxima. Qualquer mudança de filtro,
     busca ou ordem volta para a página 1; trocar de página sobe até o início dos resultados.
     No celular a paginação mostra só primeira, atual e última (1 … 4 … 11). O link direto
     `#E090` abre a ficha em qualquer página.
5. **Sobre os dados** (`#sobre`): para que serve o guia e de onde vêm as informações.
6. **Ficha do lugar**: modal com "Como chegar", "Quando abre",
   "Estrutura e atendimento" e todos os contatos. Abre por `index.html#E001`; "voltar" fecha.

## Uma regra de filtro só

Toda contagem e todo resultado da página saem da mesma função, `passa(e, ignorar)`: os
cards, as categorias do topo, os chips e os números do painel de filtros. `ignorar` só deixa a
categoria de fora para os chips mostrarem quantos lugares teriam se fossem escolhidos. O estado
também é um só (`estado`): o topo, os chips, o painel e as pílulas só leem e escrevem nele.

## A carta dos distritos (`geo-distritos`)

- Os 7 distritos do IBGE para o município 3200300 batem 1:1 com a coluna `distrito` dos dados.
  Os nomes diferem em dois: **Sede** = "Alfredo Chaves" e **São Bento de Urânia** = "Urânia"
  no IBGE.
- O contorno vem das relações de limite administrativo do OpenStreetMap (ids 1902985 a
  1902991), que trazem o código IBGE de cada distrito. Licença ODbL. Simplificado (~30 m) e
  guardado como polilinha codificada (~3,7 KB). Os pontos são as vilas-sede (nós
  `place=town/village` do OSM). Extraído em 11/10/2026.
- O JavaScript desenha a carta em SVG na abertura (`heroCarta()`). Ela some em telas
  estreitas. O bloco não é gerado pelo `gerar_site.py`, e rodar o script não o altera.

## Identidade visual

- **Direção**: guia de viagem editorial + carta topográfica. Nenhuma fotografia, nenhuma
  imagem externa, nenhuma fonte web.
- **Paleta** (variáveis no `:root`): verde-floresta `#1F3D2B` (marca, títulos, ações
  principais), verdes de sálvia para superfícies, creme `#F6F1E7` de fundo,
  terra `#9A4F2A` para sobretítulos, foco e destaques. Texto em `#1C211D`/`#4B534D`/`#646A64`
  (contraste ≥ 4,4:1 sobre creme e papel). O verde do WhatsApp continua para o botão de contato.
- **Tipografia**: títulos em serifa do sistema (Iowan Old Style / Palatino Linotype / Book
  Antiqua / Noto Serif / Georgia), corpo em `system-ui`. O peso editorial vem do contraste
  entre a serifa grande e os sobretítulos em caixa-alta espaçada.
- **Grafismos** (`assets/`), gerados por código e não por ilustração livre:
  - `relevo.svg`: linhas de nível de um relevo sintético, com curva mestra a cada cinco,
    como numa carta topográfica. Fundo do hero e do cabeçalho da ficha.
  - `serra.svg`: cinco cadeias de montanha em perspectiva atmosférica; a última tem o verde da
    seção seguinte, para a transição ser contínua.
  - A **carta dos distritos** do hero é desenhada em SVG pelo JavaScript a partir do bloco
    `geo-distritos`. É geografia real, não ilustração livre.
- **Ícones**: SVG 16×16 em traço de 1,3, desenhados no código (`ICONES`). Os do card vêm do
  visual do Power BI; os de categoria (talheres, cama, cachoeira, pote, sacola, montanha com
  bandeira, colunas) seguem o mesmo traço. Os de
  WhatsApp, Instagram e site continuam genéricos (balão, câmera, globo), sem reproduzir marcas.
- **Cor de categoria** (`APARENCIA_SEG`): só apresentação. Aparece no filete do card, no
  ponto dos chips e no ícone das categorias. Lugar com mais de uma categoria usa a
  primeira pela `ordem_chip`. Categoria nova sem entrada usa um verde neutro.

## Acessibilidade e responsividade

- Mobile-first na interação: os chips rolam na horizontal; ficha e painel de filtros ocupam a
  tela inteira. Verificado sem rolagem horizontal de 320 px a 1920 px.
- Alvos de toque de 36 a 46 px; foco visível (contorno terra de 3 px) em todo controle.
- Tudo funciona por teclado: busca, categorias e chips (`aria-pressed`), painel, ficha (Esc
  fecha).
- Total de resultados em `aria-live`.
- `prefers-reduced-motion`: sem animações de abertura e sem rolagem suave.

## Arquivos

| Arquivo | O que é |
|---|---|
| `index.html` | O site: estrutura, estilos, dados embutidos e o script. |
| `gerar_site.py` | Injeta os CSVs de `dados/tratados/` no `index.html` (sem mudanças nesta versão). |
| `assets/relevo.svg`, `assets/serra.svg` | Grafismos da identidade. |

## Como atualizar depois de mexer nos dados

```
python scripts/preparar_dados.py "dados/brutos/ALFREDO CHAVES - Diagnóstico Turístico - Formulário.xlsx" dados/tratados
python site-publico/gerar_site.py
```

O `gerar_site.py` usa os **nomes atuais** das tabelas (`dim_` / `ponte_`, convenção da §7-8
de `analise/diagnostico-modelo-dados.md`). Ele não tem a dívida de nomes antigos registrada
na seção "Pendências" do `README.md`; aquela continua valendo só para
`site-visitante/gerar_site.py` e `powerbi/mockup/v2/injetar_dados.py`.

Rodar o script com os CSVs atuais deixa o `index.html` **idêntico byte a byte** (conferido em
11/10/2026). Os dados embutidos hoje são exatamente os de
`dados/tratados/`.

O navegador bloqueia `fetch` de arquivo local quando a página é aberta do disco (`file://`).
Por isso os CSVs ficam embutidos em blocos `<script type="text/plain" id="csv-...">` e são
lidos por um parser de CSV em JavaScript na abertura. É o mesmo mecanismo do protótipo v2 do
painel.

## O que vem de onde

| Bloco embutido | Tabela | Para quê |
|---|---|---|
| `csv-dim_empreendimento` | `dim_empreendimento` (31 colunas) | tudo o que aparece no card e na ficha |
| `csv-dim_segmento` | `dim_segmento` | categorias e chips, em linguagem de visitante e na ordem oficial |
| `csv-dim_regiao` | `dim_regiao` | a lista de distritos do painel |
| `csv-dim_dia` | `dim_dia` | a lista de dias |
| `csv-dim_conveniencia` | `dim_conveniencia` | as 7 caixas de seleção |
| `csv-ponte_empreendimento_segmento` | ponte | de quais categorias cada lugar é |
| `csv-ponte_empreendimento_dia` | ponte | o filtro "aberto em" |
| `csv-ponte_empreendimento_conveniencia` | ponte | o filtro de conveniências |
| `geo-distritos` | OSM, extraído uma vez (não é gerado pelo script) | a carta dos distritos do topo |

**O site não calcula regra de negócio.** Quem é "acessível", quem "aceita pet", o texto de
cada selo, o resumo de dias e a faixa de grupos vêm prontos das colunas do `preparar_dados.py`.
Os números do topo e as contagens são só contagem de linhas. Para mudar uma regra, mexa no
Python e regere os CSVs, não no JavaScript.

## O card, campo a campo

Os nomes de classe (`ce-*`), os ícones, a regra de telefone e a divisão dos selos continuam os
do visual `CardEmpreendimento`. A ordem visual mudou: as categorias viraram sobretítulo e o
distrito ganhou linha própria.

| Parte do card | Origem |
|---|---|
| Sobretítulo (categorias) | `segmentos_rotulo` |
| Título | `nome_empreendimento` |
| Distrito · zona | `distrito` · `localizacao_rotulo` |
| Endereço | `endereco_completo` |
| Dias | `dias_resumo` (+ "· abre em feriados" quando `abre_feriados`) |
| Horário | `horario_atendimento` (ou "24 horas" quando `atende_24h`) |
| Tarja âmbar de via | `condicao_via` (só "Não pavimentada" e "Mista") |
| Selos (Sim/Não) | `aceita_cartao`, `tem_estacionamento_proprio`, `espaco_infantil`, `atende_24h` |
| Selos (texto) | `pet_texto`, `grupos_capacidade` |
| Selos (texto pronto) | `badge_acessibilidade`, `badge_idiomas` |
| Botões | `telefones`, `whatsapp`, `instagram`, `site` |
| Filete colorido | cor da primeira categoria (apresentação) |

A regra de telefone é a mesma de `src/contato.ts`: o número que também é WhatsApp aparece
uma vez, e o sufixo "· WhatsApp" só entra quando o card tem mais de um número.

## O que nunca entra neste HTML

Formalização (CNPJ/CPF), nº de funcionários, ano de início, público principal, canal de
divulgação, necessidades do município e do negócio, capacidade total de hospedagem, preço e
a autodeclaração `facilidade_pcd`/`capacitado_pcd`. Tudo isso é conteúdo do painel de gestão,
com acesso restrito. A lista de colunas está fixada em `COLS_EMP`, no topo do `gerar_site.py`.

Nome do proprietário, CPF/CNPJ, razão social e e-mail continuam fora dos CSVs: não estão nem
nos dados tratados nem aqui.

## Limitações conhecidas

- **Sem mapa interativo.** Uma versão com mapa Leaflet (distritos coloridos pelos filtros e
  pontos só para coordenadas verificadas) foi feita e retirada em 11/10/2026. Registro do que
  pesa numa próxima tentativa: nenhum dos 93 lugares tem coordenada, já que o formulário
  guardou só o endereço em texto livre e só 24 endereços têm número. Um mapa de pontos depende
  de levantar coordenadas verificadas (GPS no local ou enviadas pelo estabelecimento), e não de
  geocodificar os endereços.
- **Base de 93 lugares.** O `dados/brutos/Dados tratados.xlsx` (92 lugares, categorias de via
  mais detalhadas) ainda não foi incorporado ao `preparar_dados.py`. O site mostra o que está
  em `dados/tratados/` hoje.
- Endereços são texto livre declarado pelo estabelecimento; o site os mostra como vieram.
- Não há testes automatizados nem lint no repositório. As verificações desta versão foram
  feitas com Playwright fora do repositório (lista no resumo da entrega).
