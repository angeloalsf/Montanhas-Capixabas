# Montanhas Capixabas — Diagnóstico da oferta turística de Alfredo Chaves

Este projeto analisa o formulário de diagnóstico turístico de Alfredo Chaves (93 empreendimentos,
pesquisa de out/2025 a ago/2026) e transforma a planilha de respostas em **duas entregas que leem a
mesma base**: o buscador público do visitante (`site-visitante/`) e o painel de gestão no Power BI
(`powerbi/`).

O caminho é sempre o mesmo:

```
dados/brutos/*.xlsx  →  scripts/preparar_dados.py  →  dados/tratados/*.csv  →  site-visitante/ e powerbi/
```

**Toda a regra de negócio mora no `scripts/preparar_dados.py`.** O buscador e o `.pbix` só exibem as
colunas e tabelas que ele produz — nada é recalculado em JavaScript nem em DAX.

Estrutura revisada em 13/09/2026.

## Status

| Etapa | Situação |
|---|---|
| Levantamento da planilha, do PDF e da legenda | Concluído |
| Reorganização dos dados (modelo para o Power BI) | Concluído — `dados/tratados/` (18 CSVs) |
| Validação dos números do PDF contra a planilha | Concluído — 306 números conferidos |
| Priorização de KPIs e estrutura do painel | Concluído — `analise/relatorio-analise.md` |
| Replanejamento do conteúdo por público-alvo | Concluído — `analise/replanejamento-publico-alvo.md` |
| Buscador público do visitante (HTML) | Concluído — `site-visitante/index.html` |
| Guia de tradução do buscador para o Power BI | Concluído — `analise/guia-importacao-powerbi.md` |
| Diagnóstico técnico do modelo do `.pbix` | Concluído — `analise/diagnostico-modelo-dados.md` |
| Convenção de nomes das tabelas (`dim_`/`ponte_`/`det_`/`fato_`/`cfg_`) | Aplicada no script e nos CSVs — **pendente** no site e no injetor do mockup |
| Construção do painel de gestão no `.pbix` | **Em andamento** — `powerbi/P01 - Montanhas Capixabas.pbix` (plano de ação na §10 do diagnóstico) |

## Estrutura de pastas

```
montanhas_capixabas/
├── README.md                      ← este arquivo
├── requirements.txt               pandas e openpyxl, com versão fixada
│
├── analise/                       documentos de decisão e guias técnicos (não são gerados)
│   ├── relatorio-analise.md            achados, modelo, validação, KPIs e páginas do painel
│   ├── replanejamento-publico-alvo.md  a separação entre conteúdo de visitante e de gestão
│   ├── guia-importacao-powerbi.md      tabelas, relacionamentos, DAX e limites do Power BI
│   ├── diagnostico-modelo-dados.md     diagnóstico do .pbix + convenção de prefixos (§7-8)
│   └── validacao_pdf_x_planilha.csv    gerado por scripts/validar_pdf.py
│
├── dados/
│   ├── README.md                  de onde vem a planilha e onde o pipeline quebra
│   ├── brutos/                    fonte original — nunca alterar, nunca apagar
│   │   ├── ALFREDO CHAVES - Diagnóstico Turístico - Formulário.xlsx
│   │   ├── Estatistica descritiva - gráficos e tabelas.pdf
│   │   └── WhatsApp Image 2026-09-10 at 16.37.58.jpeg   legenda dos códigos
│   └── tratados/                  18 CSVs gerados (separador ";", UTF-8 com BOM)
│       ├── dim_*.csv              6 dimensões
│       ├── ponte_*.csv            4 pontes muitos-para-muitos
│       ├── det_*.csv              5 tabelas de detalhe do empreendimento
│       ├── fato_resposta_questionario.csv
│       ├── cfg_ficha.csv
│       └── dicionario_dados.csv   descrição de cada coluna de cada tabela
│
├── scripts/
│   ├── README.md                  tutorial de instalação e execução, do zero
│   ├── preparar_dados.py          gera dados/tratados/ — onde vive a regra de negócio
│   └── validar_pdf.py             gera analise/validacao_pdf_x_planilha.csv
│
├── powerbi/                       painel de gestão (acesso restrito)
│   ├── P01 - Montanhas Capixabas.pbix
│   ├── base.pbix                  ← a confirmar: ver "Pendências"
│   ├── recursos/                  o que é importado dentro do Power BI Desktop
│   │   ├── tema_base.json         tema aplicado no P01
│   │   └── icones/                4 PNGs usados nos visuais do P01
│   └── mockup/
│       ├── LEIA-ME.md             qual protótipo usar
│       ├── v2/                    versão atual: protótipo interativo + injetar_dados.py
│       └── v1-wireframe-tecnico/  arquivado: wireframe com notas de DAX
│
└── site-visitante/                buscador público da oferta turística
    ├── LEIA-ME.md                 o que aparece na tela e as decisões de conteúdo
    ├── index.html                 o buscador (abre com duplo clique, sem servidor)
    ├── gerar_site.py              injeta os CSVs de dados/tratados/ dentro do index.html
    └── index.backup-site-turismo.html   ← a confirmar: ver "Pendências"
```

## O que tem em cada pasta

| Pasta | Conteúdo |
|---|---|
| `dados/brutos/` | Arquivos originais, **sem alteração**: a planilha do formulário (4 abas), o PDF de estatística descritiva e a foto da legenda dos códigos. É a fonte irrecuperável do projeto. A planilha contém dados pessoais — não levar para o Power BI. |
| `dados/tratados/` | Tabelas prontas para importar. O prefixo diz o papel no modelo: `dim_` = dimensão, `ponte_` = associação muitos-para-muitos (só chaves), `det_` = detalhe/satélite do empreendimento, `fato_` = grão de evento com medida, `cfg_` = configuração de interface. A convenção está definida em `analise/diagnostico-modelo-dados.md` §7-8. Trazem endereço, horário e contato comercial (telefone/WhatsApp, Instagram, site), com publicação autorizada; nome do proprietário, CPF/CNPJ, razão social e e-mail continuam fora. |
| `analise/` | Os documentos de decisão do projeto e os guias técnicos. Nada aqui é gerado por script, com uma exceção: `validacao_pdf_x_planilha.csv`, que sai do `validar_pdf.py`. |
| `scripts/` | `preparar_dados.py` gera `dados/tratados/`; `validar_pdf.py` gera a validação. Ambos leem a planilha de `dados/brutos/` sem alterá-la. **Tutorial passo a passo: `scripts/README.md`.** |
| `powerbi/` | O painel de gestão: o `.pbix`, os recursos visuais que ele importa (`recursos/`) e os protótipos em HTML (`mockup/`). |
| `site-visitante/` | O buscador público da oferta turística, voltado a quem escolhe o que fazer no município. Ver `site-visitante/LEIA-ME.md`. |

## Como regenerar as tabelas

Na raiz do projeto, com as bibliotecas instaladas (`pip install -r requirements.txt`, de
preferência num ambiente virtual). **Se for a primeira vez, siga `scripts/README.md`** — ele
explica o passo a passo, o que esperar de saída e os erros mais comuns.

```
python scripts/preparar_dados.py "dados/brutos/ALFREDO CHAVES - Diagnóstico Turístico - Formulário.xlsx" dados/tratados
python scripts/validar_pdf.py    "dados/brutos/ALFREDO CHAVES - Diagnóstico Turístico - Formulário.xlsx" analise/validacao_pdf_x_planilha.csv
```

Depois de regerar os CSVs, atualize também os dados embutidos no buscador e no protótipo:

```
python site-visitante/gerar_site.py
python powerbi/mockup/v2/injetar_dados.py
```

> **Atenção:** esses dois últimos comandos **ainda apontam para os nomes antigos de tabela** e falham
> hoje. Ver "Pendências" abaixo.

## Pendências

### Do projeto

1. **Nomes de tabela antigos no site e no mockup.** `site-visitante/gerar_site.py`,
   `site-visitante/index.html` e `powerbi/mockup/v2/injetar_dados.py` ainda referenciam
   `fato_empreendimentos`, `fato_listas_ficha`, `fato_emp_conveniencia` e mais 6 nomes (~50
   ocorrências). Dívida assumida de propósito — detalhe em `dados/README.md` §6.
2. **O `.pbix` de gestão.** 13 tabelas importadas, 0 medidas, relações todas unidirecionais e caminho
   dos CSVs gravado literalmente. Plano de ação priorizado na §10 de
   `analise/diagnostico-modelo-dados.md`.
3. **Três decisões de base ainda abertas** (§6 do `relatorio-analise.md`): 93 ou 92 empreendimentos;
   a exclusão de 6 empreendimentos de Alimentação sem regra registrada; e a regra do nº de
   funcionários, antes de usar "empregos" como KPI.

### Arquivos aguardando decisão

| Arquivo | Por que está em dúvida |
|---|---|
| `powerbi/base.pbix` (823 KB) | **Não é deste projeto.** É o relatório "Analytics - Padrão BI Unimed Sul Capixaba", com um modelo de dados de 650 KB de outra organização e ícones de outro contexto (médico, pessoas, casa). Foi a origem do tema e dos 4 ícones hoje em `powerbi/recursos/` (bytes idênticos). Serve ainda como referência visual, ou pode sair da pasta? |
| `site-visitante/index.backup-site-turismo.html` (159 KB) | Cópia do `index.html` no estilo anterior (site de turismo), antes da reestilização Power BI. O próprio `site-visitante/LEIA-ME.md` já registra que pode ser apagada. |
| `powerbi/mockup/v1-wireframe-tecnico/painel-mockup.html` (50 KB) | Arquivado. A v2 **não substitui** o conteúdo dele: a v1 é wireframe técnico com notas de DAX, a v2 é protótipo de uso. O papel de apoio à modelagem hoje é coberto por `analise/guia-importacao-powerbi.md`. |

## Onde está documentado o quê

| Pergunta | Documento |
|---|---|
| Como rodar os scripts do zero | `scripts/README.md` |
| De onde vem a planilha e o que quebra quando o formulário muda | `dados/README.md` |
| O que cada coluna de cada tabela significa | `dados/tratados/dicionario_dados.csv` |
| Quais KPIs e páginas o painel de gestão deve ter | `analise/relatorio-analise.md` |
| Por que o conteúdo do visitante foi separado do de gestão | `analise/replanejamento-publico-alvo.md` |
| Como traduzir a lógica do buscador para tabelas, relações e DAX | `analise/guia-importacao-powerbi.md` |
| Em que estado está o `.pbix` e o que fazer nele, em ordem | `analise/diagnostico-modelo-dados.md` |
| O que o buscador mostra e o que nunca entra nele | `site-visitante/LEIA-ME.md` |
| Qual protótipo do painel de gestão usar | `powerbi/mockup/LEIA-ME.md` |
