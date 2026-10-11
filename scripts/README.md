# `scripts/preparar_dados.py` — como rodar do zero

Tutorial passo a passo. Não é preciso entender o código nem o projeto: siga as seções na
ordem. Tempo total na primeira vez: ~5 minutos (a execução em si leva menos de 10 segundos).

---

## 1. O que este script faz

Lê a planilha bruta do formulário de diagnóstico turístico de Alfredo Chaves (o `.xlsx`
exportado do Google Sheets, em `dados/brutos/`), aplica a legenda de códigos da pesquisa e as
regras de negócio do projeto, e grava **17 tabelas tratadas + 1 dicionário de dados** em CSV
na pasta `dados/tratados/`. Esses CSVs são o que o Power BI e o buscador do visitante
consomem — nenhum dos dois recalcula nada.

O script **não altera a planilha de entrada**: só lê. Rodar duas vezes seguidas produz
exatamente o mesmo resultado (os CSVs são sobrescritos).

---

## 2. Pré-requisitos

**Python 3.** Testado com **Python 3.10.12** (Ubuntu 22.04 / WSL). Deve funcionar em qualquer
Python 3.9 ou superior, mas só a 3.10 foi verificada.

Confira o que você tem:

```bash
python3 --version     # Linux, Mac, WSL
python --version      # Windows (PowerShell ou Prompt de Comando)
```

Se o comando não existir, instale o Python (<https://www.python.org/downloads/> no Windows;
`sudo apt install python3 python3-venv` no Ubuntu/WSL).

### Ambiente virtual (recomendado)

Um ambiente virtual (`venv`) isola as bibliotecas deste projeto das do resto do computador.
Todos os comandos abaixo são rodados **a partir da raiz do projeto** (a pasta que contém
`scripts/`, `dados/` e `requirements.txt`).

**Criar** (só uma vez):

```bash
python3 -m venv .venv      # Linux / Mac / WSL
python -m venv .venv       # Windows
```

**Ativar** (toda vez que abrir um terminal novo):

| Sistema | Comando |
|---|---|
| Linux / Mac / **WSL Ubuntu** | `source .venv/bin/activate` |
| Windows — PowerShell | `.venv\Scripts\Activate.ps1` |
| Windows — Prompt de Comando (cmd) | `.venv\Scripts\activate.bat` |

Deu certo quando o prompt do terminal passa a começar com `(.venv)`.

> No PowerShell, se aparecer *"a execução de scripts foi desabilitada neste sistema"*, rode
> `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` nessa mesma janela e ative de novo.

Para sair do ambiente virtual depois: `deactivate`.

---

## 3. Instalação das bibliotecas

Com o `(.venv)` ativado, na raiz do projeto:

```bash
pip install -r requirements.txt
```

Instala `pandas==2.3.3` e `openpyxl==3.1.5` (e as dependências delas). As versões estão
fixadas de propósito, para que uma atualização futura de biblioteca não mude o resultado sem
aviso.

---

## 4. Onde colocar o arquivo de entrada

O arquivo de entrada é a planilha do formulário **exportada do Google Sheets como `.xlsx`**
("Arquivo → Fazer download → Microsoft Excel (.xlsx)"), e o lugar dela é:

```
dados/brutos/
```

O arquivo que está em uso hoje se chama:

```
ALFREDO CHAVES - Diagnóstico Turístico - Formulário.xlsx
```

**O nome do arquivo não é obrigatório** — o script recebe o caminho como argumento, então
qualquer nome serve, desde que você passe esse nome no comando (entre aspas, se tiver espaços
ou acentos). Se a planilha nova vier com outro nome, você tem duas opções:

1. **Renomear o arquivo novo** para o nome acima e sobrescrever o antigo — é o mais simples,
   porque aí o comando da seção 5 continua valendo sem mudar nada (é o que o `README.md` da
   raiz e o `dados/README.md` documentam).
2. **Manter o nome novo** e trocar o caminho dentro do comando da seção 5.

O que **precisa** estar certo, independentemente do nome do arquivo, são os nomes das abas
dentro da planilha (o script confere e avisa se não bater — ver seção 8):

| Aba | Para que serve |
|---|---|
| `ALFREDO CHAVES - Diagnóstico Tu` | Aba de respostas do formulário. O nome vem do título do formulário e é **cortado em 31 caracteres** pelo Google Sheets. |
| `Dados tratados` | Aba analítica, com os 93 empreendimentos já curados à mão. É a base de tudo. |

> Como a planilha chega até aqui (é um export manual, com dois passos humanos) está explicado
> em `dados/README.md`, seção 1.

---

## 5. Como rodar

Da **raiz do projeto**, com o `(.venv)` ativado:

```bash
python scripts/preparar_dados.py "dados/brutos/ALFREDO CHAVES - Diagnóstico Turístico - Formulário.xlsx" dados/tratados
```

São dois argumentos, nessa ordem:

| Posição | Argumento | O que é |
|---|---|---|
| 1º | `"dados/brutos/....xlsx"` | Caminho da planilha de entrada. Entre aspas, porque tem espaços e acentos. |
| 2º | `dados/tratados` | Pasta de saída dos CSVs. É criada se não existir; arquivos com o mesmo nome são sobrescritos. |

Os caminhos são **relativos à pasta de onde você rodou o comando** — por isso rode sempre da
raiz do projeto, não de dentro de `scripts/`.

Rodar sem argumentos (`python scripts/preparar_dados.py`) mostra esse mesmo comando como
lembrete, sem fazer nada.

> **Não use `python -O`.** Esse modo desliga as conferências internas (`assert`) do script e
> transforma erros barulhentos em erros silenciosos. O motivo está em `dados/README.md`, §4.

---

## 6. O que o script gera

18 arquivos em `dados/tratados/`, todos em CSV com separador `;` e UTF-8 com BOM (o formato
que o Excel e o Power BI abrem sem configurar nada).

O prefixo do nome diz o papel da tabela no modelo: `dim_` = dimensão, `ponte_` = associação
muitos-para-muitos, `det_` = detalhe de um empreendimento, `fato_` = evento com medida,
`cfg_` = configuração de interface.

| Arquivo | O que é |
|---|---|
| `dim_empreendimento.csv` | **A tabela central**: um registro por empreendimento (93), com nome, distrito, endereço, horário, contato, acessibilidade, capacidade e os campos derivados usados nos filtros. |
| `dim_regiao.csv` | Os 7 distritos do município, com a quantidade anotada na legenda da pesquisa. |
| `dim_segmento.csv` | Os 7 segmentos de atuação, com o rótulo em linguagem de visitante e a ordem de exibição. |
| `dim_dia.csv` | Os 7 dias da semana, com abreviação e ordem. |
| `dim_grupo.csv` | As 4 faixas de tamanho de grupo que o empreendimento recebe. |
| `dim_conveniencia.csv` | As 7 conveniências do painel de filtros (cartão, estacionamento, pet, etc.). |
| `ponte_empreendimento_segmento.csv` | Quais segmentos cada empreendimento atende (um pode ter vários). |
| `ponte_empreendimento_conveniencia.csv` | Quais conveniências cada empreendimento oferece. |
| `ponte_empreendimento_grupo.csv` | Quais faixas de tamanho de grupo cada empreendimento atende. |
| `ponte_empreendimento_dia.csv` | Em quais dias da semana cada empreendimento abre. |
| `det_empreendimento_atributo.csv` | Atributos específicos de cada segmento (versão analítica, formato atributo/valor). |
| `det_empreendimento_atributo_ficha.csv` | Os mesmos atributos, já filtrados e com o rótulo e a ordem em que aparecem na ficha do visitante. |
| `det_empreendimento_opcao_ficha.csv` | As opções de múltipla escolha que aparecem na ficha (categoria, tipo de diária, etc.). |
| `det_empreendimento_acessibilidade.csv` | Itens de acessibilidade declarados por empreendimento. |
| `det_empreendimento_sanitario.csv` | Itens de sanitário adaptado declarados por empreendimento. |
| `fato_resposta_questionario.csv` | Todas as respostas de múltipla escolha do questionário, uma linha por opção marcada. |
| `cfg_ficha.csv` | Configuração da ficha do visitante: o que aparece em cada segmento, com que rótulo e em que ordem. |
| `dicionario_dados.csv` | Descrição de **cada coluna de cada tabela** acima. Regerado automaticamente junto com o resto. |

---

## 7. Como saber se deu certo

Sem precisar entender o script, três conferências bastam:

**1. O console termina com a lista das 17 tabelas e nenhuma mensagem de erro.** A saída
esperada é exatamente esta (números de 13/09/2026, com os 93 empreendimentos da base atual):

```
dim_empreendimento                93 linhas x 69 colunas
dim_regiao                         7 linhas x  3 colunas
dim_segmento                       7 linhas x  5 colunas
dim_dia                            7 linhas x  3 colunas
dim_grupo                          4 linhas x  4 colunas
dim_conveniencia                   7 linhas x  4 colunas
ponte_empreendimento_segmento    142 linhas x  3 colunas
ponte_empreendimento_conveniencia   329 linhas x  2 colunas
ponte_empreendimento_grupo       207 linhas x  2 colunas
ponte_empreendimento_dia         488 linhas x  4 colunas
det_empreendimento_atributo      647 linhas x  4 colunas
det_empreendimento_atributo_ficha   566 linhas x  6 colunas
det_empreendimento_opcao_ficha   682 linhas x  7 colunas
det_empreendimento_acessibilidade   115 linhas x  3 colunas
det_empreendimento_sanitario      62 linhas x  3 colunas
fato_resposta_questionario      2648 linhas x  6 colunas
cfg_ficha                         38 linhas x  5 colunas
```

A linha que mais importa é a primeira: **`dim_empreendimento` tem que ter 93 linhas**, o total
de empreendimentos da base. Se vier diferente disso, a planilha de entrada mudou — o que pode
ser legítimo (respostas novas foram curadas na aba `Dados tratados`) ou não.

**2. A pasta `dados/tratados/` tem 18 arquivos `.csv`**, com a data/hora de agora:

```bash
ls -l dados/tratados/          # Linux / Mac / WSL
dir dados\tratados\            # Windows
```

**3. Abra `dados/tratados/dim_empreendimento.csv`** — deve ter 94 linhas (1 de cabeçalho + 93
de dados), a primeira coluna `id_empreendimento` indo de `E001` a `E093`, e os acentos
aparecendo corretos (`Crubixá`, `São Bento de Urânia`).

Se o script falhar em algum ponto, ele **para e não grava nada pela metade** nas conferências
de entrada — mas se falhar já no meio do processamento, alguns CSVs podem ter sido gravados
antes do erro. Nesse caso, corrija a causa e rode de novo: os arquivos são sobrescritos.

---

## 8. Erros comuns

### `ERRO: não encontrei o arquivo de entrada: ...`

```
ERRO: não encontrei o arquivo de entrada: dados/brutos/ALFREDO CHAVES - ... .xlsx
  A pasta 'dados/brutos' não existe a partir de onde você rodou o comando.
  O caminho é relativo à pasta atual — rode o comando da raiz do projeto.
```

Duas causas, e a própria mensagem diz qual é:

- **"A pasta ... não existe a partir de onde você rodou o comando"** → você está na pasta
  errada. Volte para a raiz do projeto (`cd ..` se estiver dentro de `scripts/`) e repita.
- **"Planilhas que existem em 'dados/brutos': ..."** → a pasta está certa, mas o nome do
  arquivo não bate. A mensagem lista os arquivos que estão lá; copie o nome exato para o
  comando, entre aspas. (Ver seção 4.)

### `ERRO: a planilha não tem a(s) aba(s) que o script espera: 'Dados tratados'`

A mensagem lista as abas que o arquivo realmente tem. Normalmente é um destes casos:

- A aba foi **renomeada** dentro do Google Sheets → volte o nome original.
- O **formulário foi renomeado** → o nome da aba de respostas muda junto (e é cortado em 31
  caracteres). Ajuste as constantes `ABA_BRUTA` / `ABA_TRATADA`, no início de
  `scripts/preparar_dados.py`, para os nomes que a mensagem listou.
- Foi baixada **só uma aba** em vez da planilha inteira → baixe de novo pelo menu
  "Arquivo → Fazer download → Microsoft Excel (.xlsx)".

### `ERRO: a aba '...' não tem a(s) coluna(s) esperada(s): ...`

O cabeçalho de uma pergunta foi editado no formulário. Restaure a redação original da
pergunta na planilha, ou ajuste o nome dentro do script. Vale o mesmo para o erro parecido
`AssertionError: (34, 'texto encontrado', 'texto esperado')`, que aparece mais adiante na
execução: o primeiro número é o índice da coluna, o segundo texto é o que o script achou e o
terceiro é o que ele esperava.

### `ModuleNotFoundError: No module named 'pandas'` (ou `'openpyxl'`)

A biblioteca não está instalada **nesse** Python. Quase sempre é o ambiente virtual que não
foi ativado (o prompt não mostra `(.venv)`). Ative (seção 2) e rode `pip install -r
requirements.txt` (seção 3). Se o prompt já mostra `(.venv)`, rode o `pip install` de novo —
o ambiente foi criado mas nunca teve as bibliotecas instaladas.

### O script roda até o fim, mas os dados saem errados

Não existe mensagem de erro para esse caso. Ele acontece quando o **texto de uma opção de
resposta** muda no formulário (o cabeçalho da pergunta continua igual): o script não tem como
saber, e a opção desconhecida vira `"Outras respostas"`. O sintoma é uma tabela encolhendo sem
motivo (por exemplo `det_empreendimento_acessibilidade` caindo de 115 para 88 linhas).

Por isso a seção 7 vale a pena: **compare os números do console com os da tabela de
referência.** O diagnóstico completo desse risco, com medições, está em `dados/README.md`, §4.

---

## Documentos relacionados

| Arquivo | O que traz |
|---|---|
| `README.md` (raiz) | Visão geral do projeto, status e o que tem em cada pasta |
| `dados/README.md` | De onde vem a planilha, onde o de-para está no código e o que quebra quando o formulário muda |
| `analise/diagnostico-modelo-dados.md` | O modelo de dados, a convenção de nomes (§7-8) e o diagnóstico do `.pbix` |
| `dados/tratados/dicionario_dados.csv` | Descrição de cada coluna de cada tabela gerada |
