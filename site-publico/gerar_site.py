# -*- coding: utf-8 -*-
"""
Injeta os dados de dados/tratados/ dentro de site-publico/index.html.

O site só exibe: toda a regra de negócio (o que é "acessível", quem "aceita pet",
que faixa de grupo atende quem, o resumo de dias, os textos dos selos) já vem
pronta nas colunas geradas por scripts/preparar_dados.py. O JavaScript do
index.html lê estes CSVs embutidos e monta os cards — não recalcula nada.

Aberto direto do disco (duplo clique) o navegador bloqueia `fetch` de arquivo
local; por isso os CSVs ficam embutidos nas tags <script type="text/plain">.

Este script usa os nomes ATUAIS das tabelas (dim_ / ponte_), definidos em
analise/diagnostico-modelo-dados.md §7-8 — não os antigos fato_*.

Rode sempre que os CSVs forem regerados:
    python site-publico/gerar_site.py
"""
import csv
import io
import re
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
TRATADOS = RAIZ / "dados" / "tratados"
HTML = Path(__file__).resolve().parent / "index.html"

# Colunas de dim_empreendimento que o site exibe ou filtra.
#
# O que fica de fora, de propósito (conteúdo do painel de gestão, acesso
# restrito): formalizacao, n_funcionarios, ano_inicio/faixa_ano_inicio,
# publico_principal, canal de divulgação, necessidades do município e do
# negócio, capacidade_hospedagem_*, preço e a autodeclaração facilidade_pcd /
# capacitado_pcd. Nome do proprietário, CPF/CNPJ, razão social e e-mail não
# existem nem nos CSVs tratados.
COLS_EMP = [
    "id_empreendimento", "nome_empreendimento",
    "distrito", "localizacao_rotulo",
    "endereco_completo", "ponto_referencia",
    "segmentos_rotulo",
    "condicao_via", "via_acesso",
    "dias_resumo", "dias_funcionamento", "abre_feriados",
    "horario_atendimento", "atende_24h",
    "telefones", "whatsapp", "instagram", "site",
    "aceita_cartao", "tem_estacionamento_proprio", "estacionamento",
    "vagas_automoveis", "vagas_onibus",
    "espaco_infantil", "pet_texto", "grupos_capacidade",
    "badge_acessibilidade", "badge_idiomas", "sanitario_adaptado", "idiomas",
    "busca",
]

# Tabelas de apoio, com as colunas que o site usa.
TABELAS = {
    "dim_segmento": ["id_segmento", "rotulo_visitante", "ordem_chip"],
    "dim_regiao": ["id_regiao", "regiao"],
    "dim_dia": ["dia", "dia_abrev", "ordem_dia"],
    "dim_conveniencia": ["id_conveniencia", "conveniencia", "ordem_conveniencia"],
    "ponte_empreendimento_segmento": ["id_empreendimento", "id_segmento"],
    "ponte_empreendimento_dia": ["id_empreendimento", "dia"],
    "ponte_empreendimento_conveniencia": ["id_empreendimento", "id_conveniencia"],
}


def ler(nome):
    caminho = TRATADOS / (nome + ".csv")
    if not caminho.exists():
        raise SystemExit(
            "%s nao encontrado — rode antes:\n"
            "  python scripts/preparar_dados.py \"dados/brutos/ALFREDO CHAVES - "
            "Diagnostico Turistico - Formulario.xlsx\" dados/tratados" % caminho
        )
    with io.open(caminho, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh, delimiter=";"))


def escrever(linhas, colunas):
    saida = io.StringIO()
    w = csv.DictWriter(saida, fieldnames=colunas, delimiter=";",
                       lineterminator="\n", extrasaction="ignore")
    w.writeheader()
    w.writerows(linhas)
    return saida.getvalue().strip()


emp = ler("dim_empreendimento")
faltando = [c for c in COLS_EMP if c not in emp[0]]
if faltando:
    raise SystemExit(
        "dim_empreendimento.csv sem as colunas %s — rode scripts/preparar_dados.py" % faltando
    )

blocos = {"dim_empreendimento": escrever(emp, COLS_EMP)}
for nome, cols in TABELAS.items():
    blocos[nome] = escrever(ler(nome), cols)

html = HTML.read_text(encoding="utf-8")
for nome, csv_txt in blocos.items():
    # o conteudo vai dentro de <script type="text/plain">: so "</script>" precisa de escape
    seguro = csv_txt.replace("</script>", "<\\/script>")
    padrao = re.compile(
        r'(<script type="text/plain" id="csv-%s">).*?(</script>)' % re.escape(nome),
        re.S,
    )
    if not padrao.search(html):
        raise SystemExit("bloco csv-%s nao encontrado no HTML" % nome)
    html = padrao.sub(lambda m: m.group(1) + "\n" + seguro + "\n" + m.group(2), html, count=1)

HTML.write_text(html, encoding="utf-8")
print("dados injetados em %s (%.0f KB)" % (HTML.name, HTML.stat().st_size / 1024))
for nome, txt in blocos.items():
    print("  %-36s %5d linhas" % (nome, txt.count("\n")))
