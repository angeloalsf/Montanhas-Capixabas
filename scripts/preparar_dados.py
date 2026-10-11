#!/usr/bin/env python3
"""
Diagnóstico da oferta turística de Alfredo Chaves — preparação das tabelas para o Power BI.

Lê a planilha original (sem alterá-la) e grava as tabelas tratadas em CSV (UTF-8 com BOM, separador ';').

Convenção de nomes dos arquivos gerados (definida em analise/diagnostico-modelo-dados.md, §7-8):

    dim_    dimensão: chave única, colunas descritivas, fica no lado "1" das relações
    ponte_  só chaves estrangeiras (+ rótulo desnormalizado opcional), resolve um muitos-para-muitos
    det_    detalhe/satélite da entidade central, com payload próprio (valor, opção, item)
    fato_   grão de evento com medida como objeto de análise — só fato_resposta_questionario
    cfg_    configuração de interface, não dimensão de negócio

Este script é o único lugar onde mora a regra de negócio da oferta turística. O buscador
(site-visitante/index.html) e o Power BI apenas LEEM as colunas e tabelas geradas aqui —
nenhum dos dois recalcula "acessível", "aceita pet", faixa de grupo, resumo de dias ou a
composição da ficha. Ao mexer numa regra, mexa aqui e regere os CSVs.

Uso:
    python preparar_dados.py "<planilha.xlsx>" "<pasta_saida>"

Base analítica: aba 'Dados tratados' (93 empreendimentos = total da legenda).
A aba bruta (respostas do formulário) é usada para recuperar a data da resposta, a data de
início do empreendimento, a resposta original de nº de funcionários e — para o buscador do
visitante — endereço, ponto de referência, horário de atendimento e contato comercial.

Continuam fora dos CSVs: nome do proprietário, CPF/CNPJ, razão social e e-mail. O telefone/
WhatsApp e o Instagram/site são o contato comercial que o próprio estabelecimento divulga e
sua publicação foi autorizada (decisão registrada em analise/replanejamento-publico-alvo.md,
seção 5, item 1).
"""
import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------- legenda (imagem)
REGIOES = {1: "Crubixá", 2: "Ibitiruí", 3: "Matilde", 4: "Ribeirão do Cristo",
           5: "Sagrada Família", 6: "São Bento de Urânia", 7: "Sede"}
QTD_LEGENDA = {1: 8, 2: 2, 3: 36, 4: 3, 5: 4, 6: 9, 7: 31}
SIM_NAO = {1: True, 2: False}
LOCALIZACAO = {1: "Urbana", 2: "Rural"}

SEGMENTOS = ["Alimentação", "Hospedagem", "Entretenimento/Aventura", "Agroindústria",
             "Compras", "Cultural", "Recursos Naturais"]
DIAS = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira",
        "Sexta-feira", "Sábado", "Domingo"]

# ---------------------------------------------------------------- interface do visitante
# Rótulos, ordens e listas que antes eram constantes no JavaScript do buscador. Ficam aqui
# para existirem uma vez só: viram colunas e tabelas nos CSVs, e o buscador e o Power BI
# passam a lê-las em vez de reimplementar a mesma regra em JS e em DAX.

# Nome de cada segmento em linguagem de visitante e a ordem dos chips (que não é a ordem
# de id_segmento).
ROTULO_VISITANTE = [
    ("Alimentação", "Onde comer"),
    ("Hospedagem", "Onde dormir"),
    ("Recursos Naturais", "Cachoeiras e natureza"),
    ("Agroindústria", "Produtos locais"),
    ("Compras", "Compras e artesanato"),
    ("Entretenimento/Aventura", "Aventura"),
    ("Cultural", "Cultura"),
]
DIAS_ABREV = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]

# Faixas de tamanho de grupo. grupo_min é comparado com grupos_ordem: cada faixa inclui
# também quem recebe grupos maiores — quem vai com até 10 pessoas cabe em quem recebe 100.
# "Qualquer tamanho" não é uma faixa e não entra na dimensão: é a ausência de filtro.
GRUPOS_FAIXA = [(1, "Até 10 pessoas", 1), (2, "De 10 a 25 pessoas", 2),
                (3, "De 25 a 50 pessoas", 3), (4, "Mais de 50 pessoas", 4)]
GRUPO_ORDEM_TETO = 6        # 9 = "Não se aplica" fica de fora de todas as faixas
GRUPO_ORDEM_SEM_LIMITE = 8  # "Capacidade não definida" entra em todas as faixas

# Conveniências do painel de filtros: cada uma é uma coluna booleana de dim_empreendimento.
CONVENIENCIAS = [("cartao", "Aceita cartão", "aceita_cartao"),
                 ("estac", "Estacionamento próprio", "tem_estacionamento_proprio"),
                 ("pet", "Aceita pet", "aceita_pet"),
                 ("infantil", "Espaço infantil", "espaco_infantil"),
                 ("acess", "Acessível", "acessivel"),
                 ("idioma", "Atende em outro idioma", "tem_idioma"),
                 ("feriado", "Abre em feriados", "abre_feriados")]

# Ficha do visitante: que atributos e listas aparecem em cada segmento, com que rótulo e em
# que ordem. O que não está aqui não aparece na ficha — é o caso dos atributos de gestão
# (certificação/selo de inspeção, interesse em projetos coletivos, responsável pela gestão,
# funciona o ano todo, fluxo em eventos) e das capacidades distorcidas.
FICHA_ATRIBUTOS = {
    "Alimentação": ["Tipo de serviço", "Tipo de atendimento", "Gastronomia", "Espaço infantil"],
    "Hospedagem": ["Aceita animais de estimação", "Oferece almoço e jantar", "Espaço infantil",
                   "Descrição de leitos/quartos (texto livre)"],
    "Agroindústria": ["Produto (texto livre)", "Espaço de degustação", "Turismo de experiência"],
    "Compras": ["Tipo de comércio", "Origem dos produtos", "Embalagem para lembrança"],
    "Entretenimento/Aventura": ["Necessário reservar", "Possui guia/instrutor",
                                "Exige equipamento de proteção", "Restrição de participação",
                                "Capacidade máxima"],
    "Recursos Naturais": ["Acessível a mobilidade reduzida", "Trilhas/caminhos sinalizados",
                          "Normas de conduta/controle", "Em área pública", "Acesso por via pública",
                          "Sazonalidade (texto livre)"],
    "Cultural": ["Forma de entrada", "Necessário reservar", "Guia/monitor/material",
                 "Periodicidade de feiras/festas", "Espaço estruturado para feiras"],
}
ROTULO_ATRIBUTO = {
    "Produto (texto livre)": "Produto",
    "Descrição de leitos/quartos (texto livre)": "Acomodações",
    "Sazonalidade (texto livre)": "Melhor época",
    "Aceita animais de estimação": "Aceita pet",
    "Exige equipamento de proteção": "Exige equipamento de proteção (EPI)",
    "Normas de conduta/controle": "Normas de conduta",
    "Necessário reservar": "Precisa reservar",
    "Guia/monitor/material": "Guia ou material de apoio",
}
# Listas de múltipla escolha que aparecem na ficha, por segmento.
FICHA_LISTAS = {
    "Alimentação": [("Categoria de alimentação", "Categoria")],
    "Hospedagem": [("Tipo de diária", "Diária"), ("Forma de reserva", "Como reservar"),
                   ("Equipamento/serviço de hospedagem", "Equipamentos")],
    "Recursos Naturais": [("Tipo de atrativo natural", "Tipo de atrativo"),
                          ("Estrutura do atrativo natural", "Estrutura no local")],
    "Cultural": [("Tipo de feira", "Tipo de feira")],
}
# "Acomoda" é a única linha da ficha que não vem de det_empreendimento_atributo: é a capacidade
# de hospedagem, e por isso tem ordem 0 (aparece antes dos atributos) e origem própria.
FICHA_CAPACIDADE = ("Hospedagem", "capacidade_hospedagem_pessoas", "Acomoda", 0)
# Valores que a ficha descarta por não dizerem nada ao visitante.
SENTINELAS_FICHA = {"", "Não se aplica", "Não informado", "Outras respostas", "0"}

ABA_BRUTA = "ALFREDO CHAVES - Diagnóstico Tu"
ABA_TRATADA = "Dados tratados"


DESCRICOES = {
    "endereco_completo": "Endereço informado no formulário (recuperado da aba bruta; texto livre, sem CEP nem coordenada)",
    "ponto_referencia": "Ponto de referência informado (recuperado da aba bruta; vazio em 27 casos)",
    "horario_atendimento": "Horário para exibição: faixa padronizada quando o texto tem início e fim, senão o texto original ou 'Consultar'",
    "horario_original": "Resposta original de horário de atendimento (rastreabilidade)",
    "horario_abertura": "Hora de abertura HH:MM, só quando o horário foi padronizado",
    "horario_fechamento": "Hora de fechamento HH:MM, só quando o horário foi padronizado",
    "horario_padronizado": "Se foi possível extrair início e fim do texto livre",
    "telefone": "Primeiro telefone/WhatsApp comercial, em (DD) NNNNN-NNNN — contato divulgado pelo estabelecimento",
    "telefones": "Todos os telefones informados, separados por '; '",
    "whatsapp": "Número em formato internacional (55DDNNNNNNNNN) para link wa.me",
    "instagram": "Perfil do Instagram sem o @ (extraído do texto livre)",
    "site": "Site do empreendimento, quando informado",
    "id_regiao": "Código do distrito conforme a legenda (1 a 7)",
    "regiao": "Nome do distrito (legenda)",
    "qtd_empreendimentos_legenda": "Quantidade de estabelecimentos anotada na legenda — usada para conferência",
    "id_segmento": "Código do segmento (1 a 7)",
    "segmento": "Segmento de atuação (um empreendimento pode ter vários)",
    "incluido_graficos_gerais_pdf": "Se o segmento aparece no Gráfico 1/Tabela 1 do PDF (Cultural e Recursos Naturais não aparecem)",
    "id_empreendimento": "Chave do empreendimento (E001…E093), gerada na ordem da aba 'Dados tratados'",
    "nome_empreendimento": "Nome fantasia informado no formulário",
    "localizacao": "Urbana/Rural (legenda: 1 = Urbana, 2 = Rural)",
    "segmentos": "Lista dos segmentos do empreendimento (texto, para tooltip)",
    "qtd_segmentos": "Número de segmentos em que o empreendimento atua",
    "data_resposta": "Data da resposta ao formulário mantida (a mais recente, quando houve duplicidade)",
    "ano_inicio": "Ano de início do empreendimento (recuperado da aba bruta; vazio = não informado)",
    "faixa_ano_inicio": "Faixa do ano de início, para gráfico de abertura de empreendimentos",
    "faixa_ano_inicio_ordem": "Ordem da faixa de ano (usar em 'Classificar por coluna')",
    "condicao_via_ordem": "Ordem da condição da via (usar em 'Classificar por coluna')",
    "placa_sinalizacao_ordem": "Ordem da categoria de sinalização (usar em 'Classificar por coluna')",
    "formalizacao": "Tipo de documento: CNPJ, CPF, CPF + Inscrição Estadual ou Pendente/não informado (sem o número)",
    "via_acesso_original": "Resposta original sobre a via de acesso (rastreabilidade)",
    "via_acesso": "Via de acesso padronizada (respostas livres reclassificadas)",
    "condicao_via": "Pavimentada (asfalto/calçamento), Não pavimentada (terra/revsol) ou Mista",
    "placa_sinalizacao_detalhe": "Resposta original sobre placa de sinalização",
    "placa_sinalizacao": "Sem placa / Placa própria (sem padrão) / Placa padronizada (municipal/EMBRATUR)",
    "faixas_chao": "Existem faixas/pinturas no chão orientando até o atrativo",
    "sinalizacao_contribui": "Percepção se a sinalização contribui para segurança/orientação",
    "n_funcionarios": "Nº de funcionários conforme a aba tratada (ver funcionarios_resposta_textual)",
    "funcionarios_resposta_textual": "True quando a resposta original era texto livre e foi convertida em número",
    "tem_email": "Informou e-mail válido (o e-mail não é exportado)",
    "tem_site_ou_instagram": "Informou site e/ou Instagram (o endereço não é exportado)",
    "dias_funcionamento": "Dias de funcionamento (texto padronizado da aba tratada)",
    "qtd_dias_semana": "Quantidade de dias de funcionamento por semana (1 a 7)",
    "abre_seg_a_qui": "Abre de segunda a quinta (todos os quatro dias)",
    "abre_fim_de_semana": "Abre no sábado e/ou domingo",
    "atende_24h": "Atendimento 24 horas",
    "abre_feriados": "Funciona em feriados",
    "grupos_capacidade": "Faixa de tamanho de grupo que o empreendimento recebe",
    "grupos_ordem": "Ordem da faixa de grupos (usar em 'Classificar por coluna')",
    "publico_principal": "Público que mais frequenta o local",
    "estacionamento": "Modalidade de estacionamento",
    "tem_estacionamento_proprio": "Estacionamento próprio (gratuito ou pago)",
    "vagas_automoveis": "Vagas para automóveis (só para quem tem estacionamento próprio, como no PDF)",
    "vagas_onibus": "Vagas para ônibus (só para quem tem estacionamento próprio, como no PDF)",
    "lingua_estrangeira": "Atende em pelo menos uma língua estrangeira",
    "idiomas": "Idiomas atendidos, separados por ';'",
    "aceita_cartao": "Aceita cartão de crédito e/ou débito",
    "facilidade_pcd": "Possui facilidade para PcD/mobilidade reduzida (legenda: 1 = Sim, 2 = Não)",
    "capacitado_pcd": "Declara estar capacitado para receber ao menos um tipo de deficiência",
    "sanitario_adaptado": "Possui ao menos um item de sanitário adaptado",
    "capacidade_hospedagem_pessoas": "Capacidade máxima declarada da hospedagem (pessoas); só segmento Hospedagem",
    "capacidade_hospedagem_origem": "Como o número foi obtido: numérico, extraído de texto, ambíguo, não informado",
    "uh_adaptadas_pcd": "Unidades habitacionais adaptadas para PcD (vazio = não informou quantidade)",
    "capacidade_alimentacao_pessoas": "Capacidade máxima declarada (pessoas); só segmento Alimentação",
    "na_base_pdf": "False para o único empreendimento fora dos 5 segmentos do Gráfico 1 (reproduz a base n = 92 do PDF)",
    "pendencias_qualidade": "Problemas de qualidade encontrados no registro (texto)",
    # colunas derivadas que antes só existiam no JavaScript do buscador
    "distrito": "Nome do distrito já resolvido a partir de id_regiao (mesmo valor de dim_regiao.regiao)",
    "acessivel": "Sanitário adaptado OU algum recurso de acessibilidade cujo nome começa com 'rampa' — é o filtro 'Acessível' do buscador",
    "aceita_pet": "Hospedagem que respondeu 'Sim' ou 'Depende do porte do animal' em 'Aceita animais de estimação'",
    "pet_texto": "Resposta original de 'Aceita animais de estimação' (o card distingue 'Depende do porte do animal')",
    "espaco_infantil": "Algum dos segmentos Alimentação ou Hospedagem respondeu 'Sim' em 'Espaço infantil'",
    "tem_idioma": "Atende em pelo menos um idioma estrangeiro (idiomas não vazio)",
    "dias_resumo": "Dias como aparecem no card: 'Todos os dias', ou as abreviações na ordem da semana; sem dias marcados, repete dias_funcionamento",
    "busca": "Nome + endereço + ponto de referência + distrito + segmentos, em minúsculas e sem acento, para busca textual",
    "segmentos_rotulo": "Segmentos em linguagem de visitante, na ordem dos chips, separados por ' · ' — é a linha de segmento do card",
    "localizacao_rotulo": "Localização como o card exibe: 'zona urbana' ou 'zona rural'",
    "badge_idiomas": "Selo de idioma pronto para o card: 'Atende em inglês, francês'; vazio para quem não declarou idioma",
    "badge_acessibilidade": "Rótulo único de acessibilidade do card: 'Sanitário adaptado' tem precedência sobre 'Rampa de acesso'; vazio quando não há nenhum dos dois",
    # dimensões da interface
    "rotulo_visitante": "Nome do segmento em linguagem de visitante, usado nos chips do buscador",
    "ordem_chip": "Ordem de exibição do segmento nos chips (não é a ordem de id_segmento)",
    "dia": "Dia da semana",
    "dia_abrev": "Abreviação do dia (Seg, Ter, Qua…)",
    "ordem_dia": "Ordem do dia na semana (1 = segunda-feira)",
    "id_grupo": "Código da faixa de tamanho de grupo",
    "rotulo_grupo": "Faixa de tamanho de grupo como aparece no filtro",
    "grupo_min": "Menor grupos_ordem que atende a faixa — cada faixa inclui quem recebe grupos maiores",
    "ordem_grupo": "Ordem de exibição da faixa de grupo",
    "id_conveniencia": "Código da conveniência no filtro do buscador",
    "conveniencia": "Conveniência como aparece no filtro",
    "coluna_origem": "Coluna booleana de dim_empreendimento que define a conveniência",
    "ordem_conveniencia": "Ordem de exibição da conveniência",
    "tipo": "O que a linha configura: 'atributo' (det_empreendimento_atributo), 'lista' (fato_resposta_questionario) ou 'capacidade' (coluna de dim_empreendimento)",
    "chave": "Nome do atributo, da pergunta ou da coluna de origem",
    "rotulo": "Rótulo exibido na ficha",
    "ordem": "Ordem de exibição dentro do bloco do segmento (0 = 'Acomoda')",
    # tabelas de exibição da ficha
    "rotulo_atributo": "Rótulo do atributo como aparece na ficha",
    "ordem_atributo": "Ordem do atributo no bloco do segmento (0 = 'Acomoda')",
    "rotulo_lista": "Rótulo da lista como aparece na ficha",
    "ordem_lista": "Ordem da lista no bloco do segmento",
    "ordem_opcao_ficha": "Ordem da opção dentro da lista (a do questionário, não alfabética)",
    "item": "Item como exibido na ficha",
    "ordem_item": "Ordem do item na lista da ficha",
    "tema": "Agrupador temático da pergunta",
    "pergunta": "Pergunta de múltipla escolha (versão curta)",
    "opcao": "Opção marcada (respostas livres agrupadas em 'Outras respostas')",
    "detalhe_outras": "Texto livre quando a opção é 'Outras respostas'",
    "ordem_opcao": "Ordem de exibição (dias da semana 1–7; 'Outras respostas' por último)",
    "atributo": "Atributo específico do segmento",
    "valor": "Valor do atributo (texto)",
}


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


COMBINANTES = dict.fromkeys(range(0x300, 0x370))


def sem_acento(s):
    """Minúsculas sem acento, preservando pontuação — igual ao semAcento() do buscador."""
    return unicodedata.normalize("NFD", str(s)).translate(COMBINANTES).lower()


def maiuscula(s):
    """Primeira letra maiúscula, como a ficha exibe os itens de lista."""
    return s[:1].upper() + s[1:] if s else s


def clean(v):
    if pd.isna(v):
        return None
    v = re.sub(r"\s+", " ", str(v)).strip()
    return v or None


def tokens(v):
    v = clean(v)
    if v is None:
        return []
    return [t for t in (clean(x) for x in v.split(";")) if t]


# ---------------------------------------------------------------- contato e horário (público)
DIAS_TEXTO = re.compile(r"segunda|terca|terça|quarta|quinta|sexta|sabado|sábado|domingo|feriado", re.I)
HORA = re.compile(r"(?<![\d:])(\d{1,2})(?:\s*[:.h]\s*(\d{2}))?\s*(?:hrs|hs|horas|hora|h)?(?![\d])", re.I)
CONECTOR = re.compile(r"[\s,h]*(?:hs|hrs|horas|hora)?[\s,]*(?:as|às|ás|a|ate|até|-|–|—)[\s,]*(?:das|de)?[\s,]*", re.I)
VAGOS = {"a ser definido", "comercial", "dia todo", "diariamente",
         "n a atendimento por whatsapp", "nos horarios das acoes da associacao"}


def horas_no_texto(txt):
    """Horas legíveis do texto: lista de (minutos desde 00h00, início, fim) no texto."""
    achados = []
    for m in HORA.finditer(txt):
        h = int(m.group(1))
        mi = int(m.group(2) or 0)
        if h > 24 or mi > 59:
            continue
        # número solto sem 'h' nem ':' só conta se estiver cercado por palavras de faixa
        if m.group(2) is None and not re.search(r"h|:", m.group(0), re.I):
            ctx = txt[max(0, m.start() - 12):m.end() + 12].lower()
            if not re.search(r"\bas\b|\bàs\b|\bás\b|\ba\b|\bde\b|\bdas\b|-", ctx):
                continue
        achados.append((h % 24 * 60 + mi, m.start(), m.end()))
    return achados


def hhmm(minutos):
    return "%02dh%02d" % (minutos // 60, minutos % 60)


IRRELEVANTES = {"de", "da", "das", "do", "dos", "a", "as", "e", "em", "ate", "no", "nos", "na", "nas",
                "todos", "os", "dia", "dias", "hs", "hrs", "h", "horas", "hora", "aberto", "atendimento",
                "segunda", "terca", "quarta", "quinta", "sexta", "sabado", "sab", "dom", "domingo",
                "feira", "feiras", "seg", "ter", "qua", "qui", "sex", "feriado", "feriados"}


def palavras_uteis(trecho):
    """Palavras do trecho que não são dia da semana nem conector de faixa."""
    return [w for w in norm(trecho).split() if w and w not in IRRELEVANTES and not w.isdigit()]


def padronizar_horario(txt):
    """(texto_exibicao, abertura, fechamento, padronizado?).

    Só padroniza quando o texto descreve uma faixa única: exatamente duas horas
    legíveis ligadas apenas por um conector de intervalo ("das 9 às 17"). Se entre
    as duas horas houver qualquer outra coisa — outro dia da semana, check-out,
    um segundo turno — o texto original é mais informativo e é mantido como está.
    """
    if not txt:
        return None, None, None, False
    t = re.sub(r"[^\S\n]+", " ", txt).strip().lstrip("⁰¹²³⁴⁵⁶⁷⁸⁹")
    if norm(t) in VAGOS:
        return "Consultar", None, None, False
    if re.fullmatch(r"24\s*h(oras)?", t, re.I):
        return "24 horas", 0, 24 * 60, True
    horas = horas_no_texto(t)
    if len(horas) == 2 and horas[0][0] != horas[1][0] and "check" not in t.lower():
        meio = t[horas[0][2]:horas[1][1]]
        sobra = palavras_uteis(t[:horas[0][1]]) + palavras_uteis(t[horas[1][2]:])
        # sobra = o que o texto diz além da faixa e dos dias ("em época de safra",
        # "domingo com hora marcada"): se é muita coisa, o original informa mais
        if CONECTOR.fullmatch(meio) and len(sobra) < 3:
            ini, fim = horas[0][0], horas[1][0]
            return "%s às %s" % (hhmm(ini), hhmm(fim)), ini, fim, True
    return t, None, None, False


def formatar_telefones(v):
    """Números de telefone/WhatsApp do texto livre, em (DD) NNNNN-NNNN."""
    if not v:
        return [], []
    achados, e164 = [], []
    for bruto_num in re.findall(r"\d[\d\s\-\.\(\)]{8,}\d", str(v)):
        d = re.sub(r"\D", "", bruto_num)
        if d.startswith("55") and len(d) in (12, 13):
            d = d[2:]
        if len(d) not in (10, 11):
            continue
        dd, resto = d[:2], d[2:]
        fmt = "(%s) %s-%s" % (dd, resto[:-4], resto[-4:])
        if fmt not in achados:
            achados.append(fmt)
            e164.append("55" + d)
    return achados, e164


def extrair_instagram_site(v):
    """(instagram_sem_arroba, url_do_site) a partir do texto livre."""
    if not v:
        return None, None
    t = str(v).strip()
    if norm(t) in ("nao", "nao tem", "nao possui", "n a", "sim intagram", "instagram", "sim instagram"):
        return None, None
    insta = site = None
    m = re.search(r"instagram\.com/([A-Za-z0-9_.]+)", t, re.I)
    if m:
        insta = m.group(1)
    else:
        m = re.search(r"@([A-Za-z0-9_.]+)", t)
        if m:
            insta = m.group(1)
    for url in re.findall(r"(?:https?://)?(?:www\.)?[A-Za-z0-9][A-Za-z0-9\-.]*\.(?:com|br|net|org|esp)(?:\.br)?(?:/[^\s]*)?", t):
        if re.search(r"instagram\.com", url, re.I):
            continue
        site = url if url.lower().startswith("http") else "https://" + url
        break
    if insta is None and site is None:
        # resposta só com o @ implícito: uma palavra sem espaço vira handle
        cand = t.strip().strip(".,/ ")
        if re.fullmatch(r"[A-Za-z0-9_.]{3,40}", cand) and not cand.isdigit():
            insta = cand
    return (insta.strip(" ./") if insta else None), site


def tipo_amigavel(nome, serie):
    """Tipo para o dicionário, independente da versão do pandas."""
    if nome.startswith("data_"):
        return "Data"
    if pd.api.types.is_bool_dtype(serie):
        return "Verdadeiro/Falso"
    if pd.api.types.is_integer_dtype(serie):
        return "Número inteiro"
    return "Texto"


def primeiro_inteiro(v):
    """Extrai o primeiro número de um texto livre de capacidade. Retorna (valor, origem)."""
    if pd.isna(v):
        return None, "não informado"
    if isinstance(v, (int, float)):
        return int(v), "numérico"
    t = clean(v)
    if norm(t) in ("nao", "apenas para hospedes"):
        return None, "declarou não ter capacidade máxima"
    if norm(t).startswith("cada"):
        return None, "ambíguo (capacidade por unidade)"
    m = re.search(r"\d+", t)
    return (int(m.group()), "extraído de texto") if m else (None, "não informado")


def parse_dias(txt):
    """Converte a resposta padronizada ('Sexta-feira a Domingo', 'Todos os dias...') em lista de dias."""
    t = clean(txt)
    if t is None:
        return []
    low = t.lower()
    if low.startswith("todos os dias"):
        return DIAS[:]
    idx = {d.lower(): i for i, d in enumerate(DIAS)}
    if ";" in t:
        return [DIAS[idx[x.lower()]] for x in tokens(t)]
    if " a " in low:
        ini, fim = [x.strip() for x in low.split(" a ")]
        i, j = idx[ini], idx[fim]
        out, k = [], i
        while True:
            out.append(DIAS[k])
            if k == j:
                break
            k = (k + 1) % 7
        return out
    if " e " in low:
        return [DIAS[idx[x.strip()]] for x in low.split(" e ")]
    return [DIAS[idx[low]]]


def main(xlsx, out_dir):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    raw = pd.read_excel(xlsx, sheet_name=ABA_BRUTA)
    dt = pd.read_excel(xlsx, sheet_name=ABA_TRATADA)
    C = list(dt.columns)

    def col(i, prefix):
        assert str(C[i]).strip().startswith(prefix.strip()), (i, C[i], prefix)
        return C[i]

    # ---- ligação com a aba bruta: resposta mais recente por nome (mesma regra usada na aba tratada)
    NOME = "Nome do empreendimento?"
    raw["k"] = raw[NOME].map(norm)
    raw["ts"] = pd.to_datetime(raw["Carimbo de data/hora"].str.replace(" GMT-3", ""),
                               format="%Y/%m/%d %I:%M:%S %p")
    latest = raw.sort_values("ts").drop_duplicates("k", keep="last").set_index("k")
    dt["k"] = dt[NOME].map(norm)
    assert dt["k"].isin(latest.index).all()
    lk = latest.loc[dt["k"]].reset_index(drop=True)

    dt = dt.reset_index(drop=True)
    dt["id_empreendimento"] = [f"E{i + 1:03d}" for i in range(len(dt))]
    ID = dt["id_empreendimento"]
    pend = {i: [] for i in ID}

    # ---- checagem da legenda contra o texto original
    assert (dt[col(10, "Localização")].map(LOCALIZACAO).values == lk["Localização"].str.strip().values).all()
    fac_raw = lk["Possui alguma facilidade para pessoas com deficiência ou mobilidade reduzida?"].str.strip()
    assert (dt[col(32, "Possui alguma facilidade")].map({1: "Sim", 2: "Não"}).values == fac_raw.values).all()

    # =========================================================== dim_regiao
    dim_regiao = pd.DataFrame({
        "id_regiao": list(REGIOES),
        "regiao": [REGIOES[k] for k in REGIOES],
        "qtd_empreendimentos_legenda": [QTD_LEGENDA[k] for k in REGIOES],
    })

    # =========================================================== dim_segmento + ponte
    rot_vis = dict(ROTULO_VISITANTE)
    ordem_chip = {seg: i + 1 for i, (seg, _) in enumerate(ROTULO_VISITANTE)}
    assert set(rot_vis) == set(SEGMENTOS)
    dim_segmento = pd.DataFrame({
        "id_segmento": range(1, 8),
        "segmento": SEGMENTOS,
        "rotulo_visitante": [rot_vis[x] for x in SEGMENTOS],
        "ordem_chip": [ordem_chip[x] for x in SEGMENTOS],
        "incluido_graficos_gerais_pdf": [True] * 5 + [False, False],
    })
    seg_cols = C[0:7]
    ponte = []
    for sc, sname in zip(seg_cols, SEGMENTOS):
        assert str(sc).strip() == sname
        for i, v in zip(ID, dt[sc]):
            if SIM_NAO[int(v)]:
                ponte.append((i, SEGMENTOS.index(sname) + 1, sname))
    ponte = pd.DataFrame(ponte, columns=["id_empreendimento", "id_segmento", "segmento"])

    # =========================================================== empreendimentos
    E = pd.DataFrame({"id_empreendimento": ID})
    E["nome_empreendimento"] = dt[NOME].map(clean)
    E["id_regiao"] = dt[col(7, "Região")].astype(int)
    E["localizacao"] = dt[col(10, "Localização")].map(LOCALIZACAO)
    # endereço e ponto de referência: existem só na aba bruta e são informação pública de
    # localização do estabelecimento (não do proprietário) — necessários ao buscador do visitante
    E["endereco_completo"] = lk["Endereço completo: (Logradouro, nº, bairro/localidade/distrito)"].map(clean).values
    E["ponto_referencia"] = lk["Ponto de referência:"].map(clean).values
    segs = ponte.groupby("id_empreendimento")["segmento"].apply(lambda s: "; ".join(s))
    E["segmentos"] = E["id_empreendimento"].map(segs)
    E["qtd_segmentos"] = E["id_empreendimento"].map(ponte.groupby("id_empreendimento").size()).astype(int)
    E["data_resposta"] = lk["ts"].dt.date.values
    inicio = pd.to_datetime(lk["Data do Início do empreendimento"], errors="coerce")
    E["ano_inicio"] = inicio.dt.year.astype("Int64").values
    E["faixa_ano_inicio"] = pd.cut(E["ano_inicio"].astype("float"), [0, 1999, 2009, 2014, 2019, 2100],
                                   labels=["Até 1999", "2000–2009", "2010–2014", "2015–2019", "2020 em diante"]
                                   ).cat.add_categories("Não informado").fillna("Não informado").astype(str)
    E["faixa_ano_inicio_ordem"] = E["faixa_ano_inicio"].map(
        {"Até 1999": 1, "2000–2009": 2, "2010–2014": 3, "2015–2019": 4, "2020 em diante": 5, "Não informado": 9})

    # formalização (a aba tratada já traz só o tipo de documento, exceto 1 caso com número)
    doc = dt[col(9, "CPF/CNPJ")].map(clean)
    docmap = {"CNPJ": "CNPJ", "CPF": "CPF", "CPF e IE": "CPF + Inscrição Estadual"}
    E["formalizacao"] = doc.map(lambda v: docmap.get(v, "Pendente/não informado"))
    for i, v in zip(ID, doc):
        if v not in docmap:
            pend[i].append("documento não informado/pendente")

    # acesso e sinalização
    via = dt[col(11, "A via que dá acesso")].map(clean)
    viamap = {"Asfaltada": "Asfaltada", "Terra": "Terra", "Revsol/ Brita": "Revsol/Brita",
              "Calçada": "Calçada/calçamento", "Paralelepípedo": "Calçada/calçamento",
              "1 km estrada de chão": "Terra",
              "Asfaltada, calçamento, revsol e terra": "Mista",
              "Só em frente é do empreendimento é calçada.": "Mista"}
    assert set(via) <= set(viamap), set(via) - set(viamap)
    E["via_acesso_original"] = via
    E["via_acesso"] = via.map(viamap)
    E["condicao_via"] = E["via_acesso"].map({"Asfaltada": "Pavimentada", "Calçada/calçamento": "Pavimentada",
                                             "Terra": "Não pavimentada", "Revsol/Brita": "Não pavimentada",
                                             "Mista": "Mista"})
    E["condicao_via_ordem"] = E["condicao_via"].map({"Pavimentada": 1, "Mista": 2, "Não pavimentada": 3})
    placa = dt[col(12, "O empreendimento possui algum tipo de placa")].map(clean)
    E["placa_sinalizacao_detalhe"] = placa
    E["placa_sinalizacao"] = placa.map(lambda v: "Sem placa" if v == "Não possui"
                                       else "Placa própria (sem padrão)" if v == "Possui, sem padronização (modelo próprio)"
                                       else "Placa padronizada (municipal/EMBRATUR)")
    E["placa_sinalizacao_ordem"] = E["placa_sinalizacao"].map(
        {"Placa padronizada (municipal/EMBRATUR)": 1, "Placa própria (sem padrão)": 2, "Sem placa": 3})
    E["faixas_chao"] = dt[col(13, "Existem faixas")].map(clean).replace({"Sem atrativos": "Não se aplica"})
    E["sinalizacao_contribui"] = dt[col(14, "A sinalização contribui")].map(clean).replace(
        {"Precisa melhorar": "Precisa melhorar", "Mais ainda não tem": "Não", "Não tem": "Não"})

    # funcionários
    E["n_funcionarios"] = dt[col(15, "Quantidade de funcionários")].astype(int)
    f_raw = lk["Quantidade de funcionários? \n(Responda com o número. Exemplo: 02)"]
    f_txt = f_raw.map(lambda v: not isinstance(v, (int, float)) and not str(v).strip().isdigit()).values
    E["funcionarios_resposta_textual"] = f_txt
    for i, t in zip(ID, f_txt):
        if t:
            pend[i].append("nº de funcionários convertido de texto livre")

    # presença digital e contato comercial.
    # O telefone/WhatsApp e o Instagram/site são o contato divulgado pelo próprio estabelecimento
    # (fachada, redes) e sua publicação foi autorizada para o buscador do visitante.
    # Continuam FORA do CSV: nome do proprietário, CPF/CNPJ, razão social e e-mail.
    email = dt[col(17, "E-mail")].map(clean)
    E["tem_email"] = email.map(lambda v: isinstance(v, str) and "@" in v)
    site = dt[col(18, "O empreendimento possui site")].map(clean)
    E["tem_site_ou_instagram"] = site.map(lambda v: isinstance(v, str) and norm(v) not in ("nao", "nao tem", "nao possui"))

    tel_raw = lk["Contato telefônico e WhatsApp \n(Responda conforme exemplo: (28) 99999-9999)"].map(clean).values
    tels = [formatar_telefones(v) for v in tel_raw]
    E["telefone"] = [t[0][0] if t[0] else None for t in tels]
    E["telefones"] = ["; ".join(t[0]) or None for t in tels]
    E["whatsapp"] = [t[1][0] if t[1] else None for t in tels]
    for i, (fmt, _), v in zip(ID, tels, tel_raw):
        if v and not fmt:
            pend[i].append("telefone em formato não reconhecido")

    insta_site = [extrair_instagram_site(v) for v in site]
    E["instagram"] = [a for a, _ in insta_site]
    E["site"] = [b for _, b in insta_site]

    # funcionamento
    dias_txt = dt[col(22, "Dias de funcionamento")].map(clean)
    dias_lst = dias_txt.map(parse_dias)
    E["dias_funcionamento"] = dias_txt
    E["qtd_dias_semana"] = dias_lst.map(len)
    E["abre_seg_a_qui"] = dias_lst.map(lambda l: all(d in l for d in DIAS[:4]))
    E["abre_fim_de_semana"] = dias_lst.map(lambda l: ("Sábado" in l) or ("Domingo" in l))
    raw_dias = lk["Dias de funcionamento:"].map(lambda v: sorted(tokens(v)))
    assert all(sorted(a) == b for a, b in zip(dias_lst, raw_dias)), "dias divergentes entre abas"
    # horário de atendimento: recuperado da aba bruta (a tratada descartou a coluna).
    # Padronizado só quando o texto descreve uma faixa única; os demais mantêm o texto original,
    # que é mais informativo ("Seg a sábado: 08 as 17 / domingo: 08 as 11h").
    hor_raw = lk["Horário de atendimento:\n(Exemplo: Sexta a domingo, das XXh às YYh)"].map(clean).values
    hor = [padronizar_horario(v) for v in hor_raw]
    E["horario_atendimento"] = [h[0] for h in hor]
    E["horario_original"] = hor_raw
    E["horario_abertura"] = ["%02d:%02d" % (h[1] // 60, h[1] % 60) if h[1] is not None else None for h in hor]
    E["horario_fechamento"] = ["%02d:%02d" % (h[2] // 60, h[2] % 60) if h[2] is not None else None for h in hor]
    E["horario_padronizado"] = [h[3] for h in hor]
    E["atende_24h"] = dt[col(23, "Atendimento 24 horas")].map({"Sim": True, "Não": False})
    E["abre_feriados"] = dt[col(24, "Funcionamento em feriados")].map({"Sim": True, "Não": False})
    grupos = dt[col(25, "Seu estabelecimento recebe grupos")].map(clean)
    gord = {"Não recebe grupos": 0, "Até 10 visitantes": 1, "De 10 a 25 visitantes": 2,
            "De 25 a 50 visitantes": 3, "De 50 a 75 visitantes": 4, "De 75 a 100 visitantes": 5,
            "Acima de 100 visitantes": 6, "Capacidade não definida": 8, "Não se aplica": 9}
    assert set(grupos) <= set(gord)
    E["grupos_capacidade"] = grupos
    E["grupos_ordem"] = grupos.map(gord)
    E["publico_principal"] = dt[col(26, "Qual é o público")].map(clean)

    # estacionamento
    est = dt[col(27, "Estacionamento")].map(clean)
    E["estacionamento"] = est
    E["tem_estacionamento_proprio"] = est.str.startswith(("Gratuito", "Pago"))
    va = dt[col(28, "Quantas vagas de estacionamento tem para automóveis")]
    vo = dt[col(29, "Quantas vagas de estacionamento tem para ônibus")]
    E["vagas_automoveis"] = va.where(E["tem_estacionamento_proprio"]).astype("Int64")
    E["vagas_onibus"] = vo.where(E["tem_estacionamento_proprio"]).astype("Int64")
    for i, p, a, o in zip(ID, E["tem_estacionamento_proprio"], va, vo):
        if not p and (pd.notna(a) or pd.notna(o)):
            pend[i].append("informou vagas sem estacionamento próprio (vagas desconsideradas)")
        if p and pd.isna(a):
            pend[i].append("estacionamento próprio sem nº de vagas de automóveis")

    # atendimento
    ling = dt[col(30, "Atendimento em língua estrangeira")].map(clean)
    E["lingua_estrangeira"] = ling != "Não"
    E["idiomas"] = ling.map(lambda v: None if v == "Não" else v.replace(" e ", "; "))
    pag = dt[col(31, "Tipo de pagamento")].map(tokens)
    E["aceita_cartao"] = pag.map(lambda l: any(t.startswith("Cartão") for t in l))
    E["facilidade_pcd"] = dt[col(32, "Possui alguma facilidade")].map(SIM_NAO)
    E["capacitado_pcd"] = dt[col(33, "Local capacitado")].map(lambda v: clean(v) != "Não")
    san = dt[col(35, "Sanitário adaptado")].map(tokens)
    E["sanitario_adaptado"] = san.map(lambda l: any(t != "Não" for t in l))
    for i, l in zip(ID, san):
        if "Não" in l and len(l) > 1:
            pend[i].append("sanitário: marcou 'Não' e itens de adaptação ao mesmo tempo")

    # capacidades
    cap_h = [primeiro_inteiro(v) for v in dt[col(95, "Seu estabelecimento tem capacidade máxima")]]
    is_h = dt[seg_cols[1]] == 1
    E["capacidade_hospedagem_pessoas"] = pd.array([c[0] if h else None for c, h in zip(cap_h, is_h)], dtype="Int64")
    E["capacidade_hospedagem_origem"] = [c[1] if h else None for c, h in zip(cap_h, is_h)]
    uh = dt[col(93, "Unidades habitacionais adaptadas")]
    E["uh_adaptadas_pcd"] = pd.array([int(v) if isinstance(v, (int, float)) and pd.notna(v) else None for v in uh], dtype="Int64")
    cap_a = [primeiro_inteiro(v) for v in dt[col(41, "Seu estabelecimento tem capacidade máxima")]]
    is_a = dt[seg_cols[0]] == 1
    E["capacidade_alimentacao_pessoas"] = pd.array([c[0] if a else None for c, a in zip(cap_a, is_a)], dtype="Int64")
    for i, h, c, u in zip(ID, is_h, cap_h, uh):
        if h and c[1] in ("extraído de texto", "ambíguo (capacidade por unidade)", "não informado", "declarou não ter capacidade máxima"):
            pend[i].append(f"capacidade de hospedagem: {c[1]}")
        if h and isinstance(u, str):
            pend[i].append("UH adaptadas: respondeu 'Sim' sem quantidade")
        if h and pd.isna(u):
            pend[i].append("UH adaptadas: sem resposta")
    E["na_base_pdf"] = (dt[seg_cols[:5]] == 1).any(axis=1)
    E["pendencias_qualidade"] = E["id_empreendimento"].map(lambda i: "; ".join(pend[i]) or None)

    # =========================================================== respostas múltiplas (formato longo)
    R = []

    def add(pergunta, tema, series, padrao=None, renomear=None, outros="Outras respostas", filtro=None):
        renomear = renomear or {}
        for i, v in zip(ID, series):
            for t in (v if isinstance(v, list) else tokens(v)):
                t = renomear.get(t, t)
                if padrao is not None and t not in padrao:
                    R.append((i, tema, pergunta, outros, t))
                else:
                    R.append((i, tema, pergunta, t, None))

    add("Canal de divulgação", "Mercado e divulgação", dt[col(19, "Como é feita a divulgação")])
    add("Necessidade turística do município", "Demandas do trade", dt[col(20, "Na sua opinião")],
        padrao={"Melhoria das estradas e acessos aos atrativos turísticos",
                "Divulgação e promoção dos atrativos turísticos",
                "Melhorias na sinalização vertical (placas)",
                "Promover cursos e capacitação para profissionais do setor",
                "Melhorias na sinalização horizontal"})
    add("Necessidade do empreendimento", "Demandas do trade", dt[col(21, "Sobre seu empreendimento")],
        padrao={"Marketing e redes sociais", "Mão de obra", "Vendas e estratégias comerciais",
                "Capacitação de mão de obra", "Gestão administrativa",
                "Aspectos jurídicos/regularização", "Gestão financeira"})
    add("Dia de funcionamento", "Funcionamento", dias_lst)
    add("Forma de pagamento", "Qualificação do atendimento", pag)
    add("Idioma atendido", "Qualificação do atendimento", E["idiomas"])
    add("Deficiência atendida", "Acessibilidade", dt[col(33, "Local capacitado")],
        renomear={"Não": "Não capacitado", "Fisica": "Física"})
    add("Recurso de acessibilidade", "Acessibilidade", dt[col(34, "Quais acessibilidades")],
        padrao={"Não possui", "Estacionamento", "Rampa", "Rampa de acesso à calçada", "Corrimão",
                "Elevador", "Interprete de libras", "Plataforma elevativa"})
    add("Item de sanitário adaptado", "Acessibilidade",
        [[t for t in l if not (t == "Não" and len(l) > 1)] for l in san],
        renomear={"Não": "Não possui sanitário adaptado"})
    add("Categoria de alimentação", "Segmento: Alimentação", dt[col(38, "Onde seu empreendimento")],
        padrao={"Restaurante", "Lanchonete", "Cafeteria", "Cervejaria", "Confeitaria", "Padaria", "Pizzaria"})
    add("Tipo de diária", "Segmento: Hospedagem", dt[col(44, "Tipo de diária")])
    add("Forma de reserva", "Segmento: Hospedagem", dt[col(45, "Formas de reserva")])
    equip = []
    for _, r in dt.iterrows():
        l = []
        for c in C[58:93]:
            if clean(r[c]) == "Sim":
                l.append(re.search(r"\[(.*)\]", c).group(1).strip())
        equip.append(l)
    assert str(C[58]).endswith("[Aparelho de vídeo]") and str(C[92]).endswith("[Outros]")
    add("Equipamento/serviço de hospedagem", "Segmento: Hospedagem", equip)
    add("Tipo de feira", "Segmento: Cultural", dt[col(121, "Em caso de feiras: qual o tipo")])
    add("Tipo de atrativo natural", "Segmento: Recursos Naturais", dt[col(125, "Tipo de atrativo")])
    add("Estrutura do atrativo natural", "Segmento: Recursos Naturais", dt[col(130, "O local possui")],
        padrao={"Banheiros", "Lanchonetes/bares", "Estacionamento", "Áreas de piquenique", "Pontos de descanso"})
    R = pd.DataFrame(R, columns=["id_empreendimento", "tema", "pergunta", "opcao", "detalhe_outras"])
    ordem_dias = {d: i + 1 for i, d in enumerate(DIAS)}
    R["ordem_opcao"] = R["opcao"].map(lambda o: ordem_dias.get(o, 99 if o == "Outras respostas" else 50))

    # =========================================================== atributos específicos por segmento (longo)
    A = []

    def attr(seg, nome, series, mapa=None):
        mask = dt[seg_cols[SEGMENTOS.index(seg)]] == 1
        for i, m, v in zip(ID, mask, series):
            if m:
                val = clean(v) if not isinstance(v, (int, float)) or pd.isna(v) else str(int(v))
                if mapa:
                    val = mapa.get(val, val)
                A.append((i, seg, nome, val if val is not None else "Não informado"))

    attr("Alimentação", "Tipo de atendimento", dt[col(37, "Tipo de atendimento")])
    attr("Alimentação", "Tipo de serviço", dt[col(39, "Tipo de serviço")],
         {k: "Outras respostas" for k in ["Delivery", "Cervejaria", "Prato executivo", "Cafeteria", "Unidade",
                                           "Porções", "Lanches", "Rodízio"]})
    attr("Alimentação", "Gastronomia", dt[col(40, "Gastronomia")].map(
        lambda v: None if pd.isna(v) else ("Caseira e italiana" if norm(v) == "caseira e italiana" else
                                           "Caseira" if norm(v).startswith("caseira") else "Outras respostas")))
    attr("Alimentação", "Espaço infantil", dt[col(42, "Possui espaço infantil")])
    meses = dt[C[46:58]].apply(lambda r: (r.map(clean) == "Sim").sum(), axis=1)
    attr("Hospedagem", "Funciona o ano todo", meses.map(lambda m: "Sim" if m == 12 else "Não"))
    attr("Hospedagem", "Oferece almoço e jantar", dt[col(96, "Oferece refeições")])
    attr("Hospedagem", "Aceita animais de estimação", dt[col(97, "Aceita animais")])
    attr("Hospedagem", "Espaço infantil", dt[col(98, "Possui espaço infantil")])
    attr("Hospedagem", "Descrição de leitos/quartos (texto livre)", dt[col(94, "Qual a quantidade total de leitos")])
    attr("Entretenimento/Aventura", "Necessário reservar", dt[col(100, "É necessário reservar")])
    attr("Entretenimento/Aventura", "Capacidade máxima", dt[col(101, "Seu estabelecimento tem capacidade máxima")])
    attr("Entretenimento/Aventura", "Possui guia/instrutor", dt[col(102, "Possui guia")])
    attr("Entretenimento/Aventura", "Exige equipamento de proteção", dt[col(103, "Há necessidade do uso")])
    attr("Entretenimento/Aventura", "Restrição de participação", dt[col(104, "Existe alguma restrição")])
    attr("Agroindústria", "Produto (texto livre)", dt[col(106, "Qual tipo de produto")])
    attr("Agroindústria", "Espaço de degustação", dt[col(107, "Existe algum espaço")])
    attr("Agroindústria", "Certificação/selo de inspeção", dt[col(108, "Possui certificações")])
    attr("Agroindústria", "Turismo de experiência", dt[col(109, "Possui turismo de experiência")])
    attr("Compras", "Tipo de comércio", dt[col(111, "Tipo de comércio")])
    attr("Compras", "Embalagem para lembrança", dt[col(112, "O comércio oferece embalagens")])
    attr("Compras", "Origem dos produtos", dt[col(113, "Os produtos comercializados")])
    attr("Compras", "Interesse em projetos coletivos", dt[col(114, "Há interesse em participar")])
    attr("Cultural", "Necessário reservar", dt[col(116, "É necessário reservar")])
    attr("Cultural", "Forma de entrada", dt[col(117, "Qual a forma de entrada")])
    attr("Cultural", "Fluxo aumenta em eventos", dt[col(118, "O fluxo de turistas aumenta")])
    attr("Cultural", "Guia/monitor/material", dt[col(119, "O local possui guia")])
    attr("Cultural", "Capacidade simultânea", dt[col(120, "Qual a capacidade máxima de visitantes")])
    attr("Cultural", "Espaço estruturado para feiras", dt[col(122, "Em caso de feiras e festas:  possui")])
    attr("Cultural", "Periodicidade de feiras/festas", dt[col(123, "Em caso de feiras e festas: qual a periodicidade")])
    attr("Recursos Naturais", "Em área pública", dt[col(126, "Esse atrativo está localizado")])
    attr("Recursos Naturais", "Acesso por via pública", dt[col(127, "O acesso a esse atrativo")])
    attr("Recursos Naturais", "Responsável pela gestão", dt[col(128, "Quem é responsável")])
    attr("Recursos Naturais", "Trilhas/caminhos sinalizados", dt[col(129, "Existem trilhas")])
    attr("Recursos Naturais", "Normas de conduta/controle", dt[col(131, "Existem normas")])
    attr("Recursos Naturais", "Acessível a mobilidade reduzida", dt[col(132, "O local é acessível")])
    attr("Recursos Naturais", "Sazonalidade (texto livre)", dt[col(133, "Existe sazonalidade")])
    A = pd.DataFrame(A, columns=["id_empreendimento", "segmento", "atributo", "valor"])

    # =========================================================== colunas derivadas do buscador
    # Daqui para baixo estão as regras que existiam só dentro do JavaScript do buscador
    # (montarModelo, resumoDias, fichaHTML) e que o guia do Power BI mandava reimplementar em
    # Power Query/DAX. Passam a ser colunas e tabelas prontas: uma regra, um lugar.
    segs_por_emp = {i: [] for i in ID}
    for r in ponte.itertuples():
        segs_por_emp[r.id_empreendimento].append(r.segmento)

    # respostas múltiplas e atributos indexados, preservando a ordem do questionário
    opc = {}
    for r in R.itertuples():
        opc.setdefault((r.id_empreendimento, r.pergunta), []).append(r.opcao)
    val = {}
    for r in A.itertuples():
        chave = (r.id_empreendimento, r.segmento, r.atributo)
        assert chave not in val, chave
        val[chave] = r.valor

    def opcoes(i, pergunta):
        return opc.get((i, pergunta), [])

    def atributo(i, seg, nome):
        return clean(val.get((i, seg, nome)))

    # itens concretos de acessibilidade — a autodeclaração (facilidade_pcd/capacitado_pcd)
    # nunca entra aqui, por decisão registrada em site-visitante/LEIA-ME.md
    recursos = {i: [o for o in opcoes(i, "Recurso de acessibilidade")
                    if o not in ("Não possui", "Outras respostas")] for i in ID}
    itens_sanitario = {i: [o for o in opcoes(i, "Item de sanitário adaptado")
                           if o != "Não possui sanitário adaptado"] for i in ID}

    E["distrito"] = E["id_regiao"].map(REGIOES)
    E["acessivel"] = [bool(san) or any(sem_acento(r).startswith("rampa") for r in recursos[i])
                      for i, san in zip(ID, E["sanitario_adaptado"])]
    # Rótulo único de acessibilidade do card. A regra do buscador é EXCLUSIVA: quem tem
    # sanitário adaptado mostra só "Sanitário adaptado"; a rampa aparece apenas para quem
    # não tem sanitário. Sem esta coluna, os 16 que têm os dois mostrariam dois badges.
    E["badge_acessibilidade"] = ["Sanitário adaptado" if san else ("Rampa de acesso" if ac else None)
                                 for san, ac in zip(E["sanitario_adaptado"], E["acessivel"])]
    pet = [atributo(i, "Hospedagem", "Aceita animais de estimação") for i in ID]
    E["aceita_pet"] = [p in ("Sim", "Depende do porte do animal") for p in pet]
    E["pet_texto"] = pet
    # um empreendimento que atua em dois segmentos responde duas vezes: basta um "Sim"
    E["espaco_infantil"] = [any(atributo(i, seg, "Espaço infantil") == "Sim"
                                for seg in ("Alimentação", "Hospedagem")) for i in ID]
    E["tem_idioma"] = E["idiomas"].map(lambda v: len(tokens(v)) > 0)

    def resumir_dias(l):
        if not l:
            return None
        if len(l) == 7:
            return "Todos os dias"
        return " · ".join(DIAS_ABREV[DIAS.index(d)] for d in DIAS if d in l)

    E["dias_resumo"] = [resumir_dias(l) or t for l, t in zip(dias_lst, E["dias_funcionamento"])]
    # Segmentos como o card os exibe: rótulo de visitante, na ordem dos chips, já unidos.
    # A coluna "segmentos" ao lado guarda os nomes técnicos, na ordem de SEGMENTOS.
    E["segmentos_rotulo"] = [" · ".join(rot_vis[s] for s in sorted(segs_por_emp[i], key=ordem_chip.get))
                             for i in ID]
    # "Urbana" é o valor do questionário; "zona urbana" é como o card fala com o visitante.
    E["localizacao_rotulo"] = E["localizacao"].map(lambda v: "zona " + v.lower(), na_action="ignore")
    # Selo de idioma, já como frase: "Atende em inglês, francês". Só os 18 que declararam.
    E["badge_idiomas"] = [("Atende em " + v.lower().replace("; ", ", ")) if isinstance(v, str) and v else None
                          for v in E["idiomas"]]
    # busca textual insensível a acento: o campo já chega normalizado para o filtro só comparar
    E["busca"] = [sem_acento(" ".join(x for x in [nome, end, ref, dist, " ".join(segs_por_emp[i])] if x))
                  for i, nome, end, ref, dist in zip(ID, E["nome_empreendimento"],
                                                     E["endereco_completo"].fillna(""),
                                                     E["ponto_referencia"].fillna(""), E["distrito"])]

    # =========================================================== dimensões da interface
    dim_dia = pd.DataFrame({"dia": DIAS, "dia_abrev": DIAS_ABREV, "ordem_dia": range(1, 8)})
    dim_grupo = pd.DataFrame(GRUPOS_FAIXA, columns=["id_grupo", "rotulo_grupo", "grupo_min"])
    dim_grupo["ordem_grupo"] = dim_grupo["id_grupo"]
    dim_conveniencia = pd.DataFrame(CONVENIENCIAS,
                                    columns=["id_conveniencia", "conveniencia", "coluna_origem"])
    dim_conveniencia["ordem_conveniencia"] = range(1, len(CONVENIENCIAS) + 1)

    seg_cap, col_cap, rot_cap, ord_cap = FICHA_CAPACIDADE
    fcfg = []
    for seg, _ in ROTULO_VISITANTE:
        if seg == seg_cap:
            fcfg.append((seg, "capacidade", col_cap, rot_cap, ord_cap))
        for k, nome in enumerate(FICHA_ATRIBUTOS.get(seg, []), 1):
            fcfg.append((seg, "atributo", nome, ROTULO_ATRIBUTO.get(nome, nome), k))
        for k, (perg, rot) in enumerate(FICHA_LISTAS.get(seg, []), 1):
            fcfg.append((seg, "lista", perg, rot, k))
    cfg_ficha = pd.DataFrame(fcfg, columns=["segmento", "tipo", "chave", "rotulo", "ordem"])

    # =========================================================== fatos derivados
    # dias no grão (empreendimento, dia), já na ordem da semana
    dias_fato = pd.DataFrame(
        [(i, d, DIAS_ABREV[DIAS.index(d)], DIAS.index(d) + 1)
         for i, l in zip(ID, dias_lst) for d in DIAS if d in l],
        columns=["id_empreendimento", "dia", "dia_abrev", "ordem_dia"])

    # ponte do filtro de tamanho de grupo: a condição é ordinal, com teto e uma exceção
    grupo_faixa = pd.DataFrame(
        [(g, i) for g, _, mn in GRUPOS_FAIXA for i, o in zip(ID, E["grupos_ordem"])
         if (mn <= o <= GRUPO_ORDEM_TETO) or o == GRUPO_ORDEM_SEM_LIMITE],
        columns=["id_grupo", "id_empreendimento"])

    # ponte das conveniências: 1 linha = (empreendimento, conveniência que ele tem)
    emp_conveniencia = pd.DataFrame(
        [(i, cid) for cid, _, coluna in CONVENIENCIAS for i, v in zip(ID, E[coluna]) if v],
        columns=["id_empreendimento", "id_conveniencia"])

    # recursos de acessibilidade no grão (empreendimento, recurso). Só texto categórico
    # estável: os recursos declarados na pergunta, mais o sanitário adaptado, que vem de
    # coluna. A contagem de quartos adaptados NÃO entra aqui — ela já é o número em
    # dim_empreendimento[uh_adaptadas_pcd]; codificá-la como texto ("2 quartos adaptados")
    # fazia o domínio de `item` crescer sem limite conforme os dados mudam.
    acess_itens = []
    for i, san in zip(ID, E["sanitario_adaptado"]):
        itens = list(recursos[i])
        if san:
            itens.append("Sanitário adaptado")
        acess_itens += [(i, maiuscula(t), k) for k, t in enumerate(itens, 1)]
    acess_itens = pd.DataFrame(acess_itens, columns=["id_empreendimento", "item", "ordem_item"])

    # itens do sanitário: viram a frase "No sanitário: …" da ficha, por isso em minúsculas
    san_itens = pd.DataFrame(
        [(i, t.lower(), k) for i in ID for k, t in enumerate(itens_sanitario[i], 1)],
        columns=["id_empreendimento", "item", "ordem_item"])

    # atributos e listas da ficha, já filtrados, rotulados e ordenados
    atr_ficha, listas_ficha = [], []
    for i, cap in zip(ID, E[col_cap]):
        for seg in segs_por_emp[i]:
            if seg == seg_cap and pd.notna(cap):
                atr_ficha.append((i, seg, col_cap, rot_cap, "até %d pessoas" % cap, ord_cap))
            for k, nome in enumerate(FICHA_ATRIBUTOS.get(seg, []), 1):
                v = atributo(i, seg, nome)
                if v is None or v in SENTINELAS_FICHA:
                    continue
                atr_ficha.append((i, seg, nome, ROTULO_ATRIBUTO.get(nome, nome), v, k))
            for k, (perg, rot) in enumerate(FICHA_LISTAS.get(seg, []), 1):
                vs = [o for o in opcoes(i, perg) if o != "Outras respostas"]
                listas_ficha += [(i, seg, perg, rot, k, maiuscula(o), j)
                                 for j, o in enumerate(vs, 1)]
    atr_ficha = pd.DataFrame(atr_ficha, columns=["id_empreendimento", "segmento", "atributo",
                                                 "rotulo_atributo", "valor", "ordem_atributo"])
    listas_ficha = pd.DataFrame(listas_ficha, columns=["id_empreendimento", "segmento", "pergunta",
                                                       "rotulo_lista", "ordem_lista", "opcao",
                                                       "ordem_opcao_ficha"])

    # =========================================================== gravação
    kw = dict(index=False, sep=";", encoding="utf-8-sig")
    TABELAS = [
        # dimensões
        ("dim_empreendimento", E),
        ("dim_regiao", dim_regiao),
        ("dim_segmento", dim_segmento),
        ("dim_dia", dim_dia),
        ("dim_grupo", dim_grupo),
        ("dim_conveniencia", dim_conveniencia),
        # pontes (só chaves estrangeiras)
        ("ponte_empreendimento_segmento", ponte),
        ("ponte_empreendimento_conveniencia", emp_conveniencia),
        ("ponte_empreendimento_grupo", grupo_faixa),
        ("ponte_empreendimento_dia", dias_fato),
        # detalhes do empreendimento (payload próprio)
        ("det_empreendimento_atributo", A),
        ("det_empreendimento_atributo_ficha", atr_ficha),
        ("det_empreendimento_opcao_ficha", listas_ficha),
        ("det_empreendimento_acessibilidade", acess_itens),
        ("det_empreendimento_sanitario", san_itens),
        # fato (grão de evento com medida)
        ("fato_resposta_questionario", R),
        # configuração de interface
        ("cfg_ficha", cfg_ficha),
    ]
    for nome, d in TABELAS:
        d.to_csv(out / (nome + ".csv"), **kw)
    dic = [(nome, c, tipo_amigavel(c, d[c]), DESCRICOES.get(c, ""))
           for nome, d in TABELAS for c in d.columns]
    pd.DataFrame(dic, columns=["tabela", "coluna", "tipo", "descricao"]).to_csv(out / "dicionario_dados.csv", **kw)
    for nome, d in TABELAS:
        print(f"{nome:30s} {d.shape[0]:5d} linhas x {d.shape[1]:2d} colunas")


# ---------------------------------------------------------------- guardas de entrada (CLI)
# Conferências feitas ANTES de processar, para que um arquivo, uma aba ou uma coluna fora do
# lugar produza uma mensagem em português dizendo o que fazer, em vez de um traceback. Nada
# aqui altera a leitura ou a transformação dos dados.

COLUNAS_EXIGIDAS = {
    ABA_BRUTA: ["Carimbo de data/hora", "Nome do empreendimento?"],
    ABA_TRATADA: ["Nome do empreendimento?"],
}
MIN_COLUNAS_TRATADA = 134  # a maior coluna lida por posição é o índice 133


def erro(titulo, *o_que_fazer):
    print("\nERRO: " + titulo, file=sys.stderr)
    for linha in o_que_fazer:
        print("  " + linha, file=sys.stderr)
    print("\nPasso a passo completo: scripts/README.md\n", file=sys.stderr)
    sys.exit(1)


def conferir_entrada(xlsx):
    """Falha com mensagem legível quando a planilha de entrada não existe, não abre,
    ou não tem as abas/colunas que o script espera."""
    caminho = Path(xlsx)

    if not caminho.exists():
        pasta = caminho.parent if str(caminho.parent) else Path(".")
        if not pasta.is_dir():
            erro(f"não encontrei o arquivo de entrada: {caminho}",
                 f"A pasta '{pasta}' não existe a partir de onde você rodou o comando.",
                 "O caminho é relativo à pasta atual — rode o comando da raiz do projeto.")
        achadas = sorted(a.name for a in pasta.glob("*.xls*") if not a.name.startswith("~$"))
        erro(f"não encontrei o arquivo de entrada: {caminho}",
             "Planilhas que existem em '%s': %s" % (pasta, ", ".join(achadas) if achadas else "(nenhuma)"),
             "Coloque a planilha exportada do Google Sheets em dados/brutos/ e repita o comando",
             "com o nome exato do arquivo, entre aspas (o nome tem espaços e acentos).")

    if caminho.is_dir():
        erro(f"'{caminho}' é uma pasta, não a planilha.",
             "O primeiro argumento é o arquivo .xlsx de entrada; o segundo é a pasta de saída.")

    if caminho.suffix.lower() not in (".xlsx", ".xlsm"):
        erro(f"o arquivo de entrada não é uma planilha .xlsx: {caminho.name}",
             "No Google Sheets use 'Arquivo -> Fazer download -> Microsoft Excel (.xlsx)'.",
             "Formatos .csv, .ods e .xls antigos não servem: o script precisa das duas abas.")

    try:
        with pd.ExcelFile(caminho) as planilha:
            abas = list(planilha.sheet_names)
    except Exception as e:
        erro(f"não consegui abrir a planilha {caminho.name}.",
             f"Detalhe técnico: {type(e).__name__}: {e}",
             "Se o arquivo veio do Google Sheets, baixe de novo como 'Microsoft Excel (.xlsx)'.",
             "Se a mensagem fala em 'openpyxl', a biblioteca não está instalada:",
             "rode 'pip install -r requirements.txt' com o ambiente virtual ativado.")

    faltando = [a for a in (ABA_BRUTA, ABA_TRATADA) if a not in abas]
    if faltando:
        erro("a planilha não tem a(s) aba(s) que o script espera: %s" % ", ".join(repr(a) for a in faltando),
             "Abas encontradas no arquivo: %s" % ", ".join(repr(a) for a in abas),
             "O nome da aba de respostas vem do título do formulário e é cortado em 31 caracteres",
             "pelo Google Sheets: se o formulário foi renomeado, ajuste ABA_BRUTA / ABA_TRATADA",
             "no início deste arquivo para bater com os nomes acima.")

    for aba, exigidas in COLUNAS_EXIGIDAS.items():
        colunas = [str(c) for c in pd.read_excel(caminho, sheet_name=aba, nrows=0).columns]
        ausentes = [c for c in exigidas if c not in colunas]
        if ausentes:
            erro("a aba %r não tem a(s) coluna(s) esperada(s): %s"
                 % (aba, ", ".join(repr(c) for c in ausentes)),
                 "A aba tem %d colunas. As 5 primeiras: %s"
                 % (len(colunas), ", ".join(repr(c) for c in colunas[:5])),
                 "Provavelmente o cabeçalho da pergunta foi editado no formulário.",
                 "Restaure a redação original da pergunta ou ajuste o nome dentro do script.")

    colunas_tratada = pd.read_excel(caminho, sheet_name=ABA_TRATADA, nrows=0).columns
    if len(colunas_tratada) < MIN_COLUNAS_TRATADA:
        erro("a aba %r tem só %d colunas; o script espera pelo menos %d."
             % (ABA_TRATADA, len(colunas_tratada), MIN_COLUNAS_TRATADA),
             "O script lê essa aba por posição de coluna. Faltando colunas, ele leria a pergunta errada.",
             "Confira se a aba foi baixada inteira e se nenhuma coluna foi apagada.")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(
            "\nUso:\n"
            '    python scripts/preparar_dados.py "<planilha .xlsx de entrada>" "<pasta de saída>"\n'
            "\nExemplo (rodando da raiz do projeto):\n"
            '    python scripts/preparar_dados.py "dados/brutos/ALFREDO CHAVES - Diagnóstico Turístico - Formulário.xlsx" dados/tratados\n'
            "\nPasso a passo completo: scripts/README.md\n"
        )
    conferir_entrada(sys.argv[1])
    main(sys.argv[1], sys.argv[2])
