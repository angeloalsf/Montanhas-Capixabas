#!/usr/bin/env python3
"""
Validação cruzada: recalcula, a partir da aba 'Dados tratados', os números publicados no PDF
'Estatistica descritiva - gráficos e tabelas.pdf' e compara nas duas bases possíveis:
  - base 92: exclui o único empreendimento que não está em nenhum dos 5 segmentos do Gráfico 1 (base do PDF);
  - base 93: todos os empreendimentos da aba tratada (= total da legenda).

Uso: python validar_pdf.py "<planilha.xlsx>" "<saida.csv>"
"""
import re
import sys

import pandas as pd

TOL = {"pct": 0.051, "n": 0.0, "stat": 0.051}


def clean(v):
    return None if pd.isna(v) else (re.sub(r"\s+", " ", str(v)).strip() or None)


def toks(v):
    v = clean(v)
    return [] if v is None else [t for t in (clean(x) for x in v.split(";")) if t]


def main(xlsx, out_csv):
    dt = pd.read_excel(xlsx, sheet_name="Dados tratados")
    C = list(dt.columns)
    seg = C[0:7]
    b92 = (dt[seg[:5]] == 1).any(axis=1)
    rows = []

    def chk(ref, item, pdf, kind, fn, sub=None):
        """fn(df) -> numerador (ou estatística); sub(df) -> máscara do universo (segmento)."""
        out = {}
        for nome, base in (("base92", dt[b92]), ("base93", dt)):
            d = base if sub is None else base[sub(base)]
            v = fn(d)
            if kind == "pct":
                v = round(100 * v / len(d), 1)
            out[nome] = v
        ok = {k: abs(float(v) - float(pdf)) <= TOL[kind] + 1e-9 for k, v in out.items()}
        status = ("OK" if ok["base92"] else "OK só na base 93" if ok["base93"] else "DIVERGE")
        rows.append({"referencia_pdf": ref, "item": item, "tipo": kind, "valor_pdf": pdf,
                     "calculado_base92": out["base92"], "calculado_base93": out["base93"], "status": status})

    eq = lambda c, v: (lambda d: int((d[c].map(clean) == v).sum()))
    tk = lambda c, v: (lambda d: int(d[c].map(lambda x: v in toks(x)).sum()))
    notin = lambda c, vs: (lambda d: int(d[c].map(lambda x: clean(x) not in vs).sum()))
    has_other = lambda c, vs: (lambda d: int(d[c].map(lambda x: any(t not in vs for t in toks(x))).sum()))
    segm = lambda i: (lambda d: d[seg[i]] == 1)

    # ---- Gráfico 1 / Tabela 1
    for i, (nome, pct) in enumerate([("Alimentação", 50.0), ("Hospedagem", 44.6), ("Entretenimento/Aventura", 12.0),
                                     ("Agroindústria", 18.5), ("Compras", 19.6)]):
        chk("Gráfico 1", nome, pct, "pct", lambda d, i=i: int((d[seg[i]] == 1).sum()))
    tab1 = {1: [5, 3, 0, 1, 0, 8], 2: [1, 2, 1, 0, 0, 2], 3: [17, 23, 4, 6, 5, 36], 4: [1, 3, 2, 0, 0, 3],
            5: [1, 0, 0, 3, 2, 4], 6: [5, 2, 2, 5, 5, 9], 7: [16, 8, 2, 2, 6, 30]}
    nomes = ["Alimentação", "Hospedagem", "Entretenimento/Aventura", "Agroindústria", "Compras", "Estabelecimentos"]
    R = C[7]
    for reg, vals in tab1.items():
        for j, v in enumerate(vals):
            f = (lambda d, reg=reg, j=j: int(((d[R] == reg) & (d[seg[j]] == 1)).sum())) if j < 5 else \
                (lambda d, reg=reg: int((d[R] == reg).sum()))
            chk("Tabela 1", f"Região {reg} × {nomes[j]}", v, "n", f)
    # ---- Informações gerais
    L = C[10]
    chk("Gráfico 2", "Rural", 77.2, "pct", lambda d: int((d[L] == 2).sum()))
    chk("Gráfico 2", "Urbana", 22.8, "pct", lambda d: int((d[L] == 1).sum()))
    V = C[11]
    for k, v in [("Asfaltada", 23), ("Terra", 23), ("Revsol/ Brita", 21), ("Calçada", 16)]:
        chk("Tabela 2", f"Via: {k}", v, "n", eq(V, k))
    chk("Tabela 2", "Via: outras respostas", 9, "n", notin(V, {"Asfaltada", "Terra", "Revsol/ Brita", "Calçada"}))
    P = C[12]
    for k, v in [("Possui, sem padronização (modelo próprio)", 62.0), ("Não possui", 33.7), ("Possui, padrão municipal", 2.2),
                 ("Possui, padrão EMBRATUR", 1.1), ("Possui, padrão municipal e sem padronização (modelo próprio)", 1.1)]:
        chk("Gráfico 3", f"Placa: {k}", v, "pct", eq(P, k))
    chk("Tabela 3", "Faixas no chão: Sim", 7, "n", eq(C[13], "Sim"))
    chk("Tabela 3", "Faixas no chão: Não", 84, "n", eq(C[13], "Não"))
    chk("Tabela 3", "Sinalização contribui: Sim", 40, "n", eq(C[14], "Sim"))
    chk("Tabela 3", "Sinalização contribui: Não", 46, "n", eq(C[14], "Não"))
    chk("Tabela 3", "Sinalização contribui: outras", 6, "n", notin(C[14], {"Sim", "Não"}))
    Fn = C[15]
    for k, v, f in [("n", 92, len), ("média", 3.4, lambda s: round(s.mean(), 1)), ("mediana", 2.0, lambda s: s.median()),
                    ("desvio padrão", 5.25, lambda s: round(s.std(), 2)), ("mínimo", 0, lambda s: s.min()),
                    ("máximo", 40, lambda s: s.max()), ("Q1", 1.0, lambda s: s.quantile(.25)), ("Q3", 4.0, lambda s: s.quantile(.75))]:
        chk("Tabela 4", f"Funcionários: {k}", v, "stat", lambda d, f=f: f(d[Fn].astype(float)))
    D = C[19]
    for k, v in [("Instagram", 83.7), ("Whatsapp", 62.0), ("Facebook", 22.8), ("Sites de turismo (Booking, Trivago, Airbnb, etc)", 16.3),
                 ("Outras respostas", 13.0), ("Parceria com outros comércios (hotéis/pousadas/restaurantes)", 8.7),
                 ("Indicação de guias", 5.4), ("Tiktok", 3.3)]:
        chk("Gráfico 4", f"Divulgação: {k}", v, "pct", tk(D, k))
    NM = C[20]
    nm = [("Melhoria das estradas e acessos aos atrativos turísticos", 79), ("Divulgação e promoção dos atrativos turísticos", 75),
          ("Melhorias na sinalização vertical (placas)", 63), ("Promover cursos e capacitação para profissionais do setor", 53),
          ("Melhorias na sinalização horizontal", 30)]
    for k, v in nm:
        chk("Tabela 5", f"Necessidade município: {k}", v, "n", tk(NM, k))
    chk("Tabela 5", "Necessidade município: outras respostas", 13, "n", has_other(NM, {k for k, _ in nm}))
    NE = C[21]
    ne = [("Marketing e redes sociais", 62.0), ("Mão de obra", 47.8), ("Vendas e estratégias comerciais", 39.1),
          ("Capacitação de mão de obra", 37.0), ("Aspectos jurídicos/regularização", 17.4), ("Gestão administrativa", 17.4),
          ("Gestão financeira", 15.2)]
    for k, v in ne:
        chk("Gráfico 5", f"Necessidade empreendimento: {k}", v, "pct", tk(NE, k))
    chk("Gráfico 5", "Necessidade empreendimento: outras respostas", 12.0, "pct", has_other(NE, {k for k, _ in ne}))

    # ---- funcionamento (dias recalculados a partir da tabela tratada)
    from preparar_dados import parse_dias  # mesma regra usada nas tabelas tratadas
    DD = C[22]
    dias = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado", "Domingo"]
    for dname, v in zip(dias, [60.9, 63.0, 65.2, 72.8, 90.2, 88.0, 82.6]):
        chk("Gráfico 6", f"Abre {dname}", v, "pct", lambda d, dn=dname: int(d[DD].map(lambda x: dn in parse_dias(x)).sum()))
    for nd, v in [(2, 10), (3, 15), (4, 7), (5, 8), (6, 16), (7, 36)]:
        chk("Tabela 6", f"{nd} dias por semana", v, "n", lambda d, nd=nd: int(d[DD].map(lambda x: len(parse_dias(x)) == nd).sum()))
    chk("Tabela 7", "Atende 24h: Sim", 12, "n", eq(C[23], "Sim"))
    chk("Tabela 7", "Atende 24h: Não", 80, "n", eq(C[23], "Não"))
    chk("Tabela 7", "Abre feriados: Sim", 78, "n", eq(C[24], "Sim"))
    chk("Tabela 7", "Abre feriados: Não", 14, "n", eq(C[24], "Não"))
    G = C[25]
    for k, v in [("De 10 a 25 visitantes", 34.8), ("De 25 a 50 visitantes", 17.4), ("Até 10 visitantes", 13.0),
                 ("Não recebe grupos", 12.0), ("Acima de 100 visitantes", 9.8), ("De 75 a 100 visitantes", 6.5),
                 ("De 50 a 75 visitantes", 4.3), ("Não se aplica", 1.1), ("Capacidade não definida", 1.1)]:
        chk("Gráfico 7", f"Grupos: {k}", v, "pct", eq(G, k))
    for k, v in [("Visitantes do Espírito Santo", 63), ("Moradores locais", 24), ("Turistas de outros estados brasileiros", 5)]:
        chk("Tabela 8", f"Público: {k}", v, "n", eq(C[26], k))
    S = C[27]
    for k, v in [("Gratuito, descoberto", 68.5), ("Estacionamento na rua", 17.4), ("Não possui estacionamento", 7.6),
                 ("Gratuito, coberto e descoberto", 3.3), ("Gratuito, coberto", 2.2), ("Pago, coberto", 1.1)]:
        chk("Gráfico 8", f"Estacionamento: {k}", v, "pct", eq(S, k))
    own = lambda d: d[S].map(clean).str.startswith(("Gratuito", "Pago"))
    for c, nome, vals in [(C[28], "automóveis", [66, 3, 25.9, 20, 1, 130, 8, 30]), (C[29], "ônibus", [67, 2, 2.6, 2, 0, 20, 0, 4])]:
        fs = [("válidos", lambda s: s.notna().sum()), ("ausentes", lambda s: s.isna().sum()), ("média", lambda s: round(s.mean(), 1)),
              ("mediana", lambda s: s.median()), ("mínimo", lambda s: s.min()), ("máximo", lambda s: s.max()),
              ("Q1", lambda s: s.quantile(.25)), ("Q3", lambda s: s.quantile(.75))]
        for (k, f), v in zip(fs, vals):
            chk("Tabela 9 (vagas)", f"Vagas {nome}: {k}", v, "stat", lambda d, f=f, c=c: f(d.loc[own(d), c].astype(float)))
    LG = C[30]
    for k, v in [("Não", 80.4), ("Inglês", 9.8), ("Espanhol", 4.3), ("Italiano", 4.3), ("Inglês e Francês", 1.1)]:
        chk("Gráfico 9", f"Língua: {k}", v, "pct", eq(LG, k))
    PG = C[31]
    for k, v in [("Pix", 98.9), ("Dinheiro", 94.6), ("Cartão de débito", 58.7), ("Cartão de crédito", 57.6), ("Cheque", 7.6), ("Vale feira", 1.1)]:
        chk("Gráfico 10", f"Pagamento: {k}", v, "pct", tk(PG, k))
    chk("Gráfico 11", "Facilidade PcD: Sim (código 1)", 50.0, "pct", lambda d: int((d[C[32]] == 1).sum()))
    CP = C[33]
    for k, v in [("Não", 53.3), ("Fisica", 44.6), ("Auditiva", 19.6), ("Visual", 9.8), ("Mental", 8.7), ("Múltipla", 4.3)]:
        chk("Gráfico 12", f"Capacitado PcD: {k}", v, "pct", tk(CP, k))
    chk("Texto após Gráfico 13", "Capacitado PcD: Visual (texto)", 9.7, "pct", tk(CP, "Visual"))
    chk("Texto após Gráfico 13", "Capacitado PcD: Auditiva (texto)", 19.4, "pct", tk(CP, "Auditiva"))
    AC = C[34]
    for k, v in [("Não possui", 43.5), ("Rampa", 43.5), ("Estacionamento", 38.0), ("Rampa de acesso à calçada", 16.3),
                 ("Corrimão", 10.9), ("Elevador", 2.2), ("Interprete de libras", 1.1), ("Plataforma elevativa", 1.1)]:
        chk("Gráfico 13", f"Acessibilidade: {k}", v, "pct", tk(AC, k))
    SA = C[35]
    for k, v in [("Não", 76), ("Porta larga suficiente para entrada de cadeira de rodas", 13), ("Acesso para cadeira de rodas", 13),
                 ("Giro para cadeira de rodas", 9), ("Barra de apoio", 8), ("Pia rebaixada", 7),
                 ("Espelho rebaixado ou com ângulo de alcance visual", 5), ("Boxe ou banheira adaptada", 4), ("Torneira monocomando/alavanca", 3)]:
        chk("Tabela 10 (sanitário)", f"Sanitário: {k}", v, "n", tk(SA, k))
    chk("Texto após Tabela 10", "Sem sanitário adaptado (texto)", 82.8, "pct", tk(SA, "Não"))

    # ---- Alimentação (universo = segmento)
    A = segm(0)
    cat = [("Restaurante", 41.3), ("Lanchonete", 34.8), ("Cafeteria", 15.2), ("Cervejaria", 13.0), ("Confeitaria", 10.9),
           ("Padaria", 10.9), ("Pizzaria", 10.9)]
    for k, v in cat:
        chk("Gráfico 14 (nota diz n=45)", f"Alimentação: {k}", v, "pct", tk(C[38], k), A)
    chk("Gráfico 14 (nota diz n=45)", "Alimentação: outras respostas", 26.1, "pct", has_other(C[38], {k for k, _ in cat}), A)
    for k, v in [("À La Carte", 21), ("Self-service", 14)]:
        chk("Tabela 9 (alimentação)", f"Tipo de serviço: {k}", v, "n", eq(C[39], k), A)
    chk("Tabela 9 (alimentação)", "Tipo de serviço: não informado", 3, "n", lambda d: int(d[C[39]].isna().sum()), A)
    chk("Tabela s/ nº (gastronomia)", "Gastronomia: Caseira (inclui 'Caseira, ...')", 31, "n",
        lambda d: int(d[C[40]].map(lambda x: clean(x) is not None and clean(x).lower().startswith("caseira")
                                   and clean(x).lower() != "caseira e italiana").sum()), A)
    chk("Tabela s/ nº (gastronomia)", "Gastronomia: não informado", 4, "n", lambda d: int(d[C[40]].isna().sum()), A)
    capA = lambda d: pd.to_numeric(d[C[41]], errors="coerce")
    for k, v, f in [("válidos", 36, lambda s: s.notna().sum()), ("ausentes", 10, lambda s: s.isna().sum()),
                    ("média", 118.8, lambda s: round(s.mean(), 1)), ("mediana", 65.5, lambda s: s.median()),
                    ("mínimo", 10, lambda s: s.min()), ("máximo", 1000, lambda s: s.max()),
                    ("Q1", 36.3, lambda s: s.quantile(.25)), ("Q3", 127.5, lambda s: s.quantile(.75))]:
        chk("Tabela 10 (alimentação)", f"Capacidade: {k}", v, "stat", lambda d, f=f: f(capA(d)), A)
    chk("Gráfico 15", "Espaço infantil: Sim", 47.8, "pct", eq(C[42], "Sim"), A)
    chk("Gráfico X (alim.)", "Alimentação rural", 71.7, "pct", lambda d: int((d[L] == 2).sum()), A)
    for k, v, f in [("n", 46, len), ("média", 3.9, lambda s: round(s.mean(), 1)), ("mediana", 2.5, lambda s: s.median()),
                    ("desvio padrão", 4.1, lambda s: round(s.std(), 1)), ("máximo", 17, lambda s: s.max()),
                    ("Q1", 2.0, lambda s: s.quantile(.25)), ("Q3", 5.0, lambda s: s.quantile(.75))]:
        chk("Tabela X (alim. funcionários)", f"Funcionários alim.: {k}", v, "stat", lambda d, f=f: f(d[Fn].astype(float)), A)
    for k, v in [("Instagram", 89.1), ("Whatsapp", 58.7), ("Facebook", 30.4), ("Outras respostas", 15.2)]:
        chk("Gráfico X (alim. divulgação)", f"Divulgação alim.: {k}", v, "pct", tk(D, k), A)
    for k, v in [("Mão de obra", 58.7), ("Marketing e redes sociais", 58.7), ("Capacitação de mão de obra", 45.7),
                 ("Vendas e estratégias comerciais", 43.5), ("Gestão administrativa", 23.9), ("Gestão financeira", 21.7),
                 ("Aspectos jurídicos/regularização", 15.2)]:
        chk("Gráfico X (alim. necessidades)", f"Necessidade alim.: {k}", v, "pct", tk(NE, k), A)
    for dname, v in zip(dias, [54.3, 56.5, 58.7, 60.9, 84.8, 93.5, 84.8]):
        chk("Gráfico X (alim. dias)", f"Alim. abre {dname}", v, "pct", lambda d, dn=dname: int(d[DD].map(lambda x: dn in parse_dias(x)).sum()), A)
    chk("Tabela X (alim. 24h)", "Alim. 24h: Sim", 4, "n", eq(C[23], "Sim"), A)
    chk("Tabela X (alim. feriados)", "Alim. feriados: Sim", 39, "n", eq(C[24], "Sim"), A)
    chk("Tabela X (alim. público)", "Alim. público: Visitantes ES", 32, "n", eq(C[26], "Visitantes do Espírito Santo"), A)
    chk("Tabela X (alim. público)", "Alim. público: Moradores", 14, "n", eq(C[26], "Moradores locais"), A)
    chk("Tabela X (alim. público)", "Alim. público: outros estados (PDF omite)", 0, "n", eq(C[26], "Turistas de outros estados brasileiros"), A)
    for k, v in [("Pix", 100.0), ("Dinheiro", 97.8), ("Cartão de débito", 73.9), ("Cartão de crédito", 73.9), ("Cheque", 2.2)]:
        chk("Gráfico X (alim. pagamento)", f"Pagamento alim.: {k}", v, "pct", tk(PG, k), A)
    for k, v in [("Rampa", 54.3), ("Estacionamento", 39.1), ("Não possui", 34.8), ("Corrimão", 15.2)]:
        chk("Gráfico X (alim. acessibilidade)", f"Acessibilidade alim.: {k}", v, "pct", tk(AC, k), A)
    chk("Tabela X (alim. sanitário)", "Sanitário alim.: Não", 37, "n", tk(SA, "Não"), A)

    # ---- Hospedagem
    H = segm(1)
    for k, v in [("Sem alimentação", 43.9), ("Café da manhã", 31.7), ("Meia pensão", 14.6), ("Pensão Completa", 19.5), ("Tudo incluído", 9.8)]:
        chk("Gráfico 16", f"Diária: {k}", v, "pct", tk(C[44], k), H)
    for k, v in [("Telefone/Whatsapp", 100.0), ("Internet (Site próprio, Booking, Trivago, etc)", 51.2), ("Balcão", 17.1), ("Agência de turismo", 4.9)]:
        chk("Gráfico 17", f"Reserva: {k}", v, "pct", tk(C[45], k), H)
    chk("Tabela 11", "Funciona os 12 meses", 40, "n", lambda d: int((d[C[46:58]].map(clean) == "Sim").all(axis=1).sum()), H)
    eqp = {"Internet/ wifi": 97.6, "frigobar": 82.9, "Mesa": 82.9, "Ventilador": 82.9, "TV com canal aberto": 73.2, "Fogão": 70.7,
           "Poltrona": 70.7, "Panelas e louças": 68.3, "Micro-ondas": 56.1, "Ar-condicionado": 48.8, "Cama queen-size": 39.0,
           "hidromassagem": 19.5, "Lareira": 19.5, "Academia": 7.3, "DVD": 2.4}
    for k, v in eqp.items():
        c = [x for x in C if x.endswith(f"[{k}]")][0]
        chk("Gráfico 18", f"Equipamento: {k}", v, "pct", eq(c, "Sim"), H)
    UH = C[93]
    chk("Tabela 12 (UH adaptadas; título errado)", "UH adaptadas: não possui (0)", 27, "n", lambda d: int((d[UH].astype(str) == "0").sum()), H)
    chk("Tabela 12 (UH adaptadas; título errado)", "UH adaptadas: 'Sim' sem quantidade", 3, "n", lambda d: int((d[UH].astype(str) == "Sim").sum()), H)
    chk("Gráfico 19", "Oferece almoço/jantar: Sim", 34.1, "pct", eq(C[96], "Sim"), H)
    chk("Gráfico 20", "Aceita pets: Sim", 61.0, "pct", eq(C[97], "Sim"), H)
    # ---- Entretenimento / Agro / Compras / RN
    En = segm(2)
    chk("Gráfico 21", "Entretenimento: exige reserva", 63.6, "pct", eq(C[100], "Sim"), En)
    chk("Gráfico 22", "Entretenimento: possui guia", 54.5, "pct", eq(C[102], "Sim"), En)
    chk("Tabela 13", "Entretenimento: EPI Sim", 4, "n", eq(C[103], "Sim"), En)
    capE = lambda d: pd.to_numeric(d[C[101]], errors="coerce")
    for k, v, f in [("válidos", 8, lambda s: s.notna().sum()), ("média", 55.0, lambda s: round(s.mean(), 1)),
                    ("mediana", 45.0, lambda s: s.median()), ("Q1", 30.75, lambda s: s.quantile(.25)), ("Q3", 81.25, lambda s: s.quantile(.75))]:
        chk("Tabela 14", f"Capacidade entretenimento: {k}", v, "stat", lambda d, f=f: f(capE(d)), En)
    Ag = segm(3)
    chk("Gráfico 25", "Agro: espaço degustação Sim", 58.8, "pct", eq(C[107], "Sim"), Ag)
    chk("Tabela 16 (total impresso = 11)", "Agro: total de respostas", 11, "n", lambda d: int(d[C[108]].notna().sum()), Ag)
    chk("Gráfico 26", "Agro: turismo de experiência Sim", 23.5, "pct", eq(C[109], "Sim"), Ag)
    Co = segm(4)
    chk("Gráfico 27", "Compras: produtos típicos", 66.7, "pct", eq(C[111], "Produtos típicos/gastronômicos"), Co)
    chk("Gráfico 28", "Compras: embalagem lembrança Sim", 44.4, "pct", eq(C[112], "Sim"), Co)
    chk("Tabela 17", "Compras: produção própria", 13, "n", eq(C[113], "Produção própria"), Co)
    chk("Gráfico 29", "Compras: interesse projetos coletivos Sim", 50.0, "pct", eq(C[114], "Sim"), Co)
    RN = segm(6)
    for k, v in [("Cachoeira", 100.0), ("Mirante", 28.6), ("Pesque e pague", 14.3), ("Montanha/Pico/Pedra", 14.3)]:
        chk("Gráfico 30", f"Atrativo: {k}", v, "pct", tk(C[125], k), RN)
    chk("Tabela 18", "RN: em área pública = Não", 6, "n", eq(C[126], "Não"), RN)
    chk("Tabela 18", "RN: acesso por via pública = Parcialmente", 4, "n", eq(C[127], "Parcialmente"), RN)
    chk("Tabela 19", "RN: trilhas sinalizadas = Sim", 3, "n", eq(C[129], "Sim"), RN)
    chk("Tabela 19", "RN: normas de conduta = Sim", 4, "n", eq(C[131], "Sim"), RN)
    chk("Gráfico 31", "RN: acessível = Não", 71.4, "pct", eq(C[132], "Não"), RN)
    for k, v in [("Banheiros", 71.4), ("Áreas de piquenique", 57.1), ("Estacionamento", 57.1), ("Lanchonetes/bares", 42.9), ("Pontos de descanso", 28.6)]:
        chk("Gráfico 32", f"Estrutura RN: {k}", v, "pct", tk(C[130], k), RN)
    chk("Tabela 20", "RN: sazonalidade verão", 5, "n", lambda d: int(d[C[133]].map(lambda x: "verão" in str(x).lower()).sum()), RN)
    chk("Tabela 20", "RN: sazonalidade inverno", 3, "n", lambda d: int(d[C[133]].map(lambda x: "inverno" in str(x).lower()).sum()), RN)

    out = pd.DataFrame(rows)
    out.to_csv(out_csv, index=False, sep=";", encoding="utf-8-sig")
    print(out["status"].value_counts().to_string())
    print(out[out["status"] != "OK"].to_string(index=False))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    sys.dont_write_bytecode = True  # não criar __pycache__ na pasta do projeto
    sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.abspath(__file__)))
    main(sys.argv[1], sys.argv[2])
