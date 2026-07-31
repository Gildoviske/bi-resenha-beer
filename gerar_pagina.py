# -*- coding: utf-8 -*-
"""Gera o painel executivo do Resenha Beer a partir das planilhas do Drive.

Uso: python gerar_pagina.py  (ou clique duas vezes em publicar.bat)
Lê os arquivos em BASE (J:\\Meu Drive) e grava index.html nesta mesma pasta,
pronta para publicar no GitHub Pages.
"""
import html as html_lib
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import pandas as pd

BASE = Path(r"J:\Meu Drive")
OUT = Path(__file__).resolve().parent / "index.html"

# ícone da aba do navegador: caneca de chopp em âmbar sobre fundo escuro,
# no mesmo tom do painel (bar/copo com espuma e alça)
FAVICON_SVG = (
    "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'>"
    "<rect width='64' height='64' rx='14' fill='#1B140F'/>"
    "<path d='M40 26 q11 0 11 9 q0 9 -11 9' fill='none' stroke='#C97A2B' stroke-width='5' stroke-linecap='round'/>"
    "<rect x='16' y='28' width='24' height='22' rx='3' fill='#C97A2B'/>"
    "<rect x='16' y='20' width='24' height='10' rx='4' fill='#FBE3C1'/>"
    "<circle cx='23' cy='38' r='2' fill='#F0B563' opacity='0.7'/>"
    "<circle cx='31' cy='44' r='1.8' fill='#F0B563' opacity='0.7'/>"
    "</svg>"
)
FAVICON_HREF = "data:image/svg+xml," + quote(FAVICON_SVG)

MESES_PT = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO",
            "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]
MESES_ABR = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

FONTES = [
    "CLIENTES - RESENHA BEER.xls",
    "CONTAS A PAGAR - RESENHA BEER.xls",
    "CUSTOS X VALOR DE VENDA - RESENHA BEER.xlsx",
    "ESTOQUE - RESENHA BEER.xls",
    "FATURAMENTO 2025 - RESENHA BEER.xlsx",
    "FATURAMENTO 2026 - RESENHA BEER.xlsx",
    "FATURAMENTO MENSAL - RESENHA BEER.xls",
    "HORÁRIO DE PICO - RESENHA BEER.xls",
    "MARGEM OPERACIONAL - RESENHA BEER.xlsx",
    "MEIOS DE PAGAMENTO - RESENHA BEER.xls",
    "PLANEJAMENTO DE COMPRAS - RESENHA BEER.xlsx",
    "RANKING DE VENDAS POR CLIENTE - RESENHA BEER.xls",
    "RECEITAS E DESPESAS - RESENHA BEER.xls",
    "VENDAS POR PRODUTO - RESENHA BEER.xls",
]


def esc(v):
    return html_lib.escape(str(v), quote=False)


def brl(v):
    if v is None:
        return "R$ 0,00"
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(v, casas=1):
    return f"{v:.{casas}f}%".replace(".", ",")


def num_br(v):
    return f"{int(v):,}".replace(",", ".")


def mtime_str(nome):
    ts = (BASE / nome).stat().st_mtime
    return datetime.fromtimestamp(ts).strftime("%d/%m/%Y às %H:%M")


def path(nome):
    return BASE / nome


# ======================================================================
#  1) FATURAMENTO 2025 / 2026 / MÊS CORRENTE
# ======================================================================
df25 = pd.read_excel(path("FATURAMENTO 2025 - RESENHA BEER.xlsx"), sheet_name="FATURAMENTO")
m25 = df25[["FATURAMENTO MÊS A MÊS (2025)", "Valor (R$)", "Lucro Bruto (R$)"]].dropna()
m25.columns = ["mes", "faturamento", "lucro"]
m25_map = {r.mes: (r.faturamento, r.lucro) for r in m25.itertuples()}

df26 = pd.read_excel(path("FATURAMENTO 2026 - RESENHA BEER.xlsx"))
m26 = df26[["FATURAMENTO MÊS A MÊS (2026)", "Valor (R$)", "Lucro Bruto (R$)"]].dropna()
m26.columns = ["mes", "faturamento", "lucro"]
m26_map = {r.mes: (r.faturamento, r.lucro) for r in m26.itertuples()}

fat2025_total = sum(v[0] for v in m25_map.values())
lucro2025_total = sum(v[1] for v in m25_map.values())
vendas_hora_2025 = df25[["Hora", "Vendas"]].dropna()
vendas2025_total = int(vendas_hora_2025["Vendas"].sum())

meses_comuns = [m for m in MESES_PT if m in m25_map and m in m26_map]
fat26_ytd = sum(m26_map[m][0] for m in meses_comuns)
lucro26_ytd = sum(m26_map[m][1] for m in meses_comuns)
fat25_mesmo_periodo = sum(m25_map[m][0] for m in meses_comuns)
crescimento_h1 = ((fat26_ytd - fat25_mesmo_periodo) / fat25_mesmo_periodo * 100) if fat25_mesmo_periodo else 0

melhor_mes_25 = max(m25_map.items(), key=lambda kv: kv[1][0])
pior_mes_25 = min(m25_map.items(), key=lambda kv: kv[1][0])

yoy = []
for m in meses_comuns:
    v25 = m25_map[m][0]
    v26 = m26_map[m][0]
    yoy.append((m, (v26 - v25) / v25 * 100 if v25 else 0))
melhor_yoy = max(yoy, key=lambda kv: kv[1]) if yoy else (None, 0)
pior_yoy = min(yoy, key=lambda kv: kv[1]) if yoy else (None, 0)
meses_abaixo = [m for m, chg in yoy if chg < 0]

margem_media_mensal_25 = (lucro2025_total / fat2025_total * 100) if fat2025_total else 0

# ---- mês corrente (snapshot mais recente) ----
dfm = pd.read_excel(path("FATURAMENTO MENSAL - RESENHA BEER.xls"))
mes_atual_fat = float(dfm["Faturamento"].sum())
mes_atual_vendas = int(dfm["Vendas"].sum())
mes_atual_lucro = float(dfm["Lucro Bruto"].sum())

dfh = pd.read_excel(path("HORÁRIO DE PICO - RESENHA BEER.xls"))
dfh.columns = ["hora", "qtd", "valor"]
hora_valor = {int(r.hora): float(r.valor) for r in dfh.itertuples()}
for h in range(24):
    hora_valor.setdefault(h, 0.0)
max_hora_val = max(hora_valor.values()) if hora_valor else 0
pico_hora, pico_valor = max(hora_valor.items(), key=lambda kv: kv[1])
janela_pico = [h for h in range(24) if hora_valor[h] >= 0.55 * max_hora_val and max_hora_val > 0]
janela_pico_soma = sum(hora_valor[h] for h in janela_pico)
janela_pico_pct = (janela_pico_soma / mes_atual_fat * 100) if mes_atual_fat else 0
janela_ini, janela_fim = (min(janela_pico), max(janela_pico)) if janela_pico else (0, 0)

dfp = pd.read_excel(path("MEIOS DE PAGAMENTO - RESENHA BEER.xls"))
dfp = dfp[dfp.iloc[:, 0] != "TOTAL"]
meios_pagto = sorted(
    [(str(r[0]).lstrip("* ").strip(), float(r[1])) for r in dfp.itertuples(index=False) if pd.notna(r[1])],
    key=lambda kv: kv[1], reverse=True,
)
meios_total = sum(v for _, v in meios_pagto)
digital_pct = sum(v for n, v in meios_pagto if n in ("Cartão de Débito", "Pix")) / meios_total * 100 if meios_total else 0

# ---- receitas e despesas do mês ----
dfr = pd.read_excel(path("RECEITAS E DESPESAS - RESENHA BEER.xls"), header=None)
raw = dfr.fillna("").values.tolist()
receitas_total = despesas_total = resultado_liquido = 0.0
despesas_categorias = []
modo = None
for row in raw:
    label = str(row[0]).strip()
    valor = row[1]
    if label == "Despesas":
        modo = "despesas"
        continue
    if label == "Receitas":
        modo = "receitas"
        continue
    if label == "Total Receitas + Pendências":
        receitas_total = float(valor)
    elif label == "Total Despesas":
        despesas_total = float(valor)
    elif label == "Receitas - Despesas":
        resultado_liquido = float(valor)
    elif modo == "despesas" and label and isinstance(valor, (int, float)) and valor != "":
        despesas_categorias.append((label.title(), float(valor)))
despesas_categorias = [d for d in despesas_categorias if d[1] > 0]
despesas_categorias.sort(key=lambda kv: kv[1], reverse=True)
maior_despesa_cat, maior_despesa_val = despesas_categorias[0] if despesas_categorias else ("-", 0)
maior_despesa_pct = (maior_despesa_val / despesas_total * 100) if despesas_total else 0

# ======================================================================
#  2) MARGEM OPERACIONAL
# ======================================================================
dfmg = pd.read_excel(path("MARGEM OPERACIONAL - RESENHA BEER.xlsx"))
margem_n = len(dfmg)
margem_media = float(dfmg["MARGEM OPERACIONAL"].mean()) * 100
class_counts = dfmg["Classificação de Margem"].value_counts().to_dict()
n_alta = class_counts.get("Alta", 0)
n_media = class_counts.get("Média", 0)
n_baixa = class_counts.get("Baixa", 0)
pct_alta = n_alta / margem_n * 100 if margem_n else 0

ajustar = dfmg[dfmg["Ajustar Preço?"] == "SIM"].sort_values("MARGEM OPERACIONAL")
top_margem_reais = dfmg.sort_values("Margem em Reais", ascending=False).head(5)

# ======================================================================
#  3) CARDÁPIO DE COPÕES (fichas de custo)
# ======================================================================
xls_custo = pd.ExcelFile(path("CUSTOS X VALOR DE VENDA - RESENHA BEER.xlsx"))
cocktails = []
nome_contagem = {}
for sheet in xls_custo.sheet_names:
    dfc = xls_custo.parse(sheet, header=None)
    nome_base = sheet.title()
    for i, row in dfc.iterrows():
        for c_custo, c_preco, c_lucro, c_margem in [(3, 4, 5, 6), (11, 12, 13, 14)]:
            try:
                custo, preco, lucro, margem = row[c_custo], row[c_preco], row[c_lucro], row[c_margem]
            except (KeyError, IndexError):
                continue
            if pd.notna(preco) and pd.notna(lucro) and isinstance(preco, (int, float)) and isinstance(lucro, (int, float)):
                # o título da receita fica 1 linha acima do cabeçalho "Ingredientes",
                # na mesma coluna (3 colunas à esquerda da coluna de custo)
                col_titulo = max(c_custo - 3, 0)
                nome_receita = nome_base
                header_idx = None
                for j in range(i - 1, max(i - 15, -1), -1):
                    try:
                        v = dfc.iat[j, col_titulo]
                    except IndexError:
                        continue
                    if isinstance(v, str) and v.strip() == "Ingredientes":
                        header_idx = j
                        break
                if header_idx is not None and header_idx > 0:
                    v = dfc.iat[header_idx - 1, col_titulo]
                    if isinstance(v, str) and v.strip():
                        nome_receita = (
                            v.strip().title()
                            .replace(" De ", " de ").replace(" Com ", " com ")
                            .replace(" S/", " s/").replace(" C/", " c/")
                        )
                if nome_receita in nome_contagem:
                    nome_contagem[nome_receita] += 1
                    nome_exibido = f"{nome_receita} ({nome_contagem[nome_receita]})"
                else:
                    nome_contagem[nome_receita] = 1
                    nome_exibido = nome_receita
                cocktails.append({
                    "nome": nome_exibido, "custo": float(custo) if pd.notna(custo) else 0.0,
                    "preco": float(preco), "lucro": float(lucro),
                    "margem": float(margem) * 100 if pd.notna(margem) else 0.0,
                })
cocktails.sort(key=lambda c: c["margem"], reverse=True)
pior_copao = min(cocktails, key=lambda c: c["margem"]) if cocktails else None
melhor_copao_familia = cocktails[0]["nome"] if cocktails else "-"

# ======================================================================
#  4) ESTOQUE
# ======================================================================
dfe = pd.read_excel(path("ESTOQUE - RESENHA BEER.xls"))
estoque_n = len(dfe)
estoque_custo_total = float(dfe["Custo Total"].sum())
estoque_venda_total = float(dfe["Preço Total"].sum())
estoque_zerados = dfe[dfe["Estoque Atual"] == 0]
estoque_zerados_n = len(estoque_zerados)
estoque_zerados_pct = estoque_zerados_n / estoque_n * 100 if estoque_n else 0
estoque_por_cat = dfe.groupby("Categoria")["Custo Total"].sum().sort_values(ascending=False).head(6)

# ======================================================================
#  5) PLANEJAMENTO DE COMPRAS
# ======================================================================
dfpl = pd.read_excel(path("PLANEJAMENTO DE COMPRAS - RESENHA BEER.xlsx"), sheet_name="PLANEJAMENTO", header=1)
dfpl = dfpl[dfpl["Produto"].notna()]
planej_n = len(dfpl)
status_counts = dfpl["Status"].value_counts().to_dict()
n_critico = status_counts.get("CRÍTICO", 0)
n_atencao = status_counts.get("ATENÇÃO", 0)
n_ok = status_counts.get("OK", 0)
n_excesso = status_counts.get("EXCESSO", 0)
n_semgiro = status_counts.get("SEM GIRO", 0)
criticos_todos = dfpl[dfpl["Status"] == "CRÍTICO"]["Produto"].tolist()
criticos_exemplos = criticos_todos[:6]

# item crítico que também vende bem (cruza com top produtos, calculado abaixo)

# ======================================================================
#  6) VENDAS POR PRODUTO / CLIENTES
# ======================================================================
dfvp = pd.read_excel(path("VENDAS POR PRODUTO - RESENHA BEER.xls"))
dfvp = dfvp[dfvp["Nome"].notna()]
top_produtos = dfvp.sort_values("Qtd.", ascending=False).head(9)
top_produtos_nomes = set(top_produtos["Nome"].str.upper())
criticos_alto_giro = [p for p in criticos_todos if str(p).upper() in top_produtos_nomes]

dfrk = pd.read_excel(path("RANKING DE VENDAS POR CLIENTE - RESENHA BEER.xls"))
dfrk.columns = ["nome", "valor"]
ranking_clientes_n = len(dfrk)
ranking_clientes_total = float(dfrk["valor"].sum())
top10_clientes = dfrk.sort_values("valor", ascending=False).head(10)
top10_soma = float(top10_clientes["valor"].sum())
top10_concentracao = top10_soma / ranking_clientes_total * 100 if ranking_clientes_total else 0
top7_clientes = list(top10_clientes.head(7).itertuples(index=False))

dfcli = pd.read_excel(path("CLIENTES - RESENHA BEER.xls"))
clientes_ativos = int((dfcli["Status"] == "Ativo").sum()) if "Status" in dfcli.columns else len(dfcli)


def parse_saldo_cliente(texto):
    """Extrai valores de 'Débito- R$ X,XX' / 'Crédito R$ X,XX' (podem vir combinados
    na mesma célula, separados por quebra de linha)."""
    if not isinstance(texto, str):
        return 0.0, 0.0

    def to_float(s):
        return float(s.replace(".", "").replace(",", "."))

    deb = sum(to_float(x) for x in re.findall(r"D[ée]bito-?\s*R\$\s*([\d.,]+)", texto))
    cred = sum(to_float(x) for x in re.findall(r"Cr[ée]dito-?\s*R\$\s*([\d.,]+)", texto))
    return deb, cred


col_saldo = "Débito / Crédito"
if col_saldo in dfcli.columns:
    pares = dfcli[col_saldo].map(parse_saldo_cliente).tolist()
else:
    pares = [(0.0, 0.0)] * len(dfcli)
dfcli["_debito"], dfcli["_credito"] = zip(*pares) if pares else ([], [])
dfcli["_saldo"] = dfcli["_debito"] - dfcli["_credito"]

devedores = dfcli[dfcli["_saldo"] > 0].sort_values("_saldo", ascending=False)
n_devedores = len(devedores)
total_fiado = float(devedores["_saldo"].sum())
maior_devedor = devedores.iloc[0] if len(devedores) else None

# ======================================================================
#  7) CONTAS A PAGAR
# ======================================================================
dfcp = pd.read_excel(path("CONTAS A PAGAR - RESENHA BEER.xls"))
status_cp = dfcp["Status"].value_counts().to_dict()
n_pagas = status_cp.get("Paga", 0)
venc = dfcp[dfcp["Status"] == "Vencida"]
n_vencidas, v_vencidas = len(venc), float(venc["Valor"].sum())
a_vencer = dfcp[dfcp["Status"] == "A Vencer"]
n_a_vencer, v_a_vencer = len(a_vencer), float(a_vencer["Valor"].sum())
n_vence_hoje = status_cp.get("Vence hoje", 0)

STATUS_PENDENTES = ["Vencida", "Vence hoje", "A Vencer"]
contas_pendentes = (
    dfcp[dfcp["Status"].isin(STATUS_PENDENTES) & dfcp["Fornecedor"].notna()]
    .sort_values("Vencimento")
)
total_pendente = float(contas_pendentes["Valor"].sum())

contas_por_cat = dfcp.groupby("Categoria")["Valor"].sum().sort_values(ascending=False).head(6)
contas_por_forn = dfcp.groupby("Fornecedor")["Valor"].sum().sort_values(ascending=False).head(6)
contas_total_hist = float(dfcp["Valor"].sum())
bebidas_pct_hist = (contas_por_cat.get("BEBIDAS", 0) / contas_total_hist * 100) if contas_total_hist else 0
top_fornecedor_nome = contas_por_forn.index[0] if len(contas_por_forn) else "-"

# ======================================================================
#  DATAS / METADADOS
# ======================================================================
agora = datetime.now()
mes_atual_nome = MESES_ABR[agora.month - 1]
ultimo_mes26_idx = max(MESES_PT.index(m) for m in m26_map) if m26_map else -1
primeiro_mes25_idx = min(MESES_PT.index(m) for m in m25_map) if m25_map else 0
periodo_label = (
    f"{MESES_ABR[primeiro_mes25_idx]}/2025 – {MESES_ABR[ultimo_mes26_idx]}/2026 fechado, "
    f"{mes_atual_nome}/{agora.year} em andamento"
)
gerado_em = agora.strftime("%d/%m/%Y às %H:%M")


# ======================================================================
#  HELPERS DE HTML
# ======================================================================
def hbar_rows(items, alt=False):
    if not items:
        return '<div class="hbar-row"><span class="name">Sem dados</span></div>'
    maxv = max(v for _, v in items) or 1
    cls = "hbar alt" if alt else "hbar"
    out = []
    for label, val in items:
        w = val / maxv * 100
        out.append(
            f'<div class="hbar-row"><span class="name">{esc(label)}</span>'
            f'<div class="hbar-track"><div class="{cls}" style="width:{w:.1f}%"></div></div>'
            f'<span class="val">{brl(val)}</span></div>'
        )
    return "\n".join(out)


def margem_tag(m):
    if m >= 40:
        return "high"
    if m >= 20:
        return "mid"
    return "low"


def month_bar(label, v25, v26, escala):
    p25 = (v25 / escala * 100) if v25 is not None else None
    p26 = (v26 / escala * 100) if v26 is not None else None
    bars = ""
    if p25 is not None:
        bars += f'<div class="bar b25" style="height:{p25:.1f}%" data-tip="2025: {brl(v25)}"></div>'
    if p26 is not None:
        bars += f'<div class="bar b26" style="height:{p26:.1f}%" data-tip="2026: {brl(v26)}"></div>'
    if not bars:
        bars = '<div class="bar b25" style="height:0%" data-tip="sem dados"></div>'
    return f'<div class="month-col"><div class="bar-pair">{bars}</div><div class="month-label">{label}</div></div>'


def hour_bar(h):
    val = hora_valor.get(h, 0.0)
    height = (val / max_hora_val * 100) if max_hora_val else 0
    peak = " peak" if h in janela_pico else ""
    tick = {0: "0h", 6: "6h", 12: "12h", 18: "18h", 23: "23h"}.get(h, "")
    tip = f"{brl(val)} — pico do dia" if h == pico_hora else brl(val)
    return (
        f'<div class="hour-col"><div class="col-bar{peak}" style="height:{height:.1f}%" '
        f'data-tip="{tip}"></div><div class="hour-tick">{tick}</div></div>'
    )


# ======================================================================
#  MONTAGEM DAS SEÇÕES
# ======================================================================
escala_mes = max([v[0] for v in m25_map.values()] + [v[0] for v in m26_map.values()] + [1]) * 1.07
month_chart_html = "\n".join(
    month_bar(MESES_ABR[i], m25_map.get(m, (None, None))[0], m26_map.get(m, (None, None))[0], escala_mes)
    for i, m in enumerate(MESES_PT)
)
hour_chart_html = "\n".join(hour_bar(h) for h in range(24))

# monta a tabela de produtos manualmente (acesso por nome de coluna, robusto a reordenação)
rows_html = []
for i, r in enumerate(top_produtos.to_dict("records"), start=1):
    rows_html.append(
        f'<tr><td class="rank">{i}</td><td>{esc(r["Nome"])}</td>'
        f'<td class="num">{int(r["Qtd."])}</td><td class="num">{brl(r["Total (R$)"])}</td></tr>'
    )
top_produtos_rows = "\n".join(rows_html)

ajustar_rows = "\n".join(
    f'<tr><td>{esc(r["Produto"])}</td><td class="num">{brl(r["Custo"])}</td>'
    f'<td class="num">{brl(r["Valor de venda"])}</td><td class="num"><span class="tag low">{pct(r["MARGEM OPERACIONAL"] * 100)}</span></td></tr>'
    for r in ajustar.to_dict("records")
) or '<tr><td colspan="4">Nenhum produto sinalizado no momento.</td></tr>'

cocktail_rows = "\n".join(
    f'<tr><td>{esc(c["nome"])}</td><td class="num">{brl(c["custo"])}</td><td class="num">{brl(c["preco"])}</td>'
    f'<td class="num">{brl(c["lucro"])}</td><td class="num"><span class="tag {margem_tag(c["margem"])}">{pct(c["margem"])}</span></td></tr>'
    for c in cocktails
)

criticos_rows = "\n".join(f"<tr><td>{esc(p)}</td></tr>" for p in criticos_exemplos) or "<tr><td>Nenhum item crítico.</td></tr>"

STATUS_TAG = {"Vencida": "low", "Vence hoje": "mid", "A Vencer": "muted"}
contas_pendentes_rows = "\n".join(
    f'<tr><td><span class="tag {STATUS_TAG.get(r["Status"], "muted")}">{esc(r["Status"])}</span></td>'
    f'<td>{r["Vencimento"].strftime("%d/%m/%Y") if pd.notna(r["Vencimento"]) else "-"}</td>'
    f'<td>{esc(r["Fornecedor"])}</td><td>{esc(r["Categoria"]) if pd.notna(r["Categoria"]) else "-"}</td>'
    f'<td>{esc(r["Referente a"]) if pd.notna(r["Referente a"]) else "-"}</td>'
    f'<td class="num strong">{brl(r["Valor"])}</td></tr>'
    for r in contas_pendentes.to_dict("records")
) or '<tr><td colspan="6">Nenhuma conta pendente no momento — tudo em dia.</td></tr>'

devedores_rows = "\n".join(
    f'<tr><td>{esc(r["Nome"])}</td><td class="num">{brl(r["_debito"])}</td>'
    f'<td class="num">{brl(r["_credito"]) if r["_credito"] else "-"}</td>'
    f'<td class="num strong">{brl(r["_saldo"])}</td></tr>'
    for r in devedores.to_dict("records")
) or '<tr><td colspan="4">Nenhum cliente com saldo em aberto.</td></tr>'

top10_clientes_hbar = hbar_rows([(r.nome, r.valor) for r in top7_clientes])
top_margem_hbar = hbar_rows([(r["Produto"], r["Margem em Reais"]) for r in top_margem_reais.to_dict("records")])
estoque_cat_hbar = hbar_rows(list(estoque_por_cat.items()))
contas_cat_hbar = hbar_rows(list(contas_por_cat.items()))
contas_forn_hbar = hbar_rows(list(contas_por_forn.items()), alt=True)
meios_pagto_hbar = hbar_rows(meios_pagto)

fontes_footer = " · ".join(f"{f} ({mtime_str(f)})" for f in FONTES if (BASE / f).exists())


# ======================================================================
#  CSS (estático)
# ======================================================================
CSS = """
:root{
  --ink-950:#1B140F; --ink-900:#241B14; --ink-800:#2E2119; --ink-050:#FBF6EE;
  --surface:#FFFFFF; --surface-alt:#F3ECDF; --line:rgba(36,27,20,0.12);
  --amber-700:#A8631C; --amber-600:#C97A2B; --amber-300:#F0B563; --amber-100:#FBE3C1;
  --teal-600:#2B7A6F; --teal-300:#7FBDB4;
  --good:#3E9A5D; --good-bg:rgba(62,154,93,0.13);
  --warning:#C98F1F; --warning-bg:rgba(201,143,31,0.15);
  --critical:#C4463B; --critical-bg:rgba(196,70,59,0.13);
  --neutral-info:#8A7B6C; --neutral-info-bg:rgba(138,123,108,0.14);
  --bg: var(--ink-050); --card: var(--surface); --card-alt: var(--surface-alt);
  --text-primary: var(--ink-900); --text-secondary: rgba(36,27,20,0.72); --text-muted: rgba(36,27,20,0.52);
  --accent: var(--amber-600); --accent-strong: var(--amber-700); --accent-soft: var(--amber-100);
  --series-b: var(--teal-600); --track: rgba(36,27,20,0.08);
  --shadow: 0 1px 2px rgba(27,20,15,0.06), 0 8px 24px -12px rgba(27,20,15,0.18);
  --font-display: Georgia, 'Iowan Old Style', 'Palatino Linotype', 'Times New Roman', serif;
  --font-body: -apple-system, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
}
@media (prefers-color-scheme: dark){
  :root{
    --bg: var(--ink-950); --card: var(--ink-900); --card-alt: var(--ink-800);
    --text-primary: #F3E9DA; --text-secondary: rgba(243,233,218,0.72); --text-muted: rgba(243,233,218,0.5);
    --line: rgba(243,233,218,0.12); --accent: var(--amber-300); --accent-strong: var(--amber-300);
    --accent-soft: rgba(240,181,99,0.16); --series-b: var(--teal-300); --track: rgba(243,233,218,0.1);
    --good-bg:rgba(62,154,93,0.18); --warning-bg:rgba(201,143,31,0.2); --critical-bg:rgba(196,70,59,0.2);
    --neutral-info-bg:rgba(138,123,108,0.2); --shadow: 0 1px 2px rgba(0,0,0,0.3), 0 12px 32px -14px rgba(0,0,0,0.55);
  }
}
:root[data-theme="dark"]{
  --bg: var(--ink-950); --card: var(--ink-900); --card-alt: var(--ink-800);
  --text-primary: #F3E9DA; --text-secondary: rgba(243,233,218,0.72); --text-muted: rgba(243,233,218,0.5);
  --line: rgba(243,233,218,0.12); --accent: var(--amber-300); --accent-strong: var(--amber-300);
  --accent-soft: rgba(240,181,99,0.16); --series-b: var(--teal-300); --track: rgba(243,233,218,0.1);
  --good-bg:rgba(62,154,93,0.18); --warning-bg:rgba(201,143,31,0.2); --critical-bg:rgba(196,70,59,0.2);
  --neutral-info-bg:rgba(138,123,108,0.2); --shadow: 0 1px 2px rgba(0,0,0,0.3), 0 12px 32px -14px rgba(0,0,0,0.55);
}
:root[data-theme="light"]{
  --bg: var(--ink-050); --card: var(--surface); --card-alt: var(--surface-alt);
  --text-primary: var(--ink-900); --text-secondary: rgba(36,27,20,0.72); --text-muted: rgba(36,27,20,0.52);
  --line: rgba(36,27,20,0.12); --accent: var(--amber-600); --accent-strong: var(--amber-700);
  --accent-soft: var(--amber-100); --series-b: var(--teal-600); --track: rgba(36,27,20,0.08);
  --shadow: 0 1px 2px rgba(27,20,15,0.06), 0 8px 24px -12px rgba(27,20,15,0.18);
}
*{box-sizing:border-box;}
html,body{margin:0;padding:0;}
body{background:var(--bg); color:var(--text-primary); font-family:var(--font-body); line-height:1.5; -webkit-font-smoothing:antialiased;}
@media (prefers-reduced-motion: reduce){*{animation-duration:0.001ms !important; transition-duration:0.001ms !important;}}
.wrap{max-width:1180px; margin:0 auto; padding:0 24px 80px;}
header.hero{padding:56px 24px 40px; border-bottom:1px solid var(--line); background: radial-gradient(1200px 260px at 15% -10%, var(--accent-soft), transparent 60%);}
header.hero .wrap{padding:0; display:flex; flex-direction:column; gap:14px; max-width:1180px; margin:0 auto; padding-left:24px; padding-right:24px;}
.eyebrow{text-transform:uppercase; letter-spacing:0.12em; font-size:0.72rem; color:var(--accent-strong); font-weight:700;}
h1.brand{font-family:var(--font-display); font-size:clamp(2.1rem, 4vw, 3rem); margin:0; font-weight:700; text-wrap:balance; letter-spacing:-0.01em;}
.hero-sub{color:var(--text-secondary); font-size:1.02rem; max-width:640px;}
.hero-meta{display:flex; gap:18px; flex-wrap:wrap; margin-top:8px; font-size:0.85rem; color:var(--text-muted);}
.hero-meta strong{color:var(--text-primary);}
nav.section-nav{position:sticky; top:0; z-index:20; background:color-mix(in srgb, var(--bg) 88%, transparent); backdrop-filter: blur(8px); border-bottom:1px solid var(--line);}
nav.section-nav .nav-inner{max-width:1180px; margin:0 auto; padding:0 24px; display:flex; gap:4px; overflow-x:auto; scrollbar-width:none;}
nav.section-nav .nav-inner::-webkit-scrollbar{display:none;}
nav.section-nav a{white-space:nowrap; padding:14px 12px; font-size:0.82rem; font-weight:600; color:var(--text-muted); text-decoration:none; border-bottom:2px solid transparent;}
nav.section-nav a:hover{color:var(--text-primary);}
nav.section-nav a.active{color:var(--accent-strong); border-bottom-color:var(--accent-strong);}
section.block{padding:52px 0 8px; border-bottom:1px solid var(--line); scroll-margin-top:64px;}
section.block:last-of-type{border-bottom:none;}
.block-head{display:flex; align-items:baseline; justify-content:space-between; gap:16px; flex-wrap:wrap; margin-bottom:22px;}
.block-head h2{font-family:var(--font-display); font-size:1.6rem; margin:0; font-weight:700; text-wrap:balance;}
.block-num{font-size:0.78rem; color:var(--text-muted); font-weight:600; letter-spacing:0.06em;}
.stat-grid{display:grid; grid-template-columns:repeat(auto-fit, minmax(210px,1fr)); gap:14px; margin-bottom:24px;}
.stat{background:var(--card); border:1px solid var(--line); border-radius:10px; padding:18px 18px 16px; box-shadow:var(--shadow);}
.stat .label{font-size:0.76rem; color:var(--text-muted); text-transform:uppercase; letter-spacing:0.06em; font-weight:700;}
.stat .value{font-family:var(--font-display); font-size:1.7rem; font-weight:700; margin-top:6px; font-variant-numeric:tabular-nums;}
.stat .delta{margin-top:6px; font-size:0.82rem; font-weight:600;}
.delta.up{color:var(--good);} .delta.down{color:var(--critical);} .delta.flat{color:var(--text-muted);}
.insight{background:var(--card-alt); border:1px solid var(--line); border-left:4px solid var(--accent-strong); border-radius:8px; padding:16px 20px; margin:18px 0 30px;}
.insight h3{margin:0 0 8px; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.08em; color:var(--accent-strong); font-weight:800;}
.insight ul{margin:0; padding-left:18px; display:flex; flex-direction:column; gap:6px;}
.insight li{font-size:0.92rem; color:var(--text-secondary);}
.insight li b, .insight li strong{color:var(--text-primary);}
.panel{background:var(--card); border:1px solid var(--line); border-radius:12px; padding:22px; box-shadow:var(--shadow); margin-bottom:20px;}
.panel h3.panel-title{margin:0 0 4px; font-size:1.02rem; font-weight:700;}
.panel .panel-sub{font-size:0.82rem; color:var(--text-muted); margin-bottom:18px;}
.two-col{display:grid; grid-template-columns:1.3fr 1fr; gap:20px;}
@media (max-width:860px){.two-col{grid-template-columns:1fr;}}
.chart-scroll{overflow-x:auto;}
.month-chart{display:flex; align-items:flex-end; gap:10px; height:210px; min-width:640px; padding-top:8px;}
.month-col{flex:1; display:flex; flex-direction:column; align-items:center; justify-content:flex-end; height:100%;}
.bar-pair{display:flex; align-items:flex-end; gap:3px; width:100%; height:100%; justify-content:center;}
.bar-pair .bar{width:13px; border-radius:3px 3px 0 0; position:relative; cursor:default;}
.bar.b25{background:var(--accent);} .bar.b26{background:var(--series-b);}
.bar[data-tip]:hover::after, .hbar[data-tip]:hover::after, .col-bar[data-tip]:hover::after{
  content:attr(data-tip); position:absolute; bottom:calc(100% + 6px); left:50%; transform:translateX(-50%);
  background:var(--ink-900); color:#F3E9DA; font-size:0.72rem; font-weight:600; padding:4px 8px; border-radius:6px;
  white-space:nowrap; z-index:5; box-shadow:var(--shadow); font-variant-numeric:tabular-nums;
}
.month-label{margin-top:8px; font-size:0.68rem; color:var(--text-muted); font-weight:600;}
.legend{display:flex; gap:18px; margin-top:14px; font-size:0.8rem; color:var(--text-secondary);}
.legend span{display:inline-flex; align-items:center; gap:6px;}
.legend i{width:10px; height:10px; border-radius:3px; display:inline-block;}
.hour-chart{display:flex; align-items:flex-end; gap:3px; height:170px; padding-top:8px;}
.hour-col{flex:1; display:flex; flex-direction:column; align-items:center; justify-content:flex-end; height:100%; position:relative;}
.col-bar{width:100%; max-width:22px; border-radius:3px 3px 0 0; background:var(--track); position:relative;}
.col-bar.peak{background:var(--accent);}
.hour-tick{margin-top:6px; font-size:0.62rem; color:var(--text-muted);}
.hbar-list{display:flex; flex-direction:column; gap:11px;}
.hbar-row{display:grid; grid-template-columns:150px 1fr 84px; align-items:center; gap:10px;}
.hbar-row .name{font-size:0.85rem; color:var(--text-secondary); white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
.hbar-track{background:var(--track); border-radius:5px; height:12px; position:relative; overflow:visible;}
.hbar{position:absolute; left:0; top:0; height:100%; border-radius:5px; background:var(--accent);}
.hbar.alt{background:var(--series-b);}
.hbar-row .val{font-size:0.83rem; font-weight:700; text-align:right; font-variant-numeric:tabular-nums;}
.chip-row{display:flex; flex-wrap:wrap; gap:8px; margin-bottom:6px;}
.chip{display:inline-flex; align-items:center; gap:7px; padding:7px 12px; border-radius:999px; font-size:0.8rem; font-weight:700; border:1px solid transparent;}
.chip .dot{width:8px; height:8px; border-radius:50%;}
.chip.critical{background:var(--critical-bg); color:var(--critical);} .chip.critical .dot{background:var(--critical);}
.chip.warning{background:var(--warning-bg); color:var(--warning);} .chip.warning .dot{background:var(--warning);}
.chip.good{background:var(--good-bg); color:var(--good);} .chip.good .dot{background:var(--good);}
.chip.neutral{background:var(--neutral-info-bg); color:var(--neutral-info);} .chip.neutral .dot{background:var(--neutral-info);}
.table-scroll{overflow-x:auto;}
table{width:100%; border-collapse:collapse; font-size:0.87rem;}
thead th{text-align:left; font-size:0.7rem; text-transform:uppercase; letter-spacing:0.05em; color:var(--text-muted); padding:0 10px 10px; border-bottom:1px solid var(--line); white-space:nowrap;}
tbody td{padding:9px 10px; border-bottom:1px solid var(--line); color:var(--text-secondary); white-space:nowrap;}
tbody tr:last-child td{border-bottom:none;}
td.num, th.num{text-align:right; font-variant-numeric:tabular-nums;}
td.strong{color:var(--text-primary); font-weight:600;}
.rank{color:var(--text-muted); font-variant-numeric:tabular-nums;}
.tag{display:inline-block; padding:2px 8px; border-radius:6px; font-size:0.72rem; font-weight:700;}
.tag.low{background:var(--critical-bg); color:var(--critical);}
.tag.mid{background:var(--warning-bg); color:var(--warning);}
.tag.high{background:var(--good-bg); color:var(--good);}
.tag.muted{background:var(--neutral-info-bg); color:var(--neutral-info);}
.grid-2{display:grid; grid-template-columns:1fr 1fr; gap:20px;}
@media (max-width:860px){.grid-2{grid-template-columns:1fr;}}
.grid-3{display:grid; grid-template-columns:repeat(3,1fr); gap:14px;}
@media (max-width:860px){.grid-3{grid-template-columns:1fr;}}
footer{padding:36px 0 20px; text-align:center; color:var(--text-muted); font-size:0.78rem;}
footer a{color:var(--text-muted);}
"""

SCRIPT = """
(function(){
  var links = Array.prototype.slice.call(document.querySelectorAll('.section-nav a'));
  var sections = links.map(function(a){ return document.querySelector(a.getAttribute('href')); });
  function onScroll(){
    var pos = window.scrollY + 90;
    var current = sections[0];
    sections.forEach(function(sec){ if(sec && sec.offsetTop <= pos) current = sec; });
    links.forEach(function(a){ a.classList.toggle('active', current && a.getAttribute('href') === '#' + current.id); });
  }
  document.addEventListener('scroll', onScroll, {passive:true});
  onScroll();
})();
"""

sec_visao = f"""
  <section class="block" id="visao-geral">
    <div class="block-head"><h2>Visão geral</h2><span class="block-num">01 / 07</span></div>
    <div class="stat-grid">
      <div class="stat">
        <div class="label">Faturamento 2025</div>
        <div class="value">{brl(fat2025_total)}</div>
        <div class="delta flat">{num_br(vendas2025_total)} vendas no ano</div>
      </div>
      <div class="stat">
        <div class="label">Faturamento acumulado 2026</div>
        <div class="value">{brl(fat26_ytd)}</div>
        <div class="delta {'up' if crescimento_h1 >= 0 else 'down'}">{'▲' if crescimento_h1 >= 0 else '▼'} {pct(abs(crescimento_h1))} vs. mesmo período de 2025</div>
      </div>
      <div class="stat">
        <div class="label">{mes_atual_nome}/{agora.year} (parcial, snapshot mais recente)</div>
        <div class="value">{brl(mes_atual_fat)}</div>
        <div class="delta flat">{brl(resultado_liquido)} líquido após despesas</div>
      </div>
      <div class="stat">
        <div class="label">Lucro bruto 2025</div>
        <div class="value">{brl(lucro2025_total)}</div>
        <div class="delta flat">{pct(margem_media_mensal_25)} sobre o faturamento</div>
      </div>
    </div>
    <div class="insight">
      <h3>Pontos importantes</h3>
      <ul>
        <li>O período acumulado de 2026 fechou <b>{pct(abs(crescimento_h1))} {'acima' if crescimento_h1 >= 0 else 'abaixo'}</b> do mesmo período de 2025 (meses com dado em ambos os anos).</li>
        <li><b>{melhor_yoy[0] and melhor_yoy[0].title()}</b> teve a melhor evolução ano contra ano (<b>{'+' if melhor_yoy[1] >= 0 else ''}{pct(melhor_yoy[1])}</b>); <b>{pior_yoy[0] and pior_yoy[0].title()}</b> teve a pior (<b>{pct(pior_yoy[1])}</b>).</li>
        <li><b>{melhor_mes_25[0].title()}/2025</b> foi o melhor mês do ano fechado ({brl(melhor_mes_25[1][0])}); <b>{pior_mes_25[0].title()}/2025</b> foi o mais fraco ({brl(pior_mes_25[1][0])}).</li>
        <li>{mes_atual_nome}/{agora.year} está em <b>{brl(mes_atual_fat)}</b> faturados até o momento da última atualização desta página.</li>
      </ul>
    </div>
  </section>
"""

sec_faturamento = f"""
  <section class="block" id="faturamento">
    <div class="block-head"><h2>Faturamento &amp; lucro mês a mês</h2><span class="block-num">02 / 07</span></div>
    <div class="panel">
      <h3 class="panel-title">2025 vs. 2026, mês a mês</h3>
      <div class="panel-sub">Faturamento em R$ · passe o mouse sobre uma barra para ver o valor exato</div>
      <div class="chart-scroll"><div class="month-chart">{month_chart_html}</div></div>
      <div class="legend"><span><i style="background:var(--accent)"></i>2025</span><span><i style="background:var(--series-b)"></i>2026</span></div>
    </div>
    <div class="insight">
      <h3>Pontos importantes</h3>
      <ul>
        <li><b>{meses_abaixo and (', '.join(m.title() for m in meses_abaixo))}</b> {'vieram abaixo do ano anterior' if meses_abaixo else 'Nenhum mês veio abaixo do ano anterior'} — vale cruzar com o calendário de eventos/feriados desses meses.</li>
        <li>A margem bruta sobre faturamento em 2025 ficou em <b>{pct(margem_media_mensal_25)}</b> na média do ano, sinal de que o custo de mercadoria vendida está sob controle mesmo com a receita oscilando mês a mês.</li>
        <li>2026 tem dados fechados até <b>{MESES_ABR[ultimo_mes26_idx]}</b> nesta planilha — a seção "Operação diária" cobre o snapshot mais recente disponível.</li>
      </ul>
    </div>
  </section>
"""

sec_operacao = f"""
  <section class="block" id="operacao">
    <div class="block-head"><h2>Operação diária — snapshot mais recente</h2><span class="block-num">03 / 07</span></div>
    <div class="two-col">
      <div class="panel">
        <h3 class="panel-title">Faturamento por horário</h3>
        <div class="panel-sub">{num_br(mes_atual_vendas)} vendas · {brl(mes_atual_fat)}</div>
        <div class="hour-chart">{hour_chart_html}</div>
      </div>
      <div class="panel">
        <h3 class="panel-title">Meios de pagamento</h3>
        <div class="panel-sub">Participação sobre {brl(meios_total)} vendidos</div>
        <div class="hbar-list">{meios_pagto_hbar}</div>
      </div>
    </div>
    <div class="insight">
      <h3>Pontos importantes</h3>
      <ul>
        <li><b>{pct(janela_pico_pct)} do faturamento</b> acontece entre <b>{janela_ini}h e {janela_fim}h</b> — é a janela que justifica reforço de equipe e estoque de giro rápido.</li>
        <li><b>Cartão de débito + Pix somam {pct(digital_pct)}</b> das vendas — negociar taxa com a adquirente de cartão tem impacto direto e mensurável no lucro.</li>
        <li>O horário de pico do dia é <b>{pico_hora}h</b>, com {brl(pico_valor)} vendidos nessa hora.</li>
      </ul>
    </div>
  </section>
"""

sec_margem = f"""
  <section class="block" id="margem">
    <div class="block-head"><h2>Margem &amp; precificação</h2><span class="block-num">04 / 07</span></div>
    <div class="stat-grid">
      <div class="stat">
        <div class="label">Margem operacional média</div>
        <div class="value">{pct(margem_media)}</div>
        <div class="delta flat">sobre {margem_n} produtos precificados</div>
      </div>
      <div class="stat">
        <div class="label">Classificação de margem</div>
        <div class="chip-row" style="margin-top:8px">
          <span class="chip good"><span class="dot"></span>Alta · {n_alta}</span>
          <span class="chip warning"><span class="dot"></span>Média · {n_media}</span>
          <span class="chip critical"><span class="dot"></span>Baixa · {n_baixa}</span>
        </div>
      </div>
      <div class="stat">
        <div class="label">Produtos sinalizados p/ ajuste de preço</div>
        <div class="value">{len(ajustar)}</div>
        <div class="delta down">margem operacional abaixo do ideal</div>
      </div>
    </div>
    <div class="grid-2">
      <div class="panel">
        <h3 class="panel-title">Maior margem em R$ por produto</h3>
        <div class="panel-sub">O que mais contribui em reais por unidade vendida</div>
        <div class="hbar-list">{top_margem_hbar}</div>
      </div>
      <div class="panel">
        <h3 class="panel-title">Sinalizados para ajuste de preço</h3>
        <div class="panel-sub">Classificação "Baixa" na planilha de margem operacional</div>
        <div class="table-scroll"><table>
          <thead><tr><th>Produto</th><th class="num">Custo</th><th class="num">Venda</th><th class="num">Margem</th></tr></thead>
          <tbody>{ajustar_rows}</tbody>
        </table></div>
      </div>
    </div>
    <div class="panel">
      <h3 class="panel-title">Cardápio de copões — custo e margem por dose</h3>
      <div class="panel-sub">Extraído das fichas técnicas de custo de cada copão do cardápio</div>
      <div class="table-scroll"><table>
        <thead><tr><th>Copão</th><th class="num">Custo</th><th class="num">Preço</th><th class="num">Lucro/copo</th><th class="num">Margem</th></tr></thead>
        <tbody>{cocktail_rows}</tbody>
      </table></div>
    </div>
    <div class="insight">
      <h3>Pontos importantes</h3>
      <ul>
        <li>{'A linha <b>' + esc(pior_copao['nome']) + '</b> é a de menor margem do cardápio, com ' + pct(pior_copao['margem']) + ' — vale reajustar o preço ou revisar a ficha técnica.' if pior_copao else 'Nenhum copão cadastrado nas fichas de custo.'}</li>
        <li><b>{pct(pct_alta)}</b> do portfólio está em faixa de margem <b>Alta</b> — o problema de rentabilidade, quando existe, tende a estar concentrado em poucos itens específicos.</li>
        <li>{len(ajustar)} produto(s) sinalizados para ajuste de preço; como costumam girar bem, um pequeno reajuste tem efeito imediato no caixa.</li>
      </ul>
    </div>
  </section>
"""

sec_produtos = f"""
  <section class="block" id="produtos">
    <div class="block-head"><h2>Produtos &amp; clientes</h2><span class="block-num">05 / 07</span></div>
    <div class="grid-2">
      <div class="panel">
        <h3 class="panel-title">Produtos mais vendidos (por quantidade)</h3>
        <div class="panel-sub">Ranking do relatório de vendas por produto</div>
        <div class="table-scroll"><table>
          <thead><tr><th></th><th>Produto</th><th class="num">Qtd.</th><th class="num">Total</th></tr></thead>
          <tbody>{top_produtos_rows}</tbody>
        </table></div>
      </div>
      <div class="panel">
        <h3 class="panel-title">Top clientes por valor de compras</h3>
        <div class="panel-sub">{clientes_ativos} clientes ativos cadastrados · {ranking_clientes_n} com histórico de compras</div>
        <div class="hbar-list">{top10_clientes_hbar}</div>
      </div>
    </div>
    <div class="panel">
      <h3 class="panel-title">Clientes com saldo devedor (fiado)</h3>
      <div class="panel-sub">{n_devedores} cliente(s) com débito em aberto na conta corrente · total {brl(total_fiado)}</div>
      <div class="table-scroll"><table>
        <thead><tr><th>Cliente</th><th class="num">Débito</th><th class="num">Crédito</th><th class="num">Saldo devedor</th></tr></thead>
        <tbody>{devedores_rows}</tbody>
      </table></div>
    </div>
    <div class="insight">
      <h3>Pontos importantes</h3>
      <ul>
        <li>O produto mais vendido é <b>{esc(top_produtos.iloc[0]['Nome']) if len(top_produtos) else '-'}</b>, com {int(top_produtos.iloc[0]['Qtd.']) if len(top_produtos) else 0} unidades no período do relatório.</li>
        <li>Os <b>10 clientes do topo do ranking respondem por {pct(top10_concentracao)}</b> de todo o valor de compras rastreado — um programa simples de fidelidade para esse grupo protege receita concentrada.</li>
        <li>{('<b>' + ', '.join(esc(p) for p in criticos_alto_giro) + '</b> ' + ('está' if len(criticos_alto_giro) == 1 else 'estão') + ' ao mesmo tempo entre os mais vendidos e em situação crítica de estoque — priorizar essa(s) compra(s) primeiro.') if criticos_alto_giro else 'Nenhum produto de alto giro está em situação crítica de estoque no momento.'}</li>
        <li>{('O fiado em aberto soma <b>' + brl(total_fiado) + '</b> em ' + str(n_devedores) + ' cliente(s); <b>' + esc(maior_devedor['Nome']) + '</b> concentra o maior saldo devedor (' + brl(maior_devedor['_saldo']) + ').') if maior_devedor is not None else 'Nenhum cliente com saldo devedor em aberto no momento.'}</li>
      </ul>
    </div>
  </section>
"""

sec_estoque = f"""
  <section class="block" id="estoque">
    <div class="block-head"><h2>Estoque &amp; planejamento de compras</h2><span class="block-num">06 / 07</span></div>
    <div class="stat-grid">
      <div class="stat">
        <div class="label">Valor em estoque (custo)</div>
        <div class="value">{brl(estoque_custo_total)}</div>
        <div class="delta flat">potencial de venda: {brl(estoque_venda_total)}</div>
      </div>
      <div class="stat">
        <div class="label">Produtos com estoque zerado</div>
        <div class="value">{estoque_zerados_n} <span style="font-size:1rem;color:var(--text-muted)">de {estoque_n}</span></div>
        <div class="delta down">{pct(estoque_zerados_pct)} do catálogo sem disponibilidade</div>
      </div>
      <div class="stat">
        <div class="label">Itens em situação crítica de compra</div>
        <div class="value">{n_critico}</div>
        <div class="delta down">precisam de reposição urgente</div>
      </div>
    </div>
    <div class="grid-2">
      <div class="panel">
        <h3 class="panel-title">Valor de estoque por categoria</h3>
        <div class="panel-sub">Custo total imobilizado, principais categorias</div>
        <div class="hbar-list">{estoque_cat_hbar}</div>
      </div>
      <div class="panel">
        <h3 class="panel-title">Status do planejamento de compras</h3>
        <div class="panel-sub">{planej_n} produtos analisados por giro dos últimos 30 dias</div>
        <div class="chip-row">
          <span class="chip critical"><span class="dot"></span>Crítico · {n_critico}</span>
          <span class="chip warning"><span class="dot"></span>Atenção · {n_atencao}</span>
          <span class="chip good"><span class="dot"></span>OK · {n_ok}</span>
          <span class="chip neutral"><span class="dot"></span>Excesso · {n_excesso}</span>
          <span class="chip neutral"><span class="dot"></span>Sem giro · {n_semgiro}</span>
        </div>
        <div class="panel-sub" style="margin-top:16px; margin-bottom:8px;">Exemplos de itens críticos (comprar com urgência)</div>
        <div class="table-scroll"><table><tbody>{criticos_rows}</tbody></table></div>
      </div>
    </div>
    <div class="insight">
      <h3>Pontos importantes</h3>
      <ul>
        <li>Apenas <b>{pct(n_ok/planej_n*100 if planej_n else 0)}</b> dos produtos estão em status "OK" de estoque — o grosso está em <b>excesso ({pct(n_excesso/planej_n*100 if planej_n else 0)})</b>, <b>sem giro ({pct(n_semgiro/planej_n*100 if planej_n else 0)})</b> ou <b>crítico ({pct(n_critico/planej_n*100 if planej_n else 0)})</b>.</li>
        <li>{('<b>' + ', '.join(esc(p) for p in criticos_alto_giro) + '</b> ' + ('combina' if len(criticos_alto_giro) == 1 else 'combinam') + ' alta venda com ruptura de estoque — ' + ('é a' if len(criticos_alto_giro) == 1 else 'são a') + ' prioridade de compra.') if criticos_alto_giro else 'Nenhum item de alto giro está em ruptura no momento — bom sinal de reposição.'}</li>
        <li>Os {n_semgiro} itens "sem giro" são candidatos naturais a saírem do cardápio/prateleira, liberando espaço e capital.</li>
      </ul>
    </div>
  </section>
"""

sec_financeiro = f"""
  <section class="block" id="financeiro">
    <div class="block-head"><h2>Financeiro — contas a pagar &amp; despesas</h2><span class="block-num">07 / 07</span></div>
    <div class="chip-row">
      <span class="chip good"><span class="dot"></span>Pagas · {n_pagas}</span>
      <span class="chip critical"><span class="dot"></span>Vencidas · {n_vencidas} ({brl(v_vencidas)})</span>
      <span class="chip warning"><span class="dot"></span>Vence hoje · {n_vence_hoje}</span>
      <span class="chip neutral"><span class="dot"></span>A vencer · {n_a_vencer} ({brl(v_a_vencer)})</span>
    </div>
    <div class="panel" style="margin-top:20px;">
      <h3 class="panel-title">Contas a pagar em aberto</h3>
      <div class="panel-sub">{len(contas_pendentes)} conta(s) pendente(s) · total {brl(total_pendente)} · ordenado por vencimento</div>
      <div class="table-scroll"><table>
        <thead><tr><th>Status</th><th>Vencimento</th><th>Fornecedor</th><th>Categoria</th><th>Referente a</th><th class="num">Valor</th></tr></thead>
        <tbody>{contas_pendentes_rows}</tbody>
      </table></div>
    </div>
    <div class="grid-2" style="margin-top:20px;">
      <div class="panel">
        <h3 class="panel-title">Maiores categorias de despesa (histórico registrado)</h3>
        <div class="hbar-list">{contas_cat_hbar}</div>
      </div>
      <div class="panel">
        <h3 class="panel-title">Maiores fornecedores (histórico registrado)</h3>
        <div class="hbar-list">{contas_forn_hbar}</div>
      </div>
    </div>
    <div class="panel">
      <h3 class="panel-title">Resultado do snapshot mais recente</h3>
      <div class="panel-sub">Receitas e despesas do período mais recente disponível</div>
      <div class="grid-3">
        <div class="stat" style="box-shadow:none;"><div class="label">Receitas</div><div class="value">{brl(receitas_total)}</div></div>
        <div class="stat" style="box-shadow:none;"><div class="label">Despesas ({esc(maior_despesa_cat)} {pct(maior_despesa_pct)} do total)</div><div class="value">{brl(despesas_total)}</div></div>
        <div class="stat" style="box-shadow:none; border-color:var(--accent-strong);"><div class="label">Receitas − despesas</div><div class="value" style="color:var(--accent-strong)">{brl(resultado_liquido)}</div></div>
      </div>
    </div>
    <div class="insight">
      <h3>Pontos importantes</h3>
      <ul>
        <li>{('Apenas <b>' + str(n_vencidas) + ' conta(s) vencida(s)</b> (' + brl(v_vencidas) + ') e <b>' + str(n_vence_hoje) + ' vence(m) hoje</b> — gestão de pagamentos sob controle.') if n_vencidas <= 3 else ('<b>' + str(n_vencidas) + ' contas vencidas</b> somando ' + brl(v_vencidas) + ' — vale priorizar a regularização.')}</li>
        <li><b>{esc(maior_despesa_cat)} concentra {pct(maior_despesa_pct)}</b> das despesas do snapshot mais recente e <b>{pct(bebidas_pct_hist)}</b> de todo o histórico de contas a pagar — é a categoria com maior alavancagem para negociação junto a <b>{esc(top_fornecedor_nome)}</b>, o maior fornecedor.</li>
        <li>O resultado líquido do período (<b>{brl(resultado_liquido)}</b>) {'supera' if resultado_liquido > lucro2025_total/12 else 'fica abaixo de'} o lucro bruto médio mensal de 2025 ({brl(lucro2025_total/12)}).</li>
      </ul>
    </div>
  </section>
"""

FOOTER = f"""
<footer>
  Painel gerado automaticamente a partir das planilhas do Google Drive.<br>
  {esc(fontes_footer)}
</footer>
"""

html_out = f"""<meta charset="utf-8">
<title>Resenha Beer — Painel Executivo</title>
<link rel="icon" href="{FAVICON_HREF}">
<style>{CSS}</style>

<header class="hero">
  <div class="wrap">
    <span class="eyebrow">Painel Executivo · Resenha Beer</span>
    <h1 class="brand">Do balcão para a planilha, da planilha para a decisão.</h1>
    <p class="hero-sub">Consolidado das planilhas operacionais do bar — faturamento, margem, estoque, compras, clientes e contas a pagar — organizado em seções com os pontos que pedem atenção agora.</p>
    <div class="hero-meta">
      <span>Período: <strong>{periodo_label}</strong></span>
      <span>Atualizado em <strong>{gerado_em}</strong></span>
      <span>Fonte: exportações do sistema de gestão do bar (Google Drive)</span>
    </div>
  </div>
</header>

<nav class="section-nav">
  <div class="nav-inner">
    <a href="#visao-geral">Visão geral</a>
    <a href="#faturamento">Faturamento</a>
    <a href="#operacao">Operação diária</a>
    <a href="#margem">Margem &amp; preço</a>
    <a href="#produtos">Produtos &amp; clientes</a>
    <a href="#estoque">Estoque &amp; compras</a>
    <a href="#financeiro">Financeiro</a>
  </div>
</nav>

<div class="wrap">
{sec_visao}
{sec_faturamento}
{sec_operacao}
{sec_margem}
{sec_produtos}
{sec_estoque}
{sec_financeiro}
</div>
{FOOTER}
<script>{SCRIPT}</script>
"""

OUT.write_text(html_out, encoding="utf-8")
print(f"OK - {OUT} gerado com sucesso ({len(html_out):,} caracteres).")
