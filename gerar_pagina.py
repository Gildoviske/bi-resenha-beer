# -*- coding: utf-8 -*-
"""Gera o painel executivo do Resenha Beer a partir das planilhas do Drive.

Uso: python gerar_pagina.py  (ou clique duas vezes em publicar.bat)
Lê os arquivos em BASE (a pasta acima desta) e grava index.html nesta mesma pasta,
pronta para publicar no GitHub Pages.
"""
import colorsys
import html as html_lib
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import quote

import pandas as pd

# pasta onde ficam as planilhas: a pasta acima desta (ex.: G:\Meu Drive ou J:\Meu Drive),
# assim funciona em qualquer PC independente da letra do drive do Google Drive
BASE = Path(__file__).resolve().parent.parent
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


def ler_controles():
    """Os dois ajustes que ficam na planilha: quantos dias de venda a
    sugestão cobre (AA1) e a margem a mais aplicada em cima (AA2)."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path("PLANEJAMENTO DE COMPRAS - RESENHA BEER.xlsx"),
                           read_only=True, data_only=True)
        aba = wb["PLANEJAMENTO"]
        dias, margem = aba["AA1"].value, aba["AA2"].value
        wb.close()
        return int(dias), float(margem or 0)
    except (KeyError, TypeError, ValueError, OSError):
        return 30, 0.0


cobertura_dias, margem_compra = ler_controles()
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
#  8) LISTA DE COMPRAS — o que está pintado na planilha
#
#  O controle é a cor da linha na aba PLANEJAMENTO: verde = vou comprar,
#  roxo = já pedi. Sem data, prazo ou otimização — quem decide é quem
#  pinta, e o painel só lê e soma.
# ======================================================================
VERDE, ROXO = "verde", "roxo"
ROTULO_COR = {VERDE: "vou comprar", ROXO: "já pedi"}


def _rgb_da_celula(cel):
    """Cor de preenchimento sólido de uma célula, como (r, g, b)."""
    fill = cel.fill
    if fill is None or fill.patternType != "solid":
        return None
    cor = fill.fgColor
    if cor is None or cor.type != "rgb" or not isinstance(cor.rgb, str):
        return None
    try:
        hexa = cor.rgb[-6:]
        return tuple(int(hexa[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return None


def _classifica_cor(rgb):
    """Verde ou roxo pelo matiz, não pelo código exato: assim vale qualquer
    tom que o Excel ofereça na paleta, claro ou escuro."""
    if rgb is None:
        return None
    r, g, b = (v / 255 for v in rgb)
    matiz, sat, brilho = colorsys.rgb_to_hsv(r, g, b)
    if sat < 0.15 or brilho < 0.20:
        return None  # branco, cinza ou preto: não é marcação
    graus = matiz * 360
    if 75 <= graus <= 175:
        return VERDE
    if 250 <= graus <= 330:
        return ROXO
    return None


def ler_cores_planejamento():
    """Cor de cada linha da aba PLANEJAMENTO. Olha Código e Produto, que são
    as colunas sem formatação condicional — assim a regra de cor do Status
    não se confunde com a marcação feita à mão."""
    from openpyxl import load_workbook

    cores, achadas = {}, {}
    try:
        wb = load_workbook(path("PLANEJAMENTO DE COMPRAS - RESENHA BEER.xlsx"))
        ws = wb["PLANEJAMENTO"]
        for linha in range(3, ws.max_row + 1):
            marcas = []
            for coluna in (1, 2):
                rgb = _rgb_da_celula(ws.cell(row=linha, column=coluna))
                marca = _classifica_cor(rgb)
                if marca:
                    marcas.append(marca)
                    achadas.setdefault(marca, set()).add("#%02X%02X%02X" % rgb)
            # roxo ganha do verde: é o estado mais adiantado da mesma linha
            if ROXO in marcas:
                cores[linha] = ROXO
            elif VERDE in marcas:
                cores[linha] = VERDE
        wb.close()
    except (KeyError, OSError) as erro:
        print(f"AVISO - nao foi possivel ler as cores da planilha ({erro})")
    return cores, {k: sorted(v) for k, v in achadas.items()}


cores_por_linha, tons_encontrados = ler_cores_planejamento()

# a aba tem o cabeçalho na linha 2, então o índice 0 do dataframe é a linha 3
dfpl["_linha"] = [int(i) + 3 for i in dfpl.index]
dfpl["_cor"] = [cores_por_linha.get(linha) for linha in dfpl["_linha"]]


def _txt(v):
    """Texto limpo, ou "" quando a célula vem vazia/NaN."""
    return v.strip() if isinstance(v, str) and v.strip() and v.strip().lower() != "nan" else ""


_FORN_CONTAS = sorted(dfcp["Fornecedor"].dropna().astype(str).unique(), key=len, reverse=True)


def forn_completo(nome):
    """No cadastro de estoque o fornecedor vem truncado em 20 caracteres;
    casa pelo prefixo com o nome completo usado nas contas a pagar."""
    n = _txt(nome).upper()
    if not n or n == "0":
        return None
    for f in _FORN_CONTAS:
        fu = f.upper()
        if fu.startswith(n) or n.startswith(fu):
            return f
    return _txt(nome)


_EST_INFO = dfe.drop_duplicates("Código").set_index("Código")[
    ["Custo", "Fornecedor", "Categoria"]].to_dict("index")
_CUSTO_CAT = dfe[dfe["Custo"] > 0].groupby("Categoria")["Custo"].median().to_dict()
SEM_FORNECEDOR = "Sem fornecedor no cadastro"


def _monta_item(registro):
    info = _EST_INFO.get(registro["Código"], {})
    categoria = _txt(info.get("Categoria")) or _txt(registro.get("Categoria")) or "SEM CATEGORIA"
    custo = float(info.get("Custo") or 0)
    if custo <= 0:
        custo = float(_CUSTO_CAT.get(categoria, 0) or 0)
    fornecedor = forn_completo(info.get("Fornecedor")) or forn_completo(registro.get("Fornecedor"))
    qtd = float(pd.to_numeric(registro.get("Pedido ajustado"), errors="coerce") or 0)
    return {
        "produto": _txt(registro["Produto"]), "cat": categoria, "qtd": int(qtd),
        "custo": custo, "valor": custo * qtd, "status": _txt(registro.get("Status")),
        "forn": fornecedor or SEM_FORNECEDOR, "sem_forn": fornecedor is None,
        "estoque": float(pd.to_numeric(registro.get("Estoque atual"), errors="coerce") or 0),
    }


def _lista(cor):
    itens = [_monta_item(r) for r in dfpl[dfpl["_cor"] == cor].to_dict("records")]
    return sorted(itens, key=lambda i: (i["forn"], -i["valor"]))


itens_verdes = _lista(VERDE)
itens_roxos = _lista(ROXO)

# crítico que ainda não recebeu cor nenhuma: é o que pede decisão
itens_sem_marca = sorted(
    (_monta_item(r) for r in dfpl[dfpl["_cor"].isna() & (dfpl["Status"] == "CRÍTICO")].to_dict("records")),
    key=lambda i: -i["valor"],
)

total_verde = sum(i["valor"] for i in itens_verdes)
total_roxo = sum(i["valor"] for i in itens_roxos)
n_verde, n_roxo, n_sem_marca = len(itens_verdes), len(itens_roxos), len(itens_sem_marca)
n_sem_forn = sum(1 for i in itens_verdes if i["sem_forn"])


def _agrupa(itens):
    grupos = {}
    for it in itens:
        g = grupos.setdefault(it["forn"], {"forn": it["forn"], "itens": 0, "qtd": 0,
                                           "valor": 0.0, "criticos": 0})
        g["itens"] += 1
        g["qtd"] += it["qtd"]
        g["valor"] += it["valor"]
        g["criticos"] += 1 if it["status"] == "CRÍTICO" else 0
    return sorted(grupos.values(), key=lambda g: -g["valor"])


compras_por_fornecedor = _agrupa(itens_verdes)
pedidos_por_fornecedor = _agrupa(itens_roxos)
maior_fornecedor = compras_por_fornecedor[0] if compras_por_fornecedor else None


# ======================================================================
#  HELPERS DE HTML
# ======================================================================
def hbar_rows(items, alt=False, fmt=brl):
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
            f'<span class="val">{fmt(val)}</span></div>'
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
.bar[data-tip]:hover::after, .hbar[data-tip]:hover::after, .col-bar[data-tip]:hover::after,
  content:attr(data-tip); position:absolute; bottom:calc(100% + 6px); left:50%; transform:translateX(-50%);
  background:var(--ink-900); color:#F3E9DA; font-size:0.72rem; font-weight:600; padding:4px 8px; border-radius:6px;
  white-space:nowrap; z-index:5; box-shadow:var(--shadow); font-variant-numeric:tabular-nums;
}
.month-label{margin-top:8px; font-size:0.68rem; color:var(--text-muted); font-weight:600;}
.legend{display:flex; gap:18px; margin-top:14px; font-size:0.8rem; color:var(--text-secondary);}
.legend span{display:inline-flex; align-items:center; gap:6px;}
.legend i{width:10px; height:10px; border-radius:3px; display:inline-block;}
/* sem isto, uma tabela larga estica a coluna do grid e vaza a página */
.grid-2>*, .grid-3>*, .two-col>*{min-width:0;}
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
    <div class="block-head"><h2>Visão geral</h2><span class="block-num">01 / 08</span></div>
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
    <div class="block-head"><h2>Faturamento &amp; lucro mês a mês</h2><span class="block-num">02 / 08</span></div>
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
    <div class="block-head"><h2>Operação diária — snapshot mais recente</h2><span class="block-num">03 / 08</span></div>
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
    <div class="block-head"><h2>Margem &amp; precificação</h2><span class="block-num">04 / 08</span></div>
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
    <div class="block-head"><h2>Produtos &amp; clientes</h2><span class="block-num">05 / 08</span></div>
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
    <div class="block-head"><h2>Estoque &amp; planejamento de compras</h2><span class="block-num">06 / 08</span></div>
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
    <div class="block-head"><h2>Financeiro — contas a pagar &amp; despesas</h2><span class="block-num">07 / 08</span></div>
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

# ----------------------------------------------------------------------
#  LISTA DE COMPRAS — tabelas por cor e total por fornecedor
# ----------------------------------------------------------------------
COR_TAG = {VERDE: "high", ROXO: "mid"}
STATUS_TAG_COMPRA = {"CRÍTICO": "low", "ATENÇÃO": "mid", "OK": "high"}


def _linha_item(it, com_status=True):
    status = (f'<td class="num"><span class="tag {STATUS_TAG_COMPRA.get(it["status"], "muted")}">'
              f'{esc(it["status"] or "—")}</span></td>') if com_status else ""
    return (f'<tr><td>{esc(it["produto"])}</td>'
            f'<td>{esc(it["forn"])}</td>'
            f'{status}'
            f'<td class="num">{it["qtd"]}</td>'
            f'<td class="num">{brl(it["custo"])}</td>'
            f'<td class="num strong">{brl(it["valor"])}</td></tr>')


def _tabela_itens(itens, vazio):
    return "\n".join(_linha_item(i) for i in itens) or f'<tr><td colspan="6">{vazio}</td></tr>'


def _linha_grupo(g):
    return (f'<tr><td class="strong">{esc(g["forn"])}</td>'
            f'<td class="num">{g["itens"]}</td>'
            f'<td class="num">{g["qtd"]}</td>'
            f'<td class="num">{g["criticos"]}</td>'
            f'<td class="num strong">{brl(g["valor"])}</td></tr>')


def _tabela_grupos(grupos, vazio):
    return "\n".join(_linha_grupo(g) for g in grupos) or f'<tr><td colspan="5">{vazio}</td></tr>'


verdes_rows = _tabela_itens(itens_verdes, "Nenhuma linha pintada de verde ainda.")
roxos_rows = _tabela_itens(itens_roxos, "Nenhuma linha pintada de roxo ainda.")
sem_marca_rows = _tabela_itens(itens_sem_marca, "Todo item crítico já foi marcado.")
compras_forn_rows = _tabela_grupos(compras_por_fornecedor, "Nada marcado para comprar.")
compras_forn_hbar = hbar_rows([(g["forn"], g["valor"]) for g in compras_por_fornecedor[:7]])

tons_txt = " · ".join(
    f'{ROTULO_COR[c]}: {", ".join(t)}' for c, t in sorted(tons_encontrados.items())
) or "nenhuma cor encontrada nas colunas Código e Produto"

sec_compras = f"""
  <section class="block" id="compras">
    <div class="block-head"><h2>Lista de compras — o que está marcado na planilha</h2><span class="block-num">08 / 08</span></div>
    <p class="panel-sub" style="margin:-6px 0 22px;">
      O controle é a cor da linha na aba PLANEJAMENTO: <b>verde</b> para o que você vai comprar,
      <b>roxo</b> para o que já pediu. As quantidades saem da própria planilha, hoje ajustada
      para cobrir <b>{cobertura_dias} dias</b> de venda com <b>{pct(margem_compra * 100, 0)}</b> de margem a mais.
    </p>
    <div class="stat-grid">
      <div class="stat">
        <div class="label">Marcado para comprar</div>
        <div class="value" style="color:var(--accent-strong)">{brl(total_verde)}</div>
        <div class="delta flat">{n_verde} produtos em {len(compras_por_fornecedor)} fornecedores</div>
      </div>
      <div class="stat">
        <div class="label">Já pedido</div>
        <div class="value">{brl(total_roxo)}</div>
        <div class="delta flat">{n_roxo} produtos {'aguardando chegar' if n_roxo else 'marcados até agora'}</div>
      </div>
      <div class="stat">
        <div class="label">Crítico sem marcação</div>
        <div class="value">{n_sem_marca}</div>
        <div class="delta {'down' if n_sem_marca else 'up'}">{'produtos zerando sem decisão tomada' if n_sem_marca else 'todo crítico já foi decidido'}</div>
      </div>
    </div>
    <div class="grid-2">
      <div class="panel">
        <h3 class="panel-title">Quanto vai para cada fornecedor</h3>
        <div class="panel-sub">Somando só as linhas verdes · use antes de ligar para fechar o pedido</div>
        <div class="table-scroll"><table>
          <thead><tr><th>Fornecedor</th><th class="num">Produtos</th><th class="num">Unidades</th><th class="num">Críticos</th><th class="num">Total</th></tr></thead>
          <tbody>{compras_forn_rows}</tbody>
        </table></div>
      </div>
      <div class="panel">
        <h3 class="panel-title">Peso de cada fornecedor na compra</h3>
        <div class="panel-sub">Maiores valores entre os marcados de verde</div>
        <div class="hbar-list">{compras_forn_hbar}</div>
      </div>
    </div>
    <div class="panel" style="margin-top:20px;">
      <h3 class="panel-title">Verde — vou comprar</h3>
      <div class="panel-sub">{n_verde} produto(s) · {brl(total_verde)} · ordenado por fornecedor</div>
      <div class="table-scroll" style="max-height:440px;"><table>
        <thead><tr><th>Produto</th><th>Fornecedor</th><th class="num">Status</th><th class="num">Qtd</th><th class="num">Custo unit.</th><th class="num">Total</th></tr></thead>
        <tbody>{verdes_rows}</tbody>
      </table></div>
    </div>
    <div class="panel" style="margin-top:20px;">
      <h3 class="panel-title">Roxo — já pedi</h3>
      <div class="panel-sub">{n_roxo} produto(s) · {brl(total_roxo)} · some daqui quando você tirar a cor</div>
      <div class="table-scroll" style="max-height:360px;"><table>
        <thead><tr><th>Produto</th><th>Fornecedor</th><th class="num">Status</th><th class="num">Qtd</th><th class="num">Custo unit.</th><th class="num">Total</th></tr></thead>
        <tbody>{roxos_rows}</tbody>
      </table></div>
    </div>
    <div class="panel" style="margin-top:20px;">
      <h3 class="panel-title">Crítico e ainda sem cor</h3>
      <div class="panel-sub">Estoque acabando e nenhuma decisão tomada — pinte ou ignore conscientemente</div>
      <div class="table-scroll" style="max-height:360px;"><table>
        <thead><tr><th>Produto</th><th>Fornecedor</th><th class="num">Status</th><th class="num">Qtd sugerida</th><th class="num">Custo unit.</th><th class="num">Total</th></tr></thead>
        <tbody>{sem_marca_rows}</tbody>
      </table></div>
    </div>
    <div class="insight">
      <h3>Pontos importantes</h3>
      <ul>
        <li>{('Marcado para comprar: <b>' + brl(total_verde) + '</b> em ' + str(n_verde) + ' produtos. O maior pedido é <b>' + esc(maior_fornecedor["forn"]) + '</b>, com ' + brl(maior_fornecedor["valor"]) + ' em ' + str(maior_fornecedor["itens"]) + ' itens.') if maior_fornecedor else 'Nenhuma linha está pintada de verde — o painel não tem o que listar até você marcar o que vai comprar.'}</li>
        <li>{('<b>' + str(n_sem_marca) + ' produto(s) em estado crítico</b> ainda não foram pintados, somando ' + brl(sum(i["valor"] for i in itens_sem_marca)) + ' se forem comprados na quantidade sugerida — são as decisões que faltam.') if n_sem_marca else 'Todo produto em estado crítico já está marcado de verde ou de roxo: nenhuma decisão pendente.'}</li>
        <li>{('<b>' + str(n_roxo) + ' produto(s)</b> estão de roxo, ' + brl(total_roxo) + ' já pedidos. Quando a mercadoria chegar e o estoque for reexportado, tire a cor para eles saírem da lista.') if n_roxo else 'Nada marcado como já pedido no momento.'}</li>
        <li>{('<b>' + str(n_sem_forn) + ' item(ns) verdes</b> não têm fornecedor no cadastro de estoque e aparecem agrupados como "' + SEM_FORNECEDOR + '" — preencher esse campo organiza o total por fornecedor.') if n_sem_forn else 'Todos os itens marcados têm fornecedor cadastrado, então o total por fornecedor está completo.'}</li>
        <li>Tons lidos na planilha — {esc(tons_txt)}. A classificação é por matiz, então qualquer verde ou roxo da paleta do Excel funciona.</li>
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
    <a href="#compras">Lista de compras</a>
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
{sec_compras}
</div>
{FOOTER}
<script>{SCRIPT}</script>
"""

OUT.write_text(html_out, encoding="utf-8")
print(f"OK - {OUT} gerado com sucesso ({len(html_out):,} caracteres).")


# ======================================================================
#  LIMPEZA DA PLANILHA DE PLANEJAMENTO
#
#  Tira o que saiu de uso quando o controle virou a cor da linha: o bloco
#  AGENDA DE COMPRAS (colunas R a X) e a aba PEDIDOS. Mexe no XML na unha
#  em vez de usar o openpyxl porque reescrever o pacote inteiro apagaria
#  os valores em cache das fórmulas, e é deles que o painel se alimenta.
#  A coluna J e os controles em Z/AA não são tocados: são seus.
# ======================================================================
PLANEJ_XLSX = BASE / "PLANEJAMENTO DE COMPRAS - RESENHA BEER.xlsx"
COLS_AGENDA_ANTIGAS = "R|S|T|U|V|W|X"


def _arquivo_da_aba(partes, nome):
    """Caminho do XML de uma aba, resolvido pelo nome (a ordem pode mudar)."""
    wbx = partes["xl/workbook.xml"].decode("utf-8")
    rid = re.search(rf'<sheet name="{re.escape(nome)}"[^>]*r:id="([^"]+)"', wbx).group(1)
    rels = partes["xl/_rels/workbook.xml.rels"].decode("utf-8")
    destino = re.search(rf'<Relationship Id="{rid}"[^>]*Target="([^"]+)"', rels).group(1)
    return "xl/" + destino.lstrip("/")


def _limpar_bloco_agenda(sheet):
    """Remove as células, a mesclagem e as larguras das colunas R:X."""
    ultima = max(int(n) for n in re.findall(r'<row r="(\d+)"', sheet))
    sheet = re.sub(rf'<c r="(?:{COLS_AGENDA_ANTIGAS})\d+"(?:[^>]*/>|[^>]*>.*?</c>)',
                   "", sheet, flags=re.S)
    merge = re.search(r'<mergeCells count="(\d+)">', sheet)
    if merge and '<mergeCell ref="S1:X1"/>' in sheet:
        sheet = sheet.replace('<mergeCell ref="S1:X1"/>', "", 1)
        sheet = sheet.replace(merge.group(0), f'<mergeCells count="{int(merge.group(1)) - 1}">', 1)
    sheet = re.sub(r'<col min="(?:1[89]|2[0-4])" max="(?:1[89]|2[0-4])"[^>]*/>', "", sheet)
    # o filtro volta a cobrir só as colunas de dados; Z/AA ficam de fora dele
    sheet = re.sub(r'(<autoFilter ref=")A2:[A-Z]+\d+(")', rf"\g<1>A2:Q{ultima}\g<2>", sheet, count=1)
    sheet = re.sub(r'<dimension ref="A1:[A-Z]+\d+"/>', f'<dimension ref="A1:AA{ultima}"/>',
                   sheet, count=1)
    return sheet


def _remover_parte(partes, nomes, caminho):
    """Tira um arquivo do pacote junto com o Override que o declara."""
    ct = partes["[Content_Types].xml"].decode("utf-8")
    partes["[Content_Types].xml"] = re.sub(
        rf'<Override PartName="/{re.escape(caminho)}"[^>]*/>', "", ct).encode("utf-8")
    return [n for n in nomes if n != caminho]


def _remover_aba_pedidos(partes, nomes):
    """Apaga a aba PEDIDOS do pacote: entrada no workbook, relacionamento,
    declaração de tipo e o próprio XML da aba."""
    wbx = partes["xl/workbook.xml"].decode("utf-8")
    entrada = re.search(r'<sheet name="PEDIDOS"[^>]*/>', wbx)
    if entrada is None:
        return partes, nomes, "aba PEDIDOS ja havia sido removida"

    posicao = re.findall(r'<sheet name="([^"]+)"', wbx).index("PEDIDOS")
    rid = re.search(r'r:id="([^"]+)"', entrada.group(0)).group(1)
    wbx = wbx.replace(entrada.group(0), "", 1)

    # localSheetId é posicional: quem vinha depois da PEDIDOS anda uma casa
    wbx = re.sub(r'localSheetId="(\d+)"',
                 lambda m: f'localSheetId="{int(m.group(1)) - 1}"'
                 if int(m.group(1)) > posicao else m.group(0), wbx)
    partes["xl/workbook.xml"] = wbx.encode("utf-8")

    rels = partes["xl/_rels/workbook.xml.rels"].decode("utf-8")
    alvo = re.search(rf'<Relationship Id="{rid}"[^>]*/>', rels)
    caminho = "xl/" + re.search(r'Target="([^"]+)"', alvo.group(0)).group(1).lstrip("/")
    rels = rels.replace(alvo.group(0), "", 1)

    # a cadeia de cálculo aponta para índices de aba; sem ela o Excel refaz
    calc = re.search(r'<Relationship [^>]*calcChain[^>]*/>', rels)
    if calc:
        rels = rels.replace(calc.group(0), "", 1)
        nomes = _remover_parte(partes, nomes, "xl/calcChain.xml")
    partes["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    nomes = _remover_parte(partes, nomes, caminho)
    return partes, nomes, "aba PEDIDOS removida"


def simplificar_planilha():
    """Aplica a limpeza num temporário, confere e só então troca o original."""
    import shutil
    import zipfile

    with zipfile.ZipFile(PLANEJ_XLSX) as z:
        nomes = z.namelist()
        partes = {n: z.read(n) for n in nomes}

    arq_planej = _arquivo_da_aba(partes, "PLANEJAMENTO")  # resolve antes de mexer nos rels
    partes[arq_planej] = _limpar_bloco_agenda(
        partes[arq_planej].decode("utf-8")).encode("utf-8")
    partes, nomes, recado = _remover_aba_pedidos(partes, nomes)

    # o _FilterDatabase é o nome oculto que o Excel mantém para o autofiltro;
    # tem de encolher junto, senão a planilha reabre com o filtro largo demais
    wbx = partes["xl/workbook.xml"].decode("utf-8")
    partes["xl/workbook.xml"] = re.sub(
        r'(PLANEJAMENTO!\$A\$2:\$)[A-Z]+(\$\d+)', r"\g<1>Q\g<2>", wbx).encode("utf-8")

    tmp = PLANEJ_XLSX.with_name("~limpeza-" + PLANEJ_XLSX.name)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in nomes:
            z.writestr(n, partes[n])
    try:
        # o with fecha o arquivo: no Windows um handle aberto impede a troca
        with pd.ExcelFile(tmp) as planilha:
            conf = pd.read_excel(planilha, sheet_name="PLANEJAMENTO", header=1)
            abas = list(planilha.sheet_names)
        if conf["Média de vendas diária"].notna().sum() == 0:
            raise RuntimeError("os valores em cache das fórmulas se perderam")
        sobrando = [c for c in ("Situação", "Melhor dia de compra", "Vencimento previsto")
                    if c in conf.columns]
        if sobrando:
            raise RuntimeError(f"as colunas da agenda continuam na planilha: {sobrando}")
        if "PEDIDOS" in abas:
            raise RuntimeError("a aba PEDIDOS continua no arquivo")
        shutil.move(str(tmp), str(PLANEJ_XLSX))
    except PermissionError:
        tmp.unlink(missing_ok=True)
        raise RuntimeError("a planilha esta aberta no Excel - feche e rode de novo "
                           "(o arquivo original nao foi alterado)") from None
    except Exception:
        tmp.unlink(missing_ok=True)
        raise
    return f"{recado}; bloco da agenda retirado das colunas R a X"


try:
    print(f"OK - planilha simplificada: {simplificar_planilha()}")
except Exception as _e:
    print(f"AVISO - planilha nao foi simplificada ({type(_e).__name__}: {_e})")
