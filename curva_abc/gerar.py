"""Curva ABC por frequência de giro (meses com venda) — peças BMC Hyundai.

Uso:
    python -m curva_abc.gerar [--dados PASTA] [--saida PASTA] [--notas f_Notas.csv ...]
"""
import argparse
import re
import sys
import unicodedata
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt

DADOS = r"C:\Users\pamela.hames\OneDrive - bmchyundai.com.br\Arquivos de Breno Ipolito Bertoni - data"
SAIDA = r"C:\Users\pamela.hames\OneDrive - bmchyundai.com.br\ANALISE CURVA ABC  PEÇAS"
NOTAS_PADRAO = ["f_Notas.csv"]
PRODUTOS = "d_Produtos.csv"
ESTOQUE = "f_estoque.csv"

# Limiares de frequência (meses distintos com venda na janela)
MIN_A, MIN_B, MIN_C, MIN_D = 9, 6, 3, 1
# Limiares de recência (k = idade do mês da última venda; k=1 é o mês de referência)
MAX_E, MAX_F, MAX_G, MAX_H = 24, 34, 35, 36

PERIODOS = [
    ("2025", pd.Period("2025-12", "M"), 12),
    ("2026_YTD", pd.Period("2026-11", "M"), 11),
]

TIPMOV_VENDA = {"V"}
TIPMOV_DEVOL = {"D"}

PARETO_A, PARETO_B = 0.80, 0.95
COBERTURA_OVERSTOCK = 6.0
COBERTURA_BAIXA = 1.0

CLASSES = [
    ("A", "FAST MOVING", f"≥ {MIN_A} meses com venda na janela"),
    ("B", "FAST MOVING", f"{MIN_B} a {MIN_A - 1} meses com venda"),
    ("C", "FAST MOVING", f"{MIN_C} a {MIN_B - 1} meses com venda"),
    ("D", "SLOW MOVING", f"{MIN_D} a {MIN_C - 1} meses com venda"),
    ("E", "SLOW MOVING", f"sem venda na janela; última venda até {MAX_E} meses"),
    ("F", "SLOW MOVING", f"última venda entre {MAX_E + 1} e {MAX_F} meses"),
    ("G", "SLOW MOVING", f"última venda em {MAX_G} meses"),
    ("H", "NON-MOVING", f"última venda há {MAX_H} meses"),
    ("X", "NON-MOVING", f"sem venda há mais de {MAX_H} meses"),
    ("NO HIST", "NO HIST", "sem histórico de venda"),
]
ORDEM = {c: i for i, (c, _, _) in enumerate(CLASSES)}
GRUPO = {c: g for c, g, _ in CLASSES}
FAST = {"A", "B", "C"}

COLS = {
    "codprod": ["codprod", "cod_prod", "codigo_produto", "cod_produto", "codigo"],
    "pn": ["referencia", "part_number", "partnumber", "pn", "part_no"],
    "desc": ["descrprod", "descricao", "descricao_produto", "produto", "desc"],
    "data": ["dtfatur", "dtneg", "data", "data_venda", "dt_venda", "dtmov", "mes"],
    "qtd": ["qtdneg", "qtd", "quantidade", "qtde", "qtd_vendida"],
    "valor": ["vlrtot", "valor", "vlr_total", "valor_total", "faturamento", "vlrnota"],
    "tipmov": ["tipmov", "tipo_mov"],
    "estoque": ["estoque", "qtdestoque", "saldo", "qtd_estoque", "estoque_atual"],
    "reservado": ["reservado"],
    "estmax": ["estmax"],
    "estmin": ["estmin"],
}


def _norm(c):
    c = str(c)
    m = re.findall(r"\[([^\]]+)\]", c)
    c = m[-1] if m else c
    c = unicodedata.normalize("NFKD", c).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "_", c.strip().lower()).strip("_")


def ler_csv(path):
    for enc in ("utf-8-sig", "latin-1"):
        try:
            df = pd.read_csv(path, sep=None, engine="python", dtype=str, encoding=enc)
            break
        except UnicodeDecodeError:
            continue
    df.columns = [_norm(c) for c in df.columns]
    return df


def achar(df, chave, obrig=True):
    for c in COLS[chave]:
        if c in df.columns:
            return c
    if obrig:
        raise KeyError(f"Coluna '{chave}' não encontrada. Colunas: {list(df.columns)}")
    return None


def num(s):
    s = s.fillna("").astype(str).str.strip().str.replace(r"[R$\s]", "", regex=True)
    br = s.str.contains(",") & (s.str.rfind(",") > s.str.rfind("."))
    s = s.where(~br, s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def datas(s):
    s = s.fillna("").astype(str).str.strip()
    iso = s.str.match(r"^\d{4}-\d{2}").mean() > 0.5
    return pd.to_datetime(s, dayfirst=not iso, errors="coerce", format="mixed")


def coluna_data(df):
    for c in COLS["data"]:
        if c in df.columns and datas(df[c].head(500)).notna().mean() > 0.8:
            return c
    taxas = {c: datas(df[c].head(500)).notna().mean() for c in df.columns}
    c = max(taxas, key=taxas.get)
    if taxas[c] < 0.8:
        raise KeyError("Coluna de data não detectada.")
    return c


def carregar_vendas(dados, arquivos):
    partes = []
    for a in arquivos:
        df = ler_csv(Path(dados) / a)
        cd, cq, cv = coluna_data(df), achar(df, "qtd"), achar(df, "valor")
        ck = achar(df, "codprod", False) or achar(df, "pn")
        out = pd.DataFrame({
            "key": df[ck].str.strip(),
            "data": datas(df[cd]),
            "qtd": num(df[cq]),
            "valor": num(df[cv]),
        })
        ct = achar(df, "tipmov", False)
        if ct:
            t = df[ct].str.strip().str.upper()
            ok = t.isin(TIPMOV_VENDA | TIPMOV_DEVOL)
            if ok.any():
                sinal = np.where(t.isin(TIPMOV_DEVOL), -1.0, 1.0)
                out[["qtd", "valor"]] = out[["qtd", "valor"]].mul(sinal, axis=0)
                out = out[ok]
            else:
                print(f"[aviso] {a}: TIPMOV sem {TIPMOV_VENDA}; usando todas as linhas.")
        if "nunota" in df.columns:
            out["nunota"] = df.loc[out.index, "nunota"]
        partes.append(out)
    v = pd.concat(partes, ignore_index=True)
    if len(arquivos) > 1:
        v = v.drop_duplicates()
    v = v.dropna(subset=["data"])
    v = v[v["key"].fillna("") != ""]
    v["mes"] = v["data"].dt.to_period("M")
    return v


def carregar_produtos(dados, chave_cod):
    p = Path(dados) / PRODUTOS
    if not p.exists():
        return pd.DataFrame(columns=["key", "part_number", "descricao"])
    df = ler_csv(p)
    ck = achar(df, "codprod", False) if chave_cod else achar(df, "pn")
    cp, cdsc = achar(df, "pn", False), achar(df, "desc", False)
    out = pd.DataFrame({
        "key": df[ck].str.strip(),
        "part_number": df[cp].str.strip() if cp else df[ck],
        "descricao": df[cdsc].str.strip() if cdsc else "",
    })
    return out.drop_duplicates("key")


def carregar_estoque(dados):
    p = Path(dados) / ESTOQUE
    if not p.exists():
        return None
    df = ler_csv(p)
    ck, ce = achar(df, "codprod", False) or achar(df, "pn"), achar(df, "estoque", False)
    if not ce:
        print(f"[aviso] coluna de estoque não encontrada em {ESTOQUE}: {list(df.columns)}")
        return None
    out = pd.DataFrame({"key": df[ck].str.strip(), "estoque": num(df[ce])})
    for k in ("reservado", "estmin", "estmax"):
        c = achar(df, k, False)
        out[k] = num(df[c]) if c else 0.0
    g = out.groupby("key", as_index=False).sum()
    g["disponivel"] = (g["estoque"] - g["reservado"]).clip(lower=0)
    return g


def mensal(v):
    return v.groupby(["key", "mes"], as_index=False)[["qtd", "valor"]].sum()


def classe(meses, k):
    if pd.isna(k):
        return "NO HIST"
    if meses >= MIN_A:
        return "A"
    if meses >= MIN_B:
        return "B"
    if meses >= MIN_C:
        return "C"
    if meses >= MIN_D:
        return "D"
    if k <= MAX_E:
        return "E"
    if k <= MAX_F:
        return "F"
    if k <= MAX_G:
        return "G"
    if k <= MAX_H:
        return "H"
    return "X"


def classificar(m, universo, ref, janela):
    x = m[(m["mes"] <= ref) & (m["qtd"] > 0)].copy()
    x["k"] = ref.ordinal - x["mes"].apply(lambda p: p.ordinal) + 1
    jan = x[x["k"] <= janela]
    agg = jan.groupby("key").agg(meses_com_venda=("mes", "nunique"),
                                 qtd_total=("qtd", "sum"), valor_total=("valor", "sum"))
    ult = x.groupby("key").agg(k_ultima=("k", "min"), ultima_venda=("mes", "max"))
    r = universo.set_index("key").join(agg).join(ult)
    r[["meses_com_venda", "qtd_total", "valor_total"]] = r[
        ["meses_com_venda", "qtd_total", "valor_total"]].fillna(0)
    r["meses_com_venda"] = r["meses_com_venda"].astype(int)
    r["meses_desde_ultima_venda"] = r["k_ultima"] - 1
    r["Cod"] = [classe(a, b) for a, b in zip(r["meses_com_venda"], r["k_ultima"])]
    r["Classe"] = r["Cod"].map(GRUPO)
    r["ultima_venda"] = r["ultima_venda"].astype(str).replace("NaT", "")
    r = r.reset_index().sort_values(["valor_total"], ascending=False)
    tot = r["valor_total"].clip(lower=0).sum()
    r["pct_valor"] = r["valor_total"].clip(lower=0) / tot if tot else 0
    r["pct_acum"] = r["pct_valor"].cumsum()
    r["Curva_Valor"] = np.select(
        [r["valor_total"] <= 0, r["pct_acum"] - r["pct_valor"] < PARETO_A,
         r["pct_acum"] - r["pct_valor"] < PARETO_B], ["-", "A", "B"], "C")
    cols = ["key", "part_number", "descricao", "meses_com_venda", "qtd_total", "valor_total",
            "Cod", "Classe", "ultima_venda", "meses_desde_ultima_venda",
            "pct_valor", "pct_acum", "Curva_Valor"]
    return r[cols]


def base_mensal(m, prod, ini, fim):
    m = m[(m["mes"] >= ini) & (m["mes"] <= fim)]
    q = m.pivot_table(index="key", columns="mes", values="qtd", aggfunc="sum", fill_value=0)
    v = m.pivot_table(index="key", columns="mes", values="valor", aggfunc="sum", fill_value=0)
    q.columns = [f"{c} Qtd" for c in q.columns]
    v.columns = [f"{c} R$" for c in v.columns]
    b = prod.set_index("key")[["part_number", "descricao"]].join(
        pd.concat([q, v], axis=1), how="right").reset_index()
    return b


def resumo(abc):
    linhas = []
    for nome, df in abc.items():
        tot_itens, tot_val = len(df), df["valor_total"].clip(lower=0).sum()
        g = df.groupby("Cod").agg(itens=("key", "size"), valor=("valor_total", "sum"))
        for c, grp, regra in CLASSES:
            it = int(g["itens"].get(c, 0))
            vl = float(g["valor"].get(c, 0))
            linhas.append({"Periodo": nome, "Cod": c, "Classe": grp, "Regra": regra,
                           "Itens": it, "% Itens": it / tot_itens if tot_itens else 0,
                           "Valor": vl, "% Valor": vl / tot_val if tot_val else 0})
    return pd.DataFrame(linhas)


def pareto_resumo(abc):
    linhas = []
    for nome, df in abc.items():
        vend = df[df["valor_total"] > 0]
        for faixa in ("A", "B", "C"):
            s = vend[vend["Curva_Valor"] == faixa]
            linhas.append({"Periodo": nome, "Curva_Valor": faixa, "Itens": len(s),
                           "% Itens (vendidos)": len(s) / len(vend) if len(vend) else 0,
                           "Valor": s["valor_total"].sum(),
                           "% Valor": s["pct_valor"].sum()})
    return pd.DataFrame(linhas)


def comparativo(a25, a26):
    c = a25[["key", "part_number", "descricao", "Cod", "valor_total"]].merge(
        a26[["key", "Cod", "valor_total"]], on="key", how="outer", suffixes=("_2025", "_2026"))
    c[["Cod_2025", "Cod_2026"]] = c[["Cod_2025", "Cod_2026"]].fillna("NO HIST")
    d = c["Cod_2026"].map(ORDEM) - c["Cod_2025"].map(ORDEM)
    c["Mudou de curva"] = np.select([d < 0, d > 0], ["SUBIU", "CAIU"], "IGUAL")
    c["Classe_2025"], c["Classe_2026"] = c["Cod_2025"].map(GRUPO), c["Cod_2026"].map(GRUPO)
    c = c.fillna({"valor_total_2025": 0, "valor_total_2026": 0})
    cols = ["key", "part_number", "descricao", "Cod_2025", "Classe_2025", "valor_total_2025",
            "Cod_2026", "Classe_2026", "valor_total_2026", "Mudou de curva"]
    return c[cols].sort_values(["Mudou de curva", "valor_total_2025"], ascending=[True, False])


def cruzar_estoque(a26, est, janela):
    e = a26.merge(est, on="key", how="left").fillna(
        {"estoque": 0, "reservado": 0, "disponivel": 0, "estmin": 0, "estmax": 0})
    media = e["qtd_total"] / janela
    e["media_mensal"] = media
    e["cobertura_meses"] = np.where(media > 0, e["disponivel"] / media.replace(0, np.nan), np.inf)
    fast = e["Cod"].isin(FAST)
    over = fast & ((e["cobertura_meses"] > COBERTURA_OVERSTOCK) |
                   ((e["estmax"] > 0) & (e["disponivel"] > e["estmax"])))
    baixo = fast & ((e["cobertura_meses"] < COBERTURA_BAIXA) |
                    ((e["estmin"] > 0) & (e["disponivel"] < e["estmin"])))
    parado = ~fast & (e["disponivel"] > 0)
    e["Sinal"] = np.select([over, baixo, parado],
                           ["OVERSTOCK – bom p/ promoção", "BAIXO ESTOQUE – risco de ruptura",
                            "PARADO – liquidar/rever"], "")
    e["cobertura_meses"] = e["cobertura_meses"].replace(np.inf, np.nan)
    e = e[(e["disponivel"] > 0) | fast]
    cols = ["key", "part_number", "descricao", "Cod", "Classe", "meses_com_venda", "qtd_total",
            "valor_total", "estoque", "reservado", "disponivel", "estmin", "estmax",
            "media_mensal", "cobertura_meses", "Sinal"]
    return e[cols].sort_values(["Sinal", "valor_total"], ascending=[False, False])


BRL = '"R$" #,##0.00'
PCT = "0.0%"
NUM = "#,##0"


def salvar_excel(path, abas, notas):
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        for nome, df in abas.items():
            df.to_excel(w, sheet_name=nome, index=False)
        pd.DataFrame({"Notas": notas}).to_excel(w, sheet_name="Notas", index=False)
        bold, fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="002C5F")
        for ws in w.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for i, cell in enumerate(ws[1], 1):
                cell.font, cell.fill = bold, fill
                h = str(cell.value or "").lower()
                fmt = (BRL if ("valor" in h or "r$" in h) and "%" not in h
                       else PCT if h.startswith("%") or h.startswith("pct")
                       else "0.0" if "cobertura" in h or "media" in h
                       else NUM if any(t in h for t in ("qtd", "itens", "estoque", "disponivel",
                                                         "reservado", "estm"))
                       else None)
                col = get_column_letter(i)
                if fmt:
                    for c in ws[col][1:]:
                        c.number_format = fmt
                largura = max([len(h)] + [len(str(c.value)) for c in ws[col][1:200] if c.value])
                ws.column_dimensions[col].width = min(max(10, largura + 2), 55)


AZUL, CINZA, LARANJA = RGBColor(0x00, 0x2C, 0x5F), RGBColor(0x59, 0x59, 0x59), RGBColor(0xE6, 0x7E, 0x22)


def brl(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _titulo(prs, texto, sub=None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(0.8)).text_frame
    tb.text = texto
    tb.paragraphs[0].runs[0].font.size, tb.paragraphs[0].runs[0].font.bold = Pt(28), True
    tb.paragraphs[0].runs[0].font.color.rgb = AZUL
    if sub:
        p = tb.add_paragraph()
        p.text = sub
        p.runs[0].font.size, p.runs[0].font.color.rgb = Pt(14), CINZA
    return s


def _tabela(s, df, top=1.4, altura=5.6, fonte=11):
    r, c = df.shape
    t = s.shapes.add_table(r + 1, c, Inches(0.5), Inches(top), Inches(12.3), Inches(altura)).table
    for j, col in enumerate(df.columns):
        t.cell(0, j).text = str(col)
    for i, row in enumerate(df.itertuples(index=False), 1):
        for j, v in enumerate(row):
            t.cell(i, j).text = str(v)
    for i in range(r + 1):
        for j in range(c):
            for p in t.cell(i, j).text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(fonte)


def _bullets(s, itens, top=1.4):
    tf = s.shapes.add_textbox(Inches(0.6), Inches(top), Inches(12), Inches(5.5)).text_frame
    tf.word_wrap = True
    for i, t in enumerate(itens):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = "• " + t
        p.runs[0].font.size = Pt(16)
        p.space_after = Pt(10)


def _fmt_top(df, n=15):
    d = df.head(n)
    return pd.DataFrame({
        "Part Number": d["part_number"], "Descrição": d["descricao"].str[:40],
        "Cód": d["Cod"], "Meses c/ venda": d["meses_com_venda"],
        "Qtd": d["qtd_total"].map(lambda x: f"{x:,.0f}".replace(",", ".")),
        "Valor": d["valor_total"].map(brl)})


def salvar_ppt(path, abc, res, comp, cruz, notas):
    a25, a26 = abc["2025"], abc["2026_YTD"]
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    s = _titulo(prs, "Curva ABC por Frequência de Giro — Peças BMC Hyundai",
                f"2025 (jan–dez, 12 meses) × 2026 YTD (jan–nov, 11 meses) · gerado em "
                f"{date.today():%d/%m/%Y}")
    _bullets(s, [f"{len(a25):,} itens analisados em 2025 e {len(a26):,} em 2026 YTD".replace(",", "."),
                 "Critério: nº de meses distintos com venda (não nº de transações)"], top=3.5)

    s = _titulo(prs, "Metodologia", "Classe pelo nº de meses com ≥1 venda na janela; sem venda na janela → recência da última venda")
    _tabela(s, pd.DataFrame([(c, g, r) for c, g, r in CLASSES], columns=["Cód", "Grupo", "Regra"]),
            top=1.5, altura=5.0, fonte=12)

    s = _titulo(prs, "Distribuição de itens por classe", "2026 YTD com janela de 11 meses — comparação indicativa")
    cd = CategoryChartData()
    cd.categories = [c for c, _, _ in CLASSES]
    for p in ("2025", "2026_YTD"):
        cd.add_series(p.replace("_", " "), res[res["Periodo"] == p]["Itens"].tolist())
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.4),
                            Inches(12.3), Inches(5.8), cd).chart
    ch.has_legend, ch.legend.position, ch.legend.include_in_layout = True, XL_LEGEND_POSITION.TOP, False
    ch.plots[0].has_data_labels = True
    ch.plots[0].data_labels.font.size = Pt(10)
    for ser, cor in zip(ch.series, (AZUL, LARANJA)):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = cor

    s = _titulo(prs, "Participação em valor — Pareto (80/20)", "Eixo X: % dos itens vendidos (ordenados por valor) · Eixo Y: % acumulado do faturamento")
    xy = XyChartData()
    txt = []
    for p, df in abc.items():
        vend = df[df["valor_total"] > 0].reset_index(drop=True)
        if vend.empty:
            continue
        ser = xy.add_series(p.replace("_", " "))
        n = len(vend)
        for i in np.unique(np.linspace(0, n - 1, min(n, 200)).astype(int)):
            ser.add_data_point((i + 1) / n, float(vend.loc[i, "pct_acum"]))
        k80 = int((vend["pct_acum"] < PARETO_A).sum()) + 1
        txt.append(f"{p.replace('_', ' ')}: {k80 / n:.1%} dos itens ({k80}) = 80% do valor")
    if txt:
        ch = s.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS, Inches(0.5), Inches(1.4),
                                Inches(8.3), Inches(5.8), xy).chart
        ch.has_legend, ch.legend.position = True, XL_LEGEND_POSITION.BOTTOM
        for ax in (ch.category_axis, ch.value_axis):
            ax.maximum_scale, ax.minimum_scale = 1.0, 0.0
            ax.tick_labels.number_format, ax.tick_labels.number_format_is_linked = "0%", False
        tf = s.shapes.add_textbox(Inches(9.0), Inches(2.0), Inches(4.0), Inches(4)).text_frame
        tf.word_wrap = True
        for i, t in enumerate(txt):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text, p.runs[0].font.size, p.space_after = t, Pt(16), Pt(14)

    s = _titulo(prs, "Top fast movers (A/B) por valor — 2026 YTD")
    _tabela(s, _fmt_top(a26[a26["Cod"].isin({"A", "B"})]), fonte=10)

    s = _titulo(prs, "Slow / non-moving — itens de atenção",
                "Candidatos a liquidação ou corte (classes D–X, 2026 YTD)")
    if cruz is not None:
        d = cruz[cruz["Sinal"].str.startswith("PARADO")].copy()
        d["valor_estoque_ref"] = d["disponivel"]
        d = d.sort_values("disponivel", ascending=False).head(15)
        _tabela(s, pd.DataFrame({"Part Number": d["part_number"], "Descrição": d["descricao"].str[:40],
                                 "Cód": d["Cod"], "Disponível": d["disponivel"].map("{:,.0f}".format),
                                 "Valor 2026": d["valor_total"].map(brl)}), fonte=10)
    else:
        d = a25[a25["Cod"].isin({"D", "E", "F", "G", "H", "X"})]
        _tabela(s, _fmt_top(d), fonte=10)

    s = _titulo(prs, "Itens que mudaram de curva (2025 → 2026 YTD)")
    cnt = comp["Mudou de curva"].value_counts()
    caiu = comp[comp["Mudou de curva"] == "CAIU"].sort_values("valor_total_2025", ascending=False).head(12)
    tf = s.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(12), Inches(0.4)).text_frame
    tf.text = (f"Subiram: {cnt.get('SUBIU', 0)} · Caíram: {cnt.get('CAIU', 0)} · "
               f"Iguais: {cnt.get('IGUAL', 0)}  — abaixo, maiores quedas por valor 2025")
    tf.paragraphs[0].runs[0].font.size = Pt(14)
    _tabela(s, pd.DataFrame({"Part Number": caiu["part_number"], "Descrição": caiu["descricao"].fillna("").str[:40],
                             "2025": caiu["Cod_2025"], "2026": caiu["Cod_2026"],
                             "Valor 2025": caiu["valor_total_2025"].map(brl),
                             "Valor 2026": caiu["valor_total_2026"].map(brl)}),
            top=1.6, altura=5.4, fonte=10)

    s = _titulo(prs, "Recomendações para campanhas / Black Friday")
    rec = []
    if cruz is not None:
        n_over = (cruz["Sinal"].str.startswith("OVERSTOCK")).sum()
        n_baixo = (cruz["Sinal"].str.startswith("BAIXO")).sum()
        n_par = (cruz["Sinal"].str.startswith("PARADO")).sum()
        rec += [f"Priorizar {n_over} fast movers (A/B/C) com overstock (cobertura > "
                f"{COBERTURA_OVERSTOCK:.0f} meses): giro garantido e liberação de capital.",
                f"Evitar {n_baixo} fast movers com baixo estoque (cobertura < {COBERTURA_BAIXA:.0f} mês): "
                "risco de ruptura — repor antes de promover.",
                f"{n_par} itens slow/non-moving com saldo: liquidação/kit com fast movers ou devolução ao fornecedor."]
    rec += ["Classes C/D com bom valor unitário: testar desconto pontual para acelerar giro.",
            "Itens que CAÍRAM de curva: investigar perda de cliente/concorrência antes de descontar.",
            "Reclassificar mensalmente (janela móvel de 12 meses) e revisar a lista da campanha."]
    _bullets(s, rec)

    s = _titulo(prs, "Notas sobre os dados")
    _bullets(s, notas)
    prs.save(path)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--dados", default=DADOS)
    ap.add_argument("--saida", default=SAIDA)
    ap.add_argument("--notas", nargs="+", default=NOTAS_PADRAO)
    a = ap.parse_args(argv)

    v = carregar_vendas(a.dados, a.notas)
    por_cod = not v["key"].str.contains(r"[A-Za-z\-]").any()
    prod = carregar_produtos(a.dados, por_cod)
    est = carregar_estoque(a.dados)
    m = mensal(v)

    chaves = set(m["key"])
    if est is not None:
        chaves |= set(est.loc[est["disponivel"] > 0, "key"])
    uni = pd.DataFrame({"key": sorted(chaves)}).merge(prod, on="key", how="left")
    uni["part_number"] = uni["part_number"].fillna(uni["key"])
    uni["descricao"] = uni["descricao"].fillna("")

    abc = {n: classificar(m, uni, ref, jan) for n, ref, jan in PERIODOS}
    ini, fim = v["mes"].min(), v["mes"].max()
    notas = [f"Arquivo(s) de vendas: {', '.join(a.notas)} · {len(v):,} linhas".replace(",", "."),
             f"Histórico disponível: {ini} a {fim}",
             "2026 YTD: janela de 11 meses (jan–nov/2026), mesmos limiares — comparação indicativa."]
    for n, ref, _ in PERIODOS:
        if (ref - ini).n < MAX_H:
            notas.append(f"{n}: histórico cobre só {(ref - ini).n + 1} meses antes da referência "
                         f"{ref}; classes F/G/H/X podem estar subestimadas (aparecem como E/NO HIST).")
    if fim < PERIODOS[-1][1]:
        notas.append(f"Último mês com venda na base: {fim} (referência 2026 = {PERIODOS[-1][1]}).")
    if est is None:
        notas.append("Estoque não encontrado: aba Estoque_Cruzado omitida.")

    res = resumo(abc)
    comp = comparativo(abc["2025"], abc["2026_YTD"])
    cruz = cruzar_estoque(abc["2026_YTD"], est, PERIODOS[-1][2]) if est is not None else None

    abas = {"Base_Mensal": base_mensal(m, prod, PERIODOS[0][1] - (MAX_H - 1), PERIODOS[-1][1]),
            "ABC_2025": abc["2025"], "ABC_2026_YTD": abc["2026_YTD"],
            "Comparativo": comp, "Resumo": res, "Pareto": pareto_resumo(abc)}
    if cruz is not None:
        abas["Estoque_Cruzado"] = cruz

    out = Path(a.saida)
    out.mkdir(parents=True, exist_ok=True)
    xlsx, pptx = out / "ABC_Pecas_BMC.xlsx", out / "ABC_Pecas_BMC.pptx"
    salvar_excel(xlsx, abas, notas)
    salvar_ppt(pptx, abc, res, comp, cruz, notas)
    print("\n".join(notas))
    print(res.pivot(index="Cod", columns="Periodo", values="Itens").reindex(ORDEM).to_string())
    print(f"\nExcel: {xlsx}\nPowerPoint: {pptx}")


if __name__ == "__main__":
    sys.exit(main())
