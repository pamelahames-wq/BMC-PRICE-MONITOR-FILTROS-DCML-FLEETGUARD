"""Curva ABC de peças e de RAPs pelo faturamento de 2026.

Uso:  python -m rap.abc   →  saida/rap/Curva_ABC_2026.xlsx + abc.json
Depois: node rap/abc_deck.js  →  saida/rap/Curva_ABC_2026.pptx
Peças: NFs de venda − devoluções dos centros de peças (Balcão + RAP), 05/01 a 16/10/2026.
RAPs: as RAPs válidas da reestruturação (vendedor RAP + categoria Standard/Ouro/Diamante).
Corte: A até 80% do faturamento acumulado, B até 95%, C o restante (editável na aba Premissas).
"""
import json

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

from . import analise as A
from .saidas import BOLD, INPUT, INT, LINK, PCT, RS, RS0, TITULO, _cab, _cod, _filtro, _fmt

SAIDA = A.RAIZ / "saida" / "rap"
ARQ = SAIDA / "Curva_ABC_2026.xlsx"
INICIO_ANO = "2026-01-01"
CORTE_A, CORTE_B = 0.80, 0.95
P_A, P_B = "Premissas!$B$4", "Premissas!$B$5"


def classe(acum):
    """A/B/C pelo acumulado antes do item: o item que cruza 80% ainda é A."""
    antes = acum.shift(fill_value=0)
    return np.where(antes < CORTE_A, "A", np.where(antes < CORTE_B, "B", "C"))


def pecas(it, cu, pro):
    v = it[(it.DT >= INICIO_ANO) & (it.DT <= A.FIM)].copy()
    v = A.custo_itens(v, cu)
    v["nf"] = v.BRUTO * v.sinal
    v["imp"] = v.IMPOSTOS * v.sinal
    v["cst"] = v.custo_unit.fillna(0) * v.QTDNEG * v.sinal
    v["nf_rap"] = v.nf.where(v.CODCENCUS == "1030207", 0)
    v["sem_custo"] = v.origem_custo == "s/ custo"
    p = v.groupby("CODPROD").agg(nf=("nf", "sum"), nf_rap=("nf_rap", "sum"), imp=("imp", "sum"),
                                 cst=("cst", "sum"), qtd=("QTDNEG", "sum"), nfs=("NUNOTA", "nunique"),
                                 sem_custo=("sem_custo", "any"))
    p = p[p.nf > 0].sort_values("nf", ascending=False).reset_index()
    p = p.merge(pro[["CODPROD", "REFERENCIA", "DESCRPROD", "CODGRUPOPROD", "grupo"]], on="CODPROD", how="left")
    p["acum"] = p.nf.cumsum() / p.nf.sum()
    p["classe"] = classe(p.acum)
    p["margem"] = ((p.nf - p.imp - p.cst) / p.nf).where(~p.sem_custo)
    return p, v.DT.min(), v.DT.max()


def raps(fat):
    r = A.rodar()
    c = r["classificacao"].copy()
    f = fat[(fat.DT >= INICIO_ANO) & (fat.DT <= A.FIM)]
    ano = (f.Valor * f.sinal).groupby(f.CODPARC).sum()
    c["fat_ano"] = c.CODPARC.map(ano).fillna(0)
    c = c.sort_values("fat_ano", ascending=False).reset_index(drop=True)
    pos = c[c.fat_ano > 0].copy()
    pos["acum"] = pos.fat_ano.cumsum() / pos.fat_ano.sum()
    pos["classe"] = classe(pos.acum)
    c = c.merge(pos[["CODPARC", "acum", "classe"]], on="CODPARC", how="left")
    c["classe"] = c.classe.fillna("Sem compra")
    c["consultor"] = c.consultor.str.replace(" (ELISASILVA)", "", regex=False)
    return c


def premissas(wb, ini, fim):
    ws = wb.active
    ws.title = "Premissas"
    ws["A1"] = "Curva ABC 2026 — premissas"
    ws["A1"].font = TITULO
    ws["A2"] = "Células em azul são editáveis: mudar os cortes reclassifica peças e RAPs."
    ws["A2"].font = Font(name="Arial", italic=True)
    for i, (rot, val, nota) in enumerate([
        ("Corte da classe A (faturamento acumulado)", CORTE_A, "Itens que somam os primeiros 80% do faturamento"),
        ("Corte da classe B (faturamento acumulado)", CORTE_B, "Dos 80% aos 95%; o restante é C"),
    ], 4):
        ws.cell(i, 1, rot).font = Font(name="Arial")
        _fmt(ws, i, 2, val, PCT, INPUT)
        ws.cell(i, 3, nota).font = Font(name="Arial", italic=True, color="595959")
    notas = [
        f"Período: NFs de {ini:%d/%m/%Y} a {fim:%d/%m/%Y} (última NF na base do Sankhya).",
        "Faturamento = valor da NF (f_Notas.BRUTO) das vendas menos devoluções, só centros de peças "
        "(Balcão …204 e RAP 1030207); peças com saldo zero ou negativo ficam fora da curva.",
        "Margem = (NF − impostos − custo) ÷ NF, mesma conta da Reestruturação de Margem. Custo da NF; se zerado, "
        "último custo do histórico f_CustoUnitarioProd. 's/ custo' quando não há custo em nenhum dos dois.",
        "RAPs: as 545 RAPs válidas da reestruturação (vendedor preferencial RAP + categoria Standard/Ouro/Diamante), "
        "por CODPARC (matriz e filial separadas). RAPs sem faturamento líquido em 2026 (sem compra ou com devolução "
        "maior que a compra) ficam como 'Sem compra'.",
        "Classe pelo acumulado antes do item: o item que cruza o corte ainda entra na classe de cima.",
    ]
    ws["A8"] = "Regras"
    ws["A8"].font = BOLD
    for k, t in enumerate(notas, 9):
        ws.cell(k, 1, "• " + t).font = Font(name="Arial")
    ws.column_dimensions["A"].width = 48
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 50


def aba_curva(wb, nome, linhas, cols_fixas, larg, valor_col, extras):
    """Escreve a curva com fórmulas: % do total, acumulado corrido, % acumulado e classe."""
    ws = wb.create_sheet(nome)
    n0 = 2
    cols = ["#"] + cols_fixas + ["Faturamento 2026 (R$)", "% do total", "Acumulado (R$)", "% acumulado", "Classe"] + \
        [e[0] for e in extras]
    _cab(ws, 1, cols, [6] + larg + [16, 9, 16, 10, 8] + [e[1] for e in extras])
    k = len(cols_fixas) + 2                      # coluna do faturamento
    L = lambda c: chr(64 + c)  # noqa: E731
    F, P, AC, PA, CL = (L(k + j) for j in range(5))
    ult = n0 + len(linhas) - 1
    for i, x in enumerate(linhas, n0):
        _fmt(ws, i, 1, i - 1, INT)
        for j, v in enumerate(x["fixas"], 2):
            _fmt(ws, i, j, v)
        _fmt(ws, i, k, round(x[valor_col], 2), RS, INPUT)
        _fmt(ws, i, k + 1, f"={F}{i}/SUM(${F}${n0}:${F}${ult})", PCT)
        _fmt(ws, i, k + 2, f"={F}{i}" if i == n0 else f"={AC}{i - 1}+{F}{i}", RS0)
        _fmt(ws, i, k + 3, f"={AC}{i}/${AC}${ult}", PCT)
        antes = "0" if i == n0 else f"{PA}{i - 1}"
        _fmt(ws, i, k + 4, f'=IF({antes}<{P_A},"A",IF({antes}<{P_B},"B","C"))', None, BOLD)
        for j, (_, _, chave, fmt) in enumerate(extras, k + 5):
            v = x.get(chave)
            _fmt(ws, i, j, v if not (isinstance(v, float) and np.isnan(v)) else "s/ custo", fmt)
    _filtro(ws, 1, ult, len(cols))
    ws.freeze_panes = "C2"
    return ult, F, CL


def resumo(wb, ult_p, Fp, CLp, ult_r, Fr, CLr, p, c):
    ws = wb.create_sheet("Resumo", 1)
    ws["A1"] = "Curva ABC 2026 — resumo"
    ws["A1"].font = TITULO
    ref = lambda aba, col, ult: f"'{aba}'!${col}$2:${col}${ult}"  # noqa: E731
    blocos = [("Peças", "ABC Peças", ult_p, Fp, CLp, "Peças"), ("RAPs", "ABC RAPs", ult_r, Fr, CLr, "RAPs")]
    linha = 3
    for titulo, aba, ult, F, CL, rot in blocos:
        ws.cell(linha, 1, titulo).font = BOLD
        _cab(ws, linha + 1, ["Classe", f"Nº de {rot}", f"% das {rot}", "Faturamento 2026 (R$)",
                             "% do faturamento"], [12, 14, 12, 20, 14])
        r0 = linha + 2
        for j, cl in enumerate("ABC"):
            i = r0 + j
            ws.cell(i, 1, cl).font = BOLD
            _fmt(ws, i, 2, f'=COUNTIF({ref(aba, CL, ult)},A{i})', INT)
            _fmt(ws, i, 3, f"=B{i}/B${r0 + 3}", PCT)
            _fmt(ws, i, 4, f'=SUMIFS({ref(aba, F, ult)},{ref(aba, CL, ult)},A{i})', RS0)
            _fmt(ws, i, 5, f"=D{i}/D${r0 + 3}", PCT)
        t = r0 + 3
        ws.cell(t, 1, "Total").font = BOLD
        for col, fmt in (("B", INT), ("C", PCT), ("D", RS0), ("E", PCT)):
            _fmt(ws, t, " ABCDE".index(col), f"=SUM({col}{r0}:{col}{t - 1})", fmt, BOLD)
        linha = t + 3
    sem = int((c.classe == "Sem compra").sum())
    ws.cell(linha - 2, 1, f"Mais {sem} RAPs sem faturamento líquido em 2026 (fora da curva; listadas no fim da aba 'ABC RAPs').").font = \
        Font(name="Arial", italic=True)

    # RAPs: classe ABC × categoria atual
    ws.cell(linha, 1, "RAPs — classe ABC × categoria atual (nº de RAPs)").font = BOLD
    _cab(ws, linha + 1, ["Classe"] + A.CATEGORIAS + ["Total"])
    for j, cl in enumerate(["A", "B", "C", "Sem compra"]):
        i = linha + 2 + j
        ws.cell(i, 1, cl).font = BOLD
        for k, cat in enumerate(A.CATEGORIAS, 2):
            _fmt(ws, i, k, int(((c.classe == cl) & (c.categoria == cat)).sum()), INT)
        _fmt(ws, i, 5, f"=SUM(B{i}:D{i})", INT, BOLD)

    # peças: classe ABC × grupo (top grupos da classe A)
    linha += 8
    ws.cell(linha, 1, "Peças classe A — grupos que mais faturam").font = BOLD
    _cab(ws, linha + 1, ["CODGRUPOPROD", "Grupo", "Peças A", "Faturamento (R$)", "Margem"], [14, 40, 10, 18, 10])
    a = p[p.classe == "A"]
    g = a.groupby(["CODGRUPOPROD", "grupo"], dropna=False).agg(n=("nf", "size"), nf=("nf", "sum"),
                                                                imp=("imp", "sum"), cst=("cst", "sum"))
    g = g.sort_values("nf", ascending=False).head(15).reset_index()
    for j, x in enumerate(g.itertuples(), linha + 2):
        _fmt(ws, j, 1, _cod(x.CODGRUPOPROD) if isinstance(x.CODGRUPOPROD, str) else "")
        _fmt(ws, j, 2, x.grupo if isinstance(x.grupo, str) else "(sem grupo)")
        _fmt(ws, j, 3, int(x.n), INT)
        _fmt(ws, j, 4, round(x.nf, 2), RS0)
        _fmt(ws, j, 5, (x.nf - x.imp - x.cst) / x.nf, PCT)
    for col, w in zip("ABCDE", [14, 40, 14, 20, 14]):
        ws.column_dimensions[col].width = w


def gerar():
    cli, fat, it, pro, pre, cu = A.carregar()
    p, ini, fim = pecas(it, cu, pro)
    c = raps(fat)
    SAIDA.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    premissas(wb, ini, fim)
    lin_p = [{"fixas": [x.REFERENCIA, int(x.CODPROD), x.DESCRPROD, _cod(x.CODGRUPOPROD) if isinstance(x.CODGRUPOPROD, str) else "",
                        x.grupo if isinstance(x.grupo, str) else ""],
              "nf": x.nf, "nf_rap": round(x.nf_rap, 2), "pct_rap": x.nf_rap / x.nf, "qtd": x.qtd, "nfs": int(x.nfs),
              "margem": x.margem} for x in p.itertuples()]
    ult_p, Fp, CLp = aba_curva(wb, "ABC Peças", lin_p, ["PN", "CODPROD", "Descrição", "CODGRUPOPROD", "Grupo"],
                               [16, 9, 34, 12, 30], "nf",
                               [("Vendido para RAP (R$)", 15, "nf_rap", RS), ("% RAP", 8, "pct_rap", PCT),
                                ("Qtd vendida", 10, "qtd", INT), ("Nº de NFs", 9, "nfs", INT),
                                ("Margem", 9, "margem", PCT)])
    cp = c[c.classe != "Sem compra"]
    lin_r = [{"fixas": [x.NOMEPARC, x.CNPJ, x.consultor, x.categoria],
              "fat_ano": x.fat_ano, "fat_1s": round(x.fat_1s, 2), "fat_2s": round(x.fat, 2), "balde": x.balde}
             for x in cp.itertuples()]
    ult_r, Fr, CLr = aba_curva(wb, "ABC RAPs", lin_r, ["RAP", "CNPJ", "Consultor", "Categoria"],
                               [40, 17, 13, 11], "fat_ano",
                               [("1º sem/2026 (R$)", 15, "fat_1s", RS), ("2º sem até 16/10 (R$)", 16, "fat_2s", RS),
                                ("Balde da reestruturação", 16, "balde", None)])
    ws = wb["ABC RAPs"]
    sem = c[c.classe == "Sem compra"]
    i0 = ult_r + 3
    ws.cell(i0, 2, f"Sem faturamento líquido em 2026 ({len(sem)} RAPs)").font = BOLD
    for i, x in enumerate(sem.sort_values(["categoria", "NOMEPARC"]).itertuples(), i0 + 1):
        for j, v in enumerate([x.NOMEPARC, x.CNPJ, x.consultor, x.categoria], 2):
            _fmt(ws, i, j, v)
    resumo(wb, ult_p, Fp, CLp, ult_r, Fr, CLr, p, c)
    wb.save(ARQ)
    (SAIDA / "abc.json").write_text(json.dumps(dados_deck(p, c, ini, fim), ensure_ascii=False, indent=1, default=float))
    return p, c


def dados_deck(p, c, ini, fim):
    def tab(df, val):
        g = df.groupby("classe").agg(n=(val, "size"), v=(val, "sum"))
        return {k: {"n": int(g.loc[k, "n"]), "v": float(g.loc[k, "v"])} for k in "ABC" if k in g.index}

    a = p[p.classe == "A"]
    grp = a.groupby("grupo", dropna=False).agg(n=("nf", "size"), nf=("nf", "sum"), imp=("imp", "sum"), cst=("cst", "sum"))
    grp = grp.sort_values("nf", ascending=False).head(8)
    cp = c[c.classe != "Sem compra"]
    pm = lambda d: float((d.nf.sum() - d.imp.sum() - d.cst[~d.sem_custo].sum()) / d.nf.sum())  # noqa: E731
    return {
        "periodo": f"{ini:%d/%m/%Y} a {fim:%d/%m/%Y}",
        "pecas": {"total_n": int(len(p)), "total_v": float(p.nf.sum()), "classes": tab(p, "nf"),
                  "margem": {k: pm(p[p.classe == k]) for k in "ABC"},
                  "rap": {k: float(p[p.classe == k].nf_rap.sum() / p[p.classe == k].nf.sum()) for k in "ABC"},
                  "top": [{"pn": x.REFERENCIA, "desc": x.DESCRPROD, "grupo": x.grupo, "nf": x.nf, "pct_rap": x.nf_rap / x.nf,
                           "m": None if np.isnan(x.margem) else x.margem} for x in p.head(15).itertuples()],
                  "grupos_a": [{"g": (k if isinstance(k, str) else "(sem grupo)"), "n": int(x.n), "nf": x.nf,
                                "m": (x.nf - x.imp - x.cst) / x.nf} for k, x in grp.iterrows()],
                  "curva": [float(p.acum.iloc[min(len(p) - 1, int(len(p) * q / 100))]) for q in range(0, 101, 5)]},
        "raps": {"total_n": int(len(cp)), "total_v": float(cp.fat_ano.sum()), "sem_compra": int((c.classe == "Sem compra").sum()),
                 "classes": tab(cp, "fat_ano"),
                 "cat": {cl: {cat: int(((c.classe == cl) & (c.categoria == cat)).sum()) for cat in A.CATEGORIAS}
                         for cl in ["A", "B", "C", "Sem compra"]},
                 "top": [{"nome": x.NOMEPARC, "cons": x.consultor, "cat": x.categoria, "v": x.fat_ano, "balde": x.balde}
                         for x in cp.head(12).itertuples()],
                 "a_rebaixar": int(((c.classe == "A") & (c.balde == "REBAIXAR")).sum()),
                 "a_desh": int(((c.classe == "A") & (c.balde == "DESHOMOLOGAR")).sum()),
                 "b_desh": int(((c.classe == "B") & (c.balde == "DESHOMOLOGAR")).sum())},
    }


if __name__ == "__main__":
    p, c = gerar()
    print(ARQ)
    print(p.groupby("classe").agg(n=("nf", "size"), nf=("nf", "sum")))
    print(c.groupby("classe").agg(n=("fat_ano", "size"), v=("fat_ano", "sum")))
    print(pd.crosstab(c.classe, c.balde))
