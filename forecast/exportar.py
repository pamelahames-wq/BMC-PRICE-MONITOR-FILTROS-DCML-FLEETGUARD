"""Gera data.json (PPT) e a planilha de acompanhamento."""
import json, sys
import numpy as np
import pandas as pd
sys.path.insert(0, __file__.rsplit("/", 1)[0])
from base import S

g = pd.read_pickle(f"{S}/out/pn_full.pkl")
fam = pd.read_pickle(f"{S}/out/parque_fam.pkl")
FAMS = ["Escavadeira", "Pá carregadeira", "Empilhadeira", "Retroescavadeira"]

MANUAL = {  # inclusão manual: baixa venda, mas sempre falta
    "500119-00624A": "Vidro HL635", "500119-00625A": "Vidro HL635", "500119-00611A": "Vidro HL635",
    "71LN-26110": "Vidro HL635", "SIBU-00008": "Filtro HB640C", "SIBU-03977": "Filtro HB640C",
    "ZUAC-01448": "Correia HB640C", "ZUAC-00027": "Motor HB640C", "ZUAC-00786": "Bomba alta pressão HB640C",
    "400401-00304": "Filtro de ar HL635",
}
from base import carregar
_p, _n, _e = carregar()
_ev = _e[_e.CODLOCAL.isin(["600101", "600102"])].merge(_p[["CODPROD", "PN", "DESCRPROD"]], on="CODPROD")
_ev["DISPONIVEL"] = _ev.DISPONIVEL.clip(lower=0)
for pn in MANUAL:
    if pn not in g.index:
        x = _ev[_ev.PN == pn]
        desc = x.DESCRPROD.iloc[0] if len(x) else _p[_p.PN == pn].DESCRPROD.head(1).squeeze() if (_p.PN == pn).any() else MANUAL[pn]
        g.loc[pn, ["DESCRPROD", "curva", "familia", "modelos", "categoria"]] = [desc, "—", "Pá carregadeira" if "HL635" in MANUAL[pn] else "Retroescavadeira", "HL635" if "HL635" in MANUAL[pn] else "HB640C", "Filtros" if "Filtro" in MANUAL[pn] else "Cabine e vidros" if "Vidro" in MANUAL[pn] else "Motor"]
        g.loc[pn, ["qtd", "valor", "qtd_mes", "meses_venda", "forecast_90d", "necessidade", "pedido_pacote", "notas"]] = 0
        g.loc[pn, "est_fab"] = x[x.CODLOCAL == "600101"].DISPONIVEL.sum()
        g.loc[pn, "est_bmc"] = x[x.CODLOCAL == "600102"].DISPONIVEL.sum()
        g.loc[pn, "estoque"] = g.loc[pn, "est_fab"] + g.loc[pn, "est_bmc"]
        g.loc[pn, ["cobertura_meses", "status", "prioridade", "recorrente", "alto_giro", "sem_estoque"]] = [np.inf, "Sem estoque" if g.loc[pn, "estoque"] <= 0 else "OK", "", False, False, g.loc[pn, "estoque"] <= 0]
g["manual"] = g.index.map(MANUAL)
g["venda_mes_rs"] = g.valor / 10
g["preco_medio"] = np.where(g.qtd > 0, g.valor / g.qtd, 0)
g["valor_necessidade"] = g.necessidade * g.preco_medio

crit = g[(g.prioridade != "") | g.manual.notna()].copy()
crit["prioridade"] = crit.prioridade.replace("", "Manual")
crit.loc[crit.manual.notna() & (crit.prioridade == ""), "prioridade"] = "Manual"
ordem = {"P1": 0, "P2": 1, "P3": 2, "Manual": 3}
crit = crit.sort_values(["prioridade", "venda_mes_rs"], key=lambda s: s.map(ordem) if s.name == "prioridade" else -s)

# ---------- planilha ----------
cols = {
    "prioridade": "Prioridade", "curva": "Curva", "familia": "Família", "modelos": "Modelos",
    "categoria": "Categoria", "DESCRPROD": "Descrição", "qtd_mes": "Venda/mês (un)",
    "meses_venda": "Meses c/ venda (de 10)", "est_fab": "Estoque fábrica HHIBP", "est_bmc": "Estoque BMC",
    "cobertura_meses": "Cobertura (meses)", "forecast_90d": "Forecast 90 dias (un)",
    "necessidade": "Falta p/ 90 dias (un)", "pedido_pacote": "Qtd no pacote RAP",
}
t = crit[list(cols)].rename(columns=cols)
t.insert(5, "Part number", crit.index)
t["Observação"] = crit.manual.fillna("")
t["Cobertura (meses)"] = t["Cobertura (meses)"].replace(np.inf, np.nan).round(1)
t["Venda/mês (un)"] = t["Venda/mês (un)"].round(1)
for c in ["Pedido à fábrica (qtd)", "Data do pedido", "ETA informado pela fábrica", "Status / retorno da reunião"]:
    t[c] = ""

abc = g[["curva", "familia", "modelos", "categoria", "DESCRPROD", "qtd", "valor", "notas", "meses_venda",
         "est_fab", "est_bmc", "cobertura_meses", "status", "forecast_90d", "pedido_pacote"]].copy()
abc.insert(0, "Part number", abc.index)
abc["cobertura_meses"] = abc.cobertura_meses.replace(np.inf, np.nan).round(1)
abc.columns = ["Part number", "Curva", "Família", "Modelos", "Categoria", "Descrição", "Qtd 2026", "Faturamento 2026 (R$)",
               "Notas", "Meses c/ venda", "Estoque fábrica HHIBP", "Estoque BMC", "Cobertura (meses)", "Status",
               "Forecast 90 dias", "Qtd no pacote RAP"]

leia = pd.DataFrame({"Item": [
    "Base", "Curva ABC", "Recorrente", "Prioridade P1", "Prioridade P2", "Prioridade P3", "Manual",
    "Estoque fábrica", "Estoque BMC", "Forecast 90 dias", "Família / modelo", "Como usar"],
 "Regra": [
    "Sankhya: vendas Balcão + RAP de 05/01 a 29/10/2026 (menos devoluções); estoque de 30/10/2026. Só peças Hyundai.",
    "Por faturamento 2026: A = 80% do faturamento, B = próximos 15%, C = últimos 5%.",
    "Vendeu em 4 ou mais dos 10 meses de 2026.",
    "Curva A, recorrente, cobertura menor que 1 mês (estoque fábrica + BMC ÷ venda média mensal).",
    "Curva B, recorrente, cobertura menor que 1 mês.",
    "Curva A, recorrente, cobertura entre 1 e 3 meses (vai faltar no prazo de importação).",
    "Incluída pelo time: vende pouco, mas falta sempre que o cliente precisa (vidros HL635, itens HB640C).",
    "Local 600101 HHIBP Virtual Diário (estoque da HD Hyundai disponível para a BMC).",
    "Local 600102 das filiais BMC.",
    "Maior entre a média mensal de 2026 e a média de ago–out × 3 meses × crescimento do parque da família (máquinas novas no SAB em 2026).",
    "Plano de manutenção oficial do modelo; sem plano, pela frota SAB dos clientes que compraram a peça (≥60% numa família).",
    "Preencher na aba Peças críticas as colunas Pedido, Data, ETA e Status a cada reunião com a fábrica."]})

xl = f"{S}/out/Forecast_Curva_ABC_Pecas_BMC_Hyundai.xlsx"
with pd.ExcelWriter(xl, engine="openpyxl") as w:
    t.to_excel(w, sheet_name="Peças críticas", index=False)
    abc.to_excel(w, sheet_name="Curva ABC completa", index=False)
    leia.to_excel(w, sheet_name="Leia-me", index=False)
print(xl, len(t), len(abc))

# ---------- números do PPT ----------
rec = g[g.recorrente]
def cnt(df):
    return {"OK": int((df.status == "OK").sum()), "Atenção": int((df.status == "Atenção (1–3 meses)").sum()),
            "Crítico": int((df.status == "Crítico (<1 mês)").sum()), "Sem estoque": int((df.status == "Sem estoque").sum())}
d = {}
d["abc"] = {c: cnt(rec[rec.curva == c]) for c in "ABC"}
d["abc_pns"] = {c: int((g.curva == c).sum()) for c in "ABC"}
d["abc_valor"] = {c: float(g[g.curva == c].valor.sum()) for c in "ABC"}
d["n_rec"] = int(len(rec))
p = crit[crit.prioridade.isin(["P1", "P2", "P3"])]
d["prio"] = {k: int((crit.prioridade == k).sum()) for k in ordem}
d["p1_venda_mes"] = float(crit[crit.prioridade == "P1"].venda_mes_rs.sum())
d["p12_venda_mes"] = float(crit[crit.prioridade.isin(["P1", "P2"])].venda_mes_rs.sum())
d["alto_giro_zero"] = int((g.alto_giro & g.sem_estoque).sum())
d["p_no_pacote"] = int((p.pedido_pacote > 0).sum())
d["p_fora_pacote"] = int((p.pedido_pacote == 0).sum())
d["fab_zero_rec"] = int((rec.est_fab <= 0).sum())
cats = ["Hidráulico (bombas, motores, válvulas)", "Vedação", "Filtros", "Cabine e vidros", "Transmissão e eixos",
        "Motor", "Coxins e suportes", "Pinos, buchas e rolamentos", "Mangueiras", "Elétrica e sensores",
        "Arrefecimento e A/C", "Material rodante", "Outras peças"]
pp = crit[crit.prioridade.isin(["P1", "P2", "P3"])].copy()
pp["cat2"] = np.where(pp.categoria.isin(cats), pp.categoria, "Outras peças")
pp["fam2"] = np.where(pp.familia.isin(FAMS), pp.familia, "Várias")
m = pd.crosstab(pp.cat2, pp.fam2).reindex(index=cats, columns=FAMS + ["Várias"], fill_value=0)
d["matriz"] = {"linhas": cats, "colunas": FAMS + ["Várias"], "valores": m.values.tolist()}

def linhas(df, n):
    out = []
    for pn, r in df.head(n).iterrows():
        out.append({"pn": pn, "desc": str(r.DESCRPROD).title()[:34], "fam": r.familia if r.familia in FAMS else "Várias",
                    "mod": (r.modelos or "—")[:20], "vm": round(float(r.qtd_mes), 1), "meses": int(r.meses_venda),
                    "fab": int(r.est_fab), "bmc": int(r.est_bmc), "f90": int(r.forecast_90d), "pac": int(r.pedido_pacote),
                    "curva": r.curva, "prio": r.prioridade, "obs": r.manual if isinstance(r.manual, str) else ""})
    return out
d["p1"] = linhas(crit[crit.prioridade == "P1"].sort_values("venda_mes_rs", ascending=False), 14)
man = crit[crit.manual.notna()]
d["manual"] = linhas(man.sort_values(["manual", "est_fab"]), 12)
d["parque"] = {f: {"parque": int(fam.loc[f, "parque"]), "novas_2026": int(fam.loc[f, "novas_2026"])} for f in FAMS}
fc = []
for f in FAMS:
    x = pp[pp.fam2 == f]
    fc.append({"fam": f, "pns": int(len(x)), "f90": int(x.forecast_90d.sum()), "estoque": int(x.estoque.sum()),
               "falta": int(x.necessidade.sum()), "falta_rs": float(x.valor_necessidade.sum()),
               "fator": float(g[g.familia == f].fator_parque.max() if (g.familia == f).any() else 1)})
d["forecast"] = fc
x = pp[pp.fam2 == "Várias"]
d["forecast_varias"] = {"pns": int(len(x)), "f90": int(x.forecast_90d.sum()), "falta": int(x.necessidade.sum())}
d["falta_total_un"] = int(pp.necessidade.sum()); d["falta_total_rs"] = float(pp.valor_necessidade.sum())
json.dump(d, open(f"{S}/out/data.json", "w"), ensure_ascii=False, indent=1)
print(json.dumps({k: v for k, v in d.items() if k not in ("p1", "manual", "matriz")}, ensure_ascii=False, indent=1))
print(m)
