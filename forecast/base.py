"""Forecast Curva ABC para a fábrica Hyundai.

Uso: FORECAST_DIR=<pasta com raw/, upload/pacote.xlsx e out/> python forecast/exportar.py
(antes: python forecast/analise.py). raw/ = bases do Breno (f_Notas, f_estoque, d_Produtos, f_Fat_Clientes,
d_Clientes, sab_cliente, sab_atividade, sab_equipamento) + planos_oficiais.json do motor.

Tabela por PN (família de código): venda 2026, recorrência, estoque vendável, terceiro, ABC."""
import os
import re
import numpy as np
import pandas as pd

S = os.environ.get("FORECAST_DIR", os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "cache", "forecast"))
R = f"{S}/raw"
LOC_VEND = {"600101", "600102"}
LOC_TERC = {"600105", "50000000"}
SUF = re.compile(r"(-AS|\.BR|BR|VM|-SE|SE)$")


def base_pn(ref):
    ref = ("" if ref is None or ref != ref else str(ref)).strip().upper()
    m = re.match(r"^([0-9A-Z]{4}-[0-9A-Z]{5})(.*)$", ref)
    if m and SUF.fullmatch(m.group(2) or "x"):
        return m.group(1)
    return ref


def carregar():
    p = pd.read_csv(f"{R}/d_Produtos.csv", dtype=str)
    p.columns = [c.strip("[]") for c in p.columns]
    p["REFERENCIA"] = p.REFERENCIA.fillna("").str.strip().str.upper()
    # "USE XXXX-XXXXX" na descrição = código substituído
    use = p.DESCRPROD.fillna("").str.extract(r"^USE\s+([0-9A-Z]{4}-[0-9A-Z]{5}\S*)")[0]
    p["PN"] = np.where(use.notna(), use.fillna("").map(base_pn), p.REFERENCIA.map(base_pn))

    n = pd.read_csv(f"{R}/f_Notas.csv", dtype=str)
    n.columns = [c.strip("[]") for c in n.columns]
    for c in ["QTDNEG", "BRUTO", "CUSTO"]:
        n[c] = pd.to_numeric(n[c], errors="coerce").fillna(0)
    n["cc"] = n.CODCENCUS.str[-6:]
    n = n[n.cc.isin(["030204", "030207"])].copy()        # Balcão + RAP (fora máquinas)
    sinal = np.where(n.TIPMOV == "D", -1, 1)
    n["qtd"] = n.QTDNEG * sinal
    n["valor"] = n.BRUTO * sinal
    n["mes"] = n.DTFATUR.str[:7]
    n["canal"] = np.where(n.cc == "030207", "RAP", "Balcão")

    e = pd.read_csv(f"{R}/f_estoque.csv", dtype=str)
    e.columns = [c.split("[")[-1].strip("]").replace("tmp_Locais.", "L_") for c in e.columns]
    for c in ["DISPONIVEL", "ESTOQUE", "ESTMIN"]:
        e[c] = pd.to_numeric(e[c], errors="coerce").fillna(0)
    return p, n, e


def tabela_pn():
    p, n, e = carregar()
    n = n.merge(p[["CODPROD", "PN", "DESCRPROD", "Grupo Nivel 1", "Grupo Nivel 3"]], on="CODPROD", how="left")
    meses = n.mes.nunique()
    v = n[n.TIPMOV == "V"]
    g = n.groupby("PN").agg(qtd=("qtd", "sum"), valor=("valor", "sum"), custo=("CUSTO", "sum"))
    g["notas"] = v.groupby("PN").NUNOTA.nunique()
    g["meses_venda"] = v.groupby("PN").mes.nunique()
    g["clientes_rap"] = v[v.canal == "RAP"].groupby("PN").NUNOTA.nunique()
    g = g.fillna(0)
    g["qtd_mes"] = g.qtd / meses
    rec = n[n.DTFATUR >= "2026-08-01"]
    g["qtd_mes_rec"] = (rec.groupby("PN").qtd.sum() / 3).reindex(g.index).fillna(0)

    ev = e[e.CODLOCAL.isin(LOC_VEND)].merge(p[["CODPROD", "PN"]], on="CODPROD", how="left")
    ev["DISPONIVEL"] = ev.DISPONIVEL.clip(lower=0)
    g["est_fab"] = ev[ev.CODLOCAL == "600101"].groupby("PN").DISPONIVEL.sum()
    g["est_bmc"] = ev[ev.CODLOCAL == "600102"].groupby("PN").DISPONIVEL.sum()
    g[["est_fab", "est_bmc"]] = g[["est_fab", "est_bmc"]].fillna(0)
    g["estoque"] = g.est_fab + g.est_bmc
    g["est_min"] = e[e.CODLOCAL == "600101"].merge(p[["CODPROD", "PN"]], on="CODPROD").groupby("PN").ESTMIN.sum()
    et = e[e.CODLOCAL.isin(LOC_TERC)].merge(p[["CODPROD", "PN"]], on="CODPROD", how="left")
    g["terceiro"] = g.index.isin(et.PN.unique())
    g["terceiro_qtd"] = et.groupby("PN").ESTOQUE.sum()
    g = g.fillna({"estoque": 0, "est_min": 0, "terceiro_qtd": 0})

    # descrição/grupo: do código mais vendido da família
    d = (n.groupby(["PN", "DESCRPROD", "Grupo Nivel 1", "Grupo Nivel 3"]).valor.sum().reset_index()
         .sort_values("valor", ascending=False).drop_duplicates("PN").set_index("PN"))
    g = g.join(d[["DESCRPROD", "Grupo Nivel 1", "Grupo Nivel 3"]])
    g = g[(g["Grupo Nivel 1"] == "PECAS HYUNDAI") & (g.valor > 0)].copy()

    g = g.sort_values("valor", ascending=False)
    acum = g.valor.cumsum() / g.valor.sum()
    g["curva"] = np.where(acum <= 0.80, "A", np.where(acum <= 0.95, "B", "C"))
    g.loc[g.index[0], "curva"] = "A"
    g["cobertura_meses"] = np.where(g.qtd_mes > 0, g.estoque / g.qtd_mes, np.inf)
    g["sem_estoque"] = g.estoque <= 0
    g["alto_giro"] = g.meses_venda >= 6            # vendeu em 6+ dos ~10 meses
    return g, meses


if __name__ == "__main__":
    g, meses = tabela_pn()
    g.to_pickle(f"{S}/out/pn.pkl")
    print("meses", meses, "PNs", len(g))
    print(g.groupby("curva").agg(pns=("valor", "size"), valor=("valor", "sum"),
                                 sem_est=("sem_estoque", "sum"), giro=("alto_giro", "sum"), terc=("terceiro", "sum")))
    a = g[(g.curva == "A") & g.sem_estoque]
    print(a.head(30)[["DESCRPROD", "Grupo Nivel 3", "qtd", "valor", "notas", "meses_venda", "estoque", "terceiro"]].to_string())
