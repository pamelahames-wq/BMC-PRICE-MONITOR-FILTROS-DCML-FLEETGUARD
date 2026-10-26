"""Reestruturação RAP: classificação por régua semestral, estudo de preço e margem por categoria.

Uso:  python -m rap.analise   (imprime os números; rap.saidas gera Excel e PowerPoint)
Lê as bases do Sankhya copiadas em cache/raw (d_Clientes, f_Fat_Clientes, f_Notas, d_Produtos,
d_Precos, f_CustoUnitarioProd). Tabela de grupos/descontos em rap/grupos.tsv.
"""
from pathlib import Path

import pandas as pd
from unidecode import unidecode

from campanha import norm

RAIZ = Path(__file__).resolve().parent.parent
RAW = RAIZ / "cache" / "raw"

INICIO, FIM = "2026-07-01", "2026-10-16"          # 2º sem/2026 até a última NF da base
FIM_SEMESTRE = "2026-12-31"
INICIO_1S, FIM_1S = "2026-01-01", "2026-06-30"
VEND_RAP = {"ELISASILVA": "ESILVA (ELISASILVA)", "SLUNA": "SLUNA",
            "PLOUREIRO": "PLOUREIRO", "FALEXANDRE": "FALEXANDRE"}
CATEGORIAS = ["Diamante", "Ouro", "Standard"]
MINIMO = {"Standard": 5_000, "Ouro": 40_000, "Diamante": 100_000}
FAIXA = {"Standard": (0.20, 0.25), "Ouro": (0.30, 0.40), "Diamante": (0.45, 0.55)}
REBAIXA = {"Diamante": "Ouro", "Ouro": "Standard", "Standard": "Encerrar"}
DESC_BALCAO = 0.15
IMPOSTOS = 0.22        # média IMPOSTOS/BRUTO das NFs RAP 2026 (f_Notas) — usado só no estudo de preço
BALDES = ["DESHOMOLOGAR", "REBAIXAR", "INCENTIVAR", "MANTER"]


def _ler(nome):
    d = pd.read_csv(RAW / nome, dtype=str, low_memory=False)
    d.columns = [c.split("[")[-1].strip("]") for c in d.columns]
    return d


def _chave(s):
    return unidecode(str(s or "")).upper().strip()


def grupos():
    g = pd.read_csv(Path(__file__).with_name("grupos.tsv"), sep="\t", dtype=str).fillna("")
    g["k1"], g["k3"] = g.SUBRAIZ.map(_chave), g.DESCRGRUPOPROD.map(_chave)
    return g


def carregar():
    cli = _ler("d_Clientes.csv")
    cli["vend"] = cli.VEND_PREFERENCIAL.str.upper().str.strip()
    cli["CNPJ"] = [norm.doc(m, t) for m, t in zip(cli.CODIGO_MATRIZ, cli.TIPPESSOA)]
    cli["categoria"] = cli.AD_CLASSRAP.fillna("").str.strip()

    fat = _ler("f_Fat_Clientes.csv")
    fat["Valor"] = pd.to_numeric(fat.Valor).fillna(0)
    fat["DT"] = pd.to_datetime(fat.DTFATUR)
    # só centros de peças (…204 balcão, …207 RAP); …201 é máquina
    fat = fat[fat.CODCENCUS.str.endswith(("204", "207"))]
    fat["sinal"] = fat.TIPMOV.map({"V": 1, "D": -1}).fillna(0)

    it = _ler("f_Notas.csv")
    for c in ["QTDNEG", "BRUTO", "IMPOSTOS", "CUSTO"]:
        it[c] = pd.to_numeric(it[c]).fillna(0)
    it["DT"] = pd.to_datetime(it.DTFATUR)
    it = it[it.CODCENCUS.str.endswith(("204", "207"))].merge(fat[["NUNOTA", "CODPARC", "sinal"]], on="NUNOTA")

    pro = _ler("d_Produtos.csv")
    g = grupos()
    cod = {(a, b): c for a, b, c in zip(g.k1, g.k3, g.CODGRUPOPROD)}
    cod3 = dict(zip(g[g.k1 == "PECAS HYUNDAI"].k3, g[g.k1 == "PECAS HYUNDAI"].CODGRUPOPROD))
    pro["CODGRUPOPROD"] = [cod.get((_chave(a), _chave(b))) or cod3.get(_chave(b), "")
                           for a, b in zip(pro["Grupo Nivel 1"], pro["Grupo Nivel 3"])]
    pro["grupo"] = pro["Grupo Nivel 3"]

    pre = _ler("d_Precos.csv")
    pre["preco"] = pd.to_numeric(pre.VLRVENDA)
    pre = pre[(pre.preco > 0) & (pre.preco < 9e7)]

    cu = _ler("f_CustoUnitarioProd.csv").rename(columns={"codProduto": "CODPROD"})
    cu["DT"] = pd.to_datetime(cu.dtCusto)
    cu["custo"] = pd.to_numeric(cu.vlrCustoUnitario)
    cu["atual"] = pd.to_numeric(cu.vlrAtual)
    return cli, fat, it, pro, pre, cu


def custo_itens(it, cu):
    """Custo unitário da NF; se vier zerado, o último custo histórico até a data da NF."""
    hist = cu[cu.custo > 0].sort_values("DT")[["CODPROD", "DT", "custo"]]
    it = it.sort_values("DT")
    a = pd.merge_asof(it, hist, on="DT", by="CODPROD", direction="backward")
    b = pd.merge_asof(it, hist, on="DT", by="CODPROD", direction="forward")
    a["custo_hist"] = a.custo.fillna(b.custo)
    nf = a.CUSTO > 0
    a["custo_unit"] = a.CUSTO.where(nf, a.custo_hist)
    a["origem_custo"] = "NF"
    a.loc[~nf & a.custo_hist.notna(), "origem_custo"] = "histórico"
    a.loc[a.custo_unit.isna() | (a.custo_unit <= 0), "origem_custo"] = "s/ custo"
    return a.drop(columns=["custo", "custo_hist"])


def populacao(cli):
    rap_vend = cli.vend.isin(VEND_RAP)
    cat_ok = cli.categoria.isin(CATEGORIAS)
    rap = cli[rap_vend & cat_ok].copy()
    inconsistente = cli[~rap_vend & cat_ok].copy()
    inconsistente["Problema"] = "Categoria RAP preenchida, mas vendedor preferencial não é RAP"
    sem_cat = cli[rap_vend & ~cat_ok].copy()
    sem_cat["Problema"] = "Vendedor RAP, mas categoria fora da régua (" + \
        sem_cat.categoria.replace("", "vazia") + ")"
    return rap, pd.concat([inconsistente, sem_cat])


def faturamento(fat, ini, fim):
    f = fat[(fat.DT >= ini) & (fat.DT <= fim)]
    return (f.Valor * f.sinal).groupby(f.CODPARC).sum()


def classificar(rap, fat):
    fs = faturamento(fat, INICIO, FIM)
    f1 = faturamento(fat, INICIO_1S, FIM_1S)
    dias = (pd.Timestamp(FIM) - pd.Timestamp(INICIO)).days + 1
    dias_sem = (pd.Timestamp(FIM_SEMESTRE) - pd.Timestamp(INICIO)).days + 1
    d = rap.copy()
    d["fat"] = d.CODPARC.map(fs).fillna(0).round(2)
    d["fat_1s"] = d.CODPARC.map(f1).fillna(0).round(2)
    d["minimo"] = d.categoria.map(MINIMO)
    d["pct"] = d.fat / d.minimo
    d["balde"] = "MANTER"
    d.loc[d.pct < 1, "balde"] = "INCENTIVAR"
    d.loc[d.pct < 0.7, "balde"] = "REBAIXAR"
    d.loc[d.fat <= 0, "balde"] = "DESHOMOLOGAR"
    std_baixo = (d.balde == "REBAIXAR") & (d.categoria == "Standard")
    d.loc[std_baixo, "balde"] = "DESHOMOLOGAR"
    d["motivo"] = ""
    d.loc[d.fat <= 0, "motivo"] = "Sem faturamento no semestre"
    d.loc[std_baixo, "motivo"] = "Standard abaixo de 70% do mínimo"
    d["sugerida"] = d.categoria
    d.loc[d.balde == "REBAIXAR", "sugerida"] = d.categoria.map(REBAIXA)
    d.loc[d.balde == "DESHOMOLOGAR", "sugerida"] = "Encerrar"
    d["falta"] = (d.minimo - d.fat).where(d.balde == "INCENTIVAR", 0).round(2)
    # referência: ritmo atual projetado para o semestre inteiro
    d["projecao"] = (d.fat * dias_sem / dias).round(2)
    d["pct_proj"] = d.projecao / d.minimo
    d["consultor"] = d.vend.map(VEND_RAP)
    ordem = {b: i for i, b in enumerate(BALDES)}
    return d.sort_values(["balde", "fat"], key=lambda s: s.map(ordem) if s.name == "balde" else -s)


def itens_rap(it, cu, pro, rap):
    i = it[(it.DT >= INICIO) & (it.DT <= FIM) & it.CODPARC.isin(rap.CODPARC)]
    i = custo_itens(i, cu)
    i = i.merge(rap[["CODPARC", "categoria", "NOMEPARC"]], on="CODPARC") \
         .merge(pro[["CODPROD", "REFERENCIA", "DESCRPROD", "CODGRUPOPROD", "grupo"]], on="CODPROD", how="left")
    i["CODGRUPOPROD"] = i.CODGRUPOPROD.fillna("").replace("", "(sem código)")
    i["grupo"] = i.grupo.fillna("(sem grupo)")
    i["fat_nf"] = i.BRUTO * i.sinal
    i["impostos"] = i.IMPOSTOS * i.sinal
    i["receita_liq"] = i.fat_nf - i.impostos
    i["custo_total"] = (i.custo_unit.fillna(0) * i.QTDNEG * i.sinal).where(i.origem_custo != "s/ custo", 0)
    return i


def margem(itens, chaves):
    com = itens[itens.origem_custo != "s/ custo"]
    a = itens.groupby(chaves).agg(fat_nf=("fat_nf", "sum"), impostos=("impostos", "sum"),
                                  receita_liq=("receita_liq", "sum"))
    b = com.groupby(chaves).agg(receita_c_custo=("receita_liq", "sum"), custo=("custo_total", "sum"))
    m = a.join(b).fillna(0)
    m["margem_rs"] = m.receita_c_custo - m.custo
    m["margem_pct"] = m.margem_rs / m.receita_c_custo.where(m.receita_c_custo != 0)
    m["cobertura_custo"] = m.receita_c_custo / m.receita_liq.where(m.receita_liq != 0)
    return m.reset_index()


def desconto_efetivo(itens, pre):
    """Desconto médio praticado nas NFs de venda vs preço de tabela atual (d_Precos)."""
    v = itens[itens.sinal > 0].merge(pre[["CODPROD", "preco"]], on="CODPROD")
    v["tabela"] = v.preco * v.QTDNEG
    s = v.groupby("categoria")[["BRUTO", "tabela"]].sum()
    return (1 - s.BRUTO / s.tabela).rename("desc_efetivo")


CESTA = [  # (CODGRUPOPROD, quantos itens) — mais faturados em 2026 em cada grupo
    ("303000000", 2), ("350000000", 2), ("325000000", 1), ("360000000", 2), ("356000000", 1),
    ("321000000", 1), ("316000000", 1), ("318000000", 1), ("311000000", 1), ("310000000", 1),
    ("368000000", 1),
]


def cesta(it, pro, pre, cu):
    venda = it[it.sinal > 0].groupby("CODPROD").agg(fat_2026=("BRUTO", "sum"), qtd_2026=("QTDNEG", "sum"))
    atual = cu.sort_values("DT").groupby("CODPROD").custo.last()
    p = pro.merge(pre[["CODPROD", "preco"]], on="CODPROD").join(venda, on="CODPROD")
    p["custo"] = p.CODPROD.map(atual)
    p = p[p.fat_2026.notna()]
    out = []
    for grp, n in CESTA:
        g = p[p.CODGRUPOPROD == grp].sort_values("fat_2026", ascending=False)
        out.append(g[g.custo.notna()].head(n))
    return pd.concat(out)[["REFERENCIA", "CODPROD", "DESCRPROD", "CODGRUPOPROD", "grupo",
                           "custo", "preco", "fat_2026", "qtd_2026"]]


def rodar():
    cli, fat, it, pro, pre, cu = carregar()
    rap, inconsistente = populacao(cli)
    cls = classificar(rap, fat)
    fs = faturamento(fat, INICIO, FIM)
    inconsistente["fat"] = inconsistente.CODPARC.map(fs).fillna(0).round(2)
    itens = itens_rap(it, cu, pro, rap)
    return {
        "classificacao": cls,
        "inconsistente": inconsistente,
        "itens": itens,
        "margem_cat": margem(itens, ["categoria"]),
        "margem_cat_grupo": margem(itens, ["categoria", "CODGRUPOPROD", "grupo"]),
        "desc_efetivo": desconto_efetivo(itens, pre),
        "cesta": cesta(it, pro, pre, cu),
        "data_max": fat.DT.max(),
    }


if __name__ == "__main__":
    r = rodar()
    c = r["classificacao"]
    pd.set_option("display.width", 200)
    print(c.groupby("balde").agg(n=("CODPARC", "size"), fat=("fat", "sum")))
    print(pd.crosstab(c.categoria, c.balde, margins=True))
    print(pd.crosstab(c.consultor, c.balde, margins=True))
    print(r["margem_cat"].round(3).to_string())
    print(r["desc_efetivo"])
    print(r["itens"].origem_custo.value_counts())
    print(r["cesta"].to_string())
    print(r["inconsistente"].Problema.value_counts())
