"""Cruza as planilhas por consultor com as bases do Sankhya e do SAB, cliente = CNPJ/CPF completo.

Uso:  python -m campanha.cruzar <pasta_planilhas> <pasta_bases> [saida.xlsx]
<pasta_bases> tem f_Fat_Clientes.csv, d_Clientes.csv e, opcionais, f_notas_SAB.csv e sab_*.csv.
Nada é juntado por razão social nem por raiz de CNPJ: filial e homônimo são clientes distintos.
"""
import sys
from pathlib import Path

import pandas as pd

from . import norm
from .auditar import REVENDA, ler, salvar
from .gerar import CENTRO_RAP, CLASSES_RAP, RAIZ_EXCLUIR

INI, FIM, CORTE = pd.Timestamp("2026-01-01"), pd.Timestamp("2026-06-30"), pd.Timestamp("2026-07-01")


def _csv(arq):
    d = pd.read_csv(arq, dtype=str, keep_default_na=False)
    d.columns = [c.split("[")[-1].strip("]") for c in d.columns]
    return d


def bases(pasta):
    pasta = Path(pasta)
    cli = _csv(pasta / "d_Clientes.csv")
    cli["doc"] = [norm.doc(m, t) or norm.doc(c, t) for m, c, t in zip(cli.CODIGO_MATRIZ, cli.CNPJ_CPF, cli.TIPPESSOA)]
    fat = _csv(pasta / "f_Fat_Clientes.csv").merge(cli[["CODPARC", "doc", "NOMEPARC", "UF", "AD_CLASSRAP"]],
                                                   on="CODPARC", how="left")
    fat["Valor"] = pd.to_numeric(fat.Valor)
    fat["DT"] = pd.to_datetime(fat.DTFATUR).dt.normalize()
    cc = fat.CODCENCUS
    fat["peca"] = cc.str.endswith(("204", "207"))
    fat["balcao"] = cc.str.endswith("204") & cc.ne(str(CENTRO_RAP))
    fat["venda"] = fat.TIPMOV.eq("V")
    rap = set(fat.loc[cc.eq(str(CENTRO_RAP)) & fat.venda, "doc"])
    rap |= set(cli.loc[cli.AD_CLASSRAP.str.upper().isin(CLASSES_RAP), "doc"])
    servico = set()
    if (pasta / "f_notas_SAB.csv").exists():
        sab_nf = _csv(pasta / "f_notas_SAB.csv").merge(cli[["CODPARC", "doc"]], on="CODPARC", how="left")
        servico = set(sab_nf.doc.dropna()) - {""}
    sab = None
    if (pasta / "sab_atividade.csv").exists():
        at, eq, cl = (_csv(pasta / f"sab_{n}.csv") for n in ("atividade", "equipamento", "cliente"))
        k = ["id_ativ", "id_sub"]
        sab = at.merge(eq, on=k, suffixes=("", "_eq")).merge(cl, on=k)
        sab = sab[sab.status.ne("Cancelado")].copy()
        sab["doc"] = sab.cnpj.map(norm.doc)
        sab["necess"] = sab.necess.str.replace(r"\s+", " ", regex=True).str.strip()
        sab["data"] = pd.to_datetime(sab.dt_criacao, errors="coerce")
    return cli, fat, rap, servico, sab


def cruzar(rep, nov, cli, fat, rap, servico, sab):
    ach = []

    def add(camp, r, item, problema, planilha="", base=""):
        ach.append({"Campanha": camp, "Consultor": r["Consultor"], "Cliente": r["Cliente"], "CNPJ/CPF": r["doc"],
                    "UF": r["UF"], "Item": item, "Problema": problema, "Na planilha": planilha, "Na base": base})

    vb = fat[fat.venda & fat.balcao]
    db = fat[fat.TIPMOV.eq("D") & fat.balcao]
    s1 = vb[vb.DT.between(INI, FIM)]
    liq = s1.groupby("doc").Valor.sum().sub(db[db.DT.between(INI, FIM)].groupby("doc").Valor.sum(), fill_value=0)
    ult = s1.groupby("doc").DT.max()
    s2 = fat[fat.venda & fat.peca & (fat.DT >= CORTE)].sort_values("DT")
    s2_ult = s2.groupby("doc").last()
    docs_cad = set(cli.doc)
    # vendedor da última compra Balcão do 1º sem -> consultor da planilha (mapa aprendido pela maioria)
    ult_vend = s1.sort_values("DT").groupby("doc").CODVEND.last()
    rep["vend"] = rep.doc.map(ult_vend)
    # vendedor dividido entre consultores (ex.: conta da casa em rodízio) não define carteira
    dist = rep.dropna(subset=["vend"]).groupby("vend").Consultor.value_counts(normalize=True)
    mapa = dist[dist >= 0.8].reset_index(level=1).Consultor

    # ---------------- Repetição
    for r in rep.to_dict("records"):
        d = r["doc"]
        if d not in docs_cad:
            add("Repetição", r, "R1", "CNPJ/CPF não existe no cadastro de clientes do Sankhya")
            continue
        if d not in liq.index:
            add("Repetição", r, "R2", "Sem NF de peça Balcão no 1º sem/2026 neste CNPJ/CPF (a compra foi de outro CNPJ do mesmo nome/grupo)",
                r["Valor 1º sem/2026 (R$)"], "R$ 0,00")
            continue
        if liq[d] <= 0:
            add("Repetição", r, "R3", "Compra do 1º sem devolvida (líquido ≤ 0)", r["Valor 1º sem/2026 (R$)"], f"{liq[d]:.2f}")
        elif abs(liq[d] - r["valor"]) > 1:
            add("Repetição", r, "R4", "Valor do 1º sem diverge (planilha somou outros CNPJs do grupo ou não descontou devolução)",
                r["Valor 1º sem/2026 (R$)"], f"{liq[d]:.2f}")
        if ult[d].date() != r["ult"].date():
            add("Repetição", r, "R5", "Data da última compra diverge", r["Data da última compra"], f"{ult[d]:%Y-%m-%d}")
        if d in s2_ult.index:
            u = s2_ult.loc[d]
            add("Repetição", r, "R6", "Já comprou peça no 2º sem/2026: sai da campanha (já repetiu)",
                "", f"NF {u.NUNOTA} em {u.DT:%d/%m/%Y}, R$ {u.Valor:,.2f}")
        if d in rap:
            add("Repetição", r, "R7", "Cliente RAP (classe RAP no cadastro ou compra no centro RAP)")
        v = r["vend"]
        if isinstance(v, str) and v in mapa.index and mapa[v] != r["Consultor"]:
            add("Repetição", r, "R8", "Consultor diferente do vendedor da última compra", r["Consultor"],
                f"CODVEND {v} (carteira de {mapa[v]})")

    na_planilha = set(rep.doc)
    falt = []
    nome = cli.drop_duplicates("doc").set_index("doc")
    for d in sorted(set(liq[liq > 0].index) - set(s2_ult.index) - na_planilha - rap - {"", None}):
        if d[:8] in RAIZ_EXCLUIR or not isinstance(d, str):
            continue
        n = nome.loc[d] if d in nome.index else None
        v = ult_vend.get(d)
        falt.append({"Cliente": n.NOMEPARC if n is not None else "", "CNPJ/CPF": d, "UF": n.UF if n is not None else "",
                     "Valor 1º sem/2026 (R$)": round(liq[d], 2), "Última compra 1º sem": f"{ult[d]:%Y-%m-%d}",
                     "CODVEND última compra": v, "Consultor provável": mapa.get(v, ""),
                     "Possível revenda": bool(n is not None and REVENDA.search(norm.razao(n.NOMEPARC)))})
    falt = pd.DataFrame(falt)
    if len(falt):
        falt = falt.sort_values("Valor 1º sem/2026 (R$)", ascending=False)

    # ---------------- Novos
    peca_docs = fat[fat.venda & fat.peca].sort_values("DT").groupby("doc").last()
    et = sab[sab.necess.eq("Entrega Técnica")] if sab is not None else None
    for r in nov.drop_duplicates(["doc", "Chassi"]).to_dict("records"):
        d = r["doc"]
        if d in peca_docs.index:
            u = peca_docs.loc[d]
            add("Novos", r, "N1", "Já comprou peça em 2026 (não é cliente novo)", "",
                f"NF {u.NUNOTA} em {u.DT:%d/%m/%Y}, R$ {u.Valor:,.2f}")
        if d in rap:
            add("Novos", r, "N2", "Cliente RAP")
        if et is not None:
            e = et[et.doc == d]
            if not len(e):
                add("Novos", r, "N3", "Sem entrega técnica no SAB para este CNPJ/CPF")
            elif r["Chassi"].strip().upper() not in set(e.serie.str.strip().str.upper()):
                add("Novos", r, "N4", "Chassi não bate com as entregas técnicas do SAB deste CNPJ/CPF",
                    r["Chassi"], ", ".join(sorted(set(e.serie)))[:200])

    # quem deveria estar em Novos e não está: comprou máquina (ET) ou serviço e nunca comprou peça
    sem_peca = lambda d: d and d not in peca_docs.index and d not in rap and d[:8] not in RAIZ_EXCLUIR
    cand = {}
    if sab is not None:
        for d, g in sab[sab.doc.ne("")].groupby("doc"):
            if sem_peca(d):
                cand[d] = {"Origem": ", ".join(sorted(set(g.necess.dropna())))[:120], "Cliente": g.nome.iloc[-1],
                           "UF": g.uf.iloc[-1]}
    for d in servico:
        if sem_peca(d) and d not in cand:
            n = nome.loc[d] if d in nome.index else None
            cand[d] = {"Origem": "NF de serviço (f_notas_SAB)", "Cliente": n.NOMEPARC if n is not None else "",
                       "UF": n.UF if n is not None else ""}
    falt_n = pd.DataFrame([{"CNPJ/CPF": d, **v, "Tem ET": "Entrega" in v["Origem"]}
                           for d, v in cand.items() if d not in set(nov.doc) and d not in na_planilha])

    ach = pd.DataFrame(ach).sort_values(["Item", "Consultor", "Cliente"]) if ach else pd.DataFrame()
    por_item = ach.groupby(["Campanha", "Item", "Problema"]).agg(Ocorrencias=("Cliente", "size"),
                                                                 CNPJs=("CNPJ/CPF", "nunique")).reset_index()
    return ach, por_item, falt, falt_n, mapa


if __name__ == "__main__":
    rep, nov = ler(sys.argv[1])
    ach, por_item, falt, falt_n, mapa = cruzar(rep, nov, *bases(sys.argv[2]))
    destino = sys.argv[3] if len(sys.argv) > 3 else "cruzamento_campanha_gpd.xlsx"
    salvar(destino, Por_item=por_item, Achados=ach, Faltantes_Repeticao=falt, Faltantes_Novos=falt_n,
           Vendedor_Consultor=mapa.rename("Consultor").reset_index())
    print(por_item.to_string(index=False))
    print("faltantes repetição:", len(falt), " faltantes novos:", len(falt_n))
