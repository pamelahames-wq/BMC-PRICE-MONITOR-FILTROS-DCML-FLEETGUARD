"""Liga PN -> família/modelo de máquina (clientes SAB + planos oficiais + descrição), forecast 90 dias e prioridades."""
import json, re, sys
import numpy as np
import pandas as pd
sys.path.insert(0, __file__.rsplit("/", 1)[0])
from base import S, R, base_pn, carregar, tabela_pn

FAM = ["Escavadeira", "Pá carregadeira", "Empilhadeira", "Retroescavadeira"]


def familia(modelo):
    m = re.sub(r"[\s\-]", "", str(modelo).upper())
    m = m.replace("ESCAVADEIRA", "")
    if m.startswith("HB"):
        return "Retroescavadeira"
    if m.startswith("HL"):
        return "Pá carregadeira"
    if re.match(r"^(R\d|HX\d|7HX|ROBEX)", m):
        return "Escavadeira"
    if re.match(r"^\d{2,3}(D|L|LE|DE|B|BR|G)", m) or "EMPILH" in m:
        return "Empilhadeira"
    if m.startswith("HG"):
        return "Motoniveladora"
    return None


def norm_modelo(modelo):
    m = re.sub(r"\s+", "", str(modelo).upper()).replace("ESCAVADEIRA-", "")
    m = re.sub(r"-(LC|L|AL)?$", lambda x: x.group(0), m)
    m = m.replace("HL757-9PRIME", "HL757PRIME").replace("HL760-9PRIME", "HL760PRIME")
    m = re.sub(r"^(HL\d{3}|HX\d{2,3}|HB\d{3}|R\d{3}|HG\d{3}).*$", r"\1", m) if not m.startswith("HX") else re.sub(r"^(HX\d{2,3}).*$", r"\1", m)
    return m


def rd(f):
    d = pd.read_csv(f"{R}/{f}", dtype=str)
    d.columns = [c.split("[")[-1].strip("]") for c in d.columns]
    return d


def frota_por_cliente():
    sc, sa, se = rd("sab_cliente.csv"), rd("sab_atividade.csv"), rd("sab_equipamento.csv")
    se = se[se.marca == "HYUNDAI"].copy()
    se["fam"] = se.modelo.map(familia)
    se["mod"] = se.modelo.map(norm_modelo)
    sc["raiz"] = sc.cnpj.fillna("").str.zfill(14).str[:8]
    x = se.merge(sc[["id_ativ", "raiz"]], on="id_ativ").merge(sa[["id_ativ", "dt_criacao", "status"]], on="id_ativ")
    x = x[(x.raiz != "00000000") & x.fam.notna()]
    return x


def parque(x):
    """Máquinas únicas (série) por família/modelo e as que entraram no SAB em 2026 (proxy de venda nova)."""
    s = x[x.serie.notna() & (x.serie.str.len() > 5)].copy()
    s["dt"] = pd.to_datetime(s.dt_criacao, errors="coerce")
    prim = s.sort_values("dt").drop_duplicates("serie")
    gar = s[s.necess.fillna("").str.contains("Garantia", case=False)].serie.unique()
    prim["nova"] = prim.serie.isin(gar)
    prim["ano"] = prim.dt.dt.year
    fam = prim.groupby("fam").agg(parque=("serie", "nunique"))
    fam["novas_2025"] = prim[prim.nova & (prim.ano == 2025)].groupby("fam").serie.nunique()
    fam["novas_2026"] = prim[prim.nova & (prim.ano == 2026)].groupby("fam").serie.nunique()
    mod = prim.groupby(["fam", "mod"]).agg(parque=("serie", "nunique"))
    mod["novas_2026"] = prim[prim.nova & (prim.ano == 2026)].groupby(["fam", "mod"]).serie.nunique()
    return fam.fillna(0), mod.fillna(0).reset_index()


def pn_familia(n, p, x):
    fc = rd("f_Fat_Clientes.csv")
    cl = rd("d_Clientes.csv")[["CODPARC", "CNPJ_CPF", "TIPPESSOA"]]
    cl = cl[cl.CNPJ_CPF.str.len() == 8]
    v = n[n.TIPMOV == "V"].merge(fc[["NUNOTA", "CODPARC"]], on="NUNOTA").merge(cl, on="CODPARC")
    v = v.merge(p[["CODPROD", "PN"]], on="CODPROD")
    frota = x.drop_duplicates(["raiz", "serie"]).groupby(["raiz", "fam"]).size().rename("maq").reset_index()
    frota["peso"] = frota.maq / frota.groupby("raiz").maq.transform("sum")
    mods = x.drop_duplicates(["raiz", "serie"]).groupby(["raiz", "mod"]).size().rename("maq").reset_index()
    mods["peso"] = mods.maq / mods.groupby("raiz").maq.transform("sum")
    j = v[["PN", "NUNOTA", "CNPJ_CPF"]].drop_duplicates().merge(frota, left_on="CNPJ_CPF", right_on="raiz")
    w = j.groupby(["PN", "fam"]).peso.sum().unstack(fill_value=0)
    share = w.div(w.sum(axis=1), axis=0)
    top = share.idxmax(axis=1)
    out = pd.DataFrame({"fam_cli": top, "fam_share": share.max(axis=1), "notas_lig": j.groupby("PN").NUNOTA.nunique()})
    jm = v[["PN", "NUNOTA", "CNPJ_CPF"]].drop_duplicates().merge(mods, left_on="CNPJ_CPF", right_on="raiz")
    wm = jm.groupby(["PN", "mod"]).peso.sum()
    tm = wm.reset_index().sort_values("peso", ascending=False).groupby("PN").head(3)
    out["modelos_cli"] = tm.groupby("PN").mod.apply(lambda s: ", ".join(s))
    fm = x.drop_duplicates("mod").set_index("mod").fam
    wm = wm.reset_index()
    wm["fam"] = wm["mod"].map(fm)
    return out, wm


def pn_planos():
    pl = json.load(open(f"{R}/planos_oficiais.json"))["planos"]
    rows = []
    for m, d in pl.items():
        for it in d["itens"]:
            rows.append((base_pn(it["pn"]), m, familia(m)))
    k = pd.DataFrame(rows, columns=["PN", "mod", "fam"])
    return k.groupby("PN").agg(fam_plano=("fam", lambda s: s.mode()[0]), modelos_plano=("mod", lambda s: ", ".join(sorted(set(s)))))


def fam_descr(d):
    d = str(d).upper()
    for k, f in [("EMPILH", "Empilhadeira"), (" EMP", "Empilhadeira"), ("CARREGAD", "Pá carregadeira"),
                 ("CARREG", "Pá carregadeira"), ("ESCAV", "Escavadeira"), ("RETRO", "Retroescavadeira")]:
        if k in d:
            return f
    return None


def categoria(g3, desc):
    g3, d = str(g3), str(desc).upper()
    tem = lambda *ks: any(k in d for k in ks)
    if "FILTR" in g3 or tem("FILTRO", "ELEMENTO FILTRANTE", "ELEMENT-", "ELEMENTO-AR", "ELEMENTO FILTRO"):
        return "Filtros"
    if tem("VIDRO", "PARA BRISA", "PARA-BRISA", "PARABRISA", "JANELA", "ESPELHO", "LIMPADOR", "PALHETA",
           "MACANETA", "CONSOLE", "CABINA", "CABINE") or g3 == "CABINES":
        return "Cabine e vidros"
    if "KIT VEDACAO" in g3 or tem("VEDACAO", "VEDADOR", "RETENTOR", "O-RING", "O RING", "ANEL BORRACHA", "JUNTA", "ANEL KIT"):
        return "Vedação"
    if "MANGUEIRA" in g3 or tem("MANGUEIRA"):
        return "Mangueiras"
    if tem("RADIADOR", "CONDENSADOR", "AR CONDICIONADO", "RESERVATORIO DE AGUA", "HELICE"):
        return "Arrefecimento e A/C"
    if any(k in g3 for k in ["BOMBA", "MOT. TRACAO", "MOTORES DE TRACAO", "GIRO", "MCV", "VALVULA", "HIDRAUL", "CILINDRO",
                             "HASTE", "IMPLEMENTOS"]) or tem("ACUMULADOR", "BOMBA", "VALVULA", "SOLENOIDE", "ENGATE", "JOYSTICK"):
        return "Hidráulico (bombas, motores, válvulas)"
    if any(k in g3 for k in ["EIXOS", "DISCOS", "TRANSMISS"]) or tem("TRANSMISS", "ENGRENAGEM", "ACOPLAMENTO",
                                                                    "RODA DENTADA", "FREIO", "CUBO", "MARCHAS"):
        return "Transmissão e eixos"
    if any(k in g3 for k in ["RODA GUIA", "ROLETE", "MATERIAL RODANTE", "LAGARTA", "SAPATA", "CORRENTE"]) or tem("LAGARTA", "ROLETE"):
        return "Material rodante"
    if any(k in g3 for k in ["MOTOR", "CUMMINS", "PERKINS", "KUBOTA", "MITSUBISHI", "YANMAR", "H.M.C", "RESFRIAMENTO",
                             "ESCAPAMENTO"]) or tem("INJETOR", "ATUADOR BOMBA COMBUSTIVEL", "CORREIA"):
        return "Motor"
    if any(k in g3 for k in ["ELETRIC", "SENSOR", "CHICOTE", "LAMPADA"]) or tem("INTERRUPTOR", "FAROL", "CHAVE", "CONTROLE ELETR"):
        return "Elétrica e sensores"
    if tem("COXIM", "SUPORTE", "AMORTECEDOR"):
        return "Coxins e suportes"
    if tem("PINO", "BUCHA", "ROLAMENTO", "MANCAL", "CALCO", "TERMINAL") or "BUCHAS" in g3 or g3 == "MANCAL":
        return "Pinos, buchas e rolamentos"
    if g3 == "FPS" or tem("DENTE") or "CACAMBA" in g3:
        return "FPS / caçamba"
    if "PNEU" in g3:
        return "Pneus"
    return "Outras peças"


def montar():
    g, meses = tabela_pn()
    p, n, e = carregar()
    x = frota_por_cliente()
    fam, mod = parque(x)
    pf, wm = pn_familia(n, p, x)
    g = g.join(pf).join(pn_planos())
    g["fam_desc"] = g.DESCRPROD.map(fam_descr)
    g["familia"] = g.fam_plano.fillna(g.fam_desc)
    ok = g.familia.isna() & (g.fam_share >= 0.6) & (g.notas_lig >= 2)
    g.loc[ok, "familia"] = g.loc[ok, "fam_cli"]
    g.loc[g.index.str.startswith("SIBU-"), "familia"] = "Retroescavadeira"
    g["familia"] = g.familia.fillna("Várias / não identificada")
    # modelos dos clientes, só da família atribuída
    wm = wm.merge(g[["familia"]], left_on="PN", right_index=True)
    wm = wm[wm.fam == wm.familia].sort_values("peso", ascending=False).groupby("PN").head(3)
    g["modelos"] = g.modelos_plano.fillna(wm.groupby("PN")["mod"].apply(lambda s: ", ".join(s))).fillna("")
    g["categoria"] = [categoria(a, b) for a, b in zip(g["Grupo Nivel 3"], g.DESCRPROD)]

    # forecast 90 dias = média mensal 2026 x 3 x fator de crescimento do parque da família
    cres = (fam.novas_2026 / (fam.parque - fam.novas_2026).clip(lower=1)).clip(upper=0.5)
    fator = (1 + cres * 0.25).to_dict()            # 3 meses de um crescimento anualizado
    g["fator_parque"] = g.familia.map(fator).fillna(1.0)
    g["forecast_90d"] = np.ceil(np.maximum(g.qtd_mes, g.qtd_mes_rec).clip(lower=0) * 3 * g.fator_parque)
    g["necessidade"] = (g.forecast_90d - g.estoque).clip(lower=0)

    # pacote RAP enviado
    pk = pd.read_excel(f"{S}/upload/pacote.xlsx", sheet_name="Lista de peças", header=1)
    pk["PN"] = pk["Part number"].fillna("").astype(str).map(lambda s: base_pn(str(s)))
    pk = pk.groupby("PN")["Qtd a importar"].sum()
    g["pedido_pacote"] = pk.reindex(g.index).fillna(0)

    g["recorrente"] = g.meses_venda >= 4
    g["status"] = np.select(
        [g.sem_estoque, g.cobertura_meses < 1, g.cobertura_meses < 3],
        ["Sem estoque", "Crítico (<1 mês)", "Atenção (1–3 meses)"], "OK")
    g["prioridade"] = np.select(
        [(g.curva == "A") & g.recorrente & (g.cobertura_meses < 1),
         (g.curva == "B") & g.recorrente & (g.cobertura_meses < 1),
         (g.curva == "A") & g.recorrente & (g.cobertura_meses < 3)],
        ["P1", "P2", "P3"], "")
    g.to_pickle(f"{S}/out/pn_full.pkl")
    fam.to_pickle(f"{S}/out/parque_fam.pkl")
    mod.to_pickle(f"{S}/out/parque_mod.pkl")
    return g, fam, mod


if __name__ == "__main__":
    g, fam, mod = montar()
    print(fam)
    print(mod.sort_values("novas_2026", ascending=False).head(15))
    print(g.familia.value_counts())
    print(g.prioridade.value_counts())
    print(pd.crosstab(g[g.prioridade != ""].categoria, g[g.prioridade != ""].familia))
    print(g[g.prioridade == "P1"][["DESCRPROD", "familia", "modelos", "categoria", "qtd_mes", "meses_venda", "estoque", "forecast_90d", "pedido_pacote"]].to_string())
