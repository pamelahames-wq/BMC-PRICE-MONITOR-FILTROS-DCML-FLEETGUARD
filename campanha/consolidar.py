"""Gera as listas finais da campanha GPD (Repetição e Novos) já corrigidas pelo cruzamento com as bases.

Uso:  python -m campanha.consolidar <pasta_planilhas> <pasta_bases> <saida.xlsx>
Regras: cliente = CNPJ/CPF completo; sai quem já comprou peça no 2º sem, RAP, BMC/Hyundai e
compra do 1º sem devolvida; entram CNPJs que cumprem a regra e faltavam (filiais que a geração
anterior juntava por nome/raiz). Telefone de consultor/BMC ou repetido em 3+ clientes não vai ao Léo.
"""
import itertools
import re
import sys

import pandas as pd

from . import norm
from .auditar import REVENDA, consultor_esperado, ler, salvar
from .cruzar import bases, cruzar

REGIONAL = {"ALEF": "Nordeste", "RANNIERICKSON": "Norte", "RICARDO": "Norte", "THIAGO": "Norte",
            "YURI": "Bahia - Espírito Santo", "JHONY": "Bahia - Espírito Santo",
            "MATEUS": "Minas Gerais", "DANIEL": "Minas Gerais", "RAYLANDER": "Minas Gerais",
            "EDU": "São Paulo", "FABIO": "São Paulo", "KAREN": "São Paulo",
            "RENAN": "Sul (PR/SC/RS)", "GEISON": "Centro Oeste", "PABBLO": "Centro Oeste"}
CELULAR = re.compile(r"\+55[1-9][1-9]9\d{8}")


class Rodizio:
    def __init__(self):
        self.ciclos = {}

    def __call__(self, uf, cidade, tel):
        op = sorted(consultor_esperado(uf, cidade, tel) - {"RAYLANDER"})
        if not op:
            return "SEM CONSULTOR"
        return next(self.ciclos.setdefault(tuple(op), itertools.cycle(op)))


def telefones_internos(rep, nov, cli, sab):
    pares = [(t, d) for t, d in zip(pd.concat([rep.tel, nov.tel]), pd.concat([rep.doc, nov.doc])) if t]
    pares += [(norm.e164(f), d) for d, txt in zip(sab.doc, sab.fone) for f in norm.telefones(txt)]
    n = pd.DataFrame(pares, columns=["t", "d"]).drop_duplicates().groupby("t").d.nunique()
    internos = set(n[n >= 3].index)
    for e, t in zip(pd.concat([rep.email, nov.email]), pd.concat([rep.tel, nov.tel])):
        if "bmchyundai" in e and t:
            internos.add(t)
    return internos


def contato(doc, cli_doc, sab_doc, internos):
    """Primeiro celular válido e não interno: cadastro Sankhya, depois SAB."""
    fontes = list(cli_doc.get(doc, [])) + list(sab_doc.get(doc, []))
    fones = [norm.e164(f) for txt in fontes for f in norm.telefones(txt)]
    bons = [f for f in fones if f not in internos]
    cel = [f for f in bons if CELULAR.fullmatch(f)]
    return cel[0] if cel else (bons[0] if bons else "")


def status_tel(t, internos):
    if not t:
        return "sem telefone"
    if t in internos:
        return "telefone interno BMC/repetido"
    return "ok" if CELULAR.fullmatch(t) else "não é celular"


def consolidar(pasta, pasta_bases):
    rep, nov = ler(pasta)
    cli, fat, rap, servico, sab = bases(pasta_bases)
    ach, por_item, falt, falt_n, mapa = cruzar(rep, nov, cli, fat, rap, servico, sab)
    internos = telefones_internos(rep, nov, cli, sab)
    cli_doc = cli.groupby("doc").CONTATO.apply(list).to_dict()
    sab_doc = sab.groupby("doc").fone.apply(list).to_dict()
    email_doc = cli[cli.EMAIL.str.contains("@")].groupby("doc").EMAIL.first().to_dict()
    cidade_doc = cli.groupby("doc").CIDADE.first().to_dict()
    rodizio = Rodizio()

    def motivo(camp, itens):
        a = ach[(ach.Campanha == camp) & ach.Item.isin(itens)]
        return a.groupby("CNPJ/CPF").Problema.first().to_dict()

    # ---------------- Repetição
    sai_rep = motivo("Repetição", ["R1", "R2", "R3", "R6", "R7"])
    vb = fat[fat.venda & fat.balcao & fat.DT.between("2026-01-01", "2026-06-30")]
    db = fat[fat.TIPMOV.eq("D") & fat.balcao & fat.DT.between("2026-01-01", "2026-06-30")]
    liq = vb.groupby("doc").Valor.sum().sub(db.groupby("doc").Valor.sum(), fill_value=0)
    ult = vb.groupby("doc").DT.max()
    linhas, removidos = [], []
    for r in rep.to_dict("records"):
        d = r["doc"]
        if d in sai_rep:
            removidos.append({"Campanha": "Repetição", "Consultor": r["Consultor"], "Cliente": r["Cliente"],
                              "CNPJ/CPF": d, "UF": r["UF"], "Motivo": sai_rep[d]})
            continue
        tel = r["tel"] if r["tel"] and r["tel"] not in internos else contato(d, cli_doc, sab_doc, internos) or r["tel"]
        linhas.append({"Consultor": r["Consultor"], "Cliente": r["Cliente"], "CNPJ/CPF": d, "Cidade": r["Cidade"],
                       "UF": r["UF"], "E-mail": r["email"], "Telefone": tel,
                       "Data da última compra": f"{ult[d]:%Y-%m-%d}", "Valor 1º sem/2026 (R$)": round(liq[d], 2),
                       "O que comprava (categorias)": r["O que comprava (categorias)"],
                       "Origem": "planilha do consultor" + (" (valor/data corrigidos pela base)"
                                                           if abs(liq[d] - r["valor"]) > 1 else "")})
    for f in falt.to_dict("records"):
        d = f["CNPJ/CPF"]
        cons = f["Consultor provável"] or rodizio(f["UF"], cidade_doc.get(d, ""), "")
        linhas.append({"Consultor": cons, "Cliente": f["Cliente"], "CNPJ/CPF": d, "Cidade": cidade_doc.get(d, ""),
                       "UF": f["UF"], "E-mail": str(email_doc.get(d, "")).strip().lower(),
                       "Telefone": contato(d, cli_doc, sab_doc, internos),
                       "Data da última compra": f["Última compra 1º sem"], "Valor 1º sem/2026 (R$)": f["Valor 1º sem/2026 (R$)"],
                       "O que comprava (categorias)": "",
                       "Origem": "INCLUÍDO na validação: CNPJ próprio que cumpre a regra (filial de grupo já listado ou faltante)"})
    rep_f = pd.DataFrame(linhas)

    # ---------------- Novos (uma linha por CNPJ)
    sai_nov = motivo("Novos", ["N1", "N2"])
    for d, g in nov[nov.doc.isin(sai_nov)].groupby("doc"):
        removidos.append({"Campanha": "Novos", "Consultor": g.Consultor.iloc[0], "Cliente": g.Cliente.iloc[0],
                          "CNPJ/CPF": d, "UF": g.UF.iloc[0], "Motivo": sai_nov[d]})
    linhas = []
    for d, g in nov[~nov.doc.isin(sai_nov)].groupby("doc", sort=False):
        r = g.iloc[0]
        tel = r.tel if r.tel and r.tel not in internos else contato(d, cli_doc, sab_doc, internos) or r.tel
        cons = r.Consultor if r.Consultor in consultor_esperado(r.UF, r.Cidade, tel) else rodizio(r.UF, r.Cidade, tel)
        linhas.append({"Consultor": cons, "Consultor na planilha": r.Consultor, "Cliente": r.Cliente, "CNPJ/CPF": d, "Cidade": r.Cidade, "UF": r.UF,
                       "E-mail": r.email, "Telefone": tel, "Qtd máquinas": g.Chassi.nunique(),
                       "Modelos": ", ".join(sorted(set(g["Modelo da máquina"]))), "Chassis": ", ".join(sorted(set(g.Chassi))),
                       "Entrega técnica mais recente": f"{g.et.max():%Y-%m-%d}", "Origem": "planilha do consultor"})
    et = sab[sab.necess.eq("Entrega Técnica")]
    servico_tab = []
    for f in falt_n.to_dict("records"):
        d = f["CNPJ/CPF"]
        cidade = cidade_doc.get(d, "")
        tel = contato(d, cli_doc, sab_doc, internos)
        cons = rodizio(f["UF"], cidade, tel)
        base = {"Consultor": cons, "Cliente": f["Cliente"], "CNPJ/CPF": d, "Cidade": cidade, "UF": f["UF"],
                "E-mail": str(email_doc.get(d, "")).strip().lower(), "Telefone": tel}
        if f["Tem ET"]:
            e = et[et.doc == d]
            linhas.append({**base, "Qtd máquinas": e.serie.nunique(), "Modelos": ", ".join(sorted(set(e.modelo.dropna()))),
                           "Chassis": ", ".join(sorted(set(e.serie.dropna()))),
                           "Entrega técnica mais recente": f"{e.data.max():%Y-%m-%d}",
                           "Origem": "INCLUÍDO na validação: tem ET no SAB e nenhuma NF de peça neste CNPJ/CPF"})
        else:
            servico_tab.append({**base, "Serviços no SAB/NF de serviço": f["Origem"],
                                "Possível revenda": bool(REVENDA.search(norm.razao(f["Cliente"] or "")))})
    nov_f = pd.DataFrame(linhas)
    serv_f = pd.DataFrame(servico_tab)

    for df in (rep_f, nov_f, serv_f):
        df.insert(0, "Regional", df.Consultor.map(REGIONAL).fillna("—"))
        df["Telefone p/ Léo"] = df.Telefone.map(lambda t: status_tel(t, internos))
        df.sort_values(["Regional", "Consultor", "Cliente"], inplace=True)

    disparo = pd.concat([rep_f.assign(Campanha="Repetição"), nov_f.assign(Campanha="Novos clientes")])
    disparo = disparo[disparo["Telefone p/ Léo"].eq("ok")]
    disparo = disparo.drop_duplicates("Telefone")[["Campanha", "Regional", "Consultor", "Cliente", "CNPJ/CPF", "UF",
                                                   "Telefone", "E-mail"]]
    resumo = pd.concat([
        rep_f.groupby(["Regional", "Consultor"]).agg(**{"Repetição (CNPJs)": ("CNPJ/CPF", "nunique"),
                                                        "Repetição R$ 1º sem": ("Valor 1º sem/2026 (R$)", "sum")}),
        nov_f.groupby(["Regional", "Consultor"]).agg(**{"Novos (CNPJs)": ("CNPJ/CPF", "nunique")}),
        disparo.groupby(["Regional", "Consultor"]).agg(**{"Prontos p/ Léo": ("Telefone", "size")}),
        serv_f.groupby(["Regional", "Consultor"]).agg(**{"Candidatos só serviço (decidir)": ("CNPJ/CPF", "nunique")}),
    ], axis=1).fillna(0)
    resumo.loc[("TOTAL", ""), :] = resumo.sum()
    return dict(Resumo=resumo, Repeticao=rep_f, Novos_Clientes=nov_f, Disparo_Leo=disparo,
                Removidos=pd.DataFrame(removidos), Candidatos_So_Servico=serv_f, Divergencias=ach)


LEIA_ME = [
    "Cliente = CNPJ/CPF completo. Homônimos e filiais contam separado.",
    "Repetição: NF de peça Balcão (CC xx0204, fora o RAP) no 1º sem/2026, líquida de devoluções, e nenhuma NF de peça (Balcão ou RAP) desde 01/07/2026.",
    "Novos clientes: entrega técnica (máquina) no SAB e nenhuma NF de peça em 2026 no mesmo CNPJ/CPF.",
    "Fora das duas listas: RAP (classe no cadastro ou compra no centro RAP), CNPJ da BMC e da HD Hyundai.",
    "LIMITE: a base de NF de peças começa em 01/2026. 'Nunca comprou peça' não foi checado em 2024–2025.",
    "Candidatos_So_Servico: fizeram revisão/corretiva/NF de serviço e não compraram peça em 2026. Entram em Novos só com aprovação.",
    "Disparo_Leo: só celulares válidos, sem telefone de consultor/BMC nem número repetido em 3+ clientes, um envio por número.",
    "Telefone p/ Léo diferente de 'ok': o consultor precisa atualizar o contato antes do disparo.",
]


if __name__ == "__main__":
    abas = consolidar(sys.argv[1], sys.argv[2])
    salvar(sys.argv[3], **{"Leia-me": pd.DataFrame({"Regras e limites": LEIA_ME}), **abas})
    print(abas["Resumo"].to_string())
    print(abas["Removidos"].groupby(["Campanha", "Motivo"]).size().to_string())
