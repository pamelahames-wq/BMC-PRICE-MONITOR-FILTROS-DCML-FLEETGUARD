"""Valida as planilhas corrigidas da campanha GPD contra as bases do Breno (só leitura).

Uso:  python -m campanha.validar <pasta_corrigida> [saida.xlsx]
A pasta corrigida tem uma subpasta por consultor (Repeticao_X.xlsx, Novos_Clientes_X.xlsx)
e o Resumo_Campanha_GPD.xlsx (aba Removidos). Nada nela é alterado.
"""
import re
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd

from . import norm
from .gerar import (CENTRO_RAP, CLASSES_RAP, CONSULTORES, ES_NORTE, MG_DDD, RAIZ, RAIZ_EXCLUIR, RAW,
                    UF_MAPA, VEND_RAP, _ler, _sem_acento, sab_base)

HOJE = pd.Timestamp("2026-10-21")
INI, FIM, CORTE = pd.Timestamp("2026-01-01"), pd.Timestamp("2026-06-30"), pd.Timestamp("2026-07-01")

FILTRO_REMOCAO = re.compile(r"\bPECAS?\b|\bPARTS\b|COMERCIO DE PECA|LEASING|ARRENDAMENTO|\bBANCO\b|"
                            r"MUNICIPIO|PREFEITURA|ENERGISA|TRACAO RENTAL")
SUSPEITA_REVENDA = re.compile(r"TRATOR|DIESEL|HIDRAULIC|MECANICA|DISTRIBUID|MULTIMARCA|ROLAMENTO|"
                              r"AUTOPECA|IMPORTA|REPRESENTA|COMERCIO DE MAQUINAS|MAQUINAS E PECA|"
                              r"RETIFICA|OFICINA")
EMAIL_GRATIS = {"gmail.com", "hotmail.com", "yahoo.com.br", "yahoo.com", "outlook.com", "live.com",
                "bol.com.br", "uol.com.br", "icloud.com", "terra.com.br", "hotmail.com.br",
                "outlook.com.br", "msn.com", "ig.com.br", "globo.com"}


# ----------------------------------------------------------------- leitura
def ler_planilhas(pasta):
    rep, nov = [], []
    for arq in sorted(Path(pasta).glob("*/*.xlsx")):
        d = pd.read_excel(arq, dtype=str)
        d["Consultor"] = arq.parent.name
        (rep if arq.name.startswith("Repeticao") else nov).append(d)
    rep, nov = pd.concat(rep, ignore_index=True), pd.concat(nov, ignore_index=True)
    removidos = pd.read_excel(Path(pasta) / "Resumo_Campanha_GPD.xlsx", sheet_name="Removidos", dtype=str)
    for d in (rep, nov, removidos):
        d["doc"] = d["CNPJ/CPF"].map(norm.so_digitos)
    rep["valor"] = pd.to_numeric(rep["Valor 1º sem/2026 (R$)"])
    rep["ult"] = pd.to_datetime(rep["Data da última compra"])
    return rep, nov, removidos


class Grupos:
    """União por raiz de CNPJ (CPF inteiro) e por razão social normalizada (transitiva)."""

    def __init__(self):
        self.p = {}

    def find(self, x):
        self.p.setdefault(x, x)
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]
            x = self.p[x]
        return x

    def union(self, a, b):
        a, b = self.find(a), self.find(b)
        if a != b:
            self.p[max(a, b)] = min(a, b)

    def add(self, doc, razao):
        if not doc:
            return
        self.union("D" + norm.raiz(doc), "D" + norm.raiz(doc))
        if razao and len(razao) >= 4:
            self.union("D" + norm.raiz(doc), "N" + razao)

    def de(self, doc):
        return self.find("D" + norm.raiz(doc)) if doc else ""


def bases():
    fat = _ler("f_Fat_Clientes.csv")
    fat["Valor"] = pd.to_numeric(fat.Valor)
    fat["DT"] = pd.to_datetime(fat.DTFATUR).dt.normalize()
    fat["CODCENCUS"] = fat.CODCENCUS.astype(int)
    notas = _ler("f_Notas.csv")
    cli = _ler("d_Clientes.csv")
    cli["doc"] = [norm.doc(m, t) or norm.doc(c, t) for m, c, t in zip(cli.CODIGO_MATRIZ, cli.CNPJ_CPF, cli.TIPPESSOA)]
    cli["razao_norm"] = cli.NOMEPARC.map(norm.razao)
    vend = _ler("d_Vendedores.csv")
    apelido = dict(zip(vend.CODVEND, vend.APELIDO.str.upper().str.strip()))
    fat = fat.merge(cli[["CODPARC", "doc", "razao_norm", "NOMEPARC", "CIDADE", "UF", "AD_CLASSRAP",
                         "VEND_PREFERENCIAL"]], on="CODPARC", how="left")
    fat["vendedor"] = fat.CODVEND.map(apelido)
    fat["venda"] = fat.TIPMOV.eq("V")
    fat["peca"] = fat.CODCENCUS.astype(str).str.endswith(("204", "207"))
    fat["balcao"] = fat.CODCENCUS.astype(str).str.endswith("204") & fat.CODCENCUS.ne(CENTRO_RAP)
    sab = sab_base()
    rap_planilha = set(pd.read_csv(RAW / "carteira_rap_parcial.csv", dtype=str).CODPARC)
    return fat, notas, cli, vend, sab, rap_planilha


# ----------------------------------------------------------------- validação
def validar(pasta, destino=None):
    rep, nov, removidos = ler_planilhas(pasta)
    fat, notas, cli, vend, sab, rap_planilha = bases()

    G = Grupos()
    for d, r in zip(cli.doc, cli.razao_norm):
        G.add(d, r)
    for d in sab.doc:            # nome do SAB não une: tem nomes genéricos (filial, cidade, "BMC MATRIZ")
        G.add(d, "")
    for d in pd.concat([rep.doc, nov.doc, removidos.doc]):
        G.add(d, "")
    fat["grupo"] = fat.doc.fillna("").map(G.de)
    sab["grupo"] = sab.doc.map(G.de)
    rep["grupo"] = rep.doc.map(G.de)
    nov["grupo"] = nov.doc.map(G.de)
    removidos["grupo"] = removidos.doc.map(G.de)

    resumo, abas = [], {}

    def item(n, desc, qtd, grave=False, obs=""):
        st = "OK" if qtd == 0 else ("ERRO" if grave else "ATENÇÃO")
        resumo.append({"Item": n, "Verificação": desc, "Status": st, "Qtd": qtd, "Observação": obs})

    # RAP / exclusões fixas
    rap_grupos = set(fat.loc[fat.CODPARC.isin(rap_planilha), "grupo"])
    rap_grupos |= set(fat.loc[fat.AD_CLASSRAP.fillna("").str.upper().isin(CLASSES_RAP), "grupo"])
    rap_grupos |= set(fat.loc[fat.vendedor.isin(VEND_RAP) | fat.VEND_PREFERENCIAL.fillna("").str.upper().isin(VEND_RAP), "grupo"])
    rap_grupos |= set(fat.loc[fat.CODCENCUS.eq(CENTRO_RAP) & fat.venda, "grupo"])
    rap_cad = set(cli.loc[cli.AD_CLASSRAP.fillna("").str.upper().isin(CLASSES_RAP), "doc"].map(G.de))

    vendas_bal = fat[fat.venda & fat.balcao]
    dev_bal = fat[fat.TIPMOV.eq("D") & fat.balcao]
    s1 = vendas_bal[(vendas_bal.DT >= INI) & (vendas_bal.DT <= FIM)]
    s1_dev = dev_bal[(dev_bal.DT >= INI) & (dev_bal.DT <= FIM)]
    liq_s1 = s1.groupby("grupo").Valor.sum().sub(s1_dev.groupby("grupo").Valor.sum(), fill_value=0)
    bruto_s1 = s1.groupby("grupo").Valor.sum()
    peca_v = fat[fat.venda & fat.peca]
    s2 = peca_v[peca_v.DT >= CORTE]
    s2_grupos = set(s2.grupo)
    ult_bal = vendas_bal.groupby("grupo").DT.max()

    # ---------------- 1. Repetição: compra no 1º sem, data e valor
    div = []
    for r in rep.itertuples():
        g = r.grupo
        if g not in bruto_s1.index:
            div.append(dict(item=1, Consultor=r.Consultor, Cliente=r.Cliente, CNPJ=r.doc, problema="Sem NF de peça Balcão no 1º sem/2026"))
            continue
        if abs(liq_s1.get(g, 0) - r.valor) > 1:
            div.append(dict(item=1, Consultor=r.Consultor, Cliente=r.Cliente, CNPJ=r.doc,
                            problema="Valor 1º sem diverge", na_planilha=r.valor, na_base=round(liq_s1.get(g, 0), 2)))
        if ult_bal.get(g) != r.ult:
            div.append(dict(item=1, Consultor=r.Consultor, Cliente=r.Cliente, CNPJ=r.doc,
                            problema="Data da última compra diverge", na_planilha=r.ult.date(), na_base=ult_bal.get(g).date()))
    n1 = sum(1 for d in div if d["item"] == 1)
    item(1, "Repetição: NF de peça no 1º sem, data e valor (±R$1) batem", n1, grave=True,
         obs="valor = vendas Balcão − devoluções Balcão no 1º sem, somando o grupo (raiz/razão social)")

    # ---------------- 2. compra a partir de 01/07 no grupo
    n2 = 0
    for r in rep.itertuples():
        x = s2[s2.grupo == r.grupo]
        if len(x):
            n2 += 1
            u = x.sort_values("DT").iloc[-1]
            div.append(dict(item=2, Consultor=r.Consultor, Cliente=r.Cliente, CNPJ=r.doc,
                            problema="Grupo comprou peça depois de 01/07",
                            na_base=f"CNPJ {u.doc} ({u.NOMEPARC}) em {u.DT:%d/%m/%Y}, NF {u.NUNOTA}, R$ {u.Valor:,.2f}"))
    item(2, "Repetição: nenhum CNPJ do grupo com NF de peça desde 01/07/2026", n2, grave=True)

    # ---------------- 3. faltantes
    na_lista = set(rep.grupo) | set(removidos.grupo)
    candidatos = set(liq_s1[liq_s1 > 0].index) - s2_grupos - rap_grupos - na_lista - {""}
    falt = []
    for g in candidatos:
        h = s1[s1.grupo == g]
        docs = set(h.doc.dropna())
        if any(d[:8] in RAIZ_EXCLUIR for d in docs):
            continue
        p = h.groupby(["doc", "NOMEPARC", "CIDADE", "UF"]).Valor.sum().sort_values().index[-1]
        nome_n = norm.razao(p[1])
        if FILTRO_REMOCAO.search(nome_n):
            continue
        u = vendas_bal[vendas_bal.grupo == g].sort_values("DT").iloc[-1]
        falt.append({"Cliente": p[1], "CNPJ/CPF": p[0], "Cidade": p[2], "UF": p[3],
                     "Valor 1º sem/2026 (R$)": round(liq_s1[g], 2), "Data da última compra": u.DT.date(),
                     "Último vendedor": u.vendedor, "Outros CNPJs do grupo": ", ".join(sorted(docs - {p[0]})),
                     "Motivo provável": "não estava na base da 1ª geração (cadastro sem CNPJ ou grupo por razão social)"})
    falt = pd.DataFrame(falt).sort_values("Valor 1º sem/2026 (R$)", ascending=False) if falt else pd.DataFrame()
    item(3, "Repetição: clientes que deveriam estar e não estão", len(falt))

    # ---------------- 4. devoluções e cancelamentos
    dev_s2 = dev_bal[dev_bal.DT >= CORTE].groupby("grupo").Valor.sum()
    n4a = 0
    for r in rep.itertuples():
        d2 = dev_s2.get(r.grupo, 0)
        if d2 and liq_s1.get(r.grupo, 0) - d2 <= 1:
            n4a += 1
            div.append(dict(item=4, Consultor=r.Consultor, Cliente=r.Cliente, CNPJ=r.doc,
                            problema="Compra do 1º sem devolvida depois de 01/07 (líquido ≤ 0)",
                            na_planilha=r.valor, na_base=round(liq_s1.get(r.grupo, 0) - d2, 2)))
    # saiu porque a compra do 2º sem foi devolvida
    v2 = s2.groupby("grupo").Valor.sum()
    d2_todas = fat[fat.TIPMOV.eq("D") & fat.peca & (fat.DT >= CORTE)].groupby("grupo").Valor.sum()
    so_devolvido = [g for g in v2.index if d2_todas.get(g, 0) >= v2[g] - 1
                    and liq_s1.get(g, 0) > 0 and g not in rap_grupos and g not in na_lista]
    for g in so_devolvido:
        h = fat[fat.grupo == g].sort_values("DT").iloc[-1]
        falt = pd.concat([falt, pd.DataFrame([{
            "Cliente": h.NOMEPARC, "CNPJ/CPF": h.doc, "Cidade": h.CIDADE, "UF": h.UF,
            "Valor 1º sem/2026 (R$)": round(liq_s1[g], 2), "Último vendedor": h.vendedor,
            "Motivo provável": "compra do 2º sem foi toda devolvida"}])], ignore_index=True)
    item(4, "Devoluções: entrou com compra devolvida / saiu por compra do 2º sem devolvida",
         n4a + len(so_devolvido),
         obs=f"{n4a} na lista com 1º sem devolvido; {len(so_devolvido)} fora por 2º sem devolvido. "
             "A base não traz NF cancelada (só V e D), então cancelamento não é verificável.")

    # ---------------- 5. BMC, funcionário, garantia Hyundai, RAP
    nome_func = set()
    for e in vend[vend.ATIVO.eq("S") & vend.EMAIL.fillna("").str.contains("bmchyundai", case=False)].EMAIL:
        a = e.split("@")[0].lower().split(".")
        if len(a) == 2:
            nome_func.add((a[0].upper(), a[1].upper()))

    def funcionario(nome, doc):
        if len(doc) != 11:
            return False
        t = norm.razao(nome).split()
        return len(t) >= 2 and (t[0], t[-1]) in nome_func

    divn = []
    n5, n5b = 0, 0
    rap_forte = rap_grupos  # carteira RAP, vendedor RAP ou compra no centro RAP
    for camp, df, destino_div in (("Repetição", rep, div), ("Novos", nov, divn)):
        for r in df.itertuples():
            grave = True
            if r.doc[:8] in RAIZ_EXCLUIR:
                motivo = "CNPJ BMC / garantia Hyundai"
            elif r.grupo in rap_forte:
                motivo = "cliente RAP (carteira RAP, vendedor RAP ou compra no centro RAP)"
            elif r.grupo in rap_cad:
                motivo, grave = "AD_CLASSRAP preenchido no cadastro, mas sem vendedor/compra RAP — confirmar", False
            elif funcionario(r.Cliente, r.doc):
                motivo, grave = "possível funcionário (CPF com nome igual a e-mail de vendedor BMC) — confirmar", False
            else:
                continue
            if grave:
                n5 += 1
            else:
                n5b += 1
            destino_div.append(dict(item=5, Consultor=r.Consultor, Cliente=r.Cliente, CNPJ=r.doc, problema=motivo))
    item(5, "Sem CNPJ BMC, garantia Hyundai ou RAP confirmado", n5, grave=True)
    item("5b", "Possível RAP (só AD_CLASSRAP) ou possível funcionário", n5b,
         obs="não há cadastro de funcionários: checado CPF com nome igual a e-mail de vendedor BMC")

    # ---------------- 6. Novos: ET e modelo/chassi
    et = sab[sab.necess.eq("Entrega Técnica")]
    et_doc = et.groupby("doc").agg(chassis=("serie", lambda s: set(s.dropna().str.strip().str.upper())),
                                   modelos=("modelo", lambda s: set(s.dropna().str.strip().str.upper())))
    n6 = 0
    for r in nov.to_dict("records"):
        r = pd.Series(r)
        ch = {c.strip().upper() for c in str(r.Chassis).split(",") if c.strip() and c != "nan"}
        mo = {c.strip().upper() for c in str(r.Modelos).split(",") if c.strip() and c != "nan"}
        if r.doc not in et_doc.index:
            divn.append(dict(item=6, Consultor=r.Consultor, Cliente=r.Cliente, CNPJ=r.doc, problema="Sem entrega técnica no SAB"))
            n6 += 1
            continue
        e = et_doc.loc[r.doc]
        if not ch <= e.chassis or not mo <= e.modelos or int(r["Qtd máquinas"]) != len(ch):
            n6 += 1
            divn.append(dict(item=6, Consultor=r.Consultor, Cliente=r.Cliente, CNPJ=r.doc,
                             problema="Modelo/chassi/qtd não batem com o SAB",
                             na_planilha=f"{sorted(mo)} {sorted(ch)}", na_base=f"{sorted(e.modelos)} {sorted(e.chassis)}"))
    item(6, "Novos: todo CNPJ com ET no SAB e modelo/chassi batendo", n6, grave=True)

    # ---------------- 7. Novos: nenhuma NF de peça no grupo
    grupos_com_peca = set(peca_v.grupo)
    n7 = 0
    for r in nov.itertuples():
        if r.grupo in grupos_com_peca:
            n7 += 1
            u = peca_v[peca_v.grupo == r.grupo].sort_values("DT").iloc[-1]
            divn.append(dict(item=7, Consultor=r.Consultor, Cliente=r.Cliente, CNPJ=r.doc,
                             problema="Grupo tem NF de peça",
                             na_base=f"CNPJ {u.doc} ({u.NOMEPARC}) em {u.DT:%d/%m/%Y}, NF {u.NUNOTA}, R$ {u.Valor:,.2f}"))
    item(7, "Novos: nenhum CNPJ do grupo com NF de peça (base 2026)", n7, grave=True)

    # ---------------- 8, 9, 10
    et_min = et.data_os.min()
    item(8, "ET mais antiga no SAB", 0,
         obs=f"ET mais antiga: {et_min:%d/%m/%Y}. A base do SAB começa em 01/2025: não há máquinas anteriores a 2025 para incluir (0 CNPJs).")
    nov["et_recente"] = pd.to_datetime(nov["Entrega técnica mais recente"])
    nov["et_antiga"] = pd.to_datetime(nov["Entrega técnica mais antiga"])
    n9 = int((nov.et_recente >= HOJE - pd.Timedelta(days=90)).sum())
    resumo.append({"Item": 9, "Verificação": "Novos com ET nos últimos 90 dias (máquina nova, em garantia)",
                   "Status": "INFO", "Qtd": n9, "Observação": f"ET desde {HOJE - pd.Timedelta(days=90):%d/%m/%Y}; só informativo"})
    n10 = int((nov.et_antiga < "2026-01-01").sum())
    resumo.append({"Item": 10, "Verificação": "Novos com máquina de 2025 (risco de peça comprada em 2025)",
                   "Status": "ATENÇÃO", "Qtd": n10,
                   "Observação": "'Nunca comprou peça' verificado só com NFs de 2026 — PROVISÓRIO até chegar o histórico 2024–2025"})

    # ---------------- 11. duplicidades
    dup = []
    for nome, df in (("Repetição", rep), ("Novos", nov)):
        x = df.groupby("grupo").Consultor.nunique()
        for g in x[x > 1].index:
            dup.append(dict(item=11, problema=f"{nome}: mesmo grupo em mais de um consultor",
                            CNPJ=", ".join(df[df.grupo == g].doc), Consultor=", ".join(df[df.grupo == g].Consultor)))
        y = df.doc.value_counts()
        for d in y[y > 1].index:
            dup.append(dict(item=11, problema=f"{nome}: CNPJ repetido", CNPJ=d, Consultor=", ".join(df[df.doc == d].Consultor)))
    ambos = set(rep.grupo) & set(nov.grupo)
    for g in ambos:
        dup.append(dict(item=11, problema="Grupo nas duas campanhas", CNPJ=", ".join(set(rep[rep.grupo == g].doc) | set(nov[nov.grupo == g].doc)),
                        Consultor=", ".join(set(rep[rep.grupo == g].Consultor) | set(nov[nov.grupo == g].Consultor))))
    item(11, "Nenhum CNPJ/grupo em 2 consultores ou nas 2 campanhas", len(dup), grave=True)

    # ---------------- 12. carteira
    def esperado(uf, cidade, fone):
        uf = _sem_acento(uf)
        if uf == "ES":
            return {"YURI"} if _sem_acento(cidade) in ES_NORTE else {"JHONY"}
        if uf == "MG":
            d = norm.so_digitos(fone)[2:4] if fone else ""
            return {MG_DDD[d]} if d in MG_DDD else {"MATEUS", "DANIEL"}
        return set(UF_MAPA.get(uf, []))

    cart = []
    for r in nov.itertuples():
        e = esperado(r.UF, r.Cidade, r.Telefone if isinstance(r.Telefone, str) else "")
        if r.Consultor not in e:
            cart.append({"Campanha": "Novos", "Consultor": r.Consultor, "Cliente": r.Cliente, "CNPJ/CPF": r.doc,
                         "UF": r.UF, "Cidade": r.Cidade, "Consultor esperado": " ou ".join(sorted(e)) or "UF sem consultor no mapa",
                         "Motivo": "fora do mapa de UF"})
    ult_vend = vendas_bal.sort_values("DT").groupby("grupo").vendedor.last().map(lambda v: CONSULTORES.get(v or "", None))
    for r in rep.itertuples():
        e = esperado(r.UF, r.Cidade, r.Telefone if isinstance(r.Telefone, str) else "")
        uv = ult_vend.get(r.grupo)
        if r.Consultor not in e and r.Consultor != uv:
            cart.append({"Campanha": "Repetição", "Consultor": r.Consultor, "Cliente": r.Cliente, "CNPJ/CPF": r.doc,
                         "UF": r.UF, "Cidade": r.Cidade, "Consultor esperado": uv or " ou ".join(sorted(e)),
                         "Motivo": "fora da UF e não foi o vendedor da última compra"})
    cart = pd.DataFrame(cart)
    rep_fora_ok = sum(1 for r in rep.itertuples()
                      if r.Consultor not in esperado(r.UF, r.Cidade, r.Telefone if isinstance(r.Telefone, str) else "")
                      and r.Consultor == ult_vend.get(r.grupo))
    item(12, "Carteira conforme mapa (Novos) / último vendedor (Repetição)", len(cart),
         obs=f"{rep_fora_ok} clientes de repetição estão fora da UF mas com o vendedor da última compra (aceito)")

    # ---------------- 13. possíveis revendas
    s1n = s1.merge(notas[["NUNOTA", "CODPROD"]], on="NUNOTA", how="left")
    pns = s1n.groupby("grupo").CODPROD.nunique()
    nfs = s1.groupby("grupo").NUNOTA.nunique()
    rev = []
    for r in rep.itertuples():
        nome = norm.razao(r.Cliente)
        m = SUSPEITA_REVENDA.search(nome)
        npn, nnf = int(pns.get(r.grupo, 0)), int(nfs.get(r.grupo, 0))
        if m or npn >= 40 or r.valor >= 150_000:
            rev.append({"Campanha": "Repetição", "Consultor": r.Consultor, "Cliente": r.Cliente, "CNPJ/CPF": r.doc,
                        "UF": r.UF, "Valor 1º sem/2026 (R$)": r.valor, "PNs diferentes 1º sem": npn, "NFs 1º sem": nnf,
                        "Sinal": "; ".join(filter(None, [f"nome: {m.group(0)}" if m else "",
                                                         "muitos PNs" if npn >= 40 else "",
                                                         "volume alto" if r.valor >= 150_000 else ""]))})
    for r in nov.itertuples():
        m = SUSPEITA_REVENDA.search(norm.razao(r.Cliente))
        if m:
            rev.append({"Campanha": "Novos", "Consultor": r.Consultor, "Cliente": r.Cliente, "CNPJ/CPF": r.doc,
                        "UF": r.UF, "Sinal": f"nome: {m.group(0)}"})
    rev = pd.DataFrame(rev)
    item(13, "Possíveis revendas que escaparam do filtro por nome/volume", len(rev))

    # ---------------- 14. contatos
    internos = telefones_internos(cli, sab)
    con = []
    todos = pd.concat([rep.assign(Campanha="Repetição"), nov.assign(Campanha="Novos")], ignore_index=True)
    todos["tel"] = todos.Telefone.fillna("").str.strip()
    todos["tel_d"] = todos.tel.map(lambda t: norm.so_digitos(t)[2:])

    def add(r, contato, problema):
        con.append(dict(Campanha=r["Campanha"], Consultor=r["Consultor"], Cliente=r["Cliente"], CNPJ=r["doc"],
                        Contato=contato, Problema=problema))

    for r in todos.to_dict("records"):
        t = r["tel"]
        if r["tel_d"] in internos:
            add(r, t, f"telefone de funcionário/consultor BMC ({internos[r['tel_d']]})")
        elif t and not re.fullmatch(r"\+55[1-9][1-9]9\d{8}", t):
            add(r, t, "não é celular +55 com 13 dígitos (fixo ou incompleto)")
    g_tel = todos[(todos.tel != "") & ~todos.tel_d.isin(internos)].groupby("tel").grupo.nunique()
    for t in g_tel[g_tel > 1].index:
        x = todos[todos.tel == t]
        for r in x.to_dict("records"):
            add(r, t, f"mesmo telefone em {x.grupo.nunique()} clientes diferentes")
    todos["email"] = todos["Contato (e-mail)"].fillna("").str.strip().str.lower()
    todos["dom"] = todos.email.str.extract(r"@([\w.-]+)")[0]
    for r in todos.dropna(subset=["dom"]).to_dict("records"):
        dom = r["dom"]
        if "bmchyundai" in dom:
            add(r, r["email"], "e-mail de funcionário BMC no cadastro do cliente")
            continue
        if dom in EMAIL_GRATIS:
            continue
        base_dom = dom.split(".")[0].upper()
        palavras = [w for w in norm.razao(r["Cliente"]).split() if len(w) >= 4]
        if len(base_dom) >= 4 and base_dom not in norm.razao(r["Cliente"]).replace(" ", "") \
                and not any(w in base_dom for w in palavras):
            outro = todos[(todos.dom == dom) & (todos.grupo != r["grupo"])]
            if len(outro) or re.search(r"PECA|MAQUINA|TRATOR|CONTAB|LEASING|DIESEL", base_dom):
                add(r, r["email"], "e-mail de outra empresa" +
                    (f" (mesmo domínio em {outro.grupo.nunique()} outro(s) cliente(s))" if len(outro) else ""))
    con = pd.DataFrame(con)
    eh_int = con.Problema.str.startswith("telefone de funcionário")
    item(14, "Contatos: telefone fixo/incompleto, repetido ou e-mail de terceiro/BMC", int((~eh_int).sum()))
    item("14b", "Telefone na lista é de funcionário/consultor BMC (LÉO dispararia para a BMC)", int(eh_int.sum()), grave=True,
         obs="origem: sab_cliente.fone guarda o telefone de quem abriu a OS; o mesmo número aparece em 5+ clientes")

    # ---------------- 15. telefones recuperáveis
    precisa = todos[(todos.tel == "") | todos.tel_d.isin(internos)].copy()
    precisa["motivo"] = ["sem telefone" if t == "" else "telefone interno BMC" for t in precisa.tel]
    fontes = defaultdict(list)
    for d, txt in zip(cli.doc, cli.CONTATO):
        for f in norm.telefones(txt):
            fontes[G.de(d)].append((f, "d_Clientes (CONTATO)"))
    hs = RAIZ / "cache" / "hubspot_telefones.csv"
    if hs.exists():
        h = pd.read_csv(hs, dtype=str).merge(cli[["CODPARC", "doc"]], on="CODPARC", how="left")
        ordem = {"contact.mobilephone": 0, "contact.hs_whatsapp_phone_number": 1, "contact.phone": 2,
                 "company.phone": 3, "company.telefone_backup": 4, "company.telefone_cnpj": 5}
        h["o"] = h.origem.map(ordem).fillna(9)
        for x in h.sort_values("o").itertuples():
            for f in norm.telefones(x.numero_raw):
                nome = x.contato_nome if isinstance(x.contato_nome, str) else ""
                fontes[G.de(x.doc)].append((f, f"HubSpot {x.origem}" + (f" ({nome})" if nome else "")))
    rec = []
    for r in precisa.to_dict("records"):
        fs = [x for x in fontes.get(r["grupo"], []) if x[0] not in internos]
        cel = [x for x in fs if norm.celular(x[0])]
        esc = cel[0] if cel else (fs[0] if fs else None)
        rec.append({"Campanha": r["Campanha"], "Consultor": r["Consultor"], "Cliente": r["Cliente"], "CNPJ/CPF": r["doc"],
                    "Situação atual": r["motivo"], "Telefone atual": r["tel"],
                    "Telefone recuperado": norm.e164(esc[0]) if esc else "",
                    "Tipo": ("celular" if cel else "fixo") if esc else "",
                    "Fonte": esc[1] if esc else "não encontrado (cadastro Sankhya / HubSpot)"})
    rec = pd.DataFrame(rec)
    r374 = rec[(rec.Campanha == "Repetição") & (rec["Situação atual"] == "sem telefone")]
    resumo.append({"Item": 15, "Verificação": "Repetição sem telefone: celular recuperado no cadastro/HubSpot",
                   "Status": "INFO", "Qtd": int((r374.Tipo == "celular").sum()),
                   "Observação": f"de {len(r374)} sem telefone: {(r374.Tipo == 'celular').sum()} celular, "
                                 f"{(r374.Tipo == 'fixo').sum()} só fixo, {(r374.Tipo == '').sum()} sem nada. "
                                 f"Telefones internos BMC substituíveis: {((rec['Situação atual'] == 'telefone interno BMC') & (rec.Tipo == 'celular')).sum()} "
                                 f"de {(rec['Situação atual'] == 'telefone interno BMC').sum()}"})

    return dict(rep=rep, nov=nov, resumo=pd.DataFrame(resumo), div=pd.DataFrame(div), falt=falt,
                divn=pd.DataFrame(divn + dup), cart=cart, rev=rev, con=con, rec=rec,
                projecao=projecao(rep, nov, div, falt, cart, dup))


def telefones_internos(cli, sab):
    """Telefones da BMC: aparecem no SAB em 5+ clientes diferentes ou no cadastro com e-mail @bmchyundai."""
    linhas = [(f, norm.raiz(d)) for d, t in zip(sab.doc, sab.fone) for f in norm.telefones(t)]
    x = pd.DataFrame(linhas, columns=["f", "r"]).drop_duplicates().groupby("f").r.nunique()
    internos = {f: f"aparece no SAB em {n} clientes" for f, n in x[x >= 5].items()}
    for e, t in zip(cli.EMAIL.fillna(""), cli.CONTATO):
        if "bmchyundai" in e.lower():
            for f in norm.telefones(t):
                internos.setdefault(f, f"cadastrado com e-mail {e.strip().lower()}")
    return internos


def projecao(rep, nov, div, falt, cart, dup):
    """Contagem por consultor se as correções recomendadas forem aplicadas."""
    r = rep[["doc", "Consultor"]].copy()
    sai = {d["CNPJ"] for d in div if d["item"] in (1, 2, 4)}
    r = r[~r.doc.isin(sai)]
    if len(falt):
        f = falt.assign(doc=falt["CNPJ/CPF"])
        f["Consultor"] = [CONSULTORES.get(v or "") or "" for v in f["Último vendedor"]]
        r = pd.concat([r, f[["doc", "Consultor"]]])
    n = nov[["doc", "Consultor", "grupo"]].copy()
    if len(cart):
        mover = {x["CNPJ/CPF"]: x["Consultor esperado"].split(" ou ")[0] for x in cart.to_dict("records")
                 if x["Consultor esperado"] and "sem consultor" not in x["Consultor esperado"]}
        r["Consultor"] = [mover.get(d, c) for d, c in zip(r.doc, r.Consultor)]
        n["Consultor"] = [mover.get(d, c) for d, c in zip(n.doc, n.Consultor)]
    # mesmo grupo com 2 consultores: fica com o consultor do 1º CNPJ
    dono = n.groupby("grupo").Consultor.first()
    n["Consultor"] = n.grupo.map(dono)
    t = pd.concat([r.Consultor.value_counts().rename("Repetição"),
                   n.drop_duplicates("doc").Consultor.value_counts().rename("Novos")], axis=1).fillna(0).astype(int)
    t.loc["TOTAL"] = t.sum()
    return t


ABAS = [("Resumo", "resumo"), ("Divergencias_Repeticao", "div"), ("Faltantes_Repeticao", "falt"),
        ("Divergencias_Novos", "divn"), ("Carteira_Fora_UF", "cart"), ("Possiveis_Revendas", "rev"),
        ("Contatos_Problema", "con"), ("Telefones_Recuperados", "rec"), ("Projecao_Consultor", "projecao")]


def salvar(r, destino):
    from openpyxl.utils import get_column_letter
    Path(destino).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(destino, engine="openpyxl") as w:
        for aba, k in ABAS:
            df = r[k] if len(r[k]) else pd.DataFrame({"": ["sem ocorrências"]})
            df.to_excel(w, sheet_name=aba, index=(k == "projecao"))
            ws = w.sheets[aba]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for i, col in enumerate(([df.index.name or "Consultor"] if k == "projecao" else []) + list(df.columns), 1):
                vals = [str(v) for v in (df[col].head(300) if col in df.columns else df.index)]
                ws.column_dimensions[get_column_letter(i)].width = min(max(10, max(map(len, [str(col)] + vals)) + 2), 70)


if __name__ == "__main__":
    pasta = sys.argv[1]
    destino = sys.argv[2] if len(sys.argv) > 2 else str(RAIZ / "saida" / "validacao_campanha_gpd.xlsx")
    r = validar(pasta, destino)
    salvar(r, destino)
    print(r["resumo"].to_string(index=False))
    print(r["projecao"].to_string())
