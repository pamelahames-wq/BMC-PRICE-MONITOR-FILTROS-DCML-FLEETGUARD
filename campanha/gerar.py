"""Gera, por consultor de Balcão, as planilhas da campanha Repetição e Novos Clientes.

Uso:  python -m campanha.gerar
Lê as bases do Breno copiadas em cache/raw (f_Notas, f_Fat_Clientes, d_Clientes, d_Vendedores,
sab_atividade, sab_cliente, sab_equipamento) e a planilha de margem em dados/ (categorias dos PNs).
"""
import itertools
import shutil
import unicodedata
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter

from . import norm

RAIZ = Path(__file__).resolve().parent.parent
RAW = RAIZ / "cache" / "raw"
SAIDA = RAIZ / "saida"

INICIO_1S26, FIM_1S26, CORTE = "2026-01-01", "2026-06-30", "2026-07-01"
CENTRO_RAP = 1030207
RAIZ_EXCLUIR = {"14168536", "17343777", "17343798", "20010472",  # BMC
                "13837846"}                                         # HD Hyundai CE Brasil (garantia)
VEND_RAP = {"PLOUREIRO", "PLOUREIO", "ELISASILVA", "FALEXANDRE", "SLUNA"}
CLASSES_RAP = {"STANDARD", "OURO", "DIAMANTE"}

# apelido Sankhya -> pasta
CONSULTORES = {
    "RANNIERICKSONF": "RANNIERICKSON", "ALEFSANTOS": "ALEF", "RICARDO SILVA": "RICARDO",
    "TPEREIRA": "THIAGO", "YURI SANTOS": "YURI", "JDIAS": "JHONY", "JHONNY DIAS": "JHONY",
    "RENANARAUJO": "RENAN", "MATEUSNATANAEL": "MATEUS", "DANIELSOARES": "DANIEL",
    "GDORST": "GEISON", "PCARVALHO": "PABBLO", "EDUARDOGAZOTI": "EDU", "FABIO LOLLI": "FABIO",
    "KARENVEIGA": "KAREN", "RDIAS": "RAYLANDER",
}
UF_MAPA = {
    "MA": ["RANNIERICKSON"], "PI": ["RANNIERICKSON"],
    **{u: ["ALEF"] for u in ["CE", "RN", "PB", "PE", "AL", "SE"]},
    **{u: ["RICARDO", "THIAGO"] for u in ["GO", "PA", "TO", "AM", "RR", "AP", "DF"]},
    "BA": ["YURI"],
    **{u: ["RENAN"] for u in ["SC", "RS", "PR"]},
    **{u: ["GEISON"] for u in ["MT", "RO", "AC"]},
    "MS": ["PABBLO"],
    "SP": ["EDU", "FABIO", "KAREN"],
}
ES_NORTE = {
    "LINHARES", "SAO MATEUS", "COLATINA", "ARACRUZ", "NOVA VENECIA", "BARRA DE SAO FRANCISCO",
    "PINHEIROS", "MONTANHA", "CONCEICAO DA BARRA", "JAGUARE", "SOORETAMA", "RIO BANANAL",
    "SAO GABRIEL DA PALHA", "PEDRO CANARIO", "ECOPORANGA", "MANTENOPOLIS", "AGUIA BRANCA",
    "VILA VALERIO", "BOA ESPERANCA", "PANCAS", "BAIXO GUANDU", "MARILANDIA",
    "GOVERNADOR LINDENBERG", "JOAO NEIVA", "IBIRACU", "PONTO BELO", "MUCURICI", "VILA PAVAO",
    "SAO DOMINGOS DO NORTE", "ALTO RIO NOVO", "SAO ROQUE DO CANAA", "ITAGUACU", "ITARANA",
    "LARANJA DA TERRA", "SANTA TERESA",
}
MG_DDD = {"31": "MATEUS", "32": "MATEUS", "33": "MATEUS",
          "34": "DANIEL", "35": "DANIEL", "37": "DANIEL", "38": "DANIEL"}


def _ler(nome):
    d = pd.read_csv(RAW / nome, dtype=str, low_memory=False)
    d.columns = [c.split("[")[-1].strip("]") for c in d.columns]
    return d


def _sem_acento(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return s.upper().strip()


def _chave_cliente(doc_):
    return norm.raiz(doc_) if doc_ else ""


def carregar():
    fat = _ler("f_Fat_Clientes.csv")
    fat["Valor"] = pd.to_numeric(fat.Valor)
    fat["DT"] = pd.to_datetime(fat.DTFATUR).dt.normalize()
    fat["CODCENCUS"] = fat.CODCENCUS.astype(int)
    notas = _ler("f_Notas.csv")
    notas["BRUTO"] = pd.to_numeric(notas.BRUTO)

    cli = _ler("d_Clientes.csv")
    cli["doc"] = [norm.doc(m, t) if norm.doc(m, t) else norm.doc(c, t)
                  for m, c, t in zip(cli.CODIGO_MATRIZ, cli.CNPJ_CPF, cli.TIPPESSOA)]
    cli["chave"] = cli.doc.map(_chave_cliente)
    cli["razao_norm"] = cli.NOMEPARC.map(norm.razao)

    vend = _ler("d_Vendedores.csv")
    apelido = dict(zip(vend.CODVEND, vend.APELIDO.str.upper().str.strip()))

    cat = pd.read_excel(RAIZ / "dados" / "Reestruturacao_Margem_2026_FINAL.xlsx",
                        sheet_name="Detalhe por Part Number", usecols=["CODPROD", "Categoria"])
    categoria = dict(zip(cat.CODPROD.astype(str), cat.Categoria))

    rap = set()
    arq_rap = RAW / "carteira_rap_parcial.csv"
    if arq_rap.exists():
        rap |= set(pd.read_csv(arq_rap, dtype=str).CODPARC)
    return fat, notas, cli, apelido, categoria, rap


def fones_cliente(cli, sab_fones):
    """Primeiro celular válido (senão o primeiro fixo) por chave de cliente."""
    tel = {}
    for chave, txt in itertools.chain(zip(cli.chave, cli.CONTATO), sab_fones):
        if not chave:
            continue
        for f in norm.telefones(txt):
            atual = tel.get(chave)
            if atual is None or (not norm.celular(atual) and norm.celular(f)):
                tel[chave] = f
    return tel


class Distribuidor:
    """Consultor pela UF; UFs com mais de um consultor recebem em rodízio (lista já ordenada)."""

    def __init__(self):
        self.ciclos = {}

    def __call__(self, uf, cidade, fone, sem_compra_12m):
        uf = _sem_acento(uf)
        if uf == "ES":
            return "YURI" if _sem_acento(cidade) in ES_NORTE else "JHONY"
        if uf == "MG":
            if sem_compra_12m:
                return "RAYLANDER"
            if (fone or "")[:2] in MG_DDD:
                return MG_DDD[fone[:2]]
            uf_ciclo = "MG"
            opcoes = ["MATEUS", "DANIEL"]
            return next(self.ciclos.setdefault(uf_ciclo, itertools.cycle(opcoes)))
        opcoes = UF_MAPA.get(uf)
        if not opcoes:
            return "SEM_CONSULTOR"
        if len(opcoes) == 1:
            return opcoes[0]
        ciclo = self.ciclos.setdefault(uf, itertools.cycle(opcoes))
        return next(ciclo)


def gerar():
    fat, notas, cli, apelido, categoria, rap_planilha = carregar()
    cli_por_parc = cli.set_index("CODPARC")
    fat = fat.merge(cli[["CODPARC", "doc", "chave", "razao_norm", "NOMEPARC", "CIDADE", "UF",
                         "EMAIL", "AD_CLASSRAP", "VEND_PREFERENCIAL"]], on="CODPARC", how="left")
    fat["vendedor"] = fat.CODVEND.map(apelido)
    # mesma razão social normalizada = mesmo cliente
    chave_por_nome = (fat[fat.razao_norm.str.len() >= 4].groupby("razao_norm").chave.min())
    fat["cliente"] = fat.razao_norm.map(chave_por_nome).fillna(fat.chave)

    venda = fat.TIPMOV.eq("V")
    peca = fat.CODCENCUS.astype(str).str.endswith(("204", "207"))
    balcao = fat.CODCENCUS.astype(str).str.endswith("204") & fat.CODCENCUS.ne(CENTRO_RAP)

    # clientes RAP: planilha de carteiras, categoria RAP no cadastro, vendedor RAP ou compra no centro RAP
    rap = set(fat.loc[fat.CODPARC.isin(rap_planilha), "cliente"])
    rap |= set(fat.loc[fat.AD_CLASSRAP.fillna("").str.upper().isin(CLASSES_RAP), "cliente"])
    rap |= set(fat.loc[fat.vendedor.isin(VEND_RAP) | fat.VEND_PREFERENCIAL.fillna("").str.upper().isin(VEND_RAP), "cliente"])
    rap |= set(fat.loc[fat.CODCENCUS.eq(CENTRO_RAP) & venda, "cliente"])
    excluir = rap | {c for c in fat.cliente.dropna() if c[:8] in RAIZ_EXCLUIR}

    # telefone
    sab = sab_base()
    tel = fones_cliente(cli, zip(sab.chave, sab.fone))
    distribuir = Distribuidor()

    def consultor(ultimo_vend, uf, cidade, fone, sem_12m=False):
        c = CONSULTORES.get(ultimo_vend or "")
        return c if c else distribuir(uf, cidade, fone, sem_12m)

    # ---------------- Repetição ----------------
    f = fat[venda & balcao & fat.cliente.notna() & (fat.cliente != "")]
    devol = fat[fat.TIPMOV.eq("D") & balcao]
    s1 = f[(f.DT >= INICIO_1S26) & (f.DT <= FIM_1S26)]
    s1_dev = devol[(devol.DT >= INICIO_1S26) & (devol.DT <= FIM_1S26)].groupby("cliente").Valor.sum()
    liq = s1.groupby("cliente").Valor.sum().sub(s1_dev, fill_value=0)
    compradores = set(liq[liq > 0].index)
    voltou = set(fat[venda & peca & (fat.DT >= CORTE)].cliente)
    publico = sorted(compradores - voltou - excluir)

    s1 = s1.merge(notas[["NUNOTA", "CODPROD", "BRUTO"]], on="NUNOTA", how="left")
    s1["categoria"] = s1.CODPROD.map(categoria).fillna("SEM CATEGORIA")
    linhas = []
    for c in publico:
        h = f[f.cliente == c].sort_values("DT")
        ult = h.iloc[-1]
        # CNPJ a atacar = o que mais comprou no 1º semestre
        principal = s1[s1.cliente == c].groupby("CODPARC").BRUTO.sum().idxmax()
        p = cli_por_parc.loc[principal]
        cats = (s1[s1.cliente == c].groupby("categoria").BRUTO.sum().sort_values(ascending=False))
        fone = tel.get(p.chave, "")
        linhas.append({
            "Cliente": p.NOMEPARC, "CNPJ/CPF": p.doc, "Cidade": p.CIDADE, "UF": p.UF,
            "Contato (e-mail)": p.EMAIL if isinstance(p.EMAIL, str) else "",
            "Telefone": norm.e164(fone),
            "Data da última compra": ult.DT.date(),
            "Valor 1º sem/2026 (R$)": round(liq[c], 2),
            "O que comprava (categorias)": ", ".join(cats.index[:5]),
            "Consultor": consultor(ult.vendedor, p.UF, p.CIDADE, fone),
        })
    rep = pd.DataFrame(linhas).sort_values("Valor 1º sem/2026 (R$)", ascending=False)

    # ---------------- Novos clientes ----------------
    ja_comprou = set(fat.loc[venda & peca, "chave"]) | set(fat.loc[venda & peca, "cliente"])
    nomes_compradores = set(fat.loc[venda & peca, "razao_norm"].dropna())
    et = sab[sab.necess.eq("Entrega Técnica") & sab.chave.ne("")]
    et = et[~et.chave.isin(ja_comprou) & ~et.chave.isin(excluir)
            & ~et.nome_norm.isin(nomes_compradores) & ~et.chave.str[:8].isin(RAIZ_EXCLUIR)]
    cad = cli.drop_duplicates("chave").set_index("chave")
    linhas = []
    for (chave, serie), g in et.sort_values("data_os").groupby(["chave", "serie"], sort=False):
        r = g.iloc[0]
        fone = tel.get(chave, "")
        nome = cad.NOMEPARC.get(chave, r.nome)
        cidade, uf = r.cidade, r.uf
        linhas.append({
            "Cliente": nome, "CNPJ/CPF": r.doc, "Cidade": cidade, "UF": uf,
            "Contato (e-mail)": r.email if isinstance(r.email, str) else "",
            "Telefone": norm.e164(fone), "Modelo da máquina": r.modelo, "Chassi": r.serie,
            "Data da compra da máquina (entrega técnica)": r.data_os.date(),
            "chave_cli": chave,
        })
    nov = pd.DataFrame(linhas).sort_values("Data da compra da máquina (entrega técnica)", ascending=False)
    # um consultor por cliente (todas as máquinas do mesmo cliente ficam com o mesmo consultor)
    dono = {}
    for r in nov.drop_duplicates("chave_cli").itertuples():
        fone = norm.so_digitos(r.Telefone)[2:]
        dono[r.chave_cli] = consultor("", r.UF, r.Cidade, fone)
    nov["Consultor"] = nov.chave_cli.map(dono)
    nov = nov.drop(columns="chave_cli")

    escrever(rep, nov)
    resumo = (pd.concat([rep.Consultor.value_counts().rename("qtd repetição"),
                         nov.drop_duplicates("CNPJ/CPF").Consultor.value_counts().rename("qtd novos")], axis=1)
              .fillna(0).astype(int).sort_index())
    return rep, nov, resumo


def sab_base():
    at = _ler("sab_atividade.csv")
    eq = _ler("sab_equipamento.csv")
    cl = _ler("sab_cliente.csv")
    k = ["id_ativ", "id_sub"]
    for d in (at, eq, cl):
        d["id_sub"] = d.id_sub.fillna("")
    s = at.merge(eq, on=k, suffixes=("", "_eq")).merge(cl, on=k)
    s = s[s.status.ne("Cancelado")]
    s["data_os"] = pd.to_datetime(s.dt_criacao)
    s["doc"] = s.cnpj.map(norm.doc)
    s["chave"] = s.doc.map(_chave_cliente)
    s["nome_norm"] = s.nome.map(norm.razao)
    s["necess"] = s.necess.str.replace(r"\s+", " ", regex=True).str.strip()
    return s


def _salvar(df, arq):
    with pd.ExcelWriter(arq, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Clientes")
        ws = w.sheets["Clientes"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for i, col in enumerate(df.columns, 1):
            larg = max([len(str(col))] + [len(str(v)) for v in df[col].head(300)])
            ws.column_dimensions[get_column_letter(i)].width = min(max(10, larg + 2), 60)


def escrever(rep, nov):
    if SAIDA.exists():
        shutil.rmtree(SAIDA)
    for nome in sorted(set(rep.Consultor) | set(nov.Consultor)):
        pasta = SAIDA / nome
        pasta.mkdir(parents=True, exist_ok=True)
        _salvar(rep[rep.Consultor == nome].drop(columns="Consultor"), pasta / f"Repeticao_{nome}.xlsx")
        _salvar(nov[nov.Consultor == nome].drop(columns="Consultor"), pasta / f"Novos_Clientes_{nome}.xlsx")


if __name__ == "__main__":
    _, _, resumo = gerar()
    print(resumo.to_string())
