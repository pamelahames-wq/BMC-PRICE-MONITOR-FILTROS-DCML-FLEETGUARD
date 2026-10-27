"""Duas planilhas só com os clientes sem nenhuma exceção (Repetição e Novos), para importar no HubSpot.

Uso:  python -m campanha.listas_finais <pasta_planilhas> <pasta_bases> <pasta_saida>
Sai qualquer cliente com divergência na base, alerta da auditoria (carteira, revenda, homônimo,
mesmo telefone de cliente de Repetição) ou que não pôde ser confirmado. Fica de fora quem só
entrou na validação (filiais), porque não constava da planilha do consultor. Contato errado
(telefone/e-mail interno da BMC) não exclui o cliente: o campo é deixado em branco.
"""
import sys
from pathlib import Path

import pandas as pd

from . import norm
from .auditar import auditar, ler, salvar
from .consolidar import REGIONAL, telefones_internos
from .cruzar import bases, cruzar

ITENS_AUDITORIA = {"03", "04", "11", "12", "24", "25"}


def listas(pasta, pasta_bases):
    rep, nov = ler(pasta)
    cli, fat, rap, servico, sab = bases(pasta_bases)
    aud, _, _ = auditar(rep.copy(), nov.copy())
    ach, _, _, _, _ = cruzar(rep.copy(), nov.copy(), cli, fat, rap, servico, sab)
    internos = telefones_internos(rep, nov, cli, sab)
    fora = set(ach["CNPJ/CPF"]) | set(aud.loc[aud.Item.isin(ITENS_AUDITORIA), "CNPJ/CPF"])

    def contato(df):
        df["Telefone"] = [t if t and t not in internos else "" for t in df.tel]
        df["E-mail"] = ["" if "bmchyundai" in e else e for e in df.email]
        return df

    # Repetição: confirmada na base e sem nenhum alerta
    r = contato(rep[~rep.doc.isin(fora)].copy())
    r.insert(0, "Regional", r.Consultor.map(REGIONAL))
    r = r[["Regional", "Consultor", "Cliente", "doc", "Cidade", "UF", "E-mail", "Telefone",
           "Data da última compra", "Valor 1º sem/2026 (R$)", "O que comprava (categorias)"]]
    r = r.rename(columns={"doc": "CNPJ/CPF"}).sort_values(["Regional", "Consultor", "Cliente"])

    # Novos: além de sem alerta, só máquina entregue em 2026, sem atendimento no SAB antes de 2026
    # (sem sinal de máquina antiga) e sem NF de peça em nenhum CNPJ da mesma raiz
    raiz_peca = set(fat.loc[fat.venda & fat.peca, "doc"].dropna().map(norm.raiz))
    sab_antes = set(sab.loc[sab.data < "2026-01-01", "doc"])
    et_min = nov.groupby("doc").et.min()
    ok = nov[~nov.doc.isin(fora) & ~nov.doc.isin(sab_antes) & nov.doc.map(et_min).ge("2026-01-01")
             & ~nov.doc.map(norm.raiz).isin(raiz_peca)]
    linhas = []
    for d, g in contato(ok.copy()).groupby("doc", sort=False):
        x = g.iloc[0]
        linhas.append({"Regional": REGIONAL.get(x.Consultor), "Consultor": x.Consultor, "Cliente": x.Cliente,
                       "CNPJ/CPF": d, "Cidade": x.Cidade, "UF": x.UF, "E-mail": x["E-mail"], "Telefone": x.Telefone,
                       "Qtd máquinas": g.Chassi.nunique(), "Modelos": ", ".join(sorted(set(g["Modelo da máquina"]))),
                       "Chassis": ", ".join(sorted(set(g.Chassi))),
                       "Entrega técnica (compra da máquina)": f"{g.et.max():%Y-%m-%d}"})
    n = pd.DataFrame(linhas).sort_values(["Regional", "Consultor", "Cliente"])
    return r, n


if __name__ == "__main__":
    r, n = listas(sys.argv[1], sys.argv[2])
    saida = Path(sys.argv[3])
    saida.mkdir(parents=True, exist_ok=True)
    salvar(saida / "Clientes_Repeticao_GPD_2sem2026.xlsx", Repeticao=r)
    salvar(saida / "Clientes_Novos_GPD_2sem2026.xlsx", Novos_Clientes=n)
    print("repetição:", len(r), r["CNPJ/CPF"].is_unique, " novos:", len(n))
    print(pd.concat([r.Consultor.value_counts().rename("rep"), n.Consultor.value_counts().rename("nov")], axis=1).fillna(0).astype(int).to_string())
