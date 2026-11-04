"""PNs de alto valor por categoria do card Black Friday e modelo de Parts Question (subsídio fábrica).

Uso:  python -m campanha.blackfriday [planilha_margem.xlsx] [--top N]
Preço = "Preço novo (tabela)" da aba "Detalhe por Part Number".
Venda: se cache/raw tiver f_Fat_Clientes.csv e f_Notas.csv, só entram PNs vendidos nos últimos 12 meses;
sem a base, entram PNs com custo no Sankhya (status OK = já tiveram entrada).
"""
import argparse
import re
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RAIZ = Path(__file__).resolve().parent.parent
RAW = RAIZ / "cache" / "raw"
SAIDA = RAIZ / "saida"
PRECO = "Preço novo (tabela)"
MIN_PRECO = 1000
NAO_ESCAVADEIRA = r"EMPIL|GLP|\bGAS\b|GIRO|HIDRAULICO|VOLANTE"

CATEGORIAS = [
    ("Motor de Tração e Redutores", lambda d: d.CODGRUPOPROD.eq(339000000)),
    ("Motor de Giro e Redutores", lambda d: d.CODGRUPOPROD.eq(340000000)),
    ("Motores Diesel (R220 / HX220)", lambda d: d.CODGRUPOPROD.eq(338000000)
     & ~d.desc.str.contains(NAO_ESCAVADEIRA, regex=True)),
    ("Caçamba", lambda d: d.CODGRUPOPROD.eq(306000000)),
    ("Rompedor", lambda d: d.desc.str.contains("ROMPEDOR")),
    ("Engate Rápido", lambda d: d.desc.str.contains(r"ENGATE RAPIDO|QUICKCOUPLER", regex=True)),
    ("MCV - Comando Hidráulico", lambda d: d.CODGRUPOPROD.eq(336000000)),
    ("MCU - Módulo", lambda d: d.desc.str.contains(r"\b[CR]?MCU\b", regex=True)),
    ("Mangueiras", lambda d: d.CODGRUPOPROD.eq(335000000) & d.desc.str.contains(r"MANGUEIRA|HOSE", regex=True)),
    ("Correias", lambda d: d.desc.str.match(r"CORREIA")),
    ("Material Rodante", lambda d: d.CODGRUPOPROD.isin(
        [310000000, 315000000, 344000000, 366000000, 367000000, 368000000, 369000000, 377000000])),
]


def vendas_12m():
    if not (RAW / "f_Fat_Clientes.csv").exists():
        return None
    ler = lambda n: pd.read_csv(RAW / n, dtype=str, low_memory=False).rename(
        columns=lambda c: c.split("[")[-1].strip("]"))
    fat, notas = ler("f_Fat_Clientes.csv"), ler("f_Notas.csv")
    fat["DT"] = pd.to_datetime(fat.DTFATUR)
    fat = fat[fat.TIPMOV.eq("V") & (fat.DT >= fat.DT.max() - pd.DateOffset(months=12))]
    v = notas[notas.NUNOTA.isin(fat.NUNOTA)].assign(BRUTO=lambda x: pd.to_numeric(x.BRUTO))
    return v.groupby("CODPROD").agg(Notas_12m=("NUNOTA", "nunique"), Faturado_12m=("BRUTO", "sum"))


def selecionar(planilha, top):
    d = pd.read_excel(planilha, sheet_name="Detalhe por Part Number")
    d["desc"] = d["Descrição"].astype(str).str.upper()
    d = d[d.Status.eq("OK") & d[PRECO].ge(MIN_PRECO)]
    vend = vendas_12m()
    if vend is not None:
        d = d.join(vend, on=d.CODPROD.astype(str)).dropna(subset=["Notas_12m"])
    partes = [d[f(d)].nlargest(top, PRECO).assign(**{"Categoria Black Friday": nome})
              for nome, f in CATEGORIAS]
    cols = ["Categoria Black Friday", "Part Number", "Descrição", "Categoria", PRECO,
            "Preço atual (Sankhya)", "Custo Total"] + (["Notas_12m", "Faturado_12m"] if vend is not None else [])
    return pd.concat(partes)[cols].rename(columns={"Categoria": "Grupo Sankhya"}), vend is not None


CARTA = """To: HD Hyundai Construction Equipment – Parts Sales / Parts Marketing
From: BMC Hyundai – Parts Department (Brazil)
Subject: PARTS QUESTION – Price Support (Subsidy) Request – BMC Black Friday 2026

Dear Hyundai Parts Team,

BMC Hyundai will run a national Black Friday parts campaign from 23/11 to 30/11/2026, focused on
high-value components for the Hyundai excavator and wheel loader fleet in Brazil (R220 / HX220 and others).

To offer competitive prices against aftermarket and competitor dealers, we kindly request special
price support (subsidy) for the part numbers listed below, all with sales history at BMC.

Request details:
  • Campaign period: 23/11/2026 – 30/11/2026 (orders placed up to 30/11/2026)
  • Requested support: [__]% over current dealer purchase price (or special net price per PN)
  • Estimated campaign quantity: see column "Est. Qty"
  • Order type: stock order with campaign reference "BMC BLACK FRIDAY 2026"
  • Expected result: [__] units / BRL [__] in additional parts revenue

Please confirm the approved price / subsidy per PN, validity and any order conditions by [dd/mm/2026].

Best regards,
[Name] – Parts Manager | BMC Hyundai | [phone] | [e-mail]"""


def escrever(lista, com_venda, arq):
    pq = lista.assign(**{"Machine Model": "", "Est. Qty": "", "Requested Subsidy (%)": "",
                         "Target Net Price (BRL)": "", "HCE Approved Price": "", "HCE Comments": ""})
    pq = pq[["Categoria Black Friday", "Part Number", "Descrição", "Machine Model", PRECO, "Est. Qty",
             "Requested Subsidy (%)", "Target Net Price (BRL)", "HCE Approved Price", "HCE Comments"]]
    pq.columns = ["Category", "Part Number", "Description", "Machine Model", "Current Price BMC (BRL)",
                  "Est. Qty", "Requested Subsidy (%)", "Target Net Price (BRL)", "HCE Approved Price",
                  "HCE Comments"]
    resumo = lista.groupby("Categoria Black Friday", sort=False).agg(
        PNs=("Part Number", "count"), Maior_preco=(PRECO, "max"), Menor_preco=(PRECO, "min"))
    linhas_carta = CARTA.count("\n") + 2

    with pd.ExcelWriter(arq, engine="openpyxl") as w:
        lista.to_excel(w, sheet_name="PNs Alto Valor", index=False)
        resumo.to_excel(w, sheet_name="Resumo")
        pq.to_excel(w, sheet_name="Parts Question", index=False, startrow=linhas_carta)
        ws = w.sheets["Parts Question"]
        for i, linha in enumerate(CARTA.split("\n"), 1):
            ws.cell(i, 1, linha).font = Font(bold=i <= 3)
        if not com_venda:
            ws.cell(linhas_carta, 1, "Obs. interna: validar histórico de venda antes do envio.").font = \
                Font(italic=True, color="C00000")
        cab = PatternFill("solid", fgColor="002C5F")
        for nome, ws in w.sheets.items():
            topo = linhas_carta + 1 if nome == "Parts Question" else 1
            for c in ws[topo]:
                c.font, c.fill = Font(bold=True, color="FFFFFF"), cab
                c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.freeze_panes = ws.cell(topo + 1, 1)
            ws.auto_filter.ref = f"A{topo}:{get_column_letter(ws.max_column)}{ws.max_row}"
            for col in ws.iter_cols(min_row=topo):
                larg = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(larg + 2, 12), 45)
                if re.search(r"Pre[çc]|Price|Custo|Faturado", str(col[0].value)):
                    for c in col[1:]:
                        c.number_format = '"R$" #,##0.00'


def escrever_anexo(lista, arq):
    df = lista[["Part Number", "Descrição", "Categoria Black Friday"]]
    df.columns = ["Part Number", "Description", "Category"]
    with pd.ExcelWriter(arq, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Parts List", index=False)
        ws = w.sheets["Parts List"]
        for c in ws[1]:
            c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="002C5F")
        for col, larg in zip("ABC", (20, 45, 32)):
            ws.column_dimensions[col].width = larg
        ws.freeze_panes, ws.auto_filter.ref = "A2", ws.dimensions


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("planilha", nargs="?", default=RAIZ / "dados" / "Reestruturacao_Margem_2026_FINAL.xlsx")
    ap.add_argument("--top", type=int, default=15)
    ap.add_argument("--saida", default=SAIDA / "BlackFriday_PNs_Alto_Valor_Parts_Question.xlsx")
    a = ap.parse_args()
    lista, com_venda = selecionar(a.planilha, a.top)
    Path(a.saida).parent.mkdir(parents=True, exist_ok=True)
    escrever(lista, com_venda, a.saida)
    anexo = Path(a.saida).with_name("BlackFriday_Parts_Question_Anexo.xlsx")
    escrever_anexo(lista, anexo)
    print(f"{len(lista)} PNs | venda 12m: {'sim' if com_venda else 'base ausente (status OK)'} -> {a.saida}")


if __name__ == "__main__":
    main()
