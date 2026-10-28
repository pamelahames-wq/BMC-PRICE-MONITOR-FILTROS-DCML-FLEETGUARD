"""Gera a planilha da Reestruturação RAP (com fórmulas) e o JSON que alimenta a apresentação.

Uso:  python -m rap.saidas   →  saida/rap/Reestruturacao_RAP_2S2026.xlsx + deck.json
Depois: node rap/deck.js     →  saida/rap/Reestruturacao_RAP_2S2026.pptx
"""
import json

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import analise as A

SAIDA = A.RAIZ / "saida" / "rap"
ARQ = SAIDA / "Reestruturacao_RAP_2S2026.xlsx"

NAVY, AMARELO = "12285A", "F5B800"
F = "Arial"
H_FONT = Font(name=F, bold=True, color="FFFFFF")
H_FILL = PatternFill("solid", fgColor=NAVY)
INPUT = Font(name=F, color="0000FF")
LINK = Font(name=F, color="008000")
BASE = Font(name=F)
BOLD = Font(name=F, bold=True)
TITULO = Font(name=F, bold=True, size=14, color=NAVY)
FINA = Side(style="thin", color="BFBFBF")
RS = '"R$" #,##0.00;-"R$" #,##0.00;"-"'
RS0 = '"R$" #,##0;-"R$" #,##0;"-"'
PCT = '0.0%;-0.0%;"-"'
INT = '#,##0;-#,##0;"-"'
DATA = "DD/MM/YYYY"

# endereços fixos na aba Premissas (usados pelas fórmulas das outras abas)
P = "Premissas!"
P_INI, P_FIM, P_FIMSEM, P_DIAS, P_DIASSEM = "$B$4", "$B$5", "$B$6", "$B$7", "$B$8"
P_LIM, P_BALCAO, P_IMP, P_IMPB, P_PISO = "$B$9", "$B$10", "$B$11", "$B$12", "$B$13"
CAT_ROWS = {"Standard": 15, "Ouro": 16, "Diamante": 17}
CAT_RNG, MIN_RNG, REB_RNG = "$A$15:$A$17", "$B$15:$B$17", "$E$15:$E$17"


def _cab(ws, linha, nomes, larg=None):
    for i, n in enumerate(nomes, 1):
        c = ws.cell(linha, i, n)
        c.font, c.fill = H_FONT, H_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[linha].height = 32
    for i, w in enumerate(larg or [], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _cod(v):
    return int(v) if str(v).isdigit() else v


def _fmt(ws, linha, col, valor, fmt=None, font=BASE):
    c = ws.cell(linha, col, valor)
    c.font = font
    if fmt:
        c.number_format = fmt
    return c


def _filtro(ws, linha_cab, ult_linha, ncol):
    ws.freeze_panes = ws.cell(linha_cab + 1, 1)
    ws.auto_filter.ref = f"A{linha_cab}:{get_column_letter(ncol)}{max(ult_linha, linha_cab + 1)}"


def premissas(wb, r):
    ws = wb.active
    ws.title = "Premissas"
    ws["A1"] = "Reestruturação RAP — premissas e fontes"
    ws["A1"].font = TITULO
    ws["A2"] = "Células em azul são parâmetros: alterar aqui recalcula classificação, estudo de preço e simulações."
    ws["A2"].font = Font(name=F, italic=True)
    linhas = [
        ("Início do semestre analisado", pd.Timestamp(A.INICIO), DATA, "2º semestre/2026"),
        ("Fim do período (última NF na base)", pd.Timestamp(A.FIM), DATA,
         "Semestre ainda aberto: a base do Sankhya vai até 16/10/2026"),
        ("Fim do semestre", pd.Timestamp(A.FIM_SEMESTRE), DATA, ""),
        ("Dias decorridos", "=B5-B4+1", INT, ""),
        ("Dias do semestre", "=B6-B4+1", INT, ""),
        ("Limite para REBAIXAR (% do mínimo)", 0.7, PCT, "Abaixo disso rebaixa; entre isso e 100% = INCENTIVAR"),
        ("Desconto padrão Balcão", A.DESC_BALCAO, PCT, "Sobre o preço novo de tabela"),
        ("Impostos na NF RAP", round(r["imp_rap"], 4), PCT,
         "IMPOSTOS ÷ BRUTO das vendas RAP no período (f_Notas). A Reestruturação de Margem usa 13,3%"),
        ("Impostos na NF Balcão", round(r["imp_balcao"], 4), PCT, "IMPOSTOS ÷ BRUTO das vendas Balcão no período"),
        ("Piso de margem por grupo na venda RAP", A.PISO_MARGEM, PCT,
         "Proposta: nenhum grupo abaixo disso; define o desconto sugerido na aba 'Desconto por grupo'"),
    ]
    for i, (rot, val, fmt, nota) in enumerate(linhas, 4):
        ws.cell(i, 1, rot).font = BASE
        _fmt(ws, i, 2, val, fmt, BASE if str(val).startswith("=") else INPUT)
        ws.cell(i, 3, nota).font = Font(name=F, italic=True, color="595959")

    _cab(ws, 14, ["Categoria", "Mínimo semestral (R$)", "Desconto mín. da faixa",
                  "Desconto máx. da faixa", "Rebaixa para"])
    for cat, lin in CAT_ROWS.items():
        ws.cell(lin, 1, cat).font = BOLD
        _fmt(ws, lin, 2, A.MINIMO[cat], RS0, INPUT)
        _fmt(ws, lin, 3, A.FAIXA[cat][0], PCT, INPUT)
        _fmt(ws, lin, 4, A.FAIXA[cat][1], PCT, INPUT)
        _fmt(ws, lin, 5, A.REBAIXA[cat], None, INPUT)

    ws["A20"] = "Regras aplicadas"
    ws["A20"].font = BOLD
    regras = [
        "RAP = VEND_PREFERENCIAL (d_Clientes) ∈ {ELISASILVA, SLUNA, PLOUREIRO, FALEXANDRE} e AD_CLASSRAP ∈ "
        "{Standard, Ouro, Diamante}. O vendedor 'ESILVA' do pedido está cadastrado no Sankhya como ELISASILVA.",
        "Categoria RAP preenchida sem um desses vendedores = cadastro inconsistente (tratado como NÃO-RAP). "
        "Vendedor RAP com categoria Auto Atendimento/Inativo/vazia também listado lá (não dá para aplicar a régua).",
        "Faturamento = NFs de venda (TIPMOV V) menos devoluções (D) do CODPARC, só centros de peças "
        "(…204 Balcão e 1030207 RAP); valor da NF (f_Fat_Clientes.Valor = soma de f_Notas.BRUTO).",
        "Baldes: R$ 0 → DESHOMOLOGAR; < 70% do mínimo → REBAIXAR (Standard vai para DESHOMOLOGAR); "
        "70%–100% → INCENTIVAR; ≥ 100% → MANTER.",
        "Projeção do semestre = faturamento × dias do semestre ÷ dias decorridos (só referência; a classificação "
        "usa o realizado, como pedido).",
        "Margem = (valor da NF − impostos − custo) ÷ valor da NF, mesma convenção da Reestruturação de Margem "
        "2026, mas com o imposto real das NFs (RAP ≈ 20,5%, Balcão ≈ 13,7%) em vez dos 13,3% da planilha. "
        "Custo = CUSTO da NF × qtd; se zerado, último custo do histórico f_CustoUnitarioProd até a data da NF.",
        "Desconto por grupo de peça = tabela atual da Reestruturação de Margem 2026 (aba Resumo por grupo). As "
        "faixas por categoria acima são a régua informada para a reestruturação.",
        "Preço de tabela = d_Precos.VLRVENDA, que já é o 'Preço novo (tabela)' da Reestruturação de Margem 2026 "
        "(conferido: 995 de 995 itens iguais).",
    ]
    for i, t in enumerate(regras, 21):
        ws.cell(i, 1, "• " + t).font = BASE

    ws["A30"] = "Passo 0 — schema real usado (exportações do Sankhya no OneDrive, pasta Simulador de Margem/data)"
    ws["A30"].font = BOLD
    _cab(ws, 31, ["Necessidade", "Tabela padrão Sankhya", "Exportação usada", "Campos"])
    schema = [
        ("Cadastro de clientes", "TGFPAR / d_Clientes", "d_Clientes.csv",
         "CODPARC, NOMEPARC, CODIGO_MATRIZ (CNPJ 14 dígitos), CNPJ_CPF (só raiz de 8), VEND_PREFERENCIAL "
         "(apelido), AD_CLASSRAP (categoria RAP)"),
        ("Cabeçalho da NF", "TGFCAB", "f_Fat_Clientes.csv",
         "NUNOTA, CODVEND, CODPARC, DTFATUR, TIPMOV (V/D), CODCENCUS, ZYDON, Valor"),
        ("Itens da NF", "TGFITE", "f_Notas.csv",
         "NUNOTA, DTFATUR, CODTIPOPER, TIPMOV, CODVEND, CODPROD, QTDNEG, BRUTO (total da linha), IMPOSTOS, "
         "CUSTO (unitário), CODCENCUS"),
        ("Produto", "TGFPRO", "d_Produtos.csv",
         "CODPROD, DESCRPROD, REFERENCIA (PN), Grupo Nivel 1/3 (sem CODGRUPOPROD: mapeado pelo nome do grupo "
         "na tabela de margens/descontos)"),
        ("Preço de tabela", "TGFEXC / TGFTAB", "d_Precos.csv", "CODPROD, VLRVENDA"),
        ("Custo", "TGFCUS", "f_CustoUnitarioProd.csv", "codProduto, dtCusto, vlrCustoUnitario, vlrAtual"),
    ]
    for i, linha in enumerate(schema, 32):
        for j, v in enumerate(linha, 1):
            ws.cell(i, j, v).font = BASE
            ws.cell(i, j).alignment = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABCDE", [42, 24, 26, 90, 16]):
        ws.column_dimensions[col].width = w


def classificacao(wb, c):
    ws = wb.create_sheet("RAPs - classificação")
    cols = ["CODPARC", "CNPJ", "Nome", "Consultor (VEND_PREFERENCIAL)", "Categoria atual",
            "Faturamento semestre (R$)", "Mínimo da categoria (R$)", "% do mínimo", "Balde",
            "Categoria sugerida", "Quanto falta (R$)", "Observação", "Faturamento 1º sem/2026 (R$)",
            "Projeção semestre (R$)", "% do mínimo na projeção"]
    _cab(ws, 1, cols, [10, 17, 40, 22, 12, 16, 14, 10, 15, 14, 14, 34, 16, 16, 12])
    for i, x in enumerate(c.itertuples(), 2):
        obs = x.motivo
        if x.balde == "DESHOMOLOGAR" and x.fat_1s > 0:
            obs = (obs + "; " if obs else "") + "comprou no 1º sem/2026"
        _fmt(ws, i, 1, int(x.CODPARC))
        _fmt(ws, i, 2, x.CNPJ)
        _fmt(ws, i, 3, x.NOMEPARC)
        _fmt(ws, i, 4, x.consultor)
        _fmt(ws, i, 5, x.categoria)
        _fmt(ws, i, 6, x.fat, RS, INPUT)
        _fmt(ws, i, 7, f"=INDEX({P}{MIN_RNG},MATCH(E{i},{P}{CAT_RNG},0))", RS0, LINK)
        _fmt(ws, i, 8, f"=IF(G{i}>0,F{i}/G{i},0)", PCT)
        _fmt(ws, i, 9, f'=IF(F{i}<=0,"DESHOMOLOGAR",IF(H{i}<{P}{P_LIM},IF(E{i}="Standard",'
                       f'"DESHOMOLOGAR","REBAIXAR"),IF(H{i}<1,"INCENTIVAR","MANTER")))')
        _fmt(ws, i, 10, f'=IF(I{i}="DESHOMOLOGAR","Encerrar",IF(I{i}="REBAIXAR",'
                        f'INDEX({P}{REB_RNG},MATCH(E{i},{P}{CAT_RNG},0)),E{i}))')
        _fmt(ws, i, 11, f'=IF(I{i}="INCENTIVAR",G{i}-F{i},0)', RS)
        _fmt(ws, i, 12, obs)
        _fmt(ws, i, 13, x.fat_1s, RS, INPUT)
        _fmt(ws, i, 14, f"=F{i}*{P}{P_DIASSEM}/{P}{P_DIAS}", RS)
        _fmt(ws, i, 15, f"=IF(G{i}>0,N{i}/G{i},0)", PCT)
    _filtro(ws, 1, len(c) + 1, len(cols))
    return len(c) + 1


def resumos(wb, ult, consultores):
    ref = "'RAPs - classificação'!"
    rng = lambda col: f"{ref}${col}$2:${col}${ult}"  # noqa: E731

    ws = wb.create_sheet("Resumo por consultor")
    ws["A1"] = f"Nº de RAPs e faturamento por consultor × balde — {A.INICIO[8:]}/{A.INICIO[5:7]} a 16/10/2026"
    ws["A1"].font = TITULO
    cab = ["Consultor"]
    for b in A.BALDES:
        cab += [f"{b} (nº)", f"{b} (R$)"]
    cab += ["Total (nº)", "Total (R$)"]
    _cab(ws, 3, cab, [24] + [13, 15] * 5)
    for i, cons in enumerate(consultores + ["Total"], 4):
        ws.cell(i, 1, cons).font = BOLD if cons == "Total" else BASE
        for j, b in enumerate(A.BALDES):
            cn, cv = 2 + 2 * j, 3 + 2 * j
            if cons == "Total":
                for cc, fmt in ((cn, INT), (cv, RS0)):
                    L = get_column_letter(cc)
                    _fmt(ws, i, cc, f"=SUM({L}4:{L}{i - 1})", fmt, BOLD)
            else:
                _fmt(ws, i, cn, f'=COUNTIFS({rng("D")},$A{i},{rng("I")},"{b}")', INT)
                _fmt(ws, i, cv, f'=SUMIFS({rng("F")},{rng("D")},$A{i},{rng("I")},"{b}")', RS0)
        tn = "+".join(f"{get_column_letter(2 + 2 * j)}{i}" for j in range(4))
        tv = "+".join(f"{get_column_letter(3 + 2 * j)}{i}" for j in range(4))
        _fmt(ws, i, 10, f"={tn}", INT, BOLD)
        _fmt(ws, i, 11, f"={tv}", RS0, BOLD)

    ws = wb.create_sheet("Resumo por balde")
    ws["A1"] = "Contagem e faturamento por balde"
    ws["A1"].font = TITULO
    _cab(ws, 3, ["Balde", "Nº de RAPs", "% das RAPs", "Faturamento semestre (R$)", "Quanto falta (R$)", "Ação"],
         [18, 12, 12, 20, 18, 60])
    acoes = {"DESHOMOLOGAR": "E-mail de encerramento (após aprovação); revisar antes quem comprou no 1º sem",
             "REBAIXAR": "Comunicar nova categoria e novo desconto",
             "INCENTIVAR": "Lista de contato do consultor: mostrar quanto falta para manter a categoria",
             "MANTER": "Sem mudança"}
    n = len(A.BALDES)
    for i, b in enumerate(A.BALDES, 4):
        ws.cell(i, 1, b).font = BOLD
        _fmt(ws, i, 2, f'=COUNTIF({rng("I")},A{i})', INT)
        _fmt(ws, i, 3, f"=B{i}/B${4 + n}", PCT)
        _fmt(ws, i, 4, f'=SUMIFS({rng("F")},{rng("I")},A{i})', RS0)
        _fmt(ws, i, 5, f'=SUMIFS({rng("K")},{rng("I")},A{i})', RS0)
        _fmt(ws, i, 6, acoes[b])
    t = 4 + n
    ws.cell(t, 1, "Total").font = BOLD
    for col, fmt in (("B", INT), ("C", PCT), ("D", RS0), ("E", RS0)):
        _fmt(ws, t, " BCDE".index(col) + 1, f"=SUM({col}4:{col}{t - 1})", fmt, BOLD)

    r0 = t + 3
    ws.cell(r0 - 1, 1, "Nº de RAPs por categoria atual × balde").font = BOLD
    _cab(ws, r0, ["Categoria atual"] + A.BALDES + ["Total"])
    for i, cat in enumerate(A.CATEGORIAS, r0 + 1):
        ws.cell(i, 1, cat).font = BOLD
        for j, b in enumerate(A.BALDES, 2):
            _fmt(ws, i, j, f'=COUNTIFS({rng("E")},$A{i},{rng("I")},"{b}")', INT)
        _fmt(ws, i, 6, f"=SUM(B{i}:E{i})", INT, BOLD)


def inconsistente(wb, d):
    ws = wb.create_sheet("Cadastro inconsistente")
    cols = ["CODPARC", "CNPJ", "Nome", "VEND_PREFERENCIAL", "AD_CLASSRAP", "Cidade", "UF", "Problema",
            "Faturamento semestre (R$)"]
    _cab(ws, 1, cols, [10, 17, 40, 20, 16, 22, 6, 58, 16])
    d = d.sort_values(["Problema", "fat"], ascending=[True, False])
    for i, x in enumerate(d.itertuples(), 2):
        for j, v in enumerate([int(x.CODPARC), x.CNPJ, x.NOMEPARC, x.VEND_PREFERENCIAL, x.AD_CLASSRAP,
                               x.CIDADE, x.UF, x.Problema], 1):
            _fmt(ws, i, j, "" if pd.isna(v) else v)
        _fmt(ws, i, 9, x.fat, RS)
    _filtro(ws, 1, len(d) + 1, len(cols))


def estudo_preco(wb, cesta):
    ws = wb.create_sheet("Estudo de preço")
    ws["A1"] = ("Quanto sai a peça em cada canal — preço novo de tabela (d_Precos) menos o desconto do canal. RAP usa o "
                "desconto atual do grupo da peça (Reestruturação de Margem 2026).")
    ws["A1"].font = BOLD
    ws["A2"] = ("Margem = (preço − impostos − custo) ÷ preço, com o imposto real de cada canal (Premissas B11/B12). "
                "Diamante sugerido = desconto do grupo que respeita o piso de margem (aba 'Desconto por grupo').")
    ws["A2"].font = Font(name=F, italic=True)
    cab = ["PN", "CODPROD", "Descrição", "CODGRUPOPROD", "Grupo", "Custo atual (R$)", "Preço tabela (R$)",
           "Balcão 15%\nR$", "Balcão\nmargem",
           "Standard\ndesc. grupo", "Standard\nR$", "Standard\nmargem",
           "Ouro\ndesc. grupo", "Ouro\nR$", "Ouro\nmargem",
           "Diamante\ndesc. grupo", "Diamante\nR$", "Diamante\nmargem",
           "Diamante\nsugerido", "Diamante sug.\nR$", "Diamante sug.\nmargem",
           "Teto régua Diamante\nR$", "Teto régua Diamante\nmargem",
           "Faturado 2026 (R$)", "Qtd 2026"]
    _cab(ws, 4, cab, [15, 9, 30, 12, 24, 12, 12] + [12, 9] + [9, 12, 9] * 4 + [12, 10] + [14, 9])
    ws.row_dimensions[4].height = 46
    imp, impb = f"{P}{P_IMP}", f"{P}{P_IMPB}"
    grp = "'Desconto por grupo'!"
    mg = lambda pc: f'=IF(N($F{{i}})<=0,"s/ custo",1-{{imp}}-$F{{i}}/{pc}{{i}})'  # noqa: E731
    for i, x in enumerate(cesta.itertuples(), 5):
        for j, v in enumerate([x.REFERENCIA, int(x.CODPROD), x.DESCRPROD, _cod(x.CODGRUPOPROD), x.grupo], 1):
            _fmt(ws, i, j, v)
        _fmt(ws, i, 6, round(x.custo, 4) if pd.notna(x.custo) and x.custo > 0 else None, RS, INPUT)
        _fmt(ws, i, 7, round(x.preco, 2), RS, INPUT)
        _fmt(ws, i, 8, f"=$G{i}*(1-{P}{P_BALCAO})", RS)
        _fmt(ws, i, 9, f'=IF(N($F{i})<=0,"s/ custo",1-{impb}-$F{i}/H{i})', PCT)
        for k, (d, col_sug) in enumerate([(x.dS, None), (x.dO, None), (x.dD, None), (None, "Q")]):
            cd, cp, cm = 10 + 3 * k, 11 + 3 * k, 12 + 3 * k
            if col_sug:
                _fmt(ws, i, cd, f"=INDEX({grp}${col_sug}$5:${col_sug}$200,MATCH($D{i},{grp}$A$5:$A$200,0))", PCT, LINK)
            else:
                _fmt(ws, i, cd, d, PCT, INPUT)
            L, D_ = get_column_letter(cp), get_column_letter(cd)
            _fmt(ws, i, cp, f"=$G{i}*(1-{D_}{i})", RS)
            _fmt(ws, i, cm, mg(L).format(i=i, imp=imp), PCT)
        _fmt(ws, i, 22, f"=$G{i}*(1-{P}$D${CAT_ROWS['Diamante']})", RS)
        _fmt(ws, i, 23, mg("V").format(i=i, imp=imp), PCT)
        _fmt(ws, i, 24, round(x.fat_2026, 2), RS0)
        _fmt(ws, i, 25, x.qtd_2026, INT)
    ws.freeze_panes = "F5"
    nota = len(cesta) + 7
    ws.cell(nota, 1, "Cesta: os itens mais faturados em 2026 (Balcão + RAP) em cada grupo pedido; inclui 2 "
                     "bombas hidráulicas (303). Custo atual = último custo de f_CustoUnitarioProd.").font = \
        Font(name=F, italic=True)
    ws.cell(nota + 1, 1, "O desconto sugerido é por grupo (média do que as RAPs compraram): uma peça cara dentro "
                         "de um grupo saudável ainda pode ficar negativa, como a bomba principal.").font = \
        Font(name=F, italic=True)


def base_itens(wb, itens, precos):
    ws = wb.create_sheet("Base itens RAP")
    cols = ["NUNOTA", "Data", "CODPARC", "Cliente", "Categoria", "CODPROD", "PN", "Descrição", "CODGRUPOPROD",
            "Grupo", "TIPMOV", "Qtd", "Faturamento NF (R$)", "Impostos (R$)", "Receita líquida (R$)",
            "Custo unitário (R$)", "Origem do custo", "Custo total (R$)", "Preço tabela × qtd (R$)"]
    _cab(ws, 1, cols, [10, 11, 9, 32, 10, 9, 16, 30, 13, 26, 7, 7, 14, 13, 14, 13, 11, 14, 14])
    itens = itens.sort_values(["categoria", "DT", "NUNOTA"])
    for i, x in enumerate(itens.itertuples(), 2):
        vals = [int(x.NUNOTA), x.DT.to_pydatetime(), int(x.CODPARC), x.NOMEPARC, x.categoria, int(x.CODPROD),
                x.REFERENCIA, x.DESCRPROD, _cod(x.CODGRUPOPROD), x.grupo, x.TIPMOV, x.QTDNEG]
        for j, v in enumerate(vals, 1):
            _fmt(ws, i, j, "" if (not isinstance(v, (int, float)) and pd.isna(v)) else v,
                 DATA if j == 2 else None)
        _fmt(ws, i, 13, round(x.fat_nf, 2), RS)
        _fmt(ws, i, 14, round(x.impostos, 2), RS)
        _fmt(ws, i, 15, f"=M{i}-N{i}", RS)
        _fmt(ws, i, 16, None if pd.isna(x.custo_unit) else round(x.custo_unit, 4), RS)
        _fmt(ws, i, 17, x.origem_custo)
        _fmt(ws, i, 18, f'=IF(Q{i}="s/ custo",0,P{i}*L{i}*IF(K{i}="D",-1,1))', RS)
        p = precos.get(x.CODPROD)
        _fmt(ws, i, 19, round(p * x.QTDNEG, 2) if (p and x.TIPMOV == "V") else None, RS)
    _filtro(ws, 1, len(itens) + 1, len(cols))
    return len(itens) + 1


def margens(wb, ult, mcg):
    b = "'Base itens RAP'!"
    rng = lambda col: f"{b}${col}$2:${col}${ult}"  # noqa: E731

    ws = wb.create_sheet("Margem por categoria")
    ws["A1"] = "Margem agregada das RAPs por categoria — NFs de 01/07 a 16/10/2026 (vendas − devoluções)"
    ws["A1"].font = TITULO
    cols = ["Categoria", "Faturamento NF (R$)", "Impostos (R$)", "Receita líquida (R$)", "Custo (R$)",
            "Margem (R$)", "Margem %", "% do faturamento RAP", "Linhas s/ custo"]
    _cab(ws, 3, cols, [14, 18, 16, 18, 16, 16, 11, 14, 12])
    for i, cat in enumerate(A.CATEGORIAS, 4):
        ws.cell(i, 1, cat).font = BOLD
        _fmt(ws, i, 2, f'=SUMIFS({rng("M")},{rng("E")},$A{i})', RS0)
        _fmt(ws, i, 3, f'=SUMIFS({rng("N")},{rng("E")},$A{i})', RS0)
        _fmt(ws, i, 4, f"=B{i}-C{i}", RS0)
        _fmt(ws, i, 5, f'=SUMIFS({rng("R")},{rng("E")},$A{i})', RS0)
        _fmt(ws, i, 6, f"=D{i}-E{i}", RS0)
        _fmt(ws, i, 7, f'=IF(B{i}=0,"-",F{i}/B{i})', PCT)
        _fmt(ws, i, 8, f"=B{i}/$B$7", PCT)
        _fmt(ws, i, 9, f'=COUNTIFS({rng("E")},$A{i},{rng("Q")},"s/ custo")', INT)
    ws.cell(7, 1, "Total").font = BOLD
    for col in "BCDEF":
        _fmt(ws, 7, " ABCDEF".index(col), f"=SUM({col}4:{col}6)", RS0, BOLD)
    _fmt(ws, 7, 7, "=F7/B7", PCT, BOLD)
    _fmt(ws, 7, 8, "=SUM(H4:H6)", PCT, BOLD)
    _fmt(ws, 7, 9, "=SUM(I4:I6)", INT, BOLD)
    ws["A9"] = ("Margem = (NF − impostos − custo) ÷ NF, convenção da Reestruturação de Margem. Custo zerado na NF foi completado pelo histórico de "
                "custo (f_CustoUnitarioProd); nenhuma linha ficou sem custo.")
    ws["A9"].font = Font(name=F, italic=True)

    ws = wb.create_sheet("Margem por categoria x grupo")
    cols = ["Categoria", "CODGRUPOPROD", "Grupo", "Faturamento NF (R$)", "Receita líquida (R$)", "Custo (R$)",
            "Margem (R$)", "Margem %", "% do faturamento da categoria"]
    _cab(ws, 1, cols, [12, 13, 44, 17, 17, 15, 15, 10, 14])
    tot_row = {cat: f"'Margem por categoria'!$B${4 + k}" for k, cat in enumerate(A.CATEGORIAS)}
    ordem = {c: k for k, c in enumerate(A.CATEGORIAS)}
    mcg = mcg.sort_values(["categoria", "fat_nf"], key=lambda s: s.map(ordem) if s.name == "categoria" else -s)
    for i, x in enumerate(mcg.itertuples(), 2):
        _fmt(ws, i, 1, x.categoria)
        _fmt(ws, i, 2, _cod(x.CODGRUPOPROD))
        _fmt(ws, i, 3, x.grupo)
        crit = f'{rng("E")},$A{i},{rng("I")},$B{i},{rng("J")},$C{i}'
        _fmt(ws, i, 4, f"=SUMIFS({rng('M')},{crit})", RS0)
        _fmt(ws, i, 5, f"=D{i}-SUMIFS({rng('N')},{crit})", RS0)
        _fmt(ws, i, 6, f"=SUMIFS({rng('R')},{crit})", RS0)
        _fmt(ws, i, 7, f"=E{i}-F{i}", RS0)
        _fmt(ws, i, 8, f'=IF(D{i}=0,"-",G{i}/D{i})', PCT)
        _fmt(ws, i, 9, f"=D{i}/{tot_row[x.categoria]}", PCT)
    _filtro(ws, 1, len(mcg) + 1, len(cols))


def desconto_grupo(wb, g):
    ws = wb.create_sheet("Desconto por grupo")
    ws["A1"] = "Desconto RAP por grupo de peça — tabela atual × margem com o imposto real das NFs RAP × sugerido"
    ws["A1"].font = TITULO
    ws["A2"] = ("Margem = (preço − impostos − custo) ÷ preço. A Reestruturação de Margem calculou as colunas 'Margem RAP' "
                "com 13,3% de impostos; nas NFs RAP o imposto é ~20,5%. Desconto sugerido = menor entre o atual e o "
                "maior desconto que mantém a margem no piso (Premissas B13). Filtre 'Muda?' = SIM.")
    ws["A2"].font = Font(name=F, italic=True)
    cols = ["CODGRUPOPROD", "Grupo", "Itens", "Faturamento RAP 2º sem (R$)", "Fonte custo/imposto",
            "Imposto ÷ NF", "Custo ÷ tabela",
            "Desc STD atual", "Desc OURO atual", "Desc DIAMANTE atual",
            "Margem DIAMANTE na planilha (13,3% imp.)",
            "Margem real STD", "Margem real OURO", "Margem real DIAMANTE",
            "Desc STD sugerido", "Desc OURO sugerido", "Desc DIAMANTE sugerido",
            "Margem DIAMANTE no sugerido", "Muda?"]
    _cab(ws, 4, cols, [12, 34, 7, 15, 18, 9, 9, 9, 9, 10, 12, 9, 9, 10, 10, 10, 11, 11, 8])
    ws.row_dimensions[4].height = 58
    piso = f"{P}{P_PISO}"
    for i, x in enumerate(g.itertuples(), 5):
        _fmt(ws, i, 1, _cod(x.cod))
        _fmt(ws, i, 2, x.grupo)
        _fmt(ws, i, 3, int(x.itens))
        _fmt(ws, i, 4, round(x.fat, 2), RS0)
        _fmt(ws, i, 5, x.fonte)
        _fmt(ws, i, 6, round(x.imp, 4), PCT, INPUT)
        _fmt(ws, i, 7, round(x.r, 4), PCT, INPUT)
        for j, d in enumerate([x.dS, x.dO, x.dD], 8):
            _fmt(ws, i, j, d, PCT, INPUT)
        _fmt(ws, i, 11, x.mD, PCT, INPUT)
        for j, col in enumerate("HIJ", 12):
            _fmt(ws, i, j, f"=1-$F{i}-$G{i}/(1-{col}{i})", PCT)
        for j, col in enumerate("HIJ", 15):
            _fmt(ws, i, j, f"=MAX(0,MIN({col}{i},ROUNDDOWN(1-$G{i}/(1-$F{i}-{piso}),2)))", PCT)
        _fmt(ws, i, 18, f"=1-$F{i}-$G{i}/(1-Q{i})", PCT)
        _fmt(ws, i, 19, f'=IF(OR(O{i}<H{i},P{i}<I{i},Q{i}<J{i}),"SIM","")', None, BOLD)
    _filtro(ws, 4, len(g) + 4, len(cols))
    ws.freeze_panes = "C5"
    ws.cell(len(g) + 6, 1, "Grupos com menos de R$ 5.000 vendidos para RAP no semestre usam o custo ÷ tabela "
                           "planejado e o imposto médio RAP.").font = Font(name=F, italic=True)


def recomendacao(wb, ult, rec):
    b = "'Base itens RAP'!"
    rng = lambda col: f"{b}${col}$2:${col}${ult}"  # noqa: E731
    grp = "'Desconto por grupo'!"
    gr = lambda col: f"{grp}${col}$5:${col}$200"  # noqa: E731
    ws = wb.create_sheet("Recomendação desconto")
    ws["A1"] = "Recomendação de desconto por categoria — PARA DECISÃO DA PAM / ALEXANDRE (não é corte automático)"
    ws["A1"].font = Font(name=F, bold=True, size=13, color="C00000")
    cols = ["Categoria", "RAPs", "RAPs que bateram o mínimo", "Faturamento semestre (R$)", "% do faturamento RAP",
            "Margem % atual", "Desconto efetivo praticado", "Custo ÷ preço tabela", "Impostos ÷ NF",
            "Faixa da régua", "Margem no piso da régua", "Margem no teto da régua",
            "Grupos com margem negativa no desconto atual", "Grupos com desconto reduzido (piso)",
            "Margem com o teto por grupo", "Recomendação", "Justificativa"]
    _cab(ws, 3, cols, [12, 7, 11, 16, 10, 10, 11, 10, 9, 11, 10, 10, 12, 12, 11, 30, 70])
    ws.row_dimensions[3].height = 58
    cls = "'RAPs - classificação'!"
    col_real = {"Standard": ("L", "O", "H"), "Ouro": ("M", "P", "I"), "Diamante": ("N", "Q", "J")}
    for i, cat in enumerate(A.CATEGORIAS, 4):
        x = rec[cat]
        lin = CAT_ROWS[cat]
        real, sug, atual = col_real[cat]
        ws.cell(i, 1, cat).font = BOLD
        _fmt(ws, i, 2, f'=COUNTIF({cls}$E$2:$E$9999,A{i})', INT)
        _fmt(ws, i, 3, f'=COUNTIFS({cls}$E$2:$E$9999,A{i},{cls}$I$2:$I$9999,"MANTER")', INT)
        _fmt(ws, i, 4, f"='Margem por categoria'!B{4 + A.CATEGORIAS.index(cat)}", RS0, LINK)
        _fmt(ws, i, 5, f"=D{i}/SUM($D$4:$D$6)", PCT)
        _fmt(ws, i, 6, f"='Margem por categoria'!G{4 + A.CATEGORIAS.index(cat)}", PCT, LINK)
        _fmt(ws, i, 7, f'=1-SUMIFS({rng("M")},{rng("E")},A{i},{rng("K")},"V",{rng("S")},">0")'
                       f'/SUMIFS({rng("S")},{rng("E")},A{i})', PCT)
        _fmt(ws, i, 8, f'=SUMIFS({rng("R")},{rng("E")},A{i},{rng("K")},"V",{rng("S")},">0")'
                       f'/SUMIFS({rng("S")},{rng("E")},A{i})', PCT)
        _fmt(ws, i, 9, f'=SUMIFS({rng("N")},{rng("E")},A{i},{rng("K")},"V")'
                       f'/SUMIFS({rng("M")},{rng("E")},A{i},{rng("K")},"V")', PCT)
        _fmt(ws, i, 10, f'=TEXT({P}$C${lin},"0%")&" a "&TEXT({P}$D${lin},"0%")')
        sim = lambda d: f"=1-$I{i}-$H{i}/(1-{d})"  # noqa: E731
        _fmt(ws, i, 11, sim(f"{P}$C${lin}"), PCT)
        _fmt(ws, i, 12, sim(f"{P}$D${lin}"), PCT)
        _fmt(ws, i, 13, f'=COUNTIF({gr(real)},"<0")', INT)
        _fmt(ws, i, 14, f'=SUMPRODUCT(--({gr(sug)}<{gr(atual)}))', INT)
        _fmt(ws, i, 15, round(x["m_teto_grupo"], 4), PCT, INPUT).comment = Comment(
            "Calculado no script: mesmas NFs do semestre, trocando o desconto de cada grupo pelo sugerido.", "Claude")
        _fmt(ws, i, 16, x["acao"], None, BOLD).alignment = Alignment(wrap_text=True, vertical="top")
        c = _fmt(ws, i, 17, x["justificativa"])
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = 62
    ws["A8"] = ("Margem = (NF − impostos − custo) ÷ NF. Simulação = mesmo mix e custo do semestre trocando só o desconto "
                "sobre a tabela: 1 − impostos − (custo ÷ tabela) ÷ (1 − desconto). Desconto efetivo = 1 − NF ÷ (tabela "
                "atual × qtd); é aproximado porque a tabela pode ter mudado no período.")
    ws["A8"].font = Font(name=F, italic=True)


RECOMENDACAO = {
    "Diamante": {"acao": "Não subir para 45–55%. Manter a tabela por grupo (35–45%) com teto nos grupos fracos",
                 "titulo": "≤ 45%", "sub": "tabela atual por grupo + teto nos grupos fracos",
                 "justificativa": "Traz 94% do faturamento RAP com 18% de margem. Com o imposto real das NFs, 21 "
                                  "grupos já ficam negativos no desconto atual (FPS, consumíveis, coroas, haste, "
                                  "Cummins, material rodante); no teto da régua (55%) a categoria inteira fica "
                                  "negativa. Limitar o desconto desses grupos ao piso de margem."},
    "Ouro": {"acao": "Manter a tabela atual (22–35%); não subir para 40%",
             "titulo": "≤ 35%", "sub": "tabela atual por grupo",
             "justificativa": "Pouco volume (5% do faturamento) e nenhuma das 172 RAPs Ouro bateu o mínimo; margem "
                              "de 31% hoje. Só Haste cilindro precisa cair. Desconto maior só como prêmio de volume."},
    "Standard": {"acao": "Manter 18–20%", "titulo": "20%", "sub": "tabela atual (FPS 18%)",
                 "justificativa": "Margem saudável em todos os grupos (exceto Haste cilindro); faturamento residual "
                                  "(<1%) e 156 de 159 Standard sem atingir a régua. É porta de entrada, não de margem."},
}


def margem_teto_grupo(itens, precos, grupos):
    """Margem de cada categoria se cada grupo usasse o desconto sugerido (mesmas NFs do semestre)."""
    v = itens[itens.sinal > 0].copy()
    v["tab"] = v.CODPROD.map(precos) * v.QTDNEG
    v = v[v.tab > 0]
    out = {}
    for cat, k in [("Standard", "S"), ("Ouro", "O"), ("Diamante", "D")]:
        g = v[v.categoria == cat]
        t = g.IMPOSTOS.sum() / g.BRUTO.sum()
        efet = 1 - g.BRUTO / g.tab                                   # desconto praticado em cada linha
        corte = (grupos.set_index("cod")["d" + k] - grupos.set_index("cod")["sug" + k]).clip(lower=0)
        novo = (efet - g.CODGRUPOPROD.map(corte).fillna(0)).clip(lower=0)
        nf = g.tab * (1 - novo)
        # ganho de margem sobre a mesma base (só vendas), somado à margem realizada da categoria
        out[cat] = (1 - t - g.custo_total.sum() / nf.sum()) - (1 - t - g.custo_total.sum() / g.BRUTO.sum())
    return out


def deck_json(r, mc, rec_num):
    c = r["classificacao"]
    out = {"periodo": "01/07/2026 a 16/10/2026", "dias": [108, 184],
           "imp_rap": r["imp_rap"], "imp_balcao": r["imp_balcao"], "piso": A.PISO_MARGEM}
    bal = c.groupby("balde").agg(n=("CODPARC", "size"), fat=("fat", "sum"), falta=("falta", "sum"))
    out["baldes"] = {b: {"n": int(bal.loc[b, "n"]), "fat": float(bal.loc[b, "fat"]),
                         "falta": float(bal.loc[b, "falta"])} for b in A.BALDES}
    out["total_raps"] = int(len(c))
    out["cat_balde"] = {cat: {b: int(((c.categoria == cat) & (c.balde == b)).sum()) for b in A.BALDES}
                        for cat in A.CATEGORIAS}
    cons = sorted(c.consultor.unique())
    out["consultor"] = {k: {b: int(((c.consultor == k) & (c.balde == b)).sum()) for b in A.BALDES} for k in cons}
    out["consultor_fat"] = {k: float(c[c.consultor == k].fat.sum()) for k in cons}
    d = c[c.balde == "DESHOMOLOGAR"]
    out["desh_1s"] = {"n": int((d.fat_1s > 0).sum()), "fat": float(d[d.fat_1s > 0].fat_1s.sum())}
    out["desh_std_baixo"] = int((d.motivo == "Standard abaixo de 70% do mínimo").sum())
    out["desh_cons_cat"] = {k: {cat: int(((d.consultor == k) & (d.categoria == cat)).sum())
                                for cat in A.CATEGORIAS} for k in cons}
    out["desh_top1s"] = [{"nome": x.NOMEPARC, "cons": x.consultor, "cat": x.categoria, "fat1s": x.fat_1s}
                         for x in d[d.fat_1s > 0].sort_values("fat_1s", ascending=False).head(10).itertuples()]
    rb = c[c.balde == "REBAIXAR"]
    out["reb_trans"] = {f"{a} → {b}": int(n) for (a, b), n in rb.groupby(["categoria", "sugerida"]).size().items()}
    out["reb_top"] = [{"nome": x.NOMEPARC, "cons": x.consultor, "de": x.categoria, "para": x.sugerida,
                       "fat": x.fat, "pct": x.pct} for x in rb.sort_values("fat", ascending=False).head(12).itertuples()]
    out["reb_proj"] = int((rb.pct_proj >= A_LIM).sum())
    out["inc"] = [{"nome": x.NOMEPARC, "cons": x.consultor, "cat": x.categoria, "fat": x.fat,
                   "pct": x.pct, "falta": x.falta} for x in c[c.balde == "INCENTIVAR"].itertuples()]
    out["manter_fat"] = float(c[c.balde == "MANTER"].fat.sum())
    out["inconsistentes"] = {"cat_sem_vend": int((r["inconsistente"].Problema.str.startswith("Categoria")).sum()),
                             "vend_sem_cat": int((r["inconsistente"].Problema.str.startswith("Vendedor")).sum())}
    ir, ib = r["imp_rap"], r["imp_balcao"]
    cesta = []
    for x in r["cesta"].itertuples():
        row = {"pn": x.REFERENCIA, "desc": x.DESCRPROD, "grupo": x.CODGRUPOPROD[:3], "custo": x.custo,
               "tab": x.preco}
        canais = {"balcao": (A.DESC_BALCAO, ib), "Standard": (x.dS, ir), "Ouro": (x.dO, ir),
                  "Diamante": (x.dD, ir), "sug": (x.sugD, ir)}
        for k, (dsc, imp) in canais.items():
            p = x.preco * (1 - dsc)
            row[k] = {"d": dsc, "p": p, "m": 1 - imp - x.custo / p if x.custo > 0 else None}
        cesta.append(row)
    out["cesta"] = cesta
    g = r["grupos"]
    out["grupos_n"] = {"total": int(len(g)), "negD": int((g.mD_real < 0).sum()), "negO": int((g.mO_real < 0).sum()),
                       "negS": int((g.mS_real < 0).sum()), "muda": int(g.muda.sum()),
                       "planD_lt8": int((g.mD < 0.08).sum())}
    neg = g[(g.mD_real < A.PISO_MARGEM) & (g.fat >= 5000)].sort_values("fat", ascending=False)
    out["grupos_neg"] = [{"g": x.grupo, "cod": x.cod[:3], "fat": x.fat, "d": x.dD, "mplan": x.mD, "mreal": x.mD_real,
                          "sug": x.sugD} for x in neg.head(12).itertuples()]
    ok = g[(g.mD_real >= 0.14) & (g.fat >= 50000)].sort_values("fat", ascending=False)
    out["grupos_ok"] = [x.grupo for x in ok.head(8).itertuples()]
    out["margem"] = [{"cat": x.categoria, "fat": x.fat_nf, "liq": x.receita_liq, "custo": x.custo,
                      "mrs": x.margem_rs, "mp": x.margem_pct} for x in mc.itertuples()]
    top = r["margem_cat_grupo"].sort_values("fat_nf", ascending=False)
    out["top_grupos"] = {cat: [{"g": x.grupo, "fat": x.fat_nf, "mp": x.margem_pct}
                               for x in top[top.categoria == cat].head(4).itertuples()] for cat in A.CATEGORIAS}
    out["rec"] = {cat: {**RECOMENDACAO[cat], **rec_num[cat]} for cat in A.CATEGORIAS}
    return out


A_LIM = 0.7


def simular(itens, precos, grupos, mc):
    """Desconto efetivo e margem simulada por categoria (mesmo cálculo das fórmulas da aba Recomendação)."""
    v = itens[(itens.sinal > 0)].copy()
    v["tab"] = v.CODPROD.map(precos) * v.QTDNEG
    teto = margem_teto_grupo(itens, precos, grupos)
    out = {}
    for cat, g in v.groupby("categoria"):
        gt = g[g.tab > 0]
        r_ = gt.custo_total.sum() / gt.tab.sum()
        t = g.IMPOSTOS.sum() / g.BRUTO.sum()
        sim = lambda d: 1 - t - r_ / (1 - d)  # noqa: E731
        lo, hi = A.FAIXA[cat]
        out[cat] = {"desc_ef": 1 - gt.BRUTO.sum() / gt.tab.sum(), "m_piso": sim(lo), "m_teto": sim(hi),
                    "m_teto_grupo": float(mc.set_index("categoria").margem_pct[cat]) + teto[cat], "imp": t}
    return out


def gerar():
    r = A.rodar()
    _, _, _, _, pre, _ = A.carregar()
    precos = dict(zip(pre.CODPROD, pre.preco))
    rec_num = simular(r["itens"], precos, r["grupos"], r["margem_cat"])
    rec = {cat: {**RECOMENDACAO[cat], **rec_num[cat]} for cat in A.CATEGORIAS}
    SAIDA.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    premissas(wb, r)
    ult = classificacao(wb, r["classificacao"])
    resumos(wb, ult, sorted(r["classificacao"].consultor.unique()))
    inconsistente(wb, r["inconsistente"])
    estudo_preco(wb, r["cesta"])
    ult_b = base_itens(wb, r["itens"], precos)
    margens(wb, ult_b, r["margem_cat_grupo"])
    desconto_grupo(wb, r["grupos"])
    recomendacao(wb, ult_b, rec)
    # ordem das abas: as pedidas primeiro, base de itens no fim
    nomes = ["Premissas", "RAPs - classificação", "Resumo por consultor", "Resumo por balde",
             "Cadastro inconsistente", "Estudo de preço", "Margem por categoria", "Margem por categoria x grupo",
             "Desconto por grupo", "Recomendação desconto", "Base itens RAP"]
    wb._sheets = [wb[n] for n in nomes]
    wb.save(ARQ)
    js = deck_json(r, r["margem_cat"], rec_num)
    (SAIDA / "deck.json").write_text(json.dumps(js, ensure_ascii=False, indent=1, default=float))
    return r, rec_num


if __name__ == "__main__":
    _, rec = gerar()
    print(ARQ)
    for k, v in rec.items():
        print(k, {a: round(b, 3) for a, b in v.items()})
