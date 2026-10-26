// Apresentação da Reestruturação RAP a partir de saida/rap/deck.json (gerado por `python -m rap.saidas`).
// Uso: node rap/deck.js   (precisa do pacote pptxgenjs; NODE_PATH pode apontar para onde ele estiver)
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

const SAIDA = path.join(__dirname, "..", "saida", "rap");
const D = JSON.parse(fs.readFileSync(path.join(SAIDA, "deck.json"), "utf8"));

const NAVY = "12285A", NAVY2 = "1C3A7A", AMARELO = "F5B800", VERDE = "5E9B1E", CINZA = "5F6B7A";
const CLARO = "EEF2F8", BRANCO = "FFFFFF", TEXTO = "1D2633", VERMELHO = "B3261E";
const COR_BALDE = { DESHOMOLOGAR: VERMELHO, REBAIXAR: "E07B00", INCENTIVAR: "2F7FC1", MANTER: VERDE };
const BALDES = ["DESHOMOLOGAR", "REBAIXAR", "INCENTIVAR", "MANTER"];
const FONT = "Arial";
const IMPOSTOS = 0.22, DESC_BALCAO = 0.15;
const PISO = { Standard: 0.20, Ouro: 0.30, Diamante: 0.45 };
const TETO = { Standard: 0.25, Ouro: 0.40, Diamante: 0.55 };

const nf0 = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 });
const rs = (v) => "R$ " + nf0.format(Math.round(v));
const pct = (v) => (v == null ? "s/ custo" : (v * 100).toFixed(1).replace(".", ",") + "%");
const corMargem = (m) => (m == null ? CINZA : m < 0.10 ? VERMELHO : m < 0.25 ? "C77700" : VERDE);

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "BMC Hyundai — Peças";
pres.title = "Reestruturação RAP — Fechamento do Semestre";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

const W = 13.33, M = 0.6;
const RODAPE = `Fonte: Sankhya (d_Clientes, f_Fat_Clientes, f_Notas, d_Precos, f_CustoUnitarioProd) · período ${D.periodo}`;

pres.defineSlideMaster({
  title: "ESCURO",
  background: { color: NAVY },
  objects: [
    { text: { text: "BMC  |  HYUNDAI", options: { x: M, y: 6.75, w: 4, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: BRANCO, charSpacing: 2 } } },
  ],
});
pres.defineSlideMaster({
  title: "CONTEUDO",
  background: { color: BRANCO },
  margin: [0.5, 0.6, 0.7, 0.6],
  objects: [
    { text: { text: RODAPE, options: { x: M, y: 7.0, w: 9.5, h: 0.3, fontFace: FONT, fontSize: 9, color: CINZA } } },
    { text: { text: "BMC | HYUNDAI", options: { x: W - M - 2.2, y: 7.0, w: 1.7, h: 0.3, fontFace: FONT, fontSize: 9, bold: true, color: NAVY, align: "right" } } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.35, w: W - 2 * M, h: 0.75, fontFace: FONT, fontSize: 30, bold: true, color: NAVY, valign: "middle", align: "left" }, text: "" } },
    { placeholder: { options: { name: "sub", type: "body", x: M, y: 1.05, w: W - 2 * M, h: 0.4, fontFace: FONT, fontSize: 14, color: CINZA, valign: "top", align: "left" }, text: "" } },
  ],
  slideNumber: { x: W - M - 0.45, y: 7.0, w: 0.45, h: 0.3, fontFace: FONT, fontSize: 9, color: CINZA, align: "right" },
});

let sec = "";
function slideConteudo(secao, titulo, sub) {
  if (secao !== sec) { pres.addSection({ title: secao }); sec = secao; }
  const s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: secao });
  s.addText(titulo, { placeholder: "title" });
  if (sub) s.addText(sub, { placeholder: "sub" });
  return s;
}
function card(s, x, y, w, h, fill, name) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08, objectName: name });
}
function txt(s, text, o) {
  s.addText(text, Object.assign({ fontFace: FONT, color: TEXTO, isTextBox: true, margin: 0 }, o));
}
function tabela(s, linhas, o) {
  s.addTable(linhas, Object.assign({ fontFace: FONT, fontSize: 11, color: TEXTO, border: { type: "solid", pt: 0.5, color: "D5DBE5" }, valign: "middle", margin: [0.03, 0.06, 0.03, 0.06] }, o));
}
const H = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true, color: BRANCO, fill: { color: NAVY }, align: "center" }, o) });
const C = (t, o = {}) => ({ text: String(t), options: o });

// 1 — Capa
pres.addSection({ title: "Abertura" }); sec = "Abertura";
{
  const s = pres.addSlide({ masterName: "ESCURO", sectionTitle: "Abertura" });
  txt(s, "REESTRUTURAÇÃO RAP", { x: M, y: 1.7, w: 11, h: 0.6, fontSize: 20, bold: true, color: AMARELO, charSpacing: 4 });
  txt(s, "Fechamento do Semestre", { x: M, y: 2.3, w: 11, h: 1.2, fontSize: 54, bold: true, color: BRANCO });
  txt(s, `2º semestre/2026 · NFs de ${D.periodo} · ${D.total_raps} RAPs avaliadas`, { x: M, y: 3.7, w: 11, h: 0.5, fontSize: 18, color: "C9D4EA" });
  txt(s, "Material para aprovação da Pam e do Alexandre — nenhuma categoria foi alterada e nenhum e-mail foi enviado", { x: M, y: 4.3, w: 11, h: 0.5, fontSize: 14, italic: true, color: "C9D4EA" });
}

// 2 — Régua e descontos
{
  const s = slideConteudo("Contexto", "Régua semestral e faixas de desconto", "RAP = vendedor preferencial ESILVA, SLUNA, PLOUREIRO ou FALEXANDRE + categoria Standard, Ouro ou Diamante");
  const cats = [["Standard", 5000, "20% – 25%", "8A97AB"], ["Ouro", 40000, "30% – 40%", "C99A06"], ["Diamante", 100000, "45% – 55%", NAVY2]];
  cats.forEach(([n, min, faixa, cor], i) => {
    const x = M + i * 4.1, y = 1.75;
    card(s, x, y, 3.8, 2.55, CLARO, `card-${n}`);
    s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.3, w: 0.5, h: 0.5, fill: { color: cor }, line: { color: cor } });
    txt(s, n.toUpperCase(), { x: x + 0.95, y: y + 0.3, w: 2.7, h: 0.5, fontSize: 18, bold: true, color: NAVY, valign: "middle" });
    txt(s, "Mínimo no semestre", { x: x + 0.3, y: y + 1.0, w: 3.3, h: 0.3, fontSize: 12, color: CINZA });
    txt(s, rs(min), { x: x + 0.3, y: y + 1.28, w: 3.3, h: 0.5, fontSize: 26, bold: true, color: NAVY });
    txt(s, "Desconto sobre a tabela nova", { x: x + 0.3, y: y + 1.8, w: 3.3, h: 0.3, fontSize: 12, color: CINZA });
    txt(s, faixa, { x: x + 0.3, y: y + 2.05, w: 3.3, h: 0.4, fontSize: 20, bold: true, color: TEXTO });
  });
  txt(s, "Como cada RAP é classificada (faturamento do semestre ÷ mínimo da categoria atual)", { x: M, y: 4.6, w: 12, h: 0.35, fontSize: 14, bold: true, color: NAVY });
  const regras = [["DESHOMOLOGAR", "R$ 0 no semestre\n(ou Standard < 70%)"], ["REBAIXAR", "< 70% do mínimo\nDiamante → Ouro → Standard"], ["INCENTIVAR", "70% a 100%\nlista de contato"], ["MANTER", "≥ 100%\nsem mudança"]];
  regras.forEach(([b, t], i) => {
    const x = M + i * 3.07, y = 5.05;
    card(s, x, y, 2.85, 1.55, COR_BALDE[b], `regra-${b}`);
    txt(s, b, { x: x + 0.2, y: y + 0.15, w: 2.5, h: 0.4, fontSize: 15, bold: true, color: BRANCO });
    txt(s, t, { x: x + 0.2, y: y + 0.6, w: 2.5, h: 0.8, fontSize: 13, color: BRANCO, valign: "top" });
  });
}

// 3 — Resumo executivo
{
  const B = D.baldes;
  const s = slideConteudo("Classificação", "Resumo executivo", `${D.total_raps} RAPs · faturamento de ${D.periodo} (${D.dias[0]} de ${D.dias[1]} dias do semestre)`);
  BALDES.forEach((b, i) => {
    const x = M + i * 3.07, y = 1.7;
    card(s, x, y, 2.85, 2.6, CLARO, `kpi-${b}`);
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: y + 0.3, w: 0.28, h: 0.28, fill: { color: COR_BALDE[b] }, line: { color: COR_BALDE[b] } });
    txt(s, b, { x: x + 0.65, y: y + 0.25, w: 2.1, h: 0.38, fontSize: 13, bold: true, color: TEXTO, valign: "middle" });
    txt(s, String(B[b].n), { x: x + 0.25, y: y + 0.75, w: 2.4, h: 0.9, fontSize: 54, bold: true, color: COR_BALDE[b] });
    txt(s, `RAPs · ${Math.round((B[b].n / D.total_raps) * 100)}% do total`, { x: x + 0.25, y: y + 1.65, w: 2.5, h: 0.3, fontSize: 12, color: CINZA });
    txt(s, rs(B[b].fat), { x: x + 0.25, y: y + 1.98, w: 2.5, h: 0.4, fontSize: 17, bold: true, color: NAVY });
  });
  const pontos = [
    `MANTER: só ${B.MANTER.n} RAPs, mas ${Math.round((B.MANTER.fat / (B.MANTER.fat + B.REBAIXAR.fat + B.INCENTIVAR.fat + B.DESHOMOLOGAR.fat)) * 100)}% do faturamento RAP (${rs(B.MANTER.fat)})`,
    `REBAIXAR: ${D.reb_trans["Diamante → Ouro"]} Diamante → Ouro e ${D.reb_trans["Ouro → Standard"]} Ouro → Standard · ${rs(B.REBAIXAR.fat)} faturados`,
    `INCENTIVAR: faltam ${rs(B.INCENTIVAR.falta)} no total para as ${B.INCENTIVAR.n} RAPs manterem a categoria`,
    `DESHOMOLOGAR: ${B.DESHOMOLOGAR.n - D.desh_std_baixo} sem nenhuma compra + ${D.desh_std_baixo} Standard abaixo de 70%; ${D.desh_1s.n} delas compraram ${rs(D.desh_1s.fat)} no 1º sem → revisar antes do e-mail`,
    `Semestre aberto: no ritmo atual, ${D.reb_proj} das ${B.REBAIXAR.n} RAPs em REBAIXAR passariam de 70% até 31/12 → reclassificar no fechamento`,
  ];
  s.addText(pontos.map((p, i) => ({ text: p, options: { bullet: true, breakLine: i < pontos.length - 1 } })),
    { x: M, y: 4.55, w: W - 2 * M, h: 2.3, fontFace: FONT, fontSize: 14, color: TEXTO, paraSpaceAfter: 6, valign: "top", isTextBox: true });
}

// 4 — Por consultor (barras empilhadas)
{
  const s = slideConteudo("Classificação", "Classificação por consultor", "Nº de RAPs por balde · faturamento do semestre ao lado");
  const nomes = Object.keys(D.consultor);
  const data = BALDES.map((b) => ({ name: b, labels: nomes, values: nomes.map((n) => D.consultor[n][b]) }));
  s.addChart(pres.charts.BAR, data, {
    x: M, y: 1.6, w: 8.4, h: 5.2, barDir: "bar", barGrouping: "stacked", chartColors: BALDES.map((b) => COR_BALDE[b]),
    showLegend: true, legendPos: "b", legendFontFace: FONT, legendFontSize: 12,
    showValue: true, dataLabelPosition: "ctr", dataLabelColor: BRANCO, dataLabelFontSize: 11, dataLabelFontFace: FONT, dataLabelFormatCode: "#,##0;;",
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 12, catAxisLabelColor: TEXTO, catAxisOrientation: "maxMin",
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 10, valAxisLabelColor: CINZA,
    valGridLine: { color: "E1E6EE", size: 0.5 }, catGridLine: { style: "none" },
  });
  nomes.forEach((n, i) => {
    const y = 1.75 + i * 1.22;
    card(s, 9.35, y, 3.38, 1.05, CLARO, `fat-${n}`);
    txt(s, n, { x: 9.55, y: y + 0.1, w: 3.0, h: 0.35, fontSize: 13, bold: true, color: NAVY });
    const tot = BALDES.reduce((a, b) => a + D.consultor[n][b], 0);
    txt(s, `${rs(D.consultor_fat[n])} · ${tot} RAPs`, { x: 9.55, y: y + 0.5, w: 3.0, h: 0.4, fontSize: 14, color: TEXTO });
  });
  txt(s, "SLUNA: as 126 RAPs que ficaram na carteira depois da redistribuição não compraram nada no semestre.", { x: 9.35, y: 6.6, w: 3.4, h: 0.4, fontSize: 10, italic: true, color: CINZA });
}

// 5 — Deshomologar
{
  const B = D.baldes.DESHOMOLOGAR;
  const s = slideConteudo("Listas", `DESHOMOLOGAR — ${B.n} RAPs para encerramento`, "Lista completa com CNPJ na aba “RAPs - classificação” (filtro Balde = DESHOMOLOGAR)");
  const cats = ["Diamante", "Ouro", "Standard"];
  const nomes = Object.keys(D.desh_cons_cat);
  const linhas = [[H("Consultor", { align: "left" }), ...cats.map((c) => H(c)), H("Total")]];
  nomes.forEach((n) => {
    const v = cats.map((c) => D.desh_cons_cat[n][c]);
    linhas.push([C(n), ...v.map((x) => C(x, { align: "center" })), C(v.reduce((a, b) => a + b, 0), { align: "center", bold: true })]);
  });
  const tc = cats.map((c) => nomes.reduce((a, n) => a + D.desh_cons_cat[n][c], 0));
  linhas.push([C("Total", { bold: true }), ...tc.map((x) => C(x, { align: "center", bold: true })), C(B.n, { align: "center", bold: true })]);
  tabela(s, linhas, { x: M, y: 1.7, w: 6.0, colW: [2.4, 0.9, 0.9, 0.9, 0.9], rowH: 0.42, fontSize: 12 });
  txt(s, `Inclui ${D.desh_std_baixo} Standard que faturaram algo, mas abaixo de 70% do mínimo (regra: Standard não tem para onde rebaixar).`, { x: M, y: 4.5, w: 6.0, h: 0.7, fontSize: 12, color: CINZA });

  card(s, 7.0, 1.7, 5.73, 5.1, CLARO, "revisar");
  txt(s, `Revisar antes do disparo: ${D.desh_1s.n} compraram no 1º sem (${rs(D.desh_1s.fat)})`, { x: 7.2, y: 1.8, w: 5.4, h: 0.6, fontSize: 14, bold: true, color: VERMELHO });
  const l2 = [[H("RAP", { align: "left" }), H("Cat."), H("1º sem (R$)")]];
  D.desh_top1s.slice(0, 9).forEach((r) => l2.push([C(r.nome.slice(0, 32), { fontSize: 10 }), C(r.cat, { align: "center", fontSize: 10 }), C(rs(r.fat1s), { align: "right", fontSize: 10 })]));
  tabela(s, l2, { x: 7.2, y: 2.45, w: 5.33, colW: [3.05, 0.98, 1.3], rowH: 0.4 });
}

// 6 — Rebaixar
{
  const B = D.baldes.REBAIXAR;
  const s = slideConteudo("Listas", `REBAIXAR — ${B.n} RAPs mudam de categoria`, "Abaixo de 70% do mínimo da categoria atual · maiores faturamentos abaixo; lista completa na planilha");
  Object.entries(D.reb_trans).forEach(([t, n], i) => {
    const y = 1.7 + i * 1.6;
    card(s, M, y, 3.4, 1.4, CLARO, `trans-${i}`);
    txt(s, String(n), { x: M + 0.25, y: y + 0.12, w: 3.0, h: 0.75, fontSize: 44, bold: true, color: COR_BALDE.REBAIXAR });
    txt(s, t, { x: M + 0.25, y: y + 0.9, w: 3.0, h: 0.4, fontSize: 16, bold: true, color: NAVY });
  });
  txt(s, `No ritmo atual, ${D.reb_proj} delas chegariam a 70% até 31/12.`, { x: M, y: 5.0, w: 3.4, h: 0.8, fontSize: 12, italic: true, color: CINZA });
  const l = [[H("RAP", { align: "left" }), H("Consultor"), H("De → Para"), H("Faturamento"), H("% do mín.")]];
  D.reb_top.slice(0, 11).forEach((r) => l.push([C(r.nome.slice(0, 34), { fontSize: 10 }), C(r.cons.replace(" (ELISASILVA)", ""), { fontSize: 10, align: "center" }), C(`${r.de} → ${r.para}`, { fontSize: 10, align: "center" }), C(rs(r.fat), { fontSize: 10, align: "right" }), C(pct(r.pct), { fontSize: 10, align: "center" })]));
  tabela(s, l, { x: 4.35, y: 1.7, w: 8.38, colW: [3.3, 1.35, 1.6, 1.23, 0.9], rowH: 0.42 });
}

// 7 — Incentivar
{
  const B = D.baldes.INCENTIVAR;
  const s = slideConteudo("Listas", `INCENTIVAR — ${B.n} RAPs perto do mínimo`, `Entre 70% e 100% do mínimo · faltam ${rs(B.falta)} no total · lista de contato, sem mudança de categoria`);
  const l = [[H("RAP", { align: "left" }), H("Consultor"), H("Categoria"), H("Faturado"), H("% do mínimo"), H("Quanto falta", { fill: { color: AMARELO }, color: NAVY })]];
  D.inc.forEach((r) => l.push([C(r.nome, { fontSize: 12 }), C(r.cons.replace(" (ELISASILVA)", ""), { align: "center" }), C(r.cat, { align: "center" }), C(rs(r.fat), { align: "right" }), C(pct(r.pct), { align: "center" }), C(rs(r.falta), { align: "right", bold: true, color: NAVY })]));
  tabela(s, l, { x: M, y: 1.75, w: W - 2 * M, colW: [4.4, 1.6, 1.4, 1.6, 1.3, 1.83], rowH: 0.5, fontSize: 12 });
  card(s, M, 5.6, W - 2 * M, 1.1, CLARO, "dica");
  txt(s, "Abordagem sugerida: o consultor mostra quanto falta e oferece um pedido que feche a régua antes de 31/12 — as 2 Diamante precisam de menos de R$ 5,2 mil cada.", { x: M + 0.25, y: 5.7, w: W - 2 * M - 0.5, h: 0.9, fontSize: 14, color: TEXTO, valign: "middle" });
}

// 8 — Estudo de preço
{
  const s = slideConteudo("Preço e margem", "Preço da peça: Balcão vs Standard, Ouro e Diamante", "Preço novo de tabela menos o desconto de cada canal (piso da faixa RAP) · margem após 22% de impostos");
  const cols = [["Balcão 15%", "balcao", DESC_BALCAO], ["Standard 20%", "Standard", PISO.Standard], ["Ouro 30%", "Ouro", PISO.Ouro], ["Diamante 45%", "Diamante", PISO.Diamante]];
  const l = [[H("PN", { align: "left" }), H("Descrição", { align: "left" }), H("Grupo"), H("Custo"), H("Tabela"), ...cols.map((c) => H(c[0]))]];
  D.cesta.forEach((r) => {
    const cel = cols.map(([, , d]) => {
      const p = r.tab * (1 - d), liq = p * (1 - IMPOSTOS);
      const m = r.custo > 0 ? (liq - r.custo) / liq : null;
      return { text: [{ text: rs(p), options: { breakLine: true, fontSize: 10, color: TEXTO } }, { text: pct(m), options: { fontSize: 10, bold: true, color: corMargem(m) } }], options: { align: "center" } };
    });
    l.push([C(r.pn, { fontSize: 9 }), C(r.desc.slice(0, 28), { fontSize: 9 }), C(r.grupo, { align: "center", fontSize: 9 }), C(rs(r.custo), { align: "right", fontSize: 9 }), C(rs(r.tab), { align: "right", fontSize: 9 }), ...cel]);
  });
  tabela(s, l, { x: M, y: 1.5, w: W - 2 * M, colW: [1.45, 2.7, 0.65, 1.15, 1.15, 1.25, 1.25, 1.25, 1.28], rowH: 0.34, margin: [0.01, 0.05, 0.01, 0.05] });
  txt(s, "Margem: verde ≥ 25% · laranja 10–25% · vermelho < 10%. No teto da faixa (25/40/55%) a margem cai mais — ver aba “Estudo de preço”.", { x: M, y: 6.62, w: W - 2 * M, h: 0.3, fontSize: 10, italic: true, color: CINZA });
}

// 9 — Margem por categoria
{
  const s = slideConteudo("Preço e margem", "Margem das RAPs por categoria no semestre", "NFs de venda − devoluções · margem = receita líquida de impostos − custo · nenhum item ficou sem custo");
  const tot = D.margem.reduce((a, m) => a + m.fat, 0);
  D.margem.forEach((m, i) => {
    const x = M + i * 4.1, y = 1.65;
    card(s, x, y, 3.8, 2.3, CLARO, `marg-${m.cat}`);
    txt(s, m.cat.toUpperCase(), { x: x + 0.25, y: y + 0.18, w: 3.3, h: 0.35, fontSize: 15, bold: true, color: NAVY });
    txt(s, pct(m.mp), { x: x + 0.25, y: y + 0.55, w: 3.3, h: 0.75, fontSize: 40, bold: true, color: corMargem(m.mp) });
    txt(s, `margem · ${rs(m.mrs)}`, { x: x + 0.25, y: y + 1.32, w: 3.3, h: 0.3, fontSize: 12, color: CINZA });
    txt(s, `Faturado ${rs(m.fat)} (${Math.round((m.fat / tot) * 100)}%)`, { x: x + 0.25, y: y + 1.72, w: 3.4, h: 0.35, fontSize: 13, bold: true, color: TEXTO });
  });
  txt(s, "Grupos que mais pesam em cada categoria (faturamento · margem)", { x: M, y: 4.2, w: 12, h: 0.35, fontSize: 14, bold: true, color: NAVY });
  ["Diamante", "Ouro", "Standard"].forEach((cat, i) => {
    const l = D.top_grupos[cat].slice(0, 4).map((g) => [C(g.g.slice(0, 26), { fontSize: 10 }), C(rs(g.fat), { align: "right", fontSize: 10 }), C(pct(g.mp), { align: "center", fontSize: 10, bold: true, color: corMargem(g.mp) })]);
    tabela(s, l, { x: M + i * 4.1, y: 4.65, w: 3.8, colW: [2.0, 1.1, 0.7], rowH: 0.42 });
  });
}

// 10 — Recomendação
{
  const s = slideConteudo("Recomendação", "Recomendação de desconto por categoria", "Para decisão da Pam e do Alexandre — não é corte automático");
  ["Diamante", "Ouro", "Standard"].forEach((cat, i) => {
    const r = D.rec[cat], m = D.margem.find((x) => x.cat === cat);
    const x = M + i * 4.1, y = 1.65;
    card(s, x, y, 3.8, 3.75, CLARO, `rec-${cat}`);
    txt(s, cat.toUpperCase(), { x: x + 0.25, y: y + 0.18, w: 3.3, h: 0.35, fontSize: 15, bold: true, color: NAVY });
    txt(s, `Faixa ${Math.round(PISO[cat] * 100)}–${Math.round(TETO[cat] * 100)}% · ` + (r.desc_ef > 0.01 ? `praticado hoje ≈ ${Math.round(r.desc_ef * 100)}%` : "NF sai ≈ preço de tabela"), { x: x + 0.25, y: y + 0.55, w: 3.4, h: 0.3, fontSize: 11, color: CINZA });
    txt(s, r.faixa_txt, { x: x + 0.25, y: y + 0.9, w: 3.3, h: 0.85, fontSize: 48, bold: true, color: NAVY });
    txt(s, "desconto recomendado", { x: x + 0.25, y: y + 1.72, w: 3.3, h: 0.3, fontSize: 12, color: CINZA });
    txt(s, [
      { text: "Margem hoje ", options: { color: CINZA } }, { text: pct(m.mp), options: { bold: true, color: corMargem(m.mp), breakLine: true } },
      { text: "No recomendado ", options: { color: CINZA } }, { text: pct(r.m_rec), options: { bold: true, color: corMargem(r.m_rec), breakLine: true } },
      { text: "No teto da faixa ", options: { color: CINZA } }, { text: pct(r.m_teto), options: { bold: true, color: corMargem(r.m_teto) } },
    ], { x: x + 0.25, y: y + 2.1, w: 3.3, h: 0.95, fontSize: 13 });
    txt(s, r.justificativa.split(" → ")[1] ? "→ " + r.justificativa.split(" → ")[1] : r.justificativa, { x: x + 0.25, y: y + 3.15, w: 3.35, h: 1.3, fontSize: 12, color: TEXTO, valign: "top" });
  });
  card(s, M, 5.75, W - 2 * M, 0.8, "FFF4CC", "alerta");
  txt(s, "Atenção: a faixa Diamante de 45–55% fica acima do que se pratica hoje (≈36%). Aplicada no teto, a categoria que traz 94% do faturamento RAP passa a dar prejuízo.", { x: M + 0.2, y: 5.8, w: W - 2 * M - 0.4, h: 0.7, fontSize: 14, bold: true, color: "7A5200", valign: "middle" });
}

// 11 — Próximos passos
pres.addSection({ title: "Encerramento" });
{
  const s = pres.addSlide({ masterName: "ESCURO", sectionTitle: "Encerramento" });
  txt(s, "Próximos passos", { x: M, y: 0.5, w: 11, h: 0.9, fontSize: 36, bold: true, color: BRANCO });
  const passos = [
    ["Aprovar", "Pam e Alexandre validam baldes e descontos por categoria"],
    ["Revisar", `${D.desh_1s.n} RAPs em DESHOMOLOGAR que compraram no 1º sem e as ${D.baldes.INCENTIVAR.n} de INCENTIVAR, com os consultores`],
    ["Corrigir cadastro", `${D.inconsistentes.cat_sem_vend} clientes com categoria RAP sem vendedor RAP e ${D.inconsistentes.vend_sem_cat} RAPs sem categoria válida`],
    ["Comunicar", "E-mails de encerramento e de nova categoria — só depois da aprovação"],
    ["Atualizar o Sankhya", "AD_CLASSRAP e vendedor preferencial conforme a decisão"],
    ["Fechar o semestre", "Rodar de novo com as NFs até 31/12 antes de efetivar (hoje: 108 de 184 dias)"],
  ];
  passos.forEach(([t, d], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = M + col * 6.15, y = 1.65 + row * 1.6;
    s.addShape(pres.shapes.OVAL, { x, y, w: 0.7, h: 0.7, fill: { color: AMARELO }, line: { color: AMARELO } });
    txt(s, String(i + 1), { x, y, w: 0.7, h: 0.7, fontSize: 22, bold: true, color: NAVY, align: "center", valign: "middle" });
    txt(s, t, { x: x + 0.9, y: y - 0.02, w: 5.0, h: 0.4, fontSize: 18, bold: true, color: AMARELO });
    txt(s, d, { x: x + 0.9, y: y + 0.4, w: 5.0, h: 0.85, fontSize: 14, color: BRANCO, valign: "top" });
  });
}

pres.writeFile({ fileName: path.join(SAIDA, "Reestruturacao_RAP_2S2026.pptx") }).then((f) => console.log(f));
