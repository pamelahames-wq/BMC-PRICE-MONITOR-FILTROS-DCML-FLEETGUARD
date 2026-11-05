// Apresentação da curva ABC 2026 a partir de saida/rap/abc.json (gerado por `python -m rap.abc`).
// Uso: node rap/abc_deck.js   (precisa do pacote pptxgenjs; NODE_PATH pode apontar para onde ele estiver)
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

const SAIDA = path.join(__dirname, "..", "saida", "rap");
const D = JSON.parse(fs.readFileSync(path.join(SAIDA, "abc.json"), "utf8"));
const P = D.pecas, R = D.raps;

const NAVY = "12285A", AMARELO = "F5B800", VERDE = "5E9B1E", CINZA = "5F6B7A";
const CLARO = "EEF2F8", BRANCO = "FFFFFF", TEXTO = "1D2633", VERMELHO = "B3261E";
const COR = { A: NAVY, B: AMARELO, C: "9AA5B4" };
const FONT = "Arial";
const W = 13.33, M = 0.6;

const nf0 = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 });
const rs = (v) => "R$ " + nf0.format(Math.round(v));
const mi = (v) => "R$ " + (v / 1e6).toFixed(1).replace(".", ",") + " mi";
const pct = (v, d = 0) => (v == null ? "s/ custo" : (v * 100).toFixed(d).replace(".", ",") + "%");
const corMargem = (m) => (m == null ? CINZA : m < 0.10 ? VERMELHO : m < 0.25 ? "C77700" : VERDE);
const titulo = (s) => s.toLowerCase().replace(/(^|[\s,(])\S/g, (c) => c.toUpperCase());

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "BMC Hyundai — Peças";
pres.title = "Curva ABC 2026 — Peças e RAPs";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
const RODAPE = `Fonte: Sankhya (f_Notas, f_Fat_Clientes, d_Clientes, d_Produtos, f_CustoUnitarioProd) · NFs de ${D.periodo}`;

pres.defineSlideMaster({
  title: "ESCURO", background: { color: NAVY },
  objects: [{ text: { text: "BMC  |  HYUNDAI", options: { x: M, y: 6.75, w: 4, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: BRANCO, charSpacing: 2 } } }],
});
pres.defineSlideMaster({
  title: "CONTEUDO", background: { color: BRANCO }, margin: [0.5, 0.6, 0.7, 0.6],
  objects: [
    { text: { text: RODAPE, options: { x: M, y: 7.0, w: 9.5, h: 0.3, fontFace: FONT, fontSize: 9, color: CINZA } } },
    { text: { text: "BMC | HYUNDAI", options: { x: W - M - 2.2, y: 7.0, w: 1.7, h: 0.3, fontFace: FONT, fontSize: 9, bold: true, color: NAVY, align: "right" } } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.35, w: W - 2 * M, h: 0.75, fontFace: FONT, fontSize: 30, bold: true, color: NAVY, valign: "middle", align: "left" }, text: "" } },
    { placeholder: { options: { name: "sub", type: "body", x: M, y: 1.05, w: W - 2 * M, h: 0.4, fontFace: FONT, fontSize: 14, color: CINZA, valign: "top", align: "left" }, text: "" } },
  ],
  slideNumber: { x: W - M - 0.45, y: 7.0, w: 0.45, h: 0.3, fontFace: FONT, fontSize: 9, color: CINZA, align: "right" },
});

let sec = "";
function slide(secao, t, sub) {
  if (secao !== sec) { pres.addSection({ title: secao }); sec = secao; }
  const s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: secao });
  s.addText(t, { placeholder: "title" });
  if (sub) s.addText(sub, { placeholder: "sub" });
  return s;
}
const card = (s, x, y, w, h, fill, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08, objectName: name });
const txt = (s, text, o) => s.addText(text, Object.assign({ fontFace: FONT, color: TEXTO, isTextBox: true, margin: 0 }, o));
const tabela = (s, l, o) => s.addTable(l, Object.assign({ fontFace: FONT, fontSize: 11, color: TEXTO, border: { type: "solid", pt: 0.5, color: "D5DBE5" }, valign: "middle", margin: [0.03, 0.06, 0.03, 0.06] }, o));
const H = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true, color: BRANCO, fill: { color: NAVY }, align: "center" }, o) });
const C = (t, o = {}) => ({ text: String(t), options: o });

function cardsClasse(s, y, dados, tot, rot, extra) {
  ["A", "B", "C"].forEach((k, i) => {
    const x = M + i * 4.1, d = dados[k];
    card(s, x, y, 3.8, 1.95, CLARO, `classe-${rot}-${k}`);
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: y + 0.25, w: 0.55, h: 0.55, fill: { color: COR[k] }, line: { color: COR[k] } });
    txt(s, k, { x: x + 0.25, y: y + 0.25, w: 0.55, h: 0.55, fontSize: 20, bold: true, color: k === "B" ? NAVY : BRANCO, align: "center", valign: "middle" });
    txt(s, `${nf0.format(d.n)} ${rot} · ${pct(d.n / tot.n)}`, { x: x + 0.95, y: y + 0.22, w: 2.75, h: 0.32, fontSize: 13, bold: true, color: NAVY });
    txt(s, `${pct(d.v / tot.v)} do faturamento`, { x: x + 0.95, y: y + 0.55, w: 2.75, h: 0.3, fontSize: 12, color: CINZA });
    txt(s, mi(d.v), { x: x + 0.25, y: y + 0.95, w: 3.3, h: 0.5, fontSize: 24, bold: true, color: TEXTO });
    if (extra) txt(s, extra(k), { x: x + 0.25, y: y + 1.5, w: 3.4, h: 0.3, fontSize: 11, color: CINZA });
  });
}

// 1 — Capa
pres.addSection({ title: "Abertura" }); sec = "Abertura";
{
  const s = pres.addSlide({ masterName: "ESCURO", sectionTitle: "Abertura" });
  txt(s, "CURVA ABC 2026", { x: M, y: 1.7, w: 11, h: 0.6, fontSize: 20, bold: true, color: AMARELO, charSpacing: 4 });
  txt(s, "Peças e Revendas RAP", { x: M, y: 2.3, w: 11, h: 1.2, fontSize: 54, bold: true, color: BRANCO });
  txt(s, `NFs de ${D.periodo} · ${nf0.format(P.total_n)} peças vendidas · ${R.total_n} RAPs com faturamento`, { x: M, y: 3.7, w: 11.5, h: 0.5, fontSize: 18, color: "C9D4EA" });
  txt(s, "Classe A = itens que somam os primeiros 80% do faturamento · B = de 80% a 95% · C = os últimos 5%", { x: M, y: 4.3, w: 11.5, h: 0.5, fontSize: 14, italic: true, color: "C9D4EA" });
}

// 2 — Peças: visão geral
{
  const s = slide("Peças", "Peças: 16% dos itens fazem 80% do faturamento", `${nf0.format(P.total_n)} peças vendidas em 2026 · ${mi(P.total_v)} em NFs de Balcão + RAP (vendas − devoluções)`);
  cardsClasse(s, 1.6, P.classes, { n: P.total_n, v: P.total_v }, "peças", (k) => `margem ${pct(P.margem[k], 1)} · ${pct(P.rap[k])} vendido para RAP`);
  const labels = P.curva.map((_, i) => (i % 2 === 0 ? `${i * 5}%` : ""));
  const vals = P.curva.map((v, i) => (i === 0 ? 0 : Math.round(v * 1000) / 10));
  s.addChart(pres.charts.LINE, [{ name: "% do faturamento acumulado", labels, values: vals }], {
    x: M, y: 3.75, w: 7.6, h: 3.1, chartColors: [NAVY], lineSize: 3, lineDataSymbol: "none",
    showTitle: true, title: "Faturamento acumulado × % das peças (da mais vendida para a menos)", titleFontFace: FONT, titleFontSize: 12, titleColor: TEXTO,
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 9, catAxisLabelColor: CINZA, catAxisLabelFrequency: 1, catAxisTitle: "% das peças", showCatAxisTitle: true, catAxisTitleFontSize: 10, catAxisTitleColor: CINZA,
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 9, valAxisLabelColor: CINZA, valAxisMaxVal: 100, valAxisMinVal: 0, valAxisLabelFormatCode: '0"%"',
    valGridLine: { color: "E1E6EE", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false,
  });
  card(s, 8.6, 3.85, 4.13, 2.95, CLARO, "leitura-pecas");
  txt(s, "Leitura", { x: 8.85, y: 4.0, w: 3.7, h: 0.35, fontSize: 15, bold: true, color: NAVY });
  txt(s, [
    { text: `As ${nf0.format(P.classes.A.n)} peças A são onde falta de estoque e erro de preço mais custam.`, options: { bullet: true, breakLine: true } },
    { text: `A margem da classe A (${pct(P.margem.A, 1)}) é menor que a de B e C (~41%): o giro está nas peças caras de margem menor.`, options: { bullet: true, breakLine: true } },
    { text: `${nf0.format(P.classes.C.n)} peças C (58% dos itens) somam só 5% do faturamento.`, options: { bullet: true } },
  ], { x: 8.85, y: 4.45, w: 3.7, h: 2.25, fontSize: 12, paraSpaceAfter: 6, valign: "top" });
}

// 3 — Peças A: top 15
{
  const s = slide("Peças", "As 15 peças que mais faturam em 2026", "Todas classe A · margem = (NF − impostos − custo) ÷ NF");
  const l = [[H("#"), H("PN", { align: "left" }), H("Descrição", { align: "left" }), H("Grupo", { align: "left" }), H("Faturamento"), H("% RAP"), H("Margem")]];
  P.top.forEach((x, i) => l.push([C(i + 1, { align: "center" }), C(x.pn), C(x.desc.slice(0, 32)), C(titulo(x.grupo || "").slice(0, 30)), C(rs(x.nf), { align: "right" }), C(pct(x.pct_rap), { align: "center" }), C(pct(x.m, 1), { align: "center", bold: true, color: corMargem(x.m) })]));
  tabela(s, l, { x: M, y: 1.55, w: W - 2 * M, colW: [0.5, 1.75, 3.6, 3.0, 1.45, 0.9, 0.93], rowH: 0.31, fontSize: 10 });
}

// 4 — Peças A por grupo
{
  const s = slide("Peças", "Peças classe A: grupos que mais faturam", "Faturamento 2026 das peças A em cada grupo · nº de peças A e margem do grupo ao lado");
  const g = P.grupos_a;
  s.addChart(pres.charts.BAR, [{ name: "Faturamento", labels: g.map((x) => titulo(x.g).slice(0, 34)), values: g.map((x) => Math.round(x.nf)) }], {
    x: M, y: 1.6, w: 6.9, h: 5.2, barDir: "bar", chartColors: [NAVY], catAxisOrientation: "maxMin",
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 10, catAxisLabelColor: TEXTO, valAxisHidden: true,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
  });
  const l = [[H("Grupo", { align: "left" }), H("Faturamento"), H("Peças A"), H("Margem")]];
  g.forEach((x) => l.push([C(titulo(x.g).slice(0, 26), { fontSize: 10 }), C(rs(x.nf), { align: "right", fontSize: 10 }), C(x.n, { align: "center" }), C(pct(x.m, 1), { align: "center", bold: true, color: corMargem(x.m) })]));
  tabela(s, l, { x: 7.7, y: 1.7, w: 5.03, colW: [2.2, 1.33, 0.75, 0.75], rowH: 0.42, fontSize: 11 });
  txt(s, "Bombas hidráulicas e peças internas de bombas estão entre as que mais faturam com as menores margens.", { x: 7.7, y: 5.8, w: 5.03, h: 0.9, fontSize: 11, italic: true, color: CINZA, valign: "top" });
}

// 5 — RAPs: visão geral
{
  const s = slide("RAPs", `RAPs: 30 revendas fazem 80% do faturamento RAP`, `${R.total_n} RAPs faturaram em 2026 (${mi(R.total_v)}) · ${R.sem_compra} das 545 não tiveram faturamento líquido no ano`);
  cardsClasse(s, 1.6, R.classes, { n: R.total_n, v: R.total_v }, "RAPs", null);
  txt(s, "Classe ABC × categoria atual (nº de RAPs)", { x: M, y: 3.8, w: 7, h: 0.35, fontSize: 14, bold: true, color: NAVY });
  const cats = ["Diamante", "Ouro", "Standard"];
  const l = [[H("Classe"), ...cats.map((c) => H(c)), H("Total")]];
  ["A", "B", "C", "Sem compra"].forEach((k) => {
    const v = cats.map((c) => R.cat[k][c]);
    l.push([C(k === "Sem compra" ? "Sem compra 2026" : k, { bold: true, align: "center" }), ...v.map((x) => C(x, { align: "center" })), C(v.reduce((a, b) => a + b, 0), { align: "center", bold: true })]);
  });
  tabela(s, l, { x: M, y: 4.25, w: 6.6, colW: [1.6, 1.25, 1.25, 1.25, 1.25], rowH: 0.45, fontSize: 12 });
  card(s, 7.6, 3.85, 5.13, 2.95, CLARO, "leitura-raps");
  txt(s, "Leitura", { x: 7.85, y: 4.0, w: 4.7, h: 0.35, fontSize: 15, bold: true, color: NAVY });
  txt(s, [
    { text: "Toda a classe A é Diamante, mas só 30 dos 214 Diamante chegam lá.", options: { bullet: true, breakLine: true } },
    { text: `${R.cat["Sem compra"].Diamante} RAPs Diamante não faturaram nada em 2026 e seguem com o maior desconto.`, options: { bullet: true, breakLine: true } },
    { text: `${R.a_rebaixar} RAPs classe A estão em REBAIXAR no 2º semestre e ${R.a_desh} em DESHOMOLOGAR: valem contato antes da decisão.`, options: { bullet: true } },
  ], { x: 7.85, y: 4.45, w: 4.7, h: 2.25, fontSize: 12, paraSpaceAfter: 6, valign: "top" });
}

// 6 — RAPs A: top 12
{
  const s = slide("RAPs", "As 12 RAPs que mais faturam em 2026", "Classe A · faturamento de jan a 16/10/2026 · balde da reestruturação do 2º semestre");
  const corBalde = { MANTER: VERDE, INCENTIVAR: "2F7FC1", REBAIXAR: "E07B00", DESHOMOLOGAR: VERMELHO };
  const l = [[H("#"), H("RAP", { align: "left" }), H("Consultor"), H("Categoria"), H("Faturamento 2026"), H("% do total RAP"), H("Balde 2º sem")]];
  R.top.forEach((x, i) => l.push([C(i + 1, { align: "center" }), C(x.nome.slice(0, 40)), C(x.cons, { align: "center" }), C(x.cat, { align: "center" }), C(rs(x.v), { align: "right" }), C(pct(x.v / R.total_v, 1), { align: "center" }), C(x.balde, { align: "center", bold: true, color: corBalde[x.balde] || TEXTO })]));
  tabela(s, l, { x: M, y: 1.55, w: W - 2 * M, colW: [0.5, 4.3, 1.5, 1.3, 1.8, 1.4, 1.33], rowH: 0.38, fontSize: 11 });
}

// 7 — Próximos passos
pres.addSection({ title: "Encerramento" });
{
  const s = pres.addSlide({ masterName: "ESCURO", sectionTitle: "Encerramento" });
  txt(s, "Como usar a curva ABC", { x: M, y: 0.4, w: 11, h: 0.9, fontSize: 36, bold: true, color: BRANCO });
  const passos = [
    ["Estoque", `Priorizar disponibilidade das ${nf0.format(P.classes.A.n)} peças A; revisar estoque parado das peças C`],
    ["Preço", "Conferir margem das peças A de grupo com margem baixa (bombas, peças internas) antes de qualquer campanha"],
    ["RAPs A", "Acompanhamento próximo das 30 RAPs A: concentram 80% do faturamento RAP"],
    ["Revisar baldes", `Antes de rebaixar ou encerrar, olhar as ${R.a_rebaixar + R.a_desh} RAPs classe A e as ${R.b_desh} classe B que caíram na régua do 2º semestre`],
  ];
  passos.forEach(([t, d], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = M + col * 6.15, y = 1.7 + row * 1.9;
    s.addShape(pres.shapes.OVAL, { x, y, w: 0.7, h: 0.7, fill: { color: AMARELO }, line: { color: AMARELO } });
    txt(s, String(i + 1), { x, y, w: 0.7, h: 0.7, fontSize: 22, bold: true, color: NAVY, align: "center", valign: "middle" });
    txt(s, t, { x: x + 0.9, y: y - 0.02, w: 5.0, h: 0.4, fontSize: 18, bold: true, color: AMARELO });
    txt(s, d, { x: x + 0.9, y: y + 0.4, w: 5.0, h: 1.0, fontSize: 14, color: BRANCO, valign: "top" });
  });
}

pres.writeFile({ fileName: path.join(SAIDA, "Curva_ABC_2026.pptx") }).then((f) => console.log(f));
