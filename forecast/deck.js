const pptxgen = require("pptxgenjs");
const fs = require("fs");
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");
const S = process.env.FORECAST_DIR || require("path").join(__dirname, "..", "cache", "forecast");
const d = JSON.parse(fs.readFileSync(`${S}/out/data.json`, "utf8"));
const OUT = `${S}/out/Forecast_Curva_ABC_Pecas_BMC_Hyundai.pptx`;

const HEX = { navy: "002C5F", sky: "00AAD2", red: "C8102E", amber: "E89B00", green: "2E8B57", slate: "6B7B93",
              ink: "1A1F2B", mist: "EEF3F8", line: "D5DDE8", white: "FFFFFF" };
const THEME = { name: "BMC Hyundai", headFontFace: "Arial", bodyFontFace: "Arial",
  colors: { dk1: HEX.ink, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.mist, accent1: HEX.navy, accent2: HEX.sky,
            accent3: HEX.red, accent4: HEX.amber, accent5: HEX.slate, accent6: HEX.green, hlink: HEX.sky, folHlink: HEX.slate } };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";            // 13.33 x 7.5
pres.theme = { headFontFace: "Arial", bodyFontFace: "Arial" };
pres.title = "Forecast de peças – Curva ABC e ruptura de estoque";
pres.author = "BMC Hyundai · Pós-Vendas Peças";
const C = pres.SchemeColor;
const W = 13.333, M = 0.6;
const fmt = (n) => Math.round(n).toLocaleString("pt-BR");
const brl = (n) => n >= 1e6 ? `R$ ${(n / 1e6).toFixed(1).replace(".", ",")} mi` : `R$ ${fmt(n / 1000)} mil`;
const FOOT = "BMC Hyundai · Pós-Vendas Peças  |  Base Sankhya 30/10/2026 · vendas jan–out/2026 · SAB";

pres.defineSlideMaster({ title: "CAPA", background: { color: HEX.navy },
  objects: [{ placeholder: { options: { name: "title", type: "title", x: M, y: 2.3, w: 10.5, h: 1.9, fontSize: 40, bold: true, color: C.background1, valign: "top", align: "left" }, text: "" } }] });
pres.defineSlideMaster({ title: "CONTEUDO", background: { color: HEX.white }, margin: [0.5, 0.6, 0.7, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.35, w: W - 2 * M, h: 0.85, fontSize: 28, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: M, y: 7.0, w: 10, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: W - M - 0.5, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: HEX.slate, align: "right" } });

const box = (s, x, y, w, h, fill, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE,
  { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08, objectName: name });
const txt = (s, t, o) => s.addText(t, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top" }, o));

// ---------- 1. Capa ----------
pres.addSection({ title: "Abertura" });
let s = pres.addSlide({ masterName: "CAPA", sectionTitle: "Abertura" });
txt(s, "BMC HYUNDAI  ·  PÓS-VENDAS PEÇAS", { x: M, y: 1.5, w: 9, h: 0.4, fontSize: 14, bold: true, color: C.accent2, charSpacing: 2 });
s.addText("Forecast de peças: Curva ABC e ruptura de estoque", { placeholder: "title" });
txt(s, "Peças sem estoque por família de máquina e o que precisamos que a fábrica traga", { x: M, y: 4.3, w: 10, h: 0.5, fontSize: 18, color: C.background1 });
txt(s, "Base Sankhya 30/10/2026 · vendas Balcão + RAP jan–out/2026 · parque SAB · pacote de importação RAP", { x: M, y: 6.4, w: 11, h: 0.4, fontSize: 12, color: C.accent2 });

// ---------- 2. Resumo ----------
pres.addSection({ title: "Diagnóstico" });
const P = d.prio, n57 = P.P1 + P.P2, n127 = P.P1 + P.P2 + P.P3;
s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Diagnóstico" });
s.addText(`${n57} peças recorrentes têm menos de 1 mês de estoque`, { placeholder: "title" });
const kpis = [
  [String(P.P1), "Curva A com menos de 1 mês de cobertura (P1)", HEX.red],
  [String(P.P2), "Curva B com menos de 1 mês de cobertura (P2)", HEX.amber],
  [String(P.P3), "Curva A com 1 a 3 meses: vão faltar no prazo de importação (P3)", HEX.navy],
  [brl(d.p12_venda_mes), "de venda média por mês nas peças P1 e P2", HEX.slate],
];
kpis.forEach((k, i) => {
  const x = M + i * 3.07;
  box(s, x, 1.45, 2.85, 2.1, HEX.mist, `kpi${i}`);
  txt(s, k[0], { x: x + 0.25, y: 1.65, w: 2.4, h: 0.9, fontSize: k[0].length > 6 ? 30 : 44, bold: true, color: k[2] });
  txt(s, k[1], { x: x + 0.25, y: 2.6, w: 2.4, h: 0.85, fontSize: 13, color: C.text1 });
});
const novas = Object.values(d.parque).reduce((a, b) => a + b.novas_2026, 0);
const esc = d.forecast[0].pns;
const msgs = [
  [`Escavadeira concentra ${esc} dos ${n127} itens críticos.`, " Hidráulico (bombas, motores de giro e translação) e kits de vedação lideram."],
  [`Só ${d.p_no_pacote} dos ${n127} itens críticos estão no pacote RAP.`, ` Os outros ${d.p_fora_pacote} precisam de pedido próprio.`],
  [`O parque cresce: ${fmt(novas)} máquinas novas entraram no SAB em 2026.`, " A demanda de peças de reposição acompanha, sobretudo retroescavadeira (+90 sobre 110)."],
  [`${d.alto_giro_zero} peças de alto giro estão zeradas hoje`, " (vendem em 6 ou mais meses do ano e não têm estoque na fábrica nem nas filiais)."],
];
msgs.forEach((m, i) => {
  const y = 3.95 + i * 0.72;
  s.addShape(pres.shapes.OVAL, { x: M, y: y + 0.07, w: 0.32, h: 0.32, fill: { color: HEX.navy }, line: { color: HEX.navy }, objectName: `n${i}` });
  txt(s, String(i + 1), { x: M, y: y + 0.07, w: 0.32, h: 0.32, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
  txt(s, [{ text: m[0], options: { bold: true } }, { text: m[1] }], { x: M + 0.5, y, w: 11.5, h: 0.6, fontSize: 15, valign: "middle" });
});

// ---------- 3. ABC x disponibilidade ----------
s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Diagnóstico" });
const A = d.abc;
const riscoA = A.A["Atenção"] + A.A["Crítico"] + A.A["Sem estoque"];
const recA = riscoA + A.A.OK;
s.addText(`Curva A: ${riscoA} de ${recA} peças recorrentes com menos de 3 meses`, { placeholder: "title" });
const cats = ["Curva A", "Curva B", "Curva C"];
const ser = [["Sem estoque", HEX.red], ["Crítico (<1 mês)", HEX.amber], ["Atenção (1–3 meses)", HEX.sky]];
const keys = ["Sem estoque", "Crítico", "Atenção"];
s.addChart(pres.charts.BAR, ser.map((sr, i) => ({ name: sr[0], labels: cats, values: ["A", "B", "C"].map(c => A[c][keys[i]]) })), {
  x: M, y: 1.4, w: 7.6, h: 5.3, barDir: "bar", barGrouping: "clustered", barGapWidthPct: 60, chartColors: ser.map(x => x[1]),
  showLegend: true, legendPos: "b", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.ink,
  showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelColor: HEX.ink, dataLabelFontFace: "+mn-lt",
  catAxisLabelColor: HEX.ink, catAxisLabelFontSize: 13, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
  valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
  showTitle: true, title: "Peças recorrentes (venda em 4+ meses) com risco de ruptura", titleFontSize: 13, titleColor: HEX.ink, titleFontFace: "+mn-lt",
});
box(s, 8.65, 1.5, 4.08, 5.1, HEX.mist, "abcnote");
const pct = Math.round(100 * d.abc_valor.A / (d.abc_valor.A + d.abc_valor.B + d.abc_valor.C));
txt(s, [
  { text: "Como ler", options: { bold: true, fontSize: 16, color: C.text2, breakLine: true } },
  { text: `Curva A = ${fmt(d.abc_pns.A)} PNs que somam ${pct}% do faturamento de peças Hyundai em 2026.`, options: { breakLine: true, paraSpaceAfter: 8 } },
  { text: "Cobertura = estoque da fábrica (HHIBP) + filiais BMC, dividido pela venda média mensal.", options: { breakLine: true, paraSpaceAfter: 8 } },
  { text: "Abaixo de 3 meses a peça falta antes de uma importação chegar.", options: { bold: true, breakLine: true, paraSpaceAfter: 8 } },
  { text: `Com estoque para 3+ meses: ${A.A.OK} na curva A, ${A.B.OK} na B e ${A.C.OK} na C.` },
], { x: 8.9, y: 1.75, w: 3.6, h: 4.7, fontSize: 14 });

// ---------- 4. Matriz família x categoria ----------
s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Diagnóstico" });
s.addText("Hidráulico e vedação de escavadeira concentram a ruptura", { placeholder: "title" });
const mx = d.matriz, colsF = mx.colunas, maxv = Math.max(...mx.valores.flat());
const shade = (v) => v === 0 ? HEX.white : v >= 10 ? HEX.red : v >= 5 ? HEX.amber : "F6D9A8";
const hdr = [{ text: "Categoria de peça", options: { bold: true, color: HEX.white, fill: { color: HEX.navy } } }]
  .concat(colsF.map(c => ({ text: ({ "Pá carregadeira": "Pá carreg.", "Retroescavadeira": "Retro" })[c] || c, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, align: "center" } })))
  .concat([{ text: "Total", options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, align: "center" } }]);
const rowsM = [hdr];
const tot = colsF.map(() => 0);
mx.linhas.forEach((l, i) => {
  const vals = mx.valores[i];
  if (vals.reduce((a, b) => a + b, 0) === 0) return;
  vals.forEach((v, j) => tot[j] += v);
  rowsM.push([{ text: l, options: { color: HEX.ink } }].concat(vals.map(v => ({ text: v ? String(v) : "–",
    options: { align: "center", bold: v >= 5, color: v >= 10 ? HEX.white : HEX.ink, fill: { color: shade(v) } } })))
    .concat([{ text: String(vals.reduce((a, b) => a + b, 0)), options: { align: "center", bold: true, color: HEX.ink } }]));
});
rowsM.push([{ text: "Total", options: { bold: true } }].concat(tot.map(v => ({ text: String(v), options: { align: "center", bold: true } })))
  .concat([{ text: String(tot.reduce((a, b) => a + b, 0)), options: { align: "center", bold: true } }]));
s.addTable(rowsM, { x: M, y: 1.4, w: 9.4, colW: [3.2, 1.3, 1.05, 1.3, 0.85, 0.85, 0.85], fontSize: 12, fontFace: "Arial",
  border: { type: "solid", pt: 0.5, color: HEX.line }, rowH: 0.34, valign: "middle", objectName: "matriz" });
box(s, 10.3, 1.4, 2.43, 5.2, HEX.mist, "matnote");
txt(s, [
  { text: `${n127} itens P1 a P3`, options: { bold: true, fontSize: 16, color: C.text2, breakLine: true } },
  { text: "Contagem de part numbers por família de máquina.", options: { breakLine: true, paraSpaceAfter: 10 } },
  { text: "Família = plano de manutenção oficial do modelo; sem plano, a frota SAB dos clientes que compraram a peça.", options: { breakLine: true, paraSpaceAfter: 10 } },
  { text: "\"Várias\" = peça comum a mais de uma família.", options: { fontSize: 12, color: C.accent5 } },
], { x: 10.48, y: 1.6, w: 2.1, h: 4.9, fontSize: 13 });

// ---------- 5. Tabela P1 ----------
pres.addSection({ title: "Pedido à fábrica" });
s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Pedido à fábrica" });
s.addText("Curva A sem cobertura: o que a fábrica precisa trazer primeiro", { placeholder: "title" });
const H = (t, al) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.navy }, align: al || "center", valign: "middle" } });
const rowsP = [[H("Part number", "left"), H("Peça", "left"), H("Família · modelos", "left"), H("Venda/mês"), H("Meses c/ venda"),
  H("Est. fábrica"), H("Est. BMC"), H("Forecast 90 d"), H("No pacote RAP"), H("ETA fábrica")]];
d.p1.slice(0, 12).forEach((r, i) => {
  const z = { fill: { color: i % 2 ? HEX.mist : HEX.white } };
  const c = (t, o) => ({ text: String(t), options: Object.assign({ align: "center" }, z, o || {}) });
  rowsP.push([c(r.pn, { align: "left", bold: true }), c(r.desc.slice(0, 28), { align: "left" }), c(`${r.fam} · ${r.mod}`, { align: "left", fontSize: 10 }),
    c(String(r.vm).replace(".", ",")), c(`${r.meses}/10`), c(r.fab, { color: r.fab ? HEX.ink : HEX.red, bold: !r.fab }),
    c(r.bmc, { color: r.bmc ? HEX.ink : HEX.red, bold: !r.bmc }), c(r.f90, { bold: true }), c(r.pac ? r.pac : "não"), c("")]);
});
s.addTable(rowsP, { x: M, y: 1.3, w: W - 2 * M, colW: [1.45, 2.55, 2.35, 0.8, 0.8, 0.8, 0.7, 0.85, 0.8, 1.03],
  fontSize: 11, fontFace: "Arial", rowH: 0.34, border: { type: "solid", pt: 0.5, color: HEX.line }, valign: "middle", objectName: "p1" });
txt(s, `12 maiores em venda de ${P.P1} itens P1. Lista completa (P1, P2, P3 e manuais) na planilha de acompanhamento.`,
  { x: M, y: 6.5, w: 11, h: 0.3, fontSize: 11, color: C.accent5 });

// ---------- 6. Inclusões manuais ----------
s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Pedido à fábrica" });
s.addText("Vendem pouco, mas faltam sempre: vidros HL635 e itens HB640C", { placeholder: "title" });
const pq = d.parque;
[["Pá carregadeira HL635", "329", "máquinas no parque SAB · +59 em 2026"], ["Retroescavadeira HB640C", "200", "máquinas no parque SAB · +90 em 2026"]].forEach((k, i) => {
  const y = 1.4 + i * 2.05;
  box(s, M, y, 3.4, 1.85, HEX.navy, `parq${i}`);
  txt(s, k[0], { x: M + 0.25, y: y + 0.18, w: 3, h: 0.35, fontSize: 14, bold: true, color: C.accent2 });
  txt(s, k[1], { x: M + 0.25, y: y + 0.55, w: 3, h: 0.7, fontSize: 40, bold: true, color: C.background1 });
  txt(s, k[2], { x: M + 0.25, y: y + 1.3, w: 3, h: 0.4, fontSize: 12, color: C.background1 });
});
box(s, M, 5.55, 3.4, 1.15, HEX.mist, "hbnote");
txt(s, "Os filtros do plano preventivo da HB640C estão disponíveis na fábrica (199 a 975 un.). O risco está nos itens abaixo, fora do plano.",
  { x: M + 0.2, y: 5.62, w: 3.05, h: 1.05, fontSize: 11, color: C.text1, valign: "middle" });
const rowsMan = [[H("Part number", "left"), H("Peça", "left"), H("Modelo"), H("Venda/mês"), H("Est. fábrica"), H("Est. BMC"), H("Pacote RAP"), H("Pedido sug.")]];
d.manual.filter(r => r.prio === "Manual" || r.pn === "400401-00304").forEach((r, i) => {
  const z = { fill: { color: i % 2 ? HEX.mist : HEX.white } };
  const c = (t, o) => ({ text: String(t), options: Object.assign({ align: "center" }, z, o || {}) });
  const sug = Math.max(r.f90, 2) - r.fab - r.bmc;
  const sugTxt = sug > 0 ? String(sug) : (r.fab ? "ok" : "repor mín.");
  rowsMan.push([c(r.pn, { align: "left", bold: true }), c(r.desc.slice(0, 30), { align: "left" }), c(r.mod.startsWith("HB") ? "HB640C" : "HL635"),
    c(r.vm ? String(r.vm).replace(".", ",") : "–"), c(r.fab, { color: r.fab ? HEX.ink : HEX.red, bold: !r.fab }),
    c(r.bmc, { color: r.bmc ? HEX.ink : HEX.red, bold: !r.bmc }), c(r.pac ? r.pac : "não"), c(sugTxt, { bold: true, fontSize: sug > 0 ? 11 : 10 })]);
});
s.addTable(rowsMan, { x: 4.3, y: 1.4, w: 8.43, colW: [1.4, 2.5, 0.8, 0.75, 0.75, 0.7, 0.73, 0.8],
  fontSize: 11, fontFace: "Arial", rowH: 0.4, border: { type: "solid", pt: 0.5, color: HEX.line }, valign: "middle", objectName: "manual" });
txt(s, "Pedido sugerido: forecast de 90 dias (mínimo 2 un. para itens de baixo giro) menos o estoque atual.",
  { x: 4.3, y: 6.4, w: 8.4, h: 0.3, fontSize: 11, color: C.accent5 });

// ---------- 7. Forecast x parque ----------
pres.addSection({ title: "Forecast" });
s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Forecast" });
s.addText(`Forecast 90 dias: faltam ${fmt(d.falta_total_un)} peças (${brl(d.falta_total_rs)}) nos itens críticos`, { placeholder: "title" });
const F = ["Escavadeira", "Pá carregadeira", "Empilhadeira", "Retroescavadeira"];
s.addChart(pres.charts.BAR, [{ name: "Máquinas novas 2026", labels: F, values: F.map(f => pq[f].novas_2026) }], {
  x: M, y: 1.4, w: 5.2, h: 4.6, barDir: "bar", chartColors: [HEX.navy], showLegend: false,
  showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelColor: HEX.ink, dataLabelFontFace: "+mn-lt",
  catAxisLabelColor: HEX.ink, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
  valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
  showTitle: true, title: "Máquinas novas que entraram no SAB em 2026", titleFontSize: 13, titleColor: HEX.ink, titleFontFace: "+mn-lt" });
const rowsF = [[H("Família", "left"), H("Parque SAB"), H("Itens críticos"), H("Forecast 90 d (un)"), H("Estoque (un)"), H("Falta (un)"), H("Falta (R$)")]];
d.forecast.forEach((r, i) => {
  const z = { fill: { color: i % 2 ? HEX.mist : HEX.white } };
  const c = (t, o) => ({ text: String(t), options: Object.assign({ align: "center" }, z, o || {}) });
  rowsF.push([c(r.fam, { align: "left", bold: true }), c(fmt(pq[r.fam].parque)), c(r.pns), c(fmt(r.f90)), c(fmt(r.estoque)),
    c(fmt(r.falta), { bold: true, color: r.falta ? HEX.red : HEX.ink }), c(r.falta_rs ? brl(r.falta_rs) : "–")]);
});
const v = d.forecast_varias;
rowsF.push([{ text: "Várias famílias", options: { bold: true } }, { text: "–", options: { align: "center" } }, { text: String(v.pns), options: { align: "center" } },
  { text: fmt(v.f90), options: { align: "center" } }, { text: fmt(v.f90 - v.falta), options: { align: "center" } },
  { text: fmt(v.falta), options: { align: "center", bold: true, color: HEX.red } }, { text: "", options: {} }]);
s.addTable(rowsF, { x: 6.1, y: 1.5, w: 6.63, colW: [1.55, 0.85, 0.8, 0.95, 0.85, 0.75, 0.88], fontSize: 11, fontFace: "Arial",
  rowH: 0.45, border: { type: "solid", pt: 0.5, color: HEX.line }, valign: "middle", objectName: "fc" });
box(s, 6.1, 4.55, 6.63, 1.95, HEX.mist, "fcnote");
txt(s, [
  { text: "Como o forecast é calculado", options: { bold: true, color: C.text2, breakLine: true } },
  { text: "Maior entre a venda média de 2026 e a de ago–out, × 3 meses (prazo de importação), × crescimento do parque da família.", options: { breakLine: true, paraSpaceAfter: 6 } },
  { text: "Falta = forecast menos estoque da fábrica e das filiais, só nos itens P1 a P3. Valor a preço médio de venda.", options: {} },
], { x: 6.3, y: 4.7, w: 6.25, h: 1.7, fontSize: 12 });

// ---------- 8. Pedido e acompanhamento ----------
s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Forecast" });
s.addText("O que pedimos à fábrica e como vamos acompanhar", { placeholder: "title" });
const asks = [
  ["1", "Importar", `Programar a importação dos ${n57} itens P1 e P2 na quantidade do forecast de 90 dias, começando pelos ${P.P1} da curva A.`],
  ["2", "Informar ETA", `Dar a data de chegada por part number: ${n127} itens críticos e os 224 do pacote RAP (13 deles estão zerados na fábrica hoje).`],
  ["3", "Estoque mínimo", `Revisar o mínimo da HHIBP para peças recorrentes, considerando as ${fmt(novas)} máquinas novas de 2026 e as inclusões manuais (vidros HL635, HB640C).`],
];
asks.forEach((a, i) => {
  const x = M + i * 4.12;
  box(s, x, 1.45, 3.9, 2.6, HEX.mist, `ask${i}`);
  s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: 1.7, w: 0.6, h: 0.6, fill: { color: HEX.navy }, line: { color: HEX.navy }, objectName: `askn${i}` });
  txt(s, a[0], { x: x + 0.3, y: 1.7, w: 0.6, h: 0.6, fontSize: 20, bold: true, color: C.background1, align: "center", valign: "middle" });
  txt(s, a[1], { x: x + 1.05, y: 1.72, w: 2.7, h: 0.55, fontSize: 20, bold: true, color: C.text2, valign: "middle" });
  txt(s, a[2], { x: x + 0.3, y: 2.5, w: 3.35, h: 1.85, fontSize: 14 });
});
box(s, M, 4.4, W - 2 * M, 1.6, HEX.navy, "rito");
txt(s, [
  { text: "Acompanhamento em cada reunião", options: { bold: true, fontSize: 16, color: C.accent2, breakLine: true } },
  { text: "Planilha única por part number: prioridade, forecast, quantidade pedida, data do pedido, ETA informado e status. ", options: {} },
  { text: "Item sem ETA ou com ETA vencido volta para a pauta da reunião seguinte.", options: { bold: true } },
], { x: M + 0.35, y: 4.6, w: W - 2 * M - 0.7, h: 1.25, fontSize: 15, color: C.background1, valign: "middle" });

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log(OUT);
})();
