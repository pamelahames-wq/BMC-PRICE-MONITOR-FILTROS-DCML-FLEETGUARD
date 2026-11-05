// Modelo de carta e e-mail de encerramento da homologação RAP (mala direta com a aba "Encerramento" de
// saida/rap/Acoes_RAP_2S2026.xlsx). Campos «ENTRE_ASPAS» = colunas da planilha; [colchetes amarelos] = preencher.
// Uso: node rap/carta_encerramento.js  →  saida/rap/Encerramento_RAP_modelo.docx
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, LevelFormat, PageBreak, HeadingLevel, ShadingType,
} = require("docx");

const NAVY = "12285A", CINZA = "5F6B7A";
const AMARELO = { type: ShadingType.CLEAR, color: "auto", fill: "FFF200" };
const FONT = "Arial";

// texto com «CAMPOS» em negrito e [PREENCHER] destacado em amarelo
function runs(texto, base = {}) {
  return texto.split(/(«[^»]+»|\[[^\]]+\])/).filter(Boolean).map((t) => {
    if (t.startsWith("«")) return new TextRun({ text: t, bold: true, color: NAVY, ...base });
    if (t.startsWith("[")) return new TextRun({ text: t, shading: AMARELO, ...base });
    return new TextRun({ text: t, ...base });
  });
}
const p = (texto, o = {}) => new Paragraph({ children: runs(texto, o.run), spacing: { after: 160, line: 300 }, ...o.par });
const item = (texto) => new Paragraph({ children: runs(texto), numbering: { reference: "itens", level: 0 }, spacing: { after: 100, line: 280 } });
const titulo = (texto) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(texto)], spacing: { before: 120, after: 200 } });
const nota = (texto) => new Paragraph({ children: [new TextRun({ text: texto, italics: true, color: CINZA, size: 18 })], spacing: { after: 120 } });

const carta = [
  new Paragraph({ children: [new TextRun({ text: "BMC | HYUNDAI", bold: true, color: NAVY, size: 28 })], spacing: { after: 40 } }),
  new Paragraph({ children: [new TextRun({ text: "Peças — Revendas Autorizadas de Peças (RAP)", color: CINZA, size: 18 })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 6 } }, spacing: { after: 360 } }),
  p("[Cidade], [data de envio]", { par: { alignment: AlignmentType.RIGHT } }),
  p("À"),
  new Paragraph({ children: runs("«RAZAO_SOCIAL»"), spacing: { after: 0 } }),
  p("CNPJ «CNPJ»"),
  p("Ref.: Encerramento da homologação como Revenda Autorizada de Peças (RAP) — contrato [nº do contrato]", { run: { bold: true } }),
  p("Prezados,"),
  p("A homologação como Revenda Autorizada de Peças BMC Hyundai garante descontos especiais sobre a tabela de peças, de acordo com a categoria da revenda. Para manter essa condição, a política RAP prevê um faturamento mínimo por semestre:"),
  item("Standard: R$ 5.000,00 por semestre;"),
  item("Ouro: R$ 40.000,00 por semestre;"),
  item("Diamante: R$ 100.000,00 por semestre."),
  p("No período de 1º de julho a [data de corte] de 2026, não houve faturamento de peças em nome da «RAZAO_SOCIAL», que hoje está na categoria «CATEGORIA». Por isso, comunicamos o encerramento da homologação RAP a partir de [data de efetivação], conforme [cláusula X do contrato / prazo de aviso prévio de N dias]."),
  p("A partir dessa data:"),
  item("o desconto de revenda RAP deixa de ser aplicado nos novos pedidos;"),
  item("pedidos já faturados ou confirmados até [data de efetivação] seguem as condições atuais;"),
  item("a empresa continua podendo comprar peças BMC Hyundai nas condições de Balcão;"),
  item("uma nova homologação poderá ser solicitada no futuro, mediante análise comercial [e critérios vigentes]."),
  p("Agradecemos a parceria até aqui. Se houver interesse em retomar as compras, ou qualquer dúvida sobre este comunicado, fale com «CONSULTOR», seu consultor de peças, ou com a nossa equipe pelo [telefone/e-mail]."),
  p("Atenciosamente,", { par: { spacing: { before: 240, after: 600 } } }),
  new Paragraph({ children: [new TextRun({ text: "[Nome do responsável]", shading: AMARELO })], spacing: { after: 0 } }),
  new Paragraph({ children: [new TextRun({ text: "[Cargo] — BMC Hyundai", shading: AMARELO })] }),
];

const email = [
  new Paragraph({ children: [new PageBreak()] }),
  titulo("Corpo do e-mail"),
  nota("Para: «E-MAIL» · Anexo: carta de encerramento em PDF · Envio só após aprovação da Pam e do Alexandre."),
  p("Assunto: Encerramento da homologação RAP — «RAZAO_SOCIAL»", { run: { bold: true } }),
  p("Olá, tudo bem?"),
  p("Escrevemos para comunicar uma mudança na parceria da «RAZAO_SOCIAL» com a BMC Hyundai."),
  p("Para manter a homologação como Revenda Autorizada de Peças (RAP), a política prevê um faturamento mínimo por semestre (Standard R$ 5 mil, Ouro R$ 40 mil, Diamante R$ 100 mil). Como não houve compras de peças no 2º semestre de 2026, a homologação RAP será encerrada a partir de [data de efetivação]. Os detalhes estão na carta em anexo."),
  p("O que muda: o desconto RAP deixa de valer nos novos pedidos a partir dessa data. A «RAZAO_SOCIAL» continua podendo comprar peças BMC Hyundai nas condições de Balcão, e as portas ficam abertas para uma nova homologação no futuro."),
  p("Se quiser conversar sobre retomar as compras, o «CONSULTOR» está à disposição pelo [telefone/WhatsApp do consultor]."),
  p("Obrigado pela parceria."),
  p("Equipe de Peças — BMC Hyundai"),
  titulo("Antes de enviar"),
  item("Preencher os campos em amarelo (data de corte, data de efetivação, cláusula ou prazo de aviso do contrato, assinatura e contatos) e validar o texto com o jurídico."),
  item("Revisar na planilha as RAPs marcadas em “Revisar antes?”: as que compraram no 1º semestre (candidatas a resgate em vez de encerramento), as sem e-mail e as filiais com e-mail repetido."),
  item("Mala direta no Word: Correspondências → Selecionar Destinatários → Usar Lista Existente → Acoes_RAP_2S2026.xlsx, aba “Encerramento (385)” (cabeçalho na linha 4). Trocar cada «CAMPO» pelo campo de mesclagem de mesmo nome."),
  item("Registrar a data de envio na coluna “Enviado em”."),
];

const doc = new Document({
  creator: "BMC Hyundai — Peças",
  title: "Encerramento da homologação RAP — modelo",
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [{ id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 28, bold: true, color: NAVY, font: FONT }, paragraph: { spacing: { before: 240, after: 200 }, outlineLevel: 0 } }],
  },
  numbering: { config: [{ reference: "itens", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [{ properties: { page: { margin: { top: 1300, bottom: 1300, left: 1400, right: 1400 } } }, children: [...carta, ...email] }],
});

const saida = path.join(__dirname, "..", "saida", "rap", "Encerramento_RAP_modelo.docx");
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(saida, b); console.log(saida); });
