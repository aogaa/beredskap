/**
 * Bygger presentasjonsdokumentet for Beredskapsplan Frivilligsentralene Vestre Aker.
 * Format: Microsoft Word (.docx), A4, serif overskrifter, profesjonell profil.
 */

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, PageOrientation,
  LevelFormat, ExternalHyperlink, FootnoteReferenceRun,
  TabStopType, TabStopPosition, TableOfContents, HeadingLevel,
  BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageNumber, PageBreak,
  HorizontalPositionRelativeFrom, HorizontalPositionAlign,
  VerticalPositionRelativeFrom, VerticalPositionAlign,
} = require("docx");

// ============================================================
// DESIGN — solid, alvor, seriøsitet
// ============================================================
const COLOR_PRIMARY = "1A2B47";     // Mørk marineblå
const COLOR_ACCENT = "7B2D2F";      // Mørk burgunder
const COLOR_SECONDARY = "2D4F3D";   // Mørk skogsgrønn
const COLOR_TEXT = "2C3E50";        // Dyp grå-blå
const COLOR_MUTED = "6B7280";       // Lys grå
const COLOR_BG_FACT = "F0EEE6";     // Lys beige — faktaboks
const COLOR_BG_GOAL = "E8E6DC";     // Litt mørkere beige — måltallsboks
const COLOR_BORDER = "C4BDA8";      // Beige border
const COLOR_GRID = "D0CCB8";        // Tabellgrid

const FONT_HEADING = "Cambria";
const FONT_BODY = "Calibri";

// Sidemål A4 (DXA): 11906 × 16838
// Margin 2cm = 1134 DXA, content width = 11906 - 2268 = 9638
const PAGE_WIDTH = 11906;
const PAGE_HEIGHT = 16838;
const MARGIN = 1134;
const CONTENT_WIDTH = PAGE_WIDTH - 2 * MARGIN;

// Borders for tables and boxes
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER };
const accentBorderLeft = { style: BorderStyle.SINGLE, size: 24, color: COLOR_PRIMARY };
const factBorderLeft = { style: BorderStyle.SINGLE, size: 24, color: COLOR_SECONDARY };
const accentBorders = { top: thinBorder, bottom: thinBorder, left: accentBorderLeft, right: thinBorder };
const factBorders = { top: thinBorder, bottom: thinBorder, left: factBorderLeft, right: thinBorder };
const tableBorders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_GRID },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_GRID },
  left: { style: BorderStyle.SINGLE, size: 4, color: COLOR_GRID },
  right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_GRID },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: COLOR_GRID },
  insideVertical: { style: BorderStyle.SINGLE, size: 4, color: COLOR_GRID },
};

// ============================================================
// HJELPEFUNKSJONER
// ============================================================
function P(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, line: 300, ...(opts.spacing || {}) },
    alignment: opts.alignment || AlignmentType.JUSTIFIED,
    children: [new TextRun({
      text,
      font: opts.font || FONT_BODY,
      size: opts.size || 22, // 11pt
      color: opts.color || COLOR_TEXT,
      bold: opts.bold || false,
      italics: opts.italics || false,
    })],
  });
}

// Paragraf med flere runs (mixed formatering)
function PR(runs, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, line: 300, ...(opts.spacing || {}) },
    alignment: opts.alignment || AlignmentType.JUSTIFIED,
    children: runs.map(r => {
      if (typeof r === "string") {
        return new TextRun({ text: r, font: FONT_BODY, size: 22, color: COLOR_TEXT });
      }
      return new TextRun({
        text: r.text,
        font: r.font || FONT_BODY,
        size: r.size || 22,
        color: r.color || COLOR_TEXT,
        bold: r.bold || false,
        italics: r.italics || false,
      });
    }),
  });
}

function H1(text, opts = {}) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 200 },
    pageBreakBefore: opts.pageBreakBefore !== false,
    keepNext: true,
    keepLines: true,
    children: [new TextRun({
      text,
      font: FONT_HEADING,
      size: 36, // 18pt
      bold: true,
      color: COLOR_PRIMARY,
    })],
  });
}

function H2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 360, after: 160 },
    keepNext: true,
    keepLines: true,
    children: [new TextRun({
      text,
      font: FONT_HEADING,
      size: 28, // 14pt
      bold: true,
      color: COLOR_PRIMARY,
    })],
  });
}

function H3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 280, after: 120 },
    keepNext: true,
    keepLines: true,
    children: [new TextRun({
      text,
      font: FONT_HEADING,
      size: 24, // 12pt
      bold: true,
      color: COLOR_PRIMARY,
    })],
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { after: 80, line: 280 },
    children: [new TextRun({
      text,
      font: FONT_BODY,
      size: 22,
      color: COLOR_TEXT,
    })],
  });
}

function bulletR(runs, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { after: 80, line: 280 },
    children: runs.map(r => {
      if (typeof r === "string") {
        return new TextRun({ text: r, font: FONT_BODY, size: 22, color: COLOR_TEXT });
      }
      return new TextRun({
        text: r.text,
        font: r.font || FONT_BODY,
        size: r.size || 22,
        color: r.color || COLOR_TEXT,
        bold: r.bold || false,
        italics: r.italics || false,
      });
    }),
  });
}

function numbered(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "numbers", level },
    spacing: { after: 80, line: 280 },
    children: [new TextRun({
      text,
      font: FONT_BODY,
      size: 22,
      color: COLOR_TEXT,
    })],
  });
}

// Faktaboks med kildehenvisning (valgfri URL gjør kilde klikkbar)
function faktaboks(tittel, tekst, kilde, url) {
  const kildeRun = url
    ? new ExternalHyperlink({
        link: url,
        children: [new TextRun({
          text: "Kilde: " + kilde,
          font: FONT_BODY,
          size: 18,
          italics: true,
          color: COLOR_PRIMARY,
          underline: { type: "single", color: COLOR_PRIMARY },
        })],
      })
    : new TextRun({
        text: "Kilde: " + kilde,
        font: FONT_BODY,
        size: 18,
        italics: true,
        color: COLOR_MUTED,
      });

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: CONTENT_WIDTH, type: WidthType.DXA },
            shading: { fill: COLOR_BG_FACT, type: ShadingType.CLEAR },
            borders: factBorders,
            margins: { top: 200, bottom: 200, left: 280, right: 280 },
            children: [
              new Paragraph({
                spacing: { after: 100 },
                children: [new TextRun({
                  text: "FAKTA  /  " + tittel.toUpperCase(),
                  font: FONT_HEADING,
                  size: 18,
                  bold: true,
                  color: COLOR_SECONDARY,
                })],
              }),
              new Paragraph({
                spacing: { after: 120, line: 280 },
                alignment: AlignmentType.JUSTIFIED,
                children: [new TextRun({
                  text: tekst,
                  font: FONT_BODY,
                  size: 21,
                  color: COLOR_TEXT,
                })],
              }),
              new Paragraph({
                spacing: { after: 0 },
                alignment: AlignmentType.RIGHT,
                children: [kildeRun],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

// Sitatboks — for utdrag fra dokumenter
function sitatboks(sitat, kilde) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: CONTENT_WIDTH, type: WidthType.DXA },
            shading: { fill: "FAFAF7", type: ShadingType.CLEAR },
            borders: accentBorders,
            margins: { top: 200, bottom: 200, left: 360, right: 280 },
            children: [
              new Paragraph({
                spacing: { after: 120, line: 300 },
                children: [new TextRun({
                  text: "“" + sitat + "”",
                  font: FONT_HEADING,
                  size: 24,
                  italics: true,
                  color: COLOR_PRIMARY,
                })],
              }),
              new Paragraph({
                spacing: { after: 0 },
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({
                  text: "— " + kilde,
                  font: FONT_BODY,
                  size: 19,
                  italics: true,
                  color: COLOR_MUTED,
                })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

// Måltallsboks for hvert tiltak
function maltallsboks(maltall, suksesskriterium) {
  const rows = [
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({
          width: { size: CONTENT_WIDTH, type: WidthType.DXA },
          shading: { fill: COLOR_BG_GOAL, type: ShadingType.CLEAR },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
            left: { style: BorderStyle.SINGLE, size: 24, color: COLOR_PRIMARY },
            right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
          },
          margins: { top: 200, bottom: 200, left: 280, right: 280 },
          children: [
            new Paragraph({
              spacing: { after: 120 },
              children: [new TextRun({
                text: "MÅLTALL FØRSTE ÅR",
                font: FONT_HEADING,
                size: 20,
                bold: true,
                color: COLOR_PRIMARY,
              })],
            }),
            ...maltall.map(m => new Paragraph({
              spacing: { after: 80, line: 280 },
              indent: { left: 200, hanging: 200 },
              children: [
                new TextRun({ text: "▸  ", font: FONT_BODY, size: 22, color: COLOR_ACCENT, bold: true }),
                new TextRun({ text: m, font: FONT_BODY, size: 21, color: COLOR_TEXT }),
              ],
            })),
            new Paragraph({
              spacing: { before: 160, after: 0, line: 280 },
              children: [
                new TextRun({
                  text: "Suksesskriterium:  ",
                  font: FONT_HEADING,
                  size: 21,
                  bold: true,
                  color: COLOR_ACCENT,
                }),
                new TextRun({
                  text: suksesskriterium,
                  font: FONT_BODY,
                  size: 21,
                  italics: true,
                  color: COLOR_TEXT,
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  ];
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    rows,
  });
}

// Tiltaks-header med score-bar visualisering
function tiltakHeader(kode, navn, viktighet, kompleksitet, hoyPrioritet) {
  // Score-bars som tekst med farger
  const fullCircle = "●";
  const emptyCircle = "○";
  function scoreBar(score) {
    const filled = Math.round(score);
    let bar = "";
    for (let i = 1; i <= 5; i++) {
      bar += (i <= filled) ? fullCircle : emptyCircle;
    }
    return bar;
  }

  const kodeRuns = [
    new TextRun({
      text: kode,
      font: FONT_HEADING,
      size: 48,
      bold: true,
      color: hoyPrioritet ? COLOR_ACCENT : COLOR_PRIMARY,
    }),
  ];
  if (hoyPrioritet) {
    kodeRuns.push(new TextRun({
      text: "  HØY PRIORITET",
      font: FONT_HEADING,
      size: 20,
      bold: true,
      color: COLOR_ACCENT,
    }));
  }

  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      pageBreakBefore: true,
      spacing: { before: 0, after: 80 },
      keepNext: true,
      keepLines: true,
      children: kodeRuns,
    }),
    new Paragraph({
      spacing: { after: 280 },
      keepNext: true,
      keepLines: true,
      children: [new TextRun({
        text: navn,
        font: FONT_HEADING,
        size: 32,
        bold: true,
        color: COLOR_PRIMARY,
      })],
    }),
    // Score-tabell
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [Math.floor(CONTENT_WIDTH / 2), Math.floor(CONTENT_WIDTH / 2)],
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            new TableCell({
              width: { size: Math.floor(CONTENT_WIDTH / 2), type: WidthType.DXA },
              borders: {
                top: noBorder, bottom: noBorder, left: noBorder,
                right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
              },
              margins: { top: 100, bottom: 100, left: 0, right: 200 },
              children: [
                new Paragraph({
                  spacing: { after: 60 },
                  children: [new TextRun({
                    text: "VIKTIGHET",
                    font: FONT_HEADING, size: 18, bold: true, color: COLOR_MUTED,
                  })],
                }),
                new Paragraph({
                  spacing: { after: 0 },
                  children: [
                    new TextRun({
                      text: scoreBar(viktighet) + "  ",
                      font: FONT_BODY, size: 28, color: COLOR_PRIMARY,
                    }),
                    new TextRun({
                      text: viktighet + " / 5",
                      font: FONT_HEADING, size: 24, bold: true, color: COLOR_PRIMARY,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              width: { size: Math.floor(CONTENT_WIDTH / 2), type: WidthType.DXA },
              borders: {
                top: noBorder, bottom: noBorder,
                left: noBorder, right: noBorder,
              },
              margins: { top: 100, bottom: 100, left: 200, right: 0 },
              children: [
                new Paragraph({
                  spacing: { after: 60 },
                  children: [new TextRun({
                    text: "KOMPLEKSITET",
                    font: FONT_HEADING, size: 18, bold: true, color: COLOR_MUTED,
                  })],
                }),
                new Paragraph({
                  spacing: { after: 0 },
                  children: [
                    new TextRun({
                      text: scoreBar(kompleksitet) + "  ",
                      font: FONT_BODY, size: 28, color: COLOR_PRIMARY,
                    }),
                    new TextRun({
                      text: (Number.isInteger(kompleksitet) ? kompleksitet : kompleksitet.toFixed(1)) + " / 5",
                      font: FONT_HEADING, size: 24, bold: true, color: COLOR_PRIMARY,
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    new Paragraph({ spacing: { after: 240 } }), // litt luft
  ];
}

// Sett inn et bilde med tittel
function bilde(filepath, width, height, tittel = null, undertekst = null) {
  const elems = [];
  if (tittel) {
    elems.push(new Paragraph({
      spacing: { before: 240, after: 120 },
      alignment: AlignmentType.CENTER,
      keepNext: true,
      keepLines: true,
      children: [new TextRun({
        text: tittel,
        font: FONT_HEADING, size: 22, bold: true, color: COLOR_PRIMARY,
      })],
    }));
  }
  elems.push(new Paragraph({
    spacing: { after: 80 },
    alignment: AlignmentType.CENTER,
    keepNext: !!undertekst,  // bare keepNext hvis det er en undertekst som skal holde sammen
    keepLines: true,
    children: [new ImageRun({
      type: "png",
      data: fs.readFileSync(filepath),
      transformation: { width, height },
      altText: {
        title: tittel || "Figur",
        description: tittel || "Figur",
        name: path.basename(filepath, ".png"),
      },
    })],
  }));
  if (undertekst) {
    elems.push(new Paragraph({
      spacing: { after: 240 },
      alignment: AlignmentType.CENTER,
      keepLines: true,
      children: [new TextRun({
        text: undertekst,
        font: FONT_BODY, size: 18, italics: true, color: COLOR_MUTED,
      })],
    }));
  }
  return elems;
}

// ============================================================
// FORSIDE
// ============================================================
function forside() {
  const halfWidth = Math.floor(CONTENT_WIDTH / 2);
  return [
    // Logoene side ved side — to celler uten kantlinjer
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [halfWidth, halfWidth],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: halfWidth, type: WidthType.DXA },
              borders: {
                top: noBorder, bottom: noBorder, left: noBorder, right: noBorder,
              },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 100, bottom: 100, left: 100, right: 200 },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 0 },
                  children: [new ImageRun({
                    type: "png",
                    data: fs.readFileSync("C:/codex/beredskap/presentasjon/figurer/VAFS_med_skygge.png"),
                    transformation: { width: 220, height: 156 },
                    altText: {
                      title: "Vestre Aker Frivilligsentral",
                      description: "Logo for Vestre Aker Frivilligsentral",
                      name: "VAFS_logo",
                    },
                  })],
                }),
              ],
            }),
            new TableCell({
              width: { size: halfWidth, type: WidthType.DXA },
              borders: {
                top: noBorder, bottom: noBorder, left: noBorder, right: noBorder,
              },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 100, bottom: 100, left: 200, right: 100 },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 0 },
                  children: [new ImageRun({
                    type: "png",
                    data: fs.readFileSync("C:/codex/beredskap/presentasjon/figurer/OVFS_med_skygge.png"),
                    transformation: { width: 233, height: 156 },
                    altText: {
                      title: "Oslo Vest Frivilligsentral",
                      description: "Logo for Oslo Vest Frivilligsentral",
                      name: "OVFS_logo",
                    },
                  })],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    // Tomrom
    new Paragraph({ spacing: { before: 2400, after: 0 } }),
    new Paragraph({ spacing: { before: 800, after: 0 } }),

    // Tittel
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 600, after: 240 },
      children: [new TextRun({
        text: "Beredskapsplan for Frivilligsentralene",
        font: FONT_HEADING, size: 44, bold: true, color: COLOR_PRIMARY,
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 600 },
      children: [new TextRun({
        text: "i Bydel Vestre Aker",
        font: FONT_HEADING, size: 44, bold: true, color: COLOR_PRIMARY,
      })],
    }),

    // Skillelinje (en horisontal strek via paragraf border)
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 400 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: COLOR_ACCENT, space: 1 } },
      children: [new TextRun({ text: "" })],
    }),

    // Undertittel
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 240 },
      children: [new TextRun({
        text: "En pilot for lokal beredskap og sivil motstandsdyktighet",
        font: FONT_HEADING, size: 28, italics: true, color: COLOR_TEXT,
      })],
    }),

    new Paragraph({ spacing: { before: 2400, after: 0 } }),

    // Eierskap nederst
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 120 },
      children: [new TextRun({
        text: "Vestre Aker Frivilligsentral",
        font: FONT_HEADING, size: 24, bold: true, color: COLOR_PRIMARY,
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 800 },
      children: [new TextRun({
        text: "Oslo Vest Frivilligsentral",
        font: FONT_HEADING, size: 24, bold: true, color: COLOR_PRIMARY,
      })],
    }),

    // Dato
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 0 },
      children: [new TextRun({
        text: "Mai 2026",
        font: FONT_BODY, size: 22, color: COLOR_MUTED,
      })],
    }),

    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ============================================================
// DOKUMENT-INNHOLD
// ============================================================

function tocSeksjon() {
  return [
    H1("Innholdsfortegnelse", { pageBreakBefore: false }),
    new TableOfContents("", { hyperlink: true, headingStyleRange: "1-3" }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

function sammendrag() {
  return [
    H1("Sammendrag", { pageBreakBefore: false }),
    P("Vestre Aker Frivilligsentral og Oslo Vest Frivilligsentral utvikler en pilot for lokal beredskap og sivil motstandsdyktighet i Bydel Vestre Aker. Piloten har som formål å beskrive konkret hva frivilligsentralene kan bidra med i lokal kriseberedskap, å gjøre dette realistisk å gjennomføre med frivillig kapasitet, og å fungere som en mal som kan kopieres til andre bydeler og kommuner."),
    P("Den umiddelbare bakgrunnen er todelt. På den ene siden står Oslo kommunes Kommunalt risikobilde 2025 (KRB), Totalberedskapsmeldingen og EUs Preparedness Union Strategy — alle peker mot lokal forankring og sterkere sivilsamfunnsrolle i beredskap. På den andre siden er Oslo kommunes offisielle kontaktpunkt for Bydel Vestre Aker ved langvarig krise i dag kun Persbråten videregående skole — et reelt geografisk og funksjonelt gap som frivilligsentralene kan fylle som supplement, ikke erstatning."),
    P("Dokumentet beskriver ti konkrete tiltak. Møteplassen (T1) er grunninfrastrukturen alle andre tiltak hviler på. Beredskapsfrivillige (T2), registeret over sårbare beboere (T3) og koordineringen av spontanfrivillige (T8) er den operasjonelle kjernen. Beredskapskurs (T4), førstehjelpskurs (T5), beredskapsvenner og nabolagsberedskap (T6), og kompetansekartet (T7) bygger ut bredden i lokalsamfunnet. Ungdom og skole (T9) og flerspråklig beredskap (T10) sikrer langsiktig effekt og overførbarhet."),
    P("Tiltakene er vurdert på to akser — viktighet og kompleksitet — og prioritert deretter. Tre tiltak er identifisert som strategisk grunnlag (T1, T2, T3), fire som strategiske raske gevinster (T4, T5, T6, T7), ett som høy prioritet (T8 — koordinering av spontanfrivillige), og to som langsiktige satsinger (T9, T10)."),
    P("Hele piloten bygger på fire ledende prinsipper: frivilligsentralene er supplement og ikke konkurrent til etablerte beredskapsaktører; samarbeid med bydel og kommune er forutsetningen; hvert tiltak designes slik at det kan kopieres av andre bydeler eller kommuner uten å starte fra null; og planleggingen er ambisiøs men gjennomføringen realistisk."),
    P("Veien videre består av to faser. Fase 1 — formell dialog med beredskapsansvarlig i bydelen og Oslo kommune — pågår nå. Fase 2 — utarbeidelse av en konkret handlingsplan med ansvar, ressursbehov og fremdrift — starter når forankringen er på plass."),
  ];
}

function bakgrunn() {
  return [
    H1("Bakgrunn"),

    H2("Hvorfor lokal beredskap nå"),
    P("Norge står overfor et sammensatt risikobilde der digitale, infrastrukturelle og samfunnsmessige sårbarheter forsterker hverandre. Oslo kommunes Kommunalt risikobilde 2025 (KRB) identifiserer fire toppscenarioer av særlig betydning: bortfall av vannforsyning, strømbortfall, cyberangrep mot ekom, og pandemi. Alle fire har det til felles at de gjør digital kommunikasjon sårbar, at de rammer befolkningen direkte, og at de krever rask lokal respons før de offentlige beredskapsstrukturene rekker å mobilisere fullt."),

    faktaboks(
      "KRB Oslo 2025 — topprisikoer",
      "Bortfall av vannforsyning er identifisert som den høyeste risikoen i Oslo. Hovedstaden er avhengig av én enkelt kilde — Maridalsvannet — og reservevannskilden ferdigstilles ikke før 2028. Cyberangrep mot ekom har høyest sannsynlighet av alle scenarioer i KRB-vurderingen, og gjør digital kommunikasjon utilgjengelig i et bredt geografisk område. Pandemi er vurdert med høy sannsynlighet og svært alvorlige konsekvenser.",
      "Oslo kommune — Kommunalt risikobilde 2025"
    ),

    P("Sentrale styringsdokumenter peker i samme retning: Totalberedskapsmeldingen og NOU 2023:17 «Nå er det alvor» understreker begge behovet for styrket sivilsamfunnsrolle og lokal forankring av beredskapsarbeidet. På europeisk nivå går EU Preparedness Union Strategy (mars 2025) lenger, og slår fast at frivillige organisasjoner og sivilsamfunn skal være i kjernen av lokal resiliens. Strategien anbefaler eksplisitt lokale kontaktpunkter bemannet av trente frivillige, integrering av beredskap i skolekurrikulum, og flerspråklig beredskapskommunikasjon."),

    sitatboks(
      "Frivillige organisasjoner og sivilsamfunn skal være i kjernen av lokal resiliens.",
      "EU Preparedness Union Strategy, mars 2025"
    ),

    H2("Persbråten-gapet i Vestre Aker"),
    P("Oslo kommunes offisielle kontaktpunkt for Bydel Vestre Aker ved langvarig strømbortfall og ekom-svikt er i dag kun Persbråten videregående skole. Dette er et godt valg som del av kommunens infrastruktur, men det fungerer ikke alene for et så stort område som Vestre Aker — kupert med åser og friområder som deler bydelen i flere lokalmiljøer."),
    P("For innbyggere med redusert mobilitet — eldre, personer med funksjonsnedsettelse, barnefamilier uten bil — kan Persbråten være utilgjengelig nettopp i de scenarioene der en møteplass trengs mest. Frivilligsentralene har faste lokaler med lokal forankring, kjennskap til nærmiljøet, eksisterende tillitsnettverk og erfaring med å koordinere frivillig innsats. Dette er ressurser som ikke utnyttes systematisk i dag."),

    H2("Frivilligsentralenes rolle"),
    P("Den frivillige innsatsen i Norge bæres i dag av flere parallelle strukturer: nødetatene har sin akutte rolle, Sivilforsvaret sin understøttende rolle, Røde Kors og Norsk Folkehjelp sine fagspesifikke roller. Frivilligsentralene står i en annen posisjon: de er lavterskel møteplasser med direkte kontakt med innbyggerne i en avgrenset geografi, men har i dag ingen formell beredskapsrolle."),
    P("Det er nettopp denne avgrensningen — lokal, lavterskel, generelt borgerrettet — som gjør frivilligsentralene egnet til å fylle gap som de andre aktørene ikke kan fylle. De når mennesker offentlige systemer ikke fanger opp, særlig eldre og marginaliserte. De har kjennskap til lokal kompetanse som verken kommunen eller nødetatene har. Og de har en infrastruktur — lokale, frivillige, rutiner — som med relativt små grep kan skaleres til beredskapsfunksjoner."),

    faktaboks(
      "Forskrift om kommunal beredskapsplikt § 4c",
      "Forskriften krever at kommunen skal ha en oversikt over hvilke ressurser frivillige organisasjoner har som kan være aktuelle ved en uønsket hendelse. Frivilligsentralene ønsker å inngå i denne oversikten formelt.",
      "Forskrift om kommunal beredskapsplikt"
    ),

    H2("Pilotlogikken — fra Vestre Aker til resten av landet"),
    P("Piloten er designet som mer enn et lokalt prosjekt. Hvert tiltak dokumenteres slik at andre bydeler og kommuner kan ta modellen i bruk uten å starte fra null. Dette gir Oslo kommune en overførbar modell for sivilsamfunnets rolle i lokal beredskap, og åpner for at piloten kan bli en nasjonal referansemodell."),
    P("Pilotambisjonen pålegger oss en disiplin: alle tiltak må fungere ikke bare i Vestre Aker, men også i bydeler med andre forutsetninger — flere språkgrupper, andre demografiske mønstre, andre geografiske forutsetninger. Det er en av grunnene til at flerspråklig beredskap (T10) er med selv om andelen flerspråklige innbyggere i Vestre Aker er moderat: modellen må fungere skalerbart."),
  ];
}

function konseptOgPrinsipper() {
  return [
    H1("Konsept og prinsipper"),

    H2("De fire ledende prinsippene"),

    new Paragraph({
      spacing: { before: 200, after: 80 },
      children: [new TextRun({
        text: "1.  Supplement, ikke konkurrent",
        font: FONT_HEADING, size: 24, bold: true, color: COLOR_PRIMARY,
      })],
    }),
    P("Frivilligsentralene fyller gap i eksisterende beredskapsstruktur. Vi overtar aldri roller som tilhører blålysetater, Sivilforsvaret, bydelen eller andre etablerte aktører. Vi avlaster og utvider — vi erstatter ikke."),

    new Paragraph({
      spacing: { before: 200, after: 80 },
      children: [new TextRun({
        text: "2.  Samarbeid er forutsetningen",
        font: FONT_HEADING, size: 24, bold: true, color: COLOR_PRIMARY,
      })],
    }),
    P("Alt skjer i tett dialog med bydel, kommunen og relevante organisasjoner. Ingenting gjøres unilateralt. Beredskapsfrivillige (T2) kan ikke lære kommunens beredskapsplaner uten at kommunen deltar aktivt i opplæringen. Kontakt med DSB, Sivilforsvaret, Norsk Folkehjelp og Røde Kors etableres via Oslo kommunes beredskapsavdeling."),

    new Paragraph({
      spacing: { before: 200, after: 80 },
      children: [new TextRun({
        text: "3.  Pilotlogikk",
        font: FONT_HEADING, size: 24, bold: true, color: COLOR_PRIMARY,
      })],
    }),
    P("Hvert tiltak designes slik at en annen bydel eller kommune kan kopiere det uten å starte fra null. Dokumentasjon, maler og prosesser deles åpent. Suksess måles ikke bare i Vestre Aker, men i hvor mange andre lokalsamfunn modellen tas i bruk."),

    new Paragraph({
      spacing: { before: 200, after: 80 },
      children: [new TextRun({
        text: "4.  Ambisiøs planlegging, realistisk gjennomføring",
        font: FONT_HEADING, size: 24, bold: true, color: COLOR_PRIMARY,
      })],
    }),
    P("Vi planlegger uten ressursbegrensninger først, deretter tilpasser vi til faktisk kapasitet. Tiltakene er designet med ambisjonsnivåer (særlig T1 — møteplassen) slik at vi kan starte enkelt og bygge ut etter hvert som ressurser og kompetanse vokser."),

    H2("Juridisk og politisk forankring"),
    P("Beredskapsarbeidet i frivilligsentralene må forankres i det norske juridiske og politiske rammeverket. De viktigste hjemlene er:"),

    bulletR([
      { text: "Sivilbeskyttelsesloven §§ 14–15", bold: true },
      " — kommunal beredskapsplikt og samarbeid med frivillige aktører",
    ]),
    bulletR([
      { text: "Forskrift om kommunal beredskapsplikt § 4c", bold: true },
      " — kommunen skal ha ressursoversikt over frivillige aktørers kapasitet",
    ]),
    bulletR([
      { text: "Oslo kommunes overordnede beredskapsplan", bold: true },
      " — fire nivåer (0–3) for beredskapsaktivitet, der frivilligsentralene primært opererer på Nivå 0–1",
    ]),
    bulletR([
      { text: "Oslo kommunes beredskapsplan § 6.1", bold: true },
      " — Beredskapsetaten kan aktivere samarbeidsavtaler med frivillige organisasjoner",
    ]),

    P("Politisk forankring finnes på tre nivåer. Nasjonalt slår Totalberedskapsmeldingen og NOU 2023:17 «Nå er det alvor» fast at sivilsamfunnsrollen i beredskap må styrkes. På bydels- og kommunenivå er beredskapsvenn-konseptet (som inngår som T6 Nivå 1) allerede etablert nasjonal politikk anbefalt av både DSB og Oslo kommune. Internasjonalt peker EU Preparedness Union Strategy (mars 2025) i samme retning, og REAL-prosjektet («Resilience through Engagement and Action by Local volunteers») dokumenterer at frivillige organisasjoner med lokalkunnskap er best plassert til å nå sårbare grupper i krise."),

    H2("Aktører og samarbeidspartnere"),
    P("Piloten er bygget rundt et økosystem av aktører på fire nivåer:"),

    // Aktørtabell
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [3000, 6638],
      borders: tableBorders,
      rows: [
        // Header
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            new TableCell({
              width: { size: 3000, type: WidthType.DXA },
              shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR },
              borders: tableBorders,
              margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [new Paragraph({ children: [new TextRun({ text: "Aktør", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })],
            }),
            new TableCell({
              width: { size: 6638, type: WidthType.DXA },
              shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR },
              borders: tableBorders,
              margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [new Paragraph({ children: [new TextRun({ text: "Rolle og status", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })],
            }),
          ],
        }),
        ...[
          ["Bydel Vestre Aker", "Primær samarbeidspart, beredskapsansvarlig. Eksisterende generell samarbeidsrelasjon — formell beredskapsdialog ikke startet."],
          ["Oslo kommune / Beredskapsetaten", "Overordnet myndighet. Formell dialog ikke startet. Kanal for kontakt med DSB, Sivilforsvaret, Norsk Folkehjelp og Røde Kors."],
          ["Oslo Vest Frivilligsentral", "Medeier av piloten. Samarbeid etablert."],
          ["Bydelsmødre Vestre Aker", "Strategisk bro til minoritetsmiljøer (T10). Samarbeid etablert."],
          ["To ungdomsskoler i bydelen", "Inngang via valgfaget «Innsats for andre» (T9). Relasjoner etablert."],
          ["DSB, Sivilforsvaret, Norsk Folkehjelp, Røde Kors", "Opplæringspartnere og materialeleverandører for T2, T4, T5, T6 og T10. Kontakt etableres via Oslo kommunes beredskapsavdeling."],
        ].map(([aktor, rolle]) => new TableRow({
          cantSplit: true,
          children: [
            new TableCell({
              width: { size: 3000, type: WidthType.DXA },
              borders: tableBorders,
              margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [new Paragraph({ children: [new TextRun({ text: aktor, font: FONT_BODY, size: 21, bold: true, color: COLOR_PRIMARY })] })],
            }),
            new TableCell({
              width: { size: 6638, type: WidthType.DXA },
              borders: tableBorders,
              margins: { top: 120, bottom: 120, left: 200, right: 200 },
              children: [new Paragraph({ children: [new TextRun({ text: rolle, font: FONT_BODY, size: 21, color: COLOR_TEXT })] })],
            }),
          ],
        })),
      ],
    }),
    new Paragraph({ spacing: { after: 240 } }),
  ];
}

function visuellOversikt() {
  return [
    H1("Visuell oversikt"),
    P("De fire figurene i dette kapittelet gir et samlet overblikk over de ti tiltakene: hvordan de er prioritert, hvordan de henger sammen, når de iverksettes, og hvilke krisescenarioer de svarer på. De fungerer som inngang til de detaljerte tiltaksbeskrivelsene som følger."),

    H2("Prioriteringsmatrise — viktighet og kompleksitet"),
    P("Alle ti tiltak er vurdert på to akser: viktighet (beredskapsmessig effekt) og kompleksitet (etablerings- og driftskostnad). Resultatet er en visuell fordeling som gir grunnlag for prioritering: strategisk grunnlag (høy viktighet, høyere kompleksitet) etableres først; strategiske raske gevinster (høy viktighet, lav kompleksitet) gir tidlig effekt med begrenset innsats; T8 framheves som høy prioritet på tross av lav kompleksitet — fordi det er et tiltak som ofte glemmes."),
    ...bilde("C:/codex/beredskap/presentasjon/figurer/01_prioriteringsmatrise.png", 580, 425,
      null,
      "Figur 1: Plassering av alle ti tiltak etter viktighet og kompleksitet."),

    H2("Avhengighetsgraf — hvordan tiltakene henger sammen"),
    P("Tiltakene er ikke uavhengige. Møteplassen (T1) er grunninfrastrukturen alle andre tiltak hviler på. Beredskapsfrivillige (T2) bemanner T1 og driver T3 og T8. Kompetansekartet (T7) forsyner T2, T5 og T6 med ressurser. Diagrammet under viser hovedkoblingene — piler peker fra tiltaket mot det det er avhengig av."),
    ...bilde("C:/codex/beredskap/presentasjon/figurer/02_avhengighetsgraf.png", 560, 410,
      null,
      "Figur 2: Avhengighetsgraf — T1 er sentrum, T8 (burgunder) har høy prioritet."),

    H2("Tidslinje for første pilotår"),
    P("Første pilotår deles i tre faser. Fase A (måned 1–3) er forankring og avtaler — dialog med bydel og kommune, etablering av T1 Nivå 1, og oppstart av T3 og T7. Fase B (måned 4–8) er oppbygging — rekruttering og kursing av T2, T6 Nivå 1-kvelder, og oppgradering av T1 til Nivå 2. Fase C (måned 9–12) er aktivering og evaluering — kurs i befolkningen (T4, T5), pilot i ungdomsskole (T9), språkpakke (T10) og felles øvelse."),
    ...bilde("C:/codex/beredskap/presentasjon/figurer/03_tidslinje.png", 600, 343,
      null,
      "Figur 3: Tidslinje med tre faser — forankring, oppbygging, aktivering."),

    H2("Risiko-tiltak-matrise — KRB 2025"),
    P("Hvert av de fire toppscenarioene i KRB Oslo 2025 har tiltak som svarer på det. Møteplassen (T1), beredskapsfrivillige (T2), registeret (T3), koordineringen av spontanfrivillige (T8) og beredskapsvennene (T6) gir sterk effekt på alle fire scenarioer. Andre tiltak gir mer scenariospesifikk effekt. Matrisen viser dette samlet."),
    ...bilde("C:/codex/beredskap/presentasjon/figurer/04_risiko_matrise.png", 600, 277,
      null,
      "Figur 4: Hvilke tiltak svarer på hvilke KRB-toppscenarioer."),
  ];
}

// ============================================================
// DE TI TILTAKENE
// ============================================================

function t1() {
  return [
    ...tiltakHeader("T1", "Lokal møteplass ved krise", 5, 3.5, false),
    H2("Formål"),
    P("Etablere frivilligsentralene som kjente, desentraliserte kontaktpunkter for befolkningen i en krisesituasjon — i tillegg til det offisielle kontaktpunktet ved Persbråten VGS. Møteplassen er grunninfrastrukturen som alle andre tiltak hviler på: uten et fysisk sted med kjent adresse, tilgjengelig utstyr og trente frivillige, er beredskapskorpset (T2), registeret (T3) og koordineringen (T8) uten operasjonelt hjem."),

    H2("Problemet møteplassen løser"),
    P("Bydelen er geografisk stor og kupert, med åser og friområder som deler bydelen i flere lokalmiljøer. For innbyggere med redusert mobilitet — eldre, personer med funksjonsnedsettelse, barnefamilier uten bil — kan Persbråten være utilgjengelig nettopp i de scenarioene der en møteplass trengs mest. Frivilligsentralene har faste lokaler med lokal forankring, kjennskap til nærmiljøet, eksisterende tillitsnettverk og erfaring med å koordinere frivillig innsats."),

    H2("Ambisjonsnivåer"),
    PR([
      { text: "Nivå 1 — Minimumsmøteplass (kompleksitet 2): ", bold: true, color: COLOR_PRIMARY },
      "Fysisk adresse kommunisert i forkant. Tilgang til lokalet uavhengig av elektroniske låssystemer. Frivillige som vet de skal møte opp og hva de skal gjøre. Grunnleggende informasjonsformidling. Koordinert med og godkjent av bydelen.",
    ]),
    PR([
      { text: "Nivå 2 — Utstyrt møteplass (kompleksitet 3): ", bold: true, color: COLOR_PRIMARY },
      "Alt i Nivå 1, pluss beredskapslager, sikringsradio for kontakt med nødetater, kapasitet til enkel førstehjelp, bærbar kortbølgeradio og papirsikkerhetskopi av alle nøkkellister.",
    ]),
    PR([
      { text: "Nivå 3 — Fullt utbygd beredskapspunkt (kompleksitet 4–5): ", bold: true, color: COLOR_PRIMARY },
      "Alt i Nivå 1 og 2, pluss fast korps av beredskapsfrivillige, hjertestarter (AED) med opplært personell, nødstrøm, del av kommunens formelle ressursoversikt, regelmessige øvelser og NRRL-samarbeid.",
    ]),

    H2("Utstyrsliste på Nivå 2–3"),
    bullet("Sikringsradio (batteri/håndsvingdynamo) for kontakt med nødetater"),
    bullet("Bærbar kortbølgeradio for mottak av NRK P1 og nødsendinger"),
    bullet("Hjertestarter (AED) — krever opplæring (jf. T5)"),
    bullet("Vannreserver, beholdere og vannfilter"),
    bullet("Nødlys, hodelykt, kraftbank og megafon (alle batteridrevet)"),
    bullet("Førstehjelpsskrin, varmeposer og tepper"),
    bullet("Papirsikkerhetskopi av T3-register, T7-kompetansekart og varslingsliste"),
    bullet("Mottakssett for spontanfrivillige (jf. T8): digital registrering primær (nettbrett/laptop + skriver), forhåndstrykte skjemaer som sikkerhetskopi, oppgaveoversikt, skiftlogg, navnelapper/vester"),

    faktaboks(
      "UK Resilience Hubs — modell for forhåndslagret utstyr",
      "I Storbritannia er konseptet «pre-positioned equipment» en kjernekomponent av Resilience Hubs — lokale beredskapspunkter med utstyr og dokumentasjon klar for 24/7-tilgang. Sentralt er at lokalene er kjent for befolkningen i forkant, ikke først ved krise. Denne modellen er direkte forbilde for frivilligsentralenes møteplasser.",
      "UK Cabinet Office — Local Resilience Forum guidance"
    ),

    H2("Nøkkeltilgang"),
    P("Begge frivilligsentralene har samme system: elektronisk lås til daglig med manuell overstyringsnøkkel. Nøkkelen kopieres opp og distribueres til et knippe nøkkelpersoner som bor i nær geografisk avstand til lokalet. Nøkkelboks vurderes som for usikkert (lett å bryte opp) og benyttes ikke."),

    H2("Nødkommunikasjon når digitalt nett er nede"),
    P("Digital kommunikasjon er det mest sårbare punktet i alle tre toppscenarioer fra KRB 2025. Møteplassen må fungere uten mobilnett og internett:"),
    bullet("Sikringsradio for kontakt med nødetater"),
    bullet("Fysisk varslingstre for nøkkelfrivillige (ringer videre nedover telefonkjeden)"),
    bullet("Forhåndsavtalt møtetidspunkt ved krise («alle nøkkelfrivillige møtes kl. 08:00 dersom varsling ikke er mulig»)"),
    bullet("Oppslagstavle som primær informasjonskanal til befolkningen ved digital svikt"),
    bullet("Samarbeid med NRRL (Norsk Radio Relæ Liga) som har nødkommunikasjonskapasitet uavhengig av kommersiell infrastruktur"),

    maltallsboks(
      [
        "Begge frivilligsentraler etablert som møteplasser på minimum Nivå 1",
        "Minst én av frivilligsentralene løftet til Nivå 2 — beredskapslager med kjernekomponenter etablert",
        "Formell anerkjennelse fra bydelen om at møteplassene inngår i kommunens ressursoversikt (jf. Forskrift § 4c)",
        "Adressene og funksjonen kommunisert til befolkningen i bydelen",
      ],
      "Nøkkeltilgang fungerer uten elektronikk, og minst én øvelse av aktivering er gjennomført."
    ),
  ];
}

function t2() {
  return [
    ...tiltakHeader("T2", "Beredskapsfrivillige", 5, 3, false),
    H2("Formål"),
    P("Rekruttere og opprettholde et fast korps av frivillige som er spesifikt kurset for krisehåndtering — og som er klare til å møte opp, bemanne møteplassen (T1) og gjennomføre konkrete oppgaver når en krise inntreffer."),

    H2("Hva som skiller beredskapsfrivillige fra ordinære frivillige"),
    bullet("Forhåndsrekruttert og -kurset — ikke spontant oppmøtt (jf. T8)"),
    bullet("Definert rolle de kjenner og har øvd på"),
    bullet("Kjenner kommunens beredskapsplaner og praktiske løsninger"),
    bullet("Forpliktet seg til å møte opp ved aktivering"),
    bullet("Del av kommunens offisielle ressursoversikt (jf. Forskrift § 4c)"),

    H2("Hvem er uaktuelle, og hvem er ideelle"),
    P("Folk ansatt i nødetater, forsvar, helsevesen, Sivilforsvaret, Røde Kors og andre samfunnskritiske funksjoner er uaktuelle — de vil selv bli kalt ut ved en krise. Beredskapsfrivillige må rekrutteres utenfor denne gruppen. Samme avgrensning ligger til grunn for T7 (kompetansekartet)."),
    P("Kjernemålgruppen er pensjonister i 60-årene med tidligere yrkesbakgrunn fra forsvar, nødetater eller helsevesen. God helse er et absolutt krav. Skillet mellom uaktuell og kjernemålgruppe går på status nå, ikke yrkesbakgrunn: en aktiv ansatt sykepleier vil bli kalt til OUS, mens en pensjonert sykepleier med samme bakgrunn ikke er på noen utrykningsliste."),
    P("Vi rekrutterer ut fra dette idealet og utvider gradvis til hjemmeværende med fleksibel hverdag, selvstendig næringsdrivende med lokal tilknytning, og deltidsarbeidende utenfor samfunnskritiske sektorer."),

    H2("Korps-størrelse"),
    P("Startmålet er 6–8 aktive beredskapsfrivillige per møteplass — vurdert som realistisk og godt utgangspunkt. Langsiktig mål er 12–15 per møteplass for å sikre tilstrekkelig dekning uansett tidspunkt og årstid."),

    H2("Opplæring"),
    P("Opplæringen bygger på en norsk, Oslo-tilpasset versjon av CERT-modellen (Community Emergency Response Team — amerikansk modell etablert siden 1985), dimensjonert mot KRB 2025-scenarioene. Grunnkurset er ca. 12–16 timer fordelt på åtte moduler:"),
    bullet("Modul 1: Rollen som beredskapsfrivillig — kommunens struktur, frivilligsentralenes plass"),
    bullet("Modul 2: De viktigste risikoscenarioene — KRB 2025"),
    bullet("Modul 3: Møteplassen i praksis — utstyr, prosedyrer, nøkkeltilgang"),
    bullet("Modul 4: Enkel førstehjelp — HLR, AED, sårbehandling"),
    bullet("Modul 5: Kommunikasjon under krise — sikringsradio, informasjonsformidling"),
    bullet("Modul 6: Sårbare grupper — T3-registeret, oppsøkende arbeid"),
    bullet("Modul 7: Koordinering og spontanfrivillige — T8-prosedyren"),
    bullet("Modul 8: Simuleringsøvelse — fullskala-scenario fra start til avslutning"),

    H2("Opplæringspartnere"),
    P("Det eksisterer i dag ingen formell kontakt med DSB, Sivilforsvaret, Norsk Folkehjelp eller Røde Kors lokalt. Etablering av disse relasjonene er en forutsetning for T2. Kontakten etableres via Oslo kommunes beredskapsavdeling — ikke direkte. Dette er et konkret punkt å ta opp i fase 1-dialogen med beredskapsansvarlig i bydelen, og samme kanal brukes for T4 og T5."),

    faktaboks(
      "Svenske Frivilliga Resursgruppen (FRG)",
      "Sveriges Frivilliga Resursgrupper er den nærmeste internasjonale parallellen til T2. FRG er kommunalt forankret, har fast opplæring, aktiveres av kommunen og har vært lokalt forankret siden 2004. FRG ble etablert nettopp for å løse problemet med å mobilisere lokal kapasitet utenfor de etablerte beredskapsstrukturene. Modellen er dokumentert effektiv på små og mellomstore hendelser, særlig der nødetatene er overbelastet eller geografisk fjerne.",
      "Civilförsvarsförbundet — Frivilliga Resursgruppen",
      "https://civil.se/om-oss/frivilliga-resursgruppen/"
    ),

    H2("Aktivering"),
    bullet("Normal kommunikasjon: SMS/telefonkjede (varslingstre)"),
    bullet("Ved kommunikasjonsvikt: forhåndsavtalt fysisk oppmøte («alle møter kl. 08:00 ved lokalet»)"),
    bullet("Responstid: minimum 2–3 beredskapsfrivillige på plass innen 60 minutter, uansett tidspunkt"),

    maltallsboks(
      [
        "6–8 aktive beredskapsfrivillige per møteplass rekruttert og gjennomført grunnkurs (12–16 timer)",
        "Formell samarbeidsavtale mellom frivilligsentralen og bydelen inngått",
        "Opplæringspartnere etablert via Oslo kommunes beredskapsavdeling (DSB, Sivilforsvaret, Norsk Folkehjelp, Røde Kors)",
        "Varslingstre operasjonelt og testet",
        "Minimum én øvelse gjennomført i samarbeid med bydelen",
      ],
      "Minst 2–3 beredskapsfrivillige kan stille på møteplassen innen 60 minutter, uansett tidspunkt."
    ),
  ];
}

function t3() {
  return [
    ...tiltakHeader("T3", "Register over sårbare beboere", 5, 4, false),
    H2("Formål"),
    P("Etablere og vedlikeholde et oppdatert register over eldre og andre med hjelpebehov som ved en krisesituasjon vil trenge ikke-medisinsk bistand og ikke nødvendigvis kan nå frem til møteplassen selv. Registeret gjør innsatsen proaktiv fremfor reaktiv: frivillige vet hvem de skal oppsøke, hvor de bor, og hva de trenger — de trenger ikke vente på at folk varsler."),

    faktaboks(
      "1000+ innbyggere i målgruppen",
      "Over 1000 innbyggere i Bydel Vestre Aker mottar hjemmetjenester. Alle er potensielt i målgruppen for T3-registeret. Det forutsetter en strukturert prioriteringsmodell i tre nivåer — fra dem som må oppsøkes innen 1–2 timer, til dem som kan sjekkes per telefon.",
      "Bydel Vestre Aker (intern statistikk)"
    ),

    H2("Hva registeret inneholder"),
    P("For hver registrert person: navn, adresse (inkl. etasje, portkode, tilgangsinfo), telefonnummer (primær og pårørende), type hjelp som kan bli nødvendig, spesielle hensyn (bevegelseshemming, hørselstap, demens, språk), dato for siste oppdatering, og dokumentert samtykke."),

    H2("Format — hybrid digital primær og papirsikkerhetskopi"),
    P("Daglig drift krever digital registrering for at registeret skal være vedlikeholdbart over tid. Samtidig må informasjonen være tilgjengelig ved strøm- og ekom-svikt — som er nettopp scenarioene registeret skal brukes i. Løsningen er hybrid:"),
    PR([
      { text: "Digital arbeidskopi (primær): ", bold: true, color: COLOR_PRIMARY },
      "Lagres på ekstern, kryptert disk med backupdisk. Ikke skylagring, ikke nettverkstilkoblet. Strenge tilgangskontroller, passord, kort og dokumentert tilgangsliste. Disk og backupdisk oppbevares fysisk adskilt — én på møteplassen, én på sekundær lokasjon.",
    ]),
    PR([
      { text: "Papirsikkerhetskopi (ved krise): ", bold: true, color: COLOR_PRIMARY },
      "Skrives ut minimum hver tredje måned og etter større oppdateringer. Oppbevares i låst arkiv på møteplassen (T1), kun tilgjengelig for autoriserte beredskapsfrivillige. Utdaterte papirkopier makuleres ved hver utskrift.",
    ]),

    H2("Prioritering i tre nivåer"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [1800, 5400, 2438],
      borders: tableBorders,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            new TableCell({ width: { size: 1800, type: WidthType.DXA }, shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: "Prioritet", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })] }),
            new TableCell({ width: { size: 5400, type: WidthType.DXA }, shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: "Profil", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })] }),
            new TableCell({ width: { size: 2438, type: WidthType.DXA }, shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: "Responstid", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })] }),
          ],
        }),
        ...[
          ["Høy", "Bor alene, ingen familie i nærheten, sterkt begrenset mobilitet", "Oppsøkes innen 1–2 timer"],
          ["Middels", "Noe støttenettverk, men sårbar", "Oppsøkes innen 4–6 timer"],
          ["Lav", "Familie/naboer kan hjelpe, men bør sjekkes", "Telefon eller nabolagssjekk"],
        ].map(row => new TableRow({
          cantSplit: true,
          children: row.map((c, i) => new TableCell({
            width: { size: [1800, 5400, 2438][i], type: WidthType.DXA },
            borders: tableBorders,
            margins: { top: 120, bottom: 120, left: 200, right: 200 },
            children: [new Paragraph({ children: [new TextRun({ text: c, font: FONT_BODY, size: 21, color: COLOR_TEXT, bold: i === 0 })] })],
          })),
        })),
      ],
    }),
    new Paragraph({ spacing: { after: 240 } }),

    H2("To spor for registeret"),
    PR([
      { text: "Spor 1 — Frivilligsentralens eget register (frivillig opt-in): ", bold: true, color: COLOR_PRIMARY },
      "Frivilligsentralen har allerede en oversikt over personer som mottar tjenester i dag — dette er det naturlige utgangspunktet. Registeret formaliseres og utvides via T6-arrangementer i borettslag, T4-beredskapskurs og besøkstjeneste-prosjektet.",
    ]),
    PR([
      { text: "Spor 2 — Kommunalt datagrunnlag (krever formell avtale): ", bold: true, color: COLOR_PRIMARY },
      "Bydelen/kommunen deler kontaktinfo om eksisterende tjenestemottakere med frivilligsentralen. Krever formell samarbeidsavtale, juridisk avklaring via kommuneadvokaten, aktivt samtykke fra tjenestemottakerne og presedens fra andre kommuner.",
    ]),

    H2("GDPR — reell og opplevd barriere"),
    P("Det eksisterer godt samarbeid med bydelen og kommunale tjenester, men bydelens opplevelse av GDPR er et reelt hinder. Hinderet er like mye kulturelt og holdningsmessig som juridisk. Mange kommuner er mer restriktive enn loven krever. Kommuneadvokat-prosessen handler ikke bare om å finne en juridisk løsning — den handler om å gi bydelen trygghet og juridisk ryggrad til å handle."),

    sitatboks(
      "Kommunen bør snarest innføre en rutine der alle som søker om kommunale tjenester, bes om aktivt samtykke til at kontaktinformasjon kan deles med samarbeidende frivillige organisasjoner ved krise. Dette er et systemgrep som bør innføres uavhengig av hvem som er frivillig samarbeidspart.",
      "Politisk anbefaling fra denne piloten"
    ),

    faktaboks(
      "COVID-19 (2020) — dokumentasjon på behovet",
      "Under pandemien i 2020 oppstod et massivt tilbud av spontanfrivillige som ville hjelpe hjemmeboende eldre og sårbare. Frivilligsentralene hadde ingen plan for situasjonen. Det avgjørende hinderet var at GDPR forhindret tilgang til kontaktinformasjon — frivillige hadde ingen å kontakte. Ressursen gikk tapt. Dette er et direkte argument i dialogen med bydelen for å løse GDPR-utfordringen nå.",
      "Erfaringsgrunnlag fra Vestre Aker Frivilligsentral, 2020"
    ),

    maltallsboks(
      [
        "Spor 1 etablert — registeransvarlig utpekt, samtykkemal godkjent, hybrid lagringsmodell operasjonell",
        "50–100 personer registrert i Spor 1 via frivilligsentralens egne tjenestemottakere og T6-arrangementer",
        "Kommuneadvokat-prosess startet for å avklare juridisk grunnlag for Spor 2",
        "Kommunal politikk-anbefaling om samtykke ved tjenestesøknad fremmet for bydelen",
        "Prioriteringsmodell (Høy/Middels/Lav) tatt i bruk i registeret",
      ],
      "Aktiveringsrutinen testet i øvelse, og papirkopi på møteplassen er aktuell innenfor 3 måneder."
    ),
  ];
}

function t4() {
  return [
    ...tiltakHeader("T4", "Beredskapskurs for befolkningen", 4, 2, false),
    H2("Formål"),
    P("Øke den generelle befolkningens evne til å klare seg selv i en krisesituasjon — slik at presset på beredskapsfrivillige (T2), møteplassen (T1) og registeret (T3) reduseres."),

    faktaboks(
      "Norsk anbefaling: én uke, ikke 72 timer",
      "Den norske offisielle anbefalingen fra DSB er at en husstand skal kunne klare seg i minimum én uke uten hjelp utenfra. 72 timer er den internasjonale minimumsstandarden — norsk anbefaling er høyere. Dette presenteres trinnvis i kurset: «start med 72 timer, bygg mot en uke.»",
      "DSB — egenberedskap"
    ),

    H2("Kursinnhold — syv moduler"),
    PR([{ text: "Modul 1 — Forstå risikoen lokalt: ", bold: true, color: COLOR_PRIMARY }, "Hvilke kriser er mest sannsynlige i Oslo og Vestre Aker. KRB 2025-scenarioene oversatt til hva de betyr hjemme."]),
    PR([{ text: "Modul 2 — Vann: ", bold: true, color: COLOR_PRIMARY }, "Minimum 3 liter per person per dag × 7 dager = 21 liter. Vannfilter, alternativt råvann, rensemetoder."]),
    PR([{ text: "Modul 3 — Mat: ", bold: true, color: COLOR_PRIMARY }, "Hva bør ligge på lager. Matlaging uten strøm. Kjølemat og fryser. Spesielle behov."]),
    PR([{ text: "Modul 4 — Varme og ly: ", bold: true, color: COLOR_PRIMARY }, "Alternativer ved strømbortfall. Brannfare og karbonmonoksid ved alternativ oppvarming."]),
    PR([{ text: "Modul 5 — Kommunikasjon: ", bold: true, color: COLOR_PRIMARY }, "Bærbar radio, NRK P1, møteplassene som fysisk informasjonskilde. Kontanter."]),
    PR([{ text: "Modul 6 — Førstehjelp: ", bold: true, color: COLOR_PRIMARY }, "Førstehjelpsskrin hjemme. Medisinreserve (minimum 2 ukers reserve)."]),
    PR([{ text: "Modul 7 — Familie og naboer: ", bold: true, color: COLOR_PRIMARY }, "Beredskapsplan for husstanden. Introduksjon til beredskapsvenn (T6)."]),

    H2("Format"),
    P("Kveldsarrangement på 2–2,5 timer: presentasjon (45 min), praktisk gjennomgang med faktisk utstyr (45 min), spørsmål og diskusjon (30 min), rekruttering til T6 og T2 (15 min)."),
    P("Arenaer: frivilligsentralenes lokaler, borettslag (kombinert med T6), Deichmanske Røa, menigheter. Frekvens: minimum 2–4 kurs per år per frivilligsentral."),
    P("Kursholdere er beredskapsfrivillige (T2) med grunnkurs, eller i samarbeid med Sivilforsvaret/DSB. DSB og Sivilforsvaret har ferdig kursmateriale — ikke behov for å utvikle fra bunnen. Frivilligsentralen har erfaring med å arrangere kvelder og foredrag, men ikke med beredskapstematikk. Formatkompetansen er der — innholdet er nytt."),

    H2("Etablering av samarbeid med DSB og Sivilforsvaret"),
    P("Frivilligsentralen har i dag ingen formell kontakt med DSB eller Sivilforsvaret lokalt. Kontakten etableres via Oslo kommunes beredskapsavdeling — ikke direkte. Dette er et konkret punkt å ta opp i fase 1-dialogen med beredskapsansvarlig i bydelen."),

    H2("Rekruttering til kursene"),
    P("Frivilligsentralen har formatkompetansen — å fylle rommet er den reelle utfordringen. Kommunikasjon og markedsføring må treffe det folk faktisk er bekymret for, ikke beredskapsfaglige formuleringer. En egen rekrutteringsplan utarbeides som del av T4-arbeidet."),

    maltallsboks(
      [
        "4–6 kurs gjennomført totalt på tvers av begge frivilligsentraler",
        "Rekrutteringsplan utarbeidet og testet — kommunikasjon som treffer det folk er bekymret for",
        "Formelt samarbeid med DSB/Sivilforsvaret om kursmateriale etablert via Oslo kommunes beredskapsavdeling",
        "Kursholderkorps etablert (beredskapsfrivillige fra T2 og lokale fagpersoner fra T7)",
      ],
      "Minimum gjennomsnitt 15 deltakere per kurs, og minst halvparten av deltakerne rekrutteres videre til T6 (beredskapsvenn)."
    ),
  ];
}

function t5() {
  return [
    ...tiltakHeader("T5", "Førstehjelp for befolkningen", 4, 2, false),
    H2("Formål"),
    P("Øke andelen innbyggere i bydelen som kan yte grunnleggende livreddende førstehjelp. De første minuttene etter en ulykke eller akutt sykdom er kritiske — og ved en krise kan nødetater ha lang responstid eller være overbelastet. Jo flere i nærmiljøet som kan handle, desto bedre."),

    H2("Frivilligsentralens rolle — tilrettelegger, ikke kursholder"),
    P("Frivilligsentralen arrangerer og markedsfører kursene, stiller lokaler og rekrutterer deltakere. Selve kursinnholdet leveres av etablerte aktører med godkjent opplæring: Røde Kors, Norsk Folkehjelp eller tilsvarende. Frivilligsentralen trenger ikke bygge kurskompetansen selv."),

    H2("Lokale instruktører — en undervurdert ressurs"),
    P("Flere typer personer i nærmiljøet kan holde kurs gratis eller mot lav kompensasjon: offentlig ansatte med medisinsk bakgrunn, pensjonerte fra forsvaret eller nødetatene, og sertifiserte instruktører fra Røde Kors eller Norsk Folkehjelp i nettverket."),
    P("Aktive ansatte i nødetater og helsevesen kan godt holde kurs i fredstid på fritiden — det er en annen funksjon enn å være beredskapsressurs ved krise. De skal ikke registreres i kompetansekartet (T7) eller rekrutteres som beredskapsfrivillige (T2), fordi de blir kalt inn til sin egen tjeneste når krisen rammer. Men som kursholdere når alt er normalt, er de en verdifull ressurs."),
    P("Frivilligsentralen har allerede én slik instruktør i prosjektgruppen som holder kurs frivillig. Dette er et bevis på at modellen fungerer, og det er ressurser å mobilisere — kompetansekartet (T7) skal systematisk identifisere flere lokale, pensjonerte fagpersoner."),

    H2("Kurstyper"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [2400, 1800, 3000, 2438],
      borders: tableBorders,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            new TableCell({ width: { size: 2400, type: WidthType.DXA }, shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: "Kurstype", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })] }),
            new TableCell({ width: { size: 1800, type: WidthType.DXA }, shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: "Varighet", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })] }),
            new TableCell({ width: { size: 3000, type: WidthType.DXA }, shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: "Innhold", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })] }),
            new TableCell({ width: { size: 2438, type: WidthType.DXA }, shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: "Målgruppe", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })] }),
          ],
        }),
        ...[
          ["HLR-kurs", "2–3 timer", "Hjerte-lungeredning og AED-bruk", "Alle"],
          ["Førstehjelp grunnkurs", "8–16 timer", "Sårbehandling, sjokk, ulykker, HLR", "De som vil mer"],
          ["Hjertestarter-kurs", "1–2 timer", "AED-bruk alene", "Borettslag, arbeidsplasser"],
        ].map(row => new TableRow({
          cantSplit: true,
          children: row.map((c, i) => new TableCell({
            width: { size: [2400, 1800, 3000, 2438][i], type: WidthType.DXA },
            borders: tableBorders,
            margins: { top: 120, bottom: 120, left: 200, right: 200 },
            children: [new Paragraph({ children: [new TextRun({ text: c, font: FONT_BODY, size: 21, color: COLOR_TEXT, bold: i === 0 })] })],
          })),
        })),
      ],
    }),
    new Paragraph({ spacing: { after: 240 } }),

    P("Prioritet er HLR + AED. Livreddende effekt er størst der, og kurslengden er lav nok til at folk faktisk deltar."),

    H2("Finansiering — deltakerne skal ikke betale"),
    P("Prinsippet er at kursene skal være gratis for deltakerne. Lav terskel er avgjørende for rekkevidde. Finansieringskilder er tilskudd fra bydel, kommune, stiftelser og sentrale beredskapsmidler; inntektsgenererende aktiviteter ved frivilligsentralen (loddsalg og lignende); og frivillig instruktørtid fra lokale fagpersoner. Dette er en grunnleggende del av T5-modellen og må kommuniseres tydelig overfor bydelen som del av samarbeidsdialogen."),

    H2("Sertifisering og oppfriskning"),
    P("Deltakere får dokumentasjon på fullført kurs etter den aktuelle kursaktørens standard (Røde Kors-bevis, Norsk Folkehjelp-bevis). HLR-kompetanse forfaller raskt — både ferdigheter og selvtillit. Oppfriskningskurs hvert 2. år legges inn som standard. Det reduserer terskelen for å komme tilbake og sikrer at det faktisk finnes kompetanse å mobilisere når noe skjer."),

    H2("Målgrupper med særlig høy effekt"),
    bullet("Foreldre med småbarn — høy frekvens av akutte hendelser, høy motivasjon"),
    bullet("Eldre som bor alene eller med eldre partner — egen sårbarhet, men også evne til å handle for partner/nabo"),
    bullet("Nabolag med beredskapsgrupper eller beredskapsvenner (T6) — strukturen er allerede der, kurset bygger ut kapasitet i et nettverk som vil aktiveres ved krise"),

    maltallsboks(
      [
        "Minimum 4 kurs gjennomført totalt — HLR + AED prioritert",
        "Finansieringsmodell på plass — minimum én tilskuddssøknad sendt, og minst én inntektsgenererende aktivitet gjennomført",
        "Minimum 2 lokale instruktører rekruttert via T7-kompetansekartet (i tillegg til den ene som allerede er i prosjektgruppen)",
        "Samarbeidsavtale med Røde Kors og/eller Norsk Folkehjelp om sertifisering",
        "Oppfriskningsrutine etablert (hvert 2. år)",
      ],
      "Minst 60 personer i bydelen kurset i HLR + AED gjennom året, og minst 20 % rekrutteres videre til beredskapsgruppe (T6) eller kompetansekart (T7)."
    ),
  ];
}

function t6() {
  return [
    ...tiltakHeader("T6", "Beredskapsvenner og nabolagsberedskap", 4, 2, false),
    H2("Formål"),
    P("Bygge nabolagsbasert beredskap der innbyggere tar ansvar for seg selv og sine nærmeste først — slik at presset på møteplassen (T1), beredskapsfrivillige (T2) og nødetatene reduseres mest mulig i tidlige timer av en krise. Tiltaket har to nivåer som bygger på hverandre."),

    H2("Hvorfor to nivåer"),
    P("Beredskapsvenn-konseptet — én-til-én-avtaler mellom naboer — er allerede offisiell nasjonal politikk fra DSB og Oslo kommune. Det er et lett første skritt og en åpningsdør inn i borettslagene. Beredskapsgrupper i borettslag er et mer organisert nivå — og det er der den reelle beredskapsmessige effekten er størst."),
    P("Nivå 1 alene er ikke nok ved en større krise. Nivå 2 alene har for høy etableringsterskel for de fleste borettslag uten oppvarming. Sekvensen Nivå 1 → Nivå 2 fungerer."),

    H3("Nivå 1 — Beredskapsvenn (parvis nabolagsavtale)"),
    P("To naboer (eller en liten gruppe) avtaler å ta vare på hverandre ved en krise. De vet hvor hverandre bor, har nøkkelinformasjon, og sjekker inn først ved strømbortfall, stormvarsel eller annen hendelse."),
    P("Frivilligsentralens rolle er å åpne døren inn til borettslagene via kveldsarrangementer. Selve avtalen inngås mellom naboene — frivilligsentralen formidler konseptet og letter prosessen. DSBs eksisterende beredskapsvenn-materiale brukes som det er. Ingen utvikling fra bunnen."),
    P("Format er et kveldsarrangement på 1,5–2 timer: velkomst og lokal kontekst (15 min), beredskapsvenn-konseptet i praksis (30 min), praktisk øvelse hvor deltakere snakker med sidemann om hvem deres beredskapsvenn kan være (30 min), spørsmål og åpning mot Nivå 2 (15 min), sosial avslutning med kaffe (15 min). Arenaer er borettslag og sameier (primær), frivilligsentralenes lokaler, menigheter, eldreråd og pensjonistforeninger."),

    H3("Nivå 2 — Beredskapsgrupper i borettslag (organisert nabolagsberedskap)"),
    P("For borettslag og sameier som vil gå lengre enn parvise avtaler. Minstekrav er én beredskapskontakt per borettslag/sameie (frivillig rolle, ikke styreleder per automatikk), enkel kartlegging av nabolaget (hvem bor her, hvem trenger hjelp, hvem kan hjelpe), én årlig aktivitet (drill, informasjonskveld eller gjennomgang av egenberedskap), og tilknytning til møteplass (T1) og register (T3)."),

    faktaboks(
      "Jishu-bosai-soshiki — japansk modell",
      "Beredskapsgrupper i borettslag er inspirert av japansk jishu-bosai-soshiki — frivillige nabolagsberedskapsgrupper som finnes i 84 % av japanske nabolag. Etter jordskjelvet på Noto-halvøya i januar 2024 fikk husstander tilknyttet en nabolagsforening hjelp vesentlig raskere enn isolerte husstander. Modellen er dokumentert effektiv på akkurat det den skal: de første timene før hjelp utenfra kommer.",
      "Bajek m.fl. — Natural Hazards (Springer)",
      "https://link.springer.com/article/10.1007/s11069-007-9107-4"
    ),

    H2("Frivilligsentralens rolle på Nivå 2 — logistisk hub"),
    P("Nabolagene driver sine egne beredskapsgrupper, men trenger jevnlig logistisk støtte. Frivilligsentralen leverer denne støtten:"),
    bulletR([
      { text: "Kursmateriell — ", bold: true, color: COLOR_PRIMARY },
      "maler for kartlegging, sjekklister, presentasjonsmateriale til årlige aktiviteter. Borettslagene skal ikke utvikle fra bunnen.",
    ]),
    bulletR([
      { text: "Kursing av beredskapskontakter — ", bold: true, color: COLOR_PRIMARY },
      "strukturert grunnopplæring for nye kontakter (rolle, prosedyrer, kobling til T1/T3, hva som forventes ved aktivering). Oppfriskningskurs ved behov.",
    ]),
    bulletR([
      { text: "Nettverksgruppe for beredskapskontakter — ", bold: true, color: COLOR_PRIMARY },
      "fast møteplass på tvers av borettslag for erfaringsutveksling, problemløsning og oppdateringer. Minimum 2 samlinger per år.",
    ]),
    bulletR([
      { text: "Nøkkelpersoner med kompetanse — ", bold: true, color: COLOR_PRIMARY },
      "beredskapskontaktene har en konkret kontaktliste hos frivilligsentralen: hvem ringer jeg for spørsmål om førstehjelp, hvem ringer jeg om GDPR, hvem ringer jeg om praktiske beredskapsspørsmål? Kompetansekartet (T7) er grunnlaget.",
    ]),
    bulletR([
      { text: "Møtelokaler — ", bold: true, color: COLOR_PRIMARY },
      "frivilligsentralens lokaler er tilgjengelige for beredskapsgruppenes egne aktiviteter når egne fellesarealer ikke fungerer.",
    ]),

    P("Prinsippet er at borettslagene driver lokalt, men aldri står alene. Den jevnlige støtten er det som skiller en levende beredskapsgruppe fra en sovende."),

    H2("Kobling til nasjonal og lokal politikk"),
    bullet("DSB og Oslo kommune anbefaler aktivt beredskapsvenn-konseptet. Nivå 1 er allerede etablert nasjonal politikk."),
    bullet("Forskrift om kommunal beredskapsplikt § 4c krever ressursoversikt over frivillige aktørers kapasitet — organiserte beredskapsgrupper i borettslag styrker denne oversikten."),
    bullet("EU Preparedness Union Strategy (mars 2025) peker spesifikt på lokale strukturer i sivilsamfunnet som kjernen i beredskap."),

    maltallsboks(
      [
        "Nivå 1: 4–6 beredskapsvennkvelder gjennomført i ulike borettslag/arenaer",
        "Nivå 2: 2–3 borettslag etablerer beredskapsgruppe",
        "Kursmateriell og maler ferdigstilt og tilgjengelig for borettslagene",
        "Nettverksgruppe for beredskapskontakter etablert med første samling gjennomført",
      ],
      "Minst halvparten av deltakerne på beredskapsvennkveld identifiserer konkret beredskapsvenn samme kveld; etablerte beredskapsgrupper gjennomfører sin første årlige aktivitet."
    ),
  ];
}

function t7() {
  return [
    ...tiltakHeader("T7", "Kompetansekart over frivillige", 4, 2, false),
    H2("Formål"),
    P("Etablere og vedlikeholde et oppdatert kompetansekart over hvem som kan gi spesialisert hjelp lokalt i en krisesituasjon — som komplement til T3-registeret som kartlegger hvem som trenger hjelp. Frivilligsentralen vet noe om lokalsamfunnet som kommunen og nødetatene ikke vet, og dette er en unik ressurs som ikke utnyttes systematisk i dag."),

    H2("Det bærende prinsippet — lokalt tilgjengelig kompetanse"),
    P("Kompetansekartet skal fange den lokalt tilgjengelige kompetansen — folk som faktisk kan møte opp i nærmiljøet når noe skjer. Aktive ansatte i nødetater, forsvar, helsevesen, Sivilforsvaret, Røde Kors og andre samfunnskritiske funksjoner er ikke målgruppen. De vil bli kalt inn til sin egen tjeneste når krisen rammer — de er allerede koblet til en aktiveringsstruktur og er ikke fritt tilgjengelige lokalt."),
    P("Målgruppen er heller: pensjonister med relevant yrkesbakgrunn; hjemmeværende eller deltidsarbeidende med relevant kompetanse; selvstendig næringsdrivende med lokal tilknytning og fleksibilitet; yrkesaktive utenfor samfunnskritiske sektorer med relevant sidekompetanse; og tidligere frivillige i Røde Kors, Norsk Folkehjelp m.fl. som ikke lenger har aktiv tjenesteplikt der."),

    sitatboks(
      "En aktiv brannmann i Vestre Aker er ikke i kartet — han kjører ut. En pensjonert brannmann som bor i samme oppgang som tre eldre er gull i kartet — han kan organisere evakuering eller førstehjelp lokalt mens nødetatene er andre steder.",
      "Praktisk illustrasjon av prinsippet"
    ),

    H2("Kompetansekategorier"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [2800, 6838],
      borders: tableBorders,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            new TableCell({ width: { size: 2800, type: WidthType.DXA }, shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: "Kategori", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })] }),
            new TableCell({ width: { size: 6838, type: WidthType.DXA }, shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: "Eksempler på relevant bakgrunn", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })] }),
          ],
        }),
        ...[
          ["Helsefaglig", "Pensjonerte sykepleiere, leger, hjelpepleiere, jordmødre, ambulansepersonell"],
          ["Teknisk", "Elektrikere, VVS/rørleggere, mekanikere, IT-teknikere, snekkere"],
          ["Logistikk", "Sjåfører (med eller uten egen bil — frivilligsentralen har to biler tilgjengelig), folk med tilhenger, firehjulstrekk, stor varebil, tidligere lastebilsjåfører"],
          ["Kommunikasjon", "Tolker og oversettere (med språk spesifisert)"],
          ["Sosialt", "Pensjonerte sosionomer, psykologer, lærere med kriseerfaring"],
          ["Praktisk", "Folk med generator, stor fryseboks, vedovn, egnet lokale, NRRL-medlemmer"],
          ["Beredskapserfaring", "Pensjonerte fra forsvar, nødetater, Sivilforsvaret, tidligere frivillige redningsmannskaper"],
        ].map(([k, e]) => new TableRow({
          cantSplit: true,
          children: [
            new TableCell({ width: { size: 2800, type: WidthType.DXA }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: k, font: FONT_BODY, size: 21, color: COLOR_PRIMARY, bold: true })] })] }),
            new TableCell({ width: { size: 6838, type: WidthType.DXA }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: e, font: FONT_BODY, size: 21, color: COLOR_TEXT })] })] }),
          ],
        })),
      ],
    }),
    new Paragraph({ spacing: { after: 240 } }),

    H2("Format og oppbevaring"),
    P("Samme hybridmodell som T3-registeret. Daglig drift krever digital registrering. Samtidig må informasjonen være tilgjengelig når strøm og ekom svikter — som er nettopp scenarioene kartet skal brukes i. Digital arbeidskopi (primær) lagres på ekstern, kryptert disk med backupdisk. Ikke skylagring, ikke nettverkstilkoblet system. Papirsikkerhetskopi skrives ut jevnlig (minimum hver tredje måned) og oppbevares i låst arkiv på møteplassen sammen med T3-registeret."),
    H3("Aktivt samtykke og personvern — fleksible bruksområder"),
    P("Registrering er frivillig opt-in. Personen skriver under på hva som registreres, hvordan opplysningene oppbevares, hvilke bruksområder samtykket dekker (personen velger selv én eller flere), og rett til når som helst å bli slettet."),
    P("Bruksområder personen krysser av for:"),
    bullet("Aktivering ved krise (obligatorisk minimum)"),
    bullet("Øvelser og beredskapsforberedelser"),
    bullet("Forespørsel om kursholder-/instruktørrolle (T4, T5)"),
    bullet("Oversetterhjelp eller flerspråklig formidling (T10)"),
    bullet("Generell rekruttering til frivilligsentralen"),
    P("Dette gir personen reell kontroll over egen kompetanse, samtidig som frivilligsentralen kan aktivere ressursen i flere sammenhenger der personen ønsker det. Den minste fellesnevneren — aktivering ved krise — er den eneste obligatoriske. T7-registeret er kvalitativt enklere GDPR-messig enn T3, fordi det er personen selv som tilbyr sin kompetanse — det er ikke et register over sårbarhet."),

    H2("Strategisk verdi — utover krisehåndtering"),
    bullet("Argumenterer for frivilligsentralens unike rolle overfor bydel og kommune — vi har data ingen andre har"),
    bullet("Styrker kommunens ressursoversikt (jf. Forskrift § 4c) konkret"),
    bullet("Bygger relasjoner mellom frivilligsentralen og lokale fagpersoner"),
    bullet("Grunnlag for løpende rekruttering av kursholdere, instruktører og spesialfrivillige"),

    maltallsboks(
      [
        "Kompetanseansvarlig utpekt og opplært (kan være samme person som T3-registeransvarlig)",
        "Hybrid lagringsmodell operasjonell — kryptert ekstern disk med backup, samtykkemal godkjent",
        "40–60 personer registrert i kartet",
        "Alle 7 kompetansekategorier representert med minst én person",
        "Førsterunden av telefonsjekk/oppdatering gjennomført",
      ],
      "Kartet er testet i øvelse — minimum tre kompetansekategorier mobilisert under simulert scenario."
    ),
  ];
}

function t8() {
  return [
    ...tiltakHeader("T8", "Koordinering av spontanfrivillige", 5, 2, true),

    H2("Hvorfor høy prioritet"),
    P("Alle andre tiltak bygger kapasitet i forkant. T8 sørger for at man faktisk kan nyttiggjøre kapasiteten som dukker opp spontant. Lav kompleksitet og høy effekt — men ofte glemt. COVID-erfaringen viser at mangel på T8 gjør at ressurser går tapt nettopp når behovet er størst."),

    H2("Formål"),
    P("Etablere et enkelt, robust system for å ta imot, registrere og nyttiggjøre frivillige som møter opp under en krise uten forhåndsdefinert rolle."),

    H2("Hvorfor dette er kritisk"),
    P("Ved enhver større krise vil mennesker møte opp og ville hjelpe. Uten et system blir de et koordineringsproblem fremfor en ressurs. Uten et system vet frivillige ikke hvem de skal henvende seg til; folk med kritisk kompetanse forsvinner i mengden; beredskapsfrivillige (T2) bruker tid på å håndtere kaos fremfor jobben sin; og frivillige blir frustrerte og går hjem."),

    faktaboks(
      "COVID-19 (2020) — dokumentasjon",
      "Under pandemien i 2020 oppstod et massivt tilbud av spontanfrivillige som ville hjelpe hjemmeboende eldre og sårbare. Frivilligsentralen hadde ingen plan for situasjonen. Smittefare begrenset hva som var mulig fysisk. Det avgjørende hinderet var at GDPR forhindret tilgang til kontaktinformasjon — frivillige hadde ingen å kontakte. Ressursen gikk tapt. Dette er like mye et argument for T3 som for T8: spontanfrivillige uten register å jobbe fra er som et korps uten kart.",
      "Erfaringsgrunnlag fra Vestre Aker Frivilligsentral, 2020"
    ),

    H2("Mottakssjef — nøkkelrollen"),
    P("Hver møteplass (T1) har minst to utpekte mottakssjefer blant beredskapsfrivillige (T2). Mottakssjefen tar imot alle som ankommer uten rolle og kanaliserer dem raskt og riktig."),
    PR([
      { text: "Prinsipp: robusthet og bred deltakelse, ikke effektivitet. ", bold: true, color: COLOR_PRIMARY },
      "Alle roller bør ha minst to personer for å gjøre systemet mindre sårbart. For frivilligheten er det dessuten et mål å engasjere flest mulig — robusthet og bredde i deltakelsen vektes høyere enn driftsoptimalisering. Det betyr at vi heller har flere mottakssjefer enn strengt nødvendig.",
    ]),
    PR([
      { text: "Egnet rolle for frivillige med fysisk funksjonsnedsettelse. ", bold: true, color: COLOR_PRIMARY },
      "Mottakssjef-rollen er stillesittende, krever god kommunikasjon, oversikt og rolig håndtering — ikke fysisk mobilitet. Den er derfor godt egnet for frivillige med fysisk funksjonsnedsettelse (f.eks. rullestolbrukere) som ellers kan være vanskelig å inkludere i akutt beredskapsarbeid. Dette gir en ellers sårbar gruppe en meningsfull og verdsatt rolle i lokal beredskap. Forutsetning: transport til møteplassen må planlegges inn som del av aktiveringsrutinen.",
    ]),
    P("Frivilligsentralen har allerede frivillige som gjør mottak av tjenestemottakere. Kompetansen for å ta imot mennesker, registrere behov og koordinere videre er direkte overførbar til mottakssjef-rollen — ingen ny kompetanse trengs å bygges fra bunnen."),

    H2("Mottaksprosedyre — digital primær, papir som sikkerhetskopi"),
    P("Prosedyren skal være digital som primærløsning, men hvert utfylt registreringsskjema skrives ut umiddelbart. Det sikrer både rask innsamling og fysisk dokumentasjon som overlever digital svikt. Forhåndstrykte tomme skjemaer ligger klar på mottakspunktet i god mengde — slik at registrering kan fortsette manuelt hvis digitale løsninger ikke er tilgjengelig ved ankomst eller faller ut underveis. Samme prinsipp som T3 og T7: digital primær, papir som sikkerhetskopi."),
    PR([{ text: "Steg 1 — Registrering: ", bold: true, color: COLOR_PRIMARY }, "Alle som ankommer uten forhåndsrolle møter mottakssjefen. Registreringsskjema: navn og telefon, bor i nærheten (utholdenhet), relevant kompetanse, tilgjengelighet, fysisk form og begrensninger."]),
    PR([{ text: "Steg 2 — Oppgavetildeling: ", bold: true, color: COLOR_PRIMARY }, "Etter oppgavekategori — oppsøkende hjelp (sjekke T3-registrerte), vakt og informasjon, logistikk og transport, praktisk hjelp, administrasjon, støtte og omsorg."]),
    PR([{ text: "Steg 3 — Skiftstruktur: ", bold: true, color: COLOR_PRIMARY }, "Skiftlengde 3–4 timer. Tydelig avleveringsprosedyre."]),
    PR([{ text: "Steg 4 — Utsjekk: ", bold: true, color: COLOR_PRIMARY }, "Alle sjekker ut ved avslutning. Kort samtale: hvordan har de det? Kontaktinfo bevares for rekruttering til T2."]),

    H2("Utstyr for mottakspunktet"),
    P("Inngår i møteplassens beredskapslager (jf. T1):"),
    bullet("Nettbrett/laptop med batteribackup for digital registrering (primær)"),
    bullet("Skriver/multifunksjonsenhet for umiddelbar utskrift av registrerte skjemaer"),
    bullet("Forhåndstrykte tomme registreringsskjemaer i god mengde (sikkerhetskopi ved digital svikt)"),
    bullet("Oppgaveoversikt på stor tavle/plakat"),
    bullet("Skiftlogg (papir)"),
    bullet("Navnelapper/vester for å skille roller visuelt"),

    maltallsboks(
      [
        "Mottakssjef utpekt og opplært per møteplass (minimum 2 personer per sentral, så rollen er dekket ved fravær)",
        "Komplett mottakssett etablert på begge møteplasser — digital primær (nettbrett/laptop + skriver) og papirbackup (forhåndstrykte skjemaer, oppgaveoversikt, skiftlogg, navnelapper/vester)",
        "T8-prosedyren integrert som Modul 7 i T2-grunnkurset",
        "Mottaksprosedyren testet i øvelse med simulerte spontanfrivillige",
      ],
      "Mottakssjef kan ta imot, registrere og oppgavetildele 10+ spontanfrivillige innen 30 minutter uten digitale verktøy."
    ),
  ];
}

function t9() {
  return [
    ...tiltakHeader("T9", "Ungdom og skole i beredskap", 3, 3, false),
    H2("Formål"),
    P("Bygge beredskapsbevissthet og praktisk kompetanse hos barn og unge i Vestre Aker — gjennom samarbeid med lokale skoler. Ungdom som lærer beredskap tar kunnskapen med hjem til foreldre og søsken, og blir bærere av beredskapskulturen i neste generasjon."),

    H2("Den langsiktige gevinsten — multiplikatoreffekten"),
    P("Hvert barn som deltar på et beredskapsopplegg bringer kunnskapen videre til en hel husstand. T9 er derfor den mest effektive måten å nå voksne som ellers ikke ville oppsøke et beredskapskurs (T4). Forskning på folkehelseprogrammer viser konsekvent at barn er sterke endringsagenter inn i familien — «pappa, har vi vannflasker hjemme?» er et budskap som treffer på en måte et offentlig kampanjebudskap ikke gjør. Dette gjør T9 til et strategisk multiplikatortiltak, ikke bare et tiltak rettet mot målgruppen barn og unge."),

    H2("Strategisk inngangspunkt — «Innsats for andre»"),
    P("Ungdomsskolene har valgfaget «Innsats for andre» — et fag spesifikt designet for at elever skal yte strukturert samfunnsbidrag. Skolene er aktivt på utkikk etter konkrete tiltak og aktiviteter som passer inn. T9 er en ferdig pakke som fyller et reelt behov skolen allerede har — ikke noe vi ber dem prioritere på toppen av eksisterende arbeid. Dette er den enkleste inngangen til skolesektoren og bør være primær startflate for T9."),
    P("Frivilligsentralen har allerede etablerte relasjoner til to ungdomsskoler i bydelen som er målgruppen for T9. Det betyr at vi ikke starter med kald kontakt — vi bygger på eksisterende tillit. Dette senker kompleksiteten reelt og gjør oppstart raskere enn et generisk T9-tiltak ville være andre steder."),

    faktaboks(
      "Teen CERT — amerikansk modell",
      "I USA har Community Emergency Response Team (CERT) en ungdomsversjon kjent som Teen CERT — strukturert opplæring for ungdom 13–18 år i førstehjelp, brannvern, lett søk og redning. Modellen er veletablert i mange amerikanske skoledistrikter, og er den nærmeste internasjonale parallellen til Format 1 i T9 (juniorberedskapsfrivillige via «Innsats for andre»).",
      "FEMA — Community Emergency Response Team"
    ),

    H2("Frivilligsentralens rolle — tilbyder og tilrettelegger"),
    P("Frivilligsentralen eier ikke skoleopplegget. Skolen og lærerne styrer hva som skjer i klasserommet. Frivilligsentralen er initiativtaker som tar kontakten, praktisk tilrettelegger med utstyr og kompetanse skolen mangler, ressursbank for lærere som vil ta tematikken inn i undervisningen, og vert for aktiviteter som ikke passer i klasserommet (besøk på møteplassen, øvelser)."),

    H2("Tre konkrete formater"),
    PR([{ text: "Format 1 — Juniorberedskapsfrivillige via «Innsats for andre» (primærspor): ", bold: true, color: COLOR_PRIMARY }, "Strukturert opplegg integrert i valgfaget. Modell inspirert av Teen CERT, tilpasset norsk skole og pilotens kontekst. Grunnopplæring tilsvarende T2 grunnkurs justert for 13–15-åringer. Reelle oppgaver i løpet av skoleåret: bistå på arrangementer, delta på øvelser, drive informasjonsarbeid på skolen, hjelpe til på beredskapsdager. Ikke aktiv utrykningsrolle ved reell krise — men kompetansebygger som forbereder neste generasjons beredskapsfrivillige. Naturlig vurderingsgrunnlag for skolen i faget."]),
    PR([{ text: "Format 2 — Beredskapsdag på skolene (halvdag): ", bold: true, color: COLOR_PRIMARY }, "Tematisk dag arrangert i samarbeid med skolen, gjerne med juniorberedskapsfrivillige som medarrangører. Stasjoner elever roterer mellom: førstehjelp (HLR-demo, sårbehandling), Hva har dere hjemme? (interaktiv quiz), møteplassen i bydelen, familieoppdrag de tar med hjem."]),
    PR([{ text: "Format 3 — Læringsopplegg til andre lærere: ", bold: true, color: COLOR_PRIMARY }, "Ferdige undervisningsmoduler (1–2 timers opplegg) som skolen kan bruke uten ekstern bistand. Knyttes til kompetansemål i fagfornyelsen der det er mulig. DSB har grunnmateriale — oppgaven er lokal tilpasning og pakketering, ikke utvikling fra bunnen."]),

    H2("Sekvens"),
    numbered("Først: Møte med de to ungdomsskolene vi allerede har relasjon til. Presenter T9 som tilbud inn i «Innsats for andre»."),
    numbered("Deretter: Etabler Format 1 (juniorberedskapsfrivillige) ved minst én av skolene som pilot i et skoleår."),
    numbered("Parallelt: Gjennomfør Format 2 (beredskapsdag) ved samme skole, gjerne med juniorberedskapsfrivillige som medarrangører."),
    numbered("Senere: Skaler Format 3 (læringsopplegg) til andre skoler og lærere i bydelen når piloten er evaluert."),

    maltallsboks(
      [
        "Etablert avtale med begge ungdomsskolene",
        "Format 1 piloteres ved minimum én skole gjennom et skoleår",
        "1 beredskapsdag (Format 2) gjennomført",
      ],
      "Juniorberedskapsfrivillige fra Format 1 har deltatt aktivt som medarrangører på minst én beredskapsdag, og skolen ønsker å videreføre opplegget neste skoleår."
    ),
  ];
}

function t10() {
  return [
    ...tiltakHeader("T10", "Flerspråklig beredskap", 3, 2, false),
    H2("Formål"),
    P("Sikre at innbyggere i bydelen med andre morsmål enn norsk når frem av beredskapsinformasjon, møteplassinformasjon og varslingsprosedyrer — både i fredstid og ved aktiv krise. Standardkommunikasjon på norsk når ikke alle grupper i lokalsamfunnet, og de som ikke nås er ofte de mest sårbare."),

    H2("Begrunnelse — politisk og faglig"),
    bullet("EU Preparedness Union Strategy (mars 2025) anbefaler eksplisitt flerspråklig kommunikasjon som en kjernekomponent i nasjonal og lokal beredskap"),
    bullet("REAL-prosjektet dokumenterer at sårbare og marginaliserte grupper systematisk faller utenfor standardkommunikasjon på majoritetsspråket — og at frivillige organisasjoner med lokalkunnskap er best plassert til å løse dette"),
    bullet("KRB Oslo 2025 påpeker at varsling og informasjonsformidling ved krise er en av de mest sårbare ledd i kommunens beredskap; flerspråklighet er en konkret del av denne sårbarheten"),

    H2("Den bærende logikken — bruk eksisterende materiell"),
    P("DSB har allerede oversatt mye beredskapsmateriale til en rekke språk. Oppgaven for frivilligsentralen er ikke å lage nytt materiell fra bunnen — det ville være sløsing med ressurser og duplisering av nasjonal kapasitet. Oppgaven er å innhente eksisterende oversatt materiale fra DSB, tilpasse lokalt (bytte ut generisk informasjon med Vestre Aker-spesifikk), kvalitetssikre oversettelsene med lokale tolker (jf. T7), og distribuere gjennom kanaler som faktisk når målgruppen."),

    H2("Strategisk inngangspunkt — Bydelsmødre Vestre Aker"),
    P("Frivilligsentralen har godt etablert samarbeid med Bydelsmødre Vestre Aker. Dette er kanskje den viktigste enkeltrelasjonen for T10."),

    faktaboks(
      "Bydelsmødre — bro mellom minoritetsmiljøer og det offentlige",
      "Bydelsmødre er per definisjon broer mellom minoritetsmiljøer og det norske offentlige systemet — kvinner med minoritetsbakgrunn som er kurset til å være tillitspersoner og informasjonsformidlere i sine egne miljøer. De har direkte tillit i de språk- og kulturgruppene T10 skal nå, kunnskap om hvilke kommunikasjonsformer som faktisk fungerer, eksisterende kanaler ut til miljøene, og erfaring med å oversette offentlig informasjon til praktisk forståelig form. Dette gjør at T10 i Vestre Aker har et reelt operasjonelt fundament fra dag én.",
      "Bydelsmødre Norge — bydelsmodre.no"
    ),

    H2("Vestre Aker-kontekst og pilotlogikken"),
    P("Vestre Aker har lavere andel innbyggere med innvandrerbakgrunn enn østlige bydeler i Oslo. Det betyr at T10 i selve piloten er mindre kritisk i ren rekkevidde enn det ville være i for eksempel Stovner eller Grorud. Men T10 må likevel være med — flerspråklige innbyggere finnes også i Vestre Aker, og de skal ikke falle utenfor. Og T10 er strategisk viktig for pilotlogikken: en modell som ikke håndterer flerspråklighet er ikke overførbar til bydeler der dette er kritisk. Hele pilotambisjonen forutsetter at modellen fungerer i mer flerspråklige bydeler."),

    H2("Prioritering av språk"),
    P("Frivilligsentralen har ikke i dag oversikt over hvilke språkgrupper som er størst i bydelen. Informasjonen må skaffes som første konkrete steg i T10:"),
    numbered("Hente språkstatistikk fra bydelsadministrasjonen / SSB / Oslo kommunes statistikkbank"),
    numbered("Supplere med kvalitativ vurdering fra Bydelsmødre Vestre Aker — de har bakkenær kunnskap om hvilke språkgrupper som er aktive"),
    numbered("Krysse mot DSBs tilgjengelige oversettelser — flertallet av relevante språk er allerede dekket"),
    numbered("Identifisere eventuelle gap (språk uten DSB-materiale)"),
    numbered("Mobilisere tolker fra T7-kompetansekartet til å dekke gapene"),

    H2("Distribusjonskanaler"),
    PR([{ text: "Etablerte kanaler (relasjon på plass): ", bold: true, color: COLOR_SECONDARY },
        "Bydelsmødre Vestre Aker (primær), bydelens informasjonskanaler, borettslag med høy andel flerspråklige beboere (via T6), skoler (via T9 — barn formidler til foreldre)."]),
    PR([{ text: "Potensielle kanaler (må etableres): ", bold: true, color: COLOR_ACCENT },
        "Religiøse forsamlinger og kulturforeninger, helsestasjoner og fastleger, sosiale medier rettet mot språkgrupper."]),

    P("Det viktigste er muntlig formidling fra tillitsperson på morsmål. Skriftlig materiale er støtte — ikke hovedkanal — for grupper med lavere lesekompetanse. Bydelsmødre er allerede modellen for nettopp denne typen formidling."),

    maltallsboks(
      [
        "Språkprioritering ferdigstilt for bydelen (3–5 språk identifisert)",
        "Kjernepakka (de 3 høyt prioriterte materialene) oversatt og lokaltilpasset for prioriterte språk",
        "Distribusjon etablert via Bydelsmødre og minimum 2 andre kanaler",
      ],
      "Flerspråklige innbyggere i minimum 2 prioriterte språkgrupper bekrefter (via Bydelsmødre eller annen tillitsperson) at de kjenner møteplassen og vet hva den brukes til ved krise."
    ),
  ];
}

function veienVidere() {
  return [
    H1("Veien videre"),

    H2("Fase 1 — Dialog med bydel og kommune"),
    P("Fase 1 består av to deler: utarbeidelse av denne skissen, og formell dialog med beredskapsansvarlig i Bydel Vestre Aker og Oslo kommunes Beredskapsetat. Skissen er nå ferdig. Den formelle dialogen er ikke startet — det er neste umiddelbare skritt."),

    P("Hovedformålet med fase 1-dialogen er å oppnå:"),
    bullet("Felles forståelse av frivilligsentralenes rolle som supplement til Persbråten VGS"),
    bullet("Avklaring av juridisk grunnlag for T3 (kommuneadvokat-prosess)"),
    bullet("Avtale om kanal for kontakt med DSB, Sivilforsvaret, Norsk Folkehjelp og Røde Kors via Oslo kommunes beredskapsavdeling"),
    bullet("Foreløpig vurdering av hvilke av de ti tiltakene som prioriteres i fase 2"),
    bullet("Formell inklusjon av frivilligsentralene i kommunens ressursoversikt (jf. Forskrift § 4c)"),

    H2("Fase 2 — Handlingsplan"),
    P("Fase 2 starter når forankringen er på plass. Handlingsplanen vil omfatte:"),
    bullet("Konkret tidsplan med milepæler for hvert prioritert tiltak"),
    bullet("Ressursbehov: bemanning, materiell, lokaler, finansiering"),
    bullet("Ansvarsfordeling mellom frivilligsentralene, bydelen, kommunen og samarbeidspartnere"),
    bullet("Finansieringsstrategi: tilskudd, sentrale beredskapsmidler, stiftelser, egen inntektsgenerering"),
    bullet("Avklaring av hvilken frivilligsentral som først løftes til T1 Nivå 2 — basert på geografisk dekning, lokalets egnethet og finansiering"),
    bullet("Plan for løpende evaluering og justering"),

    H2("Overføring — fra Vestre Aker til andre bydeler og kommuner"),
    P("Piloten skal dokumenteres på en måte som gjør den overførbar. Dette betyr at maler, prosesser, kursmateriale, samarbeidsavtaler, samtykkemaler og lærdommer samles og deles åpent. Modellen testes deretter i andre bydeler i Oslo (særlig i bydeler med andre demografiske profiler, der T10 og T9 vil ha sterkere effekt), og deretter i andre kommuner."),
    P("Resultatet — om piloten lykkes — er en norsk modell for sivilsamfunnets rolle i lokal beredskap. Det er det dette dokumentet handler om."),
  ];
}

function vedleggOrdliste() {
  const ordliste = [
    ["AED", "Automated External Defibrillator — hjertestarter"],
    ["CERT", "Community Emergency Response Team (USA) — modell for trente lokale beredskapsfrivillige"],
    ["DSB", "Direktoratet for samfunnssikkerhet og beredskap"],
    ["Ekom", "Elektronisk kommunikasjon (mobilnett, internett, telefoni)"],
    ["EU PUS", "EU Preparedness Union Strategy (mars 2025)"],
    ["FRG", "Frivilliga Resursgrupper (Sverige) — svensk modell for kommunalt forankrede beredskapsfrivillige"],
    ["HLR", "Hjerte-lungeredning"],
    ["KRB", "Kommunalt risikobilde (her: Oslo 2025)"],
    ["NOU", "Norges offentlige utredninger"],
    ["NRRL", "Norsk Radio Relæ Liga — radioamatørorganisasjon med nødkommunikasjonskapasitet"],
    ["REAL", "Resilience through Engagement and Action by Local volunteers — EU-prosjekt"],
    ["Sivilforsvaret", "Statlig beredskapsorganisasjon under Justis- og beredskapsdepartementet"],
  ];

  return [
    H1("Vedlegg A — Ordliste"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [2200, 7438],
      borders: tableBorders,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            new TableCell({ width: { size: 2200, type: WidthType.DXA }, shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: "Term", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })] }),
            new TableCell({ width: { size: 7438, type: WidthType.DXA }, shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: "Forklaring", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })] }),
          ],
        }),
        ...ordliste.map(([term, forklaring]) => new TableRow({
          cantSplit: true,
          children: [
            new TableCell({ width: { size: 2200, type: WidthType.DXA }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: term, font: FONT_HEADING, size: 21, color: COLOR_PRIMARY, bold: true })] })] }),
            new TableCell({ width: { size: 7438, type: WidthType.DXA }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: forklaring, font: FONT_BODY, size: 21, color: COLOR_TEXT })] })] }),
          ],
        })),
      ],
    }),
    new Paragraph({ spacing: { after: 240 } }),
  ];
}

function vedleggKilder() {
  // Hver kilde: [navn, beskrivelse, URL eller null]
  const kilder = [
    ["KRB Oslo 2025",
     "Kommunalt risikobilde for Oslo kommune 2025 — beredskapsfaglig dimensjonering",
     "https://www.oslo.kommune.no/beredskap/"],
    ["Totalberedskapsmeldingen",
     "Meld. St. om norsk totalberedskap — nasjonal politikk for sivilsamfunnsrolle",
     "https://www.regjeringen.no/no/dokumenter/meld.-st.-9-20242025/id3082147/"],
    ["NOU 2023:17",
     "«Nå er det alvor» — utredning om totalberedskap og sivilsamfunn",
     "https://www.regjeringen.no/no/dokumenter/nou-2023-17/id2992254/"],
    ["EU Preparedness Union Strategy",
     "EU-strategi (mars 2025) for europeisk beredskap — internasjonal kontekst",
     "https://commission.europa.eu/topics/preparedness-union-strategy_en"],
    ["EU Preparedness Union Action Plan",
     "EUs handlingsplan (mars 2025) for operasjonalisering av strategien",
     "https://commission.europa.eu/topics/preparedness-union-strategy_en"],
    ["REAL Policy Recommendations",
     "Anbefalinger fra REAL-prosjektet — frivillige i lokal beredskap",
     "https://filer.frivilligsentralen.org/dokumenter/REAL/REAL%20Policy%20Recommendations_ENG.pdf"],
    ["REAL Best Practice Manual",
     "Beste praksis-håndbok fra REAL-prosjektet — operasjonell veileder for frivillige i beredskap",
     "https://filer.frivilligsentralen.org/dokumenter/REAL/bpm_eng.pdf"],
    ["Frivilliga Resursgruppen (FRG, Sverige)",
     "Svensk modell for kommunalt forankrede beredskapsfrivillige — direkte forbilde for T2",
     "https://civil.se/om-oss/frivilliga-resursgruppen/"],
    ["Bajek m.fl. — Natural Hazards (Springer)",
     "Forskningsartikkel om japanske jishu-bosai-soshiki — frivillige nabolagsberedskapsgrupper. Forbilde for T6 Nivå 2",
     "https://link.springer.com/article/10.1007/s11069-007-9107-4"],
    ["Sivilbeskyttelsesloven",
     "Lov om kommunal beredskapsplikt mm. — primær rettskilde",
     "https://lovdata.no/dokument/NL/lov/2010-06-25-45"],
    ["Forskrift om kommunal beredskapsplikt",
     "Forskrift med § 4c om ressursoversikt over frivillige",
     "https://lovdata.no/dokument/SF/forskrift/2011-08-22-894"],
    ["DSB — Egenberedskap",
     "DSBs anbefalinger om egenberedskap (1 ukes-prinsippet)",
     "https://www.sikkerhverdag.no/egenberedskap/"],
    ["DSB — Kommunal beredskapsplikt (veileder)",
     "Veileder til forskriften — operasjonell tolkning",
     "https://www.dsb.no/lover/risiko-sarbarhet-og-beredskap/veiledning-til-forskrift/veileder-til-forskrift-om-kommunal-beredskapsplikt/"],
    ["Oslo overordnet beredskapsplan",
     "Oslo kommunes plan med fire nivåer — definerer frivilligsentralenes plass",
     "https://www.oslo.kommune.no/politikk/sikkerhet-og-beredskap/"],
    ["Beredskap.oslo.kommune.no",
     "Oslo kommunes offisielle beredskapsinformasjon — bekrefter Persbråten-gapet",
     "https://www.oslo.kommune.no/beredskap/"],
    ["Aasland — Frivillighet i samvirke",
     "«Frivillige organisasjoners medvirkning i redningstjeneste og beredskapsarbeid» — utgitt av Hovedredningssentralen",
     "https://www.hovedredningssentralen.no/wp-content/uploads/2023/10/Frivillighet-i-beredskapssamvirke_Aasland.pdf"],
    ["Internasjonale modeller (intern oversikt)",
     "Sammenstilling av FRG, UK Resilience Hubs, Teen CERT, Japan-modellen m.fl.",
     null],
  ];

  // Hjelpefunksjon: bygg kilde-celle med eller uten hyperkobling
  function kildeCelle(navn, url) {
    if (url) {
      return new TableCell({
        width: { size: 3800, type: WidthType.DXA },
        borders: tableBorders,
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        children: [new Paragraph({
          children: [new ExternalHyperlink({
            link: url,
            children: [new TextRun({
              text: navn,
              font: FONT_HEADING,
              size: 21,
              color: COLOR_PRIMARY,
              bold: true,
              underline: { type: "single", color: COLOR_PRIMARY },
            })],
          })],
        })],
      });
    }
    return new TableCell({
      width: { size: 3800, type: WidthType.DXA },
      borders: tableBorders,
      margins: { top: 120, bottom: 120, left: 200, right: 200 },
      children: [new Paragraph({
        children: [new TextRun({
          text: navn,
          font: FONT_HEADING,
          size: 21,
          color: COLOR_PRIMARY,
          bold: true,
        })],
      })],
    });
  }

  return [
    H1("Vedlegg B — Kildemateriale"),
    P("Dokumentet bygger på det følgende kildematerialet. Kildenavn med understrek er klikkbare lenker til kildens primæradresse. Detaljerte referanser og uttrekk er tilgjengelig i prosjektets kildemappe."),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [3800, 5838],
      borders: tableBorders,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            new TableCell({ width: { size: 3800, type: WidthType.DXA }, shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: "Kilde", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })] }),
            new TableCell({ width: { size: 5838, type: WidthType.DXA }, shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: "Innhold og betydning", font: FONT_HEADING, size: 22, bold: true, color: "FFFFFF" })] })] }),
          ],
        }),
        ...kilder.map(([k, b, url]) => new TableRow({
          cantSplit: true,
          children: [
            kildeCelle(k, url),
            new TableCell({ width: { size: 5838, type: WidthType.DXA }, borders: tableBorders, margins: { top: 120, bottom: 120, left: 200, right: 200 }, children: [new Paragraph({ children: [new TextRun({ text: b, font: FONT_BODY, size: 21, color: COLOR_TEXT })] })] }),
          ],
        })),
      ],
    }),
    new Paragraph({ spacing: { after: 240 } }),
  ];
}

// ============================================================
// DOKUMENT
// ============================================================
const doc = new Document({
  creator: "Vestre Aker Frivilligsentral",
  title: "Beredskapsplan for Frivilligsentralene Vestre Aker",
  description: "Pilotdokument for lokal beredskap og sivil motstandsdyktighet",
  styles: {
    default: {
      document: {
        run: { font: FONT_BODY, size: 22, color: COLOR_TEXT },
        paragraph: { spacing: { line: 280 } },
      },
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: FONT_HEADING, color: COLOR_PRIMARY },
        paragraph: { spacing: { before: 480, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: FONT_HEADING, color: COLOR_PRIMARY },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: FONT_HEADING, color: COLOR_PRIMARY },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }, {
          level: 1, format: LevelFormat.BULLET, text: "◦",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1080, hanging: 360 } } },
        }],
      },
      { reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  sections: [
    // Section 1 — forside (uten header/footer)
    {
      properties: {
        page: {
          size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      children: forside(),
    },
    // Section 2 — resten av dokumentet (med header og footer)
    {
      properties: {
        page: {
          size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      headers: {
        default: new Header({
          children: [
            // Vannmerker — diskret bakgrunn på alle sider. Bruker floating + behindDocument.
            new Paragraph({
              spacing: { after: 0 },
              children: [
                new ImageRun({
                  type: "png",
                  data: fs.readFileSync("C:/codex/beredskap/presentasjon/figurer/vannmerke_VAFS.png"),
                  transformation: { width: 240, height: 160 },
                  floating: {
                    horizontalPosition: {
                      relative: HorizontalPositionRelativeFrom.PAGE,
                      offset: 1800000, // ca 5 cm fra venstre kant (EMU: 914400 = 1 inch)
                    },
                    verticalPosition: {
                      relative: VerticalPositionRelativeFrom.PAGE,
                      offset: 7200000, // ca 20 cm fra topp
                    },
                    behindDocument: true,
                  },
                  altText: {
                    title: "Vannmerke VAFS",
                    description: "Diskret vannmerke Vestre Aker Frivilligsentral",
                    name: "watermark_VAFS",
                  },
                }),
                new ImageRun({
                  type: "png",
                  data: fs.readFileSync("C:/codex/beredskap/presentasjon/figurer/vannmerke_OVFS.png"),
                  transformation: { width: 240, height: 160 },
                  floating: {
                    horizontalPosition: {
                      relative: HorizontalPositionRelativeFrom.PAGE,
                      offset: 4600000, // høyresiden
                    },
                    verticalPosition: {
                      relative: VerticalPositionRelativeFrom.PAGE,
                      offset: 7200000,
                    },
                    behindDocument: true,
                  },
                  altText: {
                    title: "Vannmerke OVFS",
                    description: "Diskret vannmerke Oslo Vest Frivilligsentral",
                    name: "watermark_OVFS",
                  },
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER, space: 1 } },
              children: [new TextRun({
                text: "Beredskapsplan Frivilligsentralene Vestre Aker",
                font: FONT_HEADING, size: 18, italics: true, color: COLOR_MUTED,
              })],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "— ", font: FONT_BODY, size: 22, color: COLOR_PRIMARY, bold: true }),
                new TextRun({ children: [PageNumber.CURRENT], font: FONT_BODY, size: 22, color: COLOR_PRIMARY, bold: true }),
                new TextRun({ text: " —", font: FONT_BODY, size: 22, color: COLOR_PRIMARY, bold: true }),
              ],
            }),
          ],
        }),
      },
      children: [
        ...tocSeksjon(),
        ...sammendrag(),
        ...bakgrunn(),
        ...konseptOgPrinsipper(),
        ...visuellOversikt(),

        // Innledning til tiltaksdelen
        H1("De ti tiltakene"),
        P("De følgende ti tiltakene er kjernen i piloten. Hvert tiltak er beskrevet med formål, hovedinnhold, koblinger til andre tiltak, og måltall for første pilotår. Tiltakene er nummerert i logisk rekkefølge, ikke etter prioritet — for prioritering se prioriteringsmatrisen i forrige kapittel."),

        ...t1(),
        ...t2(),
        ...t3(),
        ...t4(),
        ...t5(),
        ...t6(),
        ...t7(),
        ...t8(),
        ...t9(),
        ...t10(),

        ...veienVidere(),
        ...vedleggOrdliste(),
        ...vedleggKilder(),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buffer => {
  const out = "C:/codex/beredskap/presentasjon/Beredskapsplan_Vestre_Aker.docx";
  fs.writeFileSync(out, buffer);
  console.log("OK — dokument generert: " + out);
}).catch(err => {
  console.error("FEIL:", err);
  process.exit(1);
});
