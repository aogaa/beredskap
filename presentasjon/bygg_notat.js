/**
 * Bygger 4-siders kompakt notat for bydel og kommune.
 * Selvstendig dokument — oppsummerer hoveddokumentet, men kan leses alene.
 */

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType,
  LevelFormat, ExternalHyperlink, TableOfContents, HeadingLevel,
  BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageNumber, PageBreak,
} = require("docx");

// ============================================================
// DESIGN — samme profil som hoveddokumentet
// ============================================================
const COLOR_PRIMARY = "1A2B47";
const COLOR_ACCENT = "7B2D2F";
const COLOR_SECONDARY = "2D4F3D";
const COLOR_TEXT = "2C3E50";
const COLOR_MUTED = "6B7280";
const COLOR_BG_FACT = "F0EEE6";
const COLOR_BG_HIGHLIGHT = "E8E6DC";
const COLOR_BORDER = "C4BDA8";
const COLOR_GRID = "D0CCB8";

const FONT_HEADING = "Cambria";
const FONT_BODY = "Calibri";

// A4 med litt strammere marger (1,5 cm) for å få plass til mer på 4 sider
const PAGE_WIDTH = 11906;
const PAGE_HEIGHT = 16838;
const MARGIN = 850; // ca 1,5 cm
const CONTENT_WIDTH = PAGE_WIDTH - 2 * MARGIN;

const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER };
const accentBorderLeft = { style: BorderStyle.SINGLE, size: 24, color: COLOR_ACCENT };
const factBorderLeft = { style: BorderStyle.SINGLE, size: 24, color: COLOR_SECONDARY };
const tableBorders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_GRID },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_GRID },
  left: { style: BorderStyle.SINGLE, size: 4, color: COLOR_GRID },
  right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_GRID },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: COLOR_GRID },
  insideVertical: { style: BorderStyle.SINGLE, size: 4, color: COLOR_GRID },
};

// ============================================================
// HJELPEFUNKSJONER (kompakte, færre vertikale mellomrom)
// ============================================================
function P(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 100, line: 270, ...(opts.spacing || {}) },
    alignment: opts.alignment || AlignmentType.JUSTIFIED,
    children: [new TextRun({
      text,
      font: opts.font || FONT_BODY,
      size: opts.size || 21, // 10,5pt — kompakt men leselig
      color: opts.color || COLOR_TEXT,
      bold: opts.bold || false,
      italics: opts.italics || false,
    })],
  });
}

function PR(runs, opts = {}) {
  return new Paragraph({
    spacing: { after: 100, line: 270, ...(opts.spacing || {}) },
    alignment: opts.alignment || AlignmentType.JUSTIFIED,
    children: runs.map(r => {
      if (typeof r === "string") {
        return new TextRun({ text: r, font: FONT_BODY, size: 21, color: COLOR_TEXT });
      }
      return new TextRun({
        text: r.text,
        font: r.font || FONT_BODY,
        size: r.size || 21,
        color: r.color || COLOR_TEXT,
        bold: r.bold || false,
        italics: r.italics || false,
      });
    }),
  });
}

function H1(text) {
  return new Paragraph({
    spacing: { before: 0, after: 100 },
    keepNext: true,
    keepLines: true,
    children: [new TextRun({
      text,
      font: FONT_HEADING, size: 32, bold: true, color: COLOR_PRIMARY,
    })],
  });
}

function H2(text) {
  return new Paragraph({
    spacing: { before: 240, after: 80 },
    keepNext: true,
    keepLines: true,
    children: [new TextRun({
      text,
      font: FONT_HEADING, size: 24, bold: true, color: COLOR_PRIMARY,
    })],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60, line: 260 },
    children: [new TextRun({
      text,
      font: FONT_BODY, size: 21, color: COLOR_TEXT,
    })],
  });
}

function bulletR(runs) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60, line: 260 },
    children: runs.map(r => {
      if (typeof r === "string") {
        return new TextRun({ text: r, font: FONT_BODY, size: 21, color: COLOR_TEXT });
      }
      return new TextRun({
        text: r.text,
        font: r.font || FONT_BODY,
        size: r.size || 21,
        color: r.color || COLOR_TEXT,
        bold: r.bold || false,
        italics: r.italics || false,
      });
    }),
  });
}

// Uthevingsboks for "Om dette notatet"
function rammeboks(tittel, innhold) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: CONTENT_WIDTH, type: WidthType.DXA },
            shading: { fill: COLOR_BG_HIGHLIGHT, type: ShadingType.CLEAR },
            borders: {
              top: thinBorder,
              bottom: thinBorder,
              left: accentBorderLeft,
              right: thinBorder,
            },
            margins: { top: 160, bottom: 160, left: 280, right: 280 },
            children: [
              new Paragraph({
                spacing: { after: 100 },
                children: [new TextRun({
                  text: tittel,
                  font: FONT_HEADING, size: 22, bold: true, color: COLOR_ACCENT,
                })],
              }),
              ...innhold,
            ],
          }),
        ],
      }),
    ],
  });
}

// ============================================================
// HEADER MED BEGGE LOGOER
// ============================================================
function topplogoer() {
  const halfWidth = Math.floor(CONTENT_WIDTH / 2);
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [halfWidth, halfWidth],
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: halfWidth, type: WidthType.DXA },
            borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
            verticalAlign: VerticalAlign.CENTER,
            margins: { top: 60, bottom: 60, left: 0, right: 100 },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { after: 0 },
                children: [new ImageRun({
                  type: "png",
                  data: fs.readFileSync("C:/codex/beredskap/presentasjon/figurer/VAFS_med_skygge.png"),
                  transformation: { width: 130, height: 92 },
                  altText: {
                    title: "Vestre Aker Frivilligsentral",
                    description: "Logo Vestre Aker Frivilligsentral",
                    name: "VAFS_logo",
                  },
                })],
              }),
            ],
          }),
          new TableCell({
            width: { size: halfWidth, type: WidthType.DXA },
            borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
            verticalAlign: VerticalAlign.CENTER,
            margins: { top: 60, bottom: 60, left: 100, right: 0 },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                spacing: { after: 0 },
                children: [new ImageRun({
                  type: "png",
                  data: fs.readFileSync("C:/codex/beredskap/presentasjon/figurer/OVFS_med_skygge.png"),
                  transformation: { width: 138, height: 92 },
                  altText: {
                    title: "Oslo Vest Frivilligsentral",
                    description: "Logo Oslo Vest Frivilligsentral",
                    name: "OVFS_logo",
                  },
                })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

// ============================================================
// DOKUMENTINNHOLD
// ============================================================

// SIDE 1
function side1() {
  return [
    topplogoer(),
    new Paragraph({ spacing: { after: 80 } }),

    // Tittel
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 120 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: COLOR_ACCENT, space: 6 } },
      children: [new TextRun({
        text: "Beredskapsplan for Frivilligsentralene i Bydel Vestre Aker",
        font: FONT_HEADING, size: 32, bold: true, color: COLOR_PRIMARY,
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      children: [new TextRun({
        text: "Notat til bydel Vestre Aker og Oslo kommune · mai 2026",
        font: FONT_HEADING, size: 22, italics: true, color: COLOR_TEXT,
      })],
    }),

    // Intro
    P("Vestre Aker Frivilligsentral og Oslo Vest Frivilligsentral har utarbeidet en pilot for lokal beredskap og sivil motstandsdyktighet i Bydel Vestre Aker. Piloten beskriver ti konkrete tiltak som frivilligsentralene kan bidra med, som supplement til kommunens etablerte beredskapsstruktur. Dette notatet er en kortversjon for orientering. Et utfyllende hoveddokument er tilgjengelig for fordypning."),

    // Den viktige framing-boksen
    rammeboks("Om denne planen", [
      new Paragraph({
        spacing: { after: 80, line: 270 },
        alignment: AlignmentType.JUSTIFIED,
        children: [new TextRun({
          text: "Planen er bevisst designet uten restriksjoner på ressurser — verken tid, bemanning eller penger. Den representerer en idealtilstand, hva to frivilligsentraler kan bidra med dersom alt lå til rette. Faktisk gjennomføring vil måtte tilpasses de ressursene vi faktisk klarer å mobilisere.",
          font: FONT_BODY, size: 21, color: COLOR_TEXT,
        })],
      }),
      new Paragraph({
        spacing: { after: 0, line: 270 },
        alignment: AlignmentType.JUSTIFIED,
        children: [new TextRun({
          text: "Det betyr at notatet og hoveddokumentet skal leses som et utgangspunkt for dialog, ikke som en endelig forpliktelse. Vi inviterer bydel og kommune til samtale om hva som er realistisk å gjennomføre — i hvilken rekkefølge, med hvilke ressurser, og innenfor hvilke samarbeidsrammer.",
          font: FONT_BODY, size: 21, color: COLOR_TEXT,
        })],
      }),
    ]),

    H2("Bakgrunn — det gapet vi har identifisert"),
    P("Oslo kommunes offisielle kontaktpunkt for Bydel Vestre Aker ved langvarig strømbortfall og ekom-svikt er i dag kun Persbråten videregående skole. Dette er et godt valg som del av kommunens infrastruktur, men det fungerer ikke alene for et så stort område som Vestre Aker — kupert med åser og friområder som deler bydelen i flere lokalmiljøer. For innbyggere med redusert mobilitet kan Persbråten være utilgjengelig nettopp i de scenarioene der en møteplass trengs mest."),
    P("Frivilligsentralene har faste lokaler med lokal forankring, kjennskap til nærmiljøet, eksisterende tillitsnettverk og erfaring med å koordinere frivillig innsats. Disse ressursene utnyttes ikke systematisk i lokal beredskap i dag — det er der piloten kommer inn."),
  ];
}

// SIDE 2
function side2() {
  return [
    H2("Faglig og politisk forankring"),
    P("Piloten bygger på det norske og europeiske rammeverket for sivilsamfunnets rolle i beredskap. Sentrale referansepunkter:"),
    bulletR([
      { text: "Kommunalt risikobilde Oslo 2025 (KRB): ", bold: true, color: COLOR_PRIMARY },
      "Topprisikoer er bortfall av vannforsyning (Maridalsvannet — reserve klar 2028), strømbortfall, cyberangrep mot ekom (høyest sannsynlighet) og pandemi. Alle krever lokal respons før offentlige strukturer mobiliserer fullt.",
    ]),
    bulletR([
      { text: "Forskrift om kommunal beredskapsplikt § 4c: ", bold: true, color: COLOR_PRIMARY },
      "Kommunen skal ha ressursoversikt over frivillige aktørers kapasitet. Frivilligsentralene ønsker å inngå i denne oversikten.",
    ]),
    bulletR([
      { text: "Totalberedskapsmeldingen og NOU 2023:17 «Nå er det alvor»: ", bold: true, color: COLOR_PRIMARY },
      "Begge slår fast at sivilsamfunnsrollen i beredskap må styrkes.",
    ]),
    bulletR([
      { text: "EU Preparedness Union Strategy (mars 2025): ", bold: true, color: COLOR_PRIMARY },
      "Slår fast at frivillige organisasjoner og sivilsamfunn skal være i kjernen av lokal resiliens; anbefaler lokale kontaktpunkter bemannet av trente frivillige, beredskap i skolekurrikulum, og flerspråklig kommunikasjon.",
    ]),

    H2("De ti tiltakene"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [1100, 2900, 6206],
      borders: tableBorders,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: [
            new TableCell({
              width: { size: 1100, type: WidthType.DXA },
              shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR },
              borders: tableBorders,
              margins: { top: 80, bottom: 80, left: 160, right: 160 },
              children: [new Paragraph({ children: [new TextRun({ text: "Kode", font: FONT_HEADING, size: 20, bold: true, color: "FFFFFF" })] })],
            }),
            new TableCell({
              width: { size: 2900, type: WidthType.DXA },
              shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR },
              borders: tableBorders,
              margins: { top: 80, bottom: 80, left: 160, right: 160 },
              children: [new Paragraph({ children: [new TextRun({ text: "Tittel", font: FONT_HEADING, size: 20, bold: true, color: "FFFFFF" })] })],
            }),
            new TableCell({
              width: { size: 6206, type: WidthType.DXA },
              shading: { fill: COLOR_PRIMARY, type: ShadingType.CLEAR },
              borders: tableBorders,
              margins: { top: 80, bottom: 80, left: 160, right: 160 },
              children: [new Paragraph({ children: [new TextRun({ text: "Hva tiltaket innebærer", font: FONT_HEADING, size: 20, bold: true, color: "FFFFFF" })] })],
            }),
          ],
        }),
        ...[
          ["T1", "Lokal møteplass ved krise", "Frivilligsentralene som desentraliserte kontaktpunkter ved krise — supplement til Persbråten VGS. Beredskapslager, sikringsradio, papirsikkerhetskopi av nøkkellister. Tre ambisjonsnivåer fra minimumsmøteplass til fullt utbygd beredskapspunkt."],
          ["T2", "Beredskapsfrivillige", "Fast korps av 6–8 forhåndsrekrutterte og -kursete frivillige per møteplass, primært pensjonister med tidligere yrkesbakgrunn fra forsvar, nødetater eller helsevesen. Inspirert av svenske Frivilliga Resursgrupper (FRG)."],
          ["T3", "Register over sårbare beboere", "Oppdatert register over eldre og andre med ikke-medisinsk hjelpebehov. Hybrid digital primær + papirsikkerhetskopi. Aktivt samtykke. Krever kommuneadvokat-prosess for å avklare juridisk grunnlag for utvidet samarbeid med bydelen."],
          ["T4", "Beredskapskurs for befolkningen", "Kveldskurs i egenberedskap basert på DSBs én-ukes-prinsipp. Innhold: vann, mat, varme, kommunikasjon, førstehjelp, familie og naboer. Kursholdere fra T2 og T7."],
          ["T5", "Førstehjelp for befolkningen", "Regulære, gratis HLR- og AED-kurs i samarbeid med Røde Kors/Norsk Folkehjelp. Lokale pensjonerte fagpersoner som instruktører. Oppfriskningskurs hvert 2. år."],
          ["T6", "Beredskapsvenner og nabolagsberedskap", "To nivåer: parvise nabolagsavtaler (nasjonal politikk fra DSB) og organiserte beredskapsgrupper i borettslag (inspirert av japansk jishu-bosai-soshiki). Frivilligsentralen er logistisk hub for borettslagene."],
          ["T7", "Kompetansekart over frivillige", "Frivillig opt-in-register over lokalt tilgjengelig kompetanse (pensjonister, ikke aktivt utrykningspersonell). Sju kategorier: helse, teknisk, logistikk, kommunikasjon, sosialt, praktisk, beredskapserfaring."],
          ["T8", "Koordinering av spontanfrivillige  —  høy prioritet", "Mottakssystem for frivillige som dukker opp uten forhåndsrolle. Minst to opplærte mottakssjefer per møteplass; rollen er godt egnet for frivillige med fysisk funksjonsnedsettelse. Digital prosedyre med papirsikkerhetskopi."],
          ["T9", "Ungdom og skole", "Beredskapsopplegg integrert i ungdomsskolens valgfag «Innsats for andre». Etablerte relasjoner til to ungdomsskoler i bydelen. Tre formater: juniorberedskapsfrivillige, beredskapsdag, læringsopplegg til andre lærere."],
          ["T10", "Flerspråklig beredskap", "Beredskapsmateriale på de mest aktuelle fremmedspråkene i bydelen, distribuert via etablerte kanaler — Bydelsmødre Vestre Aker som strategisk inngang. Bruker eksisterende oversatt DSB-materiale, ikke nytt."],
        ].map(([k, t, b]) => new TableRow({
          cantSplit: true,
          children: [
            new TableCell({
              width: { size: 1100, type: WidthType.DXA },
              borders: tableBorders,
              margins: { top: 80, bottom: 80, left: 160, right: 160 },
              children: [new Paragraph({ children: [new TextRun({ text: k, font: FONT_HEADING, size: 21, bold: true, color: k === "T8" ? COLOR_ACCENT : COLOR_PRIMARY })] })],
            }),
            new TableCell({
              width: { size: 2900, type: WidthType.DXA },
              borders: tableBorders,
              margins: { top: 80, bottom: 80, left: 160, right: 160 },
              children: [new Paragraph({ children: [new TextRun({ text: t, font: FONT_BODY, size: 20, bold: true, color: COLOR_PRIMARY })] })],
            }),
            new TableCell({
              width: { size: 6206, type: WidthType.DXA },
              borders: tableBorders,
              margins: { top: 80, bottom: 80, left: 160, right: 160 },
              children: [new Paragraph({ children: [new TextRun({ text: b, font: FONT_BODY, size: 19, color: COLOR_TEXT })] })],
            }),
          ],
        })),
      ],
    }),
  ];
}

// SIDE 3
function side3() {
  return [
    new Paragraph({ children: [new PageBreak()] }),
    H2("Prioritering — viktighet og kompleksitet"),
    P("Tiltakene er vurdert på to akser: viktighet (beredskapsmessig effekt) og kompleksitet (etablerings- og driftskostnad). Resultatet gir et bilde av hvor vi bør starte, og hvilke tiltak som krever mer tid og forarbeid."),

    new Paragraph({
      spacing: { before: 200, after: 100 },
      alignment: AlignmentType.CENTER,
      keepNext: true,
      keepLines: true,
      children: [new ImageRun({
        type: "png",
        data: fs.readFileSync("C:/codex/beredskap/presentasjon/figurer/01_prioriteringsmatrise.png"),
        transformation: { width: 560, height: 411 },
        altText: {
          title: "Prioriteringsmatrise",
          description: "Plassering av tiltakene etter viktighet og kompleksitet",
          name: "prioriteringsmatrise",
        },
      })],
    }),
    new Paragraph({
      spacing: { after: 200 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({
        text: "Figur: Plassering av de ti tiltakene etter viktighet og kompleksitet.",
        font: FONT_BODY, size: 18, italics: true, color: COLOR_MUTED,
      })],
    }),

    H2("Hvordan vi leser matrisen"),
    bulletR([
      { text: "Strategisk grunnlag (T1, T2, T3): ", bold: true, color: COLOR_PRIMARY },
      "Høyest viktighet og høyere kompleksitet. Må etableres først fordi de andre tiltakene hviler på dem — møteplassen, beredskapsfrivillige som bemanner den, og registeret over sårbare beboere.",
    ]),
    bulletR([
      { text: "Strategiske raske gevinster (T4, T5, T6, T7): ", bold: true, color: COLOR_PRIMARY },
      "Høy viktighet og lav kompleksitet. Gir bredde-effekt på relativt kort tid og med begrenset innsats.",
    ]),
    bulletR([
      { text: "Høy prioritet (T8): ", bold: true, color: COLOR_ACCENT },
      "Koordinering av spontanfrivillige har høyeste viktighet og lav kompleksitet — men blir ofte glemt. Uten T8 risikerer møteplassen å bli kaotisk og kontraproduktiv ved en krise.",
    ]),
    bulletR([
      { text: "Langsiktige satsinger (T9, T10): ", bold: true, color: COLOR_PRIMARY },
      "Lavere kortsiktig effekt, men strategisk viktige — T9 multipliserer rekkevidden via hjembringing til foreldre, og T10 er en forutsetning for at modellen er overførbar til mer flerspråklige bydeler.",
    ]),
  ];
}

// SIDE 4
function side4() {
  return [
    new Paragraph({ children: [new PageBreak()] }),
    H2("Hvordan piloten er tenkt drevet"),
    PR([
      { text: "Supplement, ikke konkurrent. ", bold: true, color: COLOR_PRIMARY },
      "Frivilligsentralene fyller gap i eksisterende beredskapsstruktur. Vi overtar aldri roller som tilhører blålysetater, Sivilforsvaret, bydelen eller andre etablerte aktører. Vi avlaster og utvider — vi erstatter ikke.",
    ]),
    PR([
      { text: "Samarbeid er forutsetningen. ", bold: true, color: COLOR_PRIMARY },
      "Alt skjer i tett dialog med bydel, kommune og relevante organisasjoner. Beredskapsfrivillige kan ikke lære kommunens beredskapsplaner uten at kommunen deltar aktivt i opplæringen. Kontakt med DSB, Sivilforsvaret, Norsk Folkehjelp og Røde Kors etableres via Oslo kommunes beredskapsavdeling.",
    ]),
    PR([
      { text: "Pilotlogikk. ", bold: true, color: COLOR_PRIMARY },
      "Hvert tiltak er designet slik at en annen bydel eller kommune kan ta det i bruk uten å starte fra null. Dokumentasjon, maler og prosesser deles åpent. Modellen skal være overførbar.",
    ]),
    PR([
      { text: "Ambisiøs plan, realistisk gjennomføring. ", bold: true, color: COLOR_PRIMARY },
      "Planen er, som dette notatet innleder med, designet uten ressursrestriksjoner. Gjennomføringen vil tilpasses faktisk kapasitet. Tiltakene har bevisst innebygde ambisjonsnivåer (særlig T1 — fra minimumsmøteplass til fullt utbygd beredskapspunkt) slik at vi kan starte enkelt og bygge ut etter hvert.",
    ]),

    H2("Veien videre"),
    P("Piloten er bygget i to faser. Fase 1 — utarbeidelse av denne skissen og hoveddokumentet, og formell dialog med beredskapsansvarlig i bydelen og Oslo kommune. Skissen er ferdig. Den formelle dialogen er ikke startet, men er det neste umiddelbare skrittet."),
    P("Fase 2 — utarbeidelse av en konkret handlingsplan med ansvar, ressursbehov og fremdrift — starter når forankringen er på plass. Det er i fase 2 vi sammen med bydelen og kommunen avklarer hva som faktisk skal prioriteres, hvilke tiltak som skal piloteres først, og hvordan dette finansieres og bemannes."),
    P("Vi ser frem til samtalen om hvordan frivilligsentralene best kan bidra — og hvilken form og takt det får som passer inn i bydelens og kommunens øvrige beredskapsarbeid."),

    // Sluttinformasjon
    new Paragraph({
      spacing: { before: 320, after: 80 },
      keepNext: true,
      keepLines: true,
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER, space: 6 } },
      children: [new TextRun({
        text: "Om dokumentet",
        font: FONT_HEADING, size: 22, bold: true, color: COLOR_PRIMARY,
      })],
    }),
    P("Dette notatet er en kortversjon. Et utfyllende hoveddokument med detaljerte tiltaksbeskrivelser, alle fire visualiseringer (prioriteringsmatrise, avhengighetsgraf, tidslinje, risiko-tiltak-matrise), faktabokser med kildehenvisninger, ordliste og kildemateriale er tilgjengelig på forespørsel."),
    new Paragraph({
      spacing: { before: 80, after: 0 },
      children: [new TextRun({
        text: "Utarbeidet av Vestre Aker Frivilligsentral og Oslo Vest Frivilligsentral  ·  mai 2026",
        font: FONT_BODY, size: 19, italics: true, color: COLOR_MUTED,
      })],
    }),
  ];
}

// ============================================================
// DOKUMENT
// ============================================================
const doc = new Document({
  creator: "Vestre Aker Frivilligsentral",
  title: "Notat — Beredskapsplan for Frivilligsentralene Vestre Aker",
  description: "Kortversjon for bydel og kommune",
  styles: {
    default: {
      document: {
        run: { font: FONT_BODY, size: 21, color: COLOR_TEXT },
        paragraph: { spacing: { line: 270 } },
      },
    },
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 540, hanging: 300 } } },
        }],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "Notat  ·  ", font: FONT_BODY, size: 18, color: COLOR_MUTED, italics: true }),
                new TextRun({ text: "Side ", font: FONT_BODY, size: 20, color: COLOR_PRIMARY, bold: true }),
                new TextRun({ children: [PageNumber.CURRENT], font: FONT_BODY, size: 20, color: COLOR_PRIMARY, bold: true }),
                new TextRun({ text: " av ", font: FONT_BODY, size: 20, color: COLOR_PRIMARY, bold: true }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT_BODY, size: 20, color: COLOR_PRIMARY, bold: true }),
              ],
            }),
          ],
        }),
      },
      children: [
        ...side1(),
        ...side2(),
        ...side3(),
        ...side4(),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buffer => {
  const out = "C:/codex/beredskap/presentasjon/Notat_Beredskap_Vestre_Aker.docx";
  fs.writeFileSync(out, buffer);
  console.log("OK — notat generert: " + out);
}).catch(err => {
  console.error("FEIL:", err);
  process.exit(1);
});
