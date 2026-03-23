const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaBullseye, FaChartLine, FaLightbulb, FaUsers, FaCar, FaDollarSign, FaBookOpen, FaCogs, FaCheckCircle, FaGlobeAmericas } = require("react-icons/fa");

function renderIconSvg(Ic, color, size = 256) {
    return ReactDOMServer.renderToStaticMarkup(React.createElement(Ic, { color, size: String(size) }));
}
async function iconPng(Ic, color) {
    const buf = await sharp(Buffer.from(renderIconSvg(Ic, color))).png().toBuffer();
    return "image/png;base64," + buf.toString("base64");
}

const C = { dark: "0A0A0F", card: "141420", red: "E82127", redDim: "3D0A0C", white: "F0F0F0", gray: "999999", dim: "666677", border: "2A2A3A" };
const mkS = () => ({ type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.35 });

function hdr(sl, num, title, sub) {
    sl.addText(num, { x: 0.5, y: 0.15, w: 0.8, h: 0.55, fontSize: 28, fontFace: "Arial Black", color: C.red, margin: 0 });
    sl.addText(title, { x: 0.5, y: 0.55, w: 9, h: 0.45, fontSize: 22, fontFace: "Arial Black", color: C.white, margin: 0 });
    if (sub) sl.addText(sub, { x: 0.5, y: 0.95, w: 8, h: 0.3, fontSize: 10, fontFace: "Georgia", color: C.dim, italic: true, margin: 0 });
}

// Empty placeholder text for cards
const PH = "(content to be added)";
const phStyle = { fontSize: 9, fontFace: "Arial", color: C.dim, margin: 0 };

async function build() {
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    pres.title = "Tesla Model 3 - Marketing Strategy";
    const R = pres.shapes.RECTANGLE;
    const topBar = (s) => s.addShape(R, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    const botBar = (s) => s.addShape(R, { x: 0, y: 5.595, w: 10, h: 0.03, fill: { color: C.red } });

    const ic = { car: await iconPng(FaCar, "#E82127"), globe: await iconPng(FaGlobeAmericas, "#E82127") };

    // ═══════════════════════════════════════
    // SLIDE 1: COVER + TOC
    // ═══════════════════════════════════════
    let s1 = pres.addSlide();
    s1.background = { color: C.dark };
    topBar(s1); botBar(s1);
    s1.addShape(pres.shapes.OVAL, { x: 2, y: 0.5, w: 6, h: 4.5, fill: { color: C.redDim, transparency: 75 } });
    s1.addText("MARKETING STRATEGY ANALYSIS", { x: 0.5, y: 0.7, w: 9, h: 0.35, fontSize: 10, fontFace: "Arial", color: C.red, charSpacing: 5, align: "center", bold: true });
    s1.addText("TESLA", { x: 0.5, y: 1.0, w: 9, h: 1.1, fontSize: 64, fontFace: "Arial Black", color: C.white, align: "center", margin: 0 });
    s1.addText("MODEL 3", { x: 0.5, y: 2.0, w: 9, h: 0.75, fontSize: 38, fontFace: "Arial Black", color: C.red, align: "center", margin: 0 });
    s1.addShape(R, { x: 4.2, y: 2.8, w: 1.6, h: 0.025, fill: { color: C.red } });
    s1.addImage({ data: ic.car, x: 4.7, y: 2.9, w: 0.5, h: 0.5 });

    s1.addText("TABLE OF CONTENTS", { x: 1.5, y: 3.5, w: 7, h: 0.3, fontSize: 9, fontFace: "Arial", color: C.red, charSpacing: 4, align: "center", bold: true });
    const toc = [
        "Part 1: Introduction to Tesla and Model 3",
        "Part 2: Critical Analysis of Market Context (PESTEL, Porter's, SWOT)",
        "Part 3: Segmentation, Targeting and Positioning (STP)",
        "Part 4: Analysis of Marketing Mix (7Ps) & Recommendations",
        "Executive Summary & References",
    ];
    toc.forEach((t, i) => {
        s1.addText([
            { text: String(i + 1).padStart(2, '0') + "  ", options: { fontSize: 9, color: C.red, bold: true } },
            { text: t, options: { fontSize: 9, color: C.gray } },
        ], { x: 2, y: 3.85 + i * 0.3, w: 6, h: 0.26, margin: 0 });
    });

    // ═══════════════════════════════════════
    // SLIDE 2: PART 1 — INTRO
    // ═══════════════════════════════════════
    let s2 = pres.addSlide();
    s2.background = { color: C.dark };
    topBar(s2);
    hdr(s2, "01", "Introduction to Tesla & Model 3");

    // 1.1 Background
    s2.addShape(R, { x: 0.5, y: 1.35, w: 4.3, h: 2.0, fill: { color: C.card }, shadow: mkS() });
    s2.addShape(R, { x: 0.5, y: 1.35, w: 4.3, h: 0.035, fill: { color: C.red } });
    s2.addText("1.1 Background of Tesla", { x: 0.65, y: 1.45, w: 3.5, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s2.addText(PH, { x: 0.65, y: 1.78, w: 3.95, h: 1.45, ...phStyle });

    // 1.2 Story
    s2.addShape(R, { x: 5.2, y: 1.35, w: 4.3, h: 2.0, fill: { color: C.card }, shadow: mkS() });
    s2.addShape(R, { x: 5.2, y: 1.35, w: 4.3, h: 0.035, fill: { color: C.red } });
    s2.addText("1.2 Story of Model 3", { x: 5.35, y: 1.45, w: 3.5, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s2.addText(PH, { x: 5.35, y: 1.78, w: 3.95, h: 1.45, ...phStyle });

    // 1.3 Why Consumers Choose
    s2.addShape(R, { x: 0.5, y: 3.55, w: 9.0, h: 1.85, fill: { color: C.card }, shadow: mkS() });
    s2.addShape(R, { x: 0.5, y: 3.55, w: 0.05, h: 1.85, fill: { color: C.red } });
    s2.addText("1.3 Why Consumers Choose Model 3", { x: 0.72, y: 3.62, w: 4, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s2.addText(PH, { x: 0.72, y: 3.9, w: 8.5, h: 1.35, ...phStyle });

    // ═══════════════════════════════════════
    // SLIDE 3: PART 2 — MARKET ANALYSIS
    // ═══════════════════════════════════════
    let s3 = pres.addSlide();
    s3.background = { color: C.dark };
    topBar(s3);
    hdr(s3, "02", "Critical Analysis of Market Context");

    // Market Trends + PESTEL — left
    s3.addShape(R, { x: 0.5, y: 1.35, w: 4.3, h: 4.05, fill: { color: C.card }, shadow: mkS() });
    s3.addShape(R, { x: 0.5, y: 1.35, w: 4.3, h: 0.035, fill: { color: C.red } });
    s3.addText("2.1 Market Trends", { x: 0.65, y: 1.45, w: 3, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s3.addText(PH, { x: 0.65, y: 1.75, w: 3.95, h: 0.9, ...phStyle });
    s3.addText("2.2.1 PESTEL Analysis", { x: 0.65, y: 2.7, w: 3, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

    // PESTEL table structure (empty analysis column)
    const pestelH = [
        { text: "Factor", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 8, fontFace: "Arial" } },
        { text: "Impact", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 8, fontFace: "Arial", align: "center" } },
        { text: "Key Points", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 8, fontFace: "Arial" } },
    ];
    const pestelF = ["Political", "Economic", "Social", "Technological", "Environmental", "Legal"];
    const pestelD = [pestelH, ...pestelF.map(f => [
        { text: f, options: { fill: { color: C.card }, color: C.white, fontSize: 7.5, fontFace: "Arial", bold: true } },
        { text: "—", options: { fill: { color: C.card }, color: "FFB74D", fontSize: 7.5, fontFace: "Arial", align: "center" } },
        { text: PH, options: { fill: { color: C.card }, color: C.dim, fontSize: 7, fontFace: "Arial" } },
    ])];
    s3.addTable(pestelD, { x: 0.65, y: 2.95, w: 3.95, h: 2.3, border: { pt: 0.3, color: C.border }, colW: [0.7, 0.55, 2.7], rowH: [0.22, 0.33, 0.33, 0.33, 0.33, 0.33, 0.33] });

    // Porter's — right top
    s3.addShape(R, { x: 5.2, y: 1.35, w: 4.3, h: 2.15, fill: { color: C.card }, shadow: mkS() });
    s3.addShape(R, { x: 5.2, y: 1.35, w: 4.3, h: 0.035, fill: { color: C.red } });
    s3.addText("2.2.2 Porter's Five Forces", { x: 5.35, y: 1.45, w: 3.5, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

    const porterH = [
        { text: "Force", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 8, fontFace: "Arial" } },
        { text: "Level", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 8, fontFace: "Arial", align: "center" } },
        { text: "Analysis", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 8, fontFace: "Arial" } },
    ];
    const porterF = ["Suppliers", "Buyers", "New Entrants", "Substitutes", "Rivalry"];
    const porterD = [porterH, ...porterF.map(f => [
        { text: f, options: { fill: { color: C.card }, color: C.white, fontSize: 7.5, fontFace: "Arial", bold: true } },
        { text: "—", options: { fill: { color: C.card }, color: "FFB74D", fontSize: 8, fontFace: "Arial", bold: true, align: "center" } },
        { text: PH, options: { fill: { color: C.card }, color: C.dim, fontSize: 7.5, fontFace: "Arial" } },
    ])];
    s3.addTable(porterD, { x: 5.35, y: 1.75, w: 3.95, h: 1.6, border: { pt: 0.3, color: C.border }, colW: [0.85, 0.55, 2.55], rowH: [0.22, 0.27, 0.27, 0.27, 0.27, 0.27] });

    // SWOT — right bottom
    s3.addShape(R, { x: 5.2, y: 3.65, w: 4.3, h: 1.75, fill: { color: C.card }, shadow: mkS() });
    s3.addText("SWOT Analysis", { x: 5.35, y: 3.7, w: 3, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    const swotLabels = [
        { l: "S — Strengths", c: "1B5E20" },
        { l: "W — Weaknesses", c: "B71C1C" },
        { l: "O — Opportunities", c: "0D47A1" },
        { l: "T — Threats", c: "E65100" },
    ];
    swotLabels.forEach((s, i) => {
        const col = i % 2, row = Math.floor(i / 2);
        const sx = 5.35 + col * 2.05, sy = 3.98 + row * 0.7;
        s3.addShape(R, { x: sx, y: sy, w: 1.95, h: 0.62, fill: { color: C.card } });
        s3.addShape(R, { x: sx, y: sy, w: 0.04, h: 0.62, fill: { color: s.c } });
        s3.addText(s.l, { x: sx + 0.1, y: sy + 0.04, w: 1.75, h: 0.18, fontSize: 7.5, fontFace: "Arial", color: s.c, bold: true, margin: 0 });
        s3.addText(PH, { x: sx + 0.1, y: sy + 0.24, w: 1.75, h: 0.34, fontSize: 6.5, fontFace: "Arial", color: C.dim, margin: 0 });
    });

    // ═══════════════════════════════════════
    // SLIDE 4: PART 3 — STP
    // ═══════════════════════════════════════
    let s4 = pres.addSlide();
    s4.background = { color: C.dark };
    topBar(s4);
    hdr(s4, "03", "Segmentation, Targeting & Positioning");

    // Segmentation 3 cards
    const segTitles = ["3.1.1 Demographic", "3.1.2 Psychographic", "3.1.3 Behavioural"];
    segTitles.forEach((t, i) => {
        const x = 0.5 + i * 3.1;
        s4.addShape(R, { x, y: 1.3, w: 2.85, h: 1.65, fill: { color: C.card }, shadow: mkS() });
        s4.addShape(R, { x, y: 1.3, w: 2.85, h: 0.035, fill: { color: C.red } });
        s4.addText(t, { x: x + 0.12, y: 1.38, w: 2.6, h: 0.22, fontSize: 9, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
        s4.addText(PH, { x: x + 0.12, y: 1.65, w: 2.6, h: 1.2, ...phStyle });
    });

    // Targeting 2 cards
    s4.addShape(R, { x: 0.5, y: 3.15, w: 4.3, h: 1.15, fill: { color: C.card }, shadow: mkS() });
    s4.addShape(R, { x: 0.5, y: 3.15, w: 0.05, h: 1.15, fill: { color: "1B5E20" } });
    s4.addText("3.2.1 Primary Target", { x: 0.7, y: 3.2, w: 3.5, h: 0.22, fontSize: 9, fontFace: "Arial", color: "66BB6A", bold: true, margin: 0 });
    s4.addText(PH, { x: 0.7, y: 3.45, w: 3.9, h: 0.8, ...phStyle });

    s4.addShape(R, { x: 5.2, y: 3.15, w: 4.3, h: 1.15, fill: { color: C.card }, shadow: mkS() });
    s4.addShape(R, { x: 5.2, y: 3.15, w: 0.05, h: 1.15, fill: { color: "0D47A1" } });
    s4.addText("3.2.2 Secondary Target", { x: 5.4, y: 3.2, w: 3.5, h: 0.22, fontSize: 9, fontFace: "Arial", color: "42A5F5", bold: true, margin: 0 });
    s4.addText(PH, { x: 5.4, y: 3.45, w: 3.9, h: 0.8, ...phStyle });

    // Positioning + Recs
    s4.addShape(R, { x: 0.5, y: 4.5, w: 4.3, h: 0.95, fill: { color: C.card }, shadow: mkS() });
    s4.addShape(R, { x: 0.5, y: 4.5, w: 4.3, h: 0.035, fill: { color: C.red } });
    s4.addText("3.3 Positioning Strategy", { x: 0.65, y: 4.55, w: 3.5, h: 0.2, fontSize: 9, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s4.addText(PH, { x: 0.65, y: 4.78, w: 3.95, h: 0.6, ...phStyle });

    s4.addShape(R, { x: 5.2, y: 4.5, w: 4.3, h: 0.95, fill: { color: C.card }, shadow: mkS() });
    s4.addShape(R, { x: 5.2, y: 4.5, w: 0.05, h: 0.95, fill: { color: "E65100" } });
    s4.addText("3.4 STP Recommendations", { x: 5.4, y: 4.55, w: 3.5, h: 0.2, fontSize: 9, fontFace: "Arial", color: "FF9800", bold: true, margin: 0 });
    s4.addText(PH, { x: 5.4, y: 4.78, w: 3.95, h: 0.6, ...phStyle });

    // ═══════════════════════════════════════
    // SLIDE 5: PART 4 — 7Ps
    // ═══════════════════════════════════════
    let s5 = pres.addSlide();
    s5.background = { color: C.dark };
    topBar(s5);
    hdr(s5, "04", "Marketing Mix — 7Ps & Recommendations");

    // Top 4 cards
    const psT = ["4.1 Product", "4.2 Pricing", "4.3 Distribution", "4.4 Promotion"];
    psT.forEach((t, i) => {
        const x = 0.5 + i * 2.32;
        s5.addShape(R, { x, y: 1.3, w: 2.12, h: 1.85, fill: { color: C.card }, shadow: mkS() });
        s5.addShape(R, { x, y: 1.3, w: 2.12, h: 0.035, fill: { color: C.red } });
        s5.addText(t, { x: x + 0.1, y: 1.37, w: 1.9, h: 0.2, fontSize: 8.5, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
        s5.addText(PH, { x: x + 0.1, y: 1.6, w: 1.9, h: 1.45, ...phStyle });
    });

    // Bottom 3 cards
    const psB = ["4.5 People", "4.6 Process", "4.7 Physical Evidence"];
    psB.forEach((t, i) => {
        const x = 0.5 + i * 3.1;
        s5.addShape(R, { x, y: 3.35, w: 2.85, h: 1.35, fill: { color: C.card }, shadow: mkS() });
        s5.addShape(R, { x, y: 3.35, w: 2.85, h: 0.035, fill: { color: C.red } });
        s5.addText(t, { x: x + 0.1, y: 3.42, w: 2.6, h: 0.2, fontSize: 8.5, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
        s5.addText(PH, { x: x + 0.1, y: 3.65, w: 2.6, h: 1.0, ...phStyle });
    });

    // Recommendations bar
    s5.addShape(R, { x: 0.5, y: 4.85, w: 9.0, h: 0.6, fill: { color: C.card }, shadow: mkS() });
    s5.addShape(R, { x: 0.5, y: 4.85, w: 0.05, h: 0.6, fill: { color: "E65100" } });
    s5.addText("4.8 Recommendations", { x: 0.7, y: 4.88, w: 2, h: 0.2, fontSize: 8, fontFace: "Arial", color: "FF9800", bold: true, margin: 0 });
    s5.addText(PH, { x: 0.7, y: 5.08, w: 8.6, h: 0.3, ...phStyle });

    // ═══════════════════════════════════════
    // SLIDE 6: EXEC SUMMARY + REFS
    // ═══════════════════════════════════════
    let s6 = pres.addSlide();
    s6.background = { color: C.dark };
    topBar(s6); botBar(s6);

    s6.addText("Executive Summary", { x: 0.5, y: 0.4, w: 9, h: 0.5, fontSize: 24, fontFace: "Arial Black", color: C.white, align: "center", margin: 0 });

    // 7 numbered placeholder lines
    for (let i = 0; i < 7; i++) {
        s6.addText([
            { text: "0" + (i + 1) + "  ", options: { fontSize: 10, color: C.red, bold: true, fontFace: "Arial Black" } },
            { text: PH, options: { fontSize: 9, color: C.dim, fontFace: "Arial" } },
        ], { x: 1.0, y: 1.05 + i * 0.45, w: 8, h: 0.4, margin: 0 });
    }

    // References
    s6.addShape(R, { x: 0.5, y: 4.3, w: 9.0, h: 0.03, fill: { color: C.border } });
    s6.addText("REFERENCES", { x: 0.5, y: 4.38, w: 3, h: 0.2, fontSize: 8, fontFace: "Arial", color: C.red, charSpacing: 3, bold: true, margin: 0 });
    s6.addText(PH, { x: 0.5, y: 4.6, w: 9.0, h: 0.55, fontSize: 7, fontFace: "Arial", color: C.dim, margin: 0 });

    s6.addText("Thank You", { x: 0, y: 5.2, w: 10, h: 0.3, fontSize: 16, fontFace: "Arial Black", color: C.red, align: "center" });

    await pres.writeFile({ fileName: "C:/laragon/www/tesla/Tesla_Model3_Marketing.pptx" });
    console.log("Done! Saved: Tesla_Model3_Marketing.pptx (empty cards, ready to fill)");
}

build().catch(err => { console.error(err); process.exit(1); });
