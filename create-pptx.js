const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaCar, FaGlobeAmericas } = require("react-icons/fa");

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

function bullets(sl, items, x, y, w, h, fontSize) {
    const textArr = items.map((t, i) => ({
        text: t,
        options: { bullet: true, breakLine: i < items.length - 1, fontSize: fontSize || 8.5, color: C.gray, fontFace: "Arial" }
    }));
    sl.addText(textArr, { x, y, w, h, margin: 0 });
}

async function build() {
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    pres.title = "Tesla Model 3 - Marketing Strategy";
    const R = pres.shapes.RECTANGLE;
    const topBar = (s) => s.addShape(R, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    const botBar = (s) => s.addShape(R, { x: 0, y: 5.595, w: 10, h: 0.03, fill: { color: C.red } });

    const ic = { car: await iconPng(FaCar, "#E82127") };

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
        "Part 4: Analysis of Marketing Mix (4Ps) & Recommendations",
        "References",
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
    bullets(s2, [
        "2003 \u2013 Started with investment of $6.3 million by Elon Musk",
        "2008 \u2013 Tesla Roadster Launched (0-60mph in 3.7s, 395km range, lithium-ion battery)",
        "2012 \u2013 Model S, World's first premium electric sedan",
        "2015 \u2013 Model X released with Autopilot for Model S",
        "2017 \u2013 Model 3 released",
    ], 0.65, 1.75, 3.95, 1.5, 8);

    // 1.2 Story
    s2.addShape(R, { x: 5.2, y: 1.35, w: 4.3, h: 2.0, fill: { color: C.card }, shadow: mkS() });
    s2.addShape(R, { x: 5.2, y: 1.35, w: 4.3, h: 0.035, fill: { color: C.red } });
    s2.addText("1.2 Story of Model 3", { x: 5.35, y: 1.45, w: 3.5, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    bullets(s2, [
        "Hinted in 2006 as an affordable car",
        "In 2008, announced to be a family car",
        "Unveiled in 2016, 325,000 reservation within a week",
        "By August 2017, 455,000 reservations made",
        "2020 Model 3 best selling EV worldwide",
    ], 5.35, 1.75, 3.95, 1.5, 8);

    // 1.3 Why Consumers Choose
    s2.addShape(R, { x: 0.5, y: 3.55, w: 9.0, h: 1.85, fill: { color: C.card }, shadow: mkS() });
    s2.addShape(R, { x: 0.5, y: 3.55, w: 0.05, h: 1.85, fill: { color: C.red } });
    s2.addText("1.3 Why Consumers Choose Model 3", { x: 0.72, y: 3.62, w: 4, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    bullets(s2, [
        "Design and Technology \u2013 Elon believed in building a good product rather than spending on marketing. Let product do its own marketing.",
        "5 Star NCAP safety rating",
        "Priced above middle-income range but subsidy made it affordable for middle class.",
        "Elon Musk Branding [Brand Recognition, Brand Loyalty, Similar to Apple]",
        "Autopilot and Full Self Driving",
    ], 0.72, 3.9, 8.5, 1.4, 8);

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
    bullets(s3, [
        "People being more environmental conscious",
        "Shift to green energy cars",
        "Cheaper maintenance and running costs",
        "Government subsidies",
    ], 0.65, 1.72, 3.95, 0.85, 8);
    s3.addText("2.2.1 PESTEL Analysis", { x: 0.65, y: 2.65, w: 3, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

    // PESTEL table with impacts filled
    const pestelH = [
        { text: "Factor", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 8, fontFace: "Arial" } },
        { text: "Impact", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 8, fontFace: "Arial", align: "center" } },
        { text: "Key Points", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 8, fontFace: "Arial" } },
    ];
    const pestelRows = [
        ["Political", "High", "Support from government since made in US and pure EV."],
        ["Economic", "Moderate", "Made in USA and priced affordably"],
        ["Social", "High", "Change in market from ICE to EV due to many benefits."],
        ["Technological", "High", "Minimalist design, Innovative features, Regular software updates, digital display for all controls."],
        ["Environmental", "High", "Pure EV, hence uses renewable source of energy. No pollution from burning of fuels."],
        ["Legal", "Moderate", "Issues due to FSD and cybersecurity"],
    ];
    const pestelD = [pestelH, ...pestelRows.map(([f, impact, pts]) => [
        { text: f, options: { fill: { color: C.card }, color: C.white, fontSize: 7.5, fontFace: "Arial", bold: true } },
        { text: impact, options: { fill: { color: C.card }, color: impact === "High" ? "FF5252" : impact === "Moderate" ? "FFB74D" : "66BB6A", fontSize: 7.5, fontFace: "Arial", bold: true, align: "center" } },
        { text: pts, options: { fill: { color: C.card }, color: C.dim, fontSize: 7, fontFace: "Arial" } },
    ])];
    s3.addTable(pestelD, { x: 0.65, y: 2.9, w: 3.95, h: 2.4, border: { pt: 0.3, color: C.border }, colW: [0.7, 0.55, 2.7], rowH: [0.22, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35] });

    // Porter's — right top
    s3.addShape(R, { x: 5.2, y: 1.35, w: 4.3, h: 2.15, fill: { color: C.card }, shadow: mkS() });
    s3.addShape(R, { x: 5.2, y: 1.35, w: 4.3, h: 0.035, fill: { color: C.red } });
    s3.addText("2.2.2 Porter's Five Forces", { x: 5.35, y: 1.45, w: 3.5, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

    const porterH = [
        { text: "Force", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 8, fontFace: "Arial" } },
        { text: "Level", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 8, fontFace: "Arial", align: "center" } },
        { text: "Analysis", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 8, fontFace: "Arial" } },
    ];
    const porterRows = [
        ["Suppliers", "Moderate", "Most of the materials self manufactured"],
        ["Buyers", "Moderate", "Have many options."],
        ["New Entrants", "Moderate", "Difficult since requires high investment and innovative technology."],
        ["Substitutes", "Low", "Better tech and cost alternative to ICE and hybrid"],
        ["Rivalry", "High", "Competitive market, fast tech advancements"],
    ];
    const porterD = [porterH, ...porterRows.map(([f, level, analysis]) => [
        { text: f, options: { fill: { color: C.card }, color: C.white, fontSize: 7.5, fontFace: "Arial", bold: true } },
        { text: level, options: { fill: { color: C.card }, color: level === "High" ? "FF5252" : level === "Moderate" ? "FFB74D" : "66BB6A", fontSize: 8, fontFace: "Arial", bold: true, align: "center" } },
        { text: analysis, options: { fill: { color: C.card }, color: C.dim, fontSize: 7.5, fontFace: "Arial" } },
    ])];
    s3.addTable(porterD, { x: 5.35, y: 1.75, w: 3.95, h: 1.6, border: { pt: 0.3, color: C.border }, colW: [0.85, 0.55, 2.55], rowH: [0.22, 0.27, 0.27, 0.27, 0.27, 0.27] });

    // SWOT — right bottom
    s3.addShape(R, { x: 5.2, y: 3.65, w: 4.3, h: 1.75, fill: { color: C.card }, shadow: mkS() });
    s3.addText("SWOT Analysis", { x: 5.35, y: 3.7, w: 3, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    const swotData = [
        { l: "S \u2014 Strengths", c: "1B5E20", t: "Good Product\nBrand Loyalty" },
        { l: "W \u2014 Weaknesses", c: "B71C1C", t: "Production Delays\nQuality Control Issues" },
        { l: "O \u2014 Opportunities", c: "0D47A1", t: "Robotaxi, FSD\nEmerging Market" },
        { l: "T \u2014 Threats", c: "E65100", t: "Competitive Market\nCybersecurity and FSD liability risk" },
    ];
    swotData.forEach((s, i) => {
        const col = i % 2, row = Math.floor(i / 2);
        const sx = 5.35 + col * 2.05, sy = 3.98 + row * 0.7;
        s3.addShape(R, { x: sx, y: sy, w: 1.95, h: 0.62, fill: { color: C.card } });
        s3.addShape(R, { x: sx, y: sy, w: 0.04, h: 0.62, fill: { color: s.c } });
        s3.addText(s.l, { x: sx + 0.1, y: sy + 0.04, w: 1.75, h: 0.18, fontSize: 7.5, fontFace: "Arial", color: s.c, bold: true, margin: 0 });
        s3.addText(s.t, { x: sx + 0.1, y: sy + 0.24, w: 1.75, h: 0.34, fontSize: 6.5, fontFace: "Arial", color: C.dim, margin: 0 });
    });

    // ═══════════════════════════════════════
    // SLIDE 4: PART 3 — STP
    // ═══════════════════════════════════════
    let s4 = pres.addSlide();
    s4.background = { color: C.dark };
    topBar(s4);
    hdr(s4, "03", "Segmentation, Targeting & Positioning");

    // Segmentation 3 cards
    const segData = [
        { title: "3.1.1 Demographic", items: ["Mostly Male buyers than female", "30-40 are major buyers"] },
        { title: "3.1.2 Psychographic", items: ["High brand recognition", "Brand loyalty"] },
        { title: "3.1.3 Behavioural", items: ["Innovative ideas", "Touch Screen display [Netflix, Games]", "Connects to PS5", "Many new fun features"] },
    ];
    segData.forEach((seg, i) => {
        const x = 0.5 + i * 3.1;
        s4.addShape(R, { x, y: 1.3, w: 2.85, h: 1.65, fill: { color: C.card }, shadow: mkS() });
        s4.addShape(R, { x, y: 1.3, w: 2.85, h: 0.035, fill: { color: C.red } });
        s4.addText(seg.title, { x: x + 0.12, y: 1.38, w: 2.6, h: 0.22, fontSize: 9, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
        bullets(s4, seg.items, x + 0.12, 1.65, 2.6, 1.2, 8);
    });

    // Targeting 2 cards
    s4.addShape(R, { x: 0.5, y: 3.15, w: 4.3, h: 1.15, fill: { color: C.card }, shadow: mkS() });
    s4.addShape(R, { x: 0.5, y: 3.15, w: 0.05, h: 1.15, fill: { color: "1B5E20" } });
    s4.addText("3.2.1 Primary Target", { x: 0.7, y: 3.2, w: 3.5, h: 0.22, fontSize: 9, fontFace: "Arial", color: "66BB6A", bold: true, margin: 0 });
    bullets(s4, ["Young People in their 20s and 30s"], 0.7, 3.48, 3.9, 0.7, 8);

    s4.addShape(R, { x: 5.2, y: 3.15, w: 4.3, h: 1.15, fill: { color: C.card }, shadow: mkS() });
    s4.addShape(R, { x: 5.2, y: 3.15, w: 0.05, h: 1.15, fill: { color: "0D47A1" } });
    s4.addText("3.2.2 Secondary Target", { x: 5.4, y: 3.2, w: 3.5, h: 0.22, fontSize: 9, fontFace: "Arial", color: "42A5F5", bold: true, margin: 0 });
    bullets(s4, ["Eco Friendly", "Tech Savvy", "Entry Level Buyer for Electric Luxury Car"], 5.4, 3.48, 3.9, 0.7, 8);

    // Positioning + Recs
    s4.addShape(R, { x: 0.5, y: 4.5, w: 4.3, h: 0.95, fill: { color: C.card }, shadow: mkS() });
    s4.addShape(R, { x: 0.5, y: 4.5, w: 4.3, h: 0.035, fill: { color: C.red } });
    s4.addText("3.3 Positioning Strategy", { x: 0.65, y: 4.55, w: 3.5, h: 0.2, fontSize: 9, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    bullets(s4, ["Ultra high End > High End > SUV > Mid Level [Priced at $35,000]"], 0.65, 4.78, 3.95, 0.6, 8);

    s4.addShape(R, { x: 5.2, y: 4.5, w: 4.3, h: 0.95, fill: { color: C.card }, shadow: mkS() });
    s4.addShape(R, { x: 5.2, y: 4.5, w: 0.05, h: 0.95, fill: { color: "E65100" } });
    s4.addText("3.4 STP Recommendations", { x: 5.4, y: 4.55, w: 3.5, h: 0.2, fontSize: 9, fontFace: "Arial", color: "FF9800", bold: true, margin: 0 });
    bullets(s4, ["Target Women", "More affordable car options for lower income market"], 5.4, 4.78, 3.95, 0.6, 8);

    // ═══════════════════════════════════════
    // SLIDE 5: PART 4 — 4Ps
    // ═══════════════════════════════════════
    let s5 = pres.addSlide();
    s5.background = { color: C.dark };
    topBar(s5);
    hdr(s5, "04", "Marketing Mix \u2014 4Ps & Recommendations");

    // 4 cards in a 2x2 grid
    const psData = [
        { title: "4.1 Product", items: ["Great EV product", "Pure Electric", "Simple Structure", "Minimalist Design"] },
        { title: "4.2 Pricing", items: ["Priced Slightly higher than affordable range", "But due to subsidy", "Feels premium but affordable"] },
        { title: "4.3 Place", items: ["Online Shopping", "No External Dealerships", "All Stores owned and managed by Tesla"] },
        { title: "4.4 Promotion", items: ["$0 ad spent", "Elon Musk (200m followers)", "Bring in $19B of ad value", "Referral programs"] },
    ];
    psData.forEach((p, i) => {
        const col = i % 2, row = Math.floor(i / 2);
        const x = 0.5 + col * 4.6;
        const y = 1.3 + row * 1.95;
        s5.addShape(R, { x, y, w: 4.3, h: 1.75, fill: { color: C.card }, shadow: mkS() });
        s5.addShape(R, { x, y, w: 4.3, h: 0.035, fill: { color: C.red } });
        s5.addText(p.title, { x: x + 0.15, y: y + 0.08, w: 3.5, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
        bullets(s5, p.items, x + 0.15, y + 0.35, 4.0, 1.3, 8.5);
    });

    // Recommendations bar
    s5.addShape(R, { x: 0.5, y: 5.1, w: 9.0, h: 0.4, fill: { color: C.card }, shadow: mkS() });
    s5.addShape(R, { x: 0.5, y: 5.1, w: 0.05, h: 0.4, fill: { color: "E65100" } });
    s5.addText("4.5 Recommendations", { x: 0.7, y: 5.12, w: 1.8, h: 0.18, fontSize: 8, fontFace: "Arial", color: "FF9800", bold: true, margin: 0 });
    s5.addText("Spend more on digital ads, more product options, more customizations options", { x: 2.6, y: 5.12, w: 6.6, h: 0.3, fontSize: 8, fontFace: "Arial", color: C.dim, margin: 0 });

    // ═══════════════════════════════════════
    // SLIDE 6: REFERENCES
    // ═══════════════════════════════════════
    let s6 = pres.addSlide();
    s6.background = { color: C.dark };
    topBar(s6); botBar(s6);
    hdr(s6, "05", "References", "Sources & Citations");

    s6.addShape(R, { x: 0.5, y: 1.3, w: 9.0, h: 3.5, fill: { color: C.card }, shadow: mkS() });
    s6.addShape(R, { x: 0.5, y: 1.3, w: 9.0, h: 0.035, fill: { color: C.red } });
    s6.addText("References", { x: 0.65, y: 1.4, w: 4, h: 0.22, fontSize: 10, fontFace: "Arial", color: C.red, bold: true, margin: 0 });

    const refs = [
        "Model 3 \u2013 sports electric sedan | Tesla. Available at: https://www.tesla.com/model3 (Accessed: 23 March 2026).",
        "Admin (2025) Tesla\u2019s marketing strategy and its impact on EV Growth, Aeternus. Available at: https://www.aeternus.rs/driving-innovation-teslas-marketing-strategy-and-its-impact-on-electric-vehicle-adoption/ (Accessed: 24 March 2026).",
        "Hu, J. (2025) \u2018Research on the marketing strategy of Tesla marketing based on the 4P Theory and SWOT analytical method\u2019, Advances in Economics, Management and Political Sciences, 240(1), pp. 81\u201390.",
        "Hu, Z. (2022) \u2018Research on the consumer behavior characteristics and marketing strategy of new energy vehicles\u2019, BCP Business & Management, 31, pp. 168\u2013175.",
        "Mangram, M.E. (2012) \u2018The globalization of Tesla Motors: A strategic marketing plan analysis\u2019, Journal of Strategic Marketing, 20(4), pp. 289\u2013312.",
        "Young, J., Fattah, A. and Clark, P. (2020) Tesla: Direct Marketing Vs Franchising [Preprint].",
        "\u2018Tesla marketing analysis\u2019 (2023) Academic Journal of Business & Management, 5(2).",
    ];
    bullets(s6, refs, 0.65, 1.7, 8.7, 2.9, 7);

    s6.addText("Thank You", { x: 0, y: 5.1, w: 10, h: 0.35, fontSize: 18, fontFace: "Arial Black", color: C.red, align: "center" });

    await pres.writeFile({ fileName: "C:/laragon/www/tesla/Tesla_Model3_Marketing.pptx" });
    console.log("Done! Saved: Tesla_Model3_Marketing.pptx");
}

build().catch(err => { console.error(err); process.exit(1); });
