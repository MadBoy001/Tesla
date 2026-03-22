const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Icons
const { FaBullseye, FaChartLine, FaLightbulb, FaUsers, FaShieldAlt, FaCar, FaDollarSign, FaBookOpen, FaCogs, FaCheckCircle } = require("react-icons/fa");

function renderIconSvg(IconComponent, color = "#000000", size = 256) {
    return ReactDOMServer.renderToStaticMarkup(
        React.createElement(IconComponent, { color, size: String(size) })
    );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
    const svg = renderIconSvg(IconComponent, color, size);
    const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
    return "image/png;base64," + pngBuffer.toString("base64");
}

// Colors
const C = {
    dark: "0A0A0F",
    darkCard: "141420",
    red: "E82127",
    redDim: "3D0A0C",
    white: "F0F0F0",
    gray: "888899",
    dimText: "666677",
    border: "2A2A3A",
};

const mkShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.4 });

async function build() {
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    pres.author = "Tesla Marketing Analysis";
    pres.title = "Tesla Model 3 - Marketing Strategy Analysis";

    // Load icons
    const icons = {
        bullseye: await iconToBase64Png(FaBullseye, "#E82127"),
        chart: await iconToBase64Png(FaChartLine, "#E82127"),
        lightbulb: await iconToBase64Png(FaLightbulb, "#E82127"),
        users: await iconToBase64Png(FaUsers, "#E82127"),
        shield: await iconToBase64Png(FaShieldAlt, "#E82127"),
        car: await iconToBase64Png(FaCar, "#E82127"),
        dollar: await iconToBase64Png(FaDollarSign, "#E82127"),
        book: await iconToBase64Png(FaBookOpen, "#E82127"),
        cogs: await iconToBase64Png(FaCogs, "#E82127"),
        check: await iconToBase64Png(FaCheckCircle, "#E82127"),
    };

    // ═══════════════════════════════════════
    // SLIDE 1: TITLE
    // ═══════════════════════════════════════
    let s1 = pres.addSlide();
    s1.background = { color: C.dark };
    // Top red accent line
    s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    // Bottom red accent line
    s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.595, w: 10, h: 0.03, fill: { color: C.red } });
    // Subtle glow circle behind title
    s1.addShape(pres.shapes.OVAL, { x: 2.5, y: 0.8, w: 5, h: 4, fill: { color: C.redDim, transparency: 70 } });
    // Overline
    s1.addText("MARKETING STRATEGY ANALYSIS", {
        x: 0.5, y: 1.3, w: 9, h: 0.4, fontSize: 11, fontFace: "Arial",
        color: C.red, charSpacing: 6, align: "center", bold: true
    });
    // Main title
    s1.addText("TESLA", {
        x: 0.5, y: 1.7, w: 9, h: 1.4, fontSize: 72, fontFace: "Arial Black",
        color: C.white, align: "center", bold: true, margin: 0
    });
    s1.addText("MODEL 3", {
        x: 0.5, y: 2.9, w: 9, h: 0.9, fontSize: 44, fontFace: "Arial Black",
        color: C.red, align: "center", bold: true, margin: 0
    });
    // Divider
    s1.addShape(pres.shapes.RECTANGLE, { x: 4.3, y: 3.85, w: 1.4, h: 0.03, fill: { color: C.red } });
    // Tagline
    s1.addText("Redefining How Cars Are Sold in America", {
        x: 1, y: 4.0, w: 8, h: 0.5, fontSize: 14, fontFace: "Arial",
        color: C.dimText, align: "center", italic: true
    });
    // Car icon
    s1.addImage({ data: icons.car, x: 4.65, y: 4.6, w: 0.6, h: 0.6 });

    // ═══════════════════════════════════════
    // SLIDE 2: EXECUTIVE SUMMARY
    // ═══════════════════════════════════════
    let s2 = pres.addSlide();
    s2.background = { color: C.dark };
    s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });

    s2.addText("02", { x: 0.6, y: 0.3, w: 1.2, h: 0.8, fontSize: 36, fontFace: "Arial Black", color: C.red, margin: 0 });
    s2.addText("Executive Summary", { x: 0.6, y: 0.85, w: 6, h: 0.55, fontSize: 28, fontFace: "Arial Black", color: C.white, margin: 0 });

    // Left: Overview text
    s2.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.7, w: 4.4, h: 3.5, fill: { color: C.darkCard }, shadow: mkShadow() });
    s2.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.7, w: 4.4, h: 0.04, fill: { color: C.red } });
    s2.addImage({ data: icons.book, x: 0.9, y: 1.95, w: 0.35, h: 0.35 });
    s2.addText("Scope & Approach", { x: 1.35, y: 1.95, w: 3, h: 0.35, fontSize: 13, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s2.addText([
        { text: "This analysis examines Tesla Model 3's marketing strategy in the US market using established strategic frameworks.", options: { breakLine: true, fontSize: 11, color: C.dimText } },
        { text: "", options: { breakLine: true, fontSize: 6 } },
        { text: "Frameworks Applied:", options: { breakLine: true, fontSize: 11, color: C.white, bold: true } },
        { text: "SWOT Analysis", options: { bullet: true, breakLine: true, fontSize: 11, color: C.gray } },
        { text: "PESTEL Analysis", options: { bullet: true, breakLine: true, fontSize: 11, color: C.gray } },
        { text: "Porter's Five Forces", options: { bullet: true, breakLine: true, fontSize: 11, color: C.gray } },
        { text: "STP Framework", options: { bullet: true, breakLine: true, fontSize: 11, color: C.gray } },
        { text: "Marketing Mix (4Ps & 7Ps)", options: { bullet: true, fontSize: 11, color: C.gray } },
    ], { x: 0.9, y: 2.45, w: 3.8, h: 2.6 });

    // Right: Key findings cards
    s2.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.7, w: 4.3, h: 1.6, fill: { color: C.darkCard }, shadow: mkShadow() });
    s2.addText("52%", { x: 5.5, y: 1.8, w: 1.5, h: 0.7, fontSize: 32, fontFace: "Arial Black", color: C.red, margin: 0 });
    s2.addText("US EV Market Share\nTesla dominates the electric vehicle segment", { x: 5.5, y: 2.45, w: 3.8, h: 0.7, fontSize: 10, fontFace: "Arial", color: C.dimText, margin: 0 });

    s2.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.5, w: 4.3, h: 1.7, fill: { color: C.darkCard }, shadow: mkShadow() });
    s2.addText("$0", { x: 5.5, y: 3.6, w: 1.5, h: 0.7, fontSize: 32, fontFace: "Arial Black", color: C.red, margin: 0 });
    s2.addText("Annual Ad Spend\nTesla's zero-advertising model disrupts traditional\nautomotive marketing playbooks", { x: 5.5, y: 4.25, w: 3.8, h: 0.8, fontSize: 10, fontFace: "Arial", color: C.dimText, margin: 0 });

    // ═══════════════════════════════════════
    // SLIDE 3: MARKETING IMPORTANCE
    // ═══════════════════════════════════════
    let s3 = pres.addSlide();
    s3.background = { color: C.dark };
    s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });

    s3.addText("03", { x: 0.6, y: 0.3, w: 1.2, h: 0.8, fontSize: 36, fontFace: "Arial Black", color: C.red, margin: 0 });
    s3.addText("Marketing & Communication", { x: 0.6, y: 0.85, w: 8, h: 0.55, fontSize: 28, fontFace: "Arial Black", color: C.white, margin: 0 });
    s3.addText("Why strategic marketing matters in the EV revolution", { x: 0.6, y: 1.35, w: 6, h: 0.35, fontSize: 12, fontFace: "Arial", color: C.dimText, margin: 0 });

    // Three column cards: Why, What, How
    const col3Data = [
        { title: "WHY", icon: icons.lightbulb, points: ["Builds brand awareness without paid advertising", "Creates emotional connection with consumers", "Drives organic demand in competitive market"] },
        { title: "WHAT", icon: icons.bullseye, points: ["Product-led growth strategy", "CEO-driven social media presence", "Community-powered word of mouth"] },
        { title: "HOW", icon: icons.cogs, points: ["Direct-to-consumer sales model", "Referral programs ($1,000+ incentives)", "Experience-first Tesla stores & galleries"] },
    ];

    col3Data.forEach((col, i) => {
        const x = 0.6 + i * 3.1;
        s3.addShape(pres.shapes.RECTANGLE, { x, y: 1.95, w: 2.8, h: 3.2, fill: { color: C.darkCard }, shadow: mkShadow() });
        s3.addShape(pres.shapes.RECTANGLE, { x, y: 1.95, w: 2.8, h: 0.04, fill: { color: C.red } });
        s3.addImage({ data: col.icon, x: x + 0.2, y: 2.2, w: 0.3, h: 0.3 });
        s3.addText(col.title, { x: x + 0.6, y: 2.2, w: 2, h: 0.3, fontSize: 14, fontFace: "Arial Black", color: C.red, margin: 0 });
        const bullets = col.points.map((p, j) => ({
            text: p,
            options: { bullet: true, breakLine: j < col.points.length - 1, fontSize: 10, color: C.gray }
        }));
        s3.addText(bullets, { x: x + 0.2, y: 2.7, w: 2.4, h: 2.2 });
    });

    // ═══════════════════════════════════════
    // SLIDE 4: COMPANY & PRODUCT
    // ═══════════════════════════════════════
    let s4 = pres.addSlide();
    s4.background = { color: C.dark };
    s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });

    s4.addText("04", { x: 0.6, y: 0.3, w: 1.2, h: 0.8, fontSize: 36, fontFace: "Arial Black", color: C.red, margin: 0 });
    s4.addText("Company & Product", { x: 0.6, y: 0.85, w: 6, h: 0.55, fontSize: 28, fontFace: "Arial Black", color: C.white, margin: 0 });

    // Left: Company info
    s4.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.7, w: 4.4, h: 3.5, fill: { color: C.darkCard }, shadow: mkShadow() });
    s4.addImage({ data: icons.car, x: 0.9, y: 1.95, w: 0.35, h: 0.35 });
    s4.addText("Tesla Inc.", { x: 1.35, y: 1.95, w: 3, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s4.addText([
        { text: "Founded 2003 in Palo Alto, California", options: { bullet: true, breakLine: true, fontSize: 11, color: C.gray } },
        { text: "Mission: Accelerate the world's transition to sustainable energy", options: { bullet: true, breakLine: true, fontSize: 11, color: C.gray } },
        { text: "Market Cap: ~$800B+ (2025)", options: { bullet: true, breakLine: true, fontSize: 11, color: C.gray } },
        { text: "Global deliveries: 1.8M+ vehicles (2024)", options: { bullet: true, fontSize: 11, color: C.gray } },
    ], { x: 0.9, y: 2.5, w: 3.8, h: 2.2 });

    // Right: Product stats as big numbers
    s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.7, w: 4.3, h: 1.6, fill: { color: C.darkCard }, shadow: mkShadow() });
    s4.addText("$38,990", { x: 5.5, y: 1.8, w: 3, h: 0.7, fontSize: 32, fontFace: "Arial Black", color: C.red, margin: 0 });
    s4.addText("Starting MSRP — Model 3 launched 2017\nMost affordable Tesla, best-selling EV in the US", { x: 5.5, y: 2.45, w: 3.8, h: 0.6, fontSize: 10, fontFace: "Arial", color: C.dimText, margin: 0 });

    s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.5, w: 4.3, h: 1.7, fill: { color: C.darkCard }, shadow: mkShadow() });
    s4.addText("Value Proposition", { x: 5.5, y: 3.6, w: 3, h: 0.4, fontSize: 13, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s4.addText([
        { text: "Zero emissions, sustainable energy", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "Cutting-edge tech: Autopilot, OTA updates", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "60% lower fuel costs vs ICE vehicles", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "50% less maintenance required", options: { bullet: true, fontSize: 10, color: C.gray } },
    ], { x: 5.5, y: 4.05, w: 3.8, h: 1.0 });

    // ═══════════════════════════════════════
    // SLIDE 5: PESTEL & SWOT
    // ═══════════════════════════════════════
    let s5 = pres.addSlide();
    s5.background = { color: C.dark };
    s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });

    s5.addText("05", { x: 0.6, y: 0.3, w: 1.2, h: 0.8, fontSize: 36, fontFace: "Arial Black", color: C.red, margin: 0 });
    s5.addText("SWOT & PESTEL Analysis", { x: 0.6, y: 0.85, w: 8, h: 0.55, fontSize: 28, fontFace: "Arial Black", color: C.white, margin: 0 });

    // SWOT 2x2 grid
    const swotData = [
        { label: "STRENGTHS", color: "1B5E20", items: "Brand loyalty & recognition\nTech & innovation leadership\nDirect sales model" },
        { label: "WEAKNESSES", color: "B71C1C", items: "Quality control issues\nLimited service network\nHigh price perception" },
        { label: "OPPORTUNITIES", color: "0D47A1", items: "Global EV market growth ($67B+)\n$7,500 federal tax credit\nAutonomous driving expansion" },
        { label: "THREATS", color: "E65100", items: "Rising competition (BYD, Ford)\nRaw material price volatility\nRegulatory changes" },
    ];
    swotData.forEach((s, i) => {
        const col = i % 2;
        const row = Math.floor(i / 2);
        const x = 0.6 + col * 4.6;
        const y = 1.6 + row * 1.85;
        s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.3, h: 1.65, fill: { color: C.darkCard }, shadow: mkShadow() });
        s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.65, fill: { color: s.color } });
        s5.addText(s.label, { x: x + 0.2, y: y + 0.1, w: 3.8, h: 0.3, fontSize: 11, fontFace: "Arial", color: s.color, bold: true, margin: 0 });
        s5.addText(s.items, { x: x + 0.2, y: y + 0.45, w: 3.8, h: 1.1, fontSize: 10, fontFace: "Arial", color: C.gray, margin: 0 });
    });

    // ═══════════════════════════════════════
    // SLIDE 6: PORTER'S FIVE FORCES
    // ═══════════════════════════════════════
    let s6 = pres.addSlide();
    s6.background = { color: C.dark };
    s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });

    s6.addText("06", { x: 0.6, y: 0.3, w: 1.2, h: 0.8, fontSize: 36, fontFace: "Arial Black", color: C.red, margin: 0 });
    s6.addText("Porter's Five Forces", { x: 0.6, y: 0.85, w: 6, h: 0.55, fontSize: 28, fontFace: "Arial Black", color: C.white, margin: 0 });

    // Table
    const headerRow = [
        { text: "Force", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 11, fontFace: "Arial", align: "left" } },
        { text: "Impact", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 11, fontFace: "Arial", align: "center" } },
        { text: "Analysis", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 11, fontFace: "Arial", align: "left" } },
    ];
    const makeRow = (force, impact, analysis) => [
        { text: force, options: { fill: { color: C.darkCard }, color: C.white, fontSize: 10, fontFace: "Arial", align: "left" } },
        { text: impact, options: { fill: { color: C.darkCard }, color: impact === "High" ? "FF5252" : impact === "Moderate" ? "FFB74D" : "66BB6A", fontSize: 10, fontFace: "Arial", bold: true, align: "center" } },
        { text: analysis, options: { fill: { color: C.darkCard }, color: C.gray, fontSize: 10, fontFace: "Arial", align: "left" } },
    ];
    const tableData = [
        headerRow,
        makeRow("Supplier Bargaining Power", "Moderate", "Battery dependency on limited lithium/cobalt suppliers; Tesla's Gigafactory mitigates risk"),
        makeRow("Buyer Bargaining Power", "Moderate", "Growing EV alternatives increase choice; strong brand loyalty offsets switching"),
        makeRow("Threat of New Entrants", "Low", "High capital requirements ($B+ investment); regulatory and infrastructure barriers"),
        makeRow("Threat of Substitutes", "Moderate", "Hybrids, hydrogen, public transit offer alternatives; EV adoption trend reduces threat"),
        makeRow("Industry Rivalry", "High", "Intense competition from Ford, GM, BYD, Hyundai, VW entering EV space"),
    ];
    s6.addTable(tableData, {
        x: 0.6, y: 1.6, w: 8.8, h: 3.6,
        border: { pt: 0.5, color: C.border },
        colW: [2.2, 1.0, 5.6],
        rowH: [0.45, 0.63, 0.63, 0.63, 0.63, 0.63],
    });

    // ═══════════════════════════════════════
    // SLIDE 7: STP ANALYSIS
    // ═══════════════════════════════════════
    let s7 = pres.addSlide();
    s7.background = { color: C.dark };
    s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });

    s7.addText("07", { x: 0.6, y: 0.3, w: 1.2, h: 0.8, fontSize: 36, fontFace: "Arial Black", color: C.red, margin: 0 });
    s7.addText("STP Analysis", { x: 0.6, y: 0.85, w: 6, h: 0.55, fontSize: 28, fontFace: "Arial Black", color: C.white, margin: 0 });
    s7.addText("Segmentation, Targeting & Positioning", { x: 0.6, y: 1.3, w: 6, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.dimText, margin: 0 });

    // Three vertical cards
    const stpCards = [
        {
            title: "SEGMENTATION", icon: icons.users,
            items: ["Demographic: Age 25-45, Income $75K+", "Psychographic: Tech-forward, eco-conscious", "Behavioral: Early adopters, high brand loyalty (91%)"]
        },
        {
            title: "TARGETING", icon: icons.bullseye,
            items: ["Primary: Urban tech professionals", "Secondary: Eco-conscious families", "Tertiary: Fleet operators & corporate buyers"]
        },
        {
            title: "POSITIONING", icon: icons.chart,
            items: ["Innovation leader at accessible price point", "Premium tech brand, not just automaker", "Sustainability as core brand identity"]
        },
    ];
    stpCards.forEach((card, i) => {
        const x = 0.6 + i * 3.1;
        s7.addShape(pres.shapes.RECTANGLE, { x, y: 1.85, w: 2.8, h: 3.3, fill: { color: C.darkCard }, shadow: mkShadow() });
        s7.addShape(pres.shapes.RECTANGLE, { x, y: 1.85, w: 2.8, h: 0.04, fill: { color: C.red } });
        s7.addImage({ data: card.icon, x: x + 0.2, y: 2.1, w: 0.3, h: 0.3 });
        s7.addText(card.title, { x: x + 0.6, y: 2.1, w: 2, h: 0.3, fontSize: 12, fontFace: "Arial Black", color: C.red, margin: 0 });
        const bullets = card.items.map((p, j) => ({
            text: p,
            options: { bullet: true, breakLine: j < card.items.length - 1, fontSize: 10, color: C.gray }
        }));
        s7.addText(bullets, { x: x + 0.2, y: 2.6, w: 2.4, h: 2.3 });
    });

    // ═══════════════════════════════════════
    // SLIDE 8: MARKETING MIX (4Ps)
    // ═══════════════════════════════════════
    let s8 = pres.addSlide();
    s8.background = { color: C.dark };
    s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });

    s8.addText("08", { x: 0.6, y: 0.3, w: 1.2, h: 0.8, fontSize: 36, fontFace: "Arial Black", color: C.red, margin: 0 });
    s8.addText("Marketing Mix — 4Ps", { x: 0.6, y: 0.85, w: 6, h: 0.55, fontSize: 28, fontFace: "Arial Black", color: C.white, margin: 0 });

    // 2x2 grid
    const mixData = [
        { title: "PRODUCT", icon: icons.car, desc: "Premium EV with OTA updates, Autopilot, minimalist interior design, 358-mile range" },
        { title: "PRICE", icon: icons.dollar, desc: "$38,990-$52,990 range. Transparent pricing, no dealer markup, federal tax credit eligible" },
        { title: "PLACE", icon: icons.cogs, desc: "Direct-to-consumer model. Online ordering, Tesla stores, home delivery, 15,000+ Superchargers" },
        { title: "PROMOTION", icon: icons.lightbulb, desc: "Zero paid advertising. CEO social media, word of mouth, referral programs, earned media" },
    ];
    mixData.forEach((item, i) => {
        const col = i % 2;
        const row = Math.floor(i / 2);
        const x = 0.6 + col * 4.6;
        const y = 1.6 + row * 1.9;
        s8.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.3, h: 1.7, fill: { color: C.darkCard }, shadow: mkShadow() });
        s8.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.3, h: 0.04, fill: { color: C.red } });
        s8.addImage({ data: item.icon, x: x + 0.2, y: y + 0.2, w: 0.3, h: 0.3 });
        s8.addText(item.title, { x: x + 0.6, y: y + 0.2, w: 3.2, h: 0.3, fontSize: 13, fontFace: "Arial Black", color: C.red, margin: 0 });
        s8.addText(item.desc, { x: x + 0.2, y: y + 0.65, w: 3.9, h: 0.9, fontSize: 10, fontFace: "Arial", color: C.gray, margin: 0 });
    });

    // ═══════════════════════════════════════
    // SLIDE 9: RECOMMENDATIONS
    // ═══════════════════════════════════════
    let s9 = pres.addSlide();
    s9.background = { color: C.dark };
    s9.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });

    s9.addText("09", { x: 0.6, y: 0.3, w: 1.2, h: 0.8, fontSize: 36, fontFace: "Arial Black", color: C.red, margin: 0 });
    s9.addText("Recommendations", { x: 0.6, y: 0.85, w: 6, h: 0.55, fontSize: 28, fontFace: "Arial Black", color: C.white, margin: 0 });

    // Left column: 4Ps recommendations
    s9.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.6, w: 4.3, h: 3.6, fill: { color: C.darkCard }, shadow: mkShadow() });
    s9.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.6, w: 0.06, h: 3.6, fill: { color: C.red } });
    s9.addText("Marketing Mix (4Ps)", { x: 0.85, y: 1.7, w: 3.8, h: 0.35, fontSize: 13, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s9.addText([
        { text: "Product: Launch affordable sub-$30K variant to expand TAM", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "Price: Implement dynamic pricing based on demand signals", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "Place: Expand service center network by 40%", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "Promotion: Targeted digital campaigns for emerging segments", options: { bullet: true, fontSize: 10, color: C.gray } },
    ], { x: 0.85, y: 2.15, w: 3.8, h: 1.4 });

    s9.addText("Services (7Ps Extension)", { x: 0.85, y: 3.45, w: 3.8, h: 0.35, fontSize: 13, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s9.addText([
        { text: "Process: Streamline service booking and delivery logistics", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "People: Invest in customer service training programs", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "Physical Evidence: Enhance showroom experience", options: { bullet: true, fontSize: 10, color: C.gray } },
    ], { x: 0.85, y: 3.85, w: 3.8, h: 1.2 });

    // Right column: STP recommendations
    s9.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.6, w: 4.4, h: 3.6, fill: { color: C.darkCard }, shadow: mkShadow() });
    s9.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.6, w: 0.06, h: 3.6, fill: { color: "0D47A1" } });
    s9.addText("STP Strategy", { x: 5.45, y: 1.7, w: 3.8, h: 0.35, fontSize: 13, fontFace: "Arial", color: "5C9CE6", bold: true, margin: 0 });
    s9.addText([
        { text: "Broaden target to include mid-income ($50K+) consumers", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "Strengthen positioning against Chinese EV brands (BYD)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "Develop fleet/B2B segment for corporate sustainability goals", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "Localize marketing for regional US market differences", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "Leverage owner testimonials as positioning tool", options: { bullet: true, fontSize: 10, color: C.gray } },
    ], { x: 5.45, y: 2.15, w: 3.8, h: 2.0 });

    // ═══════════════════════════════════════
    // SLIDE 10: REFERENCES
    // ═══════════════════════════════════════
    let s10 = pres.addSlide();
    s10.background = { color: C.dark };
    s10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    s10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.595, w: 10, h: 0.03, fill: { color: C.red } });

    s10.addText("10", { x: 0.6, y: 0.3, w: 1.2, h: 0.8, fontSize: 36, fontFace: "Arial Black", color: C.red, margin: 0 });
    s10.addText("References", { x: 0.6, y: 0.85, w: 6, h: 0.55, fontSize: 28, fontFace: "Arial Black", color: C.white, margin: 0 });

    s10.addText([
        { text: "Statista (2025). US Electric Vehicle Market Report. Available at: statista.com", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "", options: { breakLine: true, fontSize: 5 } },
        { text: "Tesla Inc. (2024). Annual Report & Form 10-K Filing. SEC.gov", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "", options: { breakLine: true, fontSize: 5 } },
        { text: "Porter, M.E. (1979). 'How Competitive Forces Shape Strategy'. Harvard Business Review, 57(2), pp.137-145.", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "", options: { breakLine: true, fontSize: 5 } },
        { text: "Kotler, P. & Keller, K.L. (2016). Marketing Management. 15th Edition. Pearson Education.", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "", options: { breakLine: true, fontSize: 5 } },
        { text: "International Energy Agency (2025). Global EV Outlook 2025. IEA Publications.", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "", options: { breakLine: true, fontSize: 5 } },
        { text: "Bloomberg NEF (2025). Electric Vehicle Market Analysis & Forecast.", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "", options: { breakLine: true, fontSize: 5 } },
        { text: "McKinsey & Company (2024). 'The Future of Mobility: Electric Vehicles in America'.", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gray } },
        { text: "", options: { breakLine: true, fontSize: 5 } },
        { text: "Deloitte (2025). Global Automotive Consumer Study — US Market Focus.", options: { bullet: true, fontSize: 10, color: C.gray } },
    ], { x: 0.8, y: 1.5, w: 8.4, h: 3.8 });

    // Thank you footer
    s10.addText("Thank You", { x: 0, y: 4.85, w: 10, h: 0.5, fontSize: 18, fontFace: "Arial Black", color: C.red, align: "center" });

    // ═══════════════════════════════════════
    // SAVE
    // ═══════════════════════════════════════
    const outputPath = "C:/laragon/www/tesla/Tesla_Model3_Marketing_Strategy.pptx";
    await pres.writeFile({ fileName: outputPath });
    console.log("Presentation saved to: " + outputPath);
}

build().catch(err => { console.error(err); process.exit(1); });
