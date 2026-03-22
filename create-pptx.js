const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

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

const C = {
    dark: "0A0A0F", darkCard: "141420", red: "E82127", redDim: "3D0A0C",
    white: "F0F0F0", gray: "888899", dimText: "666677", border: "2A2A3A",
};
const mkShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.4 });

function bullets(items, opts = {}) {
    return items.map((text, i) => ({
        text,
        options: {
            bullet: true, breakLine: i < items.length - 1,
            fontSize: opts.fontSize || 10, color: opts.color || C.gray,
        }
    }));
}

// Narrative hook — italic quote at top of slide
function addHook(slide, text, x = 0.6, y = 1.25) {
    slide.addText(text, {
        x, y, w: 8.8, h: 0.4, fontSize: 12, fontFace: "Georgia",
        color: C.dimText, italic: true, margin: 0
    });
}

// Slide header with number
function addHeader(slide, num, title) {
    slide.addText(num, { x: 0.6, y: 0.2, w: 1.2, h: 0.7, fontSize: 32, fontFace: "Arial Black", color: C.red, margin: 0 });
    slide.addText(title, { x: 0.6, y: 0.75, w: 8, h: 0.5, fontSize: 26, fontFace: "Arial Black", color: C.white, margin: 0 });
}

async function build() {
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    pres.author = "Tesla Marketing Analysis";
    pres.title = "Tesla Model 3 - Marketing Strategy Analysis";

    const icons = {
        bullseye: await iconToBase64Png(FaBullseye, "#E82127"),
        chart: await iconToBase64Png(FaChartLine, "#E82127"),
        lightbulb: await iconToBase64Png(FaLightbulb, "#E82127"),
        users: await iconToBase64Png(FaUsers, "#E82127"),
        car: await iconToBase64Png(FaCar, "#E82127"),
        dollar: await iconToBase64Png(FaDollarSign, "#E82127"),
        book: await iconToBase64Png(FaBookOpen, "#E82127"),
        cogs: await iconToBase64Png(FaCogs, "#E82127"),
        check: await iconToBase64Png(FaCheckCircle, "#E82127"),
    };

    // ═══════════════════════════════════════
    // SLIDE 1: TITLE — Set the scene
    // ═══════════════════════════════════════
    let s1 = pres.addSlide();
    s1.background = { color: C.dark };
    s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.595, w: 10, h: 0.03, fill: { color: C.red } });
    s1.addShape(pres.shapes.OVAL, { x: 2.5, y: 0.8, w: 5, h: 4, fill: { color: C.redDim, transparency: 70 } });
    s1.addText("MARKETING STRATEGY ANALYSIS", {
        x: 0.5, y: 1.1, w: 9, h: 0.4, fontSize: 11, fontFace: "Arial",
        color: C.red, charSpacing: 6, align: "center", bold: true
    });
    s1.addText("TESLA", {
        x: 0.5, y: 1.5, w: 9, h: 1.4, fontSize: 72, fontFace: "Arial Black",
        color: C.white, align: "center", bold: true, margin: 0
    });
    s1.addText("MODEL 3", {
        x: 0.5, y: 2.75, w: 9, h: 0.9, fontSize: 44, fontFace: "Arial Black",
        color: C.red, align: "center", bold: true, margin: 0
    });
    s1.addShape(pres.shapes.RECTANGLE, { x: 4.3, y: 3.7, w: 1.4, h: 0.03, fill: { color: C.red } });
    // Story opener
    s1.addText("What happens when a company sells 1.8 million cars a year\nwithout spending a single dollar on advertising?", {
        x: 1.5, y: 3.85, w: 7, h: 0.7, fontSize: 13, fontFace: "Georgia",
        color: C.dimText, align: "center", italic: true
    });
    s1.addImage({ data: icons.car, x: 4.65, y: 4.65, w: 0.55, h: 0.55 });

    // ═══════════════════════════════════════
    // SLIDE 2: THE SETUP — Why this matters
    // ═══════════════════════════════════════
    let s2 = pres.addSlide();
    s2.background = { color: C.dark };
    s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    addHeader(s2, "01", "The $67 Billion Question");
    addHook(s2, "Every automaker spends billions on ads. Tesla spends zero. Yet they own 52% of the US EV market. How?");

    // Three big stat cards telling the story
    const stats = [
        { val: "$67B", label: "US EV Market Size", desc: "The prize everyone is\nfighting for (Statista, 2025)" },
        { val: "52%", label: "Tesla's Market Share", desc: "More than all competitors\ncombined — without a single ad" },
        { val: "$0", label: "Advertising Budget", desc: "While GM spends $3.5B and\nFord $2.8B annually on ads" },
    ];
    stats.forEach((s, i) => {
        const x = 0.6 + i * 3.1;
        s2.addShape(pres.shapes.RECTANGLE, { x, y: 1.9, w: 2.8, h: 2.6, fill: { color: C.darkCard }, shadow: mkShadow() });
        s2.addShape(pres.shapes.RECTANGLE, { x, y: 1.9, w: 2.8, h: 0.04, fill: { color: C.red } });
        s2.addText(s.val, { x, y: 2.1, w: 2.8, h: 0.8, fontSize: 36, fontFace: "Arial Black", color: C.red, align: "center", margin: 0 });
        s2.addText(s.label, { x, y: 2.85, w: 2.8, h: 0.35, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, align: "center", margin: 0 });
        s2.addText(s.desc, { x: x + 0.2, y: 3.25, w: 2.4, h: 0.8, fontSize: 10, fontFace: "Arial", color: C.dimText, align: "center", margin: 0 });
    });

    // Transition text
    s2.addText("To understand Tesla's dominance, we need to look at how they\nrewrote the rules of automotive marketing from the ground up.", {
        x: 0.6, y: 4.7, w: 8.8, h: 0.6, fontSize: 10, fontFace: "Georgia", color: C.dimText, italic: true, margin: 0
    });

    // ═══════════════════════════════════════
    // SLIDE 3: THE REBEL — Marketing approach
    // ═══════════════════════════════════════
    let s3 = pres.addSlide();
    s3.background = { color: C.dark };
    s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    addHeader(s3, "02", "The Anti-Marketing Playbook");
    addHook(s3, "Tesla didn't just skip advertising — they built an entirely new model where the product markets itself.");

    const plays = [
        { title: "The Product IS the Ad", icon: icons.car, items: [
            "91% owner satisfaction becomes word-of-mouth (JD Power)",
            "Every Tesla on the road is a rolling billboard",
            "OTA updates (40+/year) generate constant media coverage",
        ]},
        { title: "CEO as Chief Marketer", icon: icons.lightbulb, items: [
            "200M+ social media followers — largest of any CEO",
            "$19B in earned media value annually (Bloomberg)",
            "Product launches become cultural events (250M+ impressions)",
        ]},
        { title: "Owners as Salesforce", icon: icons.users, items: [
            "Referral program drove 425,000+ new orders",
            "Tesla ranks #1 in brand loyalty (Consumer Reports)",
            "Community forums and owner clubs amplify the message",
        ]},
        { title: "Direct-to-Consumer", icon: icons.cogs, items: [
            "No dealerships = no $2,000+ markup, no brand dilution",
            "Online ordering: configure & buy in 8 minutes",
            "200+ Tesla stores in premium malls for brand exposure",
        ]},
    ];
    plays.forEach((p, i) => {
        const col = i % 2, row = Math.floor(i / 2);
        const x = 0.6 + col * 4.6, y = 1.75 + row * 1.8;
        s3.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.3, h: 1.6, fill: { color: C.darkCard }, shadow: mkShadow() });
        s3.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.3, h: 0.04, fill: { color: C.red } });
        s3.addImage({ data: p.icon, x: x + 0.15, y: y + 0.15, w: 0.28, h: 0.28 });
        s3.addText(p.title, { x: x + 0.5, y: y + 0.15, w: 3.5, h: 0.28, fontSize: 12, fontFace: "Arial Black", color: C.red, margin: 0 });
        s3.addText(bullets(p.items, { fontSize: 9 }), { x: x + 0.15, y: y + 0.5, w: 4.0, h: 1.0 });
    });

    // ═══════════════════════════════════════
    // SLIDE 4: THE HERO — Model 3
    // ═══════════════════════════════════════
    let s4 = pres.addSlide();
    s4.background = { color: C.dark };
    s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    addHeader(s4, "03", "Meet the Model 3");
    addHook(s4, "In 2017, 400,000 people paid $1,000 to reserve a car they'd never seen in person. That's not a product launch — it's a movement.");

    // Left: The story
    s4.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.85, w: 4.4, h: 3.5, fill: { color: C.darkCard }, shadow: mkShadow() });
    s4.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.85, w: 0.06, h: 3.5, fill: { color: C.red } });
    s4.addText("The Car That Changed Everything", { x: 0.85, y: 1.95, w: 3.8, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s4.addText(bullets([
        "Launched March 2017 — 400K pre-orders in the first week",
        "Mission: Make EVs accessible, not just aspirational",
        "Best-selling EV in the US for 6 consecutive years (2019-2024)",
        "Range: 272-358 miles — eliminated 'range anxiety' narrative",
        "5-star NHTSA safety rating in every category tested",
        "Autopilot standard; Full Self-Driving: $12K or $199/mo",
        "Revenue contribution: ~40% of Tesla's $96.8B FY2024",
    ], { fontSize: 9 }), { x: 0.85, y: 2.35, w: 3.95, h: 2.8 });

    // Right: Big price + value props
    s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.85, w: 4.3, h: 1.5, fill: { color: C.darkCard }, shadow: mkShadow() });
    s4.addText("$38,990", { x: 5.3, y: 1.95, w: 4.3, h: 0.7, fontSize: 38, fontFace: "Arial Black", color: C.red, align: "center", margin: 0 });
    s4.addText("Starting MSRP — or just $31,490 after $7,500 federal tax credit", { x: 5.5, y: 2.65, w: 3.9, h: 0.5, fontSize: 10, fontFace: "Arial", color: C.dimText, align: "center", margin: 0 });

    s4.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.55, w: 4.3, h: 1.8, fill: { color: C.darkCard }, shadow: mkShadow() });
    s4.addText("Why Owners Never Go Back", { x: 5.5, y: 3.65, w: 3.9, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s4.addText(bullets([
        "60% lower fuel cost — $0.04/mile vs $0.12/mile gasoline (DOE)",
        "50% less maintenance — no oil changes, brakes last 100K+ mi",
        "Car improves after purchase — 40+ OTA updates per year",
        "15,000+ Supercharger stations — largest fast-charging network",
    ], { fontSize: 9 }), { x: 5.5, y: 4.05, w: 3.9, h: 1.2 });

    // ═══════════════════════════════════════
    // SLIDE 5: THE LANDSCAPE — SWOT
    // ═══════════════════════════════════════
    let s5 = pres.addSlide();
    s5.background = { color: C.dark };
    s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    addHeader(s5, "04", "Strengths, Cracks & Crossroads");
    addHook(s5, "No company is invincible. Tesla's greatest strengths also hide its biggest vulnerabilities.");

    const swot = [
        { label: "STRENGTHS", color: "1B5E20", items: [
            "91% satisfaction — owners become evangelists",
            "Only automaker with full software-defined vehicle",
            "Vertical integration: batteries, chips, charging",
            "Direct sales = $2,000+ saved per car vs dealerships",
        ]},
        { label: "WEAKNESSES", color: "B71C1C", items: [
            "Quality issues: 23% report panel gaps or paint defects",
            "Service centers: avg 45-min drive (vs 15-min for Toyota)",
            "CEO dependency — one tweet can move stock ±10%",
            "Customer service scored below industry avg (ACSI)",
        ]},
        { label: "OPPORTUNITIES", color: "0D47A1", items: [
            "US EV market hitting $67B by 2026 — still only 9% of auto",
            "$7,500 IRA tax credit accelerating mass adoption",
            "Autonomous driving: $300B TAM by 2030 (Goldman Sachs)",
            "Energy storage + solar creates recurring revenue streams",
        ]},
        { label: "THREATS", color: "E65100", items: [
            "BYD outsold Tesla globally in Q4 2023 — the gap is closing",
            "Legacy OEMs pouring $500B+ into EV catch-up collectively",
            "Lithium prices swung 300% in 2022 — supply chain risk",
            "Tax credit phase-out could eliminate $7,500 price advantage",
        ]},
    ];
    swot.forEach((s, i) => {
        const col = i % 2, row = Math.floor(i / 2);
        const x = 0.6 + col * 4.6, y = 1.75 + row * 1.85;
        s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.3, h: 1.7, fill: { color: C.darkCard }, shadow: mkShadow() });
        s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.7, fill: { color: s.color } });
        s5.addText(s.label, { x: x + 0.2, y: y + 0.08, w: 3.8, h: 0.25, fontSize: 11, fontFace: "Arial", color: s.color, bold: true, margin: 0 });
        s5.addText(bullets(s.items, { fontSize: 9 }), { x: x + 0.2, y: y + 0.38, w: 3.9, h: 1.2 });
    });

    // ═══════════════════════════════════════
    // SLIDE 6: THE BATTLEFIELD — Porter's
    // ═══════════════════════════════════════
    let s6 = pres.addSlide();
    s6.background = { color: C.dark };
    s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    addHeader(s6, "05", "The Competitive Battlefield");
    addHook(s6, "Tesla built a moat — but competitors are learning to swim. Here's what the competitive landscape really looks like.");

    const headerRow = [
        { text: "Competitive Force", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 10, fontFace: "Arial", align: "left" } },
        { text: "Threat", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 10, fontFace: "Arial", align: "center" } },
        { text: "What This Means for Tesla", options: { fill: { color: C.red }, color: "FFFFFF", bold: true, fontSize: 10, fontFace: "Arial", align: "left" } },
    ];
    const makeRow = (force, impact, analysis) => [
        { text: force, options: { fill: { color: C.darkCard }, color: C.white, fontSize: 9, fontFace: "Arial", align: "left", bold: true } },
        { text: impact, options: { fill: { color: C.darkCard }, color: impact === "HIGH" ? "FF5252" : impact === "MOD" ? "FFB74D" : "66BB6A", fontSize: 10, fontFace: "Arial", bold: true, align: "center" } },
        { text: analysis, options: { fill: { color: C.darkCard }, color: C.gray, fontSize: 8.5, fontFace: "Arial", align: "left" } },
    ];
    const tableData = [
        headerRow,
        makeRow("Supplier Power", "MOD", "Battery materials are scarce, but Tesla's Gigafactories and 4680 cells reduce dependency on any single supplier"),
        makeRow("Buyer Power", "MOD", "30+ competing EVs give buyers choice — but Tesla's 91% satisfaction and Supercharger lock-in keep them loyal"),
        makeRow("New Entrants", "LOW", "Building an EV company costs $5B+ minimum. Charging infrastructure and software expertise create massive barriers"),
        makeRow("Substitutes", "MOD", "Hybrids and hydrogen exist, but declining battery costs and government incentives are making EVs the default choice"),
        makeRow("Industry Rivalry", "HIGH", "This is the real threat. BYD, Ford, GM, Hyundai, VW are spending $500B+ combined to catch Tesla"),
    ];
    s6.addTable(tableData, {
        x: 0.6, y: 1.75, w: 8.8, h: 3.5,
        border: { pt: 0.5, color: C.border },
        colW: [1.5, 0.8, 6.5],
        rowH: [0.4, 0.62, 0.62, 0.62, 0.62, 0.62],
    });

    s6.addText("The verdict? Tesla's moat is real — but it's eroding. Speed of execution will determine the next chapter.", {
        x: 0.6, y: 5.0, w: 8.8, h: 0.35, fontSize: 10, fontFace: "Georgia", color: C.dimText, italic: true, margin: 0
    });

    // ═══════════════════════════════════════
    // SLIDE 7: THE AUDIENCE — STP
    // ═══════════════════════════════════════
    let s7 = pres.addSlide();
    s7.background = { color: C.dark };
    s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    addHeader(s7, "06", "Who Buys a Tesla — and Why");
    addHook(s7, "Tesla doesn't sell to everyone. They sell to people who want to feel like they're part of the future.");

    const stpCards = [
        { title: "WHO THEY ARE", icon: icons.users, items: [
            "Age 25-45, household income $75K+, 72% male",
            "College-educated, urban professionals",
            "3.2x more likely to own smart home devices",
            "Top 5 states = 52% of sales (CA, TX, FL, NY, WA)",
        ]},
        { title: "WHAT DRIVES THEM", icon: icons.lightbulb, items: [
            "Tech-forward: they buy innovation, not transportation",
            "Eco-conscious: sustainability is identity, not feature",
            "Status-seeking: Tesla is a lifestyle statement",
            "Early adopters on the Rogers diffusion curve",
        ]},
        { title: "WHERE TESLA POSITIONS", icon: icons.bullseye, items: [
            "Intersection of HIGH innovation + MID price",
            "Premium tech brand — not 'just another automaker'",
            "Only EV with integrated ecosystem (car + energy + charging)",
            "Primary: Urban tech pros / Secondary: Eco families",
        ]},
    ];
    stpCards.forEach((card, i) => {
        const x = 0.6 + i * 3.1;
        s7.addShape(pres.shapes.RECTANGLE, { x, y: 1.75, w: 2.8, h: 3.5, fill: { color: C.darkCard }, shadow: mkShadow() });
        s7.addShape(pres.shapes.RECTANGLE, { x, y: 1.75, w: 2.8, h: 0.04, fill: { color: C.red } });
        s7.addImage({ data: card.icon, x: x + 0.15, y: 1.95, w: 0.28, h: 0.28 });
        s7.addText(card.title, { x: x + 0.5, y: 1.95, w: 2, h: 0.28, fontSize: 11, fontFace: "Arial Black", color: C.red, margin: 0 });
        s7.addText(bullets(card.items, { fontSize: 9 }), { x: x + 0.15, y: 2.4, w: 2.5, h: 2.6 });
    });

    // ═══════════════════════════════════════
    // SLIDE 8: THE PLAYBOOK — 4Ps
    // ═══════════════════════════════════════
    let s8 = pres.addSlide();
    s8.background = { color: C.dark };
    s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    addHeader(s8, "07", "The 4P Playbook");
    addHook(s8, "Most companies optimize their marketing mix. Tesla reinvented it — every P breaks an industry convention.");

    const mixData = [
        { title: "PRODUCT — The Car Sells Itself", icon: icons.car, items: [
            "358mi range, 15\" touchscreen, minimalist interior",
            "Autopilot + FSD = car improves after you buy it",
            "40+ OTA updates/year — no other automaker does this",
            "5-star NHTSA rating in every single category",
        ]},
        { title: "PRICE — Transparent & Disruptive", icon: icons.dollar, items: [
            "$38,990 base — or $31,490 after $7,500 tax credit",
            "Same price for everyone — no haggling, no dealer games",
            "TCO saves $6,000-8,000 over 5yr vs BMW 3 Series",
            "Insurance through Tesla at 20-30% lower rates",
        ]},
        { title: "PLACE — Cut Out the Middleman", icon: icons.cogs, items: [
            "100% direct-to-consumer — first automaker to do this",
            "Buy online in 8 minutes, delivered to your driveway",
            "200+ stores in malls — test drive, don't get sold to",
            "15,000+ Superchargers now the US standard (NACS)",
        ]},
        { title: "PROMOTION — Let Others Do It", icon: icons.lightbulb, items: [
            "$0 ad spend — while competitors burn $3-5B annually",
            "CEO's tweets = $19B earned media value (Bloomberg)",
            "Referral program alone drove 425K+ vehicle sales",
            "Every launch is a cultural moment, not a commercial",
        ]},
    ];
    mixData.forEach((item, i) => {
        const col = i % 2, row = Math.floor(i / 2);
        const x = 0.6 + col * 4.6, y = 1.75 + row * 1.8;
        s8.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.3, h: 1.6, fill: { color: C.darkCard }, shadow: mkShadow() });
        s8.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.3, h: 0.04, fill: { color: C.red } });
        s8.addImage({ data: item.icon, x: x + 0.15, y: y + 0.15, w: 0.25, h: 0.25 });
        s8.addText(item.title, { x: x + 0.45, y: y + 0.13, w: 3.7, h: 0.28, fontSize: 11, fontFace: "Arial Black", color: C.red, margin: 0 });
        s8.addText(bullets(item.items, { fontSize: 9 }), { x: x + 0.15, y: y + 0.48, w: 4.0, h: 1.05 });
    });

    // ═══════════════════════════════════════
    // SLIDE 9: THE NEXT CHAPTER — Recs
    // ═══════════════════════════════════════
    let s9 = pres.addSlide();
    s9.background = { color: C.dark };
    s9.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    addHeader(s9, "08", "What Tesla Should Do Next");
    addHook(s9, "The playbook that got Tesla here won't be enough for what's coming. Here's how to stay ahead.");

    // Left
    s9.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.75, w: 4.3, h: 3.6, fill: { color: C.darkCard }, shadow: mkShadow() });
    s9.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.75, w: 0.06, h: 3.6, fill: { color: C.red } });
    s9.addText("Fix the Gaps (4Ps + 7Ps)", { x: 0.85, y: 1.85, w: 3.8, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.red, bold: true, margin: 0 });
    s9.addText(bullets([
        "Product: Launch a sub-$30K model — 60% of buyers are priced out",
        "Price: Dynamic pricing based on demand (like airlines, not cars)",
        "Place: 40% more service centers — current 2-3 week wait is unacceptable",
        "Promotion: Targeted TikTok/YouTube for Gen Z — the next wave of buyers",
        "People: Service training overhaul — ACSI score is below industry average",
        "Process: Mobile service fleet should handle 80% of repairs at home",
        "Physical: VR test drives in showrooms — let people experience the future",
    ], { fontSize: 9 }), { x: 0.85, y: 2.25, w: 3.85, h: 2.9 });

    // Right
    s9.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.75, w: 4.4, h: 3.6, fill: { color: C.darkCard }, shadow: mkShadow() });
    s9.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.75, w: 0.06, h: 3.6, fill: { color: "0D47A1" } });
    s9.addText("Expand the Story (STP)", { x: 5.45, y: 1.85, w: 3.8, h: 0.3, fontSize: 12, fontFace: "Arial", color: "5C9CE6", bold: true, margin: 0 });
    s9.addText(bullets([
        "Go broader: mid-income ($50-75K) is a massive untapped segment",
        "Go global: strengthen positioning vs BYD and Chinese EVs before they dominate",
        "Go corporate: fleet/B2B for companies with ESG mandates",
        "Go inclusive: only 28% of buyers are women — huge opportunity",
        "Go younger: all-inclusive subscription model for Gen Z/Millennials",
        "Go local: tailor messaging for rural (range) vs urban (charging) markets",
        "Reposition: 'best value in class' not 'cheapest Tesla' to protect brand",
    ], { fontSize: 9 }), { x: 5.45, y: 2.25, w: 3.95, h: 2.9 });

    // ═══════════════════════════════════════
    // SLIDE 10: THE TAKEAWAY — Closing
    // ═══════════════════════════════════════
    let s10 = pres.addSlide();
    s10.background = { color: C.dark };
    s10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.03, fill: { color: C.red } });
    s10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.595, w: 10, h: 0.03, fill: { color: C.red } });
    s10.addShape(pres.shapes.OVAL, { x: 2.5, y: 1.5, w: 5, h: 3.5, fill: { color: C.redDim, transparency: 75 } });

    s10.addText("The Bottom Line", {
        x: 0.5, y: 0.8, w: 9, h: 0.6, fontSize: 32, fontFace: "Arial Black", color: C.white, align: "center", margin: 0
    });

    s10.addText("Tesla didn't just build a better car.\nThey built a better way to sell one.", {
        x: 1, y: 1.5, w: 8, h: 0.8, fontSize: 16, fontFace: "Georgia", color: C.dimText, align: "center", italic: true
    });

    // Key takeaways as story beats
    const takeaways = [
        "Zero advertising + 52% market share = the product IS the marketing",
        "Direct-to-consumer eliminated $2,000+ dealer markup AND controlled the brand",
        "91% satisfaction turned buyers into the most powerful salesforce in automotive",
        "But the moat is eroding — BYD, Ford, and GM are closing fast",
        "The next chapter requires going broader, cheaper, and more inclusive",
    ];
    takeaways.forEach((t, i) => {
        s10.addText([
            { text: String(i + 1).padStart(2, '0') + "  ", options: { fontSize: 11, color: C.red, bold: true } },
            { text: t, options: { fontSize: 11, color: C.gray } },
        ], { x: 1.5, y: 2.5 + i * 0.42, w: 7, h: 0.38, margin: 0 });
    });

    // References
    s10.addText("Sources: Statista (2025), Tesla 10-K (2024), Porter (1979), Kotler & Keller (2016), IEA (2025), Bloomberg NEF (2025), McKinsey (2024), JD Power (2024), Consumer Reports (2024)", {
        x: 0.6, y: 4.85, w: 8.8, h: 0.4, fontSize: 7, fontFace: "Arial", color: C.dimText, margin: 0
    });

    s10.addText("Thank You", {
        x: 0, y: 5.1, w: 10, h: 0.35, fontSize: 16, fontFace: "Arial Black", color: C.red, align: "center"
    });

    // Save
    await pres.writeFile({ fileName: "C:/laragon/www/tesla/Tesla_Model3_Marketing.pptx" });
    console.log("Presentation saved!");
}

build().catch(err => { console.error(err); process.exit(1); });
