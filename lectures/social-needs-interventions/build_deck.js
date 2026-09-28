// Builds "Social Needs Interventions" (10-slide MPH lecture) with pptxgenjs.
// Requires: pptxgenjs, react, react-dom, react-icons, sharp.  Run: node build_deck.js [out.pptx]
const pptxgen = require("pptxgenjs");
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const OUT = process.argv[2] || "Social_Needs_Interventions_MPH_Lecture.pptx";

const C = {
  deep: "123C3A", teal: "1F6F66", teal2: "2E7D72", sage: "EAF3F0", sageLine: "CFE2DC",
  accent: "B84A22", accentSoft: "F4A27F", accentTint: "FBECE5",
  ink: "1D2B2A", muted: "5C6B69", white: "FFFFFF", mint: "CFE3DE", deep2: "1C4F4B",
};
const HEAD = "Cambria", BODY = "Calibri";

const URL = {
  pcoriMap: "https://www.pcori.org/sites/default/files/social-needs-interventions/index.html",
  pcoriBlog: "https://www.pcori.org/blog/2022/exploring-evidence-social-needs-interventions-improve-health-outcomes",
  pcoriReport: "https://sirenetwork.ucsf.edu/tools-resources/resources/pcori-evidence-map-social-needs-interventions-improve-health-outcomes-0",
  pcoriUpdate: "https://sirenetwork.ucsf.edu/tools-resources/resources/pcori-evidence-map-social-needs-interventions-improve-health-outcomes-1",
  cg: "https://www.thecommunityguide.org/findings/social-determinants-health-tenant-based-housing-voucher-programs.html",
  cgTffrs: "https://www.thecommunityguide.org/pages/tffrs-social-determinants-health-tenant-based-housing-voucher-programs.html",
  cpstf: "https://www.thecommunityguide.org/pages/about-community-preventive-services-task-force.html",
  jphmp: "https://pubmed.ncbi.nlm.nih.gov/36194822/",
  ha: "https://www.healthaffairs.org/do/10.1377/hblog20190115.234942/full/",
  hp2030: "https://odphp.health.gov/healthypeople/priority-areas/social-determinants-health",
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.title = "Social Needs Interventions: What the Evidence Shows";

// ---------- helpers ----------
const iconCache = {};
async function icon(name, color) {
  const key = name + color;
  if (!iconCache[key]) {
    if (!fa[name]) throw new Error("Missing icon " + name);
    let svg = renderToStaticMarkup(React.createElement(fa[name], { size: 512 }));
    svg = svg.replace(/currentColor/g, "#" + color);
    if (!svg.includes("xmlns=")) svg = svg.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
    const buf = await sharp(Buffer.from(svg)).png().toBuffer();
    iconCache[key] = "image/png;base64," + buf.toString("base64");
  }
  return iconCache[key];
}

async function iconCircle(slide, x, y, d, fill, name, iconColor = C.white, ratio = 0.5) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0.5 } });
  const s = d * ratio;
  slide.addImage({ data: await icon(name, iconColor), x: x + (d - s) / 2, y: y + (d - s) / 2, w: s, h: s });
}

function text(slide, t, o) {
  slide.addText(t, { fontFace: BODY, color: C.ink, margin: 0, valign: "top", isTextBox: true, ...o });
}

function card(slide, x, y, w, h, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill: { color: fill }, line: { color: fill, width: 0.5 } });
}

function header(slide, kicker, title, dark = false) {
  text(slide, kicker.toUpperCase(), { x: 0.6, y: 0.32, w: 12.1, h: 0.3, fontSize: 12, bold: true, charSpacing: 2, color: dark ? C.accentSoft : C.accent });
  text(slide, title, { x: 0.6, y: 0.64, w: 12.1, h: 0.72, fontFace: HEAD, fontSize: 36, bold: true, valign: "middle", color: dark ? C.white : C.deep });
}

function footer(slide, src, n) {
  if (src) text(slide, src, { x: 0.6, y: 6.98, w: 11.3, h: 0.32, fontSize: 10, color: C.muted, valign: "middle" });
  text(slide, String(n), { x: 12.13, y: 6.98, w: 0.6, h: 0.32, fontSize: 10, color: C.muted, align: "right", valign: "middle" });
}

// ---------- build ----------
(async () => {
  // 1 — Title
  {
    const s = pres.addSlide();
    s.background = { color: C.deep };
    text(s, "MPH LECTURE  ·  SOCIAL DETERMINANTS OF HEALTH", { x: 0.8, y: 1.45, w: 7.4, h: 0.4, fontSize: 14, bold: true, charSpacing: 3, color: C.accentSoft });
    text(s, "Social Needs Interventions", { x: 0.8, y: 1.95, w: 7.4, h: 1.8, fontFace: HEAD, fontSize: 50, bold: true, color: C.white });
    text(s, "What the evidence shows: from PCORI's evidence map to CPSTF-recommended housing vouchers", { x: 0.8, y: 3.95, w: 7.0, h: 1.0, fontSize: 20, color: C.mint });
    text(s, [
      { text: "Evidence sources", options: { bold: true, color: C.accentSoft, breakLine: true } },
      { text: "PCORI: Social Needs Interventions to Improve Health Outcomes (evidence map)", options: { breakLine: true } },
      { text: "The Community Guide (CPSTF): Tenant-Based Housing Voucher Programs" },
    ], { x: 0.8, y: 5.55, w: 7.4, h: 1.0, fontSize: 13, color: C.mint, paraSpaceAfter: 2 });
    const circles = [
      [9.05, 1.35, C.accent, "FaHome"], [10.6, 1.35, C.teal2, "FaUtensils"],
      [8.275, 2.75, C.teal2, "FaBus"], [9.825, 2.75, C.mint, "FaStethoscope", C.deep], [11.375, 2.75, C.accent, "FaUserFriends"],
      [9.05, 4.15, C.teal2, "FaKey"], [10.6, 4.15, C.accent, "FaGraduationCap"],
    ];
    for (const [x, y, f, n, ic] of circles) await iconCircle(s, x, y, 1.3, f, n, ic || C.white);
    s.addNotes(
      "Welcome. This session covers (1) what social needs interventions are and how they differ from the social determinants of health, " +
      "(2) what PCORI's national evidence map shows about the research base, and (3) a CPSTF-recommended example: tenant-based housing voucher programs.\n\n" +
      "Primary sources:\n- PCORI. Social Needs Interventions to Improve Health Outcomes (interactive evidence map): " + URL.pcoriMap +
      "\n- Community Preventive Services Task Force (CPSTF). Social Determinants of Health: Tenant-Based Housing Voucher Programs. The Community Guide: " + URL.cg
    );
  }

  // 2 — Learning objectives
  {
    const s = pres.addSlide();
    header(s, "Session overview", "Learning objectives");
    const objs = [
      "Distinguish social needs from the social determinants of health (SDOH).",
      "Describe what PCORI's evidence map shows about social needs interventions, and where the gaps are.",
      "Summarize the CPSTF recommendation, effects, and economics of tenant-based housing voucher programs.",
      "Compare an evidence map with a systematic review–based recommendation, and apply both to program decisions.",
    ];
    for (let i = 0; i < objs.length; i++) {
      const y = 1.75 + i * 1.18;
      s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.12, w: 0.62, h: 0.62, fill: { color: C.teal }, line: { color: C.teal, width: 0.5 } });
      text(s, String(i + 1), { x: 0.6, y: y + 0.12, w: 0.62, h: 0.62, fontSize: 20, bold: true, color: C.white, align: "center", valign: "middle" });
      text(s, objs[i], { x: 1.45, y, w: 6.55, h: 0.86, fontSize: 17, valign: "middle" });
    }
    card(s, 8.5, 1.75, 4.23, 4.5, C.sage);
    text(s, "Roadmap", { x: 8.85, y: 2.0, w: 3.6, h: 0.45, fontFace: HEAD, fontSize: 20, bold: true, color: C.deep });
    const road = [
      ["FaLayerGroup", "Concepts: needs vs. determinants"],
      ["FaMapMarkedAlt", "PCORI evidence map: scope & gaps"],
      ["FaHome", "Case: housing vouchers (CPSTF)"],
      ["FaComments", "Synthesis & discussion"],
    ];
    for (let i = 0; i < road.length; i++) {
      const y = 2.7 + i * 0.85;
      await iconCircle(s, 8.85, y, 0.55, i === 2 ? C.accent : C.teal, road[i][0]);
      text(s, road[i][1], { x: 9.6, y, w: 2.95, h: 0.55, fontSize: 15, valign: "middle" });
    }
    footer(s, "", 2);
    s.addNotes(
      "Suggested pacing for a ~50-minute session: concepts (~5 min), PCORI evidence map (~15 min), housing voucher case (~15 min), synthesis and discussion (~15 min).\n\n" +
      "Objective 4 is the methods thread: an evidence map (scoping review) describes the size and shape of a literature, whereas a systematic review with a Task Force recommendation judges whether an intervention works and is worth its cost."
    );
  }

  // 3 — Concepts
  {
    const s = pres.addSlide();
    header(s, "Part 1 · Concepts", "Social needs vs. social determinants");
    card(s, 0.6, 1.65, 5.95, 2.4, C.sage);
    await iconCircle(s, 0.85, 1.88, 0.62, C.teal, "FaCity");
    text(s, "Social determinants of health (SDOH)", { x: 1.65, y: 1.88, w: 4.75, h: 0.62, fontSize: 17, bold: true, color: C.deep, valign: "middle" });
    text(s, [
      { text: "“The conditions in the environments where people are born, live, learn, work, play, worship, and age that affect a wide range of health, functioning, and quality-of-life outcomes and risks.”", options: { breakLine: true } },
      { text: "— Healthy People 2030", options: { italic: true, color: C.muted } },
    ], { x: 0.85, y: 2.65, w: 5.5, h: 1.3, fontSize: 13.5, paraSpaceAfter: 4 });

    card(s, 0.6, 4.25, 5.95, 2.4, C.accentTint);
    await iconCircle(s, 0.85, 4.48, 0.62, C.accent, "FaUser");
    text(s, "Social needs", { x: 1.65, y: 4.48, w: 4.75, h: 0.62, fontSize: 17, bold: true, color: C.deep, valign: "middle" });
    text(s, [
      { text: "Adverse social conditions associated with poor health that are identified or prioritized from patients' perspectives.", options: { breakLine: true } },
      { text: "— PCORI evidence map", options: { italic: true, color: C.muted, breakLine: true } },
      { text: "Example: one family's food insecurity or housing instability.", options: { color: C.muted } },
    ], { x: 0.85, y: 5.25, w: 5.5, h: 1.3, fontSize: 13.5, paraSpaceAfter: 4 });

    text(s, "Social needs addressed in the PCORI map", { x: 6.95, y: 1.65, w: 5.8, h: 0.4, fontSize: 17, bold: true, color: C.deep });
    const doms = [
      ["FaUtensils", "Food insecurity"], ["FaHome", "Housing instability & quality"],
      ["FaBus", "Transportation"], ["FaShieldAlt", "Interpersonal violence"],
      ["FaWallet", "Financial strain"], ["FaBriefcase", "Employment"],
      ["FaGraduationCap", "Education"], ["FaUserFriends", "Social isolation"],
      ["FaBalanceScale", "Access to legal services"], ["FaBaby", "Childcare"],
      ["FaShapes", "Early childhood education & development"], ["FaStethoscope", "Health care & primary care"],
    ];
    for (let i = 0; i < doms.length; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 6.95 + col * 2.95, y = 2.15 + row * 0.55;
      await iconCircle(s, x, y, 0.44, row % 2 === col % 2 ? C.teal : C.teal2, doms[i][0]);
      text(s, doms[i][1], { x: x + 0.55, y: y - 0.04, w: 2.3, h: 0.52, fontSize: 12.5, valign: "middle" });
    }
    card(s, 6.95, 5.55, 5.78, 1.1, C.deep);
    text(s, [
      { text: "Key distinction: ", options: { bold: true, color: C.accentSoft } },
      { text: "meeting individual social needs falls short of addressing the social determinants of health (Castrucci & Auerbach, 2019).", options: { color: C.white } },
    ], { x: 7.2, y: 5.62, w: 5.3, h: 0.96, fontSize: 14, valign: "middle" });
    footer(s, "Sources: Healthy People 2030 (ODPHP); PCORI Social Needs Interventions evidence map; Castrucci B, Auerbach J. Health Affairs Blog, Jan 16, 2019.", 3);
    s.addNotes(
      "SDOH (Healthy People 2030): \"the conditions in the environments where people are born, live, learn, work, play, worship, and age that affect a wide range of health, functioning, and quality-of-life outcomes and risks.\" " + URL.hp2030 + "\n\n" +
      "Social needs (PCORI evidence map): adverse social conditions that are associated with poor health and are identified or prioritized based on patients' perspectives. The 12 domains on the right are the social needs PCORI lists for the map: food insecurity, transportation, housing instability and quality, interpersonal violence, education, financial strain, employment, and social isolation, plus access to legal services, childcare, early childhood education and development, and health care and primary care. " + URL.pcoriMap + "\n\n" +
      "Castrucci & Auerbach (Health Affairs Blog, January 16, 2019) warn that the terms are often used interchangeably and argue that meeting individual social needs falls short of addressing SDOH. " + URL.ha + "\n\n" +
      "Prompt: Which of these 12 domains could a clinic address for one patient this week? Which would require policy or systems change?"
    );
  }

  // 4 — PCORI map: what it is
  {
    const s = pres.addSlide();
    header(s, "Part 2 · Mapping the evidence (PCORI)", "A living map of the evidence");
    const rows = [
      ["FaSearch", "What it is", "An interactive visualization of a systematic scoping review of social needs interventions that reported behavioral, health, or health care utilization outcomes."],
      ["FaHandshake", "Who built it", "Commissioned by PCORI; designed with social needs experts from SIREN (Social Interventions Research and Evaluation Network)."],
      ["FaFilter", "How to use it", "Filter by social need, population, intervention setting, study design and quality, and outcome type and direction; open study details or download sets of studies."],
    ];
    for (let i = 0; i < rows.length; i++) {
      const y = 1.7 + i * 1.62;
      await iconCircle(s, 0.6, y, 0.66, i === 0 ? C.accent : C.teal, rows[i][0]);
      text(s, rows[i][1], { x: 1.5, y: y - 0.02, w: 4.85, h: 0.4, fontSize: 16, bold: true, color: C.deep });
      text(s, rows[i][2], { x: 1.5, y: y + 0.4, w: 4.85, h: 1.05, fontSize: 13.5 });
    }
    s.addChart(pres.charts.BAR, [{ name: "Studies", labels: ["Mar 2021", "Nov 2021", "Apr 2023"], values: [139, 157, 280] }], {
      x: 6.75, y: 1.6, w: 5.98, h: 4.15, barDir: "col", chartColors: [C.teal], barGapWidthPct: 70,
      showTitle: true, title: "Studies in the map, by search cutoff", titleFontFace: BODY, titleFontSize: 15, titleColor: C.deep,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 16, dataLabelFontBold: true, dataLabelFontFace: BODY, dataLabelColor: C.ink,
      catAxisLabelColor: C.muted, catAxisLabelFontSize: 13, catAxisLabelFontFace: BODY, catAxisLineColor: C.sageLine,
      valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 320, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      showLegend: false,
    });
    text(s, "Search cutoffs: Mar 5, 2021 (first map) · Nov 29, 2021 (Feb 2022 update: 157 studies, 178 publications) · Apr 6, 2023. PCORI also reports 285 studies in the current visualization — always cite the version and search date.",
      { x: 6.95, y: 5.85, w: 5.78, h: 0.85, fontSize: 11.5, italic: true, color: C.muted });
    footer(s, "Sources: PCORI Social Needs Interventions evidence map and program page; Viswanathan M, et al. PCORI scoping review and evidence map update (Feb 2022).", 4);
    s.addNotes(
      "PCORI commissioned this interactive evidence map. The visualization and its reports summarize a systematic scoping review of social needs interventions that reported behavioral, health, or health care utilization outcomes. It was designed with social needs experts from the Social Interventions Research and Evaluation Network (SIREN) and complements SIREN's Evidence & Resource Library.\n\n" +
      "Versions (why the count changes): 139 studies from searches through March 5, 2021 (first published August 2021); an update with searches through November 29, 2021 added 19 studies (23 publications) for a cumulative 157 studies (178 publications), published February 2022; the map later displayed 280 studies from searches through April 6, 2023. PCORI also states that 285 studies are represented in the visualization. Check the live map and record the date you accessed it.\n\n" +
      "Live demo: open " + URL.pcoriMap + ", filter to one social need (e.g., housing), then by setting, open a study record, and show the spreadsheet download."
    );
  }

  // 5 — PCORI: what the map shows
  {
    const s = pres.addSlide();
    header(s, "Part 2 · Mapping the evidence (PCORI)", "What the map shows");
    text(s, "Snapshot from the February 2022 update: N = 157 studies (searches through Nov 29, 2021)", { x: 0.6, y: 1.42, w: 6.45, h: 0.35, fontSize: 13, italic: true, color: C.muted });
    const stats = [
      ["113", "studies addressed health care services access & quality — the most frequently studied need", C.accent],
      ["95", "studies addressed housing stability & quality — the next most frequent", C.accent],
      ["Nearly ¾", "of studies addressed more than one social need", C.deep],
      ["124", "studies (more than three-quarters) focused on adults aged ≥ 18 years", C.deep],
    ];
    for (let i = 0; i < stats.length; i++) {
      const x = 0.6 + (i % 2) * 3.3, y = 1.95 + Math.floor(i / 2) * 2.35;
      card(s, x, y, 3.05, 2.15, C.sage);
      text(s, stats[i][0], { x: x + 0.25, y: y + 0.18, w: 2.6, h: 0.8, fontFace: HEAD, fontSize: i === 2 ? 34 : 44, bold: true, color: stats[i][2], valign: "middle" });
      text(s, stats[i][1], { x: x + 0.25, y: y + 1.02, w: 2.6, h: 1.0, fontSize: 13 });
    }
    card(s, 7.35, 1.95, 5.38, 4.5, C.deep);
    const rows = [
      ["FaHospital", "Outcomes measured", "Health care utilization outcomes were reported most often, e.g., emergency department visits and inpatient admissions."],
      ["FaChartBar", "Direction of results", "Most studies reported positive, mixed, or no effects; a minority reported negative outcomes (results favoring the comparison arm)."],
      ["FaLightbulb", "Read with care", "A count of study directions is not a pooled effect size. Weigh study design and quality before concluding an intervention works."],
    ];
    for (let i = 0; i < rows.length; i++) {
      const y = 2.2 + i * 1.38;
      await iconCircle(s, 7.6, y, 0.55, C.accent, rows[i][0]);
      text(s, rows[i][1], { x: 8.35, y: y - 0.02, w: 4.15, h: 0.35, fontSize: 15, bold: true, color: C.accentSoft });
      text(s, rows[i][2], { x: 8.35, y: y + 0.36, w: 4.15, h: 0.9, fontSize: 13, color: C.white });
    }
    footer(s, "Sources: PCORI blog, “Exploring the Evidence on Social Needs Interventions to Improve Health Outcomes” (2022); PCORI evidence map; Viswanathan M, et al. PCORI update (2022).", 5);
    s.addNotes(
      "Counts on the left come from the February 2022 update (N = 157 studies; searches through November 29, 2021), as reported in PCORI's 2022 blog and the update report: health care services access and quality was the most frequently studied social need (113 studies), followed by housing stability and quality (95 studies); nearly three-fourths of all studies addressed more than one social need; more than three-quarters of studies (n = 124) focused on adults aged 18 or older. " + URL.pcoriBlog + "\n\n" +
      "Across the map, health care utilization outcomes were most commonly reported and included resource-intensive events such as emergency department visits and inpatient admissions. Most studies reported positive results (consistent with the intended direction of effect), mixed results (inconsistent across multiple similar outcomes), or no effects; a minority reported negative outcomes (favoring the comparison arm).\n\n" +
      "Methods point for MPH students: tallying positive vs. null studies (vote counting) ignores effect size, precision, and study quality, so it cannot tell us how well an intervention works."
    );
  }

  // 6 — PCORI: gaps
  {
    const s = pres.addSlide();
    header(s, "Part 2 · Mapping the evidence (PCORI)", "Where the evidence is thin");
    const cols = [
      { x: 0.6, icon: "FaUserFriends", title: "Under-addressed needs" },
      { x: 4.74, icon: "FaUsers", title: "Under-studied populations" },
      { x: 8.88, icon: "FaBullseye", title: "Outcomes that matter" },
    ];
    for (const c of cols) {
      card(s, c.x, 1.65, 3.85, 3.95, C.sage);
      await iconCircle(s, c.x + 0.28, 1.9, 0.66, C.teal, c.icon);
      text(s, c.title, { x: c.x + 1.08, y: 1.9, w: 2.65, h: 0.66, fontSize: 17, bold: true, color: C.deep, valign: "middle" });
    }
    ["Social isolation", "Childcare assistance", "Financial strain"].forEach((t, i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.88, y: 2.8 + i * 0.58, w: 3.29, h: 0.46, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.sageLine, width: 1 } });
      text(s, t, { x: 1.05, y: 2.8 + i * 0.58, w: 3.0, h: 0.46, fontSize: 14, bold: true, color: C.accent, valign: "middle" });
    });
    text(s, "PCORI flags these as needing heightened attention, particularly in the context of COVID-19.", { x: 0.88, y: 4.62, w: 3.29, h: 1.2, fontSize: 13 });
    text(s, [
      { text: "Unclear whether effects differ for racially and ethnically diverse populations, or whether interventions help:", options: { breakLine: true } },
      ...["adolescents", "children", "older adults", "pregnant persons", "immigrants", "sexual and gender minority communities"].map((t, i, a) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < a.length - 1 } })),
    ], { x: 5.02, y: 2.8, w: 3.35, h: 2.7, fontSize: 13, paraSpaceAfter: 3 });
    text(s, [
      { text: "Health care utilization (e.g., ED visits, inpatient admissions) is the most commonly reported outcome.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 6 } },
      { text: "Ask of any study: are patient-centered health outcomes and effects on disparities measured?", options: { bold: true, color: C.deep } },
    ], { x: 9.16, y: 2.8, w: 3.3, h: 2.7, fontSize: 14 });
    card(s, 0.6, 5.8, 12.13, 0.85, C.deep);
    await iconCircle(s, 0.85, 5.95, 0.55, C.accent, "FaLightbulb");
    text(s, [
      { text: "For practice: ", options: { bold: true, color: C.accentSoft } },
      { text: "search the map before designing or funding a program, and design evaluations that fill these gaps.", options: { color: C.white } },
    ], { x: 1.6, y: 5.8, w: 10.9, h: 0.85, fontSize: 15, valign: "middle" });
    footer(s, "Sources: PCORI blog (2022); Viswanathan M, et al. Social Needs Interventions to Improve Health Outcomes: Review and Evidence Map (PCORI, 2021); PCORI evidence map.", 6);
    s.addNotes(
      "Under-addressed needs: PCORI's 2022 blog notes that less-well-addressed needs such as social isolation, childcare assistance, and financial strain will require heightened attention, particularly in the context of COVID-19. " + URL.pcoriBlog + "\n\n" +
      "Under-studied populations: the PCORI review reports that evidence gaps persist in understanding whether these interventions have differential effects on racially and ethnically diverse populations and whether they are impactful for adolescents, children, older adults, pregnant persons, immigrants, and sexual and gender minority communities. " + URL.pcoriReport + "\n\n" +
      "Outcomes: health care utilization outcomes (e.g., ED visits, inpatient admissions) were the most commonly reported. The question on the slide is a critical-appraisal prompt, not a PCORI finding."
    );
  }

  // 7 — Vouchers: how they work
  {
    const s = pres.addSlide();
    header(s, "Part 3 · Case study (The Community Guide)", "Case: tenant-based housing vouchers");
    text(s, "Vouchers help households with very low incomes afford safe and sanitary housing in the private market. They are tied to households, not specific units, so families can use them to move to neighborhoods with greater opportunities.",
      { x: 0.6, y: 1.45, w: 12.1, h: 0.72, fontSize: 15 });
    const steps = [
      ["FaUsers", "Eligible household", "A household with very low income (e.g., HUD's Housing Choice Voucher program, “Section 8”)."],
      ["FaTicketAlt", "Receives a voucher", "Rent is limited to 30% or less of household income."],
      ["FaSearch", "Searches the market", "Within a time limit, finds housing that passes HUD certification and a landlord who accepts vouchers."],
      ["FaMapMarkedAlt", "Moves", "Can choose a neighborhood with greater opportunity (e.g., lower poverty, better schools)."],
    ];
    for (let i = 0; i < steps.length; i++) {
      const x = 0.6 + i * 3.12, y = 2.35;
      card(s, x, y, 2.75, 2.05, C.sage);
      await iconCircle(s, x + 0.2, y + 0.2, 0.58, i === 3 ? C.accent : C.teal, steps[i][0]);
      text(s, "STEP " + (i + 1), { x: x + 0.9, y: y + 0.2, w: 1.7, h: 0.25, fontSize: 11, bold: true, charSpacing: 1, color: C.accent });
      text(s, steps[i][1], { x: x + 0.9, y: y + 0.44, w: 1.8, h: 0.34, fontSize: 14.5, bold: true, color: C.deep });
      text(s, steps[i][2], { x: x + 0.2, y: y + 0.92, w: 2.4, h: 1.05, fontSize: 12.5 });
      if (i < steps.length - 1) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 2.81, y: y + 0.87, w: 0.26, h: 0.3, fill: { color: C.teal2 }, line: { color: C.teal2, width: 0.5 } });
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 4.62, w: 5.9, h: 2.05, rectRadius: 0.1, fill: { color: C.white }, line: { color: C.sageLine, width: 1 } });
    text(s, "Programs vary in…", { x: 0.85, y: 4.75, w: 5.4, h: 0.38, fontSize: 15, bold: true, color: C.deep });
    text(s, [
      "Eligibility (e.g., family income level)",
      "Rental process (e.g., time allowed to find and rent a unit)",
      "Assistance (e.g., counseling in finding rentals)",
      "Relocation rules (e.g., housing in low-poverty neighborhoods)",
      "Short-term payments for initial costs (e.g., rental deposits)",
    ].map((t, i, a) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < a.length - 1 } })), { x: 0.85, y: 5.15, w: 5.45, h: 1.45, fontSize: 12.5, paraSpaceAfter: 1 });
    card(s, 6.83, 4.62, 5.9, 2.05, C.accentTint);
    await iconCircle(s, 7.08, 4.8, 0.55, C.accent, "FaRoute");
    text(s, "MTO-type programs (Moving to Opportunity)", { x: 7.8, y: 4.8, w: 4.75, h: 0.55, fontSize: 15, bold: true, color: C.deep, valign: "middle" });
    text(s, "Vouchers plus pre-move counseling for families with young children living in public housing; households move to lower-poverty, higher-opportunity neighborhoods. CPSTF's favorable economic finding applies to this model (slide 9).",
      { x: 7.08, y: 5.45, w: 5.45, h: 1.15, fontSize: 12.5 });
    footer(s, "Sources: The Community Guide, SDOH: Tenant-Based Housing Voucher Programs (finding and TFFRS); J Public Health Manag Pract. 2022;28(6):E795–E803.", 7);
    s.addNotes(
      "Definition (The Community Guide): tenant-based housing voucher programs help households with very low incomes afford safe and sanitary housing in the private market. Vouchers are tied to households rather than specific housing units, so households can use vouchers to move to neighborhoods with greater opportunities. HUD administers the Housing Choice Voucher program (also known as Section 8), a tenant-based voucher program. " + URL.cg + "\n\n" +
      "Voucher programs limit rent to 30% or less of household income (Community Guide systematic review, J Public Health Manag Pract 2022). " + URL.jphmp + "\n\n" +
      "Process barrier (Community Guide TFFRS): households have a limited time to find housing that passes HUD certification and a landlord who accepts vouchers. " + URL.cgTffrs + "\n\n" +
      "Programs vary in eligibility criteria, rental process, assistance, relocation requirements, and availability of short-term payments for initial expenses. MTO-type programs provide vouchers and pre-move counseling to households with young children living in public housing and require moves to neighborhoods with lower poverty and higher opportunity."
    );
  }

  // 8 — CPSTF finding
  {
    const s = pres.addSlide();
    header(s, "Part 3 · Case study (The Community Guide)", "CPSTF finding: Recommended");
    card(s, 0.6, 1.6, 4.95, 5.1, C.deep);
    await iconCircle(s, 0.9, 1.88, 0.8, C.accent, "FaCheck");
    text(s, "RECOMMENDED", { x: 1.9, y: 1.98, w: 3.4, h: 0.6, fontSize: 20, bold: true, charSpacing: 2, color: C.accentSoft, valign: "middle" });
    text(s, "CPSTF recommends tenant-based housing voucher programs to improve health and health-related outcomes for adults (sufficient evidence of effectiveness) and to advance health equity.",
      { x: 0.9, y: 2.9, w: 4.4, h: 1.75, fontSize: 16, color: C.white });
    text(s, "Evidence base: 7 studies (20 publications) of U.S. voucher programs with concurrent comparison groups; literature from 1999 to July 2019.",
      { x: 0.9, y: 4.7, w: 4.4, h: 0.95, fontSize: 12.5, color: C.mint });
    text(s, "CPSTF: an independent, nonfederal panel of public health and prevention experts appointed by the CDC Director.",
      { x: 0.9, y: 5.75, w: 4.4, h: 0.8, fontSize: 12, italic: true, color: C.mint });
    const res = [
      ["FaUser", "Adults / households", "Better housing quality and security, improved health care access, greater neighborhood opportunity, and fewer physical and mental health conditions.", C.teal],
      ["FaChild", "Children aged 12 and younger", "Improvements in education, employment, and income later in life.", C.teal],
      ["FaExclamationTriangle", "Adolescents aged 10–20 (lower-poverty moves)", "Females: better health outcomes. Males: worse physical and mental health outcomes; more research is needed to understand and address their challenges.", C.accent],
    ];
    for (let i = 0; i < res.length; i++) {
      const y = 1.6 + i * 1.75;
      card(s, 5.85, y, 6.88, 1.6, i === 2 ? C.accentTint : C.sage);
      await iconCircle(s, 6.1, y + 0.45, 0.7, res[i][3], res[i][0]);
      text(s, res[i][1], { x: 7.05, y: y + 0.17, w: 5.5, h: 0.38, fontSize: 16, bold: true, color: C.deep });
      text(s, res[i][2], { x: 7.05, y: y + 0.58, w: 5.5, h: 0.95, fontSize: 13.5 });
    }
    footer(s, "Sources: The Community Guide, SDOH: Tenant-Based Housing Voucher Programs (CPSTF finding); J Public Health Manag Pract. 2022;28(6):E795–E803; About the CPSTF.", 8);
    s.addNotes(
      "CPSTF finding: the Task Force recommends tenant-based housing voucher programs to improve health and health-related outcomes for adults, based on sufficient evidence of effectiveness, and to advance health equity. Health-related outcomes include housing quality and security, health care use, and neighborhood opportunities (e.g., lower poverty level, better schools). " + URL.cg + "\n\n" +
      "Evidence base (J Public Health Manag Pract 2022;28(6):E795–E803): 7 studies in 20 publications; U.S. voucher programs with concurrent comparison populations; literature from 1999 to July 2019. Households that used vouchers reported greater housing quality and security and fewer physical and mental health conditions than eligible households not offered vouchers. " + URL.jphmp + "\n\n" +
      "Children aged 12 years and younger whose households used vouchers showed improvements in education, employment, and income later in life. Females aged 10–20 whose families used vouchers to live in lower-poverty neighborhoods experienced better health outcomes, whereas males of the same age experienced worse physical and mental health outcomes; the Community Guide calls for more research to understand and address challenges faced by adolescent males.\n\n" +
      "About the CPSTF: an independent, nonfederal panel of public health and prevention experts whose members are appointed by the CDC Director; its findings are based on systematic reviews of effectiveness and economic evidence. " + URL.cpstf
    );
  }

  // 9 — Economics & implementation
  {
    const s = pres.addSlide();
    header(s, "Part 3 · Case study (The Community Guide)", "Worth the cost? Making it work");
    text(s, [
      { text: "CPSTF economic finding: ", options: { bold: true, color: C.deep } },
      { text: "societal benefits exceed the cost of programs that serve families with young children living in public housing, provide pre-move counseling, and move families to neighborhoods with greater opportunities." },
    ], { x: 0.6, y: 1.6, w: 5.95, h: 1.35, fontSize: 15 });
    const eco = [["$69,601", "lifetime societal economic benefit per household", C.accent], ["$9,215", "cost savings for taxpayers per household", C.deep]];
    eco.forEach(([n, l, col], i) => {
      const x = 0.6 + i * 3.1;
      card(s, x, 3.1, 2.85, 1.95, C.sage);
      text(s, n, { x: x + 0.22, y: 3.25, w: 2.45, h: 0.75, fontFace: HEAD, fontSize: 34, bold: true, color: col, valign: "middle" });
      text(s, l, { x: x + 0.22, y: 4.05, w: 2.45, h: 0.9, fontSize: 13 });
    });
    text(s, "One study comparing MTO-type programs with public housing. The economic review included 27 U.S. studies (searches 1980 to Nov 2020).",
      { x: 0.6, y: 5.2, w: 5.95, h: 0.8, fontSize: 12, italic: true, color: C.muted });

    card(s, 6.85, 1.6, 5.88, 5.07, C.sage);
    text(s, "Implementation considerations", { x: 7.15, y: 1.8, w: 5.3, h: 0.42, fontSize: 17, bold: true, color: C.deep });
    const imp = [
      ["FaClock", "Allow more time to search for housing (more than 60 days)"],
      ["FaComments", "Offer intensive pre-move counseling"],
      ["FaMoneyBillWave", "Provide short-term financial help with initial moving expenses"],
      ["FaHandshake", "Recruit landlords to participate"],
      ["FaGavel", "Enact state or local source-of-income laws that bar discrimination against renters using vouchers"],
    ];
    for (let i = 0; i < imp.length; i++) {
      const y = 2.38 + i * 0.72;
      await iconCircle(s, 7.15, y, 0.48, i === 4 ? C.accent : C.teal, imp[i][0]);
      text(s, imp[i][1], { x: 7.8, y: y - 0.08, w: 4.7, h: 0.64, fontSize: 13, valign: "middle" });
    }
    text(s, "Barrier: many landlords hold negative perceptions of voucher holders or see voucher paperwork and inspections as too cumbersome.",
      { x: 7.15, y: 5.98, w: 5.35, h: 0.6, fontSize: 12, italic: true, color: C.muted });
    footer(s, "Source: The Community Guide, SDOH: Tenant-Based Housing Voucher Programs (economic evidence, finding, and TFFRS implementation considerations).", 9);
    s.addNotes(
      "Economic evidence (The Community Guide): the economic review included 27 U.S. studies (search period 1980 to November 7, 2020). For MTO-type programs compared with public housing, one study estimated cost savings for taxpayers of $9,215 per household and a societal economic benefit of $69,601 per household over the lifetime. CPSTF finds societal benefits exceed the cost of tenant-based housing voucher programs that serve families with young children who are living in public housing, provide pre-move counseling, and move families to neighborhoods with greater opportunities. " + URL.cg + "\n\n" +
      "Implementation considerations (Community Guide TFFRS): reduce barriers by giving households more time to search for and arrange housing (more than 60 days), offering intensive pre-move counseling, providing short-term financial assistance to cover initial moving expenses, or recruiting landlords; state and local source-of-income laws or ordinances can prohibit discrimination against renters based on the source of their income, including vouchers. Many landlords have negative perceptions of voucher holders and may find the paperwork and inspection process too cumbersome. " + URL.cgTffrs
    );
  }

  // 10 — Synthesis, discussion, sources
  {
    const s = pres.addSlide();
    header(s, "Synthesis & discussion", "Putting it together");
    text(s, "Two lenses on the evidence", { x: 0.6, y: 1.45, w: 5.95, h: 0.38, fontSize: 17, bold: true, color: C.deep });
    const lenses = [
      ["FaMapMarkedAlt", "Evidence map (PCORI)", "Shows what has been studied: needs, populations, settings, designs, and direction of results. Use it to find evidence and gaps.", C.teal],
      ["FaCheckCircle", "Systematic review + recommendation (CPSTF)", "Judges effectiveness and economic value to issue guidance, e.g., housing vouchers: Recommended.", C.accent],
    ];
    for (let i = 0; i < lenses.length; i++) {
      const y = 1.9 + i * 1.33;
      card(s, 0.6, y, 5.95, 1.2, C.sage);
      await iconCircle(s, 0.82, y + 0.3, 0.6, lenses[i][3], lenses[i][0]);
      text(s, lenses[i][1], { x: 1.62, y: y + 0.1, w: 4.8, h: 0.34, fontSize: 14.5, bold: true, color: C.deep });
      text(s, lenses[i][2], { x: 1.62, y: y + 0.45, w: 4.8, h: 0.72, fontSize: 12.5 });
    }
    text(s, "Discuss", { x: 6.85, y: 1.45, w: 5.88, h: 0.38, fontSize: 17, bold: true, color: C.deep });
    card(s, 6.85, 1.9, 5.88, 2.53, C.deep);
    const qs = [
      "A clinic screens for food and housing needs and refers patients to services. Is it addressing social needs, SDOH, or both?",
      "Adolescent males in voucher households fared worse. What supports would you add to a mobility program, and how would you evaluate them?",
      "You can fund one program: a well-studied need or an under-studied one (e.g., social isolation)? Use the map to justify your choice.",
    ];
    text(s, qs.map((q, i, a) => ({ text: q, options: { bullet: { type: "number", indent: 16 }, breakLine: i < a.length - 1 } })),
      { x: 7.08, y: 2.02, w: 5.45, h: 2.3, fontSize: 13, color: C.white, paraSpaceAfter: 5, valign: "middle" });
    const refs = [
      ["PCORI. Social Needs Interventions to Improve Health Outcomes [interactive evidence map]. ", "pcori.org/sites/default/files/social-needs-interventions/index.html", URL.pcoriMap],
      ["CPSTF. Social Determinants of Health: Tenant-Based Housing Voucher Programs. The Community Guide. ", "thecommunityguide.org/findings/social-determinants-health-tenant-based-housing-voucher-programs.html", URL.cg],
      ["Tenant-Based Housing Voucher Programs: A Community Guide Systematic Review. J Public Health Manag Pract. 2022;28(6):E795–E803. ", "PMID 36194822", URL.jphmp],
      ["Viswanathan M, et al. Social Needs Interventions to Improve Health Outcomes: Scoping Review and Evidence Map (2021; update 2022). PCORI. ", "Summary (SIREN)", URL.pcoriUpdate],
      ["PCORI. Exploring the Evidence on Social Needs Interventions to Improve Health Outcomes [blog]. 2022. ", "pcori.org/blog/2022/exploring-evidence-social-needs-interventions-improve-health-outcomes", URL.pcoriBlog],
      ["Castrucci B, Auerbach J. Meeting Individual Social Needs Falls Short of Addressing Social Determinants of Health. Health Affairs Blog. Jan 16, 2019. ", "doi:10.1377/hblog20190115.234942", URL.ha],
      ["ODPHP. Healthy People 2030: Social Determinants of Health. ", "odphp.health.gov/healthypeople/priority-areas/social-determinants-health", URL.hp2030],
    ];
    text(s, "Key sources", { x: 0.6, y: 4.62, w: 12.1, h: 0.3, fontSize: 12, bold: true, color: C.accent });
    text(s, refs.map(([t, lt, u], i, a) => [
      { text: (i + 1) + ". " + t },
      { text: lt, options: { hyperlink: { url: u }, breakLine: i < a.length - 1 } },
    ]).flat(), { x: 0.6, y: 4.95, w: 12.13, h: 1.9, fontSize: 10, color: C.ink, paraSpaceAfter: 2 });
    footer(s, "", 10);
    s.addNotes(
      "Synthesis: the PCORI map tells us where research exists (many studies on health care access and housing; few on social isolation, childcare, financial strain) and which way results point; the CPSTF recommendation tells us that one specific housing intervention works for adults, advances health equity, and, in its MTO-type form, has societal benefits that exceed costs.\n\n" +
      "Discussion guidance: Q1 builds on Castrucci & Auerbach (2019). Q2 draws on the Community Guide's call for research on adolescent males. Q3 asks students to weigh evidence volume against unmet need.\n\n" +
      "Full references:\n" +
      "1. PCORI. Social Needs Interventions to Improve Health Outcomes (interactive evidence map). " + URL.pcoriMap + "\n" +
      "2. PCORI. Exploring the Evidence on Social Needs Interventions to Improve Health Outcomes (blog), 2022. " + URL.pcoriBlog + "\n" +
      "3. Viswanathan M, Kennedy S, Eder M, et al. Social Needs Interventions to Improve Health Outcomes: Scoping Review and Evidence Map (2021) and update (2022). PCORI. " + URL.pcoriReport + " ; " + URL.pcoriUpdate + "\n" +
      "4. Community Preventive Services Task Force. Social Determinants of Health: Tenant-Based Housing Voucher Programs. The Community Guide. " + URL.cg + " ; TFFRS: " + URL.cgTffrs + "\n" +
      "5. Tenant-Based Housing Voucher Programs: A Community Guide Systematic Review. J Public Health Manag Pract. 2022;28(6):E795–E803. PMID 36194822. " + URL.jphmp + "\n" +
      "6. Castrucci B, Auerbach J. Meeting Individual Social Needs Falls Short of Addressing Social Determinants of Health. Health Affairs Blog. January 16, 2019. doi:10.1377/hblog20190115.234942\n" +
      "7. Office of Disease Prevention and Health Promotion. Social Determinants of Health. Healthy People 2030. " + URL.hp2030 + "\n" +
      "8. The Community Guide. About the Community Preventive Services Task Force. " + URL.cpstf
    );
  }

  await pres.writeFile({ fileName: OUT });
  console.log("Wrote " + OUT);
})().catch((e) => { console.error(e); process.exit(1); });
