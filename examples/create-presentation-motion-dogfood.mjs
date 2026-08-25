// Run with:
//   officekit run examples/create-presentation-motion-dogfood.mjs -- <output-directory>
//
// The three decks are real authoring examples, not an automated benchmark.
// They deliberately combine editable native content, visual carriers, charts,
// semantic motion recipes, speaker-note sources, round-trip inspection, and
// per-slide SVG evidence.

import { createHash } from "node:crypto";
import { mkdir, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { Presentation, PresentationFile } from "office-kit";

const outputRoot = path.resolve(process.argv[2] || path.join(os.tmpdir(), "officekit-motion-dogfood"));
const FONT = "PingFang SC";
const C = {
  bg: "#080A0F",
  panel: "#111722",
  panel2: "#171F2C",
  gold: "#D7A845",
  gold2: "#F4D06F",
  white: "#F7F2E8",
  muted: "#9CA3AF",
  red: "#FF6B6B",
  green: "#52D39A",
  blue: "#67B7F7",
  ink: "#111318",
};

const SOURCES = {
  treasury: "https://apnews.com/article/gold-bitcoin-treasury-dollar-bessent-inflation-trump-be7df8c0eaa159e4149df8efc4000fc9",
  etf: "https://news.bitcoin.com/bitcoin-etf/bitcoin-ether-etfs-pull-in-2-6b-in-strongest-week-since-october/",
  policy: "https://www.theblock.co/news/regulation/2026-08-20-inside-oval-office-meeting-trump-bullish-clarity-act-solicits-feedback-crypto-finance-ceos-412340",
  whale: "https://news.bitcoin.com/crypto-news/bitcoin-whale-sells-576-million-btc-nears-80k/",
};

function frame(left, top, width, height) {
  return { left, top, width, height };
}

function svgDataUrl(source) {
  return `data:image/svg+xml;base64,${Buffer.from(source, "utf8").toString("base64")}`;
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function parseNdjson(value) {
  return String(value || "").split("\n").filter(Boolean).map((line) => JSON.parse(line));
}

function text(slide, name, value, position, options = {}) {
  const textPosition = { ...position, height: Math.max(41, position.height) };
  return slide.shapes.add({
    name,
    geometry: "textbox",
    position: textPosition,
    fill: options.fill || "transparent",
    line: options.line || { fill: "transparent", width: 0 },
    text: value,
    textStyle: {
      fontFamily: options.fontFamily || FONT,
      fontSize: options.fontSize || 24,
      color: options.color || C.white,
      bold: options.bold,
      italic: options.italic,
    },
  });
}

function box(slide, name, position, options = {}) {
  return slide.shapes.add({
    name,
    geometry: options.geometry || "roundRect",
    position,
    fill: options.fill || C.panel,
    line: options.line || { fill: options.lineColor || C.gold, width: options.lineWidth ?? 1 },
    shadow: options.shadow === false ? undefined : { color: "#000000", blurRadius: 10, distance: 4, direction: 45, opacity: 0.3 },
    text: options.text || "",
    textStyle: {
      fontFamily: options.fontFamily || FONT,
      fontSize: options.fontSize || 22,
      color: options.color || C.white,
      bold: options.bold,
    },
  });
}

function line(slide, name, x1, y1, x2, y2, color = C.gold, width = 2) {
  return slide.connectors.add({
    name,
    start: { x: x1, y: y1 },
    end: { x: x2, y: y2 },
    line: { fill: color, width },
  });
}

function title(slide, kicker, headline, index) {
  text(slide, `kicker-${index}`, `${String(index).padStart(2, "0")}  ${kicker.toUpperCase()}`, frame(72, 40, 600, 41), { fontSize: 14, color: C.gold, bold: true });
  text(slide, `title-${index}`, headline, frame(72, 84, 880, 76), { fontSize: 34, bold: true });
  line(slide, `title-rule-${index}`, 72, 163, 1208, 163, C.panel2, 1);
}

function footer(slide, label, urls = []) {
  text(slide, `source-${slide.index + 1}`, `来源：${label}`, frame(72, 674, 1000, 41), { fontSize: 10, color: C.muted });
  text(slide, `page-${slide.index + 1}`, String(slide.index + 1).padStart(2, "0"), frame(1148, 670, 62, 41), { fontSize: 10, color: C.gold, bold: true });
  if (urls.length) slide.addNotes(`[Sources]\n${urls.join("\n")}`);
}

function addCircuit(slide, name, position, accent = C.gold) {
  const { width, height } = position;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
    <defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop stop-color="${accent}" stop-opacity=".95"/><stop offset="1" stop-color="#67B7F7" stop-opacity=".35"/></linearGradient></defs>
    <rect width="100%" height="100%" rx="28" fill="#111722"/>
    <g fill="none" stroke="url(#g)" stroke-width="2" opacity=".8">
      <path d="M28 ${height * .25}H${width * .36}V${height * .52}H${width * .72}V${height * .18}H${width - 26}"/>
      <path d="M18 ${height * .72}H${width * .26}V${height * .42}H${width * .58}V${height * .78}H${width - 18}"/>
      <circle cx="${width * .36}" cy="${height * .52}" r="7" fill="${accent}"/>
      <circle cx="${width * .72}" cy="${height * .18}" r="7" fill="#67B7F7"/>
      <circle cx="${width * .58}" cy="${height * .78}" r="7" fill="${accent}"/>
    </g>
    <circle cx="${width * .5}" cy="${height * .5}" r="${Math.min(width, height) * .23}" fill="#080A0F" stroke="${accent}" stroke-width="4"/>
    <text x="50%" y="57%" text-anchor="middle" font-family="Arial" font-size="${Math.min(width, height) * .27}" font-weight="700" fill="${accent}">₿</text>
  </svg>`;
  return slide.images.add({ name, position, dataUrl: svgDataUrl(svg), fit: "contain", alt: "Bitcoin circuit illustration" });
}

function addWhale(slide, name, position) {
  const { width, height } = position;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
    <rect width="100%" height="100%" rx="26" fill="#111722"/>
    <path d="M${width * .14} ${height * .56}c${width * .14}-${height * .29} ${width * .42}-${height * .29} ${width * .56}-${height * .06} ${width * .08}-${height * .13} ${width * .18}-${height * .18} ${width * .28}-${height * .13}-${width * .04} ${height * .12}-${width * .12} ${height * .2}-${width * .23} ${height * .25}-${width * .1} ${height * .24}-${width * .37} ${height * .33}-${width * .61} ${height * .04}z" fill="#67B7F7" opacity=".86"/>
    <path d="M${width * .42} ${height * .25}c0-${height * .12} ${width * .06}-${height * .2} ${width * .13}-${height * .24}" fill="none" stroke="#F4D06F" stroke-width="4" stroke-linecap="round"/>
    <circle cx="${width * .28}" cy="${height * .5}" r="4" fill="#080A0F"/>
    <g fill="#F4D06F"><circle cx="${width * .72}" cy="${height * .76}" r="5"/><circle cx="${width * .8}" cy="${height * .66}" r="3"/><circle cx="${width * .87}" cy="${height * .78}" r="7"/></g>
  </svg>`;
  return slide.images.add({ name, position, dataUrl: svgDataUrl(svg), fit: "contain", alt: "Whale distribution illustration" });
}

function newDarkDeck() {
  // The deck-level design grammar is materialized through the shared helpers
  // below. Source-free theme customization remains fail-closed in the codec.
  return Presentation.create({ slideSize: { width: 1280, height: 720 } });
}

function addBaseTransition(slide) {
  slide.setTransition({ effect: "fade", speed: "fast", durationMs: 420, advanceOnClick: true });
}

function buildBitcoinDeck() {
  const deck = newDarkDeck();

  // 1. Cover: visual carrier is a large editable SVG illustration plus type.
  const cover = deck.slides.add({ name: "封面" });
  cover.setBackground(C.bg);
  box(cover, "cover-gold-rail", frame(0, 0, 22, 720), { geometry: "rect", fill: C.gold, line: { fill: C.gold, width: 0 }, shadow: false });
  text(cover, "cover-label", "DIGITAL ASSETS  /  AUG 2026", frame(74, 76, 500, 28), { fontSize: 14, color: C.gold, bold: true });
  text(cover, "cover-title", "比特币大涨", frame(74, 150, 710, 90), { fontSize: 58, bold: true });
  text(cover, "cover-title-2", "2026 年 8 月行情驱动因素解析", frame(74, 255, 730, 64), { fontSize: 32, color: C.gold2, bold: true });
  text(cover, "cover-subtitle", "一周 +22%：多因素共振下的暴涨逻辑", frame(78, 345, 620, 48), { fontSize: 24, color: C.white });
  const coverHero = addCircuit(cover, "cover-bitcoin-circuit", frame(830, 105, 360, 430));
  box(cover, "cover-stat", frame(77, 464, 335, 118), { fill: C.panel2, lineColor: C.gold, text: "+22%\n7 days", fontSize: 30, bold: true });
  text(cover, "cover-date", "2026.08.25  ·  MARKET BRIEF", frame(78, 625, 500, 24), { fontSize: 14, color: C.muted });
  cover.animations.add(coverHero, { effect: "zoom", start: "afterPrevious", durationMs: 720 });
  footer(cover, "公开市场资料整理，2026-08-25", [SOURCES.treasury, SOURCES.etf]);

  // 2. Framework: four forces orbit a central market outcome.
  const framework = deck.slides.add({ name: "四维框架" });
  framework.setBackground(C.bg);
  addBaseTransition(framework);
  title(framework, "FRAMEWORK", "上涨不是单一叙事，而是四股力量同时收紧供需", 2);
  const frameworkHero = box(framework, "market-focus-overview", frame(525, 252, 230, 230), { geometry: "ellipse", fill: C.gold, line: { fill: C.gold2, width: 3 }, color: C.ink, text: "+22%\n一周", fontSize: 34, bold: true });
  const drivers = [
    ["宏观", "长债回购扩大\n美元走弱", 126, 215, C.blue],
    ["政策", "白宫会面\nCLARITY 预期", 865, 215, C.gold2],
    ["资金", "ETF 5 日净流入\n$1.92B", 126, 465, C.green],
    ["杠杆", "空头强平\n加速追价", 865, 465, C.red],
  ];
  for (const [label, copy, left, top, accent] of drivers) {
    box(framework, `driver-${label}`, frame(left, top, 292, 142), { fill: C.panel2, lineColor: accent, text: `${label}\n${copy}`, fontSize: 22, bold: true });
  }
  line(framework, "framework-link-1", 418, 286, 525, 329, C.blue, 2);
  line(framework, "framework-link-2", 755, 329, 865, 286, C.gold2, 2);
  line(framework, "framework-link-3", 418, 536, 525, 409, C.green, 2);
  line(framework, "framework-link-4", 755, 409, 865, 536, C.red, 2);
  footer(framework, "OfficeKit 分析框架");

  // 3. Market: chart grows by time. Morph focuses framework outcome into price.
  const market = deck.slides.add({ name: "行情回顾" });
  market.setBackground(C.bg);
  title(market, "MARKET", "7 天从约 6.4 万跃升至 7.9 万美元，修复速度快于基本面变化", 3);
  const marketHero = box(market, "market-focus-detail", frame(970, 52, 205, 92), { geometry: "ellipse", fill: C.gold, line: { fill: C.gold2, width: 2 }, color: C.ink, text: "$79K", fontSize: 30, bold: true });
  market.setMorph({ from: framework, durationMs: 820, pairs: [{ key: "market-focus", from: frameworkHero, to: marketHero }] });
  const marketChart = market.charts.add("line", {
    name: "btc-price-chart",
    title: "BTC price · USD",
    position: frame(72, 184, 815, 408),
    categories: ["8/18", "8/19", "8/20", "8/21", "8/22", "8/23", "8/24", "8/25"],
    series: [{ name: "BTC", values: [64000, 68800, 71300, 74200, 76000, 77200, 78500, 79000], color: C.gold, line: { fill: C.gold, width: 3 }, marker: { symbol: "circle", size: 7, fill: C.gold2 } }],
    legend: false,
    axes: { category: { title: "Date" }, value: { title: "USD", min: 60000, max: 82000, majorUnit: 5000 } },
    dataLabels: { showValue: false },
  });
  box(market, "market-cap", frame(930, 210, 260, 136), { fill: C.panel2, lineColor: C.gold, text: "$1.59T\n估算市值", fontSize: 30, bold: true });
  box(market, "market-context", frame(930, 378, 260, 166), { fill: C.panel, lineColor: C.muted, text: "今年 5 月以来首次\n逼近 8 万美元\n\n反弹 ≠ 新 ATH", fontSize: 21, bold: true });
  market.animations.add(marketChart, { effect: "wipe", direction: "right", chartBuild: "category-element", start: "onClick", durationMs: 720, staggerMs: 95, animateChartBackground: false });
  footer(market, "公开市场价格，2026-08-25", [SOURCES.treasury]);

  // 4. Macro causal chain.
  const macro = deck.slides.add({ name: "宏观传导" });
  macro.setBackground(C.bg);
  addBaseTransition(macro);
  title(macro, "MACRO", "财政部回购先压低长端利率，再把资金推向稀缺资产", 4);
  const causal = [
    ["01", "长债回购\n单次 ≥ $4B", C.blue],
    ["02", "长端收益率\n快速回落", C.gold2],
    ["03", "美元走弱\n流动性改善", C.green],
    ["04", "黄金 + BTC\n同步重估", C.gold],
  ];
  const causalNodes = [];
  const causalLinks = [];
  causal.forEach(([num, copy, accent], index) => {
    const left = 70 + index * 300;
    const node = box(macro, `macro-node-${num}`, frame(left, 255, 230, 190), { fill: C.panel2, lineColor: accent, text: `${num}\n${copy}`, fontSize: 24, bold: true });
    causalNodes.push(node);
    if (index > 0) {
      causalLinks.push(macro.connectors.add({
        name: `macro-arrow-${index}`,
        start: { x: left - 62, y: 350 },
        end: { x: left - 10, y: 350 },
        line: { fill: C.gold, width: 3, endArrow: "triangle" },
      }));
    }
  });
  causalNodes.forEach((node, index) => {
    macro.animations.add(node, { effect: "fade", start: index === 0 ? "onClick" : "afterPrevious", durationMs: 420 });
    if (causalLinks[index]) macro.animations.add(causalLinks[index], { effect: "wipe", direction: "right", start: "afterPrevious", durationMs: 260 });
  });
  text(macro, "macro-callout", "Debasement Trade 的核心不是“放水”二字，而是长期财政可信度折价。", frame(160, 520, 960, 42), { fontSize: 23, color: C.gold2, bold: true });
  footer(macro, "AP / US Treasury，2026-08-19", [SOURCES.treasury]);

  // 5. BTC vs Gold comparison.
  const debasement = deck.slides.add({ name: "硬资产对比" });
  debasement.setBackground(C.bg);
  addBaseTransition(debasement);
  title(debasement, "DEBASEMENT", "黄金定价避险，比特币放大流动性：同叙事、不同风险曲线", 5);
  const goldSide = box(debasement, "gold-side", frame(78, 196, 500, 360), { fill: "#1A1710", lineColor: C.gold2, text: "GOLD\n\n央行与机构持仓\n波动较低\n避险共识更成熟", fontSize: 27, bold: true });
  const btcSide = box(debasement, "btc-side", frame(702, 196, 500, 360), { fill: "#101824", lineColor: C.blue, text: "BITCOIN\n\n固定供给叙事\n全天候交易\n高弹性、高波动", fontSize: 27, bold: true });
  box(debasement, "versus", frame(594, 315, 92, 92), { geometry: "ellipse", fill: C.gold, line: { fill: C.gold2, width: 2 }, color: C.ink, text: "VS", fontSize: 24, bold: true });
  text(debasement, "debasement-bottom", "组合含义：黄金守住购买力，比特币表达对流动性的高 beta。", frame(210, 594, 860, 40), { fontSize: 22, color: C.gold2, bold: true });
  debasement.animations.add(goldSide, { effect: "fly", direction: "left", start: "onClick", durationMs: 480 });
  debasement.animations.add(btcSide, { effect: "fly", direction: "right", start: "withPrevious", durationMs: 480 });
  footer(debasement, "AP，2026-08-21", [SOURCES.treasury]);

  // 6. Policy timeline.
  const policy = deck.slides.add({ name: "政策时间线" });
  policy.setBackground(C.bg);
  addBaseTransition(policy);
  title(policy, "POLICY", "白宫直接听取行业反馈，监管从“执法优先”转向“规则竞争”", 6);
  line(policy, "timeline-rail", 118, 350, 1160, 350, C.gold, 4);
  const policyItems = [
    ["8/18", "监管议题升温", 170, C.muted],
    ["8/19", "白宫会见行业高管", 455, C.gold2],
    ["8/20", "CLARITY 路径讨论", 740, C.blue],
    ["NEXT", "市场等待立法细节", 1025, C.green],
  ];
  policyItems.forEach(([date, copy, left, accent], index) => {
    const dot = box(policy, `policy-dot-${index}`, frame(left, 329, 42, 42), { geometry: "ellipse", fill: accent, line: { fill: accent, width: 0 }, shadow: false });
    const label = text(policy, `policy-date-${index}`, date, frame(left - 30, 250, 110, 30), { fontSize: 18, color: accent, bold: true });
    const body = text(policy, `policy-copy-${index}`, copy, frame(left - 72, 400, 185, 64), { fontSize: 18, bold: true });
    policy.animations.add(dot, { effect: "zoom", start: index === 0 ? "onClick" : "afterPrevious", durationMs: 280 });
    policy.animations.add(label, { effect: "fade", start: "withPrevious", durationMs: 280 });
    policy.animations.add(body, { effect: "fade", start: "afterPrevious", durationMs: 330 });
  });
  box(policy, "policy-tag", frame(78, 548, 1124, 72), { fill: C.panel2, lineColor: C.gold, text: "定价变化：政策不确定性仍在，但“被禁止”的尾部风险下降。", fontSize: 23, bold: true });
  footer(policy, "The Block，2026-08-20", [SOURCES.policy]);

  // 7. ETF daily flows.
  const etf = deck.slides.add({ name: "ETF 资金流" });
  etf.setBackground(C.bg);
  addBaseTransition(etf);
  title(etf, "ETF FLOW", "现货 ETF 单周吸金 19.2 亿美元，IBIT 拿走近七成", 7);
  const etfChart = etf.charts.add("bar", {
    name: "etf-daily-inflow",
    title: "Daily net inflow · USD million",
    position: frame(72, 185, 790, 420),
    categories: ["Mon 17", "Tue 18", "Wed 19", "Thu 20", "Fri 21"],
    series: [{ name: "Net inflow", values: [297.56, 189.30, 517.19, 606.29, 307.45], color: C.gold }],
    legend: false,
    axes: { category: { title: "Aug 2026" }, value: { title: "USD mn", min: 0, max: 700, majorUnit: 100 } },
    dataLabels: { showValue: true, position: "outsideEnd" },
  });
  box(etf, "etf-week", frame(920, 205, 275, 145), { fill: C.panel2, lineColor: C.gold, text: "$1.92B\n周净流入", fontSize: 33, bold: true });
  box(etf, "ibit-share", frame(920, 386, 275, 168), { fill: C.panel, lineColor: C.blue, text: "$1.33B\nIBIT · 69%\n\n机构买盘集中", fontSize: 24, bold: true });
  etf.animations.add(etfChart, { effect: "wipe", direction: "up", chartBuild: "category-element", start: "onClick", durationMs: 680, staggerMs: 100, animateChartBackground: false });
  footer(etf, "Bitcoin.com ETF tracker，2026-08-21", [SOURCES.etf]);

  // 8. Demand vs supply imbalance.
  const supply = deck.slides.add({ name: "供需失衡" });
  supply.setBackground(C.bg);
  addBaseTransition(supply);
  title(supply, "SUPPLY", "ETF 日均买盘远超矿工新增供给，边际价格被少量流通盘决定", 8);
  text(supply, "demand-label", "ETF BUYING", frame(90, 220, 310, 32), { fontSize: 18, color: C.gold, bold: true });
  box(supply, "demand-bar", frame(90, 270, 920, 92), { geometry: "rect", fill: C.gold, line: { fill: C.gold, width: 0 }, color: C.ink, text: "约 3,000–4,000 BTC / day", fontSize: 26, bold: true, shadow: false });
  text(supply, "supply-label", "NEW MINING SUPPLY", frame(90, 420, 350, 32), { fontSize: 18, color: C.blue, bold: true });
  box(supply, "supply-bar", frame(90, 470, 230, 92), { geometry: "rect", fill: C.blue, line: { fill: C.blue, width: 0 }, color: C.ink, text: "约 450", fontSize: 26, bold: true, shadow: false });
  box(supply, "imbalance-multiple", frame(1040, 270, 160, 292), { fill: C.panel2, lineColor: C.gold2, text: "≈ 8×\n\n需求 / 新供给", fontSize: 30, bold: true });
  text(supply, "supply-caveat", "这是边际流量比较，不代表 ETF 吸收全部成交量。", frame(90, 610, 900, 28), { fontSize: 16, color: C.muted });
  footer(supply, "ETF 流量与区块奖励估算，2026-08");

  // 9. Short squeeze feedback loop.
  const squeeze = deck.slides.add({ name: "Short Squeeze" });
  squeeze.setBackground(C.bg);
  addBaseTransition(squeeze);
  title(squeeze, "LEVERAGE", "空头强平把趋势交易变成自我强化的买入回路", 9);
  const loopNodes = [
    ["价格上涨", 510, 190, C.gold],
    ["空单触发\n保证金不足", 840, 330, C.red],
    ["交易所\n强制买回", 625, 520, C.blue],
    ["现货深度\n被继续吃掉", 250, 470, C.green],
    ["涨幅扩大", 180, 245, C.gold2],
  ];
  const squeezeNodes = loopNodes.map(([label, left, top, accent], index) => box(squeeze, `squeeze-${index}`, frame(left, top, 210, 100), { fill: C.panel2, lineColor: accent, text: label, fontSize: 22, bold: true }));
  const loopLines = [
    [720, 240, 840, 355], [920, 430, 790, 520], [625, 570, 460, 520], [250, 470, 280, 345], [390, 260, 510, 230],
  ].map(([x1, y1, x2, y2], index) => squeeze.connectors.add({ name: `squeeze-arrow-${index}`, start: { x: x1, y: y1 }, end: { x: x2, y: y2 }, line: { fill: C.gold, width: 3, endArrow: "triangle" } }));
  squeezeNodes.forEach((node, index) => {
    squeeze.animations.add(node, { effect: "fade", start: index === 0 ? "onClick" : "afterPrevious", durationMs: 360 });
    squeeze.animations.add(loopLines[index], { effect: "wipe", direction: "right", start: "afterPrevious", durationMs: 260 });
  });
  box(squeeze, "squeeze-stat", frame(930, 188, 268, 104), { fill: "#25151A", lineColor: C.red, text: ">$4B\n看空仓位被清算", fontSize: 24, bold: true });
  footer(squeeze, "AP，统计至 2026-08-21", [SOURCES.treasury]);

  // 10. Whale split signal.
  const whale = deck.slides.add({ name: "链上分歧" });
  whale.setBackground(C.bg);
  addBaseTransition(whale);
  title(whale, "ON-CHAIN", "长期增持与短期抛售同时出现，市场正在换手而非单边一致", 10);
  const whaleImage = addWhale(whale, "whale-illustration", frame(78, 210, 440, 350));
  box(whale, "accumulate", frame(575, 205, 280, 168), { fill: C.panel2, lineColor: C.green, text: "+43,000 BTC\n此前 60 日增持", fontSize: 28, bold: true });
  box(whale, "distribute", frame(900, 205, 280, 168), { fill: C.panel2, lineColor: C.red, text: "−7,700 BTC\n8/19–8/22 转移/抛售", fontSize: 26, bold: true });
  box(whale, "interpretation", frame(575, 420, 605, 140), { fill: C.panel, lineColor: C.gold, text: "多头筹码仍厚，但 8 万美元附近的边际卖压已经出现。", fontSize: 25, bold: true });
  whale.animations.add(whaleImage, { effect: "zoom", start: "afterPrevious", durationMs: 520 });
  footer(whale, "Bitcoin.com，2026-08-22", [SOURCES.whale]);

  // 11. Sentiment and levels.
  const sentiment = deck.slides.add({ name: "情绪与技术面" });
  sentiment.setBackground(C.bg);
  addBaseTransition(sentiment);
  title(sentiment, "TECHNICALS", "8 万美元是情绪与筹码的共同压力位，追涨赔率开始下降", 11);
  const gauge = box(sentiment, "sentiment-gauge", frame(90, 205, 300, 300), { geometry: "ellipse", fill: C.panel2, line: { fill: C.gold, width: 12 }, text: "GREED\n74", fontSize: 34, bold: true });
  box(sentiment, "resistance", frame(470, 205, 320, 128), { fill: "#25151A", lineColor: C.red, text: "$80K\n关键阻力", fontSize: 30, bold: true });
  box(sentiment, "support", frame(470, 377, 320, 128), { fill: "#10231E", lineColor: C.green, text: "$74K\n首要支撑", fontSize: 30, bold: true });
  box(sentiment, "rsi", frame(850, 205, 340, 300), { fill: C.panel, lineColor: C.gold2, text: "RSI\n进入超买区\n\n趋势仍强\n回撤风险上升", fontSize: 27, bold: true });
  sentiment.animations.add(gauge, { phase: "emphasis", effect: "pulse", start: "afterPrevious", durationMs: 520 });
  footer(sentiment, "市场技术指标，2026-08-25");

  // 12. Scenario table, each scenario has a distinct contour and trigger.
  const scenarios = deck.slides.add({ name: "情景推演" });
  scenarios.setBackground(C.bg);
  addBaseTransition(scenarios);
  title(scenarios, "SCENARIOS", "后市取决于资金是否持续，而不是这轮涨幅本身", 12);
  const scenarioData = [
    ["突破上行", "ETF > $300M/日\n美债收益率续降", "$82K → $88K", C.green],
    ["高位盘整", "资金流降温\n监管无新进展", "$74K–$82K", C.gold],
    ["深度回调", "ETF 转负\n美元与收益率反弹", "$68K–$72K", C.red],
  ];
  scenarioData.forEach(([name, trigger, range, accent], index) => {
    const left = 72 + index * 400;
    box(scenarios, `scenario-${index}`, frame(left, 205, 350, 360), { fill: C.panel2, lineColor: accent, text: `${name}\n\n${trigger}\n\n${range}`, fontSize: 24, bold: true });
  });
  text(scenarios, "scenario-watch", "共同观察：ETF 资金流 · 美债收益率 · CLARITY 进度 · Jackson Hole 信号", frame(150, 610, 980, 30), { fontSize: 18, color: C.gold2, bold: true });
  footer(scenarios, "OfficeKit 情景分析，2026-08-25");

  // 13. Risk and synthesis.
  const risk = deck.slides.add({ name: "风险与结论" });
  risk.setBackground(C.bg);
  addBaseTransition(risk);
  title(risk, "SYNTHESIS", "真实买盘决定底部，空头轧空决定速度；两者不能混为一谈", 13);
  box(risk, "real-demand", frame(72, 205, 520, 310), { fill: "#10231E", lineColor: C.green, text: "真实买盘\n\nETF 持续净流入\n硬资产叙事升温\n监管尾部风险下降", fontSize: 27, bold: true });
  box(risk, "squeeze-demand", frame(688, 205, 520, 310), { fill: "#25151A", lineColor: C.red, text: "杠杆放大\n\n空头被迫买回\n流动性短时变薄\n涨速超过基本面", fontSize: 27, bold: true });
  const riskNumber = box(risk, "risk-focus", frame(404, 558, 472, 76), { fill: C.gold, line: { fill: C.gold2, width: 2 }, color: C.ink, text: "关键风险：8 万美元附近获利回吐", fontSize: 22, bold: true });
  risk.animations.add(riskNumber, { phase: "emphasis", effect: "pulse", start: "afterPrevious", durationMs: 520 });
  footer(risk, "综合公开资料，2026-08-25", [SOURCES.etf, SOURCES.whale]);

  // 14. Closing disclaimer.
  const close = deck.slides.add({ name: "免责声明" });
  close.setBackground(C.bg);
  addBaseTransition(close);
  addCircuit(close, "closing-circuit", frame(810, 130, 360, 390), C.gold2);
  text(close, "close-kicker", "END / DISCLOSURE", frame(78, 110, 400, 28), { fontSize: 14, color: C.gold, bold: true });
  text(close, "close-title", "在趋势里保持判断，\n在波动里保留余地。", frame(78, 198, 670, 150), { fontSize: 46, bold: true });
  text(close, "close-disclaimer", "本演示仅用于市场研究与信息交流，不构成投资建议。\n数字资产波动剧烈，请独立判断并自行承担风险。", frame(82, 435, 640, 100), { fontSize: 21, color: C.muted });
  footer(close, "OfficeKit，2026-08-25");

  return deck;
}

function buildArchitectureDeck() {
  const deck = newDarkDeck();
  const slide = deck.slides.add({ name: "Agent artifact compiler" });
  slide.setBackground(C.bg);
  title(slide, "ARCHITECTURE", "Agent 决定内容，编译器保证文件、定位与复核", 1);
  const labels = [
    ["Brief", "目标 / 受众 / 证据", C.blue],
    ["Plan", "叙事 / 设计语法", C.gold2],
    ["Compose", "原生对象 / 布局", C.green],
    ["Review", "结构 / 视觉 / 播放", C.red],
  ];
  const nodes = [];
  const links = [];
  labels.forEach(([name, copy, accent], index) => {
    const left = 70 + index * 300;
    const node = box(slide, `arch-${name}`, frame(left, 275, 235, 150), { fill: C.panel2, lineColor: accent, text: `${name}\n${copy}`, fontSize: 23, bold: true });
    nodes.push(node);
    if (index > 0) links.push(slide.connectors.add({ name: `arch-link-${index}`, start: { x: left - 55, y: 350 }, end: { x: left - 10, y: 350 }, line: { fill: C.gold, width: 3, endArrow: "triangle" } }));
  });
  nodes.forEach((node, index) => {
    slide.animations.add(node, { effect: "fade", start: index === 0 ? "onClick" : "afterPrevious", durationMs: 380 });
    if (links[index]) slide.animations.add(links[index], { effect: "wipe", direction: "right", start: "afterPrevious", durationMs: 240 });
  });
  footer(slide, "OfficeKit Causal Reveal example");
  return deck;
}

function buildBrandDeck() {
  const deck = newDarkDeck();
  const overview = deck.slides.add({ name: "Brand overview" });
  overview.setBackground(C.bg);
  const from = box(overview, "brand-hero-overview", frame(110, 155, 360, 360), { geometry: "ellipse", fill: C.gold, line: { fill: C.gold2, width: 3 }, color: C.ink, text: "01\nSIGNAL", fontSize: 36, bold: true });
  text(overview, "brand-title", "One signal.\nOne decisive move.", frame(560, 220, 590, 150), { fontFamily: "Aptos Display", fontSize: 48, bold: true });
  text(overview, "brand-caption", "Morph continuity keeps identity while the story changes scale.", frame(565, 405, 550, 80), { fontFamily: "Aptos", fontSize: 21, color: C.muted });
  footer(overview, "OfficeKit Morph Continuity example");

  const detail = deck.slides.add({ name: "Brand detail" });
  detail.setBackground(C.bg);
  const to = box(detail, "brand-hero-detail", frame(690, 70, 500, 500), { geometry: "ellipse", fill: C.gold, line: { fill: C.gold2, width: 3 }, color: C.ink, text: "01\nSIGNAL", fontSize: 48, bold: true });
  text(detail, "brand-detail-title", "Identity stays.\nMeaning expands.", frame(80, 190, 540, 150), { fontFamily: "Aptos Display", fontSize: 46, bold: true });
  text(detail, "brand-detail-copy", "The same native object moves, grows and remains editable after round-trip.", frame(84, 390, 500, 105), { fontFamily: "Aptos", fontSize: 22, color: C.muted });
  detail.setMorph({ from: overview, durationMs: 900, pairs: [{ key: "brand-signal", from, to }] });
  footer(detail, "OfficeKit Morph Continuity example");
  return deck;
}

async function emitDeck(name, deck) {
  const deckDir = path.join(outputRoot, name);
  const renderDir = path.join(deckDir, "svg");
  await mkdir(renderDir, { recursive: true });

  const verification = deck.verify({ visualQa: true });
  const file = await PresentationFile.exportPptx(deck);
  const bytes = new Uint8Array(await file.arrayBuffer());
  const pptxPath = path.join(deckDir, `${name}.pptx`);
  await writeFile(pptxPath, bytes);

  const imported = await PresentationFile.importPptx(file);
  const secondVerification = imported.verify({ visualQa: true });
  const motion = imported.inspect({ kind: "animation,morph", maxChars: 250_000 });
  for (const slide of imported.slides.items) {
    const svg = await (await slide.export({ format: "svg" })).text();
    await writeFile(path.join(renderDir, `${String(slide.index + 1).padStart(2, "0")}.svg`), svg);
  }

  const evidence = {
    artifact: path.basename(pptxPath),
    sha256: sha256(bytes),
    bytes: bytes.byteLength,
    slides: imported.slides.count,
    authoredVerification: {
      ok: verification.ok,
      errors: verification.issues.filter((issue) => issue.severity !== "warning").length,
      warnings: verification.issues.filter((issue) => issue.severity === "warning").length,
    },
    roundTripVerification: {
      ok: secondVerification.ok,
      errors: secondVerification.issues.filter((issue) => issue.severity !== "warning").length,
      warnings: secondVerification.issues.filter((issue) => issue.severity === "warning").length,
    },
    motionRecords: parseNdjson(motion.ndjson),
    playbackEvidence: "structural",
    powerpointPlayback: "unverified",
  };
  await writeFile(path.join(deckDir, "evidence.json"), `${JSON.stringify(evidence, null, 2)}\n`);
  return { pptxPath, evidencePath: path.join(deckDir, "evidence.json"), ...evidence };
}

await mkdir(outputRoot, { recursive: true });
const outputs = [];
outputs.push(await emitDeck("bitcoin-rally-2026", buildBitcoinDeck()));
outputs.push(await emitDeck("architecture-causal-reveal", buildArchitectureDeck()));
outputs.push(await emitDeck("brand-morph-continuity", buildBrandDeck()));
await writeFile(path.join(outputRoot, "manifest.json"), `${JSON.stringify({ generatedAt: new Date().toISOString(), outputs }, null, 2)}\n`);
console.log(JSON.stringify({ outputRoot, outputs }, null, 2));
