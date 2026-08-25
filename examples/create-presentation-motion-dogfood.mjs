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

import { Presentation, PresentationFile, reviewArtifact } from "office-kit";

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
    geometry: options.geometry || "rect",
    position,
    fill: options.fill || C.panel,
    line: options.line || { fill: options.lineColor || C.gold, width: options.lineWidth ?? 1 },
    shadow: options.shadow === true ? { color: "#000000", blurRadius: 10, distance: 4, direction: 45, opacity: 0.3 } : undefined,
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

function metricPanel(slide, name, position, options) {
  const panel = box(slide, `${name}-surface`, position, {
    fill: options.fill || C.panel2,
    lineColor: options.accent || C.gold,
  });
  const valueSize = options.valueSize || 32;
  text(slide, `${name}-value`, options.value, frame(position.left + 12, position.top + 4, position.width - 24, 58), {
    fontFamily: options.valueFont || "Aptos Display",
    fontSize: valueSize,
    color: options.valueColor || C.white,
    bold: true,
  });
  text(slide, `${name}-label`, options.label, frame(position.left + 12, position.top + 62, position.width - 24, 42), {
    fontSize: options.labelSize || 20,
    color: options.labelColor || C.white,
    bold: true,
  });
  if (options.detail) {
    text(slide, `${name}-detail`, options.detail, frame(position.left + 12, position.top + 104, position.width - 24, Math.max(44, position.height - 112)), {
      fontSize: options.detailSize || 17,
      color: options.detailColor || C.muted,
      bold: options.detailBold,
    });
  }
  return panel;
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
  metricPanel(cover, "cover-stat", frame(77, 464, 335, 118), { value: "+22%", label: "7 DAYS", accent: C.gold, valueSize: 34 });
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
  title(market, "MARKET", "7 天从约 6.4 万涨至 7.9 万美元，反弹快于基本面修复", 3);
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
  metricPanel(market, "market-cap", frame(930, 210, 260, 136), { value: "$1.59T", label: "估算市值", accent: C.gold, valueSize: 31 });
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
  const goldSide = box(debasement, "gold-side", frame(78, 196, 500, 360), { fill: "#1A1710", lineColor: C.gold2, text: "GOLD\n\n央行与机构持仓\n波动较低\n避险共识更成熟", fontSize: 22, bold: true });
  const btcSide = box(debasement, "btc-side", frame(702, 196, 500, 360), { fill: "#101824", lineColor: C.blue, text: "BITCOIN\n\n固定供给叙事\n全天候交易\n高弹性、高波动", fontSize: 22, bold: true });
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
  metricPanel(etf, "etf-week", frame(920, 205, 275, 145), { value: "$1.92B", label: "周净流入", accent: C.gold, valueSize: 33 });
  metricPanel(etf, "ibit-share", frame(920, 386, 275, 168), { value: "$1.33B", label: "IBIT · 69%", detail: "机构买盘集中", accent: C.blue, valueSize: 29, detailColor: C.white, detailBold: true });
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
  metricPanel(supply, "imbalance-multiple", frame(1040, 270, 160, 292), { value: "≈ 8×", label: "需求 / 新供给", accent: C.gold2, valueSize: 34, labelSize: 17 });
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
  metricPanel(squeeze, "squeeze-stat", frame(930, 188, 268, 104), { value: ">$4B", label: "看空仓位被清算", fill: "#25151A", accent: C.red, valueSize: 28, labelSize: 18 });
  footer(squeeze, "AP，统计至 2026-08-21", [SOURCES.treasury]);

  // 10. Whale split signal.
  const whale = deck.slides.add({ name: "链上分歧" });
  whale.setBackground(C.bg);
  addBaseTransition(whale);
  title(whale, "ON-CHAIN", "长期增持与短期抛售同时出现，市场正在换手而非单边一致", 10);
  const whaleImage = addWhale(whale, "whale-illustration", frame(78, 210, 440, 350));
  metricPanel(whale, "accumulate", frame(575, 205, 280, 168), { value: "+43,000", label: "BTC · 此前 60 日增持", accent: C.green, valueSize: 28, labelSize: 17 });
  metricPanel(whale, "distribute", frame(900, 205, 280, 168), { value: "−7,700", label: "BTC · 8/19–8/22", detail: "转移 / 抛售", accent: C.red, valueSize: 28, labelSize: 17, detailColor: C.white, detailBold: true });
  box(whale, "interpretation", frame(575, 420, 605, 140), { fill: C.panel, lineColor: C.gold, text: "多头筹码仍厚，但 8 万美元附近的边际卖压已经出现。", fontSize: 23, bold: true });
  whale.animations.add(whaleImage, { effect: "zoom", start: "afterPrevious", durationMs: 520 });
  footer(whale, "Bitcoin.com，2026-08-22", [SOURCES.whale]);

  // 11. Sentiment and levels.
  const sentiment = deck.slides.add({ name: "情绪与技术面" });
  sentiment.setBackground(C.bg);
  addBaseTransition(sentiment);
  title(sentiment, "TECHNICALS", "8 万美元是情绪与筹码的共同压力位，追涨赔率开始下降", 11);
  const gauge = box(sentiment, "sentiment-gauge", frame(90, 205, 300, 300), { geometry: "ellipse", fill: C.panel2, line: { fill: C.gold, width: 12 }, text: "GREED  74", fontSize: 30, bold: true });
  metricPanel(sentiment, "resistance", frame(470, 205, 320, 128), { value: "$80K", label: "关键阻力", fill: "#25151A", accent: C.red, valueSize: 30 });
  metricPanel(sentiment, "support", frame(470, 377, 320, 128), { value: "$74K", label: "首要支撑", fill: "#10231E", accent: C.green, valueSize: 30 });
  box(sentiment, "rsi", frame(850, 205, 340, 300), { fill: C.panel, lineColor: C.gold2, text: "RSI\n进入超买区\n\n趋势仍强\n回撤风险上升", fontSize: 22, bold: true });
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

function buildManagementDeck() {
  const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const paper = "#F2EFE8";
  const ink = "#1B1D1F";
  const red = "#C54B36";
  const blue = "#285F8F";
  const green = "#3B7458";
  const rule = "#B8B2A8";

  const managementHeader = (slide, section, claim, index) => {
    text(slide, `management-section-${index}`, `${String(index).padStart(2, "0")} / ${section}`, frame(64, 42, 270, 28), { fontSize: 13, color: red, bold: true });
    text(slide, `management-claim-${index}`, claim, frame(64, 86, 1080, 76), { fontSize: 34, color: ink, bold: true });
    line(slide, `management-rule-${index}`, 64, 170, 1216, 170, rule, 1);
  };

  const cover = deck.slides.add({ name: "转正答辩：从交付到体系" });
  cover.setBackground(paper);
  box(cover, "cover-red-field", frame(0, 0, 24, 720), { fill: red, line: { fill: red, width: 0 } });
  text(cover, "cover-eyebrow", "PROBATION REVIEW / 2026", frame(72, 74, 500, 28), { fontSize: 14, color: red, bold: true });
  text(cover, "cover-management-title-1", "从完成功能，", frame(72, 146, 760, 80), { fontSize: 52, color: ink, bold: true });
  text(cover, "cover-management-title-2", "到建立可持续交付体系", frame(72, 236, 780, 80), { fontSize: 52, color: ink, bold: true });
  line(cover, "cover-management-rule", 74, 360, 760, 360, ink, 3);
  text(cover, "cover-management-subtitle", "转正答辩 · Agent 基础设施与 Office 工程", frame(74, 386, 640, 42), { fontSize: 22, color: "#56595C" });
  text(cover, "cover-management-number", "03", frame(920, 130, 240, 170), { fontSize: 118, color: red, bold: true });
  text(cover, "cover-management-number-label", "个可复用系统\n替代一次性交付", frame(930, 328, 260, 90), { fontSize: 24, color: ink, bold: true });
  footer(cover, "OfficeKit management-report example");

  const outcomes = deck.slides.add({ name: "阶段成果" });
  outcomes.setBackground(paper);
  outcomes.setTransition({ effect: "fade", durationMs: 420, advanceOnClick: true });
  managementHeader(outcomes, "OUTCOMES", "真正的增量不是功能数量，而是交付能力开始复用", 2);
  const outcomeChart = outcomes.charts.add("bar", {
    name: "delivery-compounding",
    title: "Reusable delivery capacity",
    position: frame(64, 218, 690, 380),
    categories: ["第 1 月", "第 2 月", "第 3 月"],
    series: [{ name: "可复用能力", values: [28, 61, 100], color: red }],
    legend: false,
    axes: { category: { title: "阶段" }, value: { title: "Index", min: 0, max: 110, majorUnit: 20 } },
    dataLabels: { showValue: true, position: "outsideEnd" },
  });
  text(outcomes, "outcome-large", "100", frame(850, 220, 300, 112), { fontSize: 76, color: red, bold: true });
  text(outcomes, "outcome-large-label", "能力复用指数", frame(854, 350, 300, 42), { fontSize: 20, color: ink, bold: true });
  line(outcomes, "outcome-side-rule", 850, 390, 1160, 390, rule, 1);
  text(outcomes, "outcome-list", "Office 文件原生编译\n持久任务与恢复\n结构、视觉与播放复核", frame(850, 420, 330, 150), { fontSize: 22, color: ink });
  outcomes.animations.add(outcomeChart, { effect: "wipe", direction: "up", chartBuild: "category-element", start: "onClick", durationMs: 650, staggerMs: 100 });
  footer(outcomes, "阶段复盘示例数据");

  const system = deck.slides.add({ name: "工作方式升级" });
  system.setBackground(paper);
  system.setTransition({ effect: "fade", durationMs: 420, advanceOnClick: true });
  managementHeader(system, "SYSTEM", "交付从个人记忆迁移到可恢复、可验证的工作流", 3);
  const lanes = [
    ["01", "计划", "受众、结论、证据和设计方向先落盘", blue],
    ["02", "编译", "可编辑原生对象承载图表、关系与叙事", red],
    ["03", "复核", "语义、结构、布局、视觉和交付形成证据", green],
  ];
  const laneShapes = [];
  lanes.forEach(([number, label, copy, accent], index) => {
    const top = 220 + index * 130;
    box(system, `lane-band-${number}`, frame(64, top, 1120, 98), { fill: index % 2 ? "#E7E2D9" : "#EDE9E1", line: { fill: "transparent", width: 0 } });
    text(system, `lane-number-${number}`, number, frame(86, top + 20, 80, 54), { fontSize: 28, color: accent, bold: true });
    text(system, `lane-label-${number}`, label, frame(190, top + 22, 150, 48), { fontSize: 27, color: ink, bold: true });
    const copyShape = text(system, `lane-copy-${number}`, copy, frame(390, top + 24, 700, 44), { fontSize: 21, color: ink });
    laneShapes.push(copyShape);
  });
  laneShapes.forEach((shape, index) => system.animations.add(shape, { effect: "wipe", direction: "right", start: index === 0 ? "onClick" : "afterPrevious", durationMs: 380 }));
  footer(system, "OfficeKit durable authoring workflow");

  const next = deck.slides.add({ name: "下一阶段" });
  next.setBackground(paper);
  next.setTransition({ effect: "fade", durationMs: 420, advanceOnClick: true });
  managementHeader(next, "NEXT", "下一阶段只追三件事：真实采用、复杂保真、跨平台验收", 4);
  const actions = [
    ["真实采用", "用答辩、报告和模板续写\n暴露可用性问题", "01", red],
    ["复杂保真", "让第三方 PPTX 成为\n可继续编程的初始状态", "02", blue],
    ["跨平台", "补齐 Windows 原生播放\n与 Live host 验收", "03", green],
  ];
  actions.forEach(([label, copy, number, accent], index) => {
    const left = 64 + index * 390;
    text(next, `next-number-${number}`, number, frame(left, 226, 100, 72), { fontSize: 42, color: accent, bold: true });
    line(next, `next-rule-${number}`, left, 315, left + 330, 315, accent, 3);
    text(next, `next-label-${number}`, label, frame(left, 338, 300, 52), { fontSize: 27, color: ink, bold: true });
    text(next, `next-copy-${number}`, copy, frame(left, 410, 320, 110), { fontSize: 20, color: "#4C4E50" });
  });
  text(next, "next-decision", "需要的支持：用真实任务评估价值，而不是继续用 API 数量替代产品判断。", frame(64, 598, 1100, 40), { fontSize: 20, color: ink, bold: true });
  footer(next, "OfficeKit management-report example");
  return deck;
}

function buildBrandDeck() {
  const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const black = "#050505";
  const acid = "#E7FF3D";
  const coral = "#FF5B45";
  const bone = "#F5F1E8";

  const ray = (slide, name, position, flip = false) => {
    const { width, height } = position;
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
      <rect width="100%" height="100%" fill="${black}"/>
      <g transform="${flip ? `translate(${width} 0) scale(-1 1)` : ""}">
        <path d="M0 ${height * .52}L${width} 0v${height * .16}z" fill="${acid}"/>
        <path d="M0 ${height * .52}L${width} ${height * .22}v${height * .16}z" fill="${coral}"/>
        <path d="M0 ${height * .52}L${width} ${height * .46}v${height * .08}z" fill="${bone}"/>
        <path d="M0 ${height * .52}L${width} ${height * .68}v${height * .18}z" fill="${coral}" opacity=".72"/>
        <path d="M0 ${height * .52}L${width} ${height * .9}v${height * .1}z" fill="${acid}" opacity=".78"/>
      </g>
    </svg>`;
    return slide.images.add({ name, position, dataUrl: svgDataUrl(svg), fit: "cover", alt: "High-contrast signal rays" });
  };

  const overview = deck.slides.add({ name: "Brand overview" });
  overview.setBackground(black);
  ray(overview, "brand-rays-overview", frame(708, 0, 572, 720));
  const from = box(overview, "brand-signal-overview", frame(72, 76, 34, 548), { fill: acid, line: { fill: acid, width: 0 } });
  text(overview, "brand-series", "OFFICEKIT / BRAND SIGNAL 01", frame(146, 80, 500, 26), { fontFamily: "Arial", fontSize: 13, color: acid, bold: true });
  text(overview, "brand-title-1", "MAKE", frame(140, 140, 530, 90), { fontFamily: "Arial Black", fontSize: 62, color: bone, bold: true });
  text(overview, "brand-title-2", "THE IDEA", frame(140, 230, 530, 90), { fontFamily: "Arial Black", fontSize: 62, color: bone, bold: true });
  text(overview, "brand-title-3", "UNMISSABLE.", frame(140, 320, 530, 90), { fontFamily: "Arial Black", fontSize: 62, color: bone, bold: true });
  line(overview, "brand-title-rule", 144, 480, 620, 480, coral, 6);
  text(overview, "brand-caption-1", "A launch deck creates one memory.", frame(146, 500, 500, 44), { fontFamily: "Arial", fontSize: 19, color: bone });
  text(overview, "brand-caption-2", "That memory becomes action.", frame(146, 548, 500, 44), { fontFamily: "Arial", fontSize: 19, color: bone });
  text(overview, "brand-index", "01", frame(600, 628, 70, 50), { fontFamily: "Arial Black", fontSize: 30, color: acid, bold: true });

  const detail = deck.slides.add({ name: "Brand detail" });
  detail.setBackground(acid);
  ray(detail, "brand-rays-detail", frame(0, 0, 510, 720), true);
  const to = box(detail, "brand-signal-detail", frame(600, 54, 52, 606), { fill: black, line: { fill: black, width: 0 } });
  text(detail, "brand-detail-series", "ONE SIGNAL / CONTINUOUS IDEA", frame(706, 66, 500, 64), { fontFamily: "Arial", fontSize: 13, color: black, bold: true });
  ["IDENTITY", "STAYS.", "MEANING", "EXPANDS."].forEach((value, index) => {
    text(detail, `brand-detail-title-${index + 1}`, value, frame(700, 132 + index * 82, 500, 82), { fontFamily: "Arial Black", fontSize: 54, color: black, bold: true });
  });
  text(detail, "brand-detail-copy-1", "Morph moves the signal.", frame(704, 500, 470, 44), { fontFamily: "Arial", fontSize: 18, color: black });
  text(detail, "brand-detail-copy-2", "Native, editable, continuous.", frame(704, 548, 470, 44), { fontFamily: "Arial", fontSize: 18, color: black });
  line(detail, "brand-detail-rule", 706, 626, 1154, 626, coral, 6);
  detail.setMorph({ from: overview, durationMs: 900, pairs: [{ key: "brand-signal", from, to }] });
  return deck;
}

function createStrategyPlan(deck, strategy) {
  const pages = deck.slides.items.map((slide, index) => {
    const page = strategy.pages[index] || {};
    const animations = slide.animations.items;
    const morph = slide.morph.value;
    let purpose = "continuity";
    let recipe = "calm-continuity";
    if (morph) {
      purpose = "morph";
      recipe = "morph-continuity";
    } else if (animations.some((animation) => animation.chartBuild)) {
      purpose = "data-reveal";
      recipe = "data-rise";
    } else if (animations.some((animation) => animation.effect === "pulse")) {
      purpose = "focus";
      recipe = "focus-pulse";
    } else if (animations.length > 1) {
      purpose = "causal-sequence";
      recipe = "causal-reveal";
    }
    const transition = morph ? "morph" : slide.transition.toJSON()?.effect || "none";
    const semanticUnits = animations.map((animation, unitIndex) => ({
      id: `motion-${index + 1}-${unitIndex + 1}`,
      targetRole: animation.targetKind || "native visual",
      order: unitIndex + 1,
    }));
    if (morph) semanticUnits.push({ id: `motion-${index + 1}-morph`, targetRole: "continuity signal", order: semanticUnits.length + 1 });
    return {
      id: `page-${String(index + 1).padStart(2, "0")}`,
      readerTask: page.readerTask || `Understand ${slide.name || `page ${index + 1}`}`,
      claim: page.claim || slide.name || `Page ${index + 1}`,
      evidence: page.evidence || ["Native editable presentation objects"],
      contentBudget: { maxCharacters: page.maxCharacters || 1_500, maxObjects: page.maxObjects || 80 },
      compositionIntent: page.compositionIntent || "A mixed visual carrier connects the page claim to supporting evidence",
      ...(semanticUnits.length || transition !== "none" ? {
        motionIntent: { purpose, recipe, units: semanticUnits, transition },
      } : {}),
    };
  });
  return {
    schema: "office-kit/presentation-authoring-plan/v1",
    mode: "create",
    brief: {
      audience: strategy.audience,
      purpose: strategy.purpose,
      primaryJob: strategy.primaryJob,
      supportingJobs: strategy.supportingJobs,
      expectedOutcome: strategy.expectedOutcome,
      mediumFit: "strong",
      afterUse: strategy.afterUse,
      deliveryMode: strategy.deliveryMode,
    },
    narrative: { thesis: strategy.thesis, sections: strategy.sections },
    design: {
      sourceMode: "self-directed",
      mechanismPacks: strategy.mechanismPacks,
      motionPolicy: "adaptive",
      scenario: { primary: strategy.scenario, secondary: strategy.secondaryScenario || null },
      direction: { name: strategy.directionName, rationale: strategy.directionRationale },
      designGrammar: strategy.designGrammar,
    },
    pages,
    editorial: { voice: strategy.voice, lockedFacts: strategy.lockedFacts || [], avoid: strategy.avoid || [] },
    artifactRefs: [],
    recipe: "tasks/create.md",
    unresolved: [],
    nextAction: "Review the rendered deck and fix only concrete communication or presentation failures",
  };
}

async function emitDeck(name, deck, plan) {
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
  const review = await reviewArtifact(file, {
    format: "pptx",
    outputPath: pptxPath,
    authoringPlan: plan,
    layout: false,
    playbackEvidence: "structural",
    visualReview: "requires-human",
  });

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
    strategy: {
      primaryJob: review.design.strategy.primaryJob,
      scenario: review.design.strategy.scenario,
      direction: review.design.strategy.direction,
      deliveryMode: review.design.strategy.deliveryMode,
    },
    review: {
      verdict: review.verdict,
      designStatus: review.design.status,
      motionStatus: review.motion.status,
      issues: [...review.design.issues, ...review.motion.issues].map((issue) => ({ severity: issue.severity, type: issue.type, message: issue.message })),
    },
    playbackEvidence: "structural",
    powerpointPlayback: "unverified",
  };
  await writeFile(path.join(deckDir, "evidence.json"), `${JSON.stringify(evidence, null, 2)}\n`);
  return { pptxPath, evidencePath: path.join(deckDir, "evidence.json"), ...evidence };
}

await mkdir(outputRoot, { recursive: true });
const outputs = [];
const bitcoin = buildBitcoinDeck();
outputs.push(await emitDeck("bitcoin-rally-2026", bitcoin, createStrategyPlan(bitcoin, {
  audience: "Investors and financial professionals familiar with digital assets",
  purpose: "Explain why Bitcoin rallied and frame the next decision without overstating certainty",
  primaryJob: "explain",
  supportingJobs: ["inform", "decide"],
  expectedOutcome: "The audience distinguishes persistent demand from short-squeeze acceleration and knows which signals to monitor next",
  afterUse: "Investment discussion, scenario monitoring, and decision record",
  deliveryMode: "hybrid",
  thesis: "Liquidity, policy, institutional demand and leverage reinforced one another, but only persistent flows can sustain the rally",
  sections: ["Framework", "Drivers", "Risk", "Scenarios"],
  mechanismPacks: ["enterprise-data-review", "visual-narrative"],
  scenario: "analysis-decision",
  directionName: "Black-gold market terminal",
  directionRationale: "A dark information-dense field lets evidence charts carry authority while gold marks the decision signal",
  designGrammar: {
    palette: { roles: { background: C.bg, surface: C.panel, evidence: C.blue, signal: C.gold, positive: C.green, risk: C.red } },
    typography: { roles: { display: FONT, body: FONT, numeric: "Aptos Display" } },
    geometry: "Square analytical panels, thin dividers, one circular market signal only when it encodes a cycle or gauge",
    densityRhythm: "Dense evidence pages alternate with sparse synthesis and transition pages",
    carriers: "Charts, causal diagrams, comparison fields, and data-linked native vectors",
    forbidden: ["decorative card wall", "unexplained empty space", "animation without information order"],
  },
  voice: "Professional, evidence-led, and explicit about uncertainty",
  lockedFacts: ["$1.92B weekly ETF inflow", "$2.99B liquidation event", "BTC moved from about $64K to $79K"],
  avoid: ["investment certainty", "generic crypto hype", "repetitive contrast slogans"],
  pages: Array.from({ length: 14 }, (_, index) => ({
    compositionIntent: index === 0 ? "Large editable Bitcoin SVG illustration and decisive typography" : index === 1 ? "Relationship diagram around one market outcome" : "Mixed chart, diagram, or typographic evidence carrier with bounded annotation",
  })),
})));

const management = buildManagementDeck();
outputs.push(await emitDeck("management-probation-review", management, createStrategyPlan(management, {
  audience: "Direct manager and promotion reviewers",
  purpose: "Show how probation-period delivery became a reusable operating system and request support for the next stage",
  primaryJob: "report",
  supportingJobs: ["align", "decide"],
  expectedOutcome: "Reviewers understand the compounded delivery capability and agree on the three next priorities",
  afterUse: "Review record and next-quarter alignment note",
  deliveryMode: "live",
  thesis: "The durable result is a repeatable delivery system, not a list of isolated features",
  sections: ["Outcome", "System", "Next decision"],
  mechanismPacks: ["enterprise-data-review", "editorial-minimal"],
  scenario: "management-report",
  directionName: "Editorial operating review",
  directionRationale: "Paper, rules and numbered evidence create sober managerial credibility without imitating a dashboard",
  designGrammar: {
    palette: { roles: { paper: "#F2EFE8", ink: "#1B1D1F", decision: "#C54B36", system: "#285F8F", proof: "#3B7458" } },
    typography: { roles: { title: FONT, body: FONT, numeric: "Aptos Display" } },
    geometry: "Square fields, ruled columns, broad horizontal bands, no generic cards",
    densityRhythm: "Sparse cover, one chart, one process page, one decision page",
    carriers: "Bar chart, operating bands, and ruled decision columns",
    forbidden: ["rounded card wall", "decorative dashboard chrome", "unsubstantiated owner language"],
  },
  voice: "Specific, modest, and outcome-led",
  avoid: ["feature inventory", "repeating owner", "inflated claims"],
  pages: [
    { compositionIntent: "Sparse editorial typography with one oversized numeric proof" },
    { compositionIntent: "Native bar chart and large numeric evidence rail" },
    { compositionIntent: "Horizontal process diagram using broad bands and thin rules" },
    { compositionIntent: "Typographic decision table with three ruled columns and one explicit ask" },
  ],
})));

const brand = buildBrandDeck();
outputs.push(await emitDeck("brand-morph-continuity", brand, createStrategyPlan(brand, {
  audience: "Launch-event audience",
  purpose: "Make one product identity memorable and demonstrate continuous native motion",
  primaryJob: "mobilize",
  supportingJobs: ["persuade"],
  expectedOutcome: "The audience retains one visual signal and connects it to the launch action",
  afterUse: "Launch-stage playback and editable campaign source",
  deliveryMode: "live",
  thesis: "One strong signal can hold identity while the story changes scale",
  sections: ["Signal", "Expansion"],
  mechanismPacks: ["brand-launch", "visual-narrative"],
  scenario: "brand-creative",
  directionName: "Acid signal broadcast",
  directionRationale: "Full-bleed black, acid yellow, coral rays and oversized type create recall without relying on card containers",
  designGrammar: {
    palette: { roles: { black: "#050505", acid: "#E7FF3D", coral: "#FF5B45", bone: "#F5F1E8" } },
    typography: { roles: { display: "Arial Black", body: "Arial" } },
    geometry: "Full-bleed light rays, one moving vertical signal, hard edges",
    densityRhythm: "Two high-impact sparse statements connected by Morph",
    carriers: "Oversized typography and a native signal bar over vector rays",
    forbidden: ["dashboard panels", "gold circle hero", "multiple competing motifs"],
  },
  voice: "Direct, energetic, and memorable",
  avoid: ["feature list", "decorative explanation", "generic launch superlatives"],
  pages: [
    { compositionIntent: "Full-bleed vector rays, oversized typography, and one native vertical signal" },
    { compositionIntent: "Full-bleed inverted field, oversized typography, and the same Morph signal" },
  ],
})));
await writeFile(path.join(outputRoot, "manifest.json"), `${JSON.stringify({ generatedAt: new Date().toISOString(), outputs }, null, 2)}\n`);
console.log(JSON.stringify({ outputRoot, outputs }, null, 2));
