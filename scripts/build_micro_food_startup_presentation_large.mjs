import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const OUT = path.join(ROOT, "微型餐飲創業教學簡報_葉佳山.pptx");
const PREVIEW_DIR = path.join(ROOT, "outputs", "projection-ppt-preview");
const LAYOUT_DIR = path.join(ROOT, "outputs", "projection-ppt-layout");
const artifactPath = "/Users/chiashanyeh/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const { Presentation, PresentationFile } = await import(pathToFileURL(artifactPath));

const C = {
  ink: "#223042",
  muted: "#728090",
  bg: "#FFFDF8",
  warm: "#F8F4EC",
  blue: "#2F78A6",
  teal: "#4D9A8C",
  coral: "#E1765E",
  gold: "#D8A23B",
  green: "#759E55",
  line: "#E6DED2",
  paleBlue: "#EAF4F8",
  paleTeal: "#EAF7F3",
  paleCoral: "#FFF0EB",
  paleGold: "#FFF7DF",
};

function line(fill = "transparent", width = 0) {
  return { style: "solid", fill, width };
}

function shape(slide, x, y, w, h, fill, stroke = "transparent", width = 0) {
  return slide.shapes.add({
    geometry: "rect",
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: line(stroke, width),
  });
}

function text(slide, value, x, y, w, h, options = {}) {
  const s = shape(slide, x, y, w, h, options.fill ?? "transparent", options.stroke ?? "transparent", options.strokeWidth ?? 0);
  s.text = value;
  s.text.fontSize = options.size ?? 28;
  s.text.color = options.color ?? C.ink;
  s.text.bold = Boolean(options.bold);
  s.text.typeface = "Arial Unicode MS";
  s.text.alignment = options.align ?? "left";
  s.text.verticalAlignment = options.valign ?? "top";
  s.text.insets = options.insets ?? { left: 0, right: 0, top: 0, bottom: 0 };
  return s;
}

function bg(slide, color = C.bg) {
  shape(slide, 0, 0, 1280, 720, color);
}

function footer(slide, chapter) {
  shape(slide, 70, 650, 1140, 1, "#EEE5D8");
  text(slide, chapter, 74, 666, 850, 28, { size: 18, color: "#B3A79B" });
}

function title(slide, heading, sub = "") {
  text(slide, heading, 70, 56, 980, 64, { size: 46, bold: true });
  if (sub) text(slide, sub, 74, 126, 940, 38, { size: 24, color: C.muted });
}

function card(slide, heading, body, x, y, w, h, accent, fill = "#FFFFFF") {
  shape(slide, x, y, w, h, fill, C.line, 1);
  shape(slide, x, y, 10, h, accent);
  text(slide, heading, x + 28, y + 22, w - 56, 34, { size: 26, bold: true });
  text(slide, body, x + 28, y + 72, w - 56, h - 92, { size: 22, color: "#405163" });
}

function chip(slide, label, x, y, w, fill) {
  shape(slide, x, y, w, 46, fill);
  text(slide, label, x + 16, y + 9, w - 32, 28, { size: 20, bold: true, color: "#FFFFFF", align: "center" });
}

function addCover(p) {
  const slide = p.slides.add();
  bg(slide, C.warm);
  text(slide, "微型餐飲創業\n教學計劃", 76, 92, 640, 138, { size: 56, bold: true });
  text(slide, "投影授課版｜大字、少字、好講解", 80, 258, 700, 40, { size: 30, color: C.blue, bold: true });
  chip(slide, "商業便當", 82, 340, 170, C.teal);
  chip(slide, "1-2 人小火鍋", 276, 340, 206, C.coral);
  chip(slide, "台灣特色 × 法式冰店", 506, 340, 290, C.gold);
  text(slide, "主講者｜弘光科技大學助理教授 葉佳山", 82, 438, 760, 38, { size: 28, bold: true });
  shape(slide, 912, 118, 250, 250, C.paleTeal);
  text(slide, "先小量\n驗證", 940, 150, 194, 86, { size: 42, bold: true, color: C.teal, align: "center", valign: "middle" });
  text(slide, "再穩定\n複製", 940, 250, 194, 86, { size: 42, bold: true, color: C.coral, align: "center", valign: "middle" });
  footer(slide, "封面｜微型餐飲創業教學計劃");
}

function addSimpleCards(p, heading, sub, cards, chapter, bgColor = C.bg) {
  const slide = p.slides.add();
  bg(slide, bgColor);
  title(slide, heading, sub);
  const cols = cards.length === 4 ? 2 : 3;
  const w = cols === 2 ? 500 : 340;
  const h = cards.length === 4 ? 135 : 185;
  cards.forEach((c, i) => {
    const x = cols === 2 ? 110 + (i % 2) * 560 : 90 + i * 395;
    const y = cols === 2 ? 220 + Math.floor(i / 2) * 170 : 245;
    card(slide, c[0], c[1], x, y, w, h, c[2], c[3] ?? "#FFFFFF");
  });
  footer(slide, chapter);
}

const p = Presentation.create({ slideSize: { width: 1280, height: 720 } });

addCover(p);

addSimpleCards(p, "這門課帶走什麼？", "把創業從熱情，變成可以驗證的小生意。", [
  ["看懂基本盤", "法規、稅務、食安、資金與成本。", C.blue, C.paleBlue],
  ["選對第一步", "便當、小火鍋、冰店，先選一個試。", C.teal, C.paleTeal],
  ["完成可試賣計劃", "菜單、成本、定價、30 日行動表。", C.coral, C.paleCoral],
], "課程總目標｜先小量驗證，再穩定複製");

{
  const slide = p.slides.add();
  bg(slide, C.warm);
  title(slide, "第一章：準備成為老闆", "失敗不是污點，是下一次少繳學費的資料。");
  shape(slide, 98, 245, 500, 190, C.paleGold, "#E8D6A8", 1);
  text(slide, "這一次，\n不靠衝動開店。", 138, 288, 420, 88, { size: 40, bold: true, color: C.coral, align: "center", valign: "middle" });
  text(slide, "靠準備、紀律、驗證與服務，\n把小生意做成能修正、能持續的事業。", 690, 260, 430, 140, { size: 30, bold: true });
  shape(slide, 690, 435, 430, 10, C.coral);
  footer(slide, "第一章｜創業者的準備與勇氣");
}

{
  const slide = p.slides.add();
  bg(slide);
  title(slide, "今天就去挑戰", "投影版只記 4 件事：先有資料，再做決定。");
  const items = [["列資金", C.teal], ["訪問顧客", C.coral], ["試做產品", C.gold], ["小量試賣", C.blue]];
  items.forEach((it, i) => {
    const x = 120 + i * 275;
    text(slide, String(i + 1), x, 240, 150, 86, { size: 72, bold: true, color: it[1], align: "center" });
    text(slide, it[0], x - 18, 344, 186, 42, { size: 30, bold: true, align: "center" });
  });
  text(slide, "今天先完成一個小動作，創業就開始有方向。", 164, 500, 952, 46, { size: 32, bold: true, color: C.ink, align: "center" });
  footer(slide, "第一章｜現在就去挑戰的方向");
}

addSimpleCards(p, "第二章：共同基本功", "小店最重要的不是漂亮，而是底線清楚。", [
  ["法規線", "登記、稅籍、食品業者登錄。", C.blue],
  ["食安線", "來源、溫度、保存、衛生。", C.teal],
  ["成本線", "食材、包材、人力、平台費。", C.coral],
  ["品牌線", "店名、菜單、招牌記憶點。", C.gold],
], "第二章｜微型餐飲創業的共同基本功", C.warm);

{
  const slide = p.slides.add();
  bg(slide);
  title(slide, "定價不要只看隔壁", "先算得出來，才有資格談賣多少。");
  card(slide, "每份直接成本", "主食材 + 副食材 + 調味 + 包材 + 平台費", 120, 240, 470, 160, C.blue, C.paleBlue);
  card(slide, "基本售價", "售價 = 每份直接成本 / 目標食材成本率", 690, 240, 470, 160, C.coral, C.paleCoral);
  text(slide, "例：成本 55 元，食材成本率 35%，售價約 157 元。", 120, 500, 1040, 48, { size: 32, bold: true, align: "center" });
  footer(slide, "第二章｜成本與定價基本公式");
}

addSimpleCards(p, "第三章：商業便當", "最適合從穩定需求與團訂開始。", [
  ["客群集中", "辦公室、醫護、補習班、會議。", C.teal],
  ["定位清楚", "高蛋白、少油鹽、台式經典。", C.gold],
  ["預購降低風險", "先收單，再備料，少報廢。", C.coral],
], "第三章｜商業便當創業教學", C.warm);

{
  const slide = p.slides.add();
  bg(slide);
  title(slide, "便當的一日流程", "穩定、準時、好吃、好算。");
  const steps = ["收單", "備料", "烹調", "分裝", "配送", "結算"];
  steps.forEach((s, i) => {
    const x = 120 + i * 175;
    shape(slide, x, 270, 120, 90, i % 2 ? C.paleTeal : C.paleBlue, C.line, 1);
    text(slide, s, x + 8, 298, 104, 34, { size: 28, bold: true, align: "center" });
    if (i < steps.length - 1) shape(slide, x + 132, 310, 34, 8, C.gold);
  });
  text(slide, "第一個月重點：試賣 30 份，開發 2 個團訂客戶。", 134, 480, 1010, 44, { size: 31, bold: true, color: C.coral, align: "center" });
  footer(slide, "第三章｜商業便當營運流程");
}

addSimpleCards(p, "第四章：1-2 人小火鍋", "小座位，也能有明確情境。", [
  ["一人療癒鍋", "安靜、乾淨、下班後放鬆。", C.teal],
  ["雙人儀式鍋", "約會、好友聊天、有套餐感。", C.coral],
  ["台味特色鍋", "麻油雞、剝皮辣椒、沙茶石頭。", C.gold],
], "第四章｜主題化 1-2 人小火鍋", C.warm);

{
  const slide = p.slides.add();
  bg(slide);
  title(slide, "小火鍋菜單模組", "3 湯底 × 3 主餐，先穩再變。");
  card(slide, "湯底", "沙茶石頭湯\n法式番茄蔬菜湯\n麻油雞風味湯", 155, 230, 400, 230, C.blue, C.paleBlue);
  card(slide, "主餐", "豬肉片\n雞腿肉\n海鮮或菇菇蔬食", 725, 230, 400, 230, C.coral, C.paleCoral);
  text(slide, "+ 手工丸餃 / 起司 / 香草奶油升級", 310, 520, 660, 42, { size: 30, bold: true, color: C.gold, align: "center" });
  footer(slide, "第四章｜菜單模組與成本控制");
}

addSimpleCards(p, "第五章：台灣特色 × 法式冰店", "法式不是昂貴，是細節與層次。", [
  ["台灣水果", "跟著季節走，不硬做。", C.gold],
  ["法式細節", "香緹、庫利、布蕾、酥脆感。", C.coral],
  ["社群畫面", "每一碗都要能說故事。", C.teal],
], "第五章｜台灣特色加法式風格冰店", C.warm);

{
  const slide = p.slides.add();
  bg(slide);
  title(slide, "四款招牌冰", "初期不做多，先做出清楚記憶點。");
  const items = [
    ["芒果米香香緹冰", C.gold],
    ["鐵觀音焦糖冰", C.teal],
    ["花生粉粿布蕾冰", C.coral],
    ["仙草莓果法式冰", C.blue],
  ];
  items.forEach((it, i) => {
    const x = 135 + (i % 2) * 525;
    const y = 240 + Math.floor(i / 2) * 145;
    shape(slide, x, y, 450, 86, "#FFFFFF", C.line, 1);
    shape(slide, x, y, 20, 86, it[1]);
    text(slide, it[0], x + 48, y + 25, 370, 34, { size: 28, bold: true });
  });
  footer(slide, "第五章｜招牌產品設計");
}

addSimpleCards(p, "第六章：重新出發策略", "先復盤，不自責；先測試，不擴張。", [
  ["三不", "不租大店面\n不做太多品項\n不在沒資料時擴張", C.coral],
  ["三要", "做 30 天小量測試\n每天記錄回饋\n開業前查法規食安", C.teal],
], "第六章｜失敗者的復盤與再出發", C.warm);

{
  const slide = p.slides.add();
  bg(slide);
  title(slide, "開業前 30 日", "用一個月讓想法變清楚。");
  const phases = [
    ["1-6日", "選業態\n訪問顧客", C.teal],
    ["7-12日", "試做產品\n成本定價", C.gold],
    ["13-18日", "查登記食安\n設計菜單", C.blue],
    ["19-24日", "小量試賣\n修正產品", C.coral],
    ["25-30日", "建立 SOP\n決定下一輪", C.green],
  ];
  phases.forEach((phase, i) => {
    const x = 88 + i * 232;
    shape(slide, x, 235, 180, 210, "#FFFFFF", C.line, 1);
    shape(slide, x, 235, 180, 54, phase[2]);
    text(slide, phase[0], x + 12, 248, 156, 28, { size: 22, bold: true, color: "#FFFFFF", align: "center" });
    text(slide, phase[1], x + 16, 325, 148, 74, { size: 27, bold: true, align: "center", valign: "middle" });
  });
  footer(slide, "附錄｜開業前 30 日行動表");
}

{
  const slide = p.slides.add();
  bg(slide, C.warm);
  text(slide, "今天，不必等待所有條件完美。", 110, 145, 1060, 60, { size: 44, bold: true, align: "center" });
  text(slide, "先完成一張成本表、一次訪談、一道產品、一次試賣。", 120, 270, 1040, 52, { size: 36, bold: true, color: C.blue, align: "center" });
  shape(slide, 315, 395, 650, 120, "#FFFFFF", C.line, 1);
  text(slide, "微型創業不是小夢想，\n是把大夢想放進可以執行的第一步。", 355, 425, 570, 70, { size: 28, bold: true, color: C.coral, align: "center", valign: "middle" });
  text(slide, "主講者｜弘光科技大學助理教授 葉佳山", 360, 570, 560, 34, { size: 24, align: "center" });
  footer(slide, "第八章｜講師結語");
}

await fs.mkdir(PREVIEW_DIR, { recursive: true });
await fs.mkdir(LAYOUT_DIR, { recursive: true });
for (let i = 0; i < p.slides.count; i += 1) {
  const slide = p.slides.getItem(i);
  const num = String(i + 1).padStart(2, "0");
  const png = await p.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(path.join(PREVIEW_DIR, `slide-${num}.png`), Buffer.from(await png.arrayBuffer()));
  const layout = await p.export({ slide, format: "layout" });
  await fs.writeFile(path.join(LAYOUT_DIR, `slide-${num}.layout.json`), await layout.text(), "utf8");
}

const pptx = await PresentationFile.exportPptx(p);
await pptx.save(OUT);
console.log(JSON.stringify({ output: OUT, slideCount: p.slides.count, previewDir: PREVIEW_DIR, layoutDir: LAYOUT_DIR }, null, 2));
