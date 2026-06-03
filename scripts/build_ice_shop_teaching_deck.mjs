import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const OUT = path.join(ROOT, "La_Neige_法式雪冰室創業教學簡報.pptx");
const PREVIEW = path.join(ROOT, "outputs", "la-neige-deck-preview");
const artifactPath = "/Users/chiashanyeh/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const { Presentation, PresentationFile } = await import(pathToFileURL(artifactPath));

const C = {
  cream: "#F7EFE2", paper: "#FFF9EF", burgundy: "#8B1E3F",
  green: "#1F5C4B", gold: "#C9A45C", cocoa: "#2B2420",
  muted: "#6A5650", pale: "#F4F7F2"
};

function line(fill = "transparent", width = 0) { return { style: "solid", fill, width }; }
function shape(slide, x, y, w, h, fill, stroke = "transparent", width = 0) {
  return slide.shapes.add({ geometry: "rect", position: { left: x, top: y, width: w, height: h }, fill, line: line(stroke, width) });
}
function text(slide, value, x, y, w, h, opt = {}) {
  const s = shape(slide, x, y, w, h, opt.fill ?? "transparent");
  s.text = value;
  s.text.fontSize = opt.size ?? 28;
  s.text.color = opt.color ?? C.cocoa;
  s.text.bold = Boolean(opt.bold);
  s.text.typeface = "Arial Unicode MS";
  s.text.alignment = opt.align ?? "left";
  s.text.verticalAlignment = opt.valign ?? "top";
  s.text.insets = opt.insets ?? { left: 0, right: 0, top: 0, bottom: 0 };
  return s;
}
function bg(slide, color = C.paper) { shape(slide, 0, 0, 1280, 720, color); }
function footer(slide, label) {
  shape(slide, 70, 650, 1140, 1, "#E8DCCB");
  text(slide, label, 74, 666, 900, 28, { size: 18, color: "#9D8F81" });
}
function title(slide, h, sub = "") {
  text(slide, h, 70, 58, 1050, 68, { size: 44, bold: true, color: C.burgundy });
  if (sub) text(slide, sub, 74, 126, 980, 36, { size: 22, color: C.muted });
}
function card(slide, h, b, x, y, w, hgt, accent) {
  shape(slide, x, y, w, hgt, "#FFFFFF", "#E8DCCB", 1);
  shape(slide, x, y, 12, hgt, accent);
  text(slide, h, x + 28, y + 22, w - 56, 34, { size: 25, bold: true });
  text(slide, b, x + 28, y + 68, w - 56, hgt - 84, { size: 20, color: C.muted });
}

const p = Presentation.create({ slideSize: { width: 1280, height: 720 } });

{
  const s = p.slides.add(); bg(s, C.cream);
  text(s, "La Neige\n法式雪冰室", 78, 86, 640, 150, { size: 58, bold: true, color: C.burgundy });
  text(s, "台灣冰品 x 法式生活感 x 父子退休創業", 84, 268, 740, 42, { size: 28, bold: true, color: C.green });
  text(s, "教學案例｜創業計劃書、品牌網站、會員活動與 GitHub Pages 範例", 84, 336, 900, 34, { size: 24, color: C.muted });
  shape(s, 880, 105, 250, 250, C.paper, C.gold, 6);
  text(s, "把台灣的冰\n做成一場\n法式午後", 912, 155, 190, 130, { size: 32, bold: true, align: "center", valign: "middle", color: C.cocoa });
  footer(s, "封面｜La Neige 法式雪冰室創業教學案例");
}

{
  const s = p.slides.add(); bg(s);
  title(s, "品牌不是從 LOGO 開始", "它從人、記憶和生活方式開始。");
  card(s, "琰綱", "英國建築與藝術背景，把光影、比例、法式街角感帶回台灣。", 92, 240, 330, 210, C.burgundy);
  card(s, "父親", "退休後重新開始，喜歡登山與重機，成為品牌的生活溫度。", 475, 240, 330, 210, C.green);
  card(s, "店", "一間台灣街景中令人驚艷的法式冰店，讓客人願意拍照、停留、再訪。", 858, 240, 330, 210, C.gold);
  footer(s, "品牌故事｜父子創業與生活美學");
}

{
  const s = p.slides.add(); bg(s, C.cream);
  title(s, "市場定位", "客單 90-150 元，用親民價格做出品味感。");
  card(s, "目標客群", "年輕拍照族、情侶、朋友聚會、親子、退休族、山友與重機騎士。", 110, 230, 500, 150, C.green);
  card(s, "差異化", "法式視覺、台灣味覺、透明製作、社群互動、父子真實故事。", 670, 230, 500, 150, C.burgundy);
  text(s, "核心策略：讓每一碗冰都能被拍、被說、被分享。", 170, 485, 940, 54, { size: 34, bold: true, color: C.cocoa, align: "center" });
  footer(s, "市場定位｜品味生活與社群傳播");
}

{
  const s = p.slides.add(); bg(s);
  title(s, "商品架構", "夏天賣冰，冬天靠熱飲湯品與會員回訪。");
  [["雪花冰", "焦糖伯爵、草莓玫瑰、芒果香草", C.burgundy], ["刨冰", "檸檬塔刨冰、紅豆牛奶、芋泥奶霜", C.gold], ["冬季飲品", "黑糖薑奶、桂圓紅棗、燒仙草奶茶、花生湯圓可可", C.green], ["湯品", "南瓜濃湯、番茄蔬菜湯", C.cocoa]].forEach((d, i) => {
    const x = 110 + (i % 2) * 560; const y = 215 + Math.floor(i / 2) * 160;
    card(s, d[0], d[1], x, y, 500, 120, d[2]);
  });
  footer(s, "產品規劃｜菜單不多但記憶點清楚");
}

{
  const s = p.slides.add(); bg(s, C.cream);
  title(s, "網站範例就是教學工具", "把品牌、商品、社群、會員、互動一次串起來。");
  ["品牌故事", "菜單主題", "拍照牆", "投票留言", "會員訂閱", "阿山哥俱樂部"].forEach((v, i) => {
    const x = 110 + (i % 3) * 360; const y = 220 + Math.floor(i / 3) * 130;
    shape(s, x, y, 280, 82, "#FFFFFF", "#E8DCCB", 1);
    text(s, v, x + 20, y + 24, 240, 34, { size: 26, bold: true, align: "center" });
  });
  text(s, "本機範例：http://localhost:4173｜GitHub Pages：main / docs", 120, 540, 1040, 40, { size: 24, bold: true, color: C.green, align: "center" });
  footer(s, "網站原型｜可連結教學範例");
}

{
  const s = p.slides.add(); bg(s);
  title(s, "全年主題行銷", "用插畫與每月限定，讓社群不只是一張公告。");
  [["2月", "玫瑰情人節", "草莓玫瑰雙人冰"], ["4月", "伯爵茶香季", "焦糖伯爵雪花冰"], ["7月", "芒果左岸", "芒果香草雪花冰"], ["12月", "聖誕雪花季", "聖誕草莓雪花冰"]].forEach((d, i) => {
    const x = 100 + i * 280;
    shape(s, x, 235, 230, 220, "#FFFFFF", "#E8DCCB", 1);
    text(s, d[0], x + 20, 260, 190, 42, { size: 34, bold: true, color: C.burgundy, align: "center" });
    text(s, d[1], x + 18, 325, 194, 32, { size: 22, bold: true, align: "center" });
    text(s, d[2], x + 18, 380, 194, 45, { size: 19, color: C.muted, align: "center" });
  });
  footer(s, "社群行銷｜每月主題與短影音任務");
}

{
  const s = p.slides.add(); bg(s, C.cream);
  title(s, "阿山哥俱樂部", "把父親的登山與重機生活，變成會員回訪機制。");
  card(s, "登山任務", "PO 爬山照，填山名或步道，產生 ASHAN 月份優惠碼。", 120, 240, 480, 170, C.green);
  card(s, "重機任務", "PO 騎士重機照，介紹品牌車款與騎乘心得，產生 RIDER 月份優惠碼。", 680, 240, 480, 170, C.burgundy);
  text(s, "優惠：招牌冰折 20 元，冬季熱飲 9 折。", 230, 500, 820, 46, { size: 32, bold: true, align: "center", color: C.cocoa });
  footer(s, "會員互動｜生活方式帶動再訪");
}

{
  const s = p.slides.add(); bg(s);
  title(s, "財務試算", "先看日客數，再決定擴張。");
  [["保守", "40 人/日", "約 15.6 萬/月"], ["穩定", "70 人/日", "約 27.3 萬/月"], ["旺季", "100 人/日", "約 39 萬/月"]].forEach((d, i) => {
    card(s, d[0], `${d[1]}\n客單 130 元\n${d[2]}`, 130 + i * 365, 230, 300, 220, [C.green, C.gold, C.burgundy][i]);
  });
  text(s, "食材成本 30-38%｜租金控制營收 10% 內｜冬季用會員與熱飲補淡季", 110, 535, 1060, 38, { size: 24, bold: true, align: "center" });
  footer(s, "財務規劃｜小店先穩定再複製");
}

{
  const s = p.slides.add(); bg(s, C.cream);
  text(s, "教學重點", 100, 90, 450, 60, { size: 50, bold: true, color: C.burgundy });
  text(s, "學生學到的不是一間冰店，而是把個人故事、商品、空間、社群、會員與網路展示整合成創業方案的方法。", 110, 205, 980, 120, { size: 32, bold: true, color: C.cocoa });
  text(s, "GitHub Pages 可作為長期教學連結，課堂、作業與創業討論都能隨時取用。", 110, 390, 980, 80, { size: 30, color: C.green, bold: true });
  footer(s, "結語｜從案例到可教、可改、可展示");
}

await fs.mkdir(PREVIEW, { recursive: true });
for (let i = 0; i < p.slides.count; i += 1) {
  const slide = p.slides.getItem(i);
  const png = await p.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(path.join(PREVIEW, `slide-${String(i + 1).padStart(2, "0")}.png`), Buffer.from(await png.arrayBuffer()));
}
const pptx = await PresentationFile.exportPptx(p);
await pptx.save(OUT);
console.log(JSON.stringify({ output: OUT, slideCount: p.slides.count, preview: PREVIEW }, null, 2));
