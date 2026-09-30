// deck-engine.js — 通用 deck-spec JSON 渲染引擎 v2
// 用法: node deck-engine.js <spec.json>
// v1: 7 种页面（cover/intro/product/scene/symbols/colors/finale）——完全保持兼容
// v2: 新增 18 种常用版式 —— toc/section/bullets/imageText/cards/compare/timeline/
//     process/matrix/pyramid/stats/chart/table/quote/team/gallery/case/contact
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const pptxgen = require("pptxgenjs");

const specFile = process.argv[2];
if (!specFile) { console.error("用法: node deck-engine.js <spec.json>"); process.exit(1); }
let spec;
try { spec = JSON.parse(fs.readFileSync(specFile, "utf-8")); }
catch (e) { console.error("spec JSON 解析失败: " + e.message); process.exit(1); }
const ASSETS = (spec.assets_dir || ".").replace(/\\/g, "/");
const OUT = spec.output || "deck-output.pptx";

// ---- 尺寸表（dims.py 自动生成，缺 PIL 则按方形兜底）----
const dimsPy = path.join(__dirname, "dims.py");
if (!fs.existsSync(dimsPy)) {
  fs.writeFileSync(dimsPy, [
    'import json, glob, os, sys',
    'from PIL import Image',
    'd = sys.argv[1].replace("\\\\", "/")',
    'o = {}',
    'for f in glob.glob(d + "/*.*"):',
    '    try: o[os.path.basename(f)] = list(Image.open(f).size)',
    '    except Exception: pass',
    'print(json.dumps(o))'].join("\n"));
}
let DIMS = {};
try { DIMS = JSON.parse(execFileSync("python", [dimsPy, ASSETS], { encoding: "utf-8" })); }
catch (e) { console.warn("⚠️ 尺寸表读取失败，全部按方形兜底"); }
const ratio = f => { const d = DIMS[path.basename(f)]; return d && d[1] ? d[0] / d[1] : 1; };
const fit = (f, bw, bh) => { const r = ratio(f); let w = bw, h = w / r; if (h > bh) { h = bh; w = h * r; } return { w, h }; };
const exist = f => fs.existsSync(path.join(ASSETS, f));

// ---- 校验（不过即退回，不产 pptx）----
const errors = [];
const TYPE_FIELDS = {
  cover: ["title"], intro: ["title", "rows"], product: ["title", "images", "rows"],
  scene: ["image", "chip"], symbols: ["title", "rows", "images"],
  colors: ["title", "swatches", "rows"], finale: ["title", "thumbs", "quote"],
  toc: ["title", "rows"], section: ["title"], bullets: ["title", "rows"],
  imageText: ["title", "image", "rows"], cards: ["title", "cards"],
  compare: ["title", "columns"], timeline: ["title", "nodes"], process: ["title", "steps"],
  matrix: ["title", "quadrants"], pyramid: ["title", "levels"], stats: ["title", "stats"],
  chart: ["title", "chart"], table: ["title", "head", "rows"], quote: ["quote"],
  team: ["title", "images"], gallery: ["title", "images"],
  case: ["title", "image", "rows"], contact: ["title"],
};
(spec.pages || []).forEach((pg, i) => {
  const n = "第" + (i + 1) + "页";
  if (!TYPE_FIELDS[pg.type]) { errors.push(n + " type 非法: " + pg.type); return; }
  TYPE_FIELDS[pg.type].forEach(f => { if (pg[f] === undefined) errors.push(n + " 缺字段 " + f); });
  if (pg.image === undefined && pg.images === undefined && pg.thumbs === undefined && pg.levels === undefined && pg.nodes === undefined) { /* 无害：只是没有图 */ }
  const imgs = [].concat(pg.image || [], pg.thumbs || [],
    (pg.images || []).map(x => (typeof x === "string" ? x : x.src)),
    pg.ref ? [pg.ref.src] : []);
  imgs.filter(Boolean).forEach(f => { if (!exist(f)) errors.push(n + " 图片不存在: " + f); });
});
if (errors.length) {
  console.error("校验未通过，退回修改：");
  errors.forEach(e => console.error(" - " + e));
  process.exit(1);
}

// ---- 主题与助手 ----
const T = Object.assign({ bg: "16130F", ink: "F0E9DC", sub: "C9BFA9", gold: "C9A053",
  red: "A8332A", hair: "3A332A", muted: "9A8F7D", font: "微软雅黑" }, spec.deck.theme || {});
const clean = h => String(h || "").replace("#", "");
const p = new pptxgen();
p.defineLayout({ name: "W", width: 13.333, height: 7.5 });
p.layout = "W";
p.defineSlideMaster({ title: "BG", background: { color: clean(T.bg) } });

function pad2(n) { n = String(n); return n.length < 2 ? "0" + n : n; }
function lighten(hex, f) {
  const h = clean(hex) || "16130F";
  const n = parseInt(h, 16);
  let r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  r = Math.min(255, Math.round(r + (255 - r) * f));
  g = Math.min(255, Math.round(g + (255 - g) * f));
  b = Math.min(255, Math.round(b + (255 - b) * f));
  return ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}
const CARD = clean(T.card) || lighten(T.bg, 0.11);
function txt(v) { return Array.isArray(v) ? v.join("\n") : (v === undefined || v === null ? "" : String(v)); }
function pair(v, a, b, c) {
  if (Array.isArray(v)) { const o = {}; o[a] = v[0]; o[b] = v[1]; if (c) { o[c] = v[2]; } return o; }
  return v || {};
}
function dwidth(s) {                       // 近似显示宽度（单位 em）：中日韩算 1，其余算 0.5
  let n = 0;
  const str = String(s === undefined || s === null ? "" : s);
  for (let i = 0; i < str.length; i++) { n += (str.charCodeAt(i) > 0x2E7F ? 1 : 0.5); }
  return n;
}
function fitSize(body, boxW, availH, base, min) {
  min = min || 7.5;
  const em = dwidth(body);
  for (let s = base; s >= min; s -= 0.5) {
    const perLine = Math.max((boxW * 72 * 0.96) / s, 1);
    const lines = Math.max(1, Math.ceil(em / perLine));
    if ((lines * s * 1.3) / 72 + 0.05 <= availH) { return s; }
  }
  return min;
}
function autoPitch(n, yTop, yBot, maxPitch) {   // 行数多了就把行距压到装得下为止
  return Math.min(maxPitch, (yBot - yTop) / Math.max(n, 1));
}
function T_(s, o) {
  s.addText(txt(o.t), { x: o.x, y: o.y, w: o.w, h: o.h, fontFace: T.font, fontSize: o.size || 12,
    color: o.color || T.sub, bold: !!o.bold, align: o.align || "left",
    valign: o.valign || "top", lineSpacingMultiple: 1.25,
    // 内容超量时自动缩排，别让字糊出框（讲稿写太长是常态）
    fit: o.fit || "shrink", ...(o.extra || {}) });
}
function imgCard(s, f, x, y, w, h) {
  s.addImage({ path: path.join(ASSETS, f), x, y, w, h,
    shadow: { type: "outer", angle: 90, blur: 10, color: "000000", offset: 4, opacity: 0.5 } });
}
function header(s, num, title, pill) {
  if (num) s.addText(num, { shape: "roundRect", rectRadius: 0.07, fill: { color: clean(T.red) },
    x: 0.62, y: 0.44, w: 0.56, h: 0.44, fontFace: T.font, fontSize: 14, bold: true,
    color: "FFFFFF", align: "center", valign: "middle" });
  T_(s, { t: title, x: 1.34, y: 0.36, w: 8.6, h: 0.62, size: 26, color: T.ink, bold: true });
  if (pill) s.addText(pill, { shape: "roundRect", rectRadius: 0.16, fill: { color: clean(T.bg) },
    line: { color: clean(T.gold), width: 1 }, x: 10.45, y: 0.49, w: 2.28, h: 0.36,
    fontFace: T.font, fontSize: 10.5, color: T.gold, align: "center", valign: "middle" });
  s.addShape("rect", { x: 0.62, y: 1.14, w: 12.09, h: 0.014, fill: { color: clean(T.hair) } });
}
function footer(s, text) {
  if (!text) { return; }
  s.addShape("rect", { x: 0.62, y: 7.02, w: 0.3, h: 0.05, fill: { color: clean(T.gold) } });
  T_(s, { t: text, x: 1.02, y: 6.86, w: 11, h: 0.4, size: 12, color: T.gold, bold: true });
}
function row(s, x, y, w, lead, body, hair, pitch) {
  pitch = pitch || 1.5;
  T_(s, { t: lead, x, y, w, h: 0.34, size: 13.5, color: T.gold, bold: true });
  T_(s, { t: body, x, y: y + 0.37, w, h: Math.max(pitch - 0.45, 0.3), size: fitSize(body, w, pitch - 0.45, 12) });
  if (hair) s.addShape("rect", { x, y: y + pitch - 0.08, w, h: 0.012, fill: { color: clean(T.hair) } });
}
function refBox(s, f, x, y, w, caption) {
  const b = fit(f, w, w);
  imgCard(s, f, x, y, b.w, b.h);
  T_(s, { t: caption, x, y: y + b.h + 0.08, w: w + 0.6, h: 0.4, size: 9.5, color: T.muted });
}
function strip(s, x, y, items, each) {
  (items || []).forEach((c, i) => {
    const xx = x + i * each;
    s.addShape("rect", { x: xx, y, w: 0.5, h: 0.26, fill: { color: clean(c[1]) }, line: { color: clean(T.hair), width: 0.75 } });
    T_(s, { t: c[0], x: xx + 0.56, y: y + 0.02, w: each - 0.6, h: 0.24, size: 8.5, color: T.sub });
  });
}
function panel(s, x, y, w, h, fill, line) {
  s.addShape("roundRect", { x, y, w, h, rectRadius: 0.05,
    fill: { color: clean(fill || CARD) }, line: { color: clean(line || T.hair), width: 0.5 } });
}
function dot(s, cx, cy, d, color) {
  s.addShape("ellipse", { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: clean(color || T.gold) } });
}
function miniBar(s, x, y, w, h, pct, color) {
  s.addShape("rect", { x, y, w, h, fill: { color: clean(T.hair) } });
  s.addShape("rect", { x, y, w: Math.max(w * Math.max(0, Math.min(1, pct)), 0.02), h, fill: { color: clean(color || T.gold) } });
}

// ---- 模板 ----
const R = {
  // ======================= 一期：原有 7 种（保持不变）=======================
  cover(pg) {
    const s = p.addSlide({ masterName: "BG" });
    T_(s, { t: pg.kicker || "", x: 0.62, y: 0.5, w: 8, h: 0.3, size: 11, color: T.gold, bold: true });
    T_(s, { t: pg.title, x: 0.56, y: 1.55, w: 6.2, h: 1.3, size: 50, color: T.ink, bold: true });
    if (pg.subtitle) T_(s, { t: pg.subtitle, x: 0.62, y: 3.0, w: 6.2, h: 0.4, size: 14 });
    if (pg.strip) strip(s, 0.62, 3.75, pg.strip, 1.55);
    if (pg.tagline) T_(s, { t: pg.tagline, x: 0.62, y: 4.6, w: 6.2, h: 0.5, size: 15, color: T.gold, bold: true });
    if (pg.image) { const b = fit(pg.image, 5.85, 5.44); imgCard(s, pg.image, 6.9, 1.28, b.w, b.h); }
  },
  intro(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    if (pg.image) {
      const b = fit(pg.image, 3.55, 5.0);
      imgCard(s, pg.image, 0.62, 1.42, b.w, b.h);
      if (pg.image_caption) T_(s, { t: pg.image_caption, x: 0.62, y: 1.42 + b.h + 0.1, w: 3.55, h: 0.3, size: 9.5, color: T.muted });
    }
    const pitch = autoPitch(pg.rows.length, 2.0, 6.7, pg.rows.length >= 4 ? 1.16 : 1.5);
    pg.rows.forEach((r, i) => row(s, 4.65, 2.0 + i * pitch, 8.1, r[0], r[1], i < pg.rows.length - 1, pitch));
    footer(s, pg.footer);
  },
  product(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const imgs = pg.images.map(x => (typeof x === "string" ? { src: x } : x));
    if (imgs.length === 2) {
      imgs.forEach((im, i) => {
        const b = fit(im.src, 2.24, 2.08);
        imgCard(s, im.src, 0.62 + i * 2.36, 2.35, b.w, b.h);
        if (im.label) T_(s, { t: im.label, x: 0.62 + i * 2.36, y: 2.35 + b.h + 0.1, w: 2.24, h: 0.3, size: 10.5, color: T.muted, align: "center" });
      });
    } else {
      const b = fit(imgs[0].src, 4.6, 4.6);
      imgCard(s, imgs[0].src, 0.62, 1.42, b.w, b.h);
    }
    const startY = pg.ref ? 3.5 : (imgs.length === 2 ? 1.95 : 1.95);
    const pitch = autoPitch(pg.rows.length, startY, 6.7, pg.ref ? 1.1 : (pg.rows.length > 3 ? 1.24 : 1.5));
    pg.rows.forEach((r, i) => row(s, 5.62, startY + i * pitch, 7.13, r[0], r[1], i < pg.rows.length - 1, pitch));
    if (pg.ref) refBox(s, pg.ref.src, 5.62, 1.5, 2.2, pg.ref.caption);
    footer(s, pg.footer);
  },
  scene(pg) {
    const s = p.addSlide({ masterName: "BG" });
    s.addImage({ path: path.join(ASSETS, pg.image), x: 0, y: 0.028, w: 13.333, h: 7.444 });
    if (pg.chip) s.addText(pg.chip, { shape: "roundRect", rectRadius: 0.08,
      fill: { color: clean(T.bg), transparency: 20 }, x: 0.62, y: 6.6, w: 3.9, h: 0.55,
      fontFace: T.font, fontSize: 14, color: T.gold, bold: true, align: "center", valign: "middle" });
    if (pg.title_line) {
      s.addShape("rect", { x: 0, y: 0, w: 13.333, h: 1.4, fill: { color: clean(T.bg), transparency: 32 } });
      T_(s, { t: pg.title_line, x: 0.7, y: 0.5, w: 10.5, h: 0.6, size: 24, color: "FFFFFF", bold: true });
    }
  },
  symbols(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const symPitch = autoPitch(pg.rows.length, 1.6, 6.7, 1.06);
    pg.rows.forEach((r, i) => row(s, 0.62, 1.6 + i * symPitch, 7.5, r[0], r[1], i < pg.rows.length - 1, symPitch));
    (pg.images || []).slice(0, 2).forEach((f, i) => {
      const b = fit(f, 2.6, 2.42);
      imgCard(s, f, 8.5, 1.5 + i * 2.85, b.w, b.h);
      if (pg.fig_caps && pg.fig_caps[i]) T_(s, { t: pg.fig_caps[i], x: 8.5, y: 1.5 + i * 2.85 + b.h + 0.08, w: 2.6, h: 0.26, size: 9.5, color: T.muted });
    });
    footer(s, pg.footer);
  },
  colors(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    (pg.swatches || []).forEach((c, i) => {
      s.addShape("rect", { x: 0.62, y: 1.55 + i * 1.3, w: 4.4, h: 1.05, fill: { color: clean(c[1]) }, line: { color: clean(T.hair), width: 0.75 } });
      T_(s, { t: c[0] + "  " + c[1], x: 0.92, y: 1.9 + i * 1.3, w: 3.8, h: 0.4, size: 15, color: "FFFFFF", bold: true });
    });
    const colPitch = autoPitch(pg.rows.length, 1.7, 5.6, 1.35);
    pg.rows.forEach((r, i) => row(s, 5.5, 1.7 + i * colPitch, 7.25, r[0], r[1], i < pg.rows.length - 1, colPitch));
    if (pg.ratio) {
      const total = pg.ratio.reduce((a, g) => a + g[2], 0);
      let sx = 5.5;
      pg.ratio.forEach(g => {
        const w = 7.2 * g[2] / total;
        s.addShape("rect", { x: sx, y: 5.95, w, h: 0.42, fill: { color: clean(g[1]) }, line: { color: clean(T.hair), width: 0.75 } });
        T_(s, { t: g[0] + " " + g[2] + "%", x: sx + 0.06, y: 6.03, w: Math.max(w - 0.1, 0.3), h: 0.26, size: 9.5, color: T.ink, bold: true });
        sx += w + 0.06;
      });
      T_(s, { t: pg.ratio_note || "系列用色比例", x: 5.5, y: 6.5, w: 7.2, h: 0.3, size: 10, color: T.muted });
    }
    footer(s, pg.footer);
  },
  finale(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const n = pg.thumbs.length;
    const each = Math.min(2.51, 12.39 / n);
    const x0 = (13.333 - (each * n + 0.16 * (n - 1))) / 2;
    pg.thumbs.forEach((f, i) => {
      const b = fit(f, each, 2.19);
      imgCard(s, f, x0 + i * (each + 0.16), 1.5, b.w, b.h);
    });
    (pg.quote || []).forEach((q, i) => T_(s, { t: q, x: 0.62, y: 4.35 + i * 0.65, w: 12.1, h: 0.55, size: 20, color: i ? T.gold : T.ink, bold: true }));
    if (pg.strip) strip(s, 0.62, 5.95, pg.strip, 1.55);
    footer(s, pg.footer);
  },

  // ======================= 二期：新增 18 种 =======================
  // 1) 目录页：两栏编号列表，可高亮当前章节
  toc(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const rows = pg.rows || [];
    const half = Math.ceil(rows.length / 2);
    const colW = 5.72, x0 = 0.62, x1 = 6.99, y0 = 1.72;
    const pitch = Math.min(0.88, 4.5 / Math.max(half, 1));
    rows.forEach((r, i) => {
      const col = i < half ? 0 : 1;
      const k = i < half ? i : i - half;
      const x = col ? x1 : x0;
      const y = y0 + k * pitch;
      const on = pg.current && (r[0] === pg.current);
      if (on) s.addShape("rect", { x: x - 0.16, y: y + 0.02, w: 0.06, h: pitch - 0.24, fill: { color: clean(T.gold) } });
      T_(s, { t: pad2(i + 1), x, y, w: 0.72, h: 0.36, size: 15, color: on ? T.gold : T.muted, bold: true });
      T_(s, { t: r[0], x: x + 0.76, y: y - 0.02, w: colW - 0.9, h: 0.36, size: 15.5, color: on ? T.ink : T.sub, bold: true });
      if (r[1]) T_(s, { t: r[1], x: x + 0.76, y: y + 0.32, w: colW - 0.9, h: 0.3, size: 10, color: T.muted });
      if (!on) s.addShape("rect", { x, y: y + pitch - 0.17, w: colW, h: 0.008, fill: { color: clean(T.hair) } });
    });
    footer(s, pg.footer);
  },
  // 2) 章节过渡页：大序号 + 大标题，可整页背景图
  section(pg) {
    const s = p.addSlide({ masterName: "BG" });
    if (pg.image) {
      s.addImage({ path: path.join(ASSETS, pg.image), x: 0, y: 0, w: 13.333, h: 7.5 });
      s.addShape("rect", { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: clean(T.bg), transparency: 42 } });
    } else {
      s.addShape("rect", { x: 0, y: 0, w: 0.22, h: 7.5, fill: { color: clean(T.red) } });
    }
    const hasB = !!pg.badge;
    if (hasB) T_(s, { t: pg.badge, x: 0.62, y: 2.05, w: 2.1, h: 1.3, size: 66, color: T.red, bold: true });
    const tx = hasB ? 2.75 : 0.9;
    T_(s, { t: pg.title, x: tx, y: 2.3, w: 13.333 - tx - 0.62, h: 1.0, size: 40, color: T.ink, bold: true });
    s.addShape("rect", { x: tx + 0.04, y: 3.5, w: 1.25, h: 0.06, fill: { color: clean(T.gold) } });
    if (pg.subtitle) T_(s, { t: pg.subtitle, x: tx + 0.04, y: 3.78, w: 10.5, h: 0.6, size: 14, color: T.sub });
    footer(s, pg.footer);
  },
  // 3) 要点页：一列要点（可选配图）
  bullets(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const rows = pg.rows || [];
    let x = 0.62, w = 12.09;
    if (pg.image) {
      const b = fit(pg.image, 3.9, 4.6);
      imgCard(s, pg.image, 8.8 + (3.9 - b.w) / 2, 1.62, b.w, b.h);
      if (pg.image_caption) T_(s, { t: pg.image_caption, x: 8.8, y: 1.62 + b.h + 0.1, w: 3.9, h: 0.3, size: 9.5, color: T.muted, align: "center" });
      w = 7.6;
    }
    const pitch = autoPitch(rows.length, 1.75, 6.7, rows.length >= 4 ? 1.13 : 1.45);
    rows.forEach((r, i) => row(s, x, 1.75 + i * pitch, w, r[0], r[1], i < rows.length - 1, pitch));
    footer(s, pg.footer);
  },
  // 4) 图文页：左图右文 / 右图左文（image_side: left|right）
  imageText(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const right = (pg.image_side || "right") !== "left";
    const imgX = right ? 8.05 : 0.62, txtX = right ? 0.62 : 5.55;
    const bw = right ? 4.66 : 4.55;
    const b = fit(pg.image, bw, 4.5);
    imgCard(s, pg.image, imgX + (bw - b.w) / 2, 1.62, b.w, b.h);
    if (pg.image_caption) T_(s, { t: pg.image_caption, x: imgX, y: 1.62 + b.h + 0.1, w: bw, h: 0.3, size: 9.5, color: T.muted, align: "center" });
    const tw = right ? 7.15 : 7.16;
    const pitch = autoPitch(pg.rows.length, 1.72, 6.7, pg.rows.length >= 4 ? 1.1 : 1.42);
    pg.rows.forEach((r, i) => row(s, txtX, 1.72 + i * pitch, tw, r[0], r[1], i < pg.rows.length - 1, pitch));
    footer(s, pg.footer);
  },
  // 5) 卡片页：2-4 张并列卡片（优势 / 板块 / 卖点）
  cards(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const cs = pg.cards || [];
    const n = Math.max(cs.length, 1);
    const gap = 0.28, cw = (12.09 - gap * (n - 1)) / n;
    const y = 1.78, ch = 4.3;
    cs.forEach((raw, i) => {
      const c = pair(raw, "title", "body", "icon");
      const x = 0.62 + i * (cw + gap);
      panel(s, x, y, cw, ch, CARD, T.hair);
      s.addShape("rect", { x, y, w: cw, h: 0.06, fill: { color: clean(i % 2 ? T.red : T.gold) } });
      if (c.icon) T_(s, { t: c.icon, x: x + 0.3, y: y + 0.42, w: cw - 0.6, h: 0.8, size: 30, color: T.gold, bold: true });
      else T_(s, { t: pad2(i + 1), x: x + 0.3, y: y + 0.42, w: cw - 0.6, h: 0.7, size: 28, color: T.muted, bold: true });
      T_(s, { t: c.title, x: x + 0.3, y: y + 1.32, w: cw - 0.6, h: 0.5, size: 15.5, color: T.ink, bold: true });
      T_(s, { t: c.body, x: x + 0.3, y: y + 1.88, w: cw - 0.6, h: ch - 2.1, size: fitSize(c.body, cw - 0.6, ch - 2.15, 11), color: T.sub });
    });
    footer(s, pg.footer);
  },
  // 6) 对比页：2-3 栏对照（优缺点 / 方案 A/B / 前后）
  compare(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const cols = pg.columns || [];
    const n = Math.max(cols.length, 1);
    const gap = 0.34, cw = (12.09 - gap * (n - 1)) / n;
    const y = 1.85, ch = 4.25;
    cols.forEach((raw, i) => {
      const c = pair(raw, "title", "points");
      const x = 0.62 + i * (cw + gap);
      const tone = c.tone === "down" ? T.red : (c.tone === "neutral" ? T.muted : T.gold);
      panel(s, x, y, cw, ch, CARD, T.hair);
      s.addShape("rect", { x, y, w: cw, h: 0.54, fill: { color: clean(tone) } });
      T_(s, { t: c.title, x: x + 0.2, y: y + 0.11, w: cw - 0.4, h: 0.34, size: 14, color: "FFFFFF", bold: true, align: "center" });
      const pts = c.points || [];
      const pitch = Math.min(0.68, (ch - 0.9) / Math.max(pts.length, 1));
      pts.forEach((pt, j) => {
        const py = y + 0.78 + j * pitch;
        dot(s, x + 0.34, py + 0.17, 0.11, tone);
        T_(s, { t: pt, x: x + 0.52, y: py, w: cw - 0.78, h: pitch - 0.06, size: fitSize(pt, cw - 0.78, pitch - 0.1, 11), color: T.sub });
      });
    });
    if (pg.note) T_(s, { t: pg.note, x: 0.62, y: 6.3, w: 12.09, h: 0.4, size: 10.5, color: T.muted });
    footer(s, pg.footer);
  },
  // 7) 时间轴：横向节点，上下交替
  timeline(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const nodes = (pg.nodes || []).map(v => pair(v, "time", "title", "desc"));
    const n = nodes.length;
    const y = 3.35, x0 = 0.95, x1 = 12.4;
    s.addShape("rect", { x: x0, y: y, w: x1 - x0, h: 0.03, fill: { color: clean(T.hair) } });
    nodes.forEach((nd, i) => {
      const t = n === 1 ? 0.5 : i / (n - 1);
      const cx = x0 + t * (x1 - x0);
      dot(s, cx, y + 0.015, 0.26, i % 2 ? T.red : T.gold);
      const bw = Math.min(3.5, ((x1 - x0) / Math.max(n - 1, 1)) * 0.9);
      const bx = Math.min(Math.max(cx - bw / 2, 0.62), 12.71 - bw);
      if (i % 2 === 0) {
        s.addShape("rect", { x: cx - 0.008, y: y - 0.6, w: 0.016, h: 0.5, fill: { color: clean(T.hair) } });
        T_(s, { t: nd.time, x: bx, y: y - 1.82, w: bw, h: 0.3, size: 12, color: T.gold, bold: true, align: "center" });
        T_(s, { t: nd.title, x: bx, y: y - 1.5, w: bw, h: 0.36, size: 13.5, color: T.ink, bold: true, align: "center" });
        T_(s, { t: nd.desc, x: bx, y: y - 1.12, w: bw, h: 0.5, size: 10, color: T.sub, align: "center" });
      } else {
        s.addShape("rect", { x: cx - 0.008, y: y + 0.16, w: 0.016, h: 0.5, fill: { color: clean(T.hair) } });
        T_(s, { t: nd.time, x: bx, y: y + 0.68, w: bw, h: 0.3, size: 12, color: T.gold, bold: true, align: "center" });
        T_(s, { t: nd.title, x: bx, y: y + 1.0, w: bw, h: 0.36, size: 13.5, color: T.ink, bold: true, align: "center" });
        T_(s, { t: nd.desc, x: bx, y: y + 1.38, w: bw, h: 0.5, size: 10, color: T.sub, align: "center" });
      }
    });
    footer(s, pg.footer);
  },
  // 8) 流程/步骤页：编号卡片 + 箭头
  process(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const steps = (pg.steps || []).map(v => pair(v, "name", "desc"));
    const n = Math.max(steps.length, 1);
    const gap = 0.34, cw = (12.09 - gap * (n - 1)) / n;
    const y = 2.3, ch = 3.2;
    steps.forEach((st, i) => {
      const x = 0.62 + i * (cw + gap);
      panel(s, x, y, cw, ch, CARD, T.hair);
      dot(s, x + cw / 2, y + 0.62, 0.58, i === 0 ? T.gold : T.red);
      T_(s, { t: String(i + 1), x: x + cw / 2 - 0.29, y: y + 0.46, w: 0.58, h: 0.34, size: 16, color: "FFFFFF", bold: true, align: "center" });
      T_(s, { t: st.name, x: x + 0.18, y: y + 1.14, w: cw - 0.36, h: 0.45, size: 14.5, color: T.ink, bold: true, align: "center" });
      T_(s, { t: st.desc, x: x + 0.18, y: y + 1.62, w: cw - 0.36, h: ch - 1.8, size: 10.5, color: T.sub, align: "center" });
      if (i < n - 1) T_(s, { t: "▶", x: x + cw + 0.02, y: y + 0.48, w: gap - 0.04, h: 0.4, size: 12, color: T.gold, align: "center" });
    });
    footer(s, pg.footer);
  },
  // 9) 四象限 / 矩阵（SWOT、重要性-紧迫性、业务划分）
  matrix(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const qs = (pg.quadrants || []).map(v => pair(v, "name", "desc"));
    const hasNote = !!pg.note;
    const gw = hasNote ? 7.0 : 8.4, gh = 4.3, gy = 1.72;
    const gx = hasNote ? 0.9 : (13.333 - gw) / 2;
    const cw = (gw - 0.22) / 2, ch = (gh - 0.22) / 2;
    for (let i = 0; i < 4; i++) {
      const r = Math.floor(i / 2), c = i % 2;
      const x = gx + c * (cw + 0.22), y = gy + r * (ch + 0.22);
      panel(s, x, y, cw, ch, i === 0 ? lighten(T.bg, 0.16) : CARD, T.hair);
      s.addShape("rect", { x, y, w: 0.06, h: ch, fill: { color: clean(i % 3 === 0 ? T.gold : (i % 3 === 1 ? T.red : T.muted)) } });
      const q = qs[i] || {};
      T_(s, { t: q.name || "", x: x + 0.24, y: y + 0.18, w: cw - 0.45, h: 0.4, size: 14.5, color: T.ink, bold: true });
      T_(s, { t: q.desc || "", x: x + 0.24, y: y + 0.62, w: cw - 0.45, h: ch - 0.8, size: 10.5, color: T.sub });
    }
    s.addShape("rect", { x: gx + gw / 2 - 0.01, y: gy - 0.05, w: 0.02, h: gh + 0.1, fill: { color: clean(T.gold) } });
    s.addShape("rect", { x: gx - 0.05, y: gy + gh / 2 - 0.01, w: gw + 0.1, h: 0.02, fill: { color: clean(T.gold) } });
    T_(s, { t: pg.x_label || "", x: gx, y: gy + gh + 0.12, w: gw, h: 0.35, size: 11, color: T.gold, align: "center", bold: true });
    T_(s, { t: pg.y_label || "", x: 0.62, y: gy + gh / 2 - 0.4, w: 1.15, h: 0.8, size: 11, color: T.gold, align: "center", bold: true });
    if (hasNote) {
      const nx = gx + gw + 0.42;
      panel(s, nx, gy, 12.71 - nx, gh, CARD, T.hair);
      T_(s, { t: pg.note, x: nx + 0.26, y: gy + 0.22, w: 12.71 - nx - 0.5, h: gh - 0.5, size: 10.5, color: T.sub });
    }
    footer(s, pg.footer);
  },
  // 10) 金字塔 / 层级（自下而上）
  pyramid(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const lv = (pg.levels || []).map(v => pair(v, "name", "desc"));
    const n = Math.max(lv.length, 1);
    const baseY = 6.25, h = Math.min(1.0, 4.4 / n);
    const cx = 4.9;
    const wMax = 6.2, wMin = 2.3;
    lv.forEach((l, i) => {
      const t = n === 1 ? 0 : i / (n - 1);
      const w = wMax - (wMax - wMin) * t;
      const y = baseY - h * (i + 1);
      const color = i % 2 ? T.red : T.gold;
      s.addShape("rect", { x: cx - w / 2, y, w, h: h - 0.08, fill: { color: clean(color) } });
      T_(s, { t: l.name, x: cx - w / 2, y: y + (h - 0.08) / 2 - 0.16, w, h: 0.34, size: 13.5, color: "FFFFFF", bold: true, align: "center" });
      T_(s, { t: l.desc, x: cx + wMax / 2 + 0.35, y: y + (h - 0.08) / 2 - 0.2, w: 12.71 - (cx + wMax / 2 + 0.35), h: 0.5, size: 10.5, color: T.sub });
    });
    T_(s, { t: pg.base_note || "", x: cx - wMax / 2, y: baseY + 0.08, w: wMax, h: 0.3, size: 10, color: T.muted, align: "center" });
    footer(s, pg.footer);
  },
  // 11) 数据看板：2-4 个大数字
  stats(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const st = (pg.stats || []).map(v => pair(v, "value", "label", "note"));
    const n = Math.max(st.length, 1);
    const cw = 12.09 / n;
    st.forEach((it, i) => {
      const x = 0.62 + i * cw;
      if (i) s.addShape("rect", { x: x - 0.01, y: 2.25, w: 0.012, h: 2.6, fill: { color: clean(T.hair) } });
      T_(s, { t: it.value, x: x + 0.15, y: 2.2, w: cw - 0.3, h: 1.05, size: 44, color: T.gold, bold: true, align: "center" });
      T_(s, { t: it.label, x: x + 0.15, y: 3.35, w: cw - 0.3, h: 0.4, size: 14, color: T.ink, bold: true, align: "center" });
      T_(s, { t: it.note, x: x + 0.15, y: 3.82, w: cw - 0.3, h: 0.8, size: 10.5, color: T.sub, align: "center" });
    });
    if (pg.source) T_(s, { t: pg.source, x: 0.62, y: 5.5, w: 12.09, h: 0.35, size: 10, color: T.muted, align: "center" });
    footer(s, pg.footer);
  },
  // 12) 图表页：bar / line / pie / doughnut / area
  chart(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const c = pg.chart || {};
    const data = (c.series || []).map(sr => ({ name: sr.name || "", labels: c.categories || [], values: sr.values || [] }));
    const hasNote = !!pg.note;
    s.addChart(c.type || "bar", data, {
      x: 0.62, y: 1.5, w: hasNote ? 8.4 : 12.09, h: 5.0,
      barDir: c.bar_dir || "col",
      chartColors: (c.colors || [T.gold, T.red, T.sub, T.muted, T.ink, "6E8B7B"]).map(clean),
      showLegend: data.length > 1 || (c.type === "pie" || c.type === "doughnut"),
      legendPos: c.legend_pos || "b", legendColor: T.sub, legendFontFace: T.font, legendFontSize: 10,
      catAxisLabelColor: T.sub, catAxisLabelFontFace: T.font, catAxisLabelFontSize: 10,
      valAxisLabelColor: T.muted, valAxisLabelFontFace: T.font, valAxisLabelFontSize: 9,
      valGridLine: { color: clean(T.hair), style: "solid", size: 0.5 },
      catGridLine: { style: "none" },
      showValue: !!c.show_value, dataLabelColor: T.ink, dataLabelFontFace: T.font, dataLabelFontSize: 9,
      barGapWidthPct: 45, lineSize: 2.5, lineDataSymbol: "circle",
      chartArea: { fill: { color: clean(T.bg) } }, plotArea: { fill: { color: clean(T.bg) } },
    });
    if (hasNote) {
      panel(s, 9.25, 1.5, 3.46, 5.0, CARD, T.hair);
      T_(s, { t: Array.isArray(pg.note) ? "结论" : "结论", x: 9.5, y: 1.72, w: 3.0, h: 0.35, size: 13, color: T.gold, bold: true });
      T_(s, { t: pg.note, x: 9.5, y: 2.15, w: 3.0, h: 4.1, size: 11, color: T.sub });
    }
    footer(s, pg.footer);
  },
  // 13) 表格页
  table(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const head = (pg.head || []).map(h => ({ text: txt(h), options: { bold: true, color: clean(T.bg), fill: clean(T.gold), fontFace: T.font, fontSize: 12, align: "center", valign: "middle" } }));
    const rows = (pg.rows || []).map((r, i) => r.map(cl => ({ text: txt(cl), options: { color: T.sub, fill: i % 2 ? CARD : clean(T.bg), fontFace: T.font, fontSize: 11, valign: "middle" } })));
    const colW = pg.col_widths || new Array(Math.max(head.length, 1)).fill(12.09 / Math.max(head.length, 1));
    s.addTable([head].concat(rows), { x: 0.62, y: 1.62, w: 12.09, colW,
      rowH: pg.row_h || 0.4, border: { type: "solid", color: clean(T.hair), pt: 0.5 }, autoPage: false, valign: "middle" });
    if (pg.note) T_(s, { t: pg.note, x: 0.62, y: 6.3, w: 12.09, h: 0.4, size: 10.5, color: T.muted });
    footer(s, pg.footer);
  },
  // 14) 金句 / 引用页
  quote(pg) {
    const s = p.addSlide({ masterName: "BG" });
    if (pg.image) {
      s.addImage({ path: path.join(ASSETS, pg.image), x: 0, y: 0, w: 13.333, h: 7.5 });
      s.addShape("rect", { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: clean(T.bg), transparency: 38 } });
    }
    T_(s, { t: "\u201C", x: 0.7, y: 0.7, w: 3, h: 1.7, size: 96, color: T.red, bold: true });
    const lines = pg.quote || [];
    lines.forEach((q, i) => T_(s, { t: q, x: 1.75, y: 2.2 + i * 0.85, w: 10.2, h: 0.75, size: 26, color: i ? T.gold : T.ink, bold: true }));
    if (pg.source) T_(s, { t: "—— " + pg.source, x: 1.78, y: 2.45 + lines.length * 0.85, w: 10.2, h: 0.45, size: 13, color: T.muted });
    footer(s, pg.footer);
  },
  // 15) 团队 / 人物页
  team(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const ms = (pg.images || []).map(x => (typeof x === "string" ? { src: x } : x));
    const n = Math.max(ms.length, 1);
    const gap = 0.3, cw = (12.09 - gap * (n - 1)) / n;
    ms.forEach((m, i) => {
      const x = 0.62 + i * (cw + gap);
      const b = fit(m.src, cw, 2.95);
      imgCard(s, m.src, x + (cw - b.w) / 2, 1.72, b.w, b.h);
      T_(s, { t: m.label || "", x, y: 4.85, w: cw, h: 0.4, size: 14.5, color: T.ink, bold: true, align: "center" });
      T_(s, { t: m.sub || "", x, y: 5.28, w: cw, h: 0.35, size: 11, color: T.gold, align: "center" });
      if (m.note) T_(s, { t: m.note, x, y: 5.68, w: cw, h: 0.7, size: 10, color: T.sub, align: "center" });
    });
    footer(s, pg.footer);
  },
  // 16) 图集 / 网格
  gallery(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const ms = (pg.images || []).map(x => (typeof x === "string" ? { src: x } : x));
    const n = Math.max(ms.length, 1);
    const cols = n <= 3 ? n : (n === 4 ? 2 : 3);
    const rows = Math.ceil(n / cols);
    const gap = 0.24;
    const cw = (12.09 - gap * (cols - 1)) / cols;
    const chh = (4.95 - gap * (rows - 1)) / rows;
    const caps = pg.captions || [];
    ms.forEach((m, i) => {
      const r = Math.floor(i / cols), c = i % cols;
      const x = 0.62 + c * (cw + gap), y = 1.62 + r * (chh + gap);
      const cap = m.label || caps[i] || "";
      const box = chh - (cap ? 0.32 : 0);
      const b = fit(m.src, cw, box);
      imgCard(s, m.src, x + (cw - b.w) / 2, y + (box - b.h) / 2, b.w, b.h);
      if (cap) T_(s, { t: cap, x, y: y + chh - 0.3, w: cw, h: 0.28, size: 9.5, color: T.muted, align: "center" });
    });
    footer(s, pg.footer);
  },
  // 17) 案例页：左图 + 右文 + 底部数据条
  case(pg) {
    const s = p.addSlide({ masterName: "BG" });
    header(s, pg.badge, pg.title, pg.pill);
    const b = fit(pg.image, 5.1, 3.45);
    imgCard(s, pg.image, 0.62, 1.6, b.w, b.h);
    if (pg.image_caption) T_(s, { t: pg.image_caption, x: 0.62, y: 1.6 + b.h + 0.1, w: 5.1, h: 0.3, size: 9.5, color: T.muted });
    const pitch = autoPitch(pg.rows.length, 1.62, 5.35, pg.rows.length >= 4 ? 0.95 : 1.15);
    pg.rows.forEach((r, i) => row(s, 6.2, 1.62 + i * pitch, 6.51, r[0], r[1], i < pg.rows.length - 1, pitch));
    const mt = pg.metrics || [];
    if (mt.length) {
      s.addShape("rect", { x: 0.62, y: 5.45, w: 12.09, h: 0.014, fill: { color: clean(T.hair) } });
      const mw = 12.09 / mt.length;
      mt.forEach((raw, i) => {
        const m = pair(raw, "value", "label");
        T_(s, { t: m.value, x: 0.62 + i * mw, y: 5.62, w: mw, h: 0.6, size: 26, color: T.gold, bold: true, align: "center" });
        T_(s, { t: m.label, x: 0.62 + i * mw, y: 6.22, w: mw, h: 0.35, size: 10.5, color: T.sub, align: "center" });
      });
    }
    footer(s, pg.footer);
  },
  // 18) 结尾联系页
  contact(pg) {
    const s = p.addSlide({ masterName: "BG" });
    s.addShape("rect", { x: 0, y: 0, w: 0.22, h: 7.5, fill: { color: clean(T.red) } });
    const hasImg = !!pg.image;
    const tw = hasImg ? 8.2 : 12.09;
    T_(s, { t: pg.title, x: 0.9, y: 1.95, w: tw, h: 1.1, size: 42, color: T.ink, bold: true });
    s.addShape("rect", { x: 0.94, y: 3.2, w: 1.3, h: 0.06, fill: { color: clean(T.gold) } });
    if (pg.slogan) T_(s, { t: pg.slogan, x: 0.94, y: 3.48, w: tw, h: 0.5, size: 16, color: T.gold, bold: true });
    (pg.lines || []).forEach((ln, i) => T_(s, { t: ln, x: 0.94, y: 4.35 + i * 0.42, w: tw, h: 0.36, size: 12, color: T.sub }));
    if (hasImg) {
      const b = fit(pg.image, 2.3, 2.3);
      s.addImage({ path: path.join(ASSETS, pg.image), x: 10.4 + (2.3 - b.w) / 2, y: 2.5, w: b.w, h: b.h });
      if (pg.image_caption) T_(s, { t: pg.image_caption, x: 10.0, y: 2.5 + b.h + 0.12, w: 3.1, h: 0.3, size: 9.5, color: T.muted, align: "center" });
    }
    footer(s, pg.footer);
  },
};

(spec.pages || []).forEach(pg => { if (R[pg.type]) R[pg.type](pg); });
p.writeFile({ fileName: OUT }).then(() => console.log("WROTE " + OUT + "  pages=" + spec.pages.length));
