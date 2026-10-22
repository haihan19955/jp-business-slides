// jpdeck.js — 日本のビジネス資料（コンサル型）向け pptxgenjs レイアウト集
// 使い方: const { createDeck } = require("<skill>/scripts/jpdeck.js");
//         const deck = createDeck({ font: "Meiryo" }); deck.cover({...}); ... await deck.save("out.pptx");
"use strict";
const PptxGenJS = require("pptxgenjs");

const W = 13.333, H = 7.5, MX = 0.5;
const CW = W - MX * 2;

const PALETTES = {
  navy:   { primary: "1B2A49", accent: "D9822B", text: "222222", sub: "5B6770", line: "C9CFD6", tint: "EEF1F5", tintAccent: "FBEFE2" },
  green:  { primary: "1E4D3A", accent: "C8553D", text: "222222", sub: "5B6770", line: "C9CFD6", tint: "EDF3EF", tintAccent: "F8E9E5" },
  wine:   { primary: "5A1E2E", accent: "2F6690", text: "222222", sub: "5B6770", line: "CFC9CB", tint: "F4EEF0", tintAccent: "E6EEF5" },
  mono:   { primary: "2B2B2B", accent: "C0392B", text: "222222", sub: "666666", line: "CCCCCC", tint: "F2F2F2", tintAccent: "F9E5E3" },
};

function createDeck(opts = {}) {
  const font = opts.font || "Meiryo";
  const P = PALETTES[opts.palette || "navy"];
  const footerLeft = opts.footer || "";
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: font, bodyFontFace: font };
  if (opts.title) pres.title = opts.title;
  if (opts.author) pres.author = opts.author;

  const titleOpts = {
    x: MX, y: 0.3, w: CW, h: 0.95, fontFace: font, fontSize: 22, bold: true,
    color: P.primary, align: "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.1,
  };
  pres.defineSlideMaster({
    title: "JP_CONTENT",
    background: { color: "FFFFFF" },
    margin: [0.5, 0.5, 0.75, 0.5],
    objects: [
      { placeholder: { options: { name: "title", type: "title", ...titleOpts }, text: "" } },
      ...(footerLeft ? [{ text: { text: footerLeft, options: { x: 6.2, y: 7.05, w: 5.6, h: 0.3, fontFace: font, fontSize: 9, color: P.sub, align: "right", margin: 0 } } }] : []),
    ],
    slideNumber: { x: 12.2, y: 7.05, w: 0.63, h: 0.3, fontFace: font, fontSize: 9, color: P.sub, align: "right" },
  });
  pres.defineSlideMaster({ title: "JP_PLAIN", background: { color: "FFFFFF" }, objects: [] });

  const T = (o) => ({ fontFace: font, color: P.text, margin: 0, isTextBox: true, ...o });

  // 共通: タイトル（＝メッセージ）・リード・結論ボックス・出所
  function content({ title, lead, bottomLine, source, notes }) {
    const s = pres.addSlide({ masterName: "JP_CONTENT" });
    s.addText(title, { placeholder: "title" });
    let top = 1.4, bottom = 6.85;
    if (lead) {
      s.addText(lead, T({ x: MX, y: 1.3, w: CW, h: 0.5, fontSize: 14, color: P.sub, valign: "top", objectName: "lead" }));
      top = 1.95;
    }
    if (bottomLine) {
      s.addShape(pres.shapes.RECTANGLE, { x: MX, y: 6.0, w: CW, h: 0.8, fill: { color: P.tintAccent }, line: { type: "none" } });
      s.addText([{ text: "結論　", options: { color: P.accent, bold: true } }, { text: bottomLine, options: { color: P.primary, bold: true } }],
        T({ x: MX + 0.3, y: 6.0, w: CW - 0.6, h: 0.8, fontSize: 16, valign: "middle", objectName: "bottomLine" }));
      bottom = 5.8;
    }
    if (source) s.addText(source, T({ x: MX, y: 7.05, w: 5.6, h: 0.3, fontSize: 9, color: P.sub, objectName: "source" }));
    if (notes) s.addNotes(notes);
    return { s, top, bottom };
  }

  const api = {
    pres, P, font,

    cover({ title, subtitle, date, org }) {
      const s = pres.addSlide({ masterName: "JP_PLAIN" });
      s.background = { color: P.primary };
      s.addText(title, T({ x: 0.9, y: 2.2, w: 11.5, h: 1.6, fontSize: 34, bold: true, color: "FFFFFF", valign: "bottom", lineSpacingMultiple: 1.15 }));
      if (subtitle) s.addText(subtitle, T({ x: 0.9, y: 3.95, w: 11.5, h: 0.7, fontSize: 18, color: "D6DCE4", valign: "top" }));
      const meta = [date, org].filter(Boolean).join("　｜　");
      if (meta) s.addText(meta, T({ x: 0.9, y: 6.2, w: 11.5, h: 0.4, fontSize: 12, color: "D6DCE4" }));
      return s;
    },

    section({ title, sub }) {
      const s = pres.addSlide({ masterName: "JP_PLAIN" });
      s.background = { color: P.tint };
      s.addText(title, T({ x: 0.9, y: 2.9, w: 11.5, h: 1.0, fontSize: 30, bold: true, color: P.primary }));
      if (sub) s.addText(sub, T({ x: 0.9, y: 3.95, w: 11.5, h: 0.6, fontSize: 16, color: P.sub }));
      return s;
    },

    // エグゼクティブサマリー: 番号付きの要点（見出し＋補足）
    summary({ title, lead, points, source, notes }) {
      const { s, top, bottom } = content(arguments[0]);
      const n = points.length, gap = 0.18;
      const h = Math.min(1.25, (bottom - top - gap * (n - 1)) / n);
      points.forEach((p, i) => {
        const y = top + i * (h + gap);
        s.addShape(pres.shapes.RECTANGLE, { x: MX, y, w: CW, h, fill: { color: P.tint }, line: { type: "none" } });
        s.addText(String(i + 1), T({ x: MX + 0.2, y, w: 0.6, h, fontSize: 26, bold: true, color: P.accent, align: "center", valign: "middle" }));
        s.addText([
          { text: p.head, options: { fontSize: 16, bold: true, color: P.primary, breakLine: true } },
          ...(p.body ? [{ text: p.body, options: { fontSize: 13, color: P.text } }] : []),
        ], T({ x: MX + 1.0, y: y + 0.1, w: CW - 1.25, h: h - 0.2, valign: "middle", paraSpaceAfter: 2 }));
      });
      return s;
    },

    // 現状 → 課題 → 打ち手（行ごとに横並び）
    issueSolution({ title, lead, headers = ["現状", "課題", "打ち手"], rows, source, notes }) {
      const { s, top, bottom } = content(arguments[0]);
      const colW = (CW - 0.6) / 3, arrowW = 0.3, headH = 0.45, gap = 0.15;
      headers.forEach((t, c) => {
        const x = MX + c * (colW + arrowW);
        s.addShape(pres.shapes.RECTANGLE, { x, y: top, w: colW, h: headH, fill: { color: c === 2 ? P.accent : P.primary }, line: { type: "none" } });
        s.addText(t, T({ x, y: top, w: colW, h: headH, fontSize: 14, bold: true, color: "FFFFFF", align: "center", valign: "middle" }));
      });
      const y0 = top + headH + gap;
      const rh = (bottom - y0 - gap * (rows.length - 1)) / rows.length;
      rows.forEach((r, i) => {
        const y = y0 + i * (rh + gap);
        r.forEach((cell, c) => {
          const x = MX + c * (colW + arrowW);
          s.addShape(pres.shapes.RECTANGLE, { x, y, w: colW, h: rh, fill: { color: c === 2 ? P.tintAccent : P.tint }, line: { type: "none" } });
          s.addText(cell, T({ x: x + 0.15, y: y + 0.1, w: colW - 0.3, h: rh - 0.2, fontSize: 13, valign: "middle", bold: c === 2 }));
          if (c < 2) s.addShape(pres.shapes.CHEVRON, { x: x + colW + 0.06, y: y + rh / 2 - 0.15, w: 0.18, h: 0.3, fill: { color: P.line }, line: { type: "none" } });
        });
      });
      return s;
    },

    // 比較表（推奨列をハイライト）
    compare({ title, lead, columns, rows, highlight, source, notes }) {
      const { s, top, bottom } = content(arguments[0]);
      const labelW = 2.3, colW = (CW - labelW) / columns.length;
      const head = [{ text: "", options: { fill: { color: "FFFFFF" } } }].concat(columns.map((c, i) => ({
        text: c, options: { bold: true, color: "FFFFFF", fill: { color: i === highlight ? P.accent : P.primary }, align: "center" },
      })));
      const body = rows.map((r) => [{ text: r.label, options: { bold: true, color: P.primary, fill: { color: P.tint } } }].concat(r.values.map((v, i) => ({
        text: v, options: { fill: { color: i === highlight ? P.tintAccent : "FFFFFF" }, bold: i === highlight },
      }))));
      const rowH = Math.min(0.85, (bottom - top) / (rows.length + 1));
      s.addTable([head, ...body], {
        x: MX, y: top, w: CW, colW: [labelW, ...columns.map(() => colW)], rowH,
        fontFace: font, fontSize: 13, color: P.text, valign: "middle", margin: [0.05, 0.12, 0.05, 0.12],
        border: { type: "solid", pt: 0.75, color: P.line },
      });
      return s;
    },

    // 2×2 マトリクス
    matrix({ title, lead, xAxis, yAxis, quadrants, highlight, source, notes }) {
      const { s, top, bottom } = content(arguments[0]);
      const ax = MX + 0.9, aw = CW - 1.1, ah = bottom - top - 0.55;
      const qw = aw / 2, qh = ah / 2;
      // quadrants 順序: [左上, 右上, 左下, 右下]
      quadrants.forEach((q, i) => {
        const x = ax + (i % 2) * qw, y = top + Math.floor(i / 2) * qh;
        const hi = i === highlight;
        s.addShape(pres.shapes.RECTANGLE, { x: x + 0.05, y: y + 0.05, w: qw - 0.1, h: qh - 0.1, fill: { color: hi ? P.tintAccent : P.tint }, line: { type: "none" } });
        s.addText([
          { text: q.head, options: { fontSize: 17, bold: true, color: hi ? P.accent : P.primary, breakLine: true } },
          ...(q.body ? [{ text: q.body, options: { fontSize: 13, color: P.text } }] : []),
        ], T({ x: x + 0.35, y: y + 0.2, w: qw - 0.7, h: qh - 0.4, valign: "middle", paraSpaceAfter: 6 }));
      });
      // 軸: 矢印は図形で描く（縦書きテキスト内の矢印は向きが変わるため）
      s.addShape(pres.shapes.LINE, { x: MX + 0.75, y: top, w: 0, h: ah, line: { color: P.sub, width: 1.25, beginArrowType: "triangle" } });
      s.addShape(pres.shapes.LINE, { x: ax, y: top + ah + 0.12, w: aw, h: 0, line: { color: P.sub, width: 1.25, endArrowType: "triangle" } });
      const noArrow = (t) => t.replace(/\s*[→←↑↓]\s*/g, "");
      s.addText(noArrow(yAxis), T({ x: MX, y: top, w: 0.6, h: ah, fontSize: 12, color: P.sub, align: "center", valign: "middle", vert: "eaVert" }));
      s.addText(noArrow(xAxis), T({ x: ax, y: top + ah + 0.2, w: aw, h: 0.3, fontSize: 12, color: P.sub, align: "center" }));
      return s;
    },

    // KPI カード（大きな数字＋単位＋説明）
    kpi({ title, lead, cards, bottomLine, source, notes }) {
      const { s, top, bottom } = content(arguments[0]);
      const n = cards.length, gap = 0.3, cw = (CW - gap * (n - 1)) / n, ch = bottomLine ? bottom - top : Math.min(bottom - top, 3.6);
      cards.forEach((c, i) => {
        const x = MX + i * (cw + gap);
        s.addShape(pres.shapes.RECTANGLE, { x, y: top, w: cw, h: ch, fill: { color: P.tint }, line: { type: "none" } });
        const y0 = top + Math.max(0.2, (ch - 3.2) / 2);
        s.addText(c.label, T({ x: x + 0.3, y: y0, w: cw - 0.6, h: 0.45, fontSize: 15, bold: true, color: P.primary }));
        s.addText([
          { text: c.value, options: { fontSize: 54, bold: true, color: P.accent } },
          { text: c.unit ? " " + c.unit : "", options: { fontSize: 20, bold: true, color: P.accent } },
        ], T({ x: x + 0.3, y: y0 + 0.6, w: cw - 0.6, h: 1.2, valign: "middle" }));
        if (c.note) s.addText(c.note, T({ x: x + 0.3, y: y0 + 2.0, w: cw - 0.6, h: 1.1, fontSize: 13, color: P.text, valign: "top" }));
      });
      return s;
    },

    // プロセス（番号付きステップを横に並べる）
    process({ title, lead, steps, bottomLine, source, notes }) {
      const { s, top, bottom } = content(arguments[0]);
      const n = steps.length, gap = 0.35, cw = (CW - gap * (n - 1)) / n, ch = bottomLine ? bottom - top : Math.min(bottom - top, 3.4);
      steps.forEach((st, i) => {
        const x = MX + i * (cw + gap);
        if (i < n - 1) s.addShape(pres.shapes.CHEVRON, { x: x + cw + 0.08, y: top + ch / 2 - 0.2, w: 0.2, h: 0.4, fill: { color: P.line }, line: { type: "none" } });
        s.addShape(pres.shapes.RECTANGLE, { x, y: top, w: cw, h: 0.55, fill: { color: P.primary }, line: { type: "none" } });
        s.addText(`STEP ${i + 1}`, T({ x, y: top, w: cw, h: 0.55, fontSize: 13, bold: true, color: "FFFFFF", align: "center", valign: "middle" }));
        s.addShape(pres.shapes.RECTANGLE, { x, y: top + 0.55, w: cw, h: ch - 0.55, fill: { color: P.tint }, line: { type: "none" } });
        s.addText([
          { text: st.head, options: { fontSize: 17, bold: true, color: P.primary, breakLine: true } },
          ...(st.body ? [{ text: st.body, options: { fontSize: 14, color: P.text } }] : []),
        ], T({ x: x + 0.25, y: top + 0.8, w: cw - 0.5, h: ch - 1.0, valign: "top", paraSpaceAfter: 8 }));
      });
      return s;
    },

    // ロードマップ（期間×トラックのガント）
    roadmap({ title, lead, periods, tracks, source, notes }) {
      const { s, top, bottom } = content(arguments[0]);
      const labelW = 2.2, pw = (CW - labelW) / periods.length, headH = 0.45;
      periods.forEach((p, i) => {
        const x = MX + labelW + i * pw;
        s.addShape(pres.shapes.RECTANGLE, { x: x + 0.02, y: top, w: pw - 0.04, h: headH, fill: { color: P.primary }, line: { type: "none" } });
        s.addText(p, T({ x, y: top, w: pw, h: headH, fontSize: 13, bold: true, color: "FFFFFF", align: "center", valign: "middle" }));
      });
      const th = Math.min(1.1, (bottom - top - headH - 0.1) / tracks.length);
      tracks.forEach((t, r) => {
        const y = top + headH + 0.1 + r * th;
        s.addShape(pres.shapes.LINE, { x: MX, y: y + th, w: CW, h: 0, line: { color: P.line, width: 0.75 } });
        s.addText(t.name, T({ x: MX, y, w: labelW - 0.15, h: th, fontSize: 13, bold: true, color: P.primary, valign: "middle" }));
        t.items.forEach((it) => {
          const x = MX + labelW + it.start * pw + 0.05, w = (it.end - it.start) * pw - 0.1;
          s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: y + 0.18, w, h: th - 0.36, rectRadius: 0.08, fill: { color: it.key ? P.accent : P.tint }, line: { type: "none" } });
          s.addText(it.label, T({ x: x + 0.1, y: y + 0.18, w: w - 0.2, h: th - 0.36, fontSize: 12, bold: !!it.key, color: it.key ? "FFFFFF" : P.text, valign: "middle" }));
        });
      });
      return s;
    },

    // 棒グラフ＋示唆（右側に「つまり何か」を書く）
    chartWithTakeaway({ title, lead, chart, takeaways, source, notes }) {
      const { s, top, bottom } = content(arguments[0]);
      const chartW = 7.9;
      s.addChart(pres.charts[chart.type || "BAR"], chart.data, {
        x: MX, y: top, w: chartW, h: bottom - top,
        // 系列が1つのときに複数色を渡すと棒ごとに色が変わってしまうので、1色だけ渡す
        barDir: chart.barDir || "col", chartColors: chart.colors || (chart.data.length > 1 ? [P.primary, P.accent, P.sub] : [P.primary]),
        showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: font, dataLabelFontSize: 11, dataLabelColor: P.text,
        catAxisLabelFontFace: font, catAxisLabelFontSize: 11, catAxisLabelColor: P.sub,
        valAxisLabelFontFace: font, valAxisLabelFontSize: 10, valAxisLabelColor: P.sub,
        valGridLine: { color: "E5E8EC", size: 0.5 }, catGridLine: { style: "none" },
        showLegend: chart.data.length > 1, legendPos: "t", legendFontFace: font, legendFontSize: 11,
        showTitle: !!chart.title, title: chart.title, titleFontFace: font, titleFontSize: 13, titleColor: P.primary,
      });
      // 単位は軸タイトル（横倒しになる）ではなく、グラフ左上に横書きで置く
      if (chart.unit) s.addText(`（${chart.unit}）`, T({ x: MX, y: top - 0.32, w: 1.5, h: 0.3, fontSize: 10, color: P.sub, objectName: "unit" }));
      const x = MX + chartW + 0.3, w = CW - chartW - 0.3;
      s.addShape(pres.shapes.RECTANGLE, { x, y: top, w, h: bottom - top, fill: { color: P.tint }, line: { type: "none" } });
      s.addText("示唆", T({ x: x + 0.3, y: top + 0.25, w: w - 0.6, h: 0.4, fontSize: 15, bold: true, color: P.accent }));
      s.addText(takeaways.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < takeaways.length - 1 } })),
        T({ x: x + 0.3, y: top + 0.8, w: w - 0.6, h: bottom - top - 1.0, fontSize: 15, valign: "top", paraSpaceAfter: 14 }));
      return s;
    },

    // ネクストステップ（何を・誰が・いつまでに）
    nextSteps({ title, lead, items, source, notes }) {
      const { s, top, bottom } = content(arguments[0]);
      const head = ["#", "実施事項", "担当", "期限"].map((t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: P.primary }, align: "center" } }));
      const body = items.map((it, i) => [
        { text: String(i + 1), options: { align: "center", bold: true, color: P.accent } },
        { text: it.what }, { text: it.who, options: { align: "center" } }, { text: it.when, options: { align: "center" } },
      ]);
      s.addTable([head, ...body], {
        x: MX, y: top, w: CW, colW: [0.6, CW - 0.6 - 2.4 - 2.0, 2.4, 2.0], rowH: Math.min(0.8, (bottom - top) / (items.length + 1)),
        fontFace: font, fontSize: 13, color: P.text, valign: "middle", margin: [0.05, 0.12, 0.05, 0.12],
        border: { type: "solid", pt: 0.75, color: P.line },
      });
      return s;
    },

    async save(fileName) {
      await pres.writeFile({ fileName });
      return fileName;
    },
  };
  return api;
}

module.exports = { createDeck, PALETTES };
