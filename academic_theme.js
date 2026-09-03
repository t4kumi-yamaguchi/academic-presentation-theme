/**
 * academic_theme.js
 * 学術発表・講義用 PowerPoint テンプレートモジュール(pptxgenjs)
 *
 * 使い方(AIへの指示例):
 *   「academic_theme.js の THEMES.navy を使って、以下の内容でスライドを生成して」
 *
 * const { THEMES, newDeck, addTitleSlide, addSectionSlide, addBulletSlide,
 *         addCardSlide, addChartStatSlide, addSummarySlide } = require('./academic_theme');
 *
 * 設計方針:
 *  - フォント: Meiryo UI(日本語学術用途)
 *  - 本文: 24pt 以上(講義スクリーン想定)。四角囲み(ポイント枠・カード)内も 24pt・行間 1.3
 *  - 本文色は必ず真っ黒(#000000)を使用する
 *  - 装飾は最小限: アクセントバー・下線は使わない。余白・背景色・円モチーフで区別
 *  - 16:9 WIDE (13.33 x 7.5 in)
 *
 * 改行・禁則の取扱い(重要):
 *  - テキスト内に手動改行(\n)は入れない。改行は PowerPoint の自動折返しに完全に
 *    任せる(箇条書きの項目が変わる箇所のみ段落を分ける)。
 *  - 全テキストランに lang: "ja-JP" を設定する(LANG 定数)。これにより PowerPoint
 *    標準の日本語禁則処理(「、」「。」「）」等を行頭に置かない)が有効になる。
 *    ※過去に行頭禁則が破られたのは、言語属性が既定の en-US となり日本語の
 *      禁則規則が適用されなかったため。モジュール側での事前折返しは、実描画との
 *      幅推定のずれによる不自然な改行を生むため行わない。
 *
 * 執筆・運用ルール(コンテンツ作成時に必ず遵守すること):
 *  1. 「／」(全角スラッシュ)は使用禁止。列挙は「、」または「・」で行う。
 *  2. 手動改行(\n)は原則使用しない(タイトルのみ許容)。使用する場合も改行後の
 *     行頭に禁則文字を置かないこと(checkText が検出)。
 *  3. 丸数字(①②③…)は文中では使用しない。番号付けは pptxgenjs の段落番号機能
 *     (bullet: { type: "number" })のみで行うか、(1), 1. 等の表記を用いる。
 *     (丸数字は指定フォント以外で描画されることがあるため)
 *  4. 出典を示す場合は、本文中に APA 形式の著者年表記(例: (Eysenbach, 2009))を
 *     入れた上で、同スライドの o.source に完全な書誌情報を渡す(下部に 14pt で表示)。
 *     出典が複数ある場合は必ず配列で渡すこと(1件=1段落として必ず改行される)。
 *     出典はぶら下げインデント約1cm、段落後間隔6ptで表示され、スライドの余白的に
 *     厳しい場合に限り自動的にフォントサイズを落とす(下限 10pt)。
 *  5. ファイルのプロパティ(概要)には作成者氏名のみを設定する(newDeck が処理。
 *     タイトル・件名・会社名等は設定しない)。
 *  6. 四角囲み(ポイント枠・カード)内は見出しも本文も 24pt とする。
 *  7. 「続き：」のような分割表現は使わない。列挙が長い場合も1つの箇条書きに
 *     まとめ、自動折返しで複数行にする。
 *  8. 保存は pres.writeFile ではなく、必ず writeDeck(pres, ファイル名) を使う
 *     (出典のぶら下げインデントを確定させる後処理を含むため)。
 */

const FONT = "Meiryo UI";
const LANG = "ja-JP";       // 全テキストに設定(日本語禁則処理の有効化)
const W = 13.33;
const H = 7.5;
const MARGIN = 0.7;
const LINE = 1.3;           // 標準行間
const BULLET_INDENT = 28;   // 箇条書き記号とテキストの間隔(約1cm = 28pt)

const THEMES = {
  // 案A: 紺 × アイスブルー(王道・医学/理系の学会発表向け)
  navy: {
    name: "Academic Navy",
    dark: "1E2761",      // 濃紺(タイトル・扉・まとめの背景)
    dark2: "2A3A80",     // 濃紺の明るめ(背景内モチーフ)
    text: "000000",      // 本文(必ず真っ黒 #000000 を使用する)
    sub: "5A5A5A",       // キャプション・補足
    tint: "E8F0FB",      // 薄色ボックス背景
    tintText: "1E2761",  // 薄色ボックス内の見出し
    accent: "C9A227",    // 強調(数値・キーワード)※多用しない
    chart: ["1E2761", "5B7BD5", "A8BEE8", "C9A227"],
  },
  // 案B: ティール × ミント(臨床・看護・ヘルスケア系向け)
  teal: {
    name: "Clinical Teal",
    dark: "046A73",
    dark2: "0A7E88",
    text: "000000",
    sub: "546E7A",
    tint: "E4F4F2",
    tintText: "046A73",
    accent: "02A88A",
    chart: ["046A73", "27A59B", "8FD0C9", "F2B84B"],
  },
  // 案C: チャコール × ワインレッド(ミニマル・人文/方法論講義向け)
  charcoal: {
    name: "Minimal Charcoal",
    dark: "37424A",
    dark2: "46535C",
    text: "000000",
    sub: "616161",
    tint: "F0F2F3",
    tintText: "37424A",
    accent: "8E1F2F",
    chart: ["37424A", "78909C", "C4CDD3", "8E1F2F"],
  },
};

/* ---------- テキスト検査(執筆・運用ルールの機械的チェック) ---------- */

const RE_ZEN_SLASH = /／/;
const RE_CIRCLED = /[①-⑳⓫-⓿㉑-㊿]/; // ①-⑳ 等の丸数字
const RE_HEAD_PUNCT = /\n[、。）」』〕】｝〉》・,.)!?！？]/; // 手動改行後の行頭禁則

function checkText(s, where) {
  if (typeof s !== "string") return s;
  if (RE_ZEN_SLASH.test(s)) {
    console.warn(`[academic_theme] ルール違反(「／」使用禁止): ${where}: "${s}"`);
  }
  if (RE_CIRCLED.test(s)) {
    console.warn(`[academic_theme] ルール違反(丸数字は文中使用禁止。段落番号機能か (1), 1. を使用): ${where}: "${s}"`);
  }
  if (RE_HEAD_PUNCT.test(s)) {
    console.warn(`[academic_theme] ルール違反(改行後の行頭に禁則文字): ${where}: "${s}"`);
  }
  return s;
}

/* ---------- 基本 ---------- */

function newDeck(theme, meta = {}) {
  const pptxgen = require("pptxgenjs");
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  // プロパティ(概要)は作成者氏名のみ。タイトル・件名・会社名等は設定しない。
  pres.author = meta.author || "";
  pres.title = "";
  pres.subject = "";
  pres.company = "";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  return pres;
}

/** 背景の円モチーフ(装飾は右上の大きな円1つのみ・全テーマ共通) */
function motif(slide, theme, onDark) {
  slide.addShape("ellipse", {
    x: W - 3.4, y: -2.2, w: 5.4, h: 5.4,
    fill: { color: onDark ? theme.dark2 : theme.tint },
    line: { type: "none" },
  });
}

/** 出典の必要行数を概算する(フォントサイズ自動調整のための目安計算のみに使用。
 *  改行位置の決定には一切使用しない) */
function estimateLines(text, widthIn, fontSize) {
  const em = fontSize / 72;
  let w = 0;
  for (const ch of text) {
    const c = ch.charCodeAt(0);
    w += (c >= 0x20 && c <= 0x7e) || (ch >= "｡" && ch <= "ﾟ") ? 0.5 * em : em;
  }
  return Math.max(1, Math.ceil(w / (widthIn * 0.95)));
}

/** スライド下部の出典表示(基本 14pt)。本文中は APA 著者年表記を用いること。
 *  source: string | string[]  複数の場合は1件=1段落として必ず改行される。
 *  ぶら下げインデント約1cm(不可視の空白バレットによる正規のインデント設定)、
 *  段落後間隔6pt。余白に収まらない場合に限りフォントサイズを落とす(下限 10pt)。
 *  改行位置は PowerPoint の自動折返しに任せる(手動改行はしない)。 */
function addSource(s, theme, source) {
  if (!source) return;
  const entries = (Array.isArray(source) ? source : [source])
    .map((t, i) => checkText(t, `source[${i}]`));
  const maxH = 1.05;                // 出典表示に使える最大高さ(in)
  const width = W - MARGIN * 2;
  const lineH = (size) => (size / 72) * 1.22;
  let fs = 14;
  let boxH;
  for (;;) {
    const totalLines = entries.reduce((a, t) => a + estimateLines(t, width, fs), 0);
    boxH = totalLines * lineH(fs) + (entries.length - 1) * (6 / 72);
    if (boxH <= maxH || fs <= 10) break;
    fs -= 1;
  }
  s.addText(
    entries.map((t) => ({
      text: t,
      options: {
        breakLine: true,                                  // 1件=1段落(必ず改行)
        bullet: { code: "0020", indent: BULLET_INDENT },  // 不可視バレットでぶら下げ約1cm
        paraSpaceAfter: 6,                                // 段落後間隔6pt
      },
    })),
    {
      x: MARGIN, y: H - 0.25 - Math.min(boxH, maxH), w: width, h: Math.min(boxH, maxH),
      fontFace: FONT, fontSize: fs, color: theme.sub, lang: LANG,
      align: "left", valign: "bottom", margin: 0,
    }
  );
}

/** 1. タイトルスライド */
function addTitleSlide(pres, theme, o) {
  const s = pres.addSlide();
  s.background = { color: theme.dark };
  motif(s, theme, true);
  s.addText(checkText(o.title, "title"), {
    x: MARGIN, y: 2.0, w: W - MARGIN * 2, h: 1.9,
    fontFace: FONT, fontSize: 40, bold: true, color: "FFFFFF", lang: LANG,
    align: "left", valign: "middle", margin: 0,
  });
  if (o.subtitle) {
    s.addText(checkText(o.subtitle, "subtitle"), {
      x: MARGIN, y: 3.9, w: W - MARGIN * 2, h: 1.1,
      fontFace: FONT, fontSize: 22, color: "D9E2F5", lang: LANG,
      align: "left", valign: "top", margin: 0,
    });
  }
  const lines = [o.presenter, o.affiliation, o.date].filter(Boolean).join("　|　");
  if (lines) {
    s.addText(lines, {
      x: MARGIN, y: H - 1.3, w: W - MARGIN * 2, h: 0.6,
      fontFace: FONT, fontSize: 18, color: "FFFFFF", lang: LANG,
      align: "left", valign: "middle", margin: 0,
    });
  }
  return s;
}

/** 2. セクション扉(章番号 + 章タイトル) */
function addSectionSlide(pres, theme, o) {
  const s = pres.addSlide();
  s.background = { color: theme.dark };
  motif(s, theme, true);
  if (o.number) {
    s.addText(String(o.number).padStart(2, "0"), {
      x: MARGIN, y: 2.1, w: 3.0, h: 1.7,
      fontFace: FONT, fontSize: 88, bold: true, color: theme.accent, lang: LANG,
      align: "left", valign: "middle", margin: 0,
    });
  }
  s.addText(checkText(o.title, "section title"), {
    x: MARGIN, y: 3.9, w: W - MARGIN * 2, h: 1.2,
    fontFace: FONT, fontSize: 36, bold: true, color: "FFFFFF", lang: LANG,
    align: "left", valign: "middle", margin: 0,
  });
  if (o.note) {
    s.addText(checkText(o.note, "section note"), {
      x: MARGIN, y: 5.1, w: W - MARGIN * 2, h: 0.6,
      fontFace: FONT, fontSize: 18, color: "D9E2F5", lang: LANG, margin: 0,
    });
  }
  return s;
}

/** 共通: 本文スライドのタイトル */
function contentTitle(s, theme, title) {
  s.addText(checkText(title, "content title"), {
    x: MARGIN, y: 0.5, w: W - MARGIN * 2, h: 0.9,
    fontFace: FONT, fontSize: 30, bold: true, color: theme.dark, lang: LANG,
    align: "left", valign: "middle", margin: 0,
  });
}

/** 3. 箇条書き + ポイント枠(2カラム)
 *  bullets: string[], point: {label, text}, source?: string | string[]
 *  改行は箇条書きの項目が変わる箇所のみ。項目内の折返しは自動折返しに任せる。 */
function addBulletSlide(pres, theme, o) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  contentTitle(s, theme, o.title);
  const hasPoint = !!o.point;
  const colW = hasPoint ? 7.3 : W - MARGIN * 2;
  s.addText(
    o.bullets.map((b, i) => ({
      text: checkText(b, `bullet[${i}]`),
      options: {
        bullet: { code: "2022", indent: BULLET_INDENT },
        breakLine: i < o.bullets.length - 1,
        paraSpaceAfter: 14,
      },
    })),
    {
      x: MARGIN, y: 1.7, w: colW, h: 4.9,
      fontFace: FONT, fontSize: 24, color: theme.text, lang: LANG,
      align: "left", valign: "top", lineSpacingMultiple: LINE,
    }
  );
  if (hasPoint) {
    const px = MARGIN + colW + 0.4;
    const pw = W - MARGIN - px;
    s.addShape("roundRect", {
      x: px, y: 1.9, w: pw, h: 4.5, rectRadius: 0.12,
      fill: { color: theme.tint }, line: { type: "none" },
    });
    s.addText(checkText(o.point.label || "Point", "point label"), {
      x: px + 0.35, y: 2.2, w: pw - 0.7, h: 0.6,
      fontFace: FONT, fontSize: 24, bold: true, color: theme.tintText, lang: LANG, margin: 0,
    });
    s.addText(checkText(o.point.text, "point text"), {
      x: px + 0.35, y: 2.85, w: pw - 0.7, h: 3.35,
      fontFace: FONT, fontSize: 24, color: theme.text, lang: LANG,
      valign: "top", lineSpacingMultiple: LINE, margin: 0,
    });
  }
  addSource(s, theme, o.source);
  return s;
}

/** 4. 3カードレイアウト cards: [{num?, header, text}] (最大3枚), source?: string | string[] */
function addCardSlide(pres, theme, o) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  contentTitle(s, theme, o.title);
  const n = o.cards.length;
  const gap = 0.4;
  const cw = (W - MARGIN * 2 - gap * (n - 1)) / n;
  o.cards.forEach((c, i) => {
    const x = MARGIN + i * (cw + gap);
    s.addShape("roundRect", {
      x, y: 1.9, w: cw, h: 4.6, rectRadius: 0.12,
      fill: { color: theme.tint }, line: { type: "none" },
    });
    s.addShape("ellipse", {
      x: x + 0.35, y: 2.25, w: 0.7, h: 0.7,
      fill: { color: theme.dark }, line: { type: "none" },
    });
    s.addText(c.num != null ? String(c.num) : String(i + 1), {
      x: x + 0.35, y: 2.25, w: 0.7, h: 0.7,
      fontFace: FONT, fontSize: 22, bold: true, color: "FFFFFF", lang: LANG,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(checkText(c.header, `card[${i}] header`), {
      x: x + 0.35, y: 3.15, w: cw - 0.7, h: 0.95,
      fontFace: FONT, fontSize: 24, bold: true, color: theme.tintText, lang: LANG,
      valign: "top", margin: 0,
    });
    s.addText(checkText(c.text, `card[${i}] text`), {
      x: x + 0.35, y: 4.15, w: cw - 0.7, h: 2.25,
      fontFace: FONT, fontSize: 24, color: theme.text, lang: LANG,
      valign: "top", lineSpacingMultiple: LINE, margin: 0,
    });
  });
  addSource(s, theme, o.source);
  return s;
}

/** 5. グラフ + 数値ハイライト chart: {labels, values, seriesName}, stat: {value, label}, source?: string | string[] */
function addChartStatSlide(pres, theme, o) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  contentTitle(s, theme, o.title);
  s.addChart("bar", [{ name: o.chart.seriesName || "系列1", labels: o.chart.labels, values: o.chart.values }], {
    x: MARGIN, y: 1.8, w: 7.6, h: 4.8,
    barDir: "col",
    chartColors: [theme.chart[0]],
    showLegend: false,
    showTitle: false,
    showValue: true,
    dataLabelPosition: "outEnd",
    dataLabelColor: theme.text,
    dataLabelFontFace: FONT,
    dataLabelFontSize: 14,
    catAxisLabelColor: theme.sub,
    catAxisLabelFontFace: FONT,
    catAxisLabelFontSize: 14,
    valAxisLabelColor: theme.sub,
    valAxisLabelFontFace: FONT,
    valAxisLabelFontSize: 12,
    catGridLine: { style: "none" },
    valGridLine: { color: "E0E0E0", size: 0.5 },
  });
  const px = MARGIN + 8.0;
  const pw = W - MARGIN - px;
  s.addShape("roundRect", {
    x: px, y: 2.3, w: pw, h: 3.4, rectRadius: 0.12,
    fill: { color: theme.tint }, line: { type: "none" },
  });
  s.addText(o.stat.value, {
    x: px + 0.3, y: 2.7, w: pw - 0.6, h: 1.5,
    fontFace: FONT, fontSize: 60, bold: true, color: theme.accent, lang: LANG,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText(checkText(o.stat.label, "stat label"), {
    x: px + 0.3, y: 4.3, w: pw - 0.6, h: 1.1,
    fontFace: FONT, fontSize: 24, color: theme.text, lang: LANG,
    align: "center", valign: "top", lineSpacingMultiple: LINE, margin: 0,
  });
  addSource(s, theme, o.source);
  return s;
}

/** 6. まとめ items: string[] */
function addSummarySlide(pres, theme, o) {
  const s = pres.addSlide();
  s.background = { color: theme.dark };
  motif(s, theme, true);
  s.addText(checkText(o.title || "まとめ", "summary title"), {
    x: MARGIN, y: 0.7, w: W - MARGIN * 2, h: 1.0,
    fontFace: FONT, fontSize: 34, bold: true, color: "FFFFFF", lang: LANG,
    valign: "middle", margin: 0,
  });
  o.items.forEach((t, i) => {
    const y = 2.1 + i * 1.5;
    s.addShape("ellipse", {
      x: MARGIN, y, w: 0.75, h: 0.75,
      fill: { color: theme.accent }, line: { type: "none" },
    });
    s.addText(String(i + 1), {
      x: MARGIN, y, w: 0.75, h: 0.75,
      fontFace: FONT, fontSize: 24, bold: true, color: "FFFFFF", lang: LANG,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(checkText(t, `summary item[${i}]`), {
      x: MARGIN + 1.05, y: y - 0.15, w: W - MARGIN * 2 - 1.05, h: 1.15,
      fontFace: FONT, fontSize: 24, color: "FFFFFF", lang: LANG,
      valign: "middle", lineSpacingMultiple: LINE, margin: 0,
    });
  });
  return s;
}

/** 保存処理。必ず pres.writeFile ではなくこの関数で保存すること。
 *  出典段落のぶら下げインデントは、書き出し後に不可視バレット(空白文字)を
 *  <a:buNone/> に置換して確定させる(PowerPoint は空白バレットを無視して
 *  ぶら下げを解除することがあるため、バレットなし + marL/indent の標準形に直す)。 */
async function writeDeck(pres, fileName) {
  const JSZip = require("jszip");
  const fs = require("fs");
  const buf = await pres.write({ outputType: "nodebuffer" });
  const zip = await JSZip.loadAsync(buf);
  const slideNames = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  for (const n of slideNames) {
    let xml = await zip.file(n).async("string");
    xml = xml
      .replace(/<a:buSzPct val="100000"\/><a:buChar char="&#x0020;"\/>/g, "<a:buNone/>")
      .replace(/<a:buSzPct val="100000"\/><a:buChar char=" "\/>/g, "<a:buNone/>");
    zip.file(n, xml);
  }
  const out = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(fileName, out);
  return fileName;
}

module.exports = {
  THEMES, FONT, LANG, W, H, MARGIN, LINE, BULLET_INDENT,
  newDeck, addTitleSlide, addSectionSlide, addBulletSlide,
  addCardSlide, addChartStatSlide, addSummarySlide, checkText, writeDeck,
};
