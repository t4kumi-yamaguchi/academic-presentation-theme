# Academic Presentation Theme

日本語の学術発表・講義用PowerPointをPptxGenJSで作るための、共通テーマとレイアウト関数です。

## 特徴

- 16:9ワイド画面
- Meiryo UIを使用
- 本文24pt以上、本文色は黒
- 紺、ティール、チャコールの3テーマ
- 日本語の禁則処理を有効化する`ja-JP`指定
- 全角スラッシュ、丸数字、改行直後の禁則文字を検出
- APA形式の出典欄と、ぶら下げインデントの後処理
- タイトル、章扉、箇条書き、カード、グラフ、まとめの基本レイアウト

## 必要環境

- Node.js 18以上
- PptxGenJS 4.0.1
- JSZip 3.10.1

## セットアップ

```bash
npm install
```

## 使い方

```js
const {
  THEMES,
  newDeck,
  addTitleSlide,
  addBulletSlide,
  addSummarySlide,
  writeDeck,
} = require("./academic_theme");

async function main() {
  const pres = newDeck(THEMES.teal, { author: "作成者名" });

  addTitleSlide(pres, THEMES.teal, {
    title: "研究発表タイトル",
    subtitle: "副題",
    presenter: "発表者名",
    affiliation: "所属",
    date: "2026年9月3日",
  });

  addBulletSlide(pres, THEMES.teal, {
    title: "研究の背景",
    bullets: [
      "背景を簡潔に記載する",
      "箇条書き内の折り返しはPowerPointに任せる",
    ],
    point: {
      label: "Point",
      text: "このスライドで最も伝えたい内容",
    },
    source: "Author, A. A. (2026). Article title. Journal Name, 1(1), 1–10.",
  });

  addSummarySlide(pres, THEMES.teal, {
    title: "まとめ",
    items: ["要点1", "要点2", "要点3"],
  });

  await writeDeck(pres, "presentation.pptx");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
```

AIへの指示では、たとえば次のように指定できます。

> `academic_theme.js`の`THEMES.teal`を使い、本文と出典のルールを守ってスライドを生成してください。保存には`writeDeck`を使ってください。

## テーマ

| キー | 用途の目安 |
|---|---|
| `navy` | 医学・理系の学会発表 |
| `teal` | 臨床・看護・ヘルスケア |
| `charcoal` | 人文・方法論の講義 |

## 主なAPI

- `newDeck(theme, meta)`
- `addTitleSlide(pres, theme, options)`
- `addSectionSlide(pres, theme, options)`
- `addBulletSlide(pres, theme, options)`
- `addCardSlide(pres, theme, options)`
- `addChartStatSlide(pres, theme, options)`
- `addSummarySlide(pres, theme, options)`
- `checkText(text, where)`
- `writeDeck(pres, fileName)`

## 運用上の注意

- 元の設計では、本文内の手動改行、全角スラッシュ、丸数字を避けます。
- 出典を表示する場合は、本文中に著者年表記を置き、完全な書誌情報を`source`へ渡します。
- PowerPoint生成後は、文字切れ、重なり、改行、出典欄をMicrosoft PowerPointで確認してください。
- `writeDeck`には出典欄のXML後処理が含まれるため、直接`pres.writeFile`を呼ばないでください。

## ライセンス

このリポジトリには現時点でライセンスを設定していません。外部公開や第三者への再配布を行う場合は、用途に合うライセンスを追加してください。
