const path = require("path");
const {
  THEMES,
  newDeck,
  addTitleSlide,
  addBulletSlide,
  addSummarySlide,
  writeDeck,
} = require("../academic_theme");

async function main() {
  const pres = newDeck(THEMES.teal, { author: "作成者名" });

  addTitleSlide(pres, THEMES.teal, {
    title: "研究発表タイトル",
    subtitle: "学術発表用テーマの使用例",
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
    items: ["要点1", "要点2", "要点3"],
  });

  const outputPath = process.argv[2] || path.join(process.cwd(), "academic_theme_sample.pptx");
  await writeDeck(pres, outputPath);
  console.log(outputPath);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
