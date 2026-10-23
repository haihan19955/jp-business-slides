// 折り返しシミュレータの検証用コーパス：さまざまな文・幅・フォント・サイズのテキストボックスを並べた pptx を作る
// node tests/wrap_corpus.js tests/wrap_corpus.pptx → 同名の .json に各ボックスの仕様を書き出す
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");

const SENTENCES = [
  "AI-OCRで自動読み取りし、品番マスタと自動照合する",
  "受注処理の約7割を自動化。担当者は例外処理と顧客対応に集中できる",
  "初期費用480万円に対し、年間約860万円の削減効果",
  "FAX・メール・PDFを専用フォルダに自動保存",
  "照合済みのデータをAPI経由で基幹システムに登録",
  "照合できなかった項目のみ担当者に通知し、人が最終判断",
  "繁忙期の受注件数は通常月の1.8倍で、現体制では処理しきれない",
  "3つの選択肢のうち、案Bが費用対効果と導入リスクのバランスに優れる",
  "まずは「効果が大きく、すぐ始められる」受注処理から着手する",
  "11月中にご判断いただければ、来年3月の繁忙期前に稼働できる",
  "ミスと残業の根本原因は「人による転記」に集中している",
  "業務フローを分解すると、問題の8割は注文書の読み取りと転記の工程で発生している",
  "Salesforceとkintoneのデータを連携し、営業日報の入力をゼロにする",
  "顧客満足度（CSAT）は前年比+12pt改善し、解約率は3.2%まで低下した",
  "新規事業の立ち上げに必要な体制・予算・スケジュールを整理する",
  "ユーザーインタビュー（n=24）から、価格よりもサポート体制が重視されていることが分かった",
  "海外拠点を含む全社共通のセキュリティポリシーを2027年4月までに策定する",
  "Google Workspace と Microsoft 365 の併用をやめ、ツールを一本化する",
  "在庫回転日数を45日から30日に短縮し、キャッシュフローを約1.2億円改善する",
  "オンライン商談の比率は2024年の18%から2026年には52%まで拡大した",
  "既存顧客へのアップセルが売上成長の6割を占めており、新規開拓の比重は下がっている",
  "請求書発行から入金確認までのリードタイムを平均14日から3日に短縮",
  "現場のヒアリングでは「入力項目が多すぎる」という声が最も多かった",
  "プロジェクトマネージャー1名、エンジニア3名、デザイナー1名の体制で進める",
  "ChatGPTやClaudeなどの生成AIを業務で使う際のガイドラインを整備する",
  "物流コストは燃料費の高騰により前年同期比で約15%増加している",
  "採用単価は1人あたり平均82万円で、業界平均（約103万円）を下回っている",
  "週次の定例会議を廃止し、非同期のレポートとSlackでの相談に切り替える",
  "パイロット導入した3店舗では、レジ待ち時間が平均4分短縮された",
  "サブスクリプション型の料金体系に移行し、月額9,800円からのプランを用意する",
  "デジタルトランスフォーメーションの推進には、経営層のコミットメントが不可欠である",
  "コールセンターの応答率は87%で、目標の95%に届いていない",
  "要件定義・設計・開発・テスト・リリースの5フェーズで進める",
  "ご提案内容について、ご不明な点がございましたらお気軽にお問い合わせください",
  "本資料は2026年10月時点の情報に基づいて作成しています",
  "当社のクラウド型ERPは、中堅・中小企業向けに最適化されています",
  "売上高は前年比112%の34億円、営業利益率は8.4%となった",
  "インバウンド需要の回復により、訪日外国人の客単価は前年を上回っている",
  "リスク：読み取り精度が目標（98%）に届かない場合は、手入力との併用期間を延長する",
  "ポイントは、PowerPointが実際に計算した改行位置を使うことです",
];
const FONTS = ["Meiryo", "Yu Gothic", "BIZ UDPGothic"];
const SIZES = [12, 14, 16, 18];
const WIDTHS = [1.8, 2.4, 3.0, 3.7, 4.5, 5.6];

let seed = 7;
const rand = () => ((seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648);
const pick = (a) => a[Math.floor(rand() * a.length)];

const out = process.argv[2] || "wrap_corpus.pptx";
const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
const specs = [];
let id = 0;
for (let sl = 0; sl < 100; sl++) {
  const s = pres.addSlide();
  for (let k = 0; k < 6; k++) {
    const spec = {
      name: `box${id++}`, text: pick(SENTENCES), font: pick(FONTS), size: pick(SIZES),
      bold: rand() < 0.35, w: pick(WIDTHS), margin: rand() < 0.5 ? 0 : null,
    };
    const col = k % 2, row = Math.floor(k / 2);
    const opts = {
      x: 0.3 + col * 6.4, y: 0.3 + row * 2.4, w: spec.w, h: 2.2, fontFace: spec.font, fontSize: spec.size,
      bold: spec.bold, valign: "top", isTextBox: true, objectName: spec.name,
    };
    if (spec.margin !== null) opts.margin = spec.margin;
    if (process.env.LANG_JA) opts.lang = "ja-JP";
    s.addText(spec.text, opts);
    specs.push(spec);
  }
}
pres.writeFile({ fileName: out }).then(() => {
  fs.writeFileSync(out.replace(/\.pptx$/, ".specs.json"), JSON.stringify(specs, null, 1));
  console.log("wrote", out, specs.length, "boxes");
});
