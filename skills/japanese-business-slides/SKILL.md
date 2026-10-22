---
name: japanese-business-slides
description: 日本語のビジネス資料（提案書・企画書・社内報告・稟議・営業資料などの PowerPoint / .pptx）を、コンサル品質の構成と日本語組版で作る・直すためのスキル。タイトルを「結論」にする書き方、日本語フォント、出所・ページ番号の作法に加えて、PowerPoint 実機の改行位置を調べて単語の途中で切れる「泣き別れ」や数字と単位の分離、文字のはみ出しを自動で検出・修正する。日本語のスライドやパワポを作るとき、既存の日本語資料を整えるとき、「泣き別れ」「改行がおかしい」「資料をきれいに」と言われたときに使う。Use for any Japanese-language business deck or .pptx, including fixing bad Japanese line breaks in existing decks.
---

# 日本語ビジネス資料スキル

日本の会社でそのまま出せる資料を作る。見た目より先に **構成（何を言うか）** を決め、最後に **PowerPoint 実機で改行とはみ出しを検証** してから渡す。

公式の pptx スキルがある環境では、pptxgenjs やファイル操作の細かい注意点はそちらに従う。日本語の構成・文章・組版・検証はこのスキルが優先する（欧文前提の「タイトル36pt以上」「テキストだけのスライドは避ける」などは日本のビジネス資料には当てはまらない）。

## 1. 構成を決める（書き始める前に）

1. **誰が・何のために読むか** を一文で決める。決裁してほしいのか、報告なのか、合意形成なのか。会議で投影するのか、事前に読ませるのか。分からなければユーザーに聞く。
2. **結論を先に書く。** 冒頭にエグゼクティブサマリー（結論＋根拠3〜4点）を置き、その後に各論、最後にネクストステップ（何を・誰が・いつまでに）。
3. **1スライド1メッセージ。** 各スライドのタイトルは、そのスライドの結論を1文で書いた「メッセージ」にする。
   - ✖「受注業務の現状について」「課題一覧」「スケジュール」（テーマ名だけ）
   - ○「ミスと残業の根本原因は『人による転記』に集中している」
   - ○「3か月で本番稼働し、4か月目から効果測定に入る」
4. タイトルだけを順に読めば話が通じるかを確認する（ストーリーライン）。通じなければ順番かメッセージを直す。
5. 主張には根拠（数字・事実・出所）をつける。「空（事実）→雨（解釈）→傘（打ち手）」の順で書くと流れが自然になる。

## 2. 日本語の書き方

- **タイトル**：全角40字以内、2行まで。句点「。」は付けない。言い切る（「〜する」「〜になる」「〜が必要」）。
- **本文**：1文は60字以内を目安に。箇条書きは体言止めか「です・ます」のどちらかにそろえ、「だ・である」と混ぜない。
- **数字**：半角数字。桁区切りはカンマ（3,200万円）。単位を省略しない。全角英数字・半角カナは使わない。
- **表記**：「ください」「できる」「いただく」「すべて」「か月」など、ひらがな表記にそろえる（公用文の書き方に準拠）。社名・製品名は正式名称で。
- **出所**：グラフ・表・数字には左下に「出所：〇〇（20XX年X月）」を書く。試算なら「出所：当社試算（前提：〜）」。
- **言葉を盛らない**：根拠のない「必ず」「圧倒的」「業界No.1」は書かない。

## 3. 作る（jpdeck.js）

`scripts/jpdeck.js` に日本のコンサル資料型のレイアウトをまとめてある。Node のスクリプトを書いて呼ぶ。pptxgenjs は `npm install pptxgenjs` で入れる。

```js
const { createDeck } = require("<このスキルのディレクトリ>/scripts/jpdeck.js");
const deck = createDeck({ font: "Meiryo", palette: "navy", footer: "株式会社〇〇　社外秘" });
deck.cover({ title: "…", subtitle: "…", date: "2026年10月", org: "…" });
deck.summary({ title: "結論の1文", points: [{ head: "要点", body: "補足" }, …] });
// …
await deck.save("out.pptx");
```

| メソッド | 用途 | 主な引数 |
|---|---|---|
| `cover` | 表紙 | title（`\n` で改行）, subtitle, date, org |
| `section` | 章扉 | title, sub |
| `summary` | エグゼクティブサマリー（番号付きの要点） | points: [{head, body}] |
| `issueSolution` | 現状 → 課題 → 打ち手 | rows: [[現状, 課題, 打ち手], …], headers |
| `chartWithTakeaway` | グラフ＋右側に「示唆」 | chart: {type, data, unit}, takeaways: [] |
| `compare` | 選択肢の比較表（推奨案をハイライト） | columns, rows: [{label, values}], highlight |
| `matrix` | 2×2 マトリクス | xAxis, yAxis, quadrants[左上,右上,左下,右下], highlight |
| `kpi` | 大きな数字のカード | cards: [{label, value, unit, note}] |
| `process` | 手順（STEP 1→2→3） | steps: [{head, body}] |
| `roadmap` | ガントチャート風の計画 | periods: [], tracks: [{name, items: [{label, start, end, key}]}] |
| `nextSteps` | 何を・誰が・いつまでに | items: [{what, who, when}] |

すべてのコンテンツスライドで `title`（メッセージ）, `lead`（リード文・任意）, `bottomLine`（下部の結論ボックス・任意）, `source`（出所）, `notes`（スピーカーノート）が使える。

- `font`: 既定は `Meiryo`（Windows 標準）。Mac でも開く資料は `Yu Gothic`、ユニバーサルデザイン重視なら `BIZ UDPGothic`。1デッキ1〜2書体まで。
- `palette`: `navy`（既定）/ `green` / `wine` / `mono`。会社のブランドカラーがあれば `PALETTES` に足す。
- 同じレイアウトを3枚以上続けない。数字はグラフか KPI カード、比較は表、手順は process、と中身に合った型を選ぶ。
- 1枚に入りきらなければ文字を小さくせず、スライドを分ける（本文12pt未満にしない）。

## 4. 検証する（必須）

作ったら毎回この順で確認する。**ユーザーに渡すのは error が 0 件になってから。**

```bash
# 1) 泣き別れ（単語の途中での改行）を自動修正  ※Windows + PowerPoint
python scripts/jp_fix_wrap.py out.pptx -o out_fixed.pptx

# 2) PowerPoint 実機の改行位置を書き出して、全項目をチェック
powershell -ExecutionPolicy Bypass -File scripts/export_layout.ps1 -Pptx out_fixed.pptx -Out layout.json
python scripts/jp_lint.py out_fixed.pptx --layout layout.json

# 3) 画像にして目で確認
powershell -ExecutionPolicy Bypass -File scripts/render_slides.ps1 -Pptx out_fixed.pptx -OutDir renders
```

jp_lint.py が見るもの:

| コード | 内容 |
|---|---|
| L1 | 文字が枠からはみ出している（PowerPoint が計算した実際の高さで判定） |
| L2 | 単語の途中で改行されている「泣き別れ」（例：品番マスタと自／動照合） |
| L3 | 最終行が1〜2文字だけ |
| L4 | 数字と単位が別の行に分かれている（例：860万／円） |
| T1〜T4 | タイトルが空・句点つき・テーマ名だけ・長すぎる |
| F1〜F3 | 日本語に欧文フォント、12pt未満の本文、書体の混在 |
| S1 | グラフに出所がない |
| W1〜W4 | です・ます／だ・であるの混在、全角英数字、表記ゆれ、半角カナ |

直し方:
- **L2/L4**：`jp_fix_wrap.py` が文節（BudouX）の境目に行内改行を入れて直す。「自動では直せない」と出たものは、言い回しを短くするか枠を広げる。生成スクリプト側で直すなら、テキストに `\n` を入れる。
- **L1**：文字を削る → 枠を広げる → スライドを分ける、の順で検討する。
- **T3**：タイトルをテーマ名から結論の文に書き換える。
- 画像では、はみ出し・重なり・余白の偏り・タイトル位置のずれ・同じレイアウトの連続がないかを見る。作ったコードを見たあとは思い込みで見落としやすいので、画像は新しい目で1枚ずつ確認する。

### Windows + PowerPoint がない環境

`jp_lint.py out.pptx`（--layout なし）で文言・フォント・構成のチェックだけ行う。LibreOffice は日本語の改行位置が PowerPoint と異なるため、泣き別れの判定には使わない。その旨をユーザーに伝え、手元の PowerPoint で `jp_fix_wrap.py` を実行してもらう。

## 5. 既存の資料を直すとき

ユーザーの .pptx をそのまま直す場合は、まず検証だけ行い、結果を見せてから直す。

1. `export_layout.ps1` → `jp_lint.py --layout` で問題を一覧にする
2. `jp_fix_wrap.py` で泣き別れを直す（元ファイルは上書きしない）
3. タイトル・文言の修正は、内容が変わるので提案として出し、了承を得てから反映する

## 依存

- Node.js + `pptxgenjs`（作成）
- Python 3 + `python-pptx`, `budoux`（`pip install python-pptx budoux`）
- Windows + Microsoft PowerPoint（改行位置の取得・自動修正・画像化）
