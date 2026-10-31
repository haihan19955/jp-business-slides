# japanese-business-slides

**AI が作った日本語のパワポの変な改行を、文字を1字も変えずに直す Claude スキル。無料・MIT。**
単語の途中で切れる「泣き別れ」と、「、」が行頭に来る禁則違反を自動で直します。新しい資料を日本のビジネス資料の型で作ることもできます。

> **English summary:** A Claude skill that fixes bad Japanese line breaks in PowerPoint decks (mid-word wraps and kinsoku violations) without changing a single character of the text. It also builds new decks in Japanese business style.

## 直すとこうなる

サンプル資料（10枚）から、実際の修正結果をそのまま載せています（`｜` が改行の位置）。

| 直す前 | 直した後 |
|---|---|
| 初期費用480万円に対し、年間約860万｜円の削減効果 | 初期費用480万円に対し、｜年間約860万円の削減効果 |
| 照合済みのデータをAPI経｜由で基幹システムに登録 | 照合済みのデータを｜API経由で基幹システムに｜登録 |
| FAX・メール・PDFを専用｜フォルダに自動保存 | FAX・メール・PDFを｜専用フォルダに自動保存 |
| 担当者3名の処理能力（約150時｜間）を大きく超えている | 担当者3名の処理能力｜（約150時間）を大きく｜超えている |

このサンプルでは、泣き別れなどの error が **10 件 → 0 件** になりました。

![デモ：修正前 → jp_fix_wrap.py → 修正後](images/demo.gif)

## 使い方（1分）

```bash
npx skills add haihan19955/jp-business-slides
pip install python-pptx budoux fonttools
```

あとは Claude Code に「この提案書.pptx の泣き別れをチェックして直して」と頼むだけです。元のファイルは上書きしません。新しい資料も作るときは `npm install pptxgenjs` も入れてください。

<details>
<summary>ほかの入れ方（Claude Code プラグイン、git clone）</summary>

Claude Code のプラグインとして入れる場合は、Claude Code の中で次の 2 行を実行します。スキルは `jp-business-slides:japanese-business-slides` という名前で表示されます。

```
/plugin marketplace add haihan19955/jp-business-slides
/plugin install jp-business-slides@jp-business-slides
```

手動で入れる場合:

```bash
git clone https://github.com/haihan19955/jp-business-slides.git
# スキル本体をユーザーのスキルフォルダへ
cp -r jp-business-slides/skills/japanese-business-slides ~/.claude/skills/
```

Windows（PowerShell）の場合:

```powershell
git clone https://github.com/haihan19955/jp-business-slides.git
Copy-Item -Recurse jp-business-slides\skills\japanese-business-slides "$env:USERPROFILE\.claude\skills\"
```

</details>

## 直すときに守ること

直すのは改行と、テキストの言語設定だけです。文字は足しも削りもしません。保存の前に全段落の文字を元のファイルと照合し、1字でも違えば出力を消してエラーにします（`tests/test_text_fidelity.py` で確かめています）。社外に出す資料の数字や言い回しを、ツールが勝手に変えてしまわないようにするためです。

## 仕組み

AI にスライドを作らせると、日本語は「専用｜フォルダ」「API経｜由」のように単語の途中で改行されがちです。PowerPoint には日本語を文節で折り返す機能がないため、人が1か所ずつ Shift+Enter で直すしかありませんでした。

このスキルは、

1. **PowerPoint 本体に実際にレイアウトさせて**、どこで改行されたかを取得し
2. **BudouX**（Google の日本語文節分割）で、その改行が単語の途中かどうかを判定し
3. **文節の切れ目に行内改行を入れて**、もう一度 PowerPoint で確認する

までを自動で行います。このスキルで作った資料だけでなく、**手元にある既存の .pptx にも使えます。** Windows は PowerPoint 実機で、Mac / Linux はフォントの文字幅からの再現（実機との一致率 98.8%）で動きます。

解説記事：
- note（使い方の紹介）：[AIが作ったパワポの「泣き別れ」を自動で直す。Claude Code用スキルを無料公開しました](https://note.com/glad_moose4636/n/n752279cadf64)
- Zenn（実装の解説）：[PowerPointの日本語「泣き別れ」を、PowerPoint自身に改行位置を聞いて自動で直す](https://zenn.dev/haihan19955/articles/93a10be81996a8)
- Zenn（禁則処理の検証）：[pptxgenjsやpython-pptxで作ったPowerPointは、日本語の禁則処理が効いていない](https://zenn.dev/haihan19955/articles/6bd0d7607f51cb)

![修正前後の比較](images/before_after_process.png)

![修正前後の比較（KPIカード）](images/before_after_kpi.png)

### pptxgenjs / python-pptx で作った資料は、禁則処理が効いていない

pptxgenjs はテキストの言語を `en-US` として書き出し、python-pptx は言語を書きません。PowerPoint は言語が `ja-JP` のテキストにしか日本語の禁則処理をしないので、どちらの場合も「、」「）」「ー」が行頭に、「（」が行末に来ます。600 個のテキストボックスで確かめたところ、`en-US` では禁則違反が 67 か所あり、`ja-JP` に変えると 0 か所になりました。

`jp_fix_wrap.py` は泣き別れを直す前に、日本語の段落の言語を `ja-JP` に書き換えます。

## できること

| | |
|---|---|
| 📐 **コンサル型の構成** | タイトル＝結論（メッセージライン）、エグゼクティブサマリー、現状→課題→打ち手、出所・ページ番号の作法を Claude に守らせる |
| 🧩 **日本のビジネス資料の型 11 種** | 表紙・サマリー・課題整理・グラフ＋示唆・比較表・2×2・KPI・プロセス・ロードマップ・ネクストステップ・章扉 |
| ✂️ **泣き別れ・禁則違反の自動修正** | `jp_fix_wrap.py`：言語を ja-JP に直し、改行位置をもとに文節の切れ目で改行し直す |
| 🔍 **日本語資料のチェッカー** | `jp_lint.py`：はみ出し・泣き別れ・禁則違反・数字と単位の分離・テーマ名だけのタイトル・欧文フォント・文体の混在・表記ゆれなど 18 項目 |
| 🍎 **Mac / Linux 対応** | `jp_wrap_sim.py`：PowerPoint なしで、フォントの文字幅と禁則規則から改行位置を再現（`--simulate`） |
| 🖼️ **実機レンダリング** | `render_slides.ps1`：PowerPoint 本体で PNG 化して目視確認（LibreOffice と違い、本番と同じ見た目） |

### サンプル（examples/）

架空の中小メーカー向け「受注業務DX」提案資料 10 枚。Claude がこのスキルで作成しました。`sample.pptx` が修正前（error 10 件）、`sample_fixed.pptx` が修正後（error 0 件）です。

![サンプル資料の一覧](images/gallery.png)

## 改行位置の取り方と精度

改行位置の取り方は2通りあります。

| | 改行位置の取り方 | 精度 |
|---|---|---|
| Windows + PowerPoint | PowerPoint 本体に計算させる（COM） | 実機そのもの |
| Mac / Linux（`--simulate`） | フォントの文字幅と禁則規則から再現 | PowerPoint 実機との一致率 98.8%（下記） |

**シミュレーションの精度**：600 個のテキストボックス（文・幅・書体・サイズ・太字の組み合わせ）を PowerPoint 実機でレイアウトさせ、段落ごとの改行がすべて一致するかを比べました（`tests/eval_wrap_sim.py`）。

| 書体 | 一致率 |
|---|---|
| Yu Gothic（游ゴシック） | 100.0%（193/193） |
| Meiryo（メイリオ） | 99.1%（210/212） |
| BIZ UDPGothic | 97.4%（190/195） |
| 全体 | 98.8%（593/600） |

シミュレーションだけで直した資料を PowerPoint 実機で測り直すと、サンプル資料は泣き別れ・禁則違反が 10 件 → 0 件、python-pptx で作った資料は 9 件 → 1 件でした。Mac で Meiryo のように書体ファイルがない場合は近い書体で代用するので、精度は下がります。

## 頼み方とコマンド

Claude Code にふつうに頼むだけです。

```
来週の経営会議用に、受注業務のデジタル化の提案資料を10枚くらいで作って。
決裁してもらうのが目的。数字は添付のExcelを使って。
```

```
この提案書.pptx の泣き別れとはみ出しをチェックして直して。
```

スクリプトを直接使うこともできます。

```bash
# 既存の資料の泣き別れ・禁則違反を自動修正（元ファイルは上書きしない）
python skills/japanese-business-slides/scripts/jp_fix_wrap.py 提案書.pptx -o 提案書_修正版.pptx

# チェック（PowerPoint の実際の改行位置を使う）
powershell -ExecutionPolicy Bypass -File skills/japanese-business-slides/scripts/export_layout.ps1 -Pptx 提案書_修正版.pptx -Out layout.json
python skills/japanese-business-slides/scripts/jp_lint.py 提案書_修正版.pptx --layout layout.json
```

Mac / Linux では `--simulate` を付けます。

```bash
python skills/japanese-business-slides/scripts/jp_fix_wrap.py 提案書.pptx -o 提案書_修正版.pptx --simulate
python skills/japanese-business-slides/scripts/jp_lint.py 提案書_修正版.pptx --simulate
```

出力例:

```
== スライド 3 ==
 ✖ L2 [Text 16] 単語の途中で改行されている（泣き別れ）: 「…、品番マスタと自｜動照合…」
     → 「自動照合」の前で改行する（jp_fix_wrap.py で自動修正できる）
== スライド 6 ==
 ✖ L4 [Text 13] 数字と単位が別の行に分かれている: 「…、年間約860万｜円の削減効果…」
     → 「年間約860万円の」を同じ行に収める（jp_fix_wrap.py で自動修正できる）
```

別の資料（python-pptx で作った一般的な .pptx）では、タイトルの書き方も指摘します:

```
 △ T3 タイトルがテーマ名になっている（結論が書かれていない）: 受注業務の現状について。
     → 「〜について」ではなく「〜なので、〜すべき」のように、このスライドの結論を書く
```

## チェック項目

| コード | 内容 |
|---|---|
| L1 | 文字が枠からはみ出している |
| L2 | 単語の途中で改行されている（泣き別れ） |
| L3 | 最終行が1〜2文字だけ |
| L4 | 数字と単位が別の行に分かれている |
| K1 | 禁則違反（「、」「）」「ー」が行頭、「（」が行末） |
| K2 | 日本語のテキストの言語が ja-JP になっていない（禁則処理が効かない） |
| T1–T4 | タイトルが空・句点つき・テーマ名だけ・長すぎる |
| F1–F3 | 日本語に欧文フォント・本文12pt未満・書体の混在 |
| S1 | グラフに出所がない |
| W1–W4 | です・ます／だ・であるの混在・全角英数字・表記ゆれ・半角カナ |

## 構成

```
skills/japanese-business-slides/
├── SKILL.md              # Claude への指示（構成・書き方・作成・検証の手順）
└── scripts/
    ├── jpdeck.js         # 日本のビジネス資料の型（pptxgenjs）
    ├── jp_lint.py        # チェッカー
    ├── jp_fix_wrap.py    # 泣き別れ・禁則違反の自動修正
    ├── jp_wrap_sim.py    # PowerPoint の折り返しの再現（Mac / Linux 用）
    ├── export_layout.ps1 # PowerPoint 実機の改行位置を JSON に書き出す
    └── render_slides.ps1 # PowerPoint 実機で PNG に書き出す
examples/sample.js        # サンプル資料の生成スクリプト
tests/wrap_corpus.js      # 折り返し検証用の 600 ボックスを作る
tests/eval_wrap_sim.py    # シミュレーションと PowerPoint 実機の一致率を測る
tests/test_text_fidelity.py # 修正で文字が変わらないことを確かめる
```

## 苦手なこと

- 改行を入れた結果、最終行が2文字だけになることがあります（サンプルの warn は 1 件 → 2 件に増えました）。error ではなく warn として出すので、目で見て直してください。
- 泣き別れの判定は BudouX の文節区切りに基づきます。固有名詞や専門用語で区切りが不自然な場合があります。
- 1つの文節が1行より長い場合は、「・」「、」の後ろで切るか、言い回しの変更を提案します。
- グラフ内の文字（軸ラベルなど）はチェック対象外です。
- シミュレーションは段落の最初の run の書体・太字・言語で代表させます（サイズは文字ごとに見ます）。1つの段落で書体が混在すると誤差が出ます。

## このスキルについて

このリポジトリのコード・サンプル資料・README は、Claude（Anthropic の AI）が設計・実装・テストしました。不具合や「この型がほしい」という要望は Issue へどうぞ。

## ライセンス

MIT
