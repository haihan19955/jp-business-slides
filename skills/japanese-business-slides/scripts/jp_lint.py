#!/usr/bin/env python3
"""jp_lint.py — 日本語ビジネス資料（.pptx）の品質チェッカー

使い方:
  python jp_lint.py deck.pptx                      # 文言・フォント・構成のチェック
  python jp_lint.py deck.pptx --layout layout.json # + PowerPoint 実機の改行位置で「泣き別れ」「はみ出し」を検出
  python jp_lint.py deck.pptx --json               # 機械可読な出力

layout.json は export_layout.ps1（Windows + PowerPoint）で作る。
終了コード: error が 1 件以上あれば 1、なければ 0。

依存: python-pptx, budoux（pip install python-pptx budoux）
"""
import argparse
import json
import re
import sys
import unicodedata
from collections import Counter, defaultdict

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Emu

try:
    import budoux
    _PARSER = budoux.load_default_japanese_parser()
except ImportError:  # 泣き別れチェックだけ無効になる
    _PARSER = None

NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

JP_RE = re.compile(r"[぀-ヿ㐀-鿿ｦ-ﾟ]")
ZENKAKU_ALNUM_RE = re.compile(r"[０-９Ａ-Ｚａ-ｚ]")
HANKAKU_KANA_RE = re.compile(r"[ｦ-ﾟ]")
LATIN_ONLY_FONTS = {
    "arial", "calibri", "calibri light", "aptos", "aptos display", "times new roman", "helvetica",
    "segoe ui", "verdana", "tahoma", "georgia", "cambria", "garamond", "century gothic", "roboto", "inter",
}
# 「テーマ」だけを表すタイトル末尾（メッセージになっていない可能性が高い）
TOPIC_ENDINGS = ("について", "の概要", "概要", "の現状", "まとめ", "一覧", "の比較", "の課題", "の検討", "の背景",
                 "スケジュール", "体制", "目次", "アジェンダ", "はじめに", "背景", "目的")
UNIT_START = ("円", "%", "％", "件", "時間", "か月", "ヶ月", "カ月", "万", "億", "千", "人", "名", "社", "倍", "年", "月", "日", "分", "秒", "個", "台", "点", "回")
OK_BREAK_AFTER = set("、。，．・：；）」』】〕〉》！？!?)]:;/ 　")
YURE_PAIRS = [
    ("下さい", "ください"), ("出来", "でき"), ("頂き", "いただき"), ("頂け", "いただけ"), ("致し", "いたし"),
    ("全て", "すべて"), ("様々", "さまざま"), ("及び", "および"), ("又は", "または"), ("更に", "さらに"),
    ("ヶ月", "か月"), ("カ月", "か月"), ("ケ月", "か月"), ("事が", "ことが"), ("時に", "ときに"),
]
DESU_RE = re.compile(r"(です|ます|でした|ました|ません|でしょう)[。」）]?$")
DEARU_RE = re.compile(r"(である|だ|だった|ではない|と考える|と言える)[。」）]?$")


class Report:
    def __init__(self):
        self.items = []

    def add(self, level, code, slide, msg, fix="", where=""):
        self.items.append({"level": level, "code": code, "slide": slide, "where": where, "message": msg, "fix": fix})

    @property
    def errors(self):
        return sum(1 for i in self.items if i["level"] == "error")


def visible_len(s):
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in s)


def shape_texts(shape):
    """(where, paragraphs[], run_elements[]) を返す。表のセルとグループも辿る。"""
    out = []
    if shape.shape_type == 6:  # group
        for s in shape.shapes:
            out += shape_texts(s)
        return out
    if getattr(shape, "has_table", False) and shape.has_table:
        for r, row in enumerate(shape.table.rows, 1):
            for c, cell in enumerate(row.cells, 1):
                out.append((f"{shape.name}[r{r}c{c}]", cell.text_frame))
        return out
    if shape.has_text_frame:
        out.append((shape.name, shape.text_frame))
    return out


def theme_ea_font(prs):
    """テーマのマイナーフォント（本文）の東アジア書体名。"""
    try:
        master = prs.slide_masters[0]
        for rel in master.part.rels.values():
            if rel.reltype.endswith("/theme"):
                xml = rel.target_part.blob.decode("utf-8", "ignore")
                m = re.search(r"<a:minorFont>.*?<a:ea typeface=\"([^\"]*)\"", xml, re.S)
                return m.group(1) if m else ""
    except Exception:
        pass
    return ""


def is_title(shape):
    if not shape.is_placeholder:
        return False
    try:
        return shape.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
    except Exception:
        return False


def check_deck(prs, rep):
    theme_ea = theme_ea_font(prs)
    jp_fonts = Counter()
    all_text = []
    body_endings = Counter()
    n = len(prs.slides)

    for idx, slide in enumerate(prs.slides, 1):
        slide_xml = slide._element.xml
        texts = []
        has_graphic = any(sh.shape_type in (3, 19) or getattr(sh, "has_chart", False) or getattr(sh, "has_table", False) for sh in slide.shapes)
        title_shape = next((sh for sh in slide.shapes if is_title(sh)), None)

        # --- タイトル（メッセージライン） ---
        if title_shape is not None:
            t = title_shape.text_frame.text.strip().replace("\n", "")
            if not t:
                rep.add("error", "T1", idx, "タイトルが空です", "スライドで一番伝えたい結論を1文で書く", title_shape.name)
            else:
                if t.endswith("。"):
                    rep.add("warn", "T2", idx, f"タイトルが「。」で終わっている: {t}", "タイトルの句点は取る")
                if t.rstrip("。．.").endswith(TOPIC_ENDINGS):
                    rep.add("warn", "T3", idx, f"タイトルがテーマ名になっている（結論が書かれていない）: {t}",
                            "「〜について」ではなく「〜なので、〜すべき」のように、このスライドの結論を書く")
                if visible_len(t) > 80:
                    rep.add("warn", "T4", idx, f"タイトルが長すぎる（全角{visible_len(t)//2}字相当）: {t}", "全角40字以内（2行まで）に絞る")

        for sh in slide.shapes:
            for where, tf in shape_texts(sh):
                name = where
                is_src = name.startswith("source") or tf.text.strip().startswith(("出所", "出典", "注", "※"))
                for p in tf.paragraphs:
                    ptext = "".join(r.text for r in p.runs).strip()
                    if not ptext:
                        continue
                    texts.append(ptext)
                    if not (title_shape is not None and sh is title_shape):
                        for seg in re.split(r"(?<=。)", ptext):
                            seg = seg.strip()
                            if DESU_RE.search(seg):
                                body_endings["desu"] += 1
                            elif DEARU_RE.search(seg):
                                body_endings["dearu"] += 1
                    for r in p.runs:
                        if not r.text.strip():
                            continue
                        rpr = r._r.find(NS_A + "rPr")
                        ea = None
                        if rpr is not None:
                            e = rpr.find(NS_A + "ea")
                            if e is not None:
                                ea = e.get("typeface")
                        ea_eff = ea if ea else theme_ea
                        if JP_RE.search(r.text):
                            if ea_eff and not ea_eff.startswith("+"):
                                jp_fonts[ea_eff] += 1
                            # ea が未指定なら PowerPoint は日本語フォントで代替表示するので問題にしない。
                            # 欧文専用フォントが明示されているときだけ指摘する（layout.json があればそちらの実フォントで再判定）。
                            if ea_eff and ea_eff.lower() in LATIN_ONLY_FONTS:
                                rep.add("error", "F1", idx, f"日本語に欧文フォント（{ea_eff}）が指定されている: 「{r.text[:20]}」",
                                        "東アジア用フォント（ea）に Meiryo / Yu Gothic / BIZ UDPゴシック などを指定する", where)
                        size = r.font.size.pt if r.font.size else None
                        if size and not is_src and name not in ("unit",) and "Slide Number" not in name and size < 12:
                            rep.add("warn", "F2", idx, f"本文が{size:g}ptで小さい: 「{r.text[:20]}」", "本文は12pt以上（出所・注記は9pt以上）", where)
        all_text += texts

        joined = "\n".join(texts)
        if has_graphic and not re.search(r"(出所|出典|Source)", joined) and idx not in (1,):
            if any(getattr(s, "has_chart", False) and s.has_chart for s in slide.shapes):
                rep.add("warn", "S1", idx, "グラフがあるのに出所の記載がない", "左下に「出所：〇〇（20XX年X月）」を入れる")
        if idx > 1 and 'type="sldNum"' not in slide_xml and not slide.slide_layout.placeholders and n > 3:
            pass  # ページ番号はレイアウト側にある場合が多いので layout.json 側で判定する

    # --- デッキ全体 ---
    if len(jp_fonts) > 2:
        rep.add("warn", "F3", 0, f"日本語フォントが{len(jp_fonts)}種類混在: {', '.join(jp_fonts)}", "見出し・本文あわせて1〜2書体にそろえる")
    if body_endings["desu"] >= 2 and body_endings["dearu"] >= 2:
        rep.add("warn", "W1", 0, f"本文で「です・ます」（{body_endings['desu']}箇所）と「だ・である」（{body_endings['dearu']}箇所）が混在",
                "本文の文末はどちらかに統一する（箇条書きは体言止めでもよい）")
    full = "\n".join(all_text)
    zen = sorted(set(ZENKAKU_ALNUM_RE.findall(full)))
    if zen:
        rep.add("warn", "W2", 0, f"全角の英数字がある: {''.join(zen)}", "英数字は半角にそろえる")
    if HANKAKU_KANA_RE.search(full):
        rep.add("warn", "W4", 0, "半角カタカナがある", "カタカナは全角にする")
    for a, b in YURE_PAIRS:
        if a in full and b in full:
            rep.add("warn", "W3", 0, f"表記ゆれ: 「{a}」と「{b}」が混在", f"どちらかに統一する（公用文では「{b}」が一般的）")


def phrase_boundaries(text):
    if _PARSER is None:
        return None
    pos, out = 0, {0}
    for ph in _PARSER.parse(text):
        pos += len(ph)
        out.add(pos)
    return out


def find_bad_breaks(para):
    """PowerPoint の実際の改行のうち、文節（BudouX）の途中で切れているものを返す。

    戻り値: [{kind: "L2"|"L4", break_at, insert_at, phrase, prev, next}]
      break_at  … 段落テキスト内で PowerPoint が改行した位置（文字オフセット）
      insert_at … ここに改行を入れれば文節ごと次の行へ送れる位置。直せない場合は None
    段落テキストは手動改行を除いた文字列（python-pptx の run テキストの連結と同じ）。
    """
    lines = [l for l in para["lines"] if l is not None]
    if len(lines) < 2:
        return []
    text = "".join(lines)
    bounds = phrase_boundaries(text)
    out, off, line_start = [], 0, 0
    for i in range(len(lines) - 1):
        line_start = off
        off += len(lines[i])
        prev, nxt = lines[i], lines[i + 1]
        if not prev or not nxt or prev[-1] in OK_BREAK_AFTER:
            continue
        number_split = bool(re.search(r"[0-9０-９][万億千]?$", prev)) and nxt.startswith(UNIT_START)
        if not number_split and (bounds is None or off in bounds):
            continue
        start = max(b for b in bounds if b < off) if bounds else off
        end = min(b for b in bounds if b > off) if bounds else off
        if number_split:  # 数字の頭まで戻す
            m = re.search(r"[0-9０-９.,，]+[万億千]?$", prev)
            start = min(start, off - len(m.group(0))) if m else start
        insert_at = start if start > line_start else None
        if insert_at is None:
            # 文節が1行より長い（例:「業務フロー確認・帳票収集」）→ 文節内の「・」「、」「／」の直後で切る
            for j in range(off - 1, line_start, -1):
                if text[j - 1] in "・、／/ 　":
                    insert_at = j
                    break
        out.append({
            "kind": "L4" if number_split else "L2", "break_at": off,
            "insert_at": insert_at,
            "phrase": text[start:end], "prev": prev, "next": nxt,
        })
    return out


def check_layout(layout, rep):
    for f in layout:
        sl, where = f["slide"], f["shape"]
        if "Slide Number" in where:
            continue
        # はみ出し（PowerPoint が計算した文字の高さ > 枠の内側の高さ）
        if f.get("autoSize", 0) == 0 and f["textH"] > f["innerH"] + 0.03 and f["innerH"] > 0:
            rep.add("error", "L1", sl, f"文字が枠からはみ出している（文字の高さ{f['textH']:.2f}in > 枠{f['innerH']:.2f}in）",
                    "文字を減らす・枠を広げる・スライドを分ける（本文を12pt未満にはしない）", where)
        for p in f["paragraphs"]:
            ea = (p.get("fontEA") or "").strip()
            if ea and ea.lower() in LATIN_ONLY_FONTS and JP_RE.search(p.get("text") or ""):
                rep.add("error", "F1", sl, f"日本語が欧文フォント（{ea}）で表示されている: 「{p['text'][:20]}」",
                        "東アジア用フォント（ea）に Meiryo / Yu Gothic / BIZ UDPゴシック などを指定する", where)
            for b in find_bad_breaks(p):
                ctx = f"「…{b['prev'][-8:]}｜{b['next'][:8]}…」"
                how = "jp_fix_wrap.py で自動修正できる" if b["insert_at"] is not None else "文節が1行より長いので、言い回しか枠幅を変える"
                if b["kind"] == "L4":
                    rep.add("error", "L4", sl, f"数字と単位が別の行に分かれている: {ctx}", f"「{b['phrase']}」を同じ行に収める（{how}）", where)
                else:
                    rep.add("error", "L2", sl, f"単語の途中で改行されている（泣き別れ）: {ctx}", f"「{b['phrase']}」の前で改行する（{how}）", where)
            lines = [l for l in p["lines"] if l]
            tail = re.sub(r"[、。）」]", "", lines[-1]) if lines else ""
            if 0 < len(tail) <= 2 and len(lines) >= 2:
                rep.add("warn", "L3", sl, f"最終行が{len(tail)}文字だけ: 「…{lines[-2][-6:]}｜{lines[-1]}」",
                        "文言を少し削るか改行位置を前に出して、行の長さをそろえる", where)


def main():
    ap = argparse.ArgumentParser(description="日本語ビジネス資料（.pptx）の品質チェッカー")
    ap.add_argument("pptx")
    ap.add_argument("--layout", help="export_layout.ps1 が出力した JSON")
    ap.add_argument("--json", action="store_true", help="JSON で出力")
    a = ap.parse_args()

    rep = Report()
    prs = Presentation(a.pptx)
    check_deck(prs, rep)
    if a.layout:
        with open(a.layout, encoding="utf-8-sig") as fh:
            check_layout(json.load(fh), rep)
    elif _PARSER is None:
        print("note: budoux 未インストールのため泣き別れチェックは無効", file=sys.stderr)

    rep.items.sort(key=lambda i: (i["slide"], 0 if i["level"] == "error" else 1, i["code"]))
    if a.json:
        print(json.dumps({"errors": rep.errors, "items": rep.items}, ensure_ascii=False, indent=2))
    else:
        if not rep.items:
            print("OK: 問題は見つかりませんでした")
        by = defaultdict(list)
        for i in rep.items:
            by[i["slide"]].append(i)
        for sl in sorted(by):
            print(f"\n== {'デッキ全体' if sl == 0 else f'スライド {sl}'} ==")
            for i in by[sl]:
                mark = "✖" if i["level"] == "error" else "△"
                loc = f" [{i['where']}]" if i["where"] else ""
                print(f" {mark} {i['code']}{loc} {i['message']}")
                if i["fix"]:
                    print(f"     → {i['fix']}")
        warns = len(rep.items) - rep.errors
        print(f"\nerror {rep.errors} 件 / warn {warns} 件")
    sys.exit(1 if rep.errors else 0)


if __name__ == "__main__":
    main()
