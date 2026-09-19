# -*- coding: utf-8 -*-
"""rico-slides 描画ヘルパー v2（python-pptx）

設計言語（references/01_layout.md「見た目の原則」）:
  - 白地。1枚に主役1つ。等分グリッドを既定にしない
  - 見出しは上に素で置き、意味を担う語だけ大きく・Accent 色（in-title accent）
  - 下に「持ち帰り」の帯を1本（Key 塗り・白文字・1文）
  - 奥行きは Tint（Key を薄めた色）の大きな図形を背景層に1つ。影・グラデ・3Dなし
  - 文字ブロックは見出し以外 7 個まで。余白を要素として使う

パレットは 7 色: Key / Accent / Surface / Body / Muted / White / Tint（Tint は Key から自動生成）。
フォントは Meiryo 固定。16:9（13.333in x 7.5in）。外周余白 0.7in。

最小例:
    from slide_helpers import Deck
    d = Deck(palette="ink-coral")
    d.cover("情報セキュリティ 基本の4か条", "新入社員研修", eyebrow="社内研修 2026", notes="…")
    s = d.slide("情報は4つに分類して扱う", accent="4つ", eyebrow="第3条", notes="…")
    d.checklist(s, ["極秘|経営判断に直結する情報", "秘|人事・契約・顧客の情報", "社外秘|社内限定の資料", "公開|誰が見てもよい情報"])
    d.takeaway(s, "自分が扱う情報が、どの区分かを先に確かめる")
    d.save("/workspace/scratch/deck.pptx")
"""
from __future__ import annotations

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

FONT = "Meiryo"
W, H = 13.333, 7.5
M = 0.7                      # 外周余白
HEAD_Y = 0.75                # 見出しの上端
BODY_Y = 2.0                 # 本文領域の上端
BAND_Y, BAND_H = 6.25, 0.7   # 持ち帰り帯
BODY_BOTTOM = BAND_Y - 0.25  # 本文領域の下端（帯の手前）

PALETTES = {
    # 3役（Key=濃い地色 / Surface=淡い面 / Accent=強調1点）で組む。Body/Muted は Key の色相に寄せた無彩色。
    # Accent が明るくて白地の文字に向かないときは accent_text（同じ色相で暗くした色）が自動で作られる。
    "ink-coral":     dict(key="27384A", accent="E2553F", surface="EEF1F4", body="262E36", muted="76808A", white="FFFFFF"),   # 既定。社内資料・IT・報告
    "navy-brass":    dict(key="1B3557", accent="C9992F", surface="F5F8FB", body="1E2833", muted="6C7885", white="FFFFFF"),   # 金融・士業・伝統企業
    "graphite-blue": dict(key="303B4C", accent="2F7FD1", surface="E6EBF1", body="222A34", muted="737E8C", white="FFFFFF"),   # テック・スタートアップ
    "charcoal-mint": dict(key="1D2430", accent="3BA776", surface="F4F7F9", body="1D2430", muted="6B7480", white="FFFFFF"),   # 建築・デザイン・不動産
    "mono-lime":     dict(key="18181B", accent="93C830", surface="F1F1F1", body="1A1A1A", muted="6E6E6E", white="FFFFFF"),   # 白黒＋ライム。動画映え・若い会社
    "moss-clay":     dict(key="2E6B4F", accent="C2452F", surface="F6F2E8", body="24352C", muted="6E7A72", white="FFFFFF"),   # 食・農・地域・環境
    "cocoa-gold":    dict(key="3A2E2C", accent="B9924A", surface="EDE2D2", body="2C2523", muted="7B706C", white="FFFFFF"),   # 高級・式典・ホテル・ブライダル
    "plum-ash":      dict(key="5A4A57", accent="D9553D", surface="E3DCCB", body="2E282C", muted="7D7078", white="FFFFFF"),   # 和・文化・自治体
    "ocean-sky":     dict(key="0F6F93", accent="1EA6DA", surface="F5F8FA", body="1C2B34", muted="697781", white="FFFFFF"),   # 医療・ヘルスケア・説明会
    "sand-terra":    dict(key="4A3F35", accent="D28B5A", surface="F4EFE7", body="2C2622", muted="7A7169", white="FFFFFF"),   # カフェ・雑貨・採用
}


def saturation(hex6: str) -> float:
    """色の鮮やかさ（0〜100）。10 種のパレットはどれも 70 未満に収まっている。"""
    import colorsys
    r, g, b = [int(hex6[i:i+2], 16) / 255 for i in (0, 2, 4)]
    return colorsys.rgb_to_hsv(r, g, b)[1] * 100


def contrast_ratio(fg: str, bg: str) -> float:
    """2色の明暗差（1.0〜21.0）。本文は 4.5 以上、大きな文字は 3.0 以上が必要。"""
    def rl(h):
        v = []
        for i in (0, 2, 4):
            c = int(h[i:i+2], 16) / 255
            v.append(c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4)
        return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2]
    a, b = rl(fg), rl(bg)
    if a < b:
        a, b = b, a
    return (a + 0.05) / (b + 0.05)


def readable_on(bg: str, want: str, fallback: str = "FFFFFF", need: float = 3.0) -> str:
    """地 bg の上に want を置いて読めるか確かめ、暗ければ fallback に替える。
    表紙のように地が Key のとき、差し色の強調語が沈むのを防ぐ
    （実測の失敗: 金赤 #E60033 の地に濃紺 #1B3557 の強調語を置いて読めなくなった・令和8年9月19日）。"""
    return want if contrast_ratio(want, bg) >= need else fallback


def tint(hex6: str, amount: float = 0.92) -> str:
    """Key を白に寄せた薄い色（amount=0.92 で 8% の色味）"""
    r, g, b = int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)
    f = lambda c: int(round(c + (255 - c) * amount))
    return f"{f(r):02X}{f(g):02X}{f(b):02X}"


def _lum(hex6: str) -> float:
    def ch(v):
        v = int(v, 16) / 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    return 0.2126 * ch(hex6[0:2]) + 0.7152 * ch(hex6[2:4]) + 0.0722 * ch(hex6[4:6])


def contrast(a: str, b: str) -> float:
    la, lb = _lum(a), _lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def darken_for_text(hex6: str, on: str = "FFFFFF", target: float = 3.0) -> str:
    """白地に載せる文字用: 同じ色相のまま、コントラスト target 以上になるまで暗くする（大きな文字の基準 3:1）"""
    r, g, b = int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)
    k = 1.0
    cur = hex6
    while contrast(cur, on) < target and k > 0.2:
        k -= 0.04
        cur = f"{int(r*k):02X}{int(g*k):02X}{int(b*k):02X}"
    return cur


def palette_arg(c: dict) -> str:
    """検品ゲートの --palette に渡す文字列（9色: 6色＋Tint＋Shade＋AccentText）"""
    return ",".join([c["key"], c["accent"], c["surface"], c["body"], c["muted"], c["white"],
                     c.get("tint", tint(c["key"])), c.get("shade", tint(c["key"], 0.10)),
                     c.get("accent_text", darken_for_text(c["accent"]))])


def rgb(hex6: str) -> RGBColor:
    return RGBColor.from_string(hex6.upper())


def chars_per_line(width_in: float, pt: float) -> int:
    return int(width_in * 72 / pt)


def height_for(lines: int, pt: float) -> float:
    return lines * pt * 1.35 / 72


class Deck:
    def __init__(self, palette: str = "ink-coral", custom: dict | None = None,
                 watermark_crop: float | None = None):
        # 生成画像の右下に入る透かしを隠すため、画像の下端を切る割合（既定 6%）
        self.watermark_crop = self.WATERMARK_CROP if watermark_crop is None else watermark_crop
        c = dict(custom) if custom else dict(PALETTES[palette])
        c.setdefault("tint", tint(c["key"]))
        c.setdefault("shade", tint(c["key"], 0.10))   # 表紙の縦帯用（Key を一段暗く見せる）
        c.setdefault("accent_text", darken_for_text(c["accent"]))   # 白地に置く強調語の色（Accent が明るいときだけ暗くなる）
        self.c = c
        self.prs = Presentation()
        self.prs.slide_width = Inches(W); self.prs.slide_height = Inches(H)
        self._blank = self.prs.slide_layouts[6]
        self.page = 0
        self._depths = {}

    # ---------------- 基本部品 ----------------
    def _shape(self, s, kind, x, y, w, h, fill: str, line: str | None = None, line_pt=1.0):
        shp = s.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
        shp.fill.solid(); shp.fill.fore_color.rgb = rgb(fill)
        if line:
            shp.line.color.rgb = rgb(line); shp.line.width = Pt(line_pt)
        else:
            shp.line.fill.background()
        shp.shadow.inherit = False
        shp.text_frame.text = ""
        return shp

    def rect(self, s, x, y, w, h, fill, line=None, line_pt=1.0):
        return self._shape(s, MSO_SHAPE.RECTANGLE, x, y, w, h, fill, line, line_pt)

    def round_rect(self, s, x, y, w, h, fill, radius=0.12):
        shp = self._shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill)
        shp.adjustments[0] = min(0.5, radius / min(w, h))
        return shp

    def oval(self, s, x, y, w, h, fill):
        return self._shape(s, MSO_SHAPE.OVAL, x, y, w, h, fill)

    def hline(self, s, x1, y, x2, color, pt=1.5):
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y), Inches(x2), Inches(y))
        ln.line.color.rgb = rgb(color); ln.line.width = Pt(pt); return ln

    def vline(self, s, x, y1, y2, color, pt=1.5):
        ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y1), Inches(x), Inches(y2))
        ln.line.color.rgb = rgb(color); ln.line.width = Pt(pt); return ln

    def text(self, s, x, y, w, h, content, pt, color, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.25, runs=None):
        """content: 文字列（\\n で改行）。runs を渡すと [(text, pt, color, bold), …] を1段落に並べる"""
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame; tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.04); tf.margin_top = tf.margin_bottom = Inches(0.02)
        tf.vertical_anchor = anchor
        if runs:
            p = tf.paragraphs[0]; p.alignment = align; p.line_spacing = spacing
            for t, rpt, rcol, rbold in runs:
                parts = t.split("\n")
                for k, part in enumerate(parts):
                    if k > 0:
                        p = tf.add_paragraph(); p.alignment = align; p.line_spacing = spacing
                    if part == "":
                        continue
                    r = p.add_run(); r.text = part; r.font.name = FONT; r.font.size = Pt(rpt); r.font.bold = rbold; r.font.color.rgb = rgb(rcol)
            return tb
        for i, ln in enumerate(str(content).split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align; p.line_spacing = spacing
            r = p.add_run(); r.text = ln; r.font.name = FONT; r.font.size = Pt(pt); r.font.bold = bold; r.font.color.rgb = rgb(color)
        return tb

    def _notes(self, s, notes):
        if notes:
            s.notes_slide.notes_text_frame.text = notes

    def _footer(self, s, label: str | None):
        """フッターはページ番号だけ（章名などの小さな文字は検品ゲートで本文扱いになるため置かない）"""
        self.page += 1
        self.text(s, W - M - 0.8, H - 0.45, 0.8, 0.3, str(self.page), 11, self.c["muted"], align=PP_ALIGN.RIGHT)

    def _depth(self, s, kind: str = "numeral:1"):
        """奥行き: 右に巨大な薄い1文字（項目数など意味のある数字）。それ以外の種類は無い（意味の無い図形は置かない）"""
        if not kind.startswith("numeral:"):
            return
        ch = kind.split(":", 1)[1][:1]
        shp = self.text(s, W - M - 3.6, BODY_Y - 0.2, 3.6, 3.4, ch, 200, self.c["tint"], bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP, spacing=1.0)
        self._depths[id(s)] = (kind, shp)

    def _drop_depth(self, s):
        """全幅を使う構図（2×2・左右対比・フロー・タイムライン・バー・数値）では奥行き図形が文字と重なるので取り除く"""
        kind, shp = self._depths.pop(id(s), (None, None))
        if shp is not None:
            el = shp._element; el.getparent().remove(el)

    def _depth_kind(self, s):
        return self._depths.get(id(s), (None, None))[0]

    # ---------------- スライドの骨格 ----------------
    def blank(self, notes=None):
        s = self.prs.slides.add_slide(self._blank); self._notes(s, notes); return s

    # この環境で作った画像は右下に「Gemini Notebook」の透かしが入るので、
    # 下側をこの割合だけ切り落としてから使う（令和8年9月19日・ユーザー指摘で追加）。
    # 透かしの無い画像（ユーザーが用意した写真など）に使っても、ほんの少し下が切れるだけで害はない。
    # 透かしごと見せたい・下端を切りたくないときは Deck(watermark_crop=0.0) にする。
    WATERMARK_CROP = 0.06

    def _crop_watermark(self, pic, path: str):
        """画像の下端を WATERMARK_CROP の割合だけ隠す。add_picture の後に呼ぶ。"""
        c = getattr(self, "watermark_crop", self.WATERMARK_CROP)
        if c <= 0:
            return pic
        # 既に下を切っている場合は足し込む（合計が 0.45 を超えないよう頭打ち）
        pic.crop_bottom = min(0.45, (pic.crop_bottom or 0) + c)
        return pic

    def picture_fill(self, s, path: str, x: float, y: float, w: float, h: float):
        """画像を枠 (x,y,w,h) いっぱいに敷く（縦横比を保ち、はみ出す分をトリミング）。写真・不透明のイラスト用。
        下端は透かしを避けるため少しだけ切り落とす（_crop_watermark）"""
        from PIL import Image
        with Image.open(path) as im:
            iw, ih = im.size
        pic = s.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
        box = w / h; src = iw / ih
        if src > box:
            cut = (1 - box / src) / 2; pic.crop_left = cut; pic.crop_right = cut
        elif src < box:
            cut = (1 - src / box) / 2; pic.crop_top = cut; pic.crop_bottom = cut
        return self._crop_watermark(pic, path)

    def picture_fit(self, s, path: str, x: float, y: float, w: float, h: float):
        """画像を枠の中に収める（縦横比を保ち、余白は透ける）。背景透過 PNG の切り抜きイラスト用。
        下端は透かしを避けるため少しだけ切り落とす（_crop_watermark）"""
        from PIL import Image
        with Image.open(path) as im:
            iw, ih = im.size
        box = w / h; src = iw / ih
        if src >= box:
            pw = w; ph = w / src
        else:
            ph = h; pw = h * src
        pic = s.shapes.add_picture(path, Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2), Inches(pw), Inches(ph))
        # 下を切った分だけ絵が縦に潰れないよう、枠の高さも同じ割合だけ詰める
        c = getattr(self, "watermark_crop", self.WATERMARK_CROP)
        if c > 0:
            self._crop_watermark(pic, path)
            pic.height = int(pic.height * (1 - c))
        return pic

    def cover(self, title: str, subtitle: str = "", eyebrow: str = "", notes=None, accent_word: str | None = None,
              image: str | None = None, cutout: str | None = None):
        """構図 A 表紙: Key 全面、左下に大きなタイトル。
        image=写真/不透明イラスト → 右 45% に敷く（縦横比を保ってトリミング）。
        cutout=背景透過 PNG → 右側に Key 地のまま浮かべる（枠なし）。両方は指定しない"""
        s = self.blank(notes)
        # 鮮やかすぎる Key（彩度 70 以上）を全面に塗ると刺さって見えるので、
        # 地を Surface にして Key は左の縦帯だけに絞る（令和8年9月19日・金赤 #E60033 の実測で追加）
        vivid = saturation(self.c["key"]) >= 70
        ground = self.c["surface"] if vivid else self.c["key"]
        on_ground = self.c["body"] if vivid else self.c["white"]
        self.rect(s, 0, 0, W, H, ground)
        if vivid:
            self.rect(s, 0, 0, 0.42, H, self.c["key"])       # 左の縦帯だけが主役色
        if image:
            self.picture_fill(s, image, W * 0.55, 0, W * 0.45, H)
            self.rect(s, W * 0.55, 0, 0.12, H, self.c["accent"])
            tw = W * 0.55 - M - 0.5
        elif cutout:
            self.picture_fit(s, cutout, W * 0.56, 0.9, W * 0.44 - M, H - 1.8)
            tw = W * 0.56 - M - 0.3
        elif not vivid:
            self.rect(s, W - M - 2.6, 0, 2.6, H, self.c["shade"])
            tw = 9.5
        else:
            tw = 9.5
        self.rect(s, M, 1.2, 0.5, 0.08, self.c["key"] if vivid else self.c["accent"])
        if eyebrow:
            self.text(s, M, 1.4, tw, 0.4, eyebrow, 14, on_ground)
        base, big = (44, 56) if not (image or cutout) else (38, 48)
        # 強調語が地に沈むなら替える（資料「Key の上に Accent を置かない」）
        # 淡い地なら Key を、Key の地なら差し色を使い、読めなければ白/本文色へ逃がす
        want = self.c["key"] if vivid else self.c["accent"]
        acc_col = readable_on(ground, want, on_ground)
        runs = self._accent_runs(title, accent_word, base, big, on_ground, acc_col) if accent_word else None
        self.text(s, M, 3.2, tw, 2.2, title, base, on_ground, bold=True, anchor=MSO_ANCHOR.BOTTOM, spacing=1.15, runs=runs)
        if subtitle:
            self.text(s, M, 5.55, tw, 0.6, subtitle, 20, on_ground)
        self.page += 1
        return s

    def _accent_runs(self, title: str, word: str, base_pt: float, big_pt: float, base_col: str, acc_col: str):
        if word and word not in title:
            raise ValueError(f"強調語 accent=「{word}」が見出し「{title}」に含まれていない。見出しの中の語を1つ選ぶ")
        if word and word in title:
            a, b = title.split(word, 1)
            runs = []
            if a: runs.append((a, base_pt, base_col, True))
            runs.append((word, big_pt, acc_col, True))
            if b: runs.append((b, base_pt, base_col, True))
            return runs
        return [(title, base_pt, base_col, True)]

    def slide(self, headline: str | None, accent: str | None = None, eyebrow: str | None = None, notes=None, depth: str | None = None, footer: str | None = None):
        """本文スライドの骨格: 白地・（奥行き図形）・アイブロー・見出し（accent 語だけ大きく色付き）・ページ番号
        headline=None にすると見出しなし（一点集中スライド用）"""
        if headline is not None:
            assert len(headline) <= 26, f"見出しは26字以内: {headline}"
        s = self.blank(notes)
        if depth:
            self._depth(s, depth)
        if eyebrow:
            self.oval(s, M, HEAD_Y + 0.05, 0.14, 0.14, self.c["accent"])
            self.text(s, M + 0.24, HEAD_Y - 0.09, 7, 0.4, eyebrow, 14, self.c["muted"])
        if headline is not None:
            hy = HEAD_Y + (0.4 if eyebrow else 0.0)
            hw = W - 2 * M - 3.2 if (depth and depth.startswith("numeral")) else W - 2 * M
            runs = self._accent_runs(headline, accent, 30, 40, self.c["key"], self.c["accent_text"]) if accent else None
            self.text(s, M, hy, hw, 0.9, headline, 30, self.c["key"], bold=True, anchor=MSO_ANCHOR.MIDDLE, spacing=1.1, runs=runs)
        self._footer(s, footer)
        return s

    def takeaway(self, s, sentence: str):
        """下の持ち帰り帯（1文・26字前後まで）"""
        self.round_rect(s, M, BAND_Y, W - 2 * M, BAND_H, self.c["key"], radius=0.18)
        self.oval(s, M + 0.3, BAND_Y + BAND_H / 2 - 0.09, 0.18, 0.18, self.c["accent"])
        self.text(s, M + 0.7, BAND_Y, W - 2 * M - 1.0, BAND_H, sentence, 18, self.c["white"], bold=True, anchor=MSO_ANCHOR.MIDDLE)

    # ---------------- 構図（本文領域 y 2.0〜6.0） ----------------
    def hero_number(self, s, value: str, unit: str, caption: str, side_items: list[tuple[str, str]]):
        """B ヒーロー数値: 左に 96pt の数値、右に補助指標を最大3つ（細い横線で区切る）"""
        self._drop_depth(s)
        self.text(s, M, BODY_Y + 0.1, 6.4, 2.2, value, 96, self.c["accent_text"], bold=True, anchor=MSO_ANCHOR.BOTTOM, spacing=1.0)
        self.text(s, M + 0.05, BODY_Y + 2.35, 6.0, 0.5, unit, 18, self.c["key"], bold=True)
        self.text(s, M, BODY_Y + 2.95, 6.2, 1.0, caption, 15, self.c["body"], spacing=1.35)
        x, y = 7.9, BODY_Y + 0.2
        for i, (label, val) in enumerate(side_items[:3]):
            if i: self.hline(s, x, y - 0.15, W - M, self.c["surface"], 1.0)
            self.text(s, x, y, 4.7, 0.38, label, 14, self.c["muted"])
            n = len(val)
            vpt = 26 if n <= 10 else (20 if n <= 14 else 16)     # 脇の値は 1 行に収める（10 字以内が望ましい）
            self.text(s, x, y + 0.33, 4.7, 0.75, val, vpt, self.c["key"], bold=True, anchor=MSO_ANCHOR.TOP)
            y += 1.3

    def two_column(self, s, left_title: str, left_lines: list[str], right_title: str, right_lines: list[str], emphasize_right=True):
        """C 左右対比: 左右とも Surface の面。右（後・推奨・例外）は Key の縦ルールと Key の見出しで一段強く"""
        self._drop_depth(s)
        colw, gap = 5.7, 0.55
        lx, rx = M, M + colw + gap
        nl = max(len(left_lines[:3]), len(right_lines[:3]), 1)
        block_h = 1.0 + nl * 0.58 + 0.6
        y0 = BODY_Y + max(0.15, (BODY_BOTTOM - BODY_Y - block_h) / 2)
        for x, title, lines, strong in ((lx, left_title, left_lines, not emphasize_right), (rx, right_title, right_lines, emphasize_right)):
            self.round_rect(s, x, y0, colw, block_h, self.c["surface"], radius=0.15)
            if strong:
                self.rect(s, x, y0 + 0.25, 0.1, block_h - 0.5, self.c["accent"])
            self.text(s, x + 0.4, y0 + 0.25, colw - 0.7, 0.6, title, 20, self.c["key"] if strong else self.c["muted"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
            self.text(s, x + 0.4, y0 + 0.95, colw - 0.7, nl * 0.5 + 0.3, "\n".join(lines[:3]), 16, self.c["body"], spacing=1.45)

    def checklist(self, s, items: list[str], columns: int = 1):
        """D チェックリスト: 「見出し|説明」の項目 3〜4 個。番号の丸（Key）＋太字の見出し（1行）＋その下に説明（1行・15pt）。
        右 3in は奥行き図形（numeral / circle）の置き場。説明は全角 30 字以内に削る（2行にしない）"""
        items = items[:4]
        x0 = M + 0.1
        dk = self._depth_kind(s)
        right = (W - M - 3.1) if dk else (W - M)
        avail = BODY_BOTTOM - BODY_Y - 0.2
        step = min(1.15, avail / len(items))
        y = BODY_Y + 0.2 + (avail - step * len(items)) / 2
        tx = x0 + 0.85
        for i, it in enumerate(items):
            head, _, desc = it.partition("|")
            self.oval(s, x0, y + 0.06, 0.52, 0.52, self.c["key"])
            self.text(s, x0, y + 0.06, 0.52, 0.52, str(i + 1), 16, self.c["white"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self.text(s, tx, y, right - tx, 0.5, head, 20, self.c["key"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
            if desc:
                self.text(s, tx, y + 0.5, right - tx, 0.42, desc, 15, self.c["body"], anchor=MSO_ANCHOR.TOP, spacing=1.2)
            if i < len(items) - 1:
                self.hline(s, tx, y + step - 0.08, right, self.c["surface"], 1.0)
            y += step

    def flow(self, s, steps: list[tuple[str, str]]):
        """E 横フロー: 3〜4 段階。番号の丸を線でつなぎ、下に見出しと説明。矢印の箱を並べない"""
        self._drop_depth(s)
        n = max(2, min(4, len(steps)))
        x0, x1 = M + 1.0, W - M - 1.0
        span = (x1 - x0) / (n - 1)
        y = BODY_Y + 0.9
        self.hline(s, x0, y + 0.4, x1, self.c["surface"], 8.0)
        for i, (title, desc) in enumerate(steps[:n]):
            cx = x0 + i * span
            self.oval(s, cx - 0.4, y, 0.8, 0.8, self.c["key"] if i < n - 1 else self.c["accent"])
            self.text(s, cx - 0.4, y, 0.8, 0.8, str(i + 1), 24, self.c["white"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            bw = min(3.4, span - 0.2)
            bx = min(max(cx - bw / 2, M), W - M - bw)
            self.text(s, bx, y + 1.1, bw, 0.6, title, 22, self.c["key"], bold=True, align=PP_ALIGN.CENTER)
            self.text(s, bx, y + 1.75, bw, 1.6, desc, 16, self.c["body"], align=PP_ALIGN.CENTER, spacing=1.35)

    def timeline(self, s, points: list[tuple[str, str]], accent_index: int = -1):
        """F タイムライン: 3〜5 点。上下交互。強調1点だけ Accent"""
        self._drop_depth(s)
        n = max(2, min(5, len(points)))
        tw = 2.6
        x0, x1 = M + tw / 2, W - M - tw / 2
        span = (x1 - x0) / (n - 1)
        y = BODY_Y + 2.15
        self.hline(s, x0 - 0.2, y, x1 + 0.2, self.c["key"], 2.0)
        for i, (date, desc) in enumerate(points[:n]):
            cx = x0 + i * span
            hit = (accent_index >= 0 and i == accent_index % n)
            self.oval(s, cx - 0.14, y - 0.14, 0.28, 0.28, self.c["accent"] if hit else self.c["key"])
            up = (i % 2 == 0)
            ty = y - 1.55 if up else y + 0.4
            self.text(s, cx - tw / 2, ty, tw, 0.5, date, 18, self.c["accent_text"] if hit else self.c["key"], bold=True, align=PP_ALIGN.CENTER)
            self.text(s, cx - tw / 2, ty + 0.5, tw, 0.95, desc, 15, self.c["body"], align=PP_ALIGN.CENTER, spacing=1.3)

    def bars(self, s, items: list[tuple[str, float, str]], accent_index: int = 0):
        """G 比例バー: 3〜6 項目。バー長は最大値に比例（最大 7.0in）"""
        self._drop_depth(s)
        items = items[:6]
        mx = max(v for _, v, _ in items) or 1
        avail = BODY_BOTTOM - BODY_Y
        step = min(0.8, avail / len(items))
        y = BODY_Y + (avail - step * len(items)) / 2
        for i, (label, val, shown) in enumerate(items):
            self.text(s, M, y, 3.2, step - 0.2, label, 15, self.c["body"], anchor=MSO_ANCHOR.MIDDLE)
            bl = 7.0 * val / mx
            self.round_rect(s, 4.1, y + 0.16, max(0.08, bl), step - 0.52, self.c["accent"] if i == accent_index else self.c["key"], radius=0.1)
            self.text(s, 4.1 + bl + 0.12, y, 1.6, step - 0.2, shown, 16, self.c["key"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
            y += step

    def statement(self, s, sentence: str, accent: str | None = None, note: str = ""):
        """H 一点集中: 中央に 36pt の一文。意味を担う語だけ 48pt・Accent"""
        runs = self._accent_runs(sentence, accent, 36, 48, self.c["key"], self.c["accent_text"]) if accent else None
        top, bottom = HEAD_Y + 0.6, BODY_BOTTOM
        block_h = 2.6 + (0.7 if note else 0)
        y0 = top + (bottom - top - block_h) / 2
        self.rect(s, W / 2 - 0.35, y0 - 0.25, 0.7, 0.08, self.c["accent"])
        self.text(s, M + 0.5, y0, W - 2 * M - 1.0, 2.6, sentence, 36, self.c["key"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, spacing=1.2, runs=runs)
        if note:
            self.text(s, M + 0.5, y0 + 2.7, W - 2 * M - 1.0, 0.7, note, 15, self.c["muted"], align=PP_ALIGN.CENTER)

    def quad(self, s, cards: list[tuple[str, str]]):
        """I 2×2 カード（1デッキ1回まで）: 大きな番号と短い本文。カード面は Surface、枠なし"""
        self._drop_depth(s)
        cw, ch, gap = 5.75, 1.75, 0.35
        for i, (title, body) in enumerate(cards[:4]):
            x = M + (i % 2) * (cw + gap); y = BODY_Y + 0.25 + (i // 2) * (ch + gap)
            self.round_rect(s, x, y, cw, ch, self.c["surface"], radius=0.15)
            self.text(s, x + 0.3, y + 0.2, 1.0, 0.8, f"{i + 1:02d}", 28, self.c["accent_text"] if i == 0 else self.c["key"], bold=True)
            self.text(s, x + 1.4, y + 0.28, cw - 1.7, 0.5, title, 18, self.c["key"], bold=True)
            self.text(s, x + 1.4, y + 0.85, cw - 1.7, ch - 1.0, body, 14, self.c["body"], spacing=1.3)

    def section(self, number: str, title: str, subtitle: str = "", notes=None, image: str | None = None, cutout: str | None = None):
        """J 章扉: Key 全面。左に巨大な章番号（Accent）、右にタイトル。image で右 40% に写真、cutout で透過 PNG を右に浮かべる"""
        s = self.blank(notes)
        self.rect(s, 0, 0, W, H, self.c["key"])
        if image:
            self.picture_fill(s, image, W * 0.66, 0, W * 0.34, H)
            self.rect(s, W * 0.66, 0, 0.12, H, self.c["accent"])
        elif cutout:
            self.picture_fit(s, cutout, W * 0.70, 1.4, W * 0.30 - M, H - 2.8)
        if image or cutout:          # 画像があるときは番号を一回り小さくして、題名の幅を確保する
            self.text(s, M, 1.9, 2.7, 3.0, number, 110, self.c["accent"], bold=True, anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)
            self.rect(s, 3.7, 2.6, 0.06, 2.3, self.c["shade"])
            tx, tw, tpt = 4.2, W * 0.66 - 0.4 - 4.2, 32
        else:
            self.text(s, M, 1.6, 3.6, 3.4, number, 150, self.c["accent"], bold=True, anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)
            self.rect(s, 4.9, 2.6, 0.06, 2.3, self.c["shade"])
            tx, tw, tpt = 5.4, W - M - 5.4, 36
        self.text(s, tx, 2.6, tw, 1.6, title, tpt, self.c["white"], bold=True, anchor=MSO_ANCHOR.BOTTOM, spacing=1.15)
        if subtitle:
            self.text(s, tx, 4.3, tw, 0.6, subtitle, 18, self.c["white"])
        self.page += 1
        return s

    def picture_text(self, s, image: str, lines: list[str], caption: str = "", cutout: bool = False):
        """N 画像＋文: 左 45% に画像、右に 2〜4 行の文（18pt）。
        cutout=True なら背景透過 PNG を Surface の面の上に浮かべる（枠なし）。False なら枠いっぱいに敷く。画像は文字なし"""
        self._drop_depth(s)
        iw, ih = 5.6, BODY_BOTTOM - BODY_Y - 0.1
        if cutout:
            self.round_rect(s, M, BODY_Y + 0.05, iw, ih, self.c["surface"], radius=0.18)
            self.picture_fit(s, image, M + 0.4, BODY_Y + 0.35, iw - 0.8, ih - 0.6)
        else:
            self.picture_fill(s, image, M, BODY_Y + 0.05, iw, ih)
        x = M + iw + 0.6
        nl = max(1, len(lines[:4]))
        y0 = BODY_Y + max(0.2, (ih - (nl * 0.62 + (0.6 if caption else 0))) / 2)
        self.text(s, x, y0, W - M - x, nl * 0.62 + 0.2, "\n".join(lines[:4]), 18, self.c["body"], spacing=1.6)
        if caption:
            self.text(s, x, y0 + nl * 0.62 + 0.35, W - M - x, 0.5, caption, 14, self.c["muted"])

    def kpi3(self, s, items: list[tuple[str, str, str]], accent_index: int = 0):
        """K 数値3並び: (数値, 単位, 説明) を 2〜3 個。数値 64pt、1つだけ Accent。細い縦罫で区切る"""
        self._drop_depth(s)
        items = items[:3]; n = len(items)
        colw = (W - 2 * M) / n
        y0 = BODY_Y + 0.4
        for i, (val, unit, cap) in enumerate(items):
            x = M + i * colw
            if i: self.vline(s, x, y0 + 0.2, y0 + 2.9, self.c["surface"], 2.0)
            col = self.c["accent_text"] if i == accent_index else self.c["key"]
            self.text(s, x + 0.3, y0, colw - 0.6, 1.5, val, 64, col, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM, spacing=1.0)
            self.text(s, x + 0.3, y0 + 1.55, colw - 0.6, 0.45, unit, 18, self.c["key"], bold=True, align=PP_ALIGN.CENTER)
            self.text(s, x + 0.3, y0 + 2.15, colw - 0.6, 0.9, cap, 15, self.c["body"], align=PP_ALIGN.CENTER, spacing=1.35)

    def table(self, s, headers: list[str], rows: list[list[str]], accent_col: int = -1):
        """L 表: 2〜4 列 × 3〜6 行。ヘッダは Key 塗り＋White、行は罫線のみ（縞や枠線で埋めない）。accent_col の列は Key 太字"""
        self._drop_depth(s)
        ncol = max(2, min(4, len(headers))); rows = rows[:6]
        x0, x1 = M, W - M
        widths = [(x1 - x0) / ncol] * ncol
        if ncol >= 2:
            widths[0] = (x1 - x0) * 0.3; rest = (x1 - x0 - widths[0]) / (ncol - 1)
            for k in range(1, ncol): widths[k] = rest
        rh = min(0.8, (BODY_BOTTOM - BODY_Y - 0.7) / max(1, len(rows)))
        y = BODY_Y + 0.15
        self.rect(s, x0, y, x1 - x0, 0.55, self.c["key"])
        x = x0
        for k in range(ncol):
            self.text(s, x + 0.15, y, widths[k] - 0.3, 0.55, headers[k], 15, self.c["white"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
            x += widths[k]
        y += 0.55
        for r in rows:
            x = x0
            for k in range(ncol):
                cell = r[k] if k < len(r) else ""
                strong = (k == accent_col)
                self.text(s, x + 0.15, y, widths[k] - 0.3, rh, cell, 15, self.c["key"] if strong else self.c["body"], bold=strong, anchor=MSO_ANCHOR.MIDDLE)
                x += widths[k]
            self.hline(s, x0, y + rh, x1, self.c["surface"], 1.5)
            y += rh

    def agenda(self, s, chapters: list[str]):
        """M 目次: 章番号（Accent の細字）と章名を縦に 3〜6 個。左に寄せ、右は空ける"""
        self._drop_depth(s)
        chapters = chapters[:6]
        avail = BODY_BOTTOM - BODY_Y - 0.2
        step = min(0.85, avail / len(chapters))
        y = BODY_Y + 0.2 + (avail - step * len(chapters)) / 2
        for i, name in enumerate(chapters):
            self.text(s, M, y, 1.0, step - 0.15, f"{i + 1:02d}", 24, self.c["accent_text"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
            self.text(s, M + 1.1, y, 8.0, step - 0.15, name, 22, self.c["key"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
            if i < len(chapters) - 1:
                self.hline(s, M + 1.1, y + step - 0.07, M + 9.1, self.c["surface"], 1.0)
            y += step

    def picture(self, s, path: str, x: float, y: float, w: float, h: float | None = None):
        """assets の画像を置く（文字なし・透かしなし前提）。h を省くと幅に合わせて縦横比を保つ"""
        if h is None:
            return s.shapes.add_picture(path, Inches(x), Inches(y), width=Inches(w))
        return s.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))

    def save(self, path: str, check: bool = True) -> int:
        """保存して、同梱の検品ゲートを自動実行し結果を印字する。戻り値は NG 件数（0 が合格）。
        check=False で検査を飛ばせるが、納品前の最後の save では必ず検査する"""
        self.prs.save(path)
        if not check:
            return -1
        import os, sys as _sys
        here = os.path.dirname(os.path.abspath(__file__))
        if here not in _sys.path:
            _sys.path.insert(0, here)
        import pptx_slide_gate as gate
        pal = {c.upper() for c in palette_arg(self.c).split(",")}
        prs, findings, colors, fonts, n_notes, n_media = gate.run_gates(path, pal, FONT, 26, 14.0, voice=True, key_color=self.c["key"])
        print(f"■ {path}")
        print(f"  {len(prs.slides)}枚 / ノート{n_notes}枚 / 画像{n_media}点 / 色{len(colors)}種 / フォント{sorted(fonts)}")
        if not findings:
            print("  ✅ 全ゲート通過（save 時に自動検品）")
            return 0
        print(f"  ❌ {len(findings)}件（save 時に自動検品）")
        import collections as _c
        by = _c.defaultdict(list)
        for sl, g, msg in findings:
            by[g].append((sl, msg))
        for g in sorted(by):
            print(f"  [{g}] {len(by[g])}件")
            for sl, msg in by[g][:10]:
                print(f"    {'slide' + str(sl) if sl else '全体'}: {msg}")
        return len(findings)

    def deliver(self, src: str, out_dir: str = "/workspace/out") -> str | None:
        """NG ゼロのときだけ out にコピーする。NG があれば None を返し、コピーしない"""
        import shutil, os
        ng = self.save(src)
        if ng != 0:
            print("  → NG が残っているため out にコピーしない。直してから save/deliver をやり直す")
            return None
        os.makedirs(out_dir, exist_ok=True)
        dst = os.path.join(out_dir, os.path.basename(src))
        shutil.copyfile(src, dst)
        print(f"  → 納品: {dst}")
        return dst
