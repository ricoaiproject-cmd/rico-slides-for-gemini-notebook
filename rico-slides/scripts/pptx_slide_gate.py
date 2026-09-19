# -*- coding: utf-8 -*-
"""PPTX スライド検証ゲート

slide-skill の DOM ゲート（scripts/qa_dom.py の TH）を PPTX 向けに移植したもの。
HTML と違い PPTX は全図形の座標・サイズ・フォントが数値で入っているので、
レンダリングせずに機械検査できる。

検査するゲート:
  G1 min_font        本文フォントが小さすぎる
  G2 off_canvas      図形がスライド外へはみ出している
  G3 text_overflow   テキスト量に対して図形が小さい（＝文字が枠からあふれる）
  G4 overlap         図形同士が重なりすぎている
  G5 palette         指定外の色が使われている
  G6 typeface        指定外のフォントが使われている
  G7 headline_len    見出しが規定文字数を超えている
  G8 sameness        全スライドが同じ配置の使い回しになっている（slide-skill の freshness）
  G9 notes           スピーカーノートが欠けている
  G10 rasterized     テキストが画像に焼かれている
  G15 headline_voice 本文スライドの見出しに強調語（大きさ・色の違う run）が無い／ラベル形（「：」区切り）（既定で有効・--no-voice で無効）
  G16 takeaway       本文スライドの下端に持ち帰り帯（Key 塗りの帯＋文）が無い（既定で有効・--no-voice で無効）

使い方:
  python pptx_slide_gate.py deck.pptx
  python pptx_slide_gate.py deck.pptx --palette 123A5E,E4572E,F4F6F8,333333,777777,FFFFFF \
                                      --font Meiryo --headline-max 26 --min-font 14
"""
import argparse
import collections
import math
import re
import sys
import unicodedata
import zipfile

EMU_PER_IN = 914400
EMU_PER_PT = 12700

# slide-skill scripts/qa_dom.py の TH から移植。
# 1280x720 CSS px 基準 → 13.333in x 7.5in の PPTX へ換算（1px = 0.75pt）。
TH = {
    "min_font_pt": 14.0,      # min_font_html 18px * 0.75 = 13.5pt → 14pt に丸め
    "overlap_frac": 0.22,     # 同値
    "edge_margin_in": 0.0,    # off_canvas。PPTX はスライド枠が絶対なので 0
    "overflow_ratio": 1.00,   # 必要高さ / 図形高さ がこれを超えたら溢れ
    "line_height": 1.35,      # 行送り係数
    "sameness_frac": 0.80,    # 配置シグネチャの一致率がこれ以上なら使い回し
    "orphan_chars": 2,        # 折り返し最終行がこの文字数以下なら孤立行
    "clearance_in": 0.30,     # 画像とテキストの最小クリアランス
    "edge_keepout_in": 0.40,  # 画像はスライド端からこれ以上内側
    "dead_space_in": 1.20,    # 最大の空白正方形がこれを超えたら死に空白
    "big_text_pt": 24.0,      # これ以上の級数は1行に収める（折り返したら G14）
}


def emu_in(v):
    return (v or 0) / EMU_PER_IN


def char_width(ch):
    """1文字の幅をフォントサイズ倍率で返す。全角=1.0 / 半角=0.5。

    フォントの実測が取れないときのフォールバック。
    """
    # Meiryo の英数字は 0.6〜0.68em。Notebook 側（Meiryo 無し）で甘く判定しないよう 0.62 を採る
    return 1.0 if unicodedata.east_asian_width(ch) in ("F", "W", "A") else 0.62


_FONT_CACHE = {}


def _metric_font(bold=True):
    """幅の実測に使う Meiryo。読めなければ None。"""
    key = "b" if bold else "r"
    if key not in _FONT_CACHE:
        name = "meiryob.ttc" if bold else "meiryo.ttc"
        try:
            from PIL import ImageFont
            _FONT_CACHE[key] = ImageFont.truetype(r"C:\Windows\Fonts" + "\\" + name, 72)
        except Exception:
            _FONT_CACHE[key] = None
    return _FONT_CACHE[key]


def text_em(text, bold=True):
    """文字列の幅を em（級数に対する倍率）で返す。

    半角=0.5 の近似は Meiryo の数字を大きく取り違える（実測 0.68em）。
    そのせいで「100名」が枠幅1.46in に収まると誤判定し、実際には折り返して
    単位が罫線をまたぐ欠陥を検出できなかった。フォントの実測を優先する。
    """
    f = _metric_font(bold)
    if f is not None:
        try:
            return f.getlength(text) / 72.0
        except Exception:
            pass
    return sum(char_width(c) for c in text)


def est_lines(text, font_pt, box_w_in, bold=True):
    """テキストが何行に折り返されるかの推定。"""
    if not text or font_pt <= 0 or box_w_in <= 0:
        return 0
    box_w_pt = box_w_in * 72.0
    lines = 0
    for para in text.split("\n"):
        if not para:
            lines += 1
            continue
        w = text_em(para, bold) * font_pt
        lines += max(1, math.ceil(w / box_w_pt))
    return lines


def load(path):
    from pptx import Presentation
    return Presentation(path)


A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _sz_of(el):
    """a:rPr / a:defRPr の sz（1/100 pt）を pt で返す。"""
    if el is None:
        return None
    v = el.get("sz")
    return int(v) / 100.0 if v else None


def _face_of(el):
    if el is None:
        return None
    for tag in ("latin", "ea", "cs"):
        n = el.find(A_NS + tag)
        if n is not None and n.get("typeface"):
            return n.get("typeface")
    return None


def shape_runs(shape):
    """(text, font_pt, font_name, color_hex) を返す。

    サイズ・書体は run の a:rPr に無いことが多く、その場合は
    段落の a:pPr/a:defRPr → 図形の a:lstStyle の順に継承をたどる。
    （Gemini Notebook が吐く pptx は pPr/defRPr にしか書かない）
    """
    out = []
    if not getattr(shape, "has_text_frame", False):
        return out
    tx = shape.text_frame._txBody
    lst = tx.find(A_NS + "lstStyle")
    lst_def = lst.find(A_NS + "lvl1pPr/" + A_NS + "defRPr") if lst is not None else None
    for p in tx.findall(A_NS + "p"):
        ppr = p.find(A_NS + "pPr")
        p_def = ppr.find(A_NS + "defRPr") if ppr is not None else None
        for r in p.findall(A_NS + "r"):
            t_el = r.find(A_NS + "t")
            text = t_el.text if t_el is not None and t_el.text else ""
            if not text.strip():
                continue
            rpr = r.find(A_NS + "rPr")
            sz = _sz_of(rpr) or _sz_of(p_def) or _sz_of(lst_def)
            face = _face_of(rpr) or _face_of(p_def) or _face_of(lst_def)
            col = None
            for src in (rpr, p_def, lst_def):
                if src is None:
                    continue
                c = src.find(A_NS + "solidFill/" + A_NS + "srgbClr")
                if c is not None:
                    col = (c.get("val") or "").upper()
                    break
            out.append((text, sz, face, col))
    return out


def effective_h(shape, w):
    """図形の高さ（in）。h=0 の自動フィット枠はテキスト量から推定する。

    Gemini Notebook は spAutoFit のテキスト枠を高さ0で書き出すことがある
    （PowerPoint が表示時に広げるので見た目は正常）。そのままだと「高さ0の
    装飾」とみなして占有判定から落ち、要素があるのに空白と誤判定する。
    """
    h = emu_in(shape.height)
    if h > 0.02 or not getattr(shape, "has_text_frame", False):
        return h
    li, ri, ti, bi = text_insets(shape)
    eff_w = max(0.10, w - li - ri)
    need = 0.0
    for t, sz, bd in shape_paras(shape):
        if not sz:
            continue
        need += est_lines(t, sz, eff_w, bd) * sz * TH["line_height"] / 72.0
    return need + ti + bi if need else h


def no_wrap(shape):
    """bodyPr が wrap="none"（折り返さない設定）かどうか。

    この枠は幅を超えても1行のまま伸びるので、折り返しの検査から外す。
    """
    if not getattr(shape, "has_text_frame", False):
        return False
    body = shape.text_frame._txBody.find(A_NS + "bodyPr")
    return body is not None and body.get("wrap") == "none"


def shape_paras(shape):
    """段落単位に (text, font_pt, bold) を返す。

    1つの段落を色分けで複数 run に割っても、表示上は1本の行として折り返される。
    run 単位で数えると「強調語ごとに改行している」とみなして溢れを誤検出する
    （Gemini Notebook は数値だけ橙にするため run を割る）。
    改行タグ a:br は "\\n" として残す。
    """
    out = []
    if not getattr(shape, "has_text_frame", False):
        return out
    tx = shape.text_frame._txBody
    lst = tx.find(A_NS + "lstStyle")
    lst_def = lst.find(A_NS + "lvl1pPr/" + A_NS + "defRPr") if lst is not None else None
    for p in tx.findall(A_NS + "p"):
        ppr = p.find(A_NS + "pPr")
        p_def = ppr.find(A_NS + "defRPr") if ppr is not None else None
        text, szs, bold = "", [], False
        for node in p:
            if node.tag == A_NS + "r":
                t_el = node.find(A_NS + "t")
                text += t_el.text if t_el is not None and t_el.text else ""
                rpr = node.find(A_NS + "rPr")
                sz = _sz_of(rpr) or _sz_of(p_def) or _sz_of(lst_def)
                if sz:
                    szs.append(sz)
                for src in (rpr, p_def, lst_def):
                    if src is not None and src.get("b") == "1":
                        bold = True
                        break
            elif node.tag == A_NS + "br":
                text += "\n"
        if text.strip():
            out.append((text, max(szs) if szs else None, bold))
    return out


def text_insets(shape):
    """bodyPr の lIns/rIns/tIns/bIns をインチで返す（既定 0.1in / 0.05in）。"""
    if not getattr(shape, "has_text_frame", False):
        return (0.1, 0.1, 0.05, 0.05)
    body = shape.text_frame._txBody.find(A_NS + "bodyPr")
    if body is None:
        return (0.1, 0.1, 0.05, 0.05)

    def g(k, d):
        v = body.get(k)
        return emu_in(int(v)) if v is not None else d

    return (g("lIns", 0.1), g("rIns", 0.1), g("tIns", 0.05), g("bIns", 0.05))


def collect_fills(path):
    """slideN.xml から srgbClr / typeface を直接拾う（継承分も取れる）。"""
    colors, fonts = collections.Counter(), collections.Counter()
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if re.match(r"ppt/slides/slide\d+\.xml$", n):
                x = z.read(n).decode("utf-8", "ignore")
                colors.update(c.upper() for c in re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', x))
                fonts.update(re.findall(r'typeface="([^"]+)"', x))
    return colors, fonts


def rect(sh):
    return (emu_in(sh.left), emu_in(sh.top), emu_in(sh.width), emu_in(sh.height))


def overlap_frac(a, b):
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    ix = max(0.0, min(ax + aw, bx + bw) - max(ax, bx))
    iy = max(0.0, min(ay + ah, by + bh) - max(ay, by))
    inter = ix * iy
    small = min(aw * ah, bw * bh)
    return (inter / small) if small > 0 else 0.0


def run_gates(path, palette, font, headline_max, min_font, voice=False, key_color=""):
    prs = load(path)
    n_total = len(prs.slides)
    W, H = emu_in(prs.slide_width), emu_in(prs.slide_height)
    findings = []
    sigs = []

    for i, slide in enumerate(prs.slides, 1):
        sig = []
        texts = []
        for sh in slide.shapes:
            if sh.left is None:
                continue
            x, y, w, h = rect(sh)
            sig.append((round(x, 1), round(y, 1), round(w, 1), round(h, 1)))

            # G2 off_canvas
            m = TH["edge_margin_in"]
            if x < -m or y < -m or x + w > W + m or y + h > H + m:
                findings.append((i, "G2 off_canvas",
                                 f"図形がスライド外 ({x:.2f},{y:.2f} {w:.2f}x{h:.2f} / 枠 {W:.2f}x{H:.2f})"))

            runs = shape_runs(sh)
            if runs:
                full = "\n".join(r[0] for r in runs)
                texts.append((full, x, y, w, h))
                sizes = [r[1] for r in runs if r[1]]
                # G1 min_font（装飾記号・ページ番号は読ませる文字ではないので除外）
                for t, sz, _, _ in runs:
                    body = t.strip()
                    decorative = len(body) <= 2 and not re.search(r"[0-9A-Za-zぁ-んァ-ヶ一-龥]{2,}", body)
                    if re.fullmatch(r"\d{1,3}", body) and y > 6.9 and x > 11.0:
                        decorative = True          # 右下のページ番号
                    if sz and sz < min_font and not decorative:
                        findings.append((i, "G1 min_font",
                                         f"{sz:.0f}pt (<{min_font:.0f}pt) 「{body[:24]}」"))
                # G3 text_overflow（bodyPr の内側余白を引いた実効サイズで判定）
                if sizes and w > 0 and h > 0 and not no_wrap(sh):
                    li, ri, ti, bi = text_insets(sh)
                    eff_w = w - li - ri
                    eff_h = h - ti - bi
                    need_in = 0.0
                    n_lines = 0
                    for t, sz, bd in shape_paras(sh):
                        s = sz or max(sizes)
                        ln = est_lines(t, s, eff_w, bd)
                        n_lines += ln
                        lh = 1.20 if ln == 1 else TH["line_height"]
                        need_in += ln * s * lh / 72.0
                    if eff_h > 0 and need_in / eff_h > TH["overflow_ratio"]:
                        findings.append((i, "G3 text_overflow",
                                         f"推定{n_lines}行={need_in:.2f}in > 実効枠{eff_h:.2f}in "
                                         f"(幅{eff_w:.2f}in) 「{full.strip()[:26]}」"))
                    # G14 big_text_wrap（大きな文字が枠幅に収まらず折り返されている）
                    # 枠が縦に大きいと G3 をすり抜けるが、数値と単位が上下に割れて
                    # 罫線をまたぐなど、見た目には最も目立つ壊れ方をする。
                    for t, sz, bd in shape_paras(sh):
                        if not sz or sz < TH["big_text_pt"]:
                            continue
                        ln = est_lines(t, sz, eff_w, bd)
                        if ln >= 2:
                            findings.append((i, "G14 big_text_wrap",
                                             f"{sz:.0f}pt が{ln}行に折り返し（枠幅{eff_w:.2f}in・"
                                             f"必要{text_em(t.strip(), bd) * sz / 72.0:.2f}in）"
                                             f"「{t.strip()[:20]}」"))

        # G4 overlap（テキストを持つ図形同士のみ。背景帯との重なりは対象外）
        for a in range(len(texts)):
            for b in range(a + 1, len(texts)):
                fa = overlap_frac(texts[a][1:], texts[b][1:])
                if fa > TH["overlap_frac"]:
                    findings.append((i, "G4 overlap",
                                     f"重なり{fa:.0%} 「{texts[a][0][:14]}」×「{texts[b][0][:14]}」"))

        # G11 orphan_line（折り返しの最終行が1〜2文字＝孤立行。組版として即失格）
        # 図形は全部見る。以前は x 座標が一致する最初の1つで打ち切っていて、
        # 注記の行末に句点だけが落ちる欠陥を取りこぼしていた。
        for sh in slide.shapes:
            if sh.left is None or not getattr(sh, "has_text_frame", False) or no_wrap(sh):
                continue
            li, ri, _, _ = text_insets(sh)
            eff_w = emu_in(sh.width) - li - ri
            if eff_w <= 0:
                continue
            for t, sz, bd in shape_paras(sh):
                if not sz or not t.strip():
                    continue
                n = est_lines(t, sz, eff_w, bd)
                if n < 2:
                    continue
                per = max(1, int((eff_w * 72.0) // sz))
                tail = len(t.strip()) - per * (n - 1)
                if 0 < tail <= TH["orphan_chars"]:
                    findings.append((i, "G11 orphan_line",
                                     f"最終行が{tail}文字だけ「{t.strip()[-6:]}」"
                                     f"（{n}行・1行{per}字）"))

        # G12 image_text_clearance（イラストとテキストの重なり・近接）
        for sh in slide.shapes:
            if sh.shape_type is None or sh.left is None:
                continue
            if "PICTURE" not in str(sh.shape_type):
                continue
            ir = rect(sh)
            for full, x, y, w, h in texts:
                f = overlap_frac(ir, (x, y, w, h))
                if f > 0.02:
                    findings.append((i, "G12 image_text_clearance",
                                     f"画像がテキストに{f:.0%}重なる「{full.strip()[:18]}」"))
                    continue
                gap_x = max(ir[0] - (x + w), x - (ir[0] + ir[2]))
                gap_y = max(ir[1] - (y + h), y - (ir[1] + ir[3]))
                gap = max(gap_x, gap_y)
                if gap < TH["clearance_in"] - 0.005:
                    findings.append((i, "G12 image_text_clearance",
                                     f"画像とテキストの間隔{gap:.2f}in "
                                     f"(<{TH['clearance_in']}in)「{full.strip()[:18]}」"))

        # G7 headline_len（そのスライドで最も大きいフォントのテキスト＝見出しとみなす）
        if texts:
            head = None
            best = -1
            for sh in slide.shapes:
                for t, sz, _, _ in shape_runs(sh):
                    if sz and sz > best:
                        best, head = sz, t
            if head and headline_max and len(head.strip()) > headline_max:
                findings.append((i, "G7 headline_len",
                                 f"{len(head.strip())}字 (>{headline_max}字) 「{head.strip()}」"))

        # G15/G16 見出しの声と持ち帰り帯（rico-slides の作法。--voice で有効）
        if voice and i >= 2:
            head_sh = None
            best = -1
            for sh in slide.shapes:
                if sh.left is None:
                    continue
                _, y0, _, _ = rect(sh)
                if y0 > 2.0:
                    continue          # 上部 2in に無いものは見出しではない（一点集中スライドは対象外）
                rr = shape_runs(sh)
                if len("".join(r[0] for r in rr).strip()) <= 2:
                    continue          # 巨大な薄い数字などの奥行き装飾は見出しではない
                for _, sz, _, _ in rr:
                    if sz and 20 <= sz <= 60 and sz > best:
                        best, head_sh = sz, sh
            if head_sh is not None:   # 見出しの無いスライド（一点集中）は G15/G16 の対象外
                hruns = shape_runs(head_sh)
                htext = "".join(r[0] for r in hruns).strip()
                if re.search(r"[：:]", htext):
                    findings.append((i, "G15 headline_voice",
                                     f"見出しがラベル形（「：」区切り）「{htext[:26]}」。言いたいことを文で書く"))
                elif len({r[1] for r in hruns if r[1]}) < 2 and len({r[3] for r in hruns if r[3]}) < 2:
                    findings.append((i, "G15 headline_voice",
                                     f"見出しに強調語が無い「{htext[:26]}」。accent= で意味を担う1語を渡す"))
                band_rect = False
                band_text = False
                for sh in slide.shapes:
                    if sh.left is None:
                        continue
                    x, y, w, h = rect(sh)
                    if y < 5.9:
                        continue
                    if w >= 8.0 and 0.4 <= h <= 1.2:
                        fill = None
                        try:
                            if sh.fill.type == 1:
                                fill = str(sh.fill.fore_color.rgb).upper()
                        except Exception:
                            fill = None
                        if not key_color or fill == key_color.upper():
                            band_rect = True
                    if shape_runs(sh) and len("".join(r[0] for r in shape_runs(sh)).strip()) >= 6:
                        band_text = True
                if not (band_rect and band_text):
                    findings.append((i, "G16 takeaway",
                                     "持ち帰り帯が無い。d.takeaway(s, 「だから何」の一文) を置く"))

        # G13 dead_space（列ごとに「要素が1つも無い横帯」を探す）
        #
        # 以前は最大の空白正方形を測っていたが、それだと大きな数値（72pt の「8.2%」など）の
        # 右にできる当然の余白まで「空白」として拾い、実態と合わなかった。
        # 見たいのは「列を上から下まで使えているか」なので、横帯の連続だけを見る。
        band = 0.10
        top, bottom = 1.30, 6.60
        nb = int((bottom - top) / band) + 1
        # 縦罫（細くて背の高い図形）があればそこで左右に分ける。無ければ中央。
        split = W / 2.0
        for sh in slide.shapes:
            if sh.left is None:
                continue
            sx, sy, sw, shh = rect(sh)
            if sw < 0.05 and shh > 3.0:
                split = sx
                break
        for ci, (cx0, cx1) in enumerate(((0.0, split), (split, W))):
            occ = [False] * nb
            for sh in slide.shapes:
                if sh.left is None:
                    continue
                x, y, w, h = rect(sh)
                h = effective_h(sh, w)
                # 装飾（細い線・端点ドット）は「内容」に数えない
                if min(w, h) < 0.08 or max(w, h) <= 0.12:
                    continue
                if x + w <= cx0 or x >= cx1:
                    continue
                for k in range(max(0, int((y - top) / band)),
                               min(nb, int((y + h - top) / band) + 1)):
                    occ[k] = True
            run = 0.0
            worst = 0.0
            for k in range(nb):
                run = 0.0 if occ[k] else run + band
                worst = max(worst, run)
            if worst > TH["dead_space_in"]:
                findings.append((i, "G13 dead_space",
                                 f"{'左' if ci == 0 else '右'}列に {worst:.2f}in の空き帯"
                                 f"（>{TH['dead_space_in']}in）。上から下まで要素が届いていない"))

        sigs.append(tuple(sorted(sig)))

    # G5/G6 palette & typeface
    colors, fonts = collect_fills(path)
    if palette:
        extra = [c for c in colors if c not in palette]
        if extra:
            findings.append((0, "G5 palette", f"指定外の色 {extra}"))
    if font:
        extra_f = [f for f in fonts if f != font]
        if extra_f:
            findings.append((0, "G6 typeface", f"指定外のフォント {extra_f}"))

    # G8 sameness（配置シグネチャの一致率）
    if len(sigs) >= 3:
        pairs = 0
        same = 0
        for a in range(len(sigs)):
            for b in range(a + 1, len(sigs)):
                pairs += 1
                sa, sb = set(sigs[a]), set(sigs[b])
                if sa and sb:
                    j = len(sa & sb) / len(sa | sb)
                    if j >= TH["sameness_frac"]:
                        same += 1
        if pairs and same / pairs >= 0.5:
            findings.append((0, "G8 sameness",
                             f"全{pairs}組中{same}組が同一配置 "
                             f"({same/pairs:.0%})。カードグリッドの使い回しになっている"))

    # G9 notes / G10 rasterized
    with zipfile.ZipFile(path) as z:
        n_notes = len([n for n in z.namelist() if re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", n)])
        n_media = len([n for n in z.namelist() if n.startswith("ppt/media/")])
    n_slides = len(prs.slides)
    if n_notes < n_slides:
        findings.append((0, "G9 notes", f"ノート {n_notes}/{n_slides} 枚。欠けている"))
    total_text = sum(len(shape_runs(sh)) for s in prs.slides for sh in s.shapes)
    if total_text == 0 and n_media > 0:
        findings.append((0, "G10 rasterized", "テキストが1つも無い＝画像に焼かれている"))

    return prs, findings, colors, fonts, n_notes, n_media


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--palette", default="", help="許可する色をカンマ区切り（例 123A5E,E4572E）")
    ap.add_argument("--font", default="", help="許可するフォント名（例 Meiryo）")
    ap.add_argument("--headline-max", type=int, default=0, help="見出しの最大文字数")
    ap.add_argument("--min-font", type=float, default=TH["min_font_pt"])
    ap.add_argument("--voice", action="store_true", help="（既定で有効）G15/G16（見出しの強調語・持ち帰り帯）を検査する。--palette の先頭色を Key とみなす")
    ap.add_argument("--no-voice", action="store_true", help="G15/G16 を検査しない（rico-slides 以外の pptx を見るとき）")
    a = ap.parse_args()

    palette = {c.strip().upper() for c in a.palette.split(",") if c.strip()}
    first = a.palette.split(",")[0].strip() if a.palette else ""
    prs, findings, colors, fonts, n_notes, n_media = run_gates(
        a.pptx, palette, a.font.strip(), a.headline_max, a.min_font, voice=not a.no_voice, key_color=first)

    print(f"■ {a.pptx}")
    print(f"  {len(prs.slides)}枚 / ノート{n_notes}枚 / 画像{n_media}点 / "
          f"色{len(colors)}種 / フォント{sorted(fonts)}")
    if not findings:
        print("\n  ✅ 全ゲート通過")
        return 0

    print(f"\n  ❌ {len(findings)}件")
    by = collections.defaultdict(list)
    for s, g, msg in findings:
        by[g].append((s, msg))
    for g in sorted(by):
        print(f"\n  [{g}] {len(by[g])}件")
        for s, msg in by[g][:10]:
            where = f"slide{s}" if s else "全体"
            print(f"    {where}: {msg}")
        if len(by[g]) > 10:
            print(f"    … 他 {len(by[g]) - 10} 件")
    return 1


if __name__ == "__main__":
    sys.exit(main())
