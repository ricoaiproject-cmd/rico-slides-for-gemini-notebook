# FORK_NOTES — rico-slides の出自

- 派生元: `jp-design-studio`（リコのAIラボ自作・Codex 用ポスター/チラシ/スライド制作スキル）。
- 派生日: 2026-09-14。目的: Gemini Notebook の裏側（ネット不可・Python＋python-pptx・Node なし）で動く**スライド専用**版。
- 持ち込んだもの: references の 01_layout / 02_typography / 03_color / 05_checklist を、スライド 16:9 と python-pptx の座標系に合わせて圧縮・書き直し。
  構図ライブラリは `videos/#GoogleVids_GeminiNotebook/Notebookデザイン指定書.md`（2026-08-13 v3・SlideComposition）を関数化。
  検品ゲート `scripts/pptx_slide_gate.py` は `tools/rico/lint/pptx_slide_gate.py`（同日）を同梱。
- 捨てたもの: HTML エディタ（templates/editor.html）、Playwright の描画検査（qa_render.js）、PSD/PNG 書き出し、GPT-Image 生成（gen_images.py）、印刷・入稿の知識（06_print_copy）、ポスターの迫力設計（キャッチ 90px 等）。
- 追加したもの: `scripts/slide_helpers.py`（構図 A〜I の描画関数）、Notebook 固有の注意（プレビュー豆腐・透かし・out へ1回）。
- 外したもの（令和8年9月19日）: `assets/bg_*.png` と `scripts/make_backgrounds.py`。手順書からもコードからも一度も参照されていない死蔵品だった。奥行きは Tint の図形で出しており、背景画像の出番が無い。
- ライセンス: MIT（自作）。同梱の pptx_slide_gate.py は slide-skill の閾値を移植したもの（自作）。

## v2〜v4（2026-09-14 夜）設計言語の作り直し
- slide-skill の美的原則（One hero / in-title accent / bottom takeaway band / whitespace budget / restraint）を python-pptx に移植。白地・見出しは文＋強調語1つ（30/40pt）・持ち帰り帯・奥行き1手（Tint）。
- 検品ゲートに G15 headline_voice・G16 takeaway を追加（`--voice`）。手順書に書くだけでは守られなかった作法を機械で強制する。
- 手順書に「存在する関数名はこれだけ」を明記（v1 の A/B で LLM が title_slide/headband_slide を捏造した対策）。
- Notebook への差し替えは別名 CSV でアップロードする（同名だと古い方が読まれる）。
