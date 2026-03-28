from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# Color palette
DARK_BLUE = RGBColor(0x1A, 0x37, 0x5E)   # 濃紺
MID_BLUE  = RGBColor(0x1F, 0x5C, 0x99)   # 中青
LIGHT_BLUE= RGBColor(0xD6, 0xE4, 0xF0)   # 薄青
ORANGE    = RGBColor(0xE8, 0x7A, 0x1E)   # アクセントオレンジ
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
GRAY      = RGBColor(0x60, 0x60, 0x60)
LIGHT_GRAY= RGBColor(0xF2, 0xF4, 0xF7)
GREEN     = RGBColor(0x27, 0x7A, 0x48)
RED       = RGBColor(0xC0, 0x39, 0x2B)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]  # blank layout

# ─────────────────────────────────────────────
# Helper functions
# ─────────────────────────────────────────────

def add_rect(slide, l, t, w, h, fill=None, line=None, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.line.fill.background()
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line and line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width or 1)
    else:
        shape.line.fill.background()
    return shape

def add_textbox(slide, text, l, t, w, h, size=12, bold=False, color=None,
                align=PP_ALIGN.LEFT, wrap=True, italic=False):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    return txb

def add_text_in_shape(shape, text, size=11, bold=False, color=WHITE, align=PP_ALIGN.CENTER):
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color

def slide_header(slide, title, subtitle=None):
    """Top banner with title"""
    add_rect(slide, 0, 0, 13.33, 1.1, fill=DARK_BLUE)
    add_rect(slide, 0, 1.0, 13.33, 0.08, fill=ORANGE)
    add_textbox(slide, title, 0.35, 0.12, 10, 0.75,
                size=28, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if subtitle:
        add_textbox(slide, subtitle, 0.35, 0.72, 10, 0.35,
                    size=13, color=LIGHT_BLUE, align=PP_ALIGN.LEFT)
    # slide bg
    bg = add_rect(slide, 0, 1.08, 13.33, 6.42, fill=LIGHT_GRAY)

def footer(slide, page_num, total):
    add_rect(slide, 0, 7.15, 13.33, 0.35, fill=DARK_BLUE)
    add_textbox(slide, f"大手ゼネコン5社 IT/DX部門 組織構成調査", 0.3, 7.17, 9, 0.28,
                size=9, color=LIGHT_BLUE)
    add_textbox(slide, f"{page_num} / {total}", 12.5, 7.17, 0.6, 0.28,
                size=9, color=WHITE, align=PP_ALIGN.RIGHT)

# ─────────────────────────────────────────────
# Slide 1: Title slide
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)

add_rect(slide, 0, 0, 13.33, 7.5, fill=DARK_BLUE)
add_rect(slide, 0, 3.2, 13.33, 0.1, fill=ORANGE)

add_textbox(slide, "大手ゼネコン5社", 1.0, 1.2, 11, 1.0,
            size=38, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(slide, "IT・DX部門 組織構成 調査レポート", 1.0, 2.2, 11, 0.8,
            size=28, bold=False, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

add_textbox(slide, "鹿島建設  ／  大林組  ／  清水建設  ／  大成建設  ／  竹中工務店",
            1.0, 3.5, 11, 0.5, size=15, color=WHITE, align=PP_ALIGN.CENTER)

add_textbox(slide, "調査日：2026年3月28日", 9.5, 6.8, 3.5, 0.4,
            size=11, color=LIGHT_BLUE, align=PP_ALIGN.RIGHT)

# ─────────────────────────────────────────────
# Slide 2: 概要比較表
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
slide_header(slide, "概要比較", "5社のIT/DX組織構成の全体像")
footer(slide, 2, 8)

# Table header
cols = [1.8, 2.4, 2.8, 1.5, 4.55]
headers = ["社名", "DX部門", "情報システム部門", "統合/分離", "特記事項"]
x = 0.3
for i, (w, h) in enumerate(zip(cols, headers)):
    r = add_rect(slide, x, 1.25, w - 0.05, 0.48, fill=MID_BLUE)
    add_text_in_shape(r, h, size=11, bold=True)
    x += w

# Table rows
rows = [
    ("鹿島建設",  "デジタル推進室\n（社長直轄）", "ITソリューション部",   "分離", "2021年DX専任組織新設。建設・事業・業務DXの3区分で推進"),
    ("大林組",    "DX本部\n（社長直轄）",          "DX本部に統合済み",      "統合", "2022年に情報システム・BIM・業務改革を一本化。約200名体制"),
    ("清水建設",  "DX経営推進室\n（社長直轄）",    "情報システム部\n（同室内）", "統合", "「DX注目企業2025」選定。DXコア人財120名・デジタル活用人財2,000名育成計画"),
    ("大成建設",  "DX戦略部\n（社長室傘下）",      "情報企画部\n（別組織）",    "分離", "業界初のCDO設置。情報子会社も活用した3層構造。「DX銘柄2025」選定"),
    ("竹中工務店","デジタル室",                    "デジタル室に統合",           "統合", "旧グループICT推進室を改組。AWS上に建設デジタルプラットフォーム構築"),
]

row_colors = [LIGHT_GRAY, WHITE, LIGHT_GRAY, WHITE, LIGHT_GRAY]
text_colors_flag = {"統合": GREEN, "分離": RED}

for ri, (row_data, bg) in enumerate(zip(rows, row_colors)):
    y = 1.73 + ri * 0.96
    rh = 0.92
    x = 0.3
    for ci, (w, cell) in enumerate(zip(cols, row_data)):
        r = add_rect(slide, x, y, w - 0.05, rh, fill=bg,
                     line=True, line_color=RGBColor(0xCC,0xCC,0xCC), line_width=0.5)
        # special color for 統合/分離
        if ci == 3:
            tc = text_colors_flag.get(cell, GRAY)
            add_text_in_shape(r, cell, size=11, bold=True, color=tc, align=PP_ALIGN.CENTER)
        elif ci == 0:
            add_text_in_shape(r, cell, size=11, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER)
        else:
            add_text_in_shape(r, cell, size=9, bold=False, color=GRAY, align=PP_ALIGN.CENTER)
        x += w

# ─────────────────────────────────────────────
# Slide 3: 鹿島建設
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
slide_header(slide, "鹿島建設", "DXと情報システム：分離型")
footer(slide, 3, 8)

# Badge
badge = add_rect(slide, 0.3, 1.25, 2.0, 0.38, fill=RED)
add_text_in_shape(badge, "分離型", size=13, bold=True)

# Org boxes
boxes = [
    ("社長直轄", 4.0, 1.25, 5.0, 0.38, MID_BLUE),
    ("デジタル推進室\n（Digital Promotion Office）", 3.0, 1.85, 3.2, 0.9, MID_BLUE),
    ("ITソリューション部\n（旧：情報システム部）", 7.0, 1.85, 3.2, 0.9, ORANGE),
]
for label, lx, ly, lw, lh, col in boxes:
    r = add_rect(slide, lx, ly, lw, lh, fill=col)
    add_text_in_shape(r, label, size=11, bold=True)

# Arrow line  (simple text arrow)
add_textbox(slide, "←  連携  →", 6.25, 2.1, 0.8, 0.4, size=9, color=GRAY, align=PP_ALIGN.CENTER)

# Detail boxes
details = [
    (3.0, 2.95, 3.2, "【デジタル推進室】\n・2021年1月 新設（社長直轄）\n・建設DX / 事業DX / 業務DX の3区分\n・土木・建築・事務・ITの多職種混成\n・デジタル人材育成メニューの構築\n・社内外デジタル課題の解決ハブ"),
    (7.0, 2.95, 3.2, "【ITソリューション部】\n・旧称：情報システム部\n・社内ICTインフラ整備・運用\n・情報システム管理\n・専務執行役員が管掌"),
]
for lx, ly, lw, txt in details:
    r = add_rect(slide, lx, ly, lw, 2.2, fill=WHITE,
                 line=True, line_color=RGBColor(0xBB,0xCC,0xDD), line_width=0.8)
    add_text_in_shape(r, txt, size=9, bold=False, color=GRAY, align=PP_ALIGN.LEFT)

add_textbox(slide, "▶ DX戦略立案 と IT基盤運用 を明確に分けた二層構造",
            0.3, 6.55, 12.0, 0.4, size=11, bold=True, color=DARK_BLUE)

# ─────────────────────────────────────────────
# Slide 4: 大林組
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
slide_header(slide, "大林組", "DXと情報システム：統合型（DX本部に一本化）")
footer(slide, 4, 8)

badge = add_rect(slide, 0.3, 1.25, 2.0, 0.38, fill=GREEN)
add_text_in_shape(badge, "統合型", size=13, bold=True)

# Top box
r = add_rect(slide, 4.5, 1.25, 4.0, 0.45, fill=DARK_BLUE)
add_text_in_shape(r, "DX本部（社長直轄） ── 約200名体制", size=12, bold=True)

# 3 pillars
pillar_data = [
    ("生産DX", 2.2, 2.1, 2.8, 3.5, MID_BLUE,
     "・BIM推進\n・現場デジタル化\n・自動化・ロボット化\n・スマート生産技術"),
    ("バックオフィスDX\n（旧：情報システム部門）", 5.3, 2.1, 2.8, 3.5, ORANGE,
     "・社内データ統合・活用\n・システムスリム化\n・業務自動化・省人化\n・生成AI（O-Bot Assistant）\n・クラウドシフト（2025年完了）"),
    ("情報セキュリティ", 8.4, 2.1, 2.8, 3.5, RGBColor(0x5D,0x6D,0x7E),
     "・サイバーセキュリティ強化\n・GRC統括\n・ICT投資の最適化"),
]
for label, lx, ly, lw, lh, col, detail in pillar_data:
    r = add_rect(slide, lx, ly, lw, 0.58, fill=col)
    add_text_in_shape(r, label, size=10, bold=True)
    r2 = add_rect(slide, lx, ly+0.58, lw, lh-0.58, fill=WHITE,
                  line=True, line_color=col, line_width=1.2)
    add_text_in_shape(r2, detail, size=9, bold=False, color=GRAY, align=PP_ALIGN.LEFT)

add_textbox(slide, "組織変遷：情報システムセンター（1991）→ グローバルICT推進室（2010）→ BIM推進室（2010）→ デジタル推進室（2020）→ DX本部（2022）",
            0.3, 6.0, 12.5, 0.45, size=9, italic=True, color=GRAY)
add_textbox(slide, "▶ 業界でも珍しく情報システムとDXを完全統合。200名規模の強力な専門組織",
            0.3, 6.5, 12.0, 0.38, size=11, bold=True, color=DARK_BLUE)

# ─────────────────────────────────────────────
# Slide 5: 清水建設
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
slide_header(slide, "清水建設", "DXと情報システム：統合型（一体運営）")
footer(slide, 5, 8)

badge = add_rect(slide, 0.3, 1.25, 2.0, 0.38, fill=GREEN)
add_text_in_shape(badge, "統合型", size=13, bold=True)

# Org hierarchy
r = add_rect(slide, 4.5, 1.25, 4.5, 0.45, fill=DARK_BLUE)
add_text_in_shape(r, "DX経営推進室（社長直轄）", size=12, bold=True)

sub_boxes = [
    (2.5, 2.05, 3.8, "情報システム部\n（室井俊一部長）",        MID_BLUE),
    (6.7, 2.05, 3.8, "基盤システム部\n（インフラ企画Gなど）",   ORANGE),
]
for lx, ly, lw, label, col in sub_boxes:
    r = add_rect(slide, lx, ly, lw, 0.65, fill=col)
    add_text_in_shape(r, label, size=10, bold=True)

# Detail cards
cards = [
    (1.5, 3.0, 4.5, "【中期DX戦略〈2024-2026〉の柱】\n①  経営視点でDXを推進する社長直轄組織の設置\n②  ワンストップでデータをつなげるプラットフォーム整備\n③  組織横断DX推進体制の構築\n④  事業部門間の人財ローテーション"),
    (6.5, 3.0, 5.5, "【人材育成目標（3年間）】\n・DXコア人財：120名\n  （データ×デジタルで業務変革・新規事業創出をリード）\n・デジタル活用人財：2,000名以上\n  （部門内業務改善を推進できるITツール活用層）"),
]
for lx, ly, lw, txt in cards:
    r = add_rect(slide, lx, ly, lw, 2.6, fill=WHITE,
                 line=True, line_color=RGBColor(0xBB,0xCC,0xDD), line_width=0.8)
    add_text_in_shape(r, txt, size=9, bold=False, color=GRAY, align=PP_ALIGN.LEFT)

# Award badge
award = add_rect(slide, 0.3, 3.1, 0.95, 1.8, fill=ORANGE)
add_text_in_shape(award, "DX\n注目\n企業\n2025", size=9, bold=True)

add_textbox(slide, "▶ 企画から運用まで一貫してDX経営推進室が担う統合モデル。経産省「DX注目企業2025」選定",
            0.3, 6.5, 12.5, 0.38, size=11, bold=True, color=DARK_BLUE)

# ─────────────────────────────────────────────
# Slide 6: 大成建設
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
slide_header(slide, "大成建設", "DXと情報システム：分離型（CDO・CIO設置、3層構造）")
footer(slide, 6, 8)

badge = add_rect(slide, 0.3, 1.25, 2.0, 0.38, fill=RED)
add_text_in_shape(badge, "分離型", size=13, bold=True)

# Layer structure
layers = [
    (3.5, 1.25, 6.5, 0.42, DARK_BLUE, "DX推進委員会  ／  CDO兼CIO（常務執行役員）"),
    (3.5, 1.82, 3.0, 0.42, MID_BLUE,  "DX戦略部（2024年新設）"),
    (6.8, 1.82, 3.0, 0.42, ORANGE,    "情報企画部"),
]
for lx, ly, lw, lh, col, label in layers:
    r = add_rect(slide, lx, ly, lw, lh, fill=col)
    add_text_in_shape(r, label, size=10, bold=True)

# Arrow between layers
add_textbox(slide, "連携", 6.6, 2.0, 0.7, 0.3, size=9, color=GRAY, align=PP_ALIGN.CENTER)

# Detail boxes
detail_data = [
    (1.0, 2.6, 3.5, "【DX戦略部】\n・全社横断DX戦略の策定・推進\n・デジタル活用による新規サービス創出\n・デジタル人財戦略\n・CDO直下の組織（社長室傘下）\n・社長室DX戦略部長：野村 淳 氏"),
    (5.0, 2.6, 3.5, "【情報企画部】\n・ICT戦略・企画・調達\n・GRC（ガバナンス・リスク・コンプライアンス）統括\n・全社IT投資評価・見える化\n・企画/コンサル/推進/IT調達の4室構成\n・情報子会社と連携・役割分担"),
    (9.0, 2.6, 3.8, "【大成建設ICTソリューションズ】\n（旧：大成情報システム）\n・2025年7月に社名変更\n・システム開発・保守・運用\n・グループ全体のICTサポート\n・情報企画部と一体で機能"),
]
for lx, ly, lw, txt in detail_data:
    r = add_rect(slide, lx, ly, lw, 3.0, fill=WHITE,
                 line=True, line_color=RGBColor(0xBB,0xCC,0xDD), line_width=0.8)
    add_text_in_shape(r, txt, size=9, bold=False, color=GRAY, align=PP_ALIGN.LEFT)

award = add_rect(slide, 0.3, 2.6, 0.55, 1.5, fill=MID_BLUE)
add_text_in_shape(award, "DX\n銘柄\n2025", size=8, bold=True)

add_textbox(slide, "業界初のCDO設置。DX戦略・IT基盤（情報企画部）・情報子会社の3層構造",
            0.3, 5.85, 12.5, 0.35, size=10, italic=True, color=GRAY)
add_textbox(slide, "▶ CDO/CIOを設置し、経営とITを直結。「DX銘柄2025」に建設業で初選定",
            0.3, 6.5, 12.5, 0.38, size=11, bold=True, color=DARK_BLUE)

# ─────────────────────────────────────────────
# Slide 7: 竹中工務店
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
slide_header(slide, "竹中工務店", "DXと情報システム：統合型（デジタル室）")
footer(slide, 7, 8)

badge = add_rect(slide, 0.3, 1.25, 2.0, 0.38, fill=GREEN)
add_text_in_shape(badge, "統合型", size=13, bold=True)

r = add_rect(slide, 4.0, 1.25, 5.0, 0.45, fill=DARK_BLUE)
add_text_in_shape(r, "デジタル室（2022年3月新設）\n岩下敬三 執行役員 デジタル室長", size=10, bold=True)

# Task force
r2 = add_rect(slide, 4.0, 1.9, 5.0, 0.45, fill=MID_BLUE)
add_text_in_shape(r2, "デジタル変革推進タスクフォース\n（本社管理系・プロジェクト系・各事業部デジタル推進責任者）", size=9, bold=False)

# 3 pillars
p_data = [
    (1.5,  2.7, 3.3, "【業務効率化 DX】\n・全業務のデジタル化\n・200以上のアプリをAWSへ移行\n・ITインフラコスト25%以上削減\n・新人事システムへのデジタル\n  アダプション導入"),
    (5.2,  2.7, 3.3, "【事業変革 DX】\n・建設デジタルプラットフォーム構築\n  （AWS上にデータ統合基盤）\n・営業/設計/見積/施工管理等\n  全データをBIで可視化・AI予測\n・非上場ながら先進的なICT戦略"),
    (8.9,  2.7, 3.8, "【デジタル人材育成】\n・事務系新入社員の配属ローテに\n  デジタル室を追加（2022年〜）\n・ITリテラシー・DXリテラシー・\n  データリテラシーを育成\n・旧グループICT推進室から\n  デジタル室へ改組・強化"),
]
for lx, ly, lw, txt in p_data:
    r = add_rect(slide, lx, ly, lw, 3.0, fill=WHITE,
                 line=True, line_color=RGBColor(0xBB,0xCC,0xDD), line_width=0.8)
    add_text_in_shape(r, txt, size=9, bold=False, color=GRAY, align=PP_ALIGN.LEFT)

add_textbox(slide, "組織変遷：グループICT推進室 → デジタル室（2022年3月）",
            0.3, 6.0, 9.0, 0.35, size=9, italic=True, color=GRAY)
add_textbox(slide, "▶ 非上場だが先進的なDX体制。情報システムとDXをデジタル室に統合、AWS全社基盤を構築",
            0.3, 6.5, 12.5, 0.38, size=11, bold=True, color=DARK_BLUE)

# ─────────────────────────────────────────────
# Slide 8: 総括・考察
# ─────────────────────────────────────────────
slide = prs.slides.add_slide(BLANK)
slide_header(slide, "総括・考察", "5社の傾向と示唆")
footer(slide, 8, 8)

findings = [
    ("社長直轄化（全社共通）",
     "5社すべてがDX推進組織を社長直轄または社長室傘下に配置。\n経営意思を直接現場に伝達する体制を重視。"),
    ("統合型 vs 分離型",
     "統合型（3社）：大林組・清水建設・竹中工務店\n分離型（2社）：鹿島建設・大成建設\n統合型が多数派。ただし分離型でも連携体制を整備。"),
    ("情報子会社の活用",
     "大成建設のみ情報子会社（ICTソリューションズ）を活用した3層構造。\n社名変更で「情報システム」から「ICTソリューション」へシフト。"),
    ("CDO/CIOの設置",
     "大成建設が業界初のCDO設置。他社はCIO兼任または役員管掌レベル。\nデジタルガバナンスの高度化が業界全体の課題。"),
    ("組織規模",
     "大林組が最大規模（DX本部 約200名）。\n他社は室・部レベル（数十名規模）。"),
    ("DX評価・外部認定",
     "大成建設：DX銘柄2025（建設業初）\n清水建設：DX注目企業2025\n各社が経産省認定を通じてDX成熟度を対外的にアピール。"),
]

x_positions = [0.3, 6.8]
y_start = 1.25
card_h = 1.55
card_w = 6.2

for i, (title, body) in enumerate(findings):
    col = i % 2
    row = i // 2
    lx = x_positions[col]
    ly = y_start + row * (card_h + 0.08)
    r = add_rect(slide, lx, ly, card_w, card_h, fill=WHITE,
                 line=True, line_color=MID_BLUE, line_width=0.8)
    title_r = add_rect(slide, lx, ly, card_w, 0.38, fill=MID_BLUE)
    add_text_in_shape(title_r, title, size=10, bold=True)
    add_text_in_shape(r, "\n" + body, size=9, bold=False, color=GRAY, align=PP_ALIGN.LEFT)

out_path = "/home/user/personal/ゼネコン5社_IT_DX組織構成.pptx"
prs.save(out_path)
print(f"Saved: {out_path}")
