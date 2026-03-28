from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image, ImageDraw
import io, math

# ── カラー定義 ──────────────────────────────────
NAVY    = RGBColor(0x1B, 0x3A, 0x6B)   # ENTAARネイビー
BLACK   = RGBColor(0x1A, 0x1A, 0x1A)
GRAY    = RGBColor(0x55, 0x55, 0x55)
LT_GRAY = RGBColor(0xAA, 0xAA, 0xAA)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
GREEN   = RGBColor(0x27, 0x7A, 0x48)
RED     = RGBColor(0xC0, 0x39, 0x2B)
BLUE_HL = RGBColor(0x1B, 0x3A, 0x6B)

W = 13.33   # slide width  (inches)
H = 7.5     # slide height (inches)
DPI = 96
PX_W = int(W * DPI)
PX_H = int(H * DPI)

# ── ドットグリッド背景画像を生成 ──────────────────
def make_dot_grid_bg(w_px, h_px, spacing=18, dot_r=1, dot_color=(210,215,220)):
    img = Image.new("RGB", (w_px, h_px), (255,255,255))
    draw = ImageDraw.Draw(img)
    for y in range(0, h_px, spacing):
        for x in range(0, w_px, spacing):
            draw.ellipse([x-dot_r, y-dot_r, x+dot_r, y+dot_r], fill=dot_color)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf

# ── 幾何学ウォーターマーク（タイトルスライド右側）─────
def make_watermark(w_px, h_px):
    """ENTAARロゴのような積み重なったシェブロン（灰色）"""
    img = Image.new("RGBA", (w_px, h_px), (255,255,255,0))
    draw = ImageDraw.Draw(img)
    col = (200, 205, 210, 180)
    cx, cy = int(w_px*0.72), int(h_px*0.5)
    size = int(h_px * 0.75)
    # 3段のシェブロン（右向き >>>> 積み重ね）
    for layer in range(3):
        offset = layer * int(size * 0.28)
        ox = cx - offset
        thick = int(size * 0.13)
        arm = int(size * 0.42)
        pts = [
            (ox,           cy - arm),
            (ox + size//3, cy),
            (ox,           cy + arm),
            (ox + thick,   cy + arm),
            (ox + size//3 + thick, cy),
            (ox + thick,   cy - arm),
        ]
        draw.polygon(pts, fill=col)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf

# ── ヘルパー関数 ──────────────────────────────────
def add_bg(slide, buf):
    buf.seek(0)
    pic = slide.shapes.add_picture(buf, Inches(0), Inches(0), Inches(W), Inches(H))
    slide.shapes._spTree.remove(pic._element)
    slide.shapes._spTree.insert(2, pic._element)

def add_rect(slide, l, t, w, h, fill=None, line_color=None, lw_pt=0):
    s = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    s.line.fill.background()
    if fill:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if line_color:
        s.line.color.rgb = line_color; s.line.width = Pt(lw_pt)
    else:
        s.line.fill.background()
    return s

def tb(slide, text, l, t, w, h, sz=11, bold=False, color=None,
        align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txb.word_wrap = wrap
    tf = txb.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    run = p.add_run(); run.text = text
    run.font.size = Pt(sz); run.font.bold = bold; run.font.italic = italic
    if color: run.font.color.rgb = color
    return txb

def shape_text(shape, lines, sz=10, bold=False, color=BLACK, align=PP_ALIGN.LEFT):
    """複数行テキストをshapeに挿入"""
    tf = shape.text_frame; tf.word_wrap = True
    from pptx.util import Pt as _Pt
    from pptx.oxml.ns import qn
    import copy
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        run = p.add_run(); run.text = line
        run.font.size = _Pt(sz); run.font.bold = bold
        run.font.color.rgb = color

def entaar_logo(slide, l, t, w=1.4):
    """ENTAARロゴ（テキストで近似）"""
    # アイコン（シェブロン）
    icon = add_rect(slide, l, t, 0.28, 0.28, fill=NAVY)
    tb(slide, ">|", l+0.01, t+0.01, 0.26, 0.26, sz=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, "ENTAAR", l+0.33, t+0.04, 1.1, 0.22, sz=14, bold=True, color=NAVY)

def footer_content(slide, page_num):
    tb(slide, "Company Confidential", 0.3, H-0.32, 3.5, 0.28, sz=8, color=LT_GRAY)
    tb(slide, f"Copyright ©2026 Entaar Inc. All rights reserved.  |  {page_num}",
       W-5.5, H-0.32, 5.2, 0.28, sz=8, color=LT_GRAY, align=PP_ALIGN.RIGHT)

def slide_title_bar(slide, title):
    """コンテンツスライド用タイトル（左ネイビーバー＋タイトル文字）"""
    add_rect(slide, 0.3, 0.25, 0.07, 0.55, fill=NAVY)
    tb(slide, title, 0.45, 0.22, W-0.6, 0.6, sz=20, bold=True, color=BLACK)
    # 区切り線
    add_rect(slide, 0.3, 0.85, W-0.6, 0.02, fill=RGBColor(0xDD,0xDD,0xDD))

# ═══════════════════════════════════════════════
#  プレゼンテーション作成
# ═══════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = Inches(W)
prs.slide_height = Inches(H)
BLANK = prs.slide_layouts[6]

dot_bg    = make_dot_grid_bg(PX_W, PX_H)
wm_img    = make_watermark(PX_W, PX_H)

# ─── Slide 1: タイトル ───────────────────────────
slide = prs.slides.add_slide(BLANK)
add_bg(slide, make_dot_grid_bg(PX_W, PX_H))

# ウォーターマーク（右側グレーシェブロン）
wm_buf = make_watermark(PX_W, PX_H)
wm_pic = slide.shapes.add_picture(wm_buf, Inches(0), Inches(0), Inches(W), Inches(H))

# 日付・会社名
tb(slide, "2026年3月28日", 0.5, 0.45, 5.0, 0.3, sz=10, color=GRAY)
tb(slide, "調査レポート", 0.5, 0.85, 5.0, 0.3, sz=11, bold=True, color=BLACK)

# メインタイトル
tb(slide, "大手ゼネコン5社", 0.5, 2.0, 8.0, 1.0, sz=36, bold=True, color=BLACK)
tb(slide, "IT・DX部門 組織構成 調査レポート", 0.5, 3.0, 9.0, 0.7, sz=24, bold=True, color=BLACK)

# ENTAARロゴ
entaar_logo(slide, 0.45, H-0.65)
tb(slide, "© Entaar, Inc.", W-1.8, H-0.38, 1.6, 0.28, sz=9, color=LT_GRAY, align=PP_ALIGN.RIGHT)

# ─── Slide 2: 概要比較表 ─────────────────────────
slide = prs.slides.add_slide(BLANK)
add_bg(slide, make_dot_grid_bg(PX_W, PX_H))
slide_title_bar(slide, "概要比較　── 5社のIT/DX組織構成")
footer_content(slide, 1)

# テーブルヘッダー
cols_w = [1.7, 2.2, 2.5, 1.2, 5.35]
headers = ["社名", "DX部門", "情報システム部門", "統合/分離", "特記事項"]
x = 0.3
for w_col, hdr in zip(cols_w, headers):
    r = add_rect(slide, x, 1.0, w_col-0.05, 0.4, fill=NAVY)
    tf = r.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    run = p.add_run(); run.text = hdr
    run.font.size = Pt(10); run.font.bold = True
    run.font.color.rgb = WHITE
    x += w_col

rows = [
    ("鹿島建設",   "デジタル推進室\n（社長直轄）",   "ITソリューション部",           "分離", "2021年DX専任組織新設。建設・事業・業務DXの3区分"),
    ("大林組",     "DX本部\n（社長直轄）",           "DX本部に統合",                  "統合", "2022年に情報システム・BIM・業務改革を一本化。約200名"),
    ("清水建設",   "DX経営推進室\n（社長直轄）",     "情報システム部\n（同室内）",    "統合", "DX注目企業2025選定。DXコア人財120名育成計画"),
    ("大成建設",   "DX戦略部\n（社長室傘下）",       "情報企画部\n（別組織）",        "分離", "業界初CDO設置。情報子会社含む3層構造。DX銘柄2025"),
    ("竹中工務店", "デジタル室",                    "デジタル室に統合",               "統合", "旧グループICT推進室を改組。AWS全社基盤を構築"),
]
row_bgs = [RGBColor(0xF7,0xF8,0xFA), WHITE, RGBColor(0xF7,0xF8,0xFA), WHITE, RGBColor(0xF7,0xF8,0xFA)]
flag_colors = {"統合": GREEN, "分離": RED}

for ri, (row_data, bg) in enumerate(zip(rows, row_bgs)):
    y = 1.4 + ri * 1.0
    rh = 0.96
    x = 0.3
    for ci, (w_col, cell) in enumerate(zip(cols_w, row_data)):
        r = add_rect(slide, x, y, w_col-0.05, rh, fill=bg,
                     line_color=RGBColor(0xDD,0xDD,0xDD), lw_pt=0.5)
        tf = r.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        run = p.add_run(); run.text = cell
        run.font.size = Pt(9 if ci not in (0,3) else 10)
        run.font.bold = (ci in (0,3))
        if ci == 3:
            run.font.color.rgb = flag_colors.get(cell, BLACK)
        elif ci == 0:
            run.font.color.rgb = NAVY
        else:
            run.font.color.rgb = GRAY
        x += w_col

# ─── Slide 3〜7: 各社スライド ──────────────────────
companies = [
    {
        "name": "鹿島建設",
        "tag": "分離型",
        "tag_color": RED,
        "page": 2,
        "summary": "DX戦略立案（デジタル推進室）と IT基盤運用（ITソリューション部）を明確に分けた二層構造",
        "boxes": [
            {
                "title": "デジタル推進室（社長直轄）",
                "color": NAVY,
                "items": [
                    "・2021年1月 新設（社長直轄）",
                    "・建設DX / 事業DX / 業務DX の3区分",
                    "・土木・建築・事務・ITの多職種混成",
                    "・デジタル人材育成メニューの構築",
                    "・社内外デジタル課題の解決ハブ",
                    "・中期経営計画（2024-2026）でもDXを重点テーマに設定",
                ]
            },
            {
                "title": "ITソリューション部（旧：情報システム部）",
                "color": RGBColor(0x2E,0x6D,0xAA),
                "items": [
                    "・社内ICTインフラ整備・運用",
                    "・情報システム管理",
                    "・専務執行役員が管掌",
                    "・セキュリティ基盤の整備",
                ]
            },
        ]
    },
    {
        "name": "大林組",
        "tag": "統合型",
        "tag_color": GREEN,
        "page": 3,
        "summary": "業界でも珍しく情報システムとDXを完全統合。200名規模の強力な専門組織（DX本部）を形成",
        "boxes": [
            {
                "title": "DX本部（社長直轄）　約200名",
                "color": NAVY,
                "items": [
                    "【統合前の3部門】",
                    "・旧 情報システム部門",
                    "・BIM生産基盤部門",
                    "・業務プロセス改革部門",
                    "",
                    "【変遷】情報システムセンター(1991)→ICT推進室(2010)→デジタル推進室(2020)→DX本部(2022)",
                ]
            },
            {
                "title": "DX本部 3本柱",
                "color": RGBColor(0x2E,0x6D,0xAA),
                "items": [
                    "① 生産DX",
                    "　 BIM推進・現場デジタル化・自動化",
                    "② バックオフィスDX",
                    "　 社内データ統合・生成AI（O-Bot）・クラウドシフト（2025年完了）",
                    "③ 情報セキュリティ",
                    "　 GRC統括・ICT投資最適化",
                ]
            },
        ]
    },
    {
        "name": "清水建設",
        "tag": "統合型",
        "tag_color": GREEN,
        "page": 4,
        "summary": "企画から運用まで一貫してDX経営推進室が担う統合モデル。経産省「DX注目企業2025」選定",
        "boxes": [
            {
                "title": "DX経営推進室（社長直轄）",
                "color": NAVY,
                "items": [
                    "【室内の部署】",
                    "・情報システム部（室井俊一部長）",
                    "　データ管理G・システム開発G",
                    "・基盤システム部",
                    "　インフラ企画G・運用G",
                    "",
                    "中期DX戦略〈2024-2026〉を策定・推進",
                ]
            },
            {
                "title": "人材育成目標（3年間）",
                "color": RGBColor(0x2E,0x6D,0xAA),
                "items": [
                    "① DXコア人財：120名",
                    "　 データ×デジタルで業務変革・",
                    "　 新規事業創出をリード",
                    "",
                    "② デジタル活用人財：2,000名以上",
                    "　 部門内業務改善を推進できる",
                    "　 ITツール活用層",
                    "",
                    "★ DX注目企業2025 選定（経産省）",
                ]
            },
        ]
    },
    {
        "name": "大成建設",
        "tag": "分離型",
        "tag_color": RED,
        "page": 5,
        "summary": "CDO/CIOを設置した業界初の事例。DX戦略・IT基盤・情報子会社の3層構造。「DX銘柄2025」選定",
        "boxes": [
            {
                "title": "DX推進委員会 ／ CDO兼CIO（常務執行役員）",
                "color": NAVY,
                "items": [
                    "【DX戦略部】2024年1月新設",
                    "・全社DX戦略策定・推進（社長室傘下）",
                    "・デジタル新規サービス創出",
                    "・デジタル人財戦略",
                    "",
                    "【情報企画部】",
                    "・ICT戦略・企画・調達・GRC統括",
                    "・IT投資評価（企画/コンサル/推進/調達の4室）",
                ]
            },
            {
                "title": "大成建設ICTソリューションズ（情報子会社）",
                "color": RGBColor(0x2E,0x6D,0xAA),
                "items": [
                    "・旧称：大成情報システム",
                    "・2025年7月 社名変更",
                    "・システム開発・保守・運用",
                    "・グループ全体のICTサポート",
                    "・情報企画部と一体で機能",
                    "",
                    "★ DX銘柄2025 選定（建設業初）",
                ]
            },
        ]
    },
    {
        "name": "竹中工務店",
        "tag": "統合型",
        "tag_color": GREEN,
        "page": 6,
        "summary": "非上場ながら先進的なDX体制。情報システムとDXをデジタル室に統合し、AWS全社基盤を構築",
        "boxes": [
            {
                "title": "デジタル室（2022年3月新設）",
                "color": NAVY,
                "items": [
                    "・旧 グループICT推進室 を改組",
                    "・岩下敬三 執行役員 デジタル室長",
                    "・デジタル変革推進タスクフォースの事務局",
                    "　（本社管理系・PJ系・各事業部DX推進責任者）",
                    "",
                    "【変遷】グループICT推進室 → デジタル室（2022）",
                ]
            },
            {
                "title": "建設デジタルプラットフォーム（AWS）",
                "color": RGBColor(0x2E,0x6D,0xAA),
                "items": [
                    "・200以上の業務アプリをAWSへ移行",
                    "・ITインフラコスト25%以上削減",
                    "・営業/設計/見積/施工/人事/経理を一元管理",
                    "・BIで可視化・AIで予測",
                    "",
                    "・新入社員（事務系）の配属ローテーションに",
                    "　デジタル室を追加（2022年〜）",
                ]
            },
        ]
    },
]

for co in companies:
    slide = prs.slides.add_slide(BLANK)
    add_bg(slide, make_dot_grid_bg(PX_W, PX_H))
    slide_title_bar(slide, co["name"])
    footer_content(slide, co["page"])

    # 統合/分離バッジ
    badge = add_rect(slide, 0.3, 0.92, 1.0, 0.28, fill=co["tag_color"])
    tf = badge.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    run = p.add_run(); run.text = co["tag"]
    run.font.size = Pt(9); run.font.bold = True; run.font.color.rgb = WHITE

    # サマリー行
    tb(slide, co["summary"], 1.42, 0.92, W-1.8, 0.3, sz=9, color=GRAY, italic=True)

    # 2列のカード
    card_w = 6.2
    positions = [0.3, 6.83]
    for ci, (box, lx) in enumerate(zip(co["boxes"], positions)):
        # カードヘッダー
        hdr = add_rect(slide, lx, 1.35, card_w, 0.4, fill=box["color"])
        tf = hdr.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        run = p.add_run(); run.text = box["title"]
        run.font.size = Pt(10); run.font.bold = True; run.font.color.rgb = WHITE

        # カード本文
        body = add_rect(slide, lx, 1.75, card_w, 5.3, fill=WHITE,
                        line_color=RGBColor(0xCC,0xCC,0xCC), lw_pt=0.5)
        tf = body.text_frame; tf.word_wrap = True
        for ii, item in enumerate(box["items"]):
            if ii == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run(); run.text = item
            run.font.size = Pt(9)
            run.font.color.rgb = BLACK if item else GRAY

# ─── Slide 9: 総括・考察 ──────────────────────────
slide = prs.slides.add_slide(BLANK)
add_bg(slide, make_dot_grid_bg(PX_W, PX_H))
slide_title_bar(slide, "総括・考察")
footer_content(slide, 7)

cards = [
    ("社長直轄化（全社共通）",
     "5社すべてがDX推進組織を社長直轄または社長室傘下に配置。\n経営意思を直接現場に伝達する体制を重視している。"),
    ("統合型 vs 分離型",
     "統合型（3社）：大林組・清水建設・竹中工務店\n分離型（2社）：鹿島建設・大成建設\n統合型が多数派だが、分離型でも連携体制を整備。"),
    ("情報子会社の活用",
     "大成建設のみ情報子会社（ICTソリューションズ）を活用した3層構造。\n社名変更で「情報システム」から「ICTソリューション」へシフト。"),
    ("CDO / CIO の設置",
     "大成建設が業界初のCDOを設置。他社はCIO兼任または役員管掌レベル。\nデジタルガバナンスの高度化が業界全体の課題。"),
    ("組織規模",
     "大林組が最大規模（DX本部 約200名）。\n他社は室・部レベル（数十名規模）で運営。"),
    ("DX外部認定",
     "大成建設：DX銘柄2025（建設業で初選定）\n清水建設：DX注目企業2025\n各社が経産省認定を通じてDX成熟度を対外的にアピール。"),
]

cw, ch = 4.1, 2.1
positions = [(0.3,1.0),(4.62,1.0),(8.93,1.0),(0.3,3.3),(4.62,3.3),(8.93,3.3)]
for (title, body), (lx, ly) in zip(cards, positions):
    hdr = add_rect(slide, lx, ly, cw, 0.38, fill=NAVY)
    tf = hdr.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    run = p.add_run(); run.text = title
    run.font.size = Pt(10); run.font.bold = True; run.font.color.rgb = WHITE

    body_box = add_rect(slide, lx, ly+0.38, cw, ch-0.38, fill=WHITE,
                        line_color=RGBColor(0xCC,0xCC,0xCC), lw_pt=0.5)
    tf = body_box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    run = p.add_run(); run.text = body
    run.font.size = Pt(9); run.font.color.rgb = GRAY

# ─── 保存 ─────────────────────────────────────────
out = "/home/user/personal/ゼネコン5社_IT_DX組織構成_Entaar.pptx"
prs.save(out)
print(f"Saved: {out}")
