from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm
import copy

# Color Palette
C_DARK = RGBColor(0x1A, 0x1A, 0x2E)      # Deep navy
C_PRIMARY = RGBColor(0x16, 0x21, 0x3E)    # Navy
C_ACCENT = RGBColor(0xE9, 0x4F, 0x37)     # Red-orange accent
C_ACCENT2 = RGBColor(0x0F, 0x3D, 0x6B)   # Mid blue
C_LIGHT = RGBColor(0xF5, 0xF5, 0xF5)      # Off white
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_GRAY = RGBColor(0x88, 0x88, 0x99)
C_YELLOW = RGBColor(0xFF, 0xC3, 0x00)
C_GREEN = RGBColor(0x27, 0xAE, 0x60)
C_BLUE_LIGHT = RGBColor(0x2E, 0x86, 0xAB)

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)

W = prs.slide_width
H = prs.slide_height

def add_rect(slide, x, y, w, h, fill=None, line=None, line_w=None):
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.line.fill.background()
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
        if line_w:
            shape.line.width = Pt(line_w)
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, x, y, w, h, size=14, bold=False, color=C_WHITE,
             align=PP_ALIGN.LEFT, wrap=True, italic=False):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.italic = italic
    return txBox

def add_multiline(slide, lines, x, y, w, h, size=13, bold=False, color=C_WHITE,
                  align=PP_ALIGN.LEFT, line_spacing=None):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
    return txBox

def slide_bg(slide, color=C_PRIMARY):
    bg = add_rect(slide, 0, 0, 13.33, 7.5, fill=color)
    return bg

def header_bar(slide, title, subtitle=None):
    add_rect(slide, 0, 0, 13.33, 1.2, fill=C_DARK)
    add_rect(slide, 0, 1.2, 0.08, 6.3, fill=C_ACCENT)
    add_text(slide, title, 0.25, 0.15, 11, 0.65, size=28, bold=True, color=C_WHITE)
    if subtitle:
        add_text(slide, subtitle, 0.25, 0.8, 11, 0.4, size=13, color=C_GRAY)

def page_num(slide, num, total=28):
    add_text(slide, f"{num} / {total}", 12.3, 7.1, 1.0, 0.35, size=10, color=C_GRAY, align=PP_ALIGN.RIGHT)

def accent_box(slide, x, y, w, h, title, body_lines, title_bg=C_ACCENT, body_bg=C_ACCENT2):
    add_rect(slide, x, y, w, 0.45, fill=title_bg)
    add_text(slide, title, x+0.1, y+0.05, w-0.2, 0.35, size=13, bold=True, color=C_WHITE)
    add_rect(slide, x, y+0.45, w, h-0.45, fill=body_bg)
    add_multiline(slide, body_lines, x+0.12, y+0.5, w-0.24, h-0.55, size=11.5, color=C_WHITE)

def arrow_right(slide, x, y):
    shape = slide.shapes.add_shape(13, Inches(x), Inches(y), Inches(0.5), Inches(0.35))
    shape.fill.solid()
    shape.fill.fore_color.rgb = C_ACCENT
    shape.line.fill.background()

# =====================================================================
# SLIDE 1: Cover
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide, C_DARK)
add_rect(slide, 0, 0, 0.12, 7.5, fill=C_ACCENT)
add_rect(slide, 0, 5.8, 13.33, 1.7, fill=C_ACCENT)

add_text(slide, "CONFIDENTIAL", 0.4, 0.3, 5, 0.4, size=11, color=C_GRAY, italic=True)
add_text(slide, "訪日海外インフルエンサー", 0.4, 1.2, 12, 0.85, size=42, bold=True, color=C_WHITE)
add_text(slide, "キャスティング事業", 0.4, 2.0, 12, 0.85, size=42, bold=True, color=C_WHITE)
add_text(slide, "事 業 計 画 書", 0.4, 2.95, 12, 0.6, size=28, bold=False, color=C_YELLOW)

add_rect(slide, 0.4, 3.75, 6, 0.04, fill=C_ACCENT)

add_text(slide, "海外インフルエンサー × 日本企業をつなぐコーディネート事業", 0.4, 3.95, 12, 0.5, size=16, color=RGBColor(0xCC,0xCC,0xDD))

add_text(slide, "2026年6月", 0.4, 6.0, 4, 0.4, size=14, color=C_WHITE, bold=True)
add_text(slide, "西村 代表", 9.0, 6.0, 4, 0.4, size=14, color=C_WHITE, bold=True, align=PP_ALIGN.RIGHT)

page_num(slide, 1)

# =====================================================================
# SLIDE 2: Table of Contents
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "目　次", "CONTENTS")

sections = [
    ("01", "エグゼクティブサマリー"),
    ("02", "市場背景と課題"),
    ("03", "当社の提供価値"),
    ("04", "ビジネスモデル"),
    ("05", "ターゲット顧客"),
    ("06", "インフルエンサー獲得戦略"),
    ("07", "インフルエンサー管理・評価設計"),
    ("08", "営業戦略・DMテンプレート"),
    ("09", "組織体制・業務フロー"),
    ("10", "KPI・数値目標"),
    ("11", "リスクと対策"),
    ("12", "将来構想"),
    ("13", "90日間アクションプラン"),
]

cols = [sections[:7], sections[7:]]
for ci, col in enumerate(cols):
    cx = 0.5 + ci * 6.5
    for ri, (num, title) in enumerate(col):
        ry = 1.4 + ri * 0.77
        add_rect(slide, cx, ry, 0.55, 0.55, fill=C_ACCENT)
        add_text(slide, num, cx+0.05, ry+0.05, 0.45, 0.45, size=16, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide, cx+0.6, ry+0.05, 5.6, 0.45, fill=C_ACCENT2)
        add_text(slide, title, cx+0.72, ry+0.1, 5.3, 0.35, size=13, color=C_WHITE, bold=True)

page_num(slide, 2)

# =====================================================================
# SLIDE 3: Executive Summary
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "01｜エグゼクティブサマリー", "Executive Summary")

# 3 big boxes
boxes = [
    ("事業概要", C_ACCENT, [
        "訪日予定の海外インフルエンサーを",
        "日本の飲食店・ホテル・観光施設へ",
        "キャスティングする仲介事業",
        "",
        "当社Instagramアカウントを",
        "インフルエンサー獲得チャネルとして活用",
    ]),
    ("なぜ今か", C_ACCENT2, [
        "✓ インバウンド需要 急増中",
        "✓ SNSによる海外集客ニーズ 高騰",
        "✓ 当社既存アセットで即参入可能",
        "✓ インフルエンサーからの自発的接触 増加",
        "✓ 競合少なく 先行者優位 を取れる",
    ]),
    ("期待成果", C_GREEN, [
        "3ヶ月：登録50名",
        "6ヶ月：月間5件受注",
        "12ヶ月：月間20件受注",
        "",
        "将来：日本最大級の",
        "訪日インフルエンサーDB構築",
    ]),
]

for i, (title, color, lines) in enumerate(boxes):
    bx = 0.4 + i * 4.3
    add_rect(slide, bx, 1.35, 4.0, 0.5, fill=color)
    add_text(slide, title, bx+0.12, 1.4, 3.76, 0.4, size=16, bold=True, color=C_WHITE)
    add_rect(slide, bx, 1.85, 4.0, 5.2, fill=RGBColor(0x1E,0x2D,0x50))
    add_multiline(slide, lines, bx+0.18, 1.95, 3.7, 5.0, size=13.5, color=C_WHITE)

page_num(slide, 3)

# =====================================================================
# SLIDE 4: Market Background
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "02｜市場背景と課題", "Market Background")

# 3 stakeholder problem boxes
stakeholders = [
    ("🍽️  飲食店の課題", C_ACCENT, [
        "❌ 海外集客したいが方法がない",
        "❌ 英語対応が難しい",
        "❌ インフルエンサーとの接点がない",
        "❌ 誰に頼めばよいかわからない",
    ]),
    ("🏨  ホテルの課題", C_ACCENT2, [
        "❌ インバウンド集客を強化したい",
        "❌ SNS露出を増やしたい",
        "❌ 海外向け認知を拡大したい",
        "❌ 効果測定が難しい",
    ]),
    ("📱  インフルエンサーの課題", RGBColor(0x6A,0x0D,0x83), [
        "❌ 日本の良い店舗を知らない",
        "❌ PR案件獲得が難しい",
        "❌ 日本語での交渉が難しい",
        "❌ 旅行中の情報収集に時間がかかる",
    ]),
]

for i, (title, color, lines) in enumerate(stakeholders):
    bx = 0.35 + i * 4.3
    add_rect(slide, bx, 1.35, 4.1, 0.52, fill=color)
    add_text(slide, title, bx+0.12, 1.38, 3.9, 0.46, size=14, bold=True, color=C_WHITE)
    add_rect(slide, bx, 1.87, 4.1, 4.0, fill=RGBColor(0x1A,0x26,0x44))
    add_multiline(slide, lines, bx+0.15, 1.98, 3.85, 3.7, size=13, color=C_WHITE)

# Bottom connector
add_rect(slide, 0.35, 6.0, 12.6, 0.06, fill=C_ACCENT)
add_text(slide, "→ これら3者の課題を「当社」が一括解決する", 0.5, 6.1, 12, 0.5, size=16, bold=True, color=C_YELLOW, align=PP_ALIGN.CENTER)

page_num(slide, 4)

# =====================================================================
# SLIDE 5: Value Proposition
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "03｜当社の提供価値", "Value Proposition")

# Center position label
add_rect(slide, 4.9, 3.0, 3.55, 1.5, fill=C_ACCENT)
add_text(slide, "当社", 4.9, 3.0, 3.55, 0.6, size=20, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
add_text(slide, "コーディネーター", 4.9, 3.6, 3.55, 0.55, size=13, color=C_WHITE, align=PP_ALIGN.CENTER)

# Left: Influencer
add_rect(slide, 0.3, 2.2, 4.0, 3.15, fill=C_ACCENT2)
add_text(slide, "海外インフルエンサー", 0.3, 2.2, 4.0, 0.5, size=14, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
add_multiline(slide, ["✅ 無料飲食体験", "✅ ホテル体験", "✅ 観光体験", "✅ 日本国内での案件紹介", "✅ 日本語サポート"], 0.42, 2.75, 3.7, 2.5, size=13, color=C_WHITE)

# Right: Restaurant/Hotel
add_rect(slide, 9.0, 2.2, 4.0, 3.15, fill=RGBColor(0x0F,0x3D,0x4A))
add_text(slide, "飲食店・ホテル・施設", 9.0, 2.2, 4.0, 0.5, size=14, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
add_multiline(slide, ["✅ 海外インフルエンサー紹介", "✅ 日程調整・来店管理", "✅ 投稿管理・効果測定", "✅ レポート提出", "✅ 英語コミュニケーション代行"], 9.12, 2.75, 3.7, 2.5, size=13, color=C_WHITE)

# Arrows
add_text(slide, "←→", 4.2, 3.45, 0.9, 0.6, size=22, bold=True, color=C_YELLOW, align=PP_ALIGN.CENTER)
add_text(slide, "←→", 8.4, 3.45, 0.9, 0.6, size=22, bold=True, color=C_YELLOW, align=PP_ALIGN.CENTER)

# Bottom tag line
add_rect(slide, 1.5, 5.7, 10.3, 0.65, fill=C_ACCENT)
add_text(slide, "「海外インフルエンサーと日本企業を繋ぐコーディネーター」", 1.5, 5.75, 10.3, 0.55, size=18, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

page_num(slide, 5)

# =====================================================================
# SLIDE 6: Business Model Overview
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "04｜ビジネスモデル", "Business Model — Phase Design")

phases = [
    ("Phase 1", "実績作り", "店舗負担：\n飲食代・宿泊費", "当社収益：なし\n→ 事例・投稿実績獲得", C_GRAY),
    ("Phase 2", "単発キャスティング", "クライアント負担：\nインフルエンサー費用", "¥30,000〜¥100,000 / 件", C_BLUE_LIGHT),
    ("Phase 3", "月額契約", "毎月インフルエンサー紹介\nレポート提出付き", "¥50,000〜¥300,000 / 月", C_ACCENT2),
    ("Phase 4", "自治体・観光施設", "高単価プロジェクト案件", "¥500,000〜 / 案件", C_ACCENT),
]

for i, (ph, title, desc, price, color) in enumerate(phases):
    bx = 0.4 + i * 3.2
    add_rect(slide, bx, 1.35, 3.0, 0.45, fill=color)
    add_text(slide, ph, bx+0.08, 1.38, 2.85, 0.38, size=13, bold=True, color=C_WHITE)
    add_rect(slide, bx, 1.8, 3.0, 0.55, fill=RGBColor(0x22,0x34,0x5A))
    add_text(slide, title, bx+0.08, 1.83, 2.85, 0.48, size=15, bold=True, color=C_YELLOW)
    add_rect(slide, bx, 2.35, 3.0, 1.5, fill=RGBColor(0x1A,0x26,0x44))
    add_multiline(slide, desc.split('\n'), bx+0.12, 2.42, 2.8, 1.35, size=12, color=C_WHITE)
    add_rect(slide, bx, 3.85, 3.0, 0.8, fill=color)
    add_text(slide, price, bx+0.08, 3.9, 2.85, 0.7, size=13, bold=True, color=C_WHITE)

    if i < 3:
        add_text(slide, "▶", bx+3.05, 2.95, 0.2, 0.4, size=16, bold=True, color=C_ACCENT, align=PP_ALIGN.CENTER)

# Revenue scale chart (simple bars)
add_rect(slide, 0.4, 5.0, 12.5, 0.06, fill=C_GRAY)
bar_vals = [0, 10, 40, 100]
bar_labels = ["Phase1", "Phase2", "Phase3", "Phase4"]
for i, (val, lbl) in enumerate(zip(bar_vals, bar_labels)):
    bx = 0.55 + i * 3.2
    bh = val / 100 * 1.5
    add_rect(slide, bx+0.5, 6.55 - bh, 2.0, bh, fill=[C_GRAY, C_BLUE_LIGHT, C_ACCENT2, C_ACCENT][i])
    add_text(slide, lbl, bx+0.3, 6.6, 2.4, 0.35, size=11, color=C_GRAY, align=PP_ALIGN.CENTER)

add_text(slide, "収益スケールイメージ →", 0.4, 4.85, 4, 0.3, size=11, color=C_GRAY, italic=True)

page_num(slide, 6)

# =====================================================================
# SLIDE 7: Target Customers
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "05｜ターゲット顧客", "Target Customers")

targets = [
    ("① ホテル", "最優先", C_ACCENT, [
        "• 予算規模が最も大きい",
        "• インバウンド強化ニーズ高",
        "• SNS露出の費用対効果を理解",
        "• 月額契約に移行しやすい",
    ]),
    ("② 高単価飲食店", "優先", C_ACCENT2, [
        "• インバウンド需要が高い",
        "• 外国人観光客の口コミ重視",
        "• 写真映えコンテンツと相性◎",
        "• 比較的決済が早い",
    ]),
    ("③ 観光施設", "中期", C_BLUE_LIGHT, [
        "• 集客単価が見えやすい",
        "• 地域PR予算を持つ",
        "• 自治体連携案件に発展可能",
    ]),
    ("④ 自治体", "長期", RGBColor(0x55,0x77,0x99), [
        "• 高単価・大型案件",
        "• 観光DMO連携",
        "• 実績積み上げ後に参入",
    ]),
]

for i, (title, priority, color, lines) in enumerate(targets):
    bx = 0.35 + i * 3.25
    add_rect(slide, bx, 1.35, 3.05, 0.5, fill=color)
    add_text(slide, title, bx+0.1, 1.37, 2.5, 0.42, size=15, bold=True, color=C_WHITE)
    add_rect(slide, bx+2.1, 1.35, 0.9, 0.5, fill=C_YELLOW)
    add_text(slide, priority, bx+2.1, 1.37, 0.9, 0.42, size=11, bold=True, color=C_DARK, align=PP_ALIGN.CENTER)
    add_rect(slide, bx, 1.85, 3.05, 4.0, fill=RGBColor(0x1A,0x26,0x44))
    add_multiline(slide, lines, bx+0.15, 1.95, 2.8, 3.75, size=13, color=C_WHITE)

page_num(slide, 7)

# =====================================================================
# SLIDE 8: Influencer Acquisition Strategy
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "06｜インフルエンサー獲得戦略", "Influencer Acquisition Flow")

steps = [
    ("STEP 1", "Instagramアカウント活用", "大阪観光コンテンツを発信し\n海外フォロワーを増加"),
    ("STEP 2", "フォロワーとの接触", "フォロー・いいね・コメント\nから旅行予定者を特定"),
    ("STEP 3", "旅行予定確認", "DMにて来日時期・\n滞在都市を確認"),
    ("STEP 4", "インサイト取得", "フォロワー数・エンゲージメント\n国籍・年齢層を収集"),
    ("STEP 5", "DB登録", "管理シートへ登録\nランク付けを実施"),
    ("STEP 6", "案件提案", "条件マッチする案件を\nDMにて提案"),
    ("STEP 7", "来店・体験", "日程調整・事前案内\n来店サポート"),
    ("STEP 8", "投稿・レポート", "投稿確認・保存\nクライアントへ報告"),
]

cols = 4
for i, (step, title, desc) in enumerate(steps):
    row = i // cols
    col = i % cols
    bx = 0.3 + col * 3.25
    by = 1.35 + row * 2.8
    add_rect(slide, bx, by, 3.0, 0.38, fill=C_ACCENT)
    add_text(slide, step, bx+0.08, by+0.04, 2.85, 0.3, size=11, bold=True, color=C_WHITE)
    add_rect(slide, bx, by+0.38, 3.0, 0.5, fill=C_ACCENT2)
    add_text(slide, title, bx+0.08, by+0.42, 2.85, 0.42, size=12, bold=True, color=C_WHITE)
    add_rect(slide, bx, by+0.88, 3.0, 1.7, fill=RGBColor(0x1A,0x26,0x44))
    add_multiline(slide, desc.split('\n'), bx+0.12, by+0.95, 2.8, 1.55, size=12, color=C_WHITE)

    # Arrow between steps in same row
    if col < cols - 1:
        add_text(slide, "▶", bx+3.03, by+0.6, 0.22, 0.4, size=14, bold=True, color=C_YELLOW, align=PP_ALIGN.CENTER)

# Down arrow between rows
add_text(slide, "▼", 6.5, 4.15, 0.5, 0.4, size=16, bold=True, color=C_YELLOW, align=PP_ALIGN.CENTER)

page_num(slide, 8)

# =====================================================================
# SLIDE 9: Influencer Management DB
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "07｜インフルエンサー管理設計", "Influencer Database Design")

db_items = [
    ("基本情報", C_ACCENT, [
        ("氏名", "個人特定・管理"),
        ("国籍", "ターゲット国マッチング"),
        ("Instagram URL", "アカウント確認"),
        ("メール", "連絡手段"),
        ("WhatsApp", "海外連絡主要手段"),
    ]),
    ("アカウント情報", C_ACCENT2, [
        ("フォロワー数", "影響力の基準"),
        ("平均再生数", "リーチ力評価"),
        ("平均いいね数", "エンゲージメント評価"),
        ("主要フォロワー国", "案件ターゲット合致確認"),
        ("男女比 / 年齢比率", "顧客層マッチング"),
    ]),
    ("旅行情報", C_BLUE_LIGHT, [
        ("来日時期", "案件提案のタイミング"),
        ("滞在都市", "エリアマッチング"),
        ("同行人数", "予約・手配の規模"),
        ("滞在日数", "案件組み込み計画"),
        ("訪問希望ジャンル", "マッチング精度向上"),
    ]),
]

for ci, (cat, color, rows) in enumerate(db_items):
    bx = 0.3 + ci * 4.35
    add_rect(slide, bx, 1.35, 4.15, 0.45, fill=color)
    add_text(slide, cat, bx+0.1, 1.37, 4.0, 0.4, size=14, bold=True, color=C_WHITE)
    # Column headers
    add_rect(slide, bx, 1.8, 4.15, 0.38, fill=RGBColor(0x0A,0x14,0x2E))
    add_text(slide, "取得項目", bx+0.1, 1.82, 2.0, 0.34, size=11, bold=True, color=C_GRAY)
    add_text(slide, "取得目的", bx+2.2, 1.82, 2.0, 0.34, size=11, bold=True, color=C_GRAY)
    for ri, (item, purpose) in enumerate(rows):
        ry = 2.18 + ri * 0.8
        bg = RGBColor(0x1A,0x26,0x44) if ri % 2 == 0 else RGBColor(0x1E,0x2D,0x52)
        add_rect(slide, bx, ry, 4.15, 0.75, fill=bg)
        add_rect(slide, bx, ry, 0.04, 0.75, fill=color)
        add_text(slide, item, bx+0.12, ry+0.15, 1.95, 0.45, size=12, bold=True, color=C_WHITE)
        add_text(slide, purpose, bx+2.15, ry+0.15, 2.05, 0.45, size=11, color=RGBColor(0xBB,0xCC,0xDD))

page_num(slide, 9)

# =====================================================================
# SLIDE 10: Influencer Evaluation Criteria
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "07｜インフルエンサー評価基準", "Influencer Ranking System")

ranks = [
    ("A ランク", "フォロワー 10万人以上", C_ACCENT, [
        "大手ホテル・自治体案件優先",
        "高単価キャスティング対象",
        "専任担当で丁寧にフォロー",
    ]),
    ("B ランク", "フォロワー 3万人以上", C_ACCENT2, [
        "高単価飲食店・中規模ホテル",
        "月額契約への移行を促進",
        "定期的な案件提供",
    ]),
    ("C ランク", "フォロワー 1万人以上", C_BLUE_LIGHT, [
        "Phase1の実績作り対象",
        "ニッチ訴求・特定ジャンル強み",
        "エンゲージメント率を重視評価",
    ]),
]

for i, (rank, threshold, color, lines) in enumerate(ranks):
    bx = 0.4 + i * 4.3
    # Big rank badge
    add_rect(slide, bx, 1.35, 1.3, 1.3, fill=color)
    add_text(slide, rank[0], bx+0.15, 1.38, 1.0, 0.9, size=48, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, rank[2:], bx+0.05, 2.25, 1.2, 0.35, size=12, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    add_rect(slide, bx+1.35, 1.35, 2.75, 0.5, fill=color)
    add_text(slide, rank, bx+1.45, 1.38, 2.6, 0.42, size=16, bold=True, color=C_WHITE)
    add_rect(slide, bx+1.35, 1.85, 2.75, 0.48, fill=RGBColor(0x22,0x34,0x5A))
    add_text(slide, threshold, bx+1.45, 1.88, 2.6, 0.4, size=13, bold=True, color=C_YELLOW)
    add_rect(slide, bx, 2.75, 4.1, 3.1, fill=RGBColor(0x1A,0x26,0x44))
    add_multiline(slide, ["活用方針:"] + [f"  ✓ {l}" for l in lines], bx+0.15, 2.85, 3.85, 2.8, size=13, color=C_WHITE)

# Evaluation KPIs
add_rect(slide, 0.4, 6.0, 12.5, 0.06, fill=C_ACCENT)
kpis = ["再生数", "エンゲージメント率", "ターゲット国", "投稿品質・世界観"]
add_text(slide, "共通評価指標：", 0.5, 6.12, 2.5, 0.42, size=13, bold=True, color=C_GRAY)
for ki, kpi in enumerate(kpis):
    add_rect(slide, 3.0 + ki * 2.45, 6.1, 2.3, 0.42, fill=C_ACCENT2)
    add_text(slide, kpi, 3.05 + ki * 2.45, 6.12, 2.2, 0.38, size=12, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

page_num(slide, 10)

# =====================================================================
# SLIDE 11: Sales Strategy
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "09｜営業戦略", "Sales Strategy")

sales = [
    ("🍽️ 飲食店向け", C_ACCENT, [
        "アプローチ方法",
        "・食べログ / Googleマップで高評価店をリスト化",
        "・DM / メールで初回接触（日本語）",
        "・「海外インフルエンサーに無料で紹介」訴求",
        "",
        "営業トーク例",
        "「現在、大阪を訪れる海外インフルエンサー（フォロワー数万〜数十万）への",
        "飲食店PR案件を展開しています。初回は完全無料でご紹介可能です。」",
    ]),
    ("🏨 ホテル向け", C_ACCENT2, [
        "アプローチ方法",
        "・インバウンド担当へ直接アポ",
        "・マーケ部門へ提案書を送付",
        "・宿泊無料提供＋SNS露出を提案",
        "",
        "営業トーク例",
        "「海外SNSで数万〜数十万フォロワーを持つ訪日インフルエンサーに",
        "貴ホテルを体験いただき、リールやストーリーで発信してもらいます。」",
    ]),
    ("🏯 観光施設向け", C_BLUE_LIGHT, [
        "アプローチ方法",
        "・施設担当・広報部門へメール",
        "・観光協会経由での紹介",
        "・自治体観光課へのアプローチ",
        "",
        "営業トーク例",
        "「訪日インフルエンサーに施設を体験していただき、",
        "海外への認知拡大・インバウンド集客強化を支援します。」",
    ]),
]

for i, (title, color, lines) in enumerate(sales):
    bx = 0.3 + i * 4.35
    add_rect(slide, bx, 1.35, 4.15, 0.48, fill=color)
    add_text(slide, title, bx+0.12, 1.38, 4.0, 0.42, size=15, bold=True, color=C_WHITE)
    add_rect(slide, bx, 1.83, 4.15, 5.35, fill=RGBColor(0x1A,0x26,0x44))
    add_multiline(slide, lines, bx+0.15, 1.9, 3.95, 5.15, size=11.5, color=C_WHITE)

page_num(slide, 11)

# =====================================================================
# SLIDE 12: DM Templates
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "10｜DMテンプレート", "DM & Email Templates")

templates = [
    ("📱 インフルエンサー初回DM（英語）", C_ACCENT, RGBColor(0x1A,0x26,0x44),
     'Hi [Name]! 😊\nWe noticed you\'re planning a trip to Japan!\nWe connect overseas influencers with the best restaurants, hotels,\nand attractions in Osaka — completely FREE.\nWould you be interested in a collaboration? 🍣🏯'),
    ("📊 インサイト回収DM（英語）", C_ACCENT2, RGBColor(0x1A,0x26,0x44),
     'Thanks for your interest! 🙌\nTo set up the best experience for you, could you share\nyour Instagram insights screenshot?\n(Followers / Reach / Audience country)\nThis helps us match you with the perfect partners!'),
    ("📧 飲食店向け営業メール（日本語）", C_BLUE_LIGHT, RGBColor(0x1A,0x26,0x44),
     '件名：海外インフルエンサーPRのご提案\n\n突然のご連絡失礼いたします。\n弊社は訪日予定の海外インフルエンサーと\n日本国内の飲食店・ホテルをマッチングする事業を展開しております。\n初回は完全無料でご案内可能です。\nお気軽にご返信ください。'),
]

for i, (title, color, bg, body) in enumerate(templates):
    by = 1.35 + i * 1.95
    add_rect(slide, 0.35, by, 4.0, 0.42, fill=color)
    add_text(slide, title, 0.48, by+0.04, 3.8, 0.35, size=12, bold=True, color=C_WHITE)
    add_rect(slide, 0.35, by+0.42, 12.6, 1.42, fill=bg)
    add_multiline(slide, body.split('\n'), 0.5, by+0.5, 12.3, 1.28, size=11.5, color=RGBColor(0xDD,0xEE,0xFF))

page_num(slide, 12)

# =====================================================================
# SLIDE 13: Org Chart
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "11｜組織体制", "Organization Structure")

# Org box - CEO
add_rect(slide, 4.7, 1.35, 3.9, 0.52, fill=C_ACCENT)
add_text(slide, "代表　西村", 4.7, 1.35, 3.9, 0.52, size=18, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

# Connector
add_rect(slide, 6.6, 1.87, 0.06, 0.5, fill=C_GRAY)
add_rect(slide, 2.5, 2.37, 8.3, 0.06, fill=C_GRAY)
add_rect(slide, 2.5, 2.37, 0.06, 0.5, fill=C_GRAY)
add_rect(slide, 10.74, 2.37, 0.06, 0.5, fill=C_GRAY)

# CEO Role box
add_rect(slide, 2.55, 2.88, 4.0, 0.42, fill=C_ACCENT)
add_text(slide, "代表（西村）の役割", 2.55, 2.88, 4.0, 0.42, size=13, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
add_rect(slide, 2.55, 3.3, 4.0, 2.9, fill=RGBColor(0x1A,0x26,0x44))
add_multiline(slide, [
    "✅ インフルエンサー獲得",
    "✅ クライアント営業",
    "✅ 商談・クロージング",
    "✅ 案件判断・意思決定",
    "✅ サービス設計",
], 2.7, 3.38, 3.7, 2.7, size=13, color=C_WHITE)

# Ops box
add_rect(slide, 6.8, 2.88, 4.0, 0.42, fill=C_ACCENT2)
add_text(slide, "運用担当の役割", 6.8, 2.88, 4.0, 0.42, size=13, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
add_rect(slide, 6.8, 3.3, 4.0, 2.9, fill=RGBColor(0x1A,0x26,0x44))
add_multiline(slide, [
    "✅ データベース管理",
    "✅ インサイト回収",
    "✅ 日程調整・来店サポート",
    "✅ 投稿確認・保存",
    "✅ レポート作成・提出",
], 6.95, 3.38, 3.7, 2.7, size=13, color=C_WHITE)

# Bottom note
add_rect(slide, 0.5, 6.45, 12.3, 0.65, fill=RGBColor(0x22,0x34,0x5A))
add_text(slide, "★ スモールスタート2名体制から開始。案件増加に応じて業務委託・採用を検討", 0.6, 6.5, 12.1, 0.55, size=13, color=C_YELLOW, bold=True)

page_num(slide, 13)

# =====================================================================
# SLIDE 14: Business Flow
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "12｜業務フロー", "Business Operation Flow")

flow_steps = [
    ("①", "インフルエンサー獲得", "代表"),
    ("②", "旅行予定確認", "代表"),
    ("③", "インサイト回収", "運用"),
    ("④", "案件提案", "代表"),
    ("⑤", "来店調整", "運用"),
    ("⑥", "来店・投稿", "インフルエンサー"),
    ("⑦", "レポート提出", "運用"),
]

for i, (num, title, owner) in enumerate(flow_steps):
    bx = 0.3 + i * 1.85
    color = C_ACCENT if owner == "代表" else (C_ACCENT2 if owner == "運用" else C_GREEN)
    add_rect(slide, bx, 1.5, 1.65, 0.45, fill=color)
    add_text(slide, f"{num} {owner}", bx+0.08, 1.52, 1.5, 0.4, size=11, bold=True, color=C_WHITE)
    add_rect(slide, bx, 1.95, 1.65, 4.3, fill=RGBColor(0x1A,0x26,0x44))
    add_rect(slide, bx, 1.95, 0.06, 4.3, fill=color)
    add_text(slide, title, bx+0.15, 3.6, 1.45, 0.8, size=12, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    if i < 6:
        add_text(slide, "▶", bx+1.68, 3.7, 0.2, 0.4, size=14, bold=True, color=C_YELLOW, align=PP_ALIGN.CENTER)

# Legend
add_rect(slide, 0.3, 6.55, 12.7, 0.06, fill=C_GRAY)
legend = [("■ 代表", C_ACCENT), ("■ 運用担当", C_ACCENT2), ("■ インフルエンサー", C_GREEN)]
for li, (label, color) in enumerate(legend):
    add_rect(slide, 2.5 + li * 3.2, 6.65, 0.25, 0.3, fill=color)
    add_text(slide, label, 2.8 + li * 3.2, 6.65, 2.5, 0.3, size=11, color=C_GRAY)

page_num(slide, 14)

# =====================================================================
# SLIDE 15: KPI
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "13｜KPI設計", "KPI Design")

# CEO KPIs
add_rect(slide, 0.35, 1.35, 6.0, 0.48, fill=C_ACCENT)
add_text(slide, "代表（西村）KPI", 0.5, 1.38, 5.8, 0.42, size=16, bold=True, color=C_WHITE)

ceo_kpis = [
    ("新規インフルエンサー獲得数", "月10名〜", "認知・DB成長"),
    ("商談数", "月5件〜", "営業活動量"),
    ("受注数", "月3件〜", "収益化KPI"),
    ("月間売上", "目標額達成率", "事業成立確認"),
]

for ri, (kpi, target, note) in enumerate(ceo_kpis):
    ry = 1.85 + ri * 0.82
    bg = RGBColor(0x1A,0x26,0x44) if ri % 2 == 0 else RGBColor(0x1E,0x2D,0x52)
    add_rect(slide, 0.35, ry, 6.0, 0.78, fill=bg)
    add_rect(slide, 0.35, ry, 0.06, 0.78, fill=C_ACCENT)
    add_text(slide, kpi, 0.5, ry+0.08, 2.8, 0.62, size=12, bold=True, color=C_WHITE)
    add_text(slide, target, 3.4, ry+0.08, 1.6, 0.62, size=13, bold=True, color=C_YELLOW, align=PP_ALIGN.CENTER)
    add_text(slide, note, 5.1, ry+0.08, 1.2, 0.62, size=11, color=C_GRAY)

# Ops KPIs
add_rect(slide, 6.9, 1.35, 6.0, 0.48, fill=C_ACCENT2)
add_text(slide, "運用担当 KPI", 7.05, 1.38, 5.8, 0.42, size=16, bold=True, color=C_WHITE)

ops_kpis = [
    ("インサイト回収率", "80%以上", "DB品質管理"),
    ("案件管理数", "同時5件〜", "業務処理能力"),
    ("投稿完了率", "90%以上", "サービス品質"),
    ("レポート提出率", "100%", "顧客満足度"),
]

for ri, (kpi, target, note) in enumerate(ops_kpis):
    ry = 1.85 + ri * 0.82
    bg = RGBColor(0x1A,0x26,0x44) if ri % 2 == 0 else RGBColor(0x1E,0x2D,0x52)
    add_rect(slide, 6.9, ry, 6.0, 0.78, fill=bg)
    add_rect(slide, 6.9, ry, 0.06, 0.78, fill=C_ACCENT2)
    add_text(slide, kpi, 7.05, ry+0.08, 2.8, 0.62, size=12, bold=True, color=C_WHITE)
    add_text(slide, target, 9.95, ry+0.08, 1.6, 0.62, size=13, bold=True, color=C_YELLOW, align=PP_ALIGN.CENTER)
    add_text(slide, note, 11.65, ry+0.08, 1.2, 0.62, size=11, color=C_GRAY)

# Monthly review cycle
add_rect(slide, 0.35, 5.35, 12.6, 0.48, fill=RGBColor(0x22,0x34,0x5A))
add_text(slide, "📅 KPIレビュー：毎月月初に前月振り返り → 翌月アクション策定", 0.5, 5.38, 12.4, 0.42, size=13, bold=True, color=C_YELLOW)

page_num(slide, 15)

# =====================================================================
# SLIDE 16: Numerical Targets
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "14｜数値目標", "Numerical Targets")

milestones = [
    ("3ヶ月", C_BLUE_LIGHT, [
        ("登録インフルエンサー", "50名"),
        ("月間案件数", "1〜2件"),
        ("月間売上", "0〜10万円"),
        ("主な活動", "実績作り"),
    ]),
    ("6ヶ月", C_ACCENT2, [
        ("登録インフルエンサー", "100名"),
        ("月間案件数", "5件"),
        ("月間売上", "20〜50万円"),
        ("主な活動", "有料化スタート"),
    ]),
    ("12ヶ月", C_ACCENT, [
        ("登録インフルエンサー", "300名"),
        ("月間案件数", "20件"),
        ("月間売上", "100〜300万円"),
        ("主な活動", "自治体・大型案件"),
    ]),
]

for i, (period, color, metrics) in enumerate(milestones):
    bx = 0.4 + i * 4.3
    # Period badge
    add_rect(slide, bx, 1.35, 4.0, 0.65, fill=color)
    add_text(slide, period, bx+0.1, 1.38, 3.85, 0.58, size=24, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    for mi, (label, value) in enumerate(metrics):
        my = 2.05 + mi * 1.1
        bg = RGBColor(0x1A,0x26,0x44) if mi % 2 == 0 else RGBColor(0x1E,0x2D,0x52)
        add_rect(slide, bx, my, 4.0, 1.05, fill=bg)
        add_rect(slide, bx, my, 0.06, 1.05, fill=color)
        add_text(slide, label, bx+0.15, my+0.08, 3.7, 0.42, size=12, color=C_GRAY)
        add_text(slide, value, bx+0.15, my+0.5, 3.7, 0.5, size=18, bold=True, color=C_YELLOW)

# Growth arrow
add_text(slide, "→  成長軌道", 4.9, 6.5, 3.5, 0.5, size=16, bold=True, color=C_ACCENT, align=PP_ALIGN.CENTER)

page_num(slide, 16)

# =====================================================================
# SLIDE 17: Risks & Countermeasures
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "15｜リスクと対策", "Risk Management")

risks = [
    ("投稿未実施", "HIGH", C_ACCENT, "事前に投稿義務を明記した\n合意書（DM文書化）を取得\n投稿期限を明確に設定"),
    ("来店キャンセル", "MED", C_ACCENT2, "前日リマインドDM送信\nキャンセルポリシーを\n事前合意"),
    ("フォロワー購入\nアカウント", "HIGH", C_ACCENT, "インサイト確認を必須化\nエンゲージメント率で判定\n(目安：1%以上)"),
    ("成果測定困難", "MED", C_BLUE_LIGHT, "投稿URLを保存・記録\nリーチ数・いいね数を\nレポートに明記"),
    ("競合参入", "LOW", RGBColor(0x55,0x77,0x99), "DB規模で差別化\n先行優位を早期確立\nリレーション深化"),
]

# Table header
add_rect(slide, 0.35, 1.35, 3.0, 0.45, fill=RGBColor(0x0A,0x14,0x2E))
add_text(slide, "リスク", 0.5, 1.38, 2.8, 0.38, size=13, bold=True, color=C_GRAY)
add_rect(slide, 3.35, 1.35, 1.2, 0.45, fill=RGBColor(0x0A,0x14,0x2E))
add_text(slide, "レベル", 3.42, 1.38, 1.1, 0.38, size=13, bold=True, color=C_GRAY, align=PP_ALIGN.CENTER)
add_rect(slide, 4.55, 1.35, 8.45, 0.45, fill=RGBColor(0x0A,0x14,0x2E))
add_text(slide, "対　策", 4.7, 1.38, 8.2, 0.38, size=13, bold=True, color=C_GRAY)

for ri, (risk, level, color, action) in enumerate(risks):
    ry = 1.82 + ri * 1.08
    bg = RGBColor(0x1A,0x26,0x44) if ri % 2 == 0 else RGBColor(0x1E,0x2D,0x52)
    add_rect(slide, 0.35, ry, 3.0, 1.02, fill=bg)
    add_rect(slide, 0.35, ry, 0.06, 1.02, fill=color)
    add_multiline(slide, risk.split('\n'), 0.5, ry+0.15, 2.75, 0.75, size=13, bold=True, color=C_WHITE)
    add_rect(slide, 3.35, ry, 1.2, 1.02, fill=color)
    add_text(slide, level, 3.35, ry+0.28, 1.2, 0.45, size=14, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide, 4.55, ry, 8.45, 1.02, fill=bg)
    add_multiline(slide, action.split('\n'), 4.7, ry+0.1, 8.1, 0.85, size=12.5, color=C_WHITE)

page_num(slide, 17)

# =====================================================================
# SLIDE 18: Future Vision
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "16｜将来構想", "Future Vision")

phases_future = [
    ("Phase 1\n〜6ヶ月", "飲食店\nキャスティング", C_BLUE_LIGHT),
    ("Phase 2\n〜12ヶ月", "ホテル\n月額契約", C_ACCENT2),
    ("Phase 3\n〜18ヶ月", "観光施設\n大型案件", C_ACCENT),
    ("Phase 4\n〜24ヶ月", "自治体\nDMO連携", RGBColor(0x8B,0x00,0x8B)),
    ("Phase 5\n〜3年", "日本最大級\nINFLUENCER DB", C_YELLOW),
]

for i, (period, title, color) in enumerate(phases_future):
    bx = 0.3 + i * 2.55
    height = 1.0 + i * 0.7
    by = 5.5 - height
    add_rect(slide, bx, by, 2.3, height, fill=color)
    add_multiline(slide, title.split('\n'), bx+0.1, by+0.15, 2.1, height-0.2, size=13, bold=True, color=C_DARK if color == C_YELLOW else C_WHITE, align=PP_ALIGN.CENTER)
    add_text(slide, period.replace('\n',' '), bx+0.05, 5.55, 2.2, 0.55, size=10, color=C_GRAY, align=PP_ALIGN.CENTER)

# Ground line
add_rect(slide, 0.3, 5.5, 12.7, 0.06, fill=C_GRAY)

# Final vision statement
add_rect(slide, 0.3, 6.0, 12.7, 0.85, fill=RGBColor(0x22,0x34,0x5A))
add_text(slide, "🌏 最終目標：訪日インフルエンサーのデータベースを日本一の規模へ　→　プラットフォーム化・SaaS化", 0.5, 6.1, 12.5, 0.65, size=14, bold=True, color=C_YELLOW, align=PP_ALIGN.CENTER)

page_num(slide, 18)

# =====================================================================
# SLIDE 19: 90-Day Action Plan (Roadmap)
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide)
header_bar(slide, "17｜今後90日間アクションプラン", "90-Day Action Roadmap")

# Month columns
months = ["Month 1（7月）", "Month 2（8月）", "Month 3（9月）"]
m_colors = [C_BLUE_LIGHT, C_ACCENT2, C_ACCENT]
for mi, (m, mc) in enumerate(zip(months, m_colors)):
    bx = 0.35 + mi * 4.3
    add_rect(slide, bx, 1.35, 4.1, 0.5, fill=mc)
    add_text(slide, m, bx+0.1, 1.37, 3.95, 0.44, size=16, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

# Tasks per month
tasks = [
    [
        "【DB構築】",
        "・管理シート設計・整備",
        "・既存フォロワーへのDM開始",
        "・インサイト回収フロー確立",
        "【営業準備】",
        "・営業リスト作成（飲食20件）",
        "・提案資料作成",
        "・DMテンプレート完成",
        "【目標】",
        "登録インフルエンサー20名",
    ],
    [
        "【初回PR実施】",
        "・飲食店PR 3〜5件実行",
        "・投稿レポート作成・蓄積",
        "【営業開始】",
        "・飲食店10社へ営業",
        "・ホテル5社へアプローチ",
        "・初回有料受注を目指す",
        "【目標】",
        "登録50名 / 有料1件",
    ],
    [
        "【スケールアップ】",
        "・月額契約 2〜3社獲得",
        "・ホテル有料案件 1件受注",
        "【品質向上】",
        "・レポートテンプレート完成",
        "・投稿品質チェックリスト",
        "・DB 300名体制設計",
        "【目標】",
        "月間案件5件 / 売上20万円",
    ],
]

for mi, (month_tasks, mc) in enumerate(zip(tasks, m_colors)):
    bx = 0.35 + mi * 4.3
    add_rect(slide, bx, 1.85, 4.1, 4.8, fill=RGBColor(0x1A,0x26,0x44))
    add_multiline(slide, month_tasks, bx+0.18, 1.95, 3.85, 4.6, size=11.5, color=C_WHITE)

# KPI bar at bottom
add_rect(slide, 0.35, 6.75, 12.65, 0.55, fill=C_ACCENT)
add_text(slide, "90日目標：登録インフルエンサー50名｜有料案件5件｜月間売上20万円達成", 0.5, 6.8, 12.5, 0.45, size=14, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

page_num(slide, 19)

# =====================================================================
# SLIDE 20: Closing / Next Steps
# =====================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide_bg(slide, C_DARK)
add_rect(slide, 0, 0, 0.12, 7.5, fill=C_ACCENT)
add_rect(slide, 0, 5.5, 13.33, 2.0, fill=C_ACCENT)

add_text(slide, "NEXT STEPS", 0.5, 1.0, 12, 0.6, size=14, bold=True, color=C_GRAY, italic=True)
add_text(slide, "今すぐ始める3つのアクション", 0.5, 1.6, 12, 0.8, size=32, bold=True, color=C_WHITE)

actions = [
    ("01", "インフルエンサーDBシート作成", "Google スプレッドシートでフォーマット整備"),
    ("02", "既存フォロワーへのDM開始", "旅行予定インフルエンサーへの初回接触"),
    ("03", "飲食店営業リスト20件作成", "食べログ・Googleマップで大阪エリアをリスト化"),
]

for i, (num, title, desc) in enumerate(actions):
    bx = 0.7 + i * 4.1
    add_rect(slide, bx, 2.7, 3.8, 2.3, fill=RGBColor(0x22,0x34,0x5A))
    add_rect(slide, bx, 2.7, 0.06, 2.3, fill=C_ACCENT)
    add_text(slide, num, bx+0.15, 2.75, 0.8, 0.6, size=22, bold=True, color=C_ACCENT)
    add_text(slide, title, bx+0.18, 3.3, 3.5, 0.55, size=14, bold=True, color=C_WHITE)
    add_text(slide, desc, bx+0.18, 3.9, 3.5, 0.9, size=12, color=RGBColor(0xBB,0xCC,0xDD))

add_text(slide, "事業成功に向けて — 本日の意思決定をお願いします", 0.5, 5.7, 12.3, 0.55, size=18, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
add_text(slide, "2026年6月　西村", 0.5, 6.5, 12.3, 0.5, size=14, color=C_WHITE, align=PP_ALIGN.CENTER)

page_num(slide, 20)

prs.save("/home/user/-/influencer_casting_business_plan.pptx")
print("DONE")
