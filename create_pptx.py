from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm
import copy

# Color palette
BLACK = RGBColor(0x0A, 0x0A, 0x0A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
NAVY = RGBColor(0x0D, 0x1B, 0x3E)
NAVY_LIGHT = RGBColor(0x1A, 0x2E, 0x5A)
GOLD = RGBColor(0xC9, 0xA4, 0x2A)
GOLD_LIGHT = RGBColor(0xE8, 0xC4, 0x6A)
DARK_GREEN = RGBColor(0x1B, 0x4D, 0x3E)
GRAY_LIGHT = RGBColor(0xF5, 0xF5, 0xF0)
GRAY_MID = RGBColor(0xCC, 0xCC, 0xCC)
GOLD2 = RGBColor(0xC9, 0xA4, 0x2A)
GOLD_ACCENT = RGBColor(0xD4, 0xAF, 0x37)
DARK_GREEN2 = RGBColor(0x1B, 0x4D, 0x3E)

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)

blank_layout = prs.slide_layouts[6]

def add_rect(slide, left, top, width, height, fill_color=None, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape

def add_textbox(slide, text, left, top, width, height, font_size=14, bold=False,
                color=WHITE, align=PP_ALIGN.LEFT, font_name="Yu Gothic", wrap=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font_name
    return txBox

def set_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

# ─── Slide 1: 表紙 ───────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, NAVY)

# Gold accent bar top
add_rect(slide, 0, 0, Inches(13.33), Inches(0.12), GOLD_ACCENT)
# Gold accent bar bottom
add_rect(slide, 0, Inches(7.38), Inches(13.33), Inches(0.12), GOLD_ACCENT)

# Left dark panel
add_rect(slide, 0, 0, Inches(0.5), Inches(7.5), DARK_GREEN2)

# Center decoration line
add_rect(slide, Inches(1.5), Inches(2.8), Inches(10.33), Inches(0.04), GOLD_ACCENT)
add_rect(slide, Inches(1.5), Inches(5.2), Inches(10.33), Inches(0.04), GOLD_ACCENT)

# Main title
add_textbox(slide, "大阪ローカル DMC / DMO", Inches(1.5), Inches(1.3), Inches(10.3), Inches(1.0),
            font_size=42, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(slide, "形成支援事業", Inches(1.5), Inches(2.2), Inches(10.3), Inches(0.8),
            font_size=42, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)

# Subtitle
add_textbox(slide, "大阪のローカル地域資源を、海外向け高付加価値体験へ",
            Inches(1.5), Inches(3.1), Inches(10.3), Inches(0.6),
            font_size=18, bold=False, color=GRAY_MID, align=PP_ALIGN.CENTER)

# Tags
tags = ["東大阪 × 八尾 × 蒲生四丁目", "インバウンド × ローカル × 高付加価値", "DMC / DMO形成支援"]
for i, tag in enumerate(tags):
    x = Inches(1.5 + i * 3.5)
    box = add_rect(slide, x, Inches(3.9), Inches(3.2), Inches(0.45), DARK_GREEN2)
    box.line.color.rgb = GOLD_ACCENT
    box.line.width = Pt(0.75)
    add_textbox(slide, tag, x, Inches(3.95), Inches(3.2), Inches(0.4),
                font_size=11, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)

# Date / Confidential
add_textbox(slide, "社内企画資料　｜　Confidential　｜　2026年", Inches(1.5), Inches(6.8), Inches(10.3), Inches(0.4),
            font_size=10, color=GRAY_MID, align=PP_ALIGN.CENTER)

# ─── Helper: standard slide layout ───────────────────────────────
def make_standard_slide(title_text, subtitle_text=""):
    slide = prs.slides.add_slide(blank_layout)
    set_bg(slide, WHITE)
    # Top navy bar
    add_rect(slide, 0, 0, Inches(13.33), Inches(1.15), NAVY)
    # Gold accent line
    add_rect(slide, 0, Inches(1.15), Inches(13.33), Inches(0.06), GOLD_ACCENT)
    # Left accent bar
    add_rect(slide, 0, 0, Inches(0.08), Inches(7.5), DARK_GREEN2)
    # Bottom bar
    add_rect(slide, 0, Inches(7.3), Inches(13.33), Inches(0.2), NAVY)

    add_textbox(slide, title_text, Inches(0.35), Inches(0.18), Inches(10), Inches(0.7),
                font_size=26, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if subtitle_text:
        add_textbox(slide, subtitle_text, Inches(0.35), Inches(0.82), Inches(12.5), Inches(0.4),
                    font_size=13, color=GOLD_ACCENT, align=PP_ALIGN.LEFT)
    return slide

# ─── Slide 2: 背景 ───────────────────────────────────────────────
slide = make_standard_slide("本企画の背景", "なぜ今、大阪ローカルなのか")

# 4 cards
card_data = [
    ("📈", "訪日需要の急拡大", "2025年 訪日外客数\n過去最高水準更新\nインバウンド消費額も急増"),
    ("🗾", "有名観光地への集中", "京都・東京・大阪都心に\n過集中。地方・ローカルへの\n分散ニーズが高まる"),
    ("🌏", "ローカル体験需要", "富裕層・教育機関・企業が\n求めるのは「本物体験」\n工場・職人・食・文化"),
    ("🤝", "万博人脈の活用", "2025年大阪万博で形成した\n海外ネットワークを\nビジネスに転換するチャンス"),
]

for i, (icon, ttl, body) in enumerate(card_data):
    x = Inches(0.35 + i * 3.24)
    y = Inches(1.45)
    add_rect(slide, x, y, Inches(3.0), Inches(5.5), NAVY)
    r = slide.shapes[-1]
    r.line.color.rgb = GOLD_ACCENT
    r.line.width = Pt(1)
    add_textbox(slide, icon, x, y + Inches(0.2), Inches(3.0), Inches(0.5),
                font_size=28, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide, x + Inches(0.15), y + Inches(0.85), Inches(2.7), Inches(0.04), GOLD_ACCENT)
    add_textbox(slide, ttl, x, y + Inches(0.95), Inches(3.0), Inches(0.5),
                font_size=14, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, body, x + Inches(0.15), y + Inches(1.55), Inches(2.7), Inches(3.5),
                font_size=12, color=GRAY_MID, align=PP_ALIGN.CENTER)

# ─── Slide 3: 課題認識 ────────────────────────────────────────────
slide = make_standard_slide("課題認識", "地域資源はある。しかし、海外向けに商品化されていない")

problems = [
    "英語対応・予約システムが未整備",
    "海外PR・情報発信力が不足",
    "団体受け入れのオペレーション体制なし",
    "地域事業者単独では海外営業が困難",
    "地域資源のコンセプト設計・ブランディングが未着手",
]
solutions = [
    "英語コンテンツ制作・Web整備支援",
    "SNS・インフルエンサー活用の海外PR",
    "PM・全国通訳案内士による現地運営",
    "旅行会社・海外エージェントとの連携",
    "DMC機能によるコンセプト設計・商品化",
]

# Table header
add_rect(slide, Inches(0.35), Inches(1.45), Inches(5.8), Inches(0.45), NAVY)
add_textbox(slide, "課題（現状）", Inches(0.35), Inches(1.5), Inches(5.8), Inches(0.4),
            font_size=13, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)
add_rect(slide, Inches(7.18), Inches(1.45), Inches(5.8), Inches(0.45), DARK_GREEN2)
add_textbox(slide, "自社の解決策", Inches(7.18), Inches(1.5), Inches(5.8), Inches(0.4),
            font_size=13, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)

# Arrow
add_textbox(slide, "→", Inches(6.1), Inches(3.5), Inches(1.1), Inches(0.5),
            font_size=30, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)

for i, (prob, sol) in enumerate(zip(problems, solutions)):
    y = Inches(1.95 + i * 0.95)
    bg_col = GRAY_LIGHT if i % 2 == 0 else WHITE
    add_rect(slide, Inches(0.35), y, Inches(5.8), Inches(0.85), bg_col)
    r = slide.shapes[-1]; r.line.color.rgb = GRAY_MID; r.line.width = Pt(0.5)
    add_textbox(slide, f"✗  {prob}", Inches(0.5), y + Inches(0.1), Inches(5.5), Inches(0.65),
                font_size=12, color=BLACK, align=PP_ALIGN.LEFT)

    add_rect(slide, Inches(7.18), y, Inches(5.8), Inches(0.85), bg_col)
    r2 = slide.shapes[-1]; r2.line.color.rgb = GRAY_MID; r2.line.width = Pt(0.5)
    add_textbox(slide, f"✓  {sol}", Inches(7.33), y + Inches(0.1), Inches(5.5), Inches(0.65),
                font_size=12, color=DARK_GREEN2, align=PP_ALIGN.LEFT)

# ─── Slide 4: 事業コンセプト ──────────────────────────────────────
slide = make_standard_slide("事業コンセプト", "大阪ローカルの地域資源を、海外向けの高付加価値体験に変える")

# Center big message
add_rect(slide, Inches(1.0), Inches(1.5), Inches(11.33), Inches(1.5), NAVY)
r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1.5)
add_textbox(slide, "大阪のローカル地域資源 × DMC機能 × 海外高付加価値体験",
            Inches(1.0), Inches(1.65), Inches(11.33), Inches(1.2),
            font_size=20, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)

# 3 pillars
pillars = [
    ("地域資源の発掘・商品化", "ものづくり・食・文化・街並みを\n体験プログラムに変える"),
    ("DMO形成支援", "観光協会・自治体・地域事業者を\nつなぎ、地域一体で受け入れ体制を構築"),
    ("海外向けDMC事業", "富裕層・教育機関・企業団体向けに\n高付加価値ツアーを提供・運営"),
]
for i, (ttl, body) in enumerate(pillars):
    x = Inches(0.6 + i * 4.2)
    add_rect(slide, x, Inches(3.3), Inches(3.8), Inches(3.6), DARK_GREEN2)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1)
    add_textbox(slide, f"0{i+1}", x + Inches(0.15), Inches(3.45), Inches(0.6), Inches(0.55),
                font_size=24, bold=True, color=GOLD_ACCENT)
    add_rect(slide, x + Inches(0.15), Inches(4.05), Inches(3.5), Inches(0.04), GOLD_ACCENT)
    add_textbox(slide, ttl, x + Inches(0.15), Inches(4.15), Inches(3.5), Inches(0.6),
                font_size=14, bold=True, color=WHITE)
    add_textbox(slide, body, x + Inches(0.15), Inches(4.85), Inches(3.5), Inches(1.8),
                font_size=11, color=GRAY_MID)

# ─── Slide 5: 事業ポジション ──────────────────────────────────────
slide = make_standard_slide("事業ポジション", "自社は「実行パートナー」として地域と海外をつなぐ")

# Positioning diagram
# Row 1
add_rect(slide, Inches(0.5), Inches(1.45), Inches(12.33), Inches(0.5), NAVY)
add_textbox(slide, "登録DMO（行政・観光協会）", Inches(0.5), Inches(1.5), Inches(4.0), Inches(0.4),
            font_size=12, color=GRAY_MID, align=PP_ALIGN.CENTER)
add_textbox(slide, "旅行会社（第1種・第2種）", Inches(4.6), Inches(1.5), Inches(4.0), Inches(0.4),
            font_size=12, color=GRAY_MID, align=PP_ALIGN.CENTER)
add_textbox(slide, "地域事業者・ガイド", Inches(9.3), Inches(1.5), Inches(3.5), Inches(0.4),
            font_size=12, color=GRAY_MID, align=PP_ALIGN.CENTER)

# Self position box
add_rect(slide, Inches(3.0), Inches(2.2), Inches(7.33), Inches(2.0), DARK_GREEN2)
r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(2)
add_textbox(slide, "自社ポジション", Inches(3.0), Inches(2.3), Inches(7.33), Inches(0.4),
            font_size=12, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)
add_textbox(slide, "DMC / 企画・現地コーディネート会社\n（観光地域づくり推進・実行パートナー）",
            Inches(3.0), Inches(2.75), Inches(7.33), Inches(1.3),
            font_size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# NOT box
add_rect(slide, Inches(0.5), Inches(4.5), Inches(5.5), Inches(1.0), RGBColor(0x3A, 0x10, 0x10))
r = slide.shapes[-1]; r.line.color.rgb = RGBColor(0xCC, 0x33, 0x33); r.line.width = Pt(1)
add_textbox(slide, "✗  登録DMOではない（行政登録は不要）\n✗  旅行業登録なしで旅行商品を販売しない",
            Inches(0.6), Inches(4.55), Inches(5.3), Inches(0.9),
            font_size=11, color=RGBColor(0xFF, 0x99, 0x99))

add_rect(slide, Inches(7.33), Inches(4.5), Inches(5.5), Inches(1.5), DARK_GREEN2)
r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1)
add_textbox(slide, "✓  旅行会社と連携して商品販売\n✓  自社はDMC機能・企画・PM・現地運営\n✓  将来的に旅行業登録も検討",
            Inches(7.43), Inches(4.55), Inches(5.3), Inches(1.4),
            font_size=11, color=WHITE)

# ─── Slide 6: 対象エリア ──────────────────────────────────────────
slide = make_standard_slide("対象エリア", "大阪のローカル3エリアに集中した深掘り")

area_data = [
    ("東大阪", "ものづくり × スポーツ", [
        "中小製造業・町工場が集積",
        "東大阪ラグビーの街",
        "職人体験・工場見学",
        "モノづくりミュージアム",
    ], NAVY),
    ("八尾", "伝統工芸 × 中小製造業", [
        "河内木綿・地場産業",
        "中小製造業の集積地",
        "工場体験・B to Bツアー",
        "伝統と最新技術の融合",
    ], DARK_GREEN2),
    ("蒲生四丁目", "古民家 × ローカルフード", [
        "古民家再生・昭和の街並み",
        "個性的な商店街",
        "ローカルフード・日本酒",
        "クリエイター・アーティスト集積",
    ], RGBColor(0x2C, 0x3E, 0x2D)),
]

for i, (name, sub, items, color) in enumerate(area_data):
    x = Inches(0.4 + i * 4.3)
    add_rect(slide, x, Inches(1.45), Inches(4.0), Inches(5.7), color)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1.5)
    add_textbox(slide, name, x, Inches(1.55), Inches(4.0), Inches(0.65),
                font_size=22, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)
    add_rect(slide, x + Inches(0.2), Inches(2.25), Inches(3.6), Inches(0.04), GOLD_ACCENT)
    add_textbox(slide, sub, x, Inches(2.35), Inches(4.0), Inches(0.45),
                font_size=13, color=WHITE, align=PP_ALIGN.CENTER)
    for j, item in enumerate(items):
        add_textbox(slide, f"• {item}", x + Inches(0.25), Inches(2.95 + j * 0.9), Inches(3.5), Inches(0.75),
                    font_size=12, color=GRAY_LIGHT)

# ─── Slide 7: ターゲット ──────────────────────────────────────────
slide = make_standard_slide("ターゲット顧客", "海外富裕層・教育機関・企業・政府団体")

targets = [
    ("🏆", "海外富裕層", "FIT・グループ旅行", [
        "プレミアム体験への高い支払い意欲",
        "オーセンティックなローカル体験を求める",
        "SNS発信力でPR効果も期待",
    ]),
    ("🎓", "海外大学・教育機関", "教育旅行・研修", [
        "日本のものづくり・地域産業を学ぶ",
        "大学の研修旅行・フィールドワーク",
        "長期滞在・リピート可能性",
    ]),
    ("🏢", "海外企業・政府団体", "ビジネス視察・研修", [
        "製造業・中小企業政策の視察",
        "インセンティブ旅行",
        "大人数・高予算案件",
    ]),
    ("🎯", "MICE / インセンティブ", "コンベンション連携", [
        "万博・国際会議の前後ツアー",
        "企業表彰旅行の組み込み",
        "カスタマイズ性の高いプログラム",
    ]),
]

for i, (icon, name, sub, items) in enumerate(targets):
    col = i % 2
    row = i // 2
    x = Inches(0.35 + col * 6.65)
    y = Inches(1.5 + row * 2.85)
    add_rect(slide, x, y, Inches(6.3), Inches(2.65), NAVY if col == 0 else DARK_GREEN2)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1)
    add_textbox(slide, icon + "  " + name, x + Inches(0.15), y + Inches(0.1), Inches(4.0), Inches(0.5),
                font_size=16, bold=True, color=GOLD_ACCENT)
    add_textbox(slide, sub, x + Inches(0.15), y + Inches(0.6), Inches(5.9), Inches(0.35),
                font_size=11, color=GRAY_MID)
    add_rect(slide, x + Inches(0.15), y + Inches(0.98), Inches(5.9), Inches(0.04), GOLD_ACCENT)
    for j, item in enumerate(items):
        add_textbox(slide, f"• {item}", x + Inches(0.2), y + Inches(1.1 + j * 0.48), Inches(5.9), Inches(0.42),
                    font_size=11, color=WHITE)

# ─── Slide 8: 提供サービス ────────────────────────────────────────
slide = make_standard_slide("提供サービス", "地域資源の発掘から海外営業・現地運営まで一気通貫")

services = [
    ("01", "地域資源発掘", "地域事業者・自治体とともにコンテンツを掘り起こす"),
    ("02", "コンセプト設計", "ターゲットに合わせた体験プログラムを設計"),
    ("03", "ツアー造成", "日程・価格・オペレーションを設計"),
    ("04", "英語資料化", "英語パンフレット・Web・資料作成"),
    ("05", "SNS / PR", "@mr.sasuke_japan等を活用した海外向け発信"),
    ("06", "海外営業支援", "旅行会社・海外エージェントへの提案営業"),
    ("07", "現地運営", "PM・全国通訳案内士による当日コーディネート"),
    ("08", "レポート作成", "実施後の効果測定・自治体・旅行会社向けレポート"),
]

for i, (num, ttl, desc) in enumerate(services):
    col = i % 4
    row = i // 4
    x = Inches(0.35 + col * 3.24)
    y = Inches(1.5 + row * 2.8)
    bg = NAVY if (col + row) % 2 == 0 else DARK_GREEN2
    add_rect(slide, x, y, Inches(3.0), Inches(2.6), bg)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(0.75)
    add_textbox(slide, num, x + Inches(0.12), y + Inches(0.1), Inches(0.7), Inches(0.45),
                font_size=20, bold=True, color=GOLD_ACCENT)
    add_textbox(slide, ttl, x + Inches(0.12), y + Inches(0.6), Inches(2.75), Inches(0.45),
                font_size=14, bold=True, color=WHITE)
    add_rect(slide, x + Inches(0.12), y + Inches(1.08), Inches(2.75), Inches(0.04), GOLD_ACCENT)
    add_textbox(slide, desc, x + Inches(0.12), y + Inches(1.2), Inches(2.75), Inches(1.2),
                font_size=10, color=GRAY_LIGHT)

# ─── Slide 9: ツアー商品案 ────────────────────────────────────────
slide = make_standard_slide("ツアー商品案", "4つのフラッグシップ体験プログラム")

tours = [
    ("Higashi Osaka\nManufacturing Experience",
     "東大阪", "半日〜1日", "USD 200〜500/人",
     "町工場訪問・職人体験・ものづくりの哲学\n富裕層・大学・企業研修向け"),
    ("Yao Craft &\nLocal Industry Tour",
     "八尾", "半日〜1日", "USD 180〜400/人",
     "伝統工芸×最新製造業\n工場見学・中小企業の経営哲学"),
    ("Gamoyon Deep Osaka\nFood & Kominka Walk",
     "蒲生四丁目", "3〜5時間", "USD 150〜350/人",
     "古民家・商店街・ローカルフード\n食の深掘り体験"),
    ("Executive Local\nBusiness Study Tour",
     "東大阪・八尾", "1〜2日", "USD 800〜2,000/人",
     "政府・企業団体向けカスタム\n産業視察＋文化体験のプレミアムパッケージ"),
]

for i, (name, area, duration, price, desc) in enumerate(tours):
    x = Inches(0.35 + i * 3.24)
    add_rect(slide, x, Inches(1.45), Inches(3.0), Inches(5.7), NAVY)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1)
    add_textbox(slide, name, x + Inches(0.12), Inches(1.6), Inches(2.75), Inches(1.1),
                font_size=13, bold=True, color=WHITE)
    add_rect(slide, x + Inches(0.12), Inches(2.75), Inches(2.75), Inches(0.04), GOLD_ACCENT)
    add_textbox(slide, f"📍 {area}", x + Inches(0.12), Inches(2.85), Inches(2.75), Inches(0.35),
                font_size=11, color=GOLD_ACCENT)
    add_textbox(slide, f"⏱ {duration}", x + Inches(0.12), Inches(3.2), Inches(2.75), Inches(0.35),
                font_size=11, color=GRAY_LIGHT)
    add_textbox(slide, price, x + Inches(0.12), Inches(3.6), Inches(2.75), Inches(0.45),
                font_size=14, bold=True, color=GOLD_ACCENT)
    add_textbox(slide, desc, x + Inches(0.12), Inches(4.1), Inches(2.75), Inches(2.8),
                font_size=10, color=GRAY_LIGHT)

# ─── Slide 10: 事業スキーム ───────────────────────────────────────
slide = make_standard_slide("事業スキーム", "旅行会社と連携しながら自社がDMC機能を担う")

# Flow chart
nodes = [
    (Inches(5.2), Inches(1.5), Inches(2.9), Inches(0.7), "海外大学・企業・政府団体", GRAY_LIGHT, BLACK),
    (Inches(5.2), Inches(2.7), Inches(2.9), Inches(0.7), "登録済み旅行会社", GRAY_LIGHT, BLACK),
    (Inches(3.8), Inches(3.9), Inches(5.7), Inches(1.1), "自社（DMC/企画/現地コーディネート/PM/PR）", DARK_GREEN2, WHITE),
    (Inches(5.2), Inches(5.4), Inches(2.9), Inches(0.7), "地域事業者・ガイド", NAVY, WHITE),
]

for lft, top, wid, hgt, txt, bg, fg in nodes:
    add_rect(slide, lft, top, wid, hgt, bg)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1)
    add_textbox(slide, txt, lft, top + Inches(0.12), wid, hgt - Inches(0.15),
                font_size=12, bold=True, color=fg, align=PP_ALIGN.CENTER)

# Arrows
for ay in [Inches(2.22), Inches(3.42), Inches(5.02)]:
    add_textbox(slide, "▼", Inches(6.42), ay, Inches(0.5), Inches(0.35),
                font_size=18, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)

# Side notes
side_labels = [
    (Inches(0.4), Inches(1.6), "問い合わせ・受注"),
    (Inches(0.4), Inches(2.8), "契約・販売"),
    (Inches(0.4), Inches(4.1), "企画・PM・当日運営・PR"),
    (Inches(0.4), Inches(5.5), "体験提供・協力"),
]
for sx, sy, stxt in side_labels:
    add_textbox(slide, stxt, sx, sy, Inches(2.8), Inches(0.5),
                font_size=10, color=DARK_GREEN2, align=PP_ALIGN.LEFT)

# Revenue flows
add_rect(slide, Inches(9.5), Inches(1.5), Inches(3.4), Inches(4.6), NAVY)
r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1)
add_textbox(slide, "収益フロー", Inches(9.5), Inches(1.6), Inches(3.4), Inches(0.45),
            font_size=12, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)
rev_items = ["企画費・コーディネート費", "当日運営費・ガイド費", "SNS/PR制作費", "業務委託費", "送客手数料"]
for j, rev in enumerate(rev_items):
    add_textbox(slide, f"• {rev}", Inches(9.65), Inches(2.15 + j * 0.72), Inches(3.1), Inches(0.6),
                font_size=10, color=WHITE)

# ─── Slide 11: チーム体制 ─────────────────────────────────────────
slide = make_standard_slide("チーム体制", "専門資格とネットワークを持つ実行チーム")

# Org chart
add_rect(slide, Inches(4.8), Inches(1.5), Inches(3.7), Inches(0.75), DARK_GREEN2)
r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1.5)
add_textbox(slide, "事業責任者 / Producer", Inches(4.8), Inches(1.58), Inches(3.7), Inches(0.6),
            font_size=13, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)

add_textbox(slide, "│", Inches(6.55), Inches(2.27), Inches(0.3), Inches(0.3),
            font_size=16, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)

add_rect(slide, Inches(5.4), Inches(2.6), Inches(2.5), Inches(0.65), NAVY)
r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1)
add_textbox(slide, "PM（プロジェクトマネージャー）", Inches(5.4), Inches(2.68), Inches(2.5), Inches(0.5),
            font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Guide
add_textbox(slide, "├──────────────────────┤", Inches(3.2), Inches(3.3), Inches(6.93), Inches(0.35),
            font_size=12, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)

guide_roles = [
    ("全国通訳案内士", "英語・多言語ガイド", Inches(1.0)),
    ("総合旅程管理主任者", "ツアーオペレーション", Inches(4.8)),
    ("SNS担当", "@mr.sasuke_japan\n海外向けPR", Inches(8.6)),
]
for ttl, sub, gx in guide_roles:
    add_rect(slide, gx, Inches(3.6), Inches(3.5), Inches(1.2), NAVY)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(0.75)
    add_textbox(slide, ttl, gx, Inches(3.7), Inches(3.5), Inches(0.45),
                font_size=11, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, sub, gx, Inches(4.15), Inches(3.5), Inches(0.6),
                font_size=10, color=WHITE, align=PP_ALIGN.CENTER)

# External partners
add_rect(slide, Inches(0.35), Inches(5.1), Inches(12.63), Inches(0.5), DARK_GREEN2)
r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(0.75)
add_textbox(slide, "外部パートナー", Inches(0.35), Inches(5.15), Inches(2.0), Inches(0.4),
            font_size=11, bold=True, color=GOLD_ACCENT)
partners = "旅行会社　｜　地域事業者　｜　自治体・観光協会　｜　交通事業者　｜　インフルエンサー"
add_textbox(slide, partners, Inches(2.5), Inches(5.15), Inches(10.3), Inches(0.4),
            font_size=11, color=WHITE)

# ─── Slide 12: 収益モデル ─────────────────────────────────────────
slide = make_standard_slide("収益モデル", "複数の収益源で安定したキャッシュフローを確保")

# Table
headers = ["収益項目", "内容", "想定単価・規模"]
rows = [
    ["企画費", "ツアー・コンテンツ設計", "30〜100万円/件"],
    ["コーディネート費", "現地手配・調整業務", "10〜30万円/件"],
    ["当日運営費", "PM・スタッフ費用", "5〜20万円/日"],
    ["ガイド・通訳費", "全国通訳案内士費用", "3〜8万円/日"],
    ["SNS/PR制作費", "コンテンツ・発信費用", "10〜50万円/件"],
    ["コンテンツ造成費", "自治体向け補助事業", "50〜200万円/件"],
    ["業務委託費", "旅行会社からの委託", "ツアー代金の10〜20%"],
    ["送客手数料", "エージェント連携", "ツアー代金の5〜15%"],
]

col_widths = [Inches(3.2), Inches(4.8), Inches(4.8)]
col_starts = [Inches(0.35), Inches(3.55), Inches(8.35)]

# Header
for ci, (hdr, cw, cx) in enumerate(zip(headers, col_widths, col_starts)):
    add_rect(slide, cx, Inches(1.5), cw, Inches(0.5), NAVY)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(0.75)
    add_textbox(slide, hdr, cx, Inches(1.56), cw, Inches(0.4),
                font_size=12, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)

for ri, row in enumerate(rows):
    bg = GRAY_LIGHT if ri % 2 == 0 else WHITE
    y = Inches(2.05 + ri * 0.62)
    for ci, (cell, cw, cx) in enumerate(zip(row, col_widths, col_starts)):
        add_rect(slide, cx, y, cw, Inches(0.6), bg)
        r = slide.shapes[-1]; r.line.color.rgb = GRAY_MID; r.line.width = Pt(0.5)
        fc = DARK_GREEN2 if ci == 2 else BLACK
        add_textbox(slide, cell, cx + Inches(0.1), y + Inches(0.08), cw - Inches(0.1), Inches(0.5),
                    font_size=11, color=fc, bold=(ci == 2))

# ─── Slide 13: 法務・資格・許認可 ────────────────────────────────
slide = make_standard_slide("法務・資格・許認可", "適切なスキームで法的リスクを回避する")

legal_items = [
    ("旅行商品の販売・契約・決済", "初期フェーズ", "登録済み旅行会社が担当\n自社は企画・コーディネートに専念", "NAVY"),
    ("旅行業登録", "自社販売時", "自社で旅行商品を販売する場合\n旅行業登録（第2種・第3種）を検討", "DARK_GREEN"),
    ("旅行サービス手配業登録", "手配業務時", "宿泊・交通・有償ガイドを手配する場合\n登録の必要性を行政書士に確認", "DARK_GREEN"),
    ("スキーム確認", "実行前", "行政書士にスキーム全体を確認\n弁護士に契約書を確認", "NAVY"),
]

for i, (ttl, phase, body, col_name) in enumerate(legal_items):
    row = i // 2
    c = i % 2
    x = Inches(0.35 + c * 6.65)
    y = Inches(1.5 + row * 2.8)
    bg = NAVY if col_name == "NAVY" else DARK_GREEN2
    add_rect(slide, x, y, Inches(6.3), Inches(2.55), bg)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1)
    add_rect(slide, x + Inches(0.15), y + Inches(0.12), Inches(1.2), Inches(0.38), GOLD_ACCENT)
    add_textbox(slide, phase, x + Inches(0.15), y + Inches(0.15), Inches(1.2), Inches(0.32),
                font_size=10, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_textbox(slide, ttl, x + Inches(1.5), y + Inches(0.12), Inches(4.6), Inches(0.42),
                font_size=14, bold=True, color=WHITE)
    add_rect(slide, x + Inches(0.15), y + Inches(0.65), Inches(5.9), Inches(0.04), GOLD_ACCENT)
    add_textbox(slide, body, x + Inches(0.15), y + Inches(0.78), Inches(5.9), Inches(1.6),
                font_size=12, color=GRAY_LIGHT)

# ─── Slide 14: 参考モデル ─────────────────────────────────────────
slide = make_standard_slide("参考モデル", "先行事例を参考にしながら大阪ローカル密着型で差別化")

models = [
    ("HIS 訪日コンサルティング", "旅行大手のインバウンド支援モデル\n自社は大阪ローカルに特化し差別化"),
    ("地方創生インバウンド戦略", "自治体・DMO主導の地域活性化\n自社は実行パートナーとして参画"),
    ("DMO形成支援事例", "観光庁・地域DMO形成のノウハウ\n認定DMOとの連携で信頼性を確保"),
    ("アクティビティ開発", "体験型観光の商品化事例\n地域事業者との共同開発モデル"),
    ("MICE・富裕層手配", "高付加価値層向けの専門エージェント\n価格設計・コンテンツ品質を参考"),
]

for i, (ttl, desc) in enumerate(models):
    x = Inches(0.35 + (i % 3) * 4.3) if i < 3 else Inches(2.1 + (i - 3) * 4.3)
    y = Inches(1.5) if i < 3 else Inches(4.2)
    add_rect(slide, x, y, Inches(4.0), Inches(2.4), NAVY)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(0.75)
    add_textbox(slide, ttl, x + Inches(0.12), y + Inches(0.12), Inches(3.75), Inches(0.55),
                font_size=13, bold=True, color=GOLD_ACCENT)
    add_rect(slide, x + Inches(0.12), y + Inches(0.73), Inches(3.75), Inches(0.04), GOLD_ACCENT)
    add_textbox(slide, desc, x + Inches(0.12), y + Inches(0.85), Inches(3.75), Inches(1.4),
                font_size=11, color=GRAY_LIGHT)

# Differentiation box
add_rect(slide, Inches(0.35), Inches(6.5), Inches(12.63), Inches(0.7), DARK_GREEN2)
r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1)
add_textbox(slide, "自社差別化：大阪ローカル密着 × 万博人脈 × インバウンドSNS発信力 × PM体制",
            Inches(0.35), Inches(6.57), Inches(12.63), Inches(0.5),
            font_size=13, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)

# ─── Slide 15: 自社の強み ─────────────────────────────────────────
slide = make_standard_slide("自社の強み", "5つの競合優位性")

strengths = [
    ("🌐", "万博人脈", "2025年大阪万博で形成した\n海外大学・企業・政府とのネットワーク\nダイレクトな案件獲得に直結"),
    ("📱", "インバウンドSNS\n発信力", "@mr.sasuke_japan\n海外向けリール・コンテンツ\n本物のローカル情報発信"),
    ("🗺️", "大阪ローカル理解", "東大阪・八尾・蒲生四丁目の\n深いローカル知識・人脈\n地域事業者との信頼関係"),
    ("📋", "専門資格体制", "全国通訳案内士\n総合旅程管理主任者\nPMによる安定したオペレーション"),
    ("🤝", "柔軟な連携力", "旅行会社・自治体・観光協会と\n連携できる実行パートナー機能\n登録不要で動ける機動力"),
]

for i, (icon, ttl, desc) in enumerate(strengths):
    x = Inches(0.35 + i * 2.59)
    add_rect(slide, x, Inches(1.5), Inches(2.45), Inches(5.7), NAVY)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1)
    add_textbox(slide, icon, x, Inches(1.6), Inches(2.45), Inches(0.55),
                font_size=28, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide, x + Inches(0.15), Inches(2.22), Inches(2.15), Inches(0.04), GOLD_ACCENT)
    add_textbox(slide, ttl, x, Inches(2.32), Inches(2.45), Inches(0.75),
                font_size=13, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, desc, x + Inches(0.12), Inches(3.15), Inches(2.2), Inches(3.8),
                font_size=11, color=GRAY_LIGHT)

# ─── Slide 16: 初期3ヶ月アクションプラン ────────────────────────
slide = make_standard_slide("初期3ヶ月アクションプラン", "小さく始め、実証しながら拡大する")

months = [
    ("Month 1", "基盤構築", [
        "社内企画資料作成・共有",
        "行政書士へスキーム相談",
        "旅行会社パートナー相談開始",
        "万博人脈の整理・ヒアリング",
        "対象エリアの現地調査",
    ], NAVY),
    ("Month 2", "商品設計", [
        "モデルツアーの詳細設計",
        "地域事業者の開拓・交渉",
        "海外人脈へのニーズヒアリング",
        "英語版コンテンツ資料作成",
        "価格設計・収益モデル確定",
    ], DARK_GREEN2),
    ("Month 3", "実証・PR", [
        "テストツアーの実施",
        "SNS・インスタ発信開始",
        "自治体・観光協会への提案",
        "旅行会社への商品提案",
        "実証結果レポート作成",
    ], RGBColor(0x2C, 0x3A, 0x5E)),
]

for i, (month, phase, tasks, bg) in enumerate(months):
    x = Inches(0.4 + i * 4.3)
    add_rect(slide, x, Inches(1.5), Inches(4.0), Inches(5.7), bg)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1.5)
    add_textbox(slide, month, x, Inches(1.58), Inches(4.0), Inches(0.5),
                font_size=18, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, phase, x, Inches(2.08), Inches(4.0), Inches(0.45),
                font_size=14, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(slide, x + Inches(0.2), Inches(2.58), Inches(3.6), Inches(0.04), GOLD_ACCENT)
    for j, task in enumerate(tasks):
        add_textbox(slide, f"✓  {task}", x + Inches(0.2), Inches(2.7 + j * 0.87), Inches(3.6), Inches(0.8),
                    font_size=11, color=GRAY_LIGHT)

# ─── Slide 17: 今後の検討事項 ────────────────────────────────────
slide = make_standard_slide("今後の検討事項", "事業推進に向けて並行して検討が必要な事項")

todos = [
    ("旅行会社パートナー選定", "第1種・第2種旅行会社との\n業務委託・連携先の確定"),
    ("許認可確認", "旅行業・旅行サービス手配業\n登録要否の確定"),
    ("価格設計", "ターゲット別の価格帯・\nパッケージ構成の策定"),
    ("地域事業者向け資料", "協力依頼・説明資料の\n日英版作成"),
    ("海外営業リスト", "大学・企業・政府団体の\nターゲットリスト整備"),
    ("収益分配設計", "旅行会社・地域事業者との\n収益分配ルールの策定"),
    ("保険・責任範囲", "ツアー中の事故・キャンセル\nリスクの整理と保険設計"),
]

for i, (ttl, desc) in enumerate(todos):
    col = i % 4 if i < 4 else (i - 4)
    row = 0 if i < 4 else 1
    x = Inches(0.35 + col * 3.24)
    y = Inches(1.55 + row * 2.9)
    w = Inches(3.0) if i < 6 else Inches(6.3)
    add_rect(slide, x, y, w, Inches(2.6), NAVY)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(0.75)
    add_textbox(slide, f"{i+1:02d}", x + Inches(0.12), y + Inches(0.1), Inches(0.5), Inches(0.45),
                font_size=18, bold=True, color=GOLD_ACCENT)
    add_textbox(slide, ttl, x + Inches(0.12), y + Inches(0.6), w - Inches(0.25), Inches(0.5),
                font_size=13, bold=True, color=WHITE)
    add_rect(slide, x + Inches(0.12), y + Inches(1.15), w - Inches(0.25), Inches(0.04), GOLD_ACCENT)
    add_textbox(slide, desc, x + Inches(0.12), y + Inches(1.25), w - Inches(0.25), Inches(1.2),
                font_size=11, color=GRAY_LIGHT)

# ─── Slide 18: まとめ / Next Action ──────────────────────────────
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, NAVY)
add_rect(slide, 0, 0, Inches(13.33), Inches(0.12), GOLD_ACCENT)
add_rect(slide, 0, Inches(7.38), Inches(13.33), Inches(0.12), GOLD_ACCENT)
add_rect(slide, 0, 0, Inches(0.08), Inches(7.5), DARK_GREEN2)

add_textbox(slide, "まとめ / Next Action", Inches(0.35), Inches(0.2), Inches(10), Inches(0.7),
            font_size=28, bold=True, color=WHITE)
add_rect(slide, Inches(0.35), Inches(0.95), Inches(12.63), Inches(0.06), GOLD_ACCENT)

next_actions = [
    ("STEP 1", "旅行会社との連携", "まず旅行会社パートナーを確保し、\nスキームを確定。小さく実証を開始する。"),
    ("STEP 2", "万博人脈の活用", "海外大学・企業・政府団体に\nダイレクトにアプローチ。ニーズを確認する。"),
    ("STEP 3", "実績づくり", "大阪ローカルDMCとして\nテストツアーを実施し、実績と信頼を積む。"),
]

for i, (step, ttl, body) in enumerate(next_actions):
    x = Inches(0.5 + i * 4.3)
    add_rect(slide, x, Inches(1.2), Inches(4.0), Inches(4.0), DARK_GREEN2)
    r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1.5)
    add_textbox(slide, step, x, Inches(1.3), Inches(4.0), Inches(0.55),
                font_size=16, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)
    add_rect(slide, x + Inches(0.2), Inches(1.9), Inches(3.6), Inches(0.04), GOLD_ACCENT)
    add_textbox(slide, ttl, x, Inches(2.02), Inches(4.0), Inches(0.55),
                font_size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(slide, body, x + Inches(0.2), Inches(2.65), Inches(3.6), Inches(2.4),
                font_size=12, color=GRAY_LIGHT, align=PP_ALIGN.CENTER)

# Summary message
add_rect(slide, Inches(0.5), Inches(5.4), Inches(12.33), Inches(1.5), RGBColor(0x0D, 0x1B, 0x3E))
r = slide.shapes[-1]; r.line.color.rgb = GOLD_ACCENT; r.line.width = Pt(1.5)
add_textbox(slide, "大阪のローカルに眠る資源を、世界に届ける体験へ。",
            Inches(0.5), Inches(5.5), Inches(12.33), Inches(0.6),
            font_size=18, bold=True, color=GOLD_ACCENT, align=PP_ALIGN.CENTER)
add_textbox(slide, "DMC × DMO形成支援 × 海外高付加価値体験  ｜  まず動く、小さく実証、大きく育てる",
            Inches(0.5), Inches(6.1), Inches(12.33), Inches(0.55),
            font_size=13, color=WHITE, align=PP_ALIGN.CENTER)

output_path = "/home/user/-/大阪ローカルDMC_DMO形成支援事業_企画書.pptx"
prs.save(output_path)
print(f"Saved: {output_path}")
print(f"Slides: {len(prs.slides)}")
