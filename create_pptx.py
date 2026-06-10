from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pptx.oxml.ns as pns
from lxml import etree
import copy

# Color palette
NAVY       = RGBColor(0x0A, 0x1A, 0x3A)   # deep navy bg
NAVY_MID   = RGBColor(0x0D, 0x26, 0x55)   # mid navy
ACCENT_P   = RGBColor(0xC0, 0x5A, 0xC8)   # purple accent
ACCENT_B   = RGBColor(0x4A, 0x90, 0xD9)   # blue accent
ACCENT_PK  = RGBColor(0xE8, 0x5A, 0x9A)   # pink accent
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xCC, 0xD6, 0xE8)
GOLD       = RGBColor(0xF0, 0xC0, 0x40)

W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

BLANK = prs.slide_layouts[6]  # completely blank

# ── helpers ──────────────────────────────────────────────────────────────────

def add_rect(slide, l, t, w, h, fill=None, line=None, alpha=None):
    shape = slide.shapes.add_shape(1, l, t, w, h)
    shape.line.fill.background()
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, l, t, w, h,
             size=18, bold=False, color=WHITE,
             align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txb = slide.shapes.add_textbox(l, t, w, h)
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.italic = italic
    return txb

def add_gradient_bg(slide, c1=NAVY, c2=NAVY_MID):
    """Dark navy gradient background via solid fill (two layered rects)."""
    add_rect(slide, 0, 0, W, H, fill=c1)
    # subtle lighter stripe at bottom-right
    add_rect(slide, W*0.5, H*0.5, W*0.5, H*0.5, fill=c2)

def accent_bar(slide, color=ACCENT_P, height=Pt(4)):
    add_rect(slide, 0, H - height, W, height, fill=color)

def slide_header(slide, title, subtitle=None,
                 title_size=32, sub_size=16,
                 title_color=WHITE, sub_color=LIGHT_GRAY):
    add_text(slide, title,
             Inches(0.7), Inches(0.35), Inches(11.9), Inches(0.9),
             size=title_size, bold=True, color=title_color)
    if subtitle:
        add_text(slide, subtitle,
                 Inches(0.7), Inches(1.1), Inches(11.9), Inches(0.55),
                 size=sub_size, color=sub_color)

def divider(slide, y, color=ACCENT_P, width=Inches(12.0), x=Inches(0.65)):
    add_rect(slide, x, y, width, Pt(2), fill=color)

def big_kpi(slide, number, label, l, t, w=Inches(2.8), h=Inches(1.5),
            num_color=ACCENT_B, num_size=40):
    add_text(slide, number, l, t, w, Inches(0.8),
             size=num_size, bold=True, color=num_color, align=PP_ALIGN.CENTER)
    add_text(slide, label, l, t + Inches(0.75), w, Inches(0.7),
             size=13, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

def card(slide, l, t, w, h, fill=NAVY_MID, border=ACCENT_P):
    add_rect(slide, l, t, w, h, fill=fill, line=border)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 1  Cover
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl, NAVY, RGBColor(0x05, 0x10, 0x28))

# decorative circle blobs
for cx, cy, r, col in [
    (Inches(10.5), Inches(1.0), Inches(3.2), ACCENT_P),
    (Inches(11.8), Inches(5.5), Inches(2.0), ACCENT_B),
    (Inches(0.3),  Inches(6.5), Inches(1.8), ACCENT_PK),
]:
    blob = sl.shapes.add_shape(9, cx - r, cy - r, r*2, r*2)  # ellipse
    blob.fill.solid(); blob.fill.fore_color.rgb = col
    blob.line.fill.background()
    sp = blob.element
    sp.attrib['style'] = ''
    # reduce opacity via transparency hack via xml
    # just keep as-is for color pop

# Tag line top
add_text(sl, "SPONSORSHIP PROPOSAL",
         Inches(0.7), Inches(0.5), Inches(9), Inches(0.6),
         size=13, color=ACCENT_P, bold=True)

# Main title
add_text(sl, "Kansai Global",
         Inches(0.7), Inches(1.2), Inches(10), Inches(1.4),
         size=64, bold=True, color=WHITE)
add_text(sl, "Creators Meetup",
         Inches(0.7), Inches(2.4), Inches(10), Inches(1.3),
         size=64, bold=True, color=ACCENT_P)

# Subtitle JP
add_text(sl,
         "関西から世界へ。訪日観光・ローカル体験を発信する\nグローバルクリエイターコミュニティ",
         Inches(0.7), Inches(3.75), Inches(9.5), Inches(1.0),
         size=17, color=LIGHT_GRAY)

# Presented by
add_text(sl, "Presented by  Mr.sasuke inc.",
         Inches(0.7), Inches(6.5), Inches(7), Inches(0.6),
         size=14, color=ACCENT_B)

# KPI mini badges
for i, (n, l) in enumerate([("35", "Creators"), ("4.7M+", "Followers"), ("30-50", "Per Event")]):
    bx = Inches(0.7 + i * 2.1)
    card(sl, bx, Inches(5.0), Inches(1.9), Inches(1.2), fill=RGBColor(0x10,0x28,0x55), border=ACCENT_B)
    add_text(sl, n, bx, Inches(5.05), Inches(1.9), Inches(0.65),
             size=22, bold=True, color=ACCENT_B, align=PP_ALIGN.CENTER)
    add_text(sl, l, bx, Inches(5.65), Inches(1.9), Inches(0.55),
             size=11, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

accent_bar(sl, ACCENT_P)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 2  Concept
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, '関西の「リアルな魅力」を世界へ届けるクリエイターコミュニティ', title_size=28)
divider(sl, Inches(1.55))

points = [
    ("🌏", "Global Creators Hub",
     "関西在住・関西に強いグローバルクリエイターが集まるオフライン交流会"),
    ("📱", "Inbound SNS Amplification",
     "訪日外国人に向けて グルメ・観光・文化・ローカル体験を海外向けに発信"),
    ("🤝", "Community Connection",
     "SNS発信者同士がつながり 関西の魅力をより広く届けるコミュニティ拠点"),
]
for i, (icon, ttl, desc) in enumerate(points):
    bx = Inches(0.7); by = Inches(1.9 + i * 1.7)
    card(sl, bx, by, Inches(11.9), Inches(1.5), fill=RGBColor(0x0D,0x22,0x4A), border=ACCENT_B)
    add_text(sl, icon, bx + Inches(0.1), by + Inches(0.1), Inches(0.7), Inches(0.7), size=28)
    add_text(sl, ttl, bx + Inches(0.9), by + Inches(0.1), Inches(4.5), Inches(0.6),
             size=17, bold=True, color=ACCENT_B)
    add_text(sl, desc, bx + Inches(0.9), by + Inches(0.65), Inches(10.5), Inches(0.7),
             size=14, color=LIGHT_GRAY)

# right accent
add_rect(sl, Inches(12.9), Inches(1.7), Inches(0.25), Inches(4.8), fill=ACCENT_P)
accent_bar(sl, ACCENT_P)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 3  Why Now
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "インバウンドPRは「広告」から\n「体験され、語られる場」へ", title_size=26)
divider(sl, Inches(1.7))

whys = [
    "訪日客はSNSで行き先・食事・体験を決める時代",
    "企業や地域にとって、海外目線で発信できるクリエイターとの接点が重要",
    "関西には魅力的なローカル資源が多いが、海外に届き切っていない",
    "クリエイターが実際に体験し、自然にシェアしたくなる場をつくることが鍵",
]
for i, w in enumerate(whys):
    by = Inches(2.0 + i * 1.2)
    add_rect(sl, Inches(0.7), by + Inches(0.18), Inches(0.08), Inches(0.55), fill=ACCENT_PK)
    add_text(sl, w, Inches(1.05), by, Inches(11.5), Inches(0.9),
             size=17, color=WHITE)

# highlight box
hb_y = Inches(6.8) - Inches(1.2)
card(sl, Inches(0.7), Inches(5.8), Inches(11.9), Inches(1.05),
     fill=RGBColor(0x2A,0x08,0x3A), border=ACCENT_P)
add_text(sl,
         "「広告を見せる」のではなく、クリエイターが体験・撮影・投稿したくなる場をデザインする",
         Inches(0.9), Inches(5.88), Inches(11.5), Inches(0.85),
         size=16, bold=True, color=ACCENT_P, align=PP_ALIGN.CENTER)
accent_bar(sl, ACCENT_P)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 4  Event Overview
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "Kansai Global Creators Meetup とは", title_size=30)
divider(sl, Inches(1.5))

overview = [
    ("📍 開催地",      "関西エリア（大阪・京都・神戸 等）"),
    ("👥 参加者",      "関西在住／関西発信のグローバルクリエイター・インバウンド系インフルエンサー"),
    ("🎯 内容",       "ネットワーキング・情報交換・スポンサー体験・商品紹介・撮影・交流"),
    ("📊 想定規模",    "30〜50名程度（開催日程・会場規模・招待条件により変動）"),
    ("🔄 開催頻度",    "定期開催を想定（年複数回）"),
    ("🎥 参加ジャンル", "旅行・グルメ・日本語/日本文化・ライフスタイル・ホテル・ローカル体験"),
]
for i, (label, val) in enumerate(overview):
    bx = Inches(0.65 + (i % 2) * 6.4)
    by = Inches(1.8 + (i // 2) * 1.7)
    card(sl, bx, by, Inches(6.0), Inches(1.45), fill=RGBColor(0x0D,0x22,0x4A), border=ACCENT_B)
    add_text(sl, label, bx + Inches(0.15), by + Inches(0.1), Inches(5.7), Inches(0.5),
             size=13, bold=True, color=ACCENT_B)
    add_text(sl, val, bx + Inches(0.15), by + Inches(0.55), Inches(5.7), Inches(0.8),
             size=13, color=WHITE)

accent_bar(sl, ACCENT_B)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 5  Creator Network
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "Creator Network", sub_size=14,
             subtitle="関西インバウンドに強いグローバルクリエイターネットワーク")
divider(sl, Inches(1.5))

kpis = [
    ("35",     "Creators\nin Network"),
    ("4.7M+",  "Instagram\nFollowers"),
    ("30–50",  "Expected\nAttendees"),
    ("10+",    "Countries\nRepresented"),
]
for i, (n, l) in enumerate(kpis):
    bx = Inches(0.65 + i * 3.1)
    card(sl, bx, Inches(1.8), Inches(2.8), Inches(2.2),
         fill=RGBColor(0x12,0x2A,0x5A), border=ACCENT_P)
    add_text(sl, n, bx, Inches(2.0), Inches(2.8), Inches(0.95),
             size=44, bold=True, color=ACCENT_P, align=PP_ALIGN.CENTER)
    add_text(sl, l, bx, Inches(2.85), Inches(2.8), Inches(0.75),
             size=13, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

# Creator examples
add_text(sl, "ネットワーク内 招待候補クリエイター例",
         Inches(0.65), Inches(4.25), Inches(11.9), Inches(0.5),
         size=14, bold=True, color=ACCENT_B)

creators = [
    ("100万超", "Japan Travel / Halal"),
    ("90万規模", "Japan Lifestyle"),
    ("39万規模", "日本語・カルチャー"),
    ("28万規模", "Japan Travel"),
    ("21万規模", "Kansai Life"),
]
for i, (f, g) in enumerate(creators):
    bx = Inches(0.65 + i * 2.5)
    card(sl, bx, Inches(4.8), Inches(2.3), Inches(1.5),
         fill=RGBColor(0x0D,0x1E,0x42), border=ACCENT_PK)
    add_text(sl, f, bx, Inches(4.88), Inches(2.3), Inches(0.65),
             size=18, bold=True, color=ACCENT_PK, align=PP_ALIGN.CENTER)
    add_text(sl, g, bx, Inches(5.5), Inches(2.3), Inches(0.55),
             size=11, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

add_text(sl,
         "※参加者数・投稿数・再生数は保証値ではありません。開催条件・参加可否・企画内容により変動します。投稿保証型PRは別途オプションとして設計可能です。",
         Inches(0.65), Inches(6.7), Inches(12.5), Inches(0.5),
         size=9, color=RGBColor(0x88,0x99,0xBB), italic=True)
accent_bar(sl, ACCENT_P)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 6  Past Results
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "過去開催実績", title_size=32)
divider(sl, Inches(1.4))

# Two event cards
for idx, (ev, att, fol, avg) in enumerate([
    ("第1回", "30名", "約871万", "平均 約29万"),
    ("第2回", "30名", "約181万", "平均 約6万"),
]):
    bx = Inches(0.65 + idx * 6.2)
    card(sl, bx, Inches(1.7), Inches(5.85), Inches(3.5),
         fill=RGBColor(0x0D,0x22,0x4A), border=ACCENT_B)
    add_text(sl, ev, bx + Inches(0.2), Inches(1.85), Inches(5.4), Inches(0.65),
             size=22, bold=True, color=ACCENT_B)
    for j, (lbl, val, col) in enumerate([
        ("参加者数",        att,  WHITE),
        ("SNS合計フォロワー", fol,  ACCENT_PK),
        (avg,               "",   LIGHT_GRAY),
    ]):
        add_text(sl, lbl, bx + Inches(0.2), Inches(2.55 + j * 0.85), Inches(3.5), Inches(0.5),
                 size=12, color=LIGHT_GRAY)
        add_text(sl, val, bx + Inches(0.2), Inches(2.9 + j * 0.85), Inches(5.4), Inches(0.65),
                 size=26, bold=True, color=col)
    if idx == 0:
        # label for avg inside the third row
        add_text(sl, avg, bx + Inches(0.2), Inches(4.25 + 0.85*0.5), Inches(5.4), Inches(0.55),
                 size=14, color=LIGHT_GRAY)

# Bottom row: network total
card(sl, Inches(0.65), Inches(5.45), Inches(12.0), Inches(1.3),
     fill=RGBColor(0x1A,0x08,0x3A), border=ACCENT_P)
add_text(sl, "関西クリエイターネットワーク　35名　Instagram合計 約470万フォロワー　（トップ層：100万超クリエイター在籍）",
         Inches(0.85), Inches(5.6), Inches(11.6), Inches(1.0),
         size=17, bold=True, color=ACCENT_P, align=PP_ALIGN.CENTER)

add_text(sl, "※数字は参加者・ネットワークのSNSフォロワー集計に基づく参考値です。",
         Inches(0.65), Inches(7.0), Inches(12.5), Inches(0.35),
         size=9, color=RGBColor(0x77,0x88,0xAA), italic=True)
accent_bar(sl, ACCENT_B)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 7  Audience Profile
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "海外に届く発信力を持つクリエイターが集まる", title_size=27,
             subtitle="Audience Profile — グローバルオーディエンス属性")
divider(sl, Inches(1.6))

# Countries
add_text(sl, "🌍 主なフォロワー国",
         Inches(0.65), Inches(1.8), Inches(6), Inches(0.5),
         size=15, bold=True, color=ACCENT_B)
countries = ["🇺🇸 USA","🇯🇵 Japan","🇮🇳 India","🇬🇧 UK","🇫🇷 France",
             "🇸🇬 Singapore","🇩🇪 Germany","🇮🇩 Indonesia","🇦🇺 Australia",
             "🇵🇭 Philippines","🇨🇦 Canada","🇻🇳 Vietnam","🇹🇭 Thailand"]
for i, c in enumerate(countries):
    bx = Inches(0.65 + (i % 3) * 1.9)
    by = Inches(2.3 + (i // 3) * 0.7)
    add_text(sl, c, bx, by, Inches(1.85), Inches(0.6), size=13, color=WHITE)

# Age bars
add_text(sl, "📊 主なフォロワー年齢層",
         Inches(7.2), Inches(1.8), Inches(5.8), Inches(0.5),
         size=15, bold=True, color=ACCENT_B)
age_data = [
    ("25〜34歳", 0.42, ACCENT_P),
    ("35〜44歳", 0.30, ACCENT_B),
    ("18〜24歳", 0.20, ACCENT_PK),
]
for i, (label, ratio, col) in enumerate(age_data):
    by = Inches(2.4 + i * 0.9)
    add_text(sl, label, Inches(7.2), by, Inches(1.6), Inches(0.55), size=13, color=LIGHT_GRAY)
    bar_w = Inches(4.0 * ratio)
    add_rect(sl, Inches(9.0), by + Inches(0.12), Inches(4.0), Inches(0.4),
             fill=RGBColor(0x1A,0x2E,0x5A))
    add_rect(sl, Inches(9.0), by + Inches(0.12), bar_w, Inches(0.4), fill=col)
    add_text(sl, f"{int(ratio*100)}%", Inches(9.0) + bar_w + Inches(0.1), by,
             Inches(0.7), Inches(0.55), size=13, bold=True, color=col)

# Genres
add_text(sl, "🎬 発信ジャンル",
         Inches(0.65), Inches(5.3), Inches(12), Inches(0.5),
         size=15, bold=True, color=ACCENT_B)
genres = ["Japan Travel", "Osaka / Kansai Food", "Japanese Culture", "Local Experience",
          "Language Learning", "Lifestyle", "Hotel / Tourism / Nightlife"]
for i, g in enumerate(genres):
    bx = Inches(0.65 + i * 1.82)
    card(sl, bx, Inches(5.85), Inches(1.72), Inches(0.7),
         fill=RGBColor(0x0D,0x22,0x4A), border=ACCENT_PK)
    add_text(sl, g, bx + Inches(0.05), Inches(5.92), Inches(1.65), Inches(0.6),
             size=11, color=WHITE, align=PP_ALIGN.CENTER)

accent_bar(sl, ACCENT_PK)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 8  Survey Results
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "参加者満足度の高いコミュニティ",
             subtitle="第2回アンケート結果（回答数：16件）", title_size=29)
divider(sl, Inches(1.55))

scores = [
    ("イベント全体満足度",   "4.88", "/5"),
    ("内容の有益性",        "4.75", "/5"),
    ("今後も参加したい",    "5.0",  "/5"),
]
for i, (label, num, denom) in enumerate(scores):
    bx = Inches(0.65 + i * 4.1)
    card(sl, bx, Inches(1.8), Inches(3.85), Inches(2.0),
         fill=RGBColor(0x10,0x28,0x5A), border=ACCENT_P)
    add_text(sl, num,    bx, Inches(2.0), Inches(3.85), Inches(1.1),
             size=54, bold=True, color=ACCENT_P, align=PP_ALIGN.CENTER)
    add_text(sl, denom,  bx, Inches(2.95), Inches(3.85), Inches(0.6),
             size=22, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)
    add_text(sl, label,  bx, Inches(3.5),  Inches(3.85), Inches(0.6),
             size=12, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

# Comments
add_text(sl, "参加者の声",
         Inches(0.65), Inches(4.1), Inches(12), Inches(0.5),
         size=14, bold=True, color=ACCENT_B)
comments = [
    '"Great atmosphere and great people."',
    '"Everyone welcoming and supporting each other."',
    '"関西人が多い事が要因なのか、初対面でも壁を感じなかった。"',
    '"みんなで関西を盛り上げていきたいという想いが素敵でした。"',
]
for i, c in enumerate(comments):
    bx = Inches(0.65 + (i % 2) * 6.2)
    by = Inches(4.7 + (i // 2) * 0.95)
    add_rect(sl, bx, by, Inches(0.06), Inches(0.65), fill=ACCENT_PK)
    add_text(sl, c, bx + Inches(0.2), by, Inches(5.8), Inches(0.75),
             size=13, color=WHITE, italic=True)

accent_bar(sl, ACCENT_P)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 9  Why Sponsor
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "スポンサー協力により、より価値ある体験設計へ", title_size=27)
divider(sl, Inches(1.5))

# Finance context box
card(sl, Inches(0.65), Inches(1.75), Inches(11.9), Inches(1.1),
     fill=RGBColor(0x0A,0x1A,0x38), border=ACCENT_B)
add_text(sl,
         "現在は参加費中心で運営しているため、会場・飲食・体験設計に限界があります。\n"
         "スポンサー協力により、クリエイターが撮影・交流・体験しやすい場をつくり、"
         "スポンサー企業にとっても自然な接点とUGC創出につながるイベントに進化できます。",
         Inches(0.85), Inches(1.83), Inches(11.5), Inches(0.95),
         size=14, color=WHITE)

# Participant requests
add_text(sl, "参加者アンケートより — 次回への期待・改善要望",
         Inches(0.65), Inches(3.05), Inches(12), Inches(0.5),
         size=14, bold=True, color=ACCENT_B)
requests = [
    "もっと長い時間交流したい",
    "より雰囲気の良い会場",
    "食事・ドリンクの充実",
    "ミニゲームや交流企画",
    "スポンサーや会場とのコラボ企画",
    "昼開催や頻度アップ",
]
for i, r in enumerate(requests):
    bx = Inches(0.65 + (i % 3) * 4.05)
    by = Inches(3.6 + (i // 3) * 0.85)
    add_rect(sl, bx, by + Inches(0.2), Inches(0.25), Inches(0.25), fill=ACCENT_PK)
    add_text(sl, r, bx + Inches(0.4), by, Inches(3.5), Inches(0.75), size=14, color=WHITE)

# CTA
card(sl, Inches(0.65), Inches(6.2), Inches(11.9), Inches(0.9),
     fill=RGBColor(0x20,0x05,0x35), border=ACCENT_P)
add_text(sl,
         "これらすべての改善は「スポンサー協力」によって実現できます。あなたの場所・商品・ブランドが体験の中心になります。",
         Inches(0.85), Inches(6.3), Inches(11.5), Inches(0.7),
         size=14, bold=True, color=ACCENT_P, align=PP_ALIGN.CENTER)

accent_bar(sl, ACCENT_P)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 10  Value for Sponsors
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "スポンサーが得られる 5 つの価値", title_size=30)
divider(sl, Inches(1.45))

values = [
    ("01", "グローバルクリエイターとの直接接点",
     "関西在住・関西発信のインフルエンサーに直接ブランドを体験してもらえる"),
    ("02", "訪日客向けSNS発信のきっかけづくり",
     "海外フォロワーを持つクリエイターが自然に投稿・紹介したくなる導線を設計"),
    ("03", "商品・店舗・施設の体験導入",
     "広告ではなくリアルな体験として商品・サービスを届けるUGC起点のPR"),
    ("04", "公式レポート・写真素材・SNS露出",
     "公式アカウントでのイベントレポート投稿、写真素材の提供、SNS掲載"),
    ("05", "関西インバウンドコミュニティへのブランド浸透",
     "単発広告ではなく、継続的なコミュニティとの関係構築で認知を積み上げる"),
]
for i, (num, title, desc) in enumerate(values):
    by = Inches(1.75 + i * 1.08)
    add_rect(sl, Inches(0.65), by + Inches(0.1), Inches(0.06), Inches(0.75), fill=ACCENT_P)
    add_text(sl, num,   Inches(0.85), by + Inches(0.08), Inches(0.7),  Inches(0.55),
             size=15, bold=True, color=ACCENT_P)
    add_text(sl, title, Inches(1.55), by + Inches(0.05), Inches(5.0),  Inches(0.55),
             size=16, bold=True, color=WHITE)
    add_text(sl, desc,  Inches(1.55), by + Inches(0.5),  Inches(11.2), Inches(0.55),
             size=13, color=LIGHT_GRAY)

accent_bar(sl, ACCENT_B)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 11  Activation Ideas
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "スポンサー参加の具体例", title_size=30,
             subtitle="業種別 — スポンサー活用イメージ")
divider(sl, Inches(1.5))

acts = [
    ("🏨 ホテル",
     "会場提供 / 館内ツアー\nレストラン・バー紹介\n公式レポート投稿"),
    ("🍽 飲食店",
     "会場提供 / フード提供\nクリエイター向け試食体験\n訪日客向け認知拡大"),
    ("🥤 飲料・食品",
     "商品サンプリング\n乾杯ドリンク提供\nフォトスポット設置"),
    ("🗾 観光・自治体",
     "地域PR / クリエイターツアー導線\nローカル体験紹介\n将来的なFAMツアー企画"),
    ("⚽ スポーツ",
     "試合観戦導線\nスタジアム・グッズ体験\n海外ファン向け認知拡大"),
]
for i, (cat, items) in enumerate(acts):
    bx = Inches(0.4 + i * 2.55)
    card(sl, bx, Inches(1.85), Inches(2.45), Inches(5.05),
         fill=RGBColor(0x0D,0x22,0x4A), border=ACCENT_B)
    add_text(sl, cat, bx + Inches(0.1), Inches(1.95), Inches(2.3), Inches(0.65),
             size=14, bold=True, color=ACCENT_B)
    add_text(sl, items, bx + Inches(0.1), Inches(2.65), Inches(2.3), Inches(4.1),
             size=12, color=WHITE)

accent_bar(sl, ACCENT_P)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 12  Sponsorship Menu
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "Sponsorship Menu", title_size=32)
divider(sl, Inches(1.4))

plans = [
    ("Venue\nSponsor",    ACCENT_B,
     "会場提供\nドリンク・軽食提供\nロゴ掲載\n公式SNS紹介\n当日写真・レポート共有",
     "会場提供\nまたは飲食提供"),
    ("Product\nSampling", ACCENT_PK,
     "商品サンプリング\n体験ブース設置\n参加者への商品配布\n公式SNS紹介",
     "商品提供\n+ 協賛費 5〜10万円"),
    ("Main\nSponsor",     ACCENT_P,
     "イベントメイン協賛\n会場内ブランド露出\n1分ブランド紹介\n公式レポート投稿\nクリエイター交流機会",
     "15万〜30万円"),
    ("Title\nSponsor",    GOLD,
     "イベント名への冠掲載\nトップスポンサー掲載\n企画設計から共同実施\nクリエイター体験導線設計\n別途インフルエンサー投稿施策相談",
     "30万〜50万円\n以上"),
]
for i, (name, col, items, price) in enumerate(plans):
    bx = Inches(0.5 + i * 3.2)
    card(sl, bx, Inches(1.7), Inches(3.05), Inches(5.15),
         fill=RGBColor(0x0D,0x22,0x4A), border=col)
    # top color bar
    add_rect(sl, bx, Inches(1.7), Inches(3.05), Inches(0.4), fill=col)
    add_text(sl, name, bx + Inches(0.1), Inches(1.72), Inches(2.9), Inches(0.55),
             size=14, bold=True, color=NAVY)
    add_text(sl, items, bx + Inches(0.1), Inches(2.2), Inches(2.9), Inches(2.8),
             size=11.5, color=WHITE)
    # price box
    add_rect(sl, bx, Inches(6.65), Inches(3.05), Inches(0.65), fill=col)
    add_text(sl, price, bx + Inches(0.05), Inches(6.67), Inches(2.98), Inches(0.6),
             size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

add_text(sl,
         "※料金は提案用の目安です。確定料金は開催規模・会場・提供内容により個別見積もりとなります。",
         Inches(0.5), Inches(7.1), Inches(12.5), Inches(0.35),
         size=9, color=RGBColor(0x77,0x88,0xAA), italic=True)
accent_bar(sl, GOLD)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 13  UGC & Exposure
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "自然なUGCを生み出す設計", title_size=30,
             subtitle="投稿・露出について")
divider(sl, Inches(1.5))

ugc = [
    ("📸", "撮影しやすい導線",
     "ブランド体験・フォトスポット・商品展示を設計し、自然に写真・動画が生まれる環境を用意"),
    ("🎉", "体験コンテンツ",
     "商品サンプリング・試食・交流企画を通じて、クリエイターが投稿したくなる体験を提供"),
    ("📢", "公式アカウント発信",
     "イベントレポートを公式SNSアカウントで投稿、スポンサー企業・会場を紹介"),
    ("🔗", "任意投稿を促進",
     "参加クリエイターへの自然な体験提供により、任意投稿・UGCの創出を促進"),
    ("⭐", "投稿保証オプション",
     "投稿数・リーチを保証するキャスティング施策は別途有料オプションとして設計可能"),
]
for i, (icon, title, desc) in enumerate(ugc):
    by = Inches(1.85 + i * 1.05)
    card(sl, Inches(0.65), by, Inches(11.9), Inches(0.9),
         fill=RGBColor(0x0D,0x1E,0x42), border=ACCENT_B if i < 4 else ACCENT_P)
    add_text(sl, icon, Inches(0.75), by + Inches(0.12), Inches(0.7), Inches(0.65), size=20)
    add_text(sl, title, Inches(1.5), by + Inches(0.08), Inches(3.0), Inches(0.45),
             size=14, bold=True, color=ACCENT_B if i < 4 else ACCENT_P)
    add_text(sl, desc, Inches(4.6), by + Inches(0.08), Inches(7.7), Inches(0.75),
             size=13, color=WHITE)

add_text(sl,
         "※参加者の個別投稿は任意であり、投稿本数・再生数を保証するものではありません。",
         Inches(0.65), Inches(7.15), Inches(12.5), Inches(0.35),
         size=9, color=RGBColor(0x77,0x88,0xAA), italic=True)
accent_bar(sl, ACCENT_B)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 14  Event Flow
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "イベント当日の流れ", title_size=32)
divider(sl, Inches(1.45))

steps = [
    ("①", "受付・チェックイン",       "参加者迎え入れ、名札配布"),
    ("②", "ウェルカムドリンク",        "スポンサー提供ドリンクで乾杯"),
    ("③", "主催挨拶",                 "イベント趣旨・スポンサー紹介"),
    ("④", "クリエイター自己紹介",      "全員で簡単な自己紹介・SNS交換"),
    ("⑤", "商品・施設体験",           "スポンサーサンプリング・展示・体験コーナー"),
    ("⑥", "ミニゲーム・ネットワーキング", "交流企画・フォトタイム"),
    ("⑦", "集合写真",                 "公式レポート用撮影"),
    ("⑧", "アフターレポート配信",      "公式SNSでイベントレポート投稿"),
]
for i, (num, step, note) in enumerate(steps):
    bx = Inches(0.5 + (i % 4) * 3.2)
    by = Inches(1.8 + (i // 4) * 2.55)
    card(sl, bx, by, Inches(3.0), Inches(2.3),
         fill=RGBColor(0x0D,0x22,0x4A), border=ACCENT_P)
    add_text(sl, num, bx + Inches(0.1), by + Inches(0.1), Inches(2.8), Inches(0.55),
             size=24, bold=True, color=ACCENT_P)
    add_text(sl, step, bx + Inches(0.1), by + Inches(0.65), Inches(2.8), Inches(0.75),
             size=13, bold=True, color=WHITE)
    add_text(sl, note, bx + Inches(0.1), by + Inches(1.4), Inches(2.8), Inches(0.8),
             size=11, color=LIGHT_GRAY)

accent_bar(sl, ACCENT_P)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 15  Future Vision
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "関西発のグローバルクリエイターコミュニティへ", title_size=27,
             subtitle="今後の展開 — Next Steps")
divider(sl, Inches(1.55))

visions = [
    ("🔄 定期開催",
     "年複数回の定期開催でコミュニティを継続成長させる"),
    ("🤝 コラボイベント",
     "スポンサー企業との共同イベント・コラボ企画を実施"),
    ("✈️ FAMツアー",
     "地域・観光施設とのファムトリップでインバウンドPRを強化"),
    ("🏙 インバウンドPR支援",
     "飲食店・ホテル・商業施設の訪日客向けPRをクリエイター起点で支援"),
    ("🌐 ネットワーク拡大",
     "関西から日本全国・海外へ発信するグローバルクリエイターネットワークへ"),
]
for i, (title, desc) in enumerate(visions):
    bx = Inches(0.65 + (i % 3) * 4.1) if i < 3 else Inches(0.65 + (i - 3) * 6.15 + 1.0)
    by = Inches(1.9) if i < 3 else Inches(4.55)
    bw = Inches(3.8) if i < 3 else Inches(4.9)
    card(sl, bx, by, bw, Inches(2.3), fill=RGBColor(0x0D,0x22,0x4A), border=ACCENT_B)
    add_text(sl, title, bx + Inches(0.15), by + Inches(0.15), bw - Inches(0.3), Inches(0.7),
             size=14, bold=True, color=ACCENT_B)
    add_text(sl, desc, bx + Inches(0.15), by + Inches(0.85), bw - Inches(0.3), Inches(1.3),
             size=13, color=WHITE)

accent_bar(sl, ACCENT_B)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 16  Sponsor Targets
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl)
slide_header(sl, "想定スポンサー領域", title_size=32,
             subtitle="ご提案先・協賛候補カテゴリ（候補として記載）")
divider(sl, Inches(1.5))

targets = [
    ("🏨 ホテル・宿泊",
     "W Osaka、Moxy など\n関西の個性的なホテル・ブティックホテル",
     ACCENT_B),
    ("🍺 飲料・酒類",
     "日本盛、アサヒビール、伊藤園 など\n乾杯・サンプリングスポンサーに最適",
     ACCENT_PK),
    ("🍫 食品・お菓子",
     "ブラックサンダー、森永 など\n参加者へのサンプリング・体験展示",
     ACCENT_P),
    ("⚽ スポーツ",
     "セレッソ大阪 など\n海外ファン向け認知拡大・体験イベント",
     GOLD),
    ("🗾 観光・自治体",
     "観光施設・地域自治体\nインバウンドPR・ローカル体験発信",
     ACCENT_B),
    ("🏬 商業施設",
     "商業施設・ショッピングモール\nインバウンド消費促進・クリエイター誘導",
     ACCENT_P),
]
for i, (cat, desc, col) in enumerate(targets):
    bx = Inches(0.5 + (i % 3) * 4.2)
    by = Inches(1.85 + (i // 3) * 2.45)
    card(sl, bx, by, Inches(4.0), Inches(2.2), fill=RGBColor(0x0D,0x22,0x4A), border=col)
    add_rect(sl, bx, by, Inches(4.0), Inches(0.35), fill=col)
    add_text(sl, cat, bx + Inches(0.1), by + Inches(0.35), Inches(3.8), Inches(0.6),
             size=14, bold=True, color=col)
    add_text(sl, desc, bx + Inches(0.1), by + Inches(0.9), Inches(3.8), Inches(1.2),
             size=12, color=WHITE)

add_text(sl, "※記載はすべて候補・想定として表記しています。営業先に応じてスライドを差し替え可能です。",
         Inches(0.5), Inches(7.1), Inches(12.5), Inches(0.35),
         size=9, color=RGBColor(0x77,0x88,0xAA), italic=True)
accent_bar(sl, ACCENT_P)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 17  Closing
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
add_gradient_bg(sl, NAVY, RGBColor(0x08,0x14,0x30))

# decorative
for cx, cy, r, col in [
    (Inches(12.5), Inches(0.5), Inches(2.8), ACCENT_P),
    (Inches(0.5),  Inches(7.0), Inches(2.0), ACCENT_B),
]:
    blob = sl.shapes.add_shape(9, cx - r, cy - r, r*2, r*2)
    blob.fill.solid(); blob.fill.fore_color.rgb = col
    blob.line.fill.background()

add_text(sl, "関西の魅力を、世界に届く体験へ。",
         Inches(0.8), Inches(1.1), Inches(11.5), Inches(1.2),
         size=38, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(sl,
         "スポンサー企業様とともに、グローバルクリエイターが集まる場をつくりたいと思います。\n"
         "商品・施設・地域の魅力を、リアルな体験として届けましょう。\n"
         "会場提供・商品提供・協賛など、形は柔軟にご相談いただけます。",
         Inches(1.5), Inches(2.5), Inches(10.3), Inches(1.8),
         size=16, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

divider(sl, Inches(4.5), color=ACCENT_P, width=Inches(8.0), x=Inches(2.65))

# Contact card
card(sl, Inches(2.65), Inches(4.75), Inches(8.0), Inches(2.2),
     fill=RGBColor(0x10,0x28,0x5A), border=ACCENT_P)
contacts = [
    ("Company",   "株式会社 Mr.sasuke"),
    ("Contact",   "西村 雄矢"),
    ("Instagram", "@mr.sasuke_japan"),
    ("Email",     "（担当者よりご連絡いたします）"),
]
for i, (k, v) in enumerate(contacts):
    bx = Inches(3.0)
    by = Inches(4.95 + i * 0.48)
    add_text(sl, k + " :", bx, by, Inches(1.5), Inches(0.45),
             size=13, color=ACCENT_B, bold=True)
    add_text(sl, v, bx + Inches(1.6), by, Inches(6.5), Inches(0.45),
             size=13, color=WHITE)

add_text(sl, "Presented by  Mr.sasuke inc.",
         Inches(0), Inches(7.1), Inches(13.33), Inches(0.35),
         size=11, color=RGBColor(0x55,0x66,0x88), align=PP_ALIGN.CENTER)
accent_bar(sl, ACCENT_P)

# ─────────────────────────────────────────────────────────────────────────────
# Save
# ─────────────────────────────────────────────────────────────────────────────
out = "/home/user/-/KansaiGlobalCreatorsMeetup_SponsorshipProposal.pptx"
prs.save(out)
print(f"Saved → {out}")
print(f"Slides: {len(prs.slides)}")
