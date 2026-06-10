from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

# ── Color Palette ─────────────────────────────────────────────────────────────
NAVY        = RGBColor(0x0D, 0x1F, 0x3C)   # deep navy (headings, accents)
NAVY_MID    = RGBColor(0x1A, 0x3A, 0x6E)   # mid navy
ACCENT_B    = RGBColor(0x1A, 0x6F, 0xC4)   # blue
ACCENT_P    = RGBColor(0x7C, 0x3A, 0xED)   # purple
ACCENT_PK   = RGBColor(0xE0, 0x45, 0x8A)   # pink
ACCENT_TEAL = RGBColor(0x0B, 0x9A, 0x8A)   # teal
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE   = RGBColor(0xF7, 0xF8, 0xFC)
LIGHT_BLUE  = RGBColor(0xE8, 0xF1, 0xFB)
LIGHT_GRAY  = RGBColor(0xF0, 0xF2, 0xF7)
MID_GRAY    = RGBColor(0x8A, 0x96, 0xAA)
DARK_GRAY   = RGBColor(0x3A, 0x4A, 0x5C)
GOLD        = RGBColor(0xF0, 0xA8, 0x00)

W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]

# ── Helpers ───────────────────────────────────────────────────────────────────

def add_rect(slide, l, t, w, h, fill=None, line_color=None, line_pt=1):
    shape = slide.shapes.add_shape(1, l, t, w, h)
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_pt)
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, l, t, w, h,
             size=14, bold=False, color=DARK_GRAY,
             align=PP_ALIGN.LEFT, italic=False):
    txb = slide.shapes.add_textbox(l, t, w, h)
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.italic = italic
    return txb

def white_bg(slide):
    add_rect(slide, 0, 0, W, H, fill=WHITE)

def top_bar(slide, color=NAVY, h=Inches(0.08)):
    add_rect(slide, 0, 0, W, h, fill=color)

def bottom_bar(slide, color=NAVY, h=Pt(4)):
    add_rect(slide, 0, H - h, W, h, fill=color)

def left_accent(slide, color=ACCENT_B, w=Pt(5), t=Inches(0.7), h=Inches(0.6)):
    add_rect(slide, Inches(0.55), t, w, h, fill=color)

def slide_header(slide, title, subtitle=None, title_size=28, sub_size=13,
                 title_color=NAVY, sub_color=MID_GRAY):
    top_bar(slide, NAVY)
    add_text(slide, title,
             Inches(0.7), Inches(0.22), Inches(11.9), Inches(0.85),
             size=title_size, bold=True, color=title_color)
    if subtitle:
        add_text(slide, subtitle,
                 Inches(0.7), Inches(0.95), Inches(11.9), Inches(0.45),
                 size=sub_size, color=sub_color)
    # thin divider
    add_rect(slide, Inches(0.7), Inches(1.35), Inches(11.93), Pt(1.5), fill=LIGHT_GRAY)
    bottom_bar(slide)

def kpi_card(slide, number, label, l, t, w=Inches(2.7), h=Inches(1.55),
             num_color=ACCENT_B, bg=LIGHT_BLUE, border=None):
    add_rect(slide, l, t, w, h, fill=bg,
             line_color=border or ACCENT_B, line_pt=1.2)
    add_text(slide, number, l, t + Inches(0.12), w, Inches(0.85),
             size=38, bold=True, color=num_color, align=PP_ALIGN.CENTER)
    add_text(slide, label, l, t + Inches(0.88), w, Inches(0.6),
             size=11, color=DARK_GRAY, align=PP_ALIGN.CENTER)

def section_tag(slide, text, l, t, color=ACCENT_B):
    add_rect(slide, l, t, Inches(0.04), Inches(0.4), fill=color)
    add_text(slide, text, l + Inches(0.12), t - Inches(0.02), Inches(6), Inches(0.45),
             size=12, bold=True, color=color)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 1  Cover
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)

# Left panel — navy
add_rect(sl, 0, 0, Inches(5.8), H, fill=NAVY)
# thin color strip between panels
add_rect(sl, Inches(5.8), 0, Inches(0.06), H, fill=ACCENT_B)

# Right panel subtle grid lines
for i in range(8):
    add_rect(sl, Inches(6.0), Inches(i * 0.95), Inches(7.33), Pt(0.5),
             fill=LIGHT_GRAY)

# Tag
add_text(sl, "SPONSORSHIP PROPOSAL  2025",
         Inches(0.55), Inches(1.0), Inches(5.0), Inches(0.45),
         size=11, color=ACCENT_B, bold=True)

# Title
add_text(sl, "Kansai Global",
         Inches(0.55), Inches(1.55), Inches(5.2), Inches(1.15),
         size=52, bold=True, color=WHITE)
add_text(sl, "Creators Meetup",
         Inches(0.55), Inches(2.55), Inches(5.2), Inches(1.15),
         size=52, bold=True, color=ACCENT_B)

# Subtitle JP
add_rect(sl, Inches(0.55), Inches(3.75), Inches(4.9), Pt(2), fill=ACCENT_B)
add_text(sl,
         "関西から世界へ。訪日観光・ローカル体験を発信する\nグローバルクリエイターコミュニティ",
         Inches(0.55), Inches(3.95), Inches(5.0), Inches(0.95),
         size=13.5, color=RGBColor(0xCC, 0xD8, 0xF0))

# Presented by
add_text(sl, "Presented by  Mr.sasuke inc.",
         Inches(0.55), Inches(6.7), Inches(5.0), Inches(0.5),
         size=12, color=MID_GRAY)

# Right panel KPIs
kpi_data = [
    ("35",    "Creators in Network", ACCENT_B),
    ("4.7M+", "Instagram Followers",  ACCENT_P),
    ("30–50", "Expected Attendees",   ACCENT_PK),
    ("10+",   "Countries Reached",    ACCENT_TEAL),
]
for i, (n, l, c) in enumerate(kpi_data):
    bx = Inches(6.3 + (i % 2) * 3.4)
    by = Inches(1.5 + (i // 2) * 2.4)
    add_rect(sl, bx, by, Inches(3.0), Inches(1.9),
             fill=WHITE, line_color=c, line_pt=1.5)
    add_rect(sl, bx, by, Inches(3.0), Inches(0.35), fill=c)
    add_text(sl, n, bx, by + Inches(0.42), Inches(3.0), Inches(0.95),
             size=42, bold=True, color=c, align=PP_ALIGN.CENTER)
    add_text(sl, l, bx, by + Inches(1.32), Inches(3.0), Inches(0.5),
             size=11.5, color=DARK_GRAY, align=PP_ALIGN.CENTER)

bottom_bar(sl, ACCENT_B)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 2  Concept
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, '関西の「リアルな魅力」を世界へ届けるクリエイターコミュニティ', title_size=24)

points = [
    (ACCENT_B,  "Global Creators Hub",
     "関西在住・関西に強いグローバルクリエイターが集まるオフライン交流会。リアルなつながりから発信が生まれます。"),
    (ACCENT_P,  "Inbound SNS Amplification",
     "訪日外国人に向けてグルメ・観光・文化・ローカル体験を発信。フォロワーの多くは日本旅行に関心を持つ海外層。"),
    (ACCENT_PK, "Community Connection",
     "SNS発信者同士がつながり、関西の魅力をより広く届けるコミュニティ拠点。継続的な関係構築が可能です。"),
]
for i, (col, ttl, desc) in enumerate(points):
    by = Inches(1.65 + i * 1.8)
    add_rect(sl, Inches(0.55), by, Inches(0.35), Inches(1.55), fill=col)
    add_rect(sl, Inches(0.9), by, Inches(11.8), Inches(1.55), fill=LIGHT_GRAY)
    add_text(sl, ttl,  Inches(1.1), by + Inches(0.18), Inches(11.3), Inches(0.55),
             size=17, bold=True, color=NAVY)
    add_text(sl, desc, Inches(1.1), by + Inches(0.72), Inches(11.3), Inches(0.75),
             size=13.5, color=DARK_GRAY)

bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 3  Why Now
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "インバウンドPRは「広告」から「体験され、語られる場」へ",
             subtitle="Why This Event Matters — 今このイベントが必要な理由", title_size=23)

whys = [
    ("訪日客はSNSで行き先・食事・体験を決める時代",
     "旅行前に検索するのはガイドブックではなくInstagram・TikTok・YouTube。インフルエンサーの発信が意思決定を左右します。"),
    ("海外目線で発信できるクリエイターとの接点が重要",
     "日本語が話せない訪日客に届く発信をするには、英語・多言語で発信するクリエイターとのパートナーシップが不可欠です。"),
    ("関西のローカル資源は、海外にまだ届き切っていない",
     "大阪・京都以外にも魅力的な体験・グルメ・カルチャーが多数あります。クリエイターが発信することで認知が広がります。"),
    ("体験→投稿→拡散のサイクルをつくることが鍵",
     "広告を見せるのではなく、クリエイターが実際に体験し、自然にシェアしたくなる場を設計することが重要です。"),
]
for i, (ttl, desc) in enumerate(whys):
    bx = Inches(0.55 + (i % 2) * 6.3)
    by = Inches(1.6 + (i // 2) * 2.4)
    add_rect(sl, bx, by, Inches(5.9), Inches(2.15), fill=OFF_WHITE, line_color=LIGHT_GRAY)
    add_rect(sl, bx, by, Inches(5.9), Inches(0.05), fill=ACCENT_B)
    add_text(sl, f"0{i+1}", bx + Inches(0.18), by + Inches(0.15), Inches(0.6), Inches(0.65),
             size=28, bold=True, color=RGBColor(0xD0,0xE4,0xF8))
    add_text(sl, ttl, bx + Inches(0.75), by + Inches(0.15), Inches(5.0), Inches(0.6),
             size=14, bold=True, color=NAVY)
    add_text(sl, desc, bx + Inches(0.18), by + Inches(0.8), Inches(5.55), Inches(1.25),
             size=12.5, color=DARK_GRAY)

bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 4  Event Overview
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "Kansai Global Creators Meetup とは",
             subtitle="Event Overview — イベント概要", title_size=28)

items = [
    ("📍", "開催地",       "関西エリア（大阪・京都・神戸 等）",              ACCENT_B),
    ("👥", "参加者",       "関西在住／関西発信のグローバルクリエイター\nインバウンド系インフルエンサー", ACCENT_P),
    ("🎯", "内容",        "ネットワーキング・情報交換・スポンサー体験\n商品紹介・撮影・交流",         ACCENT_PK),
    ("📊", "想定規模",     "30〜50名程度\n（開催条件により変動）",            ACCENT_TEAL),
    ("🔄", "開催頻度",     "定期開催を想定（年複数回）",                       ACCENT_B),
    ("🎥", "参加ジャンル", "旅行・グルメ・日本語/文化\nライフスタイル・ホテル・ローカル体験",       ACCENT_P),
]
for i, (icon, label, val, col) in enumerate(items):
    bx = Inches(0.55 + (i % 3) * 4.25)
    by = Inches(1.6 + (i // 3) * 2.55)
    add_rect(sl, bx, by, Inches(4.0), Inches(2.3), fill=WHITE, line_color=LIGHT_GRAY)
    add_rect(sl, bx, by, Inches(4.0), Inches(0.42), fill=col)
    add_text(sl, label, bx + Inches(0.12), by + Inches(0.08), Inches(3.75), Inches(0.3),
             size=12, bold=True, color=WHITE)
    add_text(sl, icon + "  " + val, bx + Inches(0.12), by + Inches(0.55),
             Inches(3.75), Inches(1.65), size=14, color=DARK_GRAY)

bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 5  Creator Network
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "Creator Network",
             subtitle="関西インバウンドに強いグローバルクリエイターネットワーク", title_size=30)

kpi_data = [
    ("35",    "Creators\nin Network",    ACCENT_B,  LIGHT_BLUE),
    ("4.7M+", "Instagram\nFollowers",    ACCENT_P,  RGBColor(0xF0,0xEB,0xFF)),
    ("30–50", "Expected\nAttendees",     ACCENT_PK, RGBColor(0xFF, 0xF0, 0xF6)),
    ("10+",   "Countries\nRepresented",  ACCENT_TEAL, RGBColor(0xE6,0xF8,0xF6)),
]
for i, (n, l, col, bg) in enumerate(kpi_data):
    kpi_card(sl, n, l, Inches(0.55 + i * 3.2), Inches(1.6),
             w=Inches(3.0), h=Inches(1.75), num_color=col, bg=bg, border=col)

# Creator examples header
section_tag(sl, "ネットワーク内 招待候補クリエイター例", Inches(0.55), Inches(3.65))

creators = [
    ("100万超",  "Japan Travel\n/ Halal",   ACCENT_B),
    ("90万規模", "Japan\nLifestyle",         ACCENT_P),
    ("39万規模", "日本語・\nカルチャー",      ACCENT_PK),
    ("28万規模", "Japan\nTravel",            ACCENT_TEAL),
    ("21万規模", "Kansai\nLife",             ACCENT_B),
]
for i, (f, g, col) in enumerate(creators):
    bx = Inches(0.55 + i * 2.55)
    by = Inches(4.1)
    add_rect(sl, bx, by, Inches(2.35), Inches(2.2), fill=WHITE, line_color=col, line_pt=1.5)
    add_rect(sl, bx, by, Inches(2.35), Inches(0.38), fill=col)
    add_text(sl, f,  bx, by + Inches(0.45), Inches(2.35), Inches(0.75),
             size=22, bold=True, color=col, align=PP_ALIGN.CENTER)
    add_text(sl, g,  bx, by + Inches(1.2),  Inches(2.35), Inches(0.9),
             size=11.5, color=DARK_GRAY, align=PP_ALIGN.CENTER)

add_text(sl,
         "※参加者数・投稿数・再生数は保証値ではありません。開催条件・参加可否・企画内容により変動します。投稿保証型PRは別途オプションとして設計可能です。",
         Inches(0.55), Inches(7.05), Inches(12.5), Inches(0.35),
         size=9, color=MID_GRAY, italic=True)
bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 6  Past Results
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "過去開催実績",
             subtitle="Track Record — 実績データ", title_size=30)

# Event 1 & 2 cards
for idx, (ev, att, fol, avg_lbl, avg_val, col) in enumerate([
    ("第1回", "30名", "約871万", "平均フォロワー", "約290,547", ACCENT_B),
    ("第2回", "30名", "約181万", "平均フォロワー", "約60,525",  ACCENT_P),
]):
    bx = Inches(0.55 + idx * 6.25)
    add_rect(sl, bx, Inches(1.6), Inches(5.95), Inches(3.75), fill=WHITE, line_color=LIGHT_GRAY)
    add_rect(sl, bx, Inches(1.6), Inches(5.95), Inches(0.5), fill=col)
    add_text(sl, ev, bx + Inches(0.2), Inches(1.62), Inches(5.5), Inches(0.45),
             size=16, bold=True, color=WHITE)
    for j, (lbl, val, vc) in enumerate([
        ("参加者数",          att, NAVY),
        ("SNS合計フォロワー",  fol, col),
        (avg_lbl,             avg_val, DARK_GRAY),
    ]):
        add_text(sl, lbl, bx + Inches(0.2), Inches(2.3 + j * 1.15), Inches(5.5), Inches(0.4),
                 size=11, color=MID_GRAY)
        add_text(sl, val, bx + Inches(0.2), Inches(2.65 + j * 1.15), Inches(5.5), Inches(0.6),
                 size=28 if j < 2 else 18, bold=True, color=vc)
        if j < 2:
            add_rect(sl, bx + Inches(0.2), Inches(3.2 + j * 1.15), Inches(5.5), Pt(0.8),
                     fill=LIGHT_GRAY)

# Bottom: network total
add_rect(sl, Inches(0.55), Inches(5.65), Inches(12.25), Inches(1.4),
         fill=LIGHT_BLUE, line_color=ACCENT_B)
add_text(sl,
         "関西クリエイターネットワーク  35名  ·  Instagram合計 約470万フォロワー  ·  トップ層：100万超クリエイター在籍",
         Inches(0.75), Inches(5.82), Inches(11.85), Inches(0.6),
         size=15, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
add_text(sl, "※数字は参加者・ネットワークのSNSフォロワー集計に基づく参考値です。",
         Inches(0.75), Inches(6.48), Inches(11.85), Inches(0.35),
         size=9, color=MID_GRAY, italic=True, align=PP_ALIGN.CENTER)
bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 7  Audience Profile
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "海外に届く発信力を持つクリエイターが集まる",
             subtitle="Audience Profile — グローバルオーディエンス属性", title_size=25)

# Countries section
section_tag(sl, "主なフォロワー国", Inches(0.55), Inches(1.58))
countries = [
    "🇺🇸 USA", "🇯🇵 Japan", "🇮🇳 India", "🇬🇧 UK", "🇫🇷 France",
    "🇸🇬 Singapore", "🇩🇪 Germany", "🇮🇩 Indonesia",
    "🇦🇺 Australia", "🇵🇭 Philippines", "🇨🇦 Canada",
    "🇻🇳 Vietnam", "🇹🇭 Thailand",
]
for i, c in enumerate(countries):
    bx = Inches(0.55 + (i % 7) * 1.84)
    by = Inches(2.05 + (i // 7) * 0.72)
    add_rect(sl, bx, by, Inches(1.75), Inches(0.6), fill=LIGHT_GRAY, line_color=LIGHT_GRAY)
    add_text(sl, c, bx + Inches(0.08), by + Inches(0.07), Inches(1.6), Inches(0.48),
             size=12.5, color=DARK_GRAY)

# Age bars
section_tag(sl, "主なフォロワー年齢層", Inches(0.55), Inches(3.68), color=ACCENT_P)
age_data = [
    ("25〜34歳", 0.42, ACCENT_P),
    ("35〜44歳", 0.30, ACCENT_B),
    ("18〜24歳", 0.20, ACCENT_PK),
]
for i, (label, ratio, col) in enumerate(age_data):
    by = Inches(4.15 + i * 0.75)
    add_text(sl, label, Inches(0.55), by, Inches(1.3), Inches(0.55), size=13, color=DARK_GRAY)
    add_rect(sl, Inches(2.0), by + Inches(0.1), Inches(7.0), Inches(0.38), fill=LIGHT_GRAY)
    add_rect(sl, Inches(2.0), by + Inches(0.1), Inches(7.0 * ratio), Inches(0.38), fill=col)
    add_text(sl, f"{int(ratio*100)}%", Inches(2.0) + Inches(7.0 * ratio) + Inches(0.12), by,
             Inches(0.65), Inches(0.55), size=13, bold=True, color=col)

# Genres
section_tag(sl, "発信ジャンル", Inches(0.55), Inches(6.28), color=ACCENT_PK)
genres = ["Japan Travel", "Osaka/Kansai Food", "Japanese Culture", "Local Experience",
          "Language Learning", "Lifestyle", "Hotel / Tourism"]
for i, g in enumerate(genres):
    bx = Inches(0.55 + i * 1.83)
    add_rect(sl, bx, Inches(6.72), Inches(1.73), Inches(0.58),
             fill=WHITE, line_color=ACCENT_PK)
    add_text(sl, g, bx + Inches(0.06), Inches(6.78), Inches(1.62), Inches(0.48),
             size=10.5, color=ACCENT_PK, align=PP_ALIGN.CENTER)

bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 8  Survey Results
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "参加者満足度の高いコミュニティ",
             subtitle="第2回アンケート結果（回答数：16件）", title_size=28)

scores = [
    ("イベント全体満足度",  "4.88", ACCENT_B,  LIGHT_BLUE),
    ("内容の有益性",       "4.75", ACCENT_P,  RGBColor(0xF0,0xEB,0xFF)),
    ("今後も参加したい",   "5.0",  ACCENT_PK, RGBColor(0xFF, 0xF0, 0xF6)),
]
for i, (label, num, col, bg) in enumerate(scores):
    bx = Inches(0.55 + i * 4.25)
    add_rect(sl, bx, Inches(1.6), Inches(4.0), Inches(2.5), fill=bg, line_color=col, line_pt=1.5)
    add_rect(sl, bx, Inches(1.6), Inches(4.0), Inches(0.42), fill=col)
    add_text(sl, label, bx + Inches(0.1), Inches(1.62), Inches(3.8), Inches(0.38),
             size=12, bold=True, color=WHITE)
    add_text(sl, num, bx, Inches(2.1), Inches(4.0), Inches(1.2),
             size=58, bold=True, color=col, align=PP_ALIGN.CENTER)
    add_text(sl, "/ 5", bx, Inches(3.1), Inches(4.0), Inches(0.5),
             size=18, color=MID_GRAY, align=PP_ALIGN.CENTER)

# Additional scores
for i, (lbl, val) in enumerate([
    ("スペース満足度", "4.50/5"),
    ("参加費の妥当性", "4.38/5"),
    ("準備物の充実度", "4.31/5"),
]):
    bx = Inches(0.55 + i * 4.25)
    add_text(sl, lbl + "  " + val,
             bx, Inches(4.3), Inches(4.0), Inches(0.5),
             size=12, color=DARK_GRAY, align=PP_ALIGN.CENTER)

# Comments
section_tag(sl, "参加者の声", Inches(0.55), Inches(4.85))
comments = [
    '"Great atmosphere and great people."',
    '"Everyone welcoming and supporting each other."',
    '"関西人が多い事が要因なのか、初対面でも壁を感じなかった。"',
    '"みんなで関西を盛り上げていきたいという想いが素敵でした。"',
]
for i, c in enumerate(comments):
    bx = Inches(0.55 + (i % 2) * 6.3)
    by = Inches(5.3 + (i // 2) * 0.9)
    add_rect(sl, bx, by + Inches(0.08), Inches(0.04), Inches(0.55), fill=ACCENT_B)
    add_text(sl, c, bx + Inches(0.18), by, Inches(5.9), Inches(0.72),
             size=12.5, color=DARK_GRAY, italic=True)

bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 9  Why Sponsor
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "スポンサー協力により、より価値ある体験設計へ",
             subtitle="Why We Need Sponsors — スポンサーが必要な理由", title_size=24)

# Context box
add_rect(sl, Inches(0.55), Inches(1.6), Inches(12.25), Inches(1.15),
         fill=LIGHT_BLUE, line_color=ACCENT_B)
add_text(sl,
         "現在は参加費中心で運営しているため、会場・飲食・体験設計に限界があります。スポンサー協力により、クリエイターが"
         "撮影・交流・体験しやすい場をつくり、スポンサー企業にとっても自然な接点とUGC創出につながるイベントに進化できます。",
         Inches(0.78), Inches(1.7), Inches(11.85), Inches(0.95),
         size=13.5, color=NAVY)

section_tag(sl, "参加者アンケートより — 次回への期待・改善要望", Inches(0.55), Inches(2.95))
requests = [
    ("もっと長い時間交流したい",    "ゆったりとした時間設計・会場選定"),
    ("より雰囲気の良い会場",        "ホテル・バー・おしゃれな空間での開催"),
    ("食事・ドリンクの充実",        "スポンサー提供による飲食クオリティアップ"),
    ("ミニゲームや交流企画",        "スポンサーとのコラボ体験コンテンツ"),
    ("スポンサー・会場とのコラボ",  "ブランド体験をイベントに自然に組み込む"),
    ("昼開催・頻度アップ",          "複数スポンサーによる開催コスト分散"),
]
for i, (req, sol) in enumerate(requests):
    bx = Inches(0.55 + (i % 3) * 4.25)
    by = Inches(3.45 + (i // 3) * 1.6)
    add_rect(sl, bx, by, Inches(4.0), Inches(1.45), fill=OFF_WHITE, line_color=LIGHT_GRAY)
    add_rect(sl, bx, by, Inches(0.05), Inches(1.45), fill=ACCENT_PK)
    add_text(sl, req, bx + Inches(0.15), by + Inches(0.1), Inches(3.75), Inches(0.5),
             size=13, bold=True, color=NAVY)
    add_text(sl, "→ " + sol, bx + Inches(0.15), by + Inches(0.65), Inches(3.75), Inches(0.7),
             size=11.5, color=ACCENT_B)

bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 10  Value for Sponsors
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "スポンサーが得られる 5 つの価値",
             subtitle="Sponsor Benefits", title_size=28)

values = [
    (ACCENT_B,  "グローバルクリエイターとの直接接点",
     "関西在住・関西発信のインフルエンサーに直接ブランドを体験してもらえる機会を提供"),
    (ACCENT_P,  "訪日客向けSNS発信のきっかけづくり",
     "海外フォロワーを持つクリエイターが自然に投稿・紹介したくなる導線を設計します"),
    (ACCENT_PK, "商品・店舗・施設の体験導入",
     "広告ではなくリアルな体験として商品・サービスを届けるUGC起点のインバウンドPR"),
    (ACCENT_TEAL,"公式レポート・写真素材・SNS露出",
     "公式アカウントでのイベントレポート投稿、写真素材の提供、SNS掲載による露出"),
    (GOLD,       "関西インバウンドコミュニティへのブランド浸透",
     "単発広告ではなく継続的なコミュニティとの関係構築で認知を積み上げる長期戦略"),
]
for i, (col, title, desc) in enumerate(values):
    by = Inches(1.6 + i * 1.07)
    add_rect(sl, Inches(0.55), by, Inches(12.25), Inches(0.98), fill=OFF_WHITE, line_color=LIGHT_GRAY)
    add_rect(sl, Inches(0.55), by, Inches(0.42), Inches(0.98), fill=col)
    add_text(sl, f"0{i+1}", Inches(1.08), by + Inches(0.12), Inches(0.65), Inches(0.7),
             size=20, bold=True, color=col)
    add_text(sl, title, Inches(1.75), by + Inches(0.1), Inches(4.5), Inches(0.45),
             size=15, bold=True, color=NAVY)
    add_text(sl, desc,  Inches(6.4),  by + Inches(0.1), Inches(6.2), Inches(0.75),
             size=13, color=DARK_GRAY)

bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 11  Activation Ideas
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "スポンサー参加の具体例",
             subtitle="業種別 — スポンサー活用イメージ", title_size=28)

acts = [
    ("🏨", "ホテル・宿泊",
     "会場提供\n館内ツアー・ルーム体験\nレストラン・バー紹介\n公式レポート投稿",
     ACCENT_B),
    ("🍽", "飲食店・レストラン",
     "会場提供\nフード・ドリンク提供\nクリエイター向け試食体験\n訪日客向け認知拡大",
     ACCENT_P),
    ("🥤", "飲料・食品メーカー",
     "商品サンプリング\n乾杯ドリンク提供\nフォトスポット設置\n新商品体験展示",
     ACCENT_PK),
    ("🗾", "観光施設・自治体",
     "地域PR・ローカル体験紹介\nクリエイターツアー導線\n将来的なFAMツアー企画",
     ACCENT_TEAL),
    ("⚽", "スポーツチーム",
     "試合観戦・スタジアム体験\nグッズ体験・選手交流\n海外ファン向け認知拡大",
     GOLD),
]
for i, (icon, cat, items, col) in enumerate(acts):
    bx = Inches(0.42 + i * 2.5)
    add_rect(sl, bx, Inches(1.6), Inches(2.38), Inches(5.4), fill=WHITE, line_color=col, line_pt=1.5)
    add_rect(sl, bx, Inches(1.6), Inches(2.38), Inches(0.5), fill=col)
    add_text(sl, icon, bx + Inches(0.08), Inches(1.62), Inches(0.5), Inches(0.42), size=18, color=WHITE)
    add_text(sl, cat,  bx + Inches(0.58), Inches(1.65), Inches(1.75), Inches(0.42),
             size=12, bold=True, color=WHITE)
    add_text(sl, items, bx + Inches(0.12), Inches(2.25), Inches(2.2), Inches(4.6),
             size=12.5, color=DARK_GRAY)

bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 12  Sponsorship Menu
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "Sponsorship Menu",
             subtitle="スポンサープラン — 参加形式と目安費用", title_size=30)

plans = [
    ("Venue\nSponsor",    ACCENT_B,
     "・会場提供\n・ドリンク・軽食提供\n・ロゴ掲載\n・公式SNS紹介\n・当日写真・レポート共有",
     "会場提供\nまたは飲食提供"),
    ("Product\nSampling", ACCENT_P,
     "・商品サンプリング\n・体験ブース設置\n・参加者への商品配布\n・公式SNS紹介",
     "商品提供\n＋協賛費\n5〜10万円"),
    ("Main\nSponsor",     ACCENT_PK,
     "・イベントメイン協賛\n・会場内ブランド露出\n・1分ブランド紹介\n・公式レポート投稿\n・クリエイター交流機会",
     "15万〜\n30万円"),
    ("Title\nSponsor",    GOLD,
     "・イベント名への冠掲載\n・トップスポンサー掲載\n・企画設計から共同実施\n・クリエイター体験導線設計\n・インフルエンサー施策相談",
     "30万〜\n50万円以上"),
]
for i, (name, col, items, price) in enumerate(plans):
    bx = Inches(0.42 + i * 3.25)
    # card
    add_rect(sl, bx, Inches(1.6), Inches(3.1), Inches(5.35), fill=WHITE, line_color=col, line_pt=1.5)
    add_rect(sl, bx, Inches(1.6), Inches(3.1), Inches(0.65), fill=col)
    add_text(sl, name, bx + Inches(0.12), Inches(1.62), Inches(2.9), Inches(0.6),
             size=14, bold=True, color=WHITE)
    add_text(sl, items, bx + Inches(0.12), Inches(2.35), Inches(2.9), Inches(3.0),
             size=12, color=DARK_GRAY)
    # price badge
    add_rect(sl, bx, Inches(6.48), Inches(3.1), Inches(0.68), fill=col)
    add_text(sl, price, bx, Inches(6.5), Inches(3.1), Inches(0.62),
             size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

add_text(sl,
         "※料金は提案用の目安です。確定料金は開催規模・会場・提供内容により個別見積もりとなります。",
         Inches(0.55), Inches(7.12), Inches(12.5), Inches(0.3),
         size=9, color=MID_GRAY, italic=True)
bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 13  UGC
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "自然なUGCを生み出す設計",
             subtitle="投稿・露出について — SNS Exposure Design", title_size=28)

ugc = [
    (ACCENT_B,  "📸", "撮影しやすい導線",
     "ブランド体験・フォトスポット・商品展示を設計し、自然に写真・動画が生まれる環境を用意します。"),
    (ACCENT_P,  "🎉", "体験コンテンツ",
     "商品サンプリング・試食・交流企画を通じて、クリエイターが「投稿したくなる」体験を提供します。"),
    (ACCENT_PK, "📢", "公式アカウント発信",
     "イベントレポートを公式SNSアカウントで投稿。スポンサー企業・会場・提供商品を紹介します。"),
    (ACCENT_TEAL,"🔗", "任意投稿を促進",
     "参加クリエイターへの自然な体験提供により、任意投稿・UGCの創出を促進する導線を設計します。"),
    (MID_GRAY,  "⭐", "投稿保証オプション",
     "投稿数・リーチを保証するキャスティング施策は、別途有料オプションとして設計可能です。"),
]
for i, (col, icon, title, desc) in enumerate(ugc):
    by = Inches(1.65 + i * 1.07)
    add_rect(sl, Inches(0.55), by, Inches(12.25), Inches(0.95), fill=OFF_WHITE, line_color=LIGHT_GRAY)
    add_rect(sl, Inches(0.55), by, Inches(0.06), Inches(0.95), fill=col)
    add_text(sl, icon,  Inches(0.72), by + Inches(0.1),  Inches(0.6),  Inches(0.7), size=20)
    add_text(sl, title, Inches(1.45), by + Inches(0.1),  Inches(3.8),  Inches(0.45),
             size=14, bold=True, color=NAVY)
    add_text(sl, desc,  Inches(5.4),  by + Inches(0.1),  Inches(7.25), Inches(0.75),
             size=12.5, color=DARK_GRAY)

add_text(sl,
         "※参加者の個別投稿は任意であり、投稿本数・再生数を保証するものではありません。投稿保証型PRは別途有料施策としてご提案可能です。",
         Inches(0.55), Inches(7.1), Inches(12.5), Inches(0.35),
         size=9, color=MID_GRAY, italic=True)
bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 14  Event Flow
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "イベント当日の流れ",
             subtitle="Event Flow", title_size=30)

steps = [
    ("01", "受付・チェックイン",         "名札配布・参加確認",              ACCENT_B),
    ("02", "ウェルカムドリンク",          "スポンサー提供ドリンクで乾杯",    ACCENT_P),
    ("03", "主催挨拶・スポンサー紹介",    "趣旨説明・スポンサー紹介",        ACCENT_PK),
    ("04", "クリエイター自己紹介",        "全員で簡単な自己紹介・SNS交換",   ACCENT_TEAL),
    ("05", "商品・施設体験",             "サンプリング・体験コーナー",       ACCENT_B),
    ("06", "ミニゲーム・ネットワーキング","交流企画・フォトタイム",           ACCENT_P),
    ("07", "集合写真",                   "公式レポート用撮影",               ACCENT_PK),
    ("08", "アフターレポート配信",        "公式SNSでイベントレポート投稿",    ACCENT_TEAL),
]
for i, (num, step, note, col) in enumerate(steps):
    bx = Inches(0.42 + (i % 4) * 3.25)
    by = Inches(1.6 + (i // 4) * 2.85)
    add_rect(sl, bx, by, Inches(3.1), Inches(2.65), fill=WHITE, line_color=LIGHT_GRAY)
    add_rect(sl, bx, by, Inches(3.1), Inches(0.45), fill=col)
    add_text(sl, num,  bx + Inches(0.12), by + Inches(0.05), Inches(2.85), Inches(0.38),
             size=16, bold=True, color=WHITE)
    add_text(sl, step, bx + Inches(0.12), by + Inches(0.55), Inches(2.85), Inches(0.75),
             size=14, bold=True, color=NAVY)
    add_text(sl, note, bx + Inches(0.12), by + Inches(1.35), Inches(2.85), Inches(1.2),
             size=12, color=DARK_GRAY)

bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 15  Future Vision
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "関西発のグローバルクリエイターコミュニティへ",
             subtitle="今後の展開 — Future Vision", title_size=25)

visions = [
    ("🔄", "定期開催",
     "年複数回の定期開催でコミュニティを継続成長させる", ACCENT_B),
    ("🤝", "コラボイベント",
     "スポンサー企業との共同イベント・コラボ企画を実施", ACCENT_P),
    ("✈️", "FAMツアー",
     "地域・観光施設とのファムトリップでインバウンドPRを強化", ACCENT_PK),
    ("🏙", "インバウンドPR支援",
     "飲食店・ホテル・商業施設の訪日客向けPRをクリエイター起点で支援", ACCENT_TEAL),
    ("🌐", "ネットワーク拡大",
     "関西から日本全国・海外へ発信するグローバルネットワークへ", ACCENT_B),
]
for i, (icon, title, desc, col) in enumerate(visions):
    bx = Inches(0.42 + (i % 3) * 4.28) if i < 3 else Inches(0.42 + (i - 3) * 6.45 + 1.05)
    by = Inches(1.65) if i < 3 else Inches(4.4)
    bw = Inches(4.08) if i < 3 else Inches(5.35)
    add_rect(sl, bx, by, bw, Inches(2.5), fill=WHITE, line_color=col, line_pt=1.5)
    add_rect(sl, bx, by, bw, Inches(0.45), fill=col)
    add_text(sl, icon + "  " + title, bx + Inches(0.12), by + Inches(0.05),
             bw - Inches(0.25), Inches(0.38), size=13, bold=True, color=WHITE)
    add_text(sl, desc, bx + Inches(0.12), by + Inches(0.65),
             bw - Inches(0.25), Inches(1.75), size=13, color=DARK_GRAY)

bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 16  Sponsor Targets
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)
slide_header(sl, "想定スポンサー領域",
             subtitle="Potential Sponsors — ご提案先カテゴリ（候補として記載）", title_size=28)

targets = [
    ("🏨", "ホテル・宿泊",
     "W Osaka、Moxy など\n関西の個性的なホテル・ブティックホテル",      ACCENT_B),
    ("🍺", "飲料・酒類",
     "日本盛、アサヒビール、伊藤園 など\n乾杯・サンプリングスポンサー", ACCENT_P),
    ("🍫", "食品・お菓子",
     "ブラックサンダー、森永 など\nサンプリング・体験展示",              ACCENT_PK),
    ("⚽", "スポーツ",
     "セレッソ大阪 など\n海外ファン向け認知拡大・体験イベント",          ACCENT_TEAL),
    ("🗾", "観光・自治体",
     "観光施設・地域自治体\nインバウンドPR・ローカル体験発信",           ACCENT_B),
    ("🏬", "商業施設",
     "商業施設・ショッピングモール\nインバウンド消費促進・導線設計",      ACCENT_P),
]
for i, (icon, cat, desc, col) in enumerate(targets):
    bx = Inches(0.42 + (i % 3) * 4.28)
    by = Inches(1.6 + (i // 3) * 2.55)
    add_rect(sl, bx, by, Inches(4.08), Inches(2.38), fill=WHITE, line_color=LIGHT_GRAY)
    add_rect(sl, bx, by, Inches(4.08), Inches(0.42), fill=col)
    add_text(sl, icon + "  " + cat, bx + Inches(0.12), by + Inches(0.05),
             Inches(3.85), Inches(0.35), size=13, bold=True, color=WHITE)
    add_text(sl, desc, bx + Inches(0.12), by + Inches(0.58),
             Inches(3.85), Inches(1.7), size=12.5, color=DARK_GRAY)

add_text(sl, "※記載はすべて候補・想定として表記しています。営業先に応じてスライドを差し替え可能です。",
         Inches(0.55), Inches(7.1), Inches(12.5), Inches(0.35),
         size=9, color=MID_GRAY, italic=True)
bottom_bar(sl)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 17  Closing
# ─────────────────────────────────────────────────────────────────────────────
sl = prs.slides.add_slide(BLANK)
white_bg(sl)

# Left navy panel
add_rect(sl, 0, 0, Inches(5.5), H, fill=NAVY)
add_rect(sl, Inches(5.5), 0, Inches(0.06), H, fill=ACCENT_B)

# Left content
add_text(sl, "CONTACT",
         Inches(0.55), Inches(0.9), Inches(4.8), Inches(0.5),
         size=11, bold=True, color=ACCENT_B)
add_rect(sl, Inches(0.55), Inches(1.38), Inches(4.6), Pt(1.5), fill=ACCENT_B)
contacts = [
    ("Company",   "株式会社 Mr.sasuke"),
    ("Contact",   "西村 雄矢"),
    ("Instagram", "@mr.sasuke_japan"),
    ("Email",     "（担当者よりご連絡いたします）"),
]
for i, (k, v) in enumerate(contacts):
    by = Inches(1.6 + i * 0.95)
    add_text(sl, k, Inches(0.55), by, Inches(1.4), Inches(0.45),
             size=11, color=ACCENT_B, bold=True)
    add_text(sl, v, Inches(0.55), by + Inches(0.38), Inches(4.85), Inches(0.52),
             size=14, color=WHITE)

add_text(sl, "Presented by  Mr.sasuke inc.",
         Inches(0.55), Inches(6.85), Inches(4.8), Inches(0.45),
         size=11, color=MID_GRAY)

# Right content
add_text(sl, "関西の魅力を、\n世界に届く体験へ。",
         Inches(5.85), Inches(1.2), Inches(7.2), Inches(2.5),
         size=36, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
add_rect(sl, Inches(5.85), Inches(3.6), Inches(5.5), Pt(2), fill=ACCENT_B)
add_text(sl,
         "スポンサー企業様とともに、グローバルクリエイターが集まる\n"
         "場をつくりたいと思います。商品・施設・地域の魅力を、\n"
         "リアルな体験として届けましょう。\n\n"
         "会場提供・商品提供・協賛など、形は柔軟にご相談いただけます。",
         Inches(5.85), Inches(3.8), Inches(7.2), Inches(2.4),
         size=14, color=DARK_GRAY)

bottom_bar(sl, NAVY)

# ─────────────────────────────────────────────────────────────────────────────
# Save
# ─────────────────────────────────────────────────────────────────────────────
out = "/home/user/-/KansaiGlobalCreatorsMeetup_SponsorshipProposal.pptx"
prs.save(out)
print(f"Saved → {out}")
print(f"Slides: {len(prs.slides)}")
