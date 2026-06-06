from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pptx.oxml.ns as nsmap
from lxml import etree

# Color palette
BLACK  = RGBColor(0x0D, 0x0D, 0x0D)
YELLOW = RGBColor(0xFF, 0xD1, 0x00)
GREEN  = RGBColor(0x39, 0xB5, 0x4A)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
DARK_GRAY = RGBColor(0x1A, 0x1A, 0x1A)
MID_GRAY  = RGBColor(0x2A, 0x2A, 0x2A)
LIGHT_GRAY = RGBColor(0x3A, 0x3A, 0x3A)
ACCENT_GREEN = RGBColor(0x00, 0xE6, 0x76)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]  # completely blank

def add_rect(slide, l, t, w, h, fill_color=None, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, l, t, w, h, font_size=18, bold=False, color=WHITE,
             align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txBox = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txBox.word_wrap = wrap
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.italic = italic
    run.font.name = "Arial"
    return txBox

def add_multiline(slide, lines, l, t, w, h, font_size=14, bold=False, color=WHITE,
                  align=PP_ALIGN.LEFT, line_spacing=1.0):
    txBox = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txBox.word_wrap = True
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, (text, fsize, fbold, fcolor) in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(fsize)
        run.font.bold = fbold
        run.font.color.rgb = fcolor
        run.font.name = "Arial"
    return txBox

def slide_bg(slide, color=BLACK):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_accent_bar(slide, color=YELLOW, height=0.06):
    add_rect(slide, 0, 0, 13.33, height, fill_color=color)

def stat_box(slide, l, t, w, h, number, label, note=""):
    add_rect(slide, l, t, w, h, fill_color=MID_GRAY)
    add_rect(slide, l, t, w, 0.05, fill_color=YELLOW)
    add_text(slide, number, l+0.1, t+0.15, w-0.2, h*0.45,
             font_size=30, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
    add_text(slide, label, l+0.1, t+h*0.48, w-0.2, h*0.3,
             font_size=11, bold=False, color=WHITE, align=PP_ALIGN.CENTER)
    if note:
        add_text(slide, note, l+0.1, t+h*0.78, w-0.2, h*0.18,
                 font_size=7, bold=False, color=RGBColor(0xAA,0xAA,0xAA), align=PP_ALIGN.CENTER)

def card(slide, l, t, w, h, title, body_lines, title_color=YELLOW):
    add_rect(slide, l, t, w, h, fill_color=MID_GRAY)
    add_rect(slide, l, t, w, 0.04, fill_color=title_color)
    add_text(slide, title, l+0.12, t+0.08, w-0.24, 0.35,
             font_size=13, bold=True, color=title_color, align=PP_ALIGN.LEFT)
    y = t + 0.45
    for line in body_lines:
        add_text(slide, "▸  " + line, l+0.12, y, w-0.24, 0.28,
                 font_size=10, color=WHITE)
        y += 0.28


# ─────────────────────────────────────────────
# SLIDE 1 — Cover
# ─────────────────────────────────────────────
s1 = prs.slides.add_slide(BLANK)
slide_bg(s1, BLACK)

# Full yellow diagonal stripe (decorative)
add_rect(s1, 8.5, -0.5, 5.5, 9.0, fill_color=YELLOW)
add_rect(s1, 9.2, -0.5, 4.6, 9.0, fill_color=BLACK)

# Green accent left bar
add_rect(s1, 0, 0, 0.18, 7.5, fill_color=GREEN)

# Title block
add_text(s1, "BIG GROOVE", 0.4, 1.0, 8.5, 1.5,
         font_size=72, bold=True, color=YELLOW, align=PP_ALIGN.LEFT)
add_text(s1, "JAPAN PR OPPORTUNITY", 0.4, 2.55, 9.5, 0.9,
         font_size=32, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
add_text(s1, "Viral Food & Entertainment Activation for Japan",
         0.4, 3.45, 9.0, 0.5,
         font_size=16, bold=False, color=GREEN, align=PP_ALIGN.LEFT)

# Divider
add_rect(s1, 0.4, 4.1, 6.0, 0.04, fill_color=YELLOW)

# Sub info
add_text(s1, "Clive Ibizugbe  |  Boston · Nigeria · Worldwide",
         0.4, 4.25, 9.0, 0.4,
         font_size=13, color=RGBColor(0xCC,0xCC,0xCC))
add_text(s1, "Dancer  ·  Foodie  ·  Fitness  ·  Viral Entertainer",
         0.4, 4.65, 9.0, 0.4,
         font_size=13, color=RGBColor(0xCC,0xCC,0xCC))

# Stats teaser
for i, (num, lbl) in enumerate([("6B+","Global Views"), ("4.6M","TikTok Followers"), ("2,000+","Restaurant Activations")]):
    xl = 0.4 + i * 2.6
    add_rect(s1, xl, 5.45, 2.3, 1.5, fill_color=DARK_GRAY)
    add_rect(s1, xl, 5.45, 2.3, 0.04, fill_color=YELLOW)
    add_text(s1, num, xl+0.1, 5.55, 2.1, 0.65,
             font_size=28, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
    add_text(s1, lbl, xl+0.1, 6.2, 2.1, 0.4,
             font_size=10, color=WHITE, align=PP_ALIGN.CENTER)

add_text(s1, "Publicly available data / To be verified with latest insights",
         0.4, 7.15, 12.5, 0.3,
         font_size=7, color=RGBColor(0x77,0x77,0x77))


# ─────────────────────────────────────────────
# SLIDE 2 — Executive Summary
# ─────────────────────────────────────────────
s2 = prs.slides.add_slide(BLANK)
slide_bg(s2, BLACK)
add_accent_bar(s2, YELLOW)

add_text(s2, "EXECUTIVE SUMMARY", 0.4, 0.2, 12.5, 0.5,
         font_size=11, bold=True, color=YELLOW)
add_text(s2, "Big Grooveとは何者か？", 0.4, 0.75, 12.5, 0.7,
         font_size=38, bold=True, color=WHITE)
add_rect(s2, 0.4, 1.5, 5.0, 0.04, fill_color=GREEN)

# Main quote box
add_rect(s2, 0.4, 1.65, 12.5, 1.4, fill_color=DARK_GRAY)
add_rect(s2, 0.4, 1.65, 0.12, 1.4, fill_color=YELLOW)
add_text(s2,
    'Big Grooveは、飲食体験を「記憶に残るSNSコンテンツ」へ変えるグローバルエンターテイナーです。',
    0.7, 1.75, 12.0, 0.6,
    font_size=20, bold=True, color=WHITE)
add_text(s2,
    "ダンス・フード・フィットネスを掛け合わせた独自スタイルで、レストランやブランドへの来店・認知をバイラルに爆発させます。",
    0.7, 2.4, 12.0, 0.5,
    font_size=14, color=RGBColor(0xCC,0xCC,0xCC))

# 3 pillars
pillars = [
    ("🌍 Global Reach", "6B+ビューを誇るグローバルなリーチ。\n40カ国以上でコンテンツが拡散。"),
    ("🍜 Food Activation Expert", "2,000以上のレストラン訪問実績。\n飲食店との親和性が圧倒的に高い。"),
    ("📱 Viral Creation Machine", "TikTok・Instagram・Facebookで\nリアルタイムに話題を爆発させる。"),
]
for i, (title, body) in enumerate(pillars):
    xl = 0.4 + i * 4.3
    add_rect(s2, xl, 3.3, 4.0, 2.8, fill_color=MID_GRAY)
    add_rect(s2, xl, 3.3, 4.0, 0.06, fill_color=GREEN)
    add_text(s2, title, xl+0.15, 3.42, 3.7, 0.45,
             font_size=14, bold=True, color=YELLOW)
    add_text(s2, body, xl+0.15, 3.9, 3.7, 1.9,
             font_size=12, color=WHITE)

add_text(s2, "Publicly available data / To be verified",
         0.4, 7.2, 12.5, 0.25, font_size=7,
         color=RGBColor(0x66,0x66,0x66))


# ─────────────────────────────────────────────
# SLIDE 3 — Creator Profile
# ─────────────────────────────────────────────
s3 = prs.slides.add_slide(BLANK)
slide_bg(s3, BLACK)
add_accent_bar(s3, GREEN)

add_text(s3, "CREATOR PROFILE", 0.4, 0.2, 12.5, 0.4,
         font_size=11, bold=True, color=GREEN)
add_text(s3, "Clive Ibizugbe / Big Groove", 0.4, 0.65, 12.5, 0.65,
         font_size=38, bold=True, color=WHITE)
add_rect(s3, 0.4, 1.35, 4.5, 0.04, fill_color=YELLOW)

# Left column — profile details
left_items = [
    ("🎤  Creator Name", "Big Groove (Clive Ibizugbe)"),
    ("📍  Base", "Boston, MA / Nigeria / Worldwide"),
    ("🎯  Genre", "Dancer · Foodie · Fitness · Viral Entertainer"),
    ("🌐  Reach", "40+ Countries  |  Featured Worldwide"),
    ("📧  Booking", "booking@movewitgroove.com"),
]
y = 1.55
for label, val in left_items:
    add_text(s3, label, 0.4, y, 4.5, 0.28, font_size=9,
             color=RGBColor(0x99,0x99,0x99))
    add_text(s3, val,   0.4, y+0.27, 4.5, 0.32, font_size=12,
             bold=True, color=WHITE)
    y += 0.65

# Right column — world view description
add_rect(s3, 5.3, 1.55, 7.6, 5.4, fill_color=DARK_GRAY)
add_rect(s3, 5.3, 1.55, 0.1, 5.4, fill_color=YELLOW)
add_text(s3, "Big Grooveの世界観", 5.55, 1.7, 7.0, 0.4,
         font_size=15, bold=True, color=YELLOW)
desc_lines = [
    "Big Grooveは、単なるダンサーではありません。",
    "",
    "レストランに入り、料理に感動し、そのエネルギーをダンスとユーモアで爆発させる。そのスタイルが世界中で共感を呼び、数十億回の視聴を記録しています。",
    "",
    "飲食・フィットネス・エンターテイメントという異なるジャンルを自然に融合させた彼の存在は、食体験を「シェアしたくなる瞬間」へと昇華させます。",
    "",
    "Snoop Dogg、Deion Sanders、Bobby Shmurda、Feastablesなど、グローバルブランド・セレブリティとの接点も多数。",
    "",
    "日本の飲食シーンに彼を呼び込むことで、インバウンド層へのアピールと国内SNS拡散の両方を同時に実現できます。",
]
y2 = 2.2
for line in desc_lines:
    if line == "":
        y2 += 0.12
        continue
    add_text(s3, line, 5.55, y2, 7.1, 0.35, font_size=11, color=WHITE)
    y2 += 0.38

add_text(s3, "Publicly available data / To be verified",
         0.4, 7.2, 12.5, 0.25, font_size=7,
         color=RGBColor(0x66,0x66,0x66))


# ─────────────────────────────────────────────
# SLIDE 4 — Social Media Insights
# ─────────────────────────────────────────────
s4 = prs.slides.add_slide(BLANK)
slide_bg(s4, BLACK)
add_accent_bar(s4, YELLOW)

add_text(s4, "SOCIAL MEDIA INSIGHTS", 0.4, 0.2, 12.5, 0.4,
         font_size=11, bold=True, color=YELLOW)
add_text(s4, "数字が証明するグローバルな影響力", 0.4, 0.65, 12.5, 0.65,
         font_size=34, bold=True, color=WHITE)
add_rect(s4, 0.4, 1.35, 5.0, 0.04, fill_color=GREEN)

# Big stats row
stats_top = [
    ("6B+",      "Global Total Views",    "Publicly available"),
    ("100M+",    "Views (Official Site)",  "Publicly available"),
    ("40+",      "Countries Reached",      "Publicly available"),
    ("2,000+",   "Restaurant Activations", "Publicly available"),
]
for i, (num, lbl, note) in enumerate(stats_top):
    stat_box(s4, 0.3 + i*3.2, 1.6, 3.0, 1.8, num, lbl, note)

# Platform breakdown
platforms = [
    ("TikTok",     "4,600,000",  "Followers",   "Favikon Nigeria TikTok #18"),
    ("Instagram",  "~2,000,000", "Followers",   "~1,600 posts"),
    ("Facebook",   "2,009,903",  "Page Likes",  "273,673 talking about"),
]
for i, (platform, followers, label, sub) in enumerate(platforms):
    xl = 0.4 + i*4.3
    add_rect(s4, xl, 3.75, 4.0, 2.5, fill_color=MID_GRAY)
    colors_p = [GREEN, RGBColor(0xE1,0x30,0x6C), RGBColor(0x18,0x77,0xF2)]
    add_rect(s4, xl, 3.75, 4.0, 0.06, fill_color=colors_p[i])
    add_text(s4, platform, xl+0.15, 3.85, 3.7, 0.4,
             font_size=16, bold=True, color=colors_p[i])
    add_text(s4, followers, xl+0.15, 4.3, 3.7, 0.65,
             font_size=32, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
    add_text(s4, label, xl+0.15, 4.95, 3.7, 0.3,
             font_size=11, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s4, sub, xl+0.15, 5.28, 3.7, 0.35,
             font_size=9, color=RGBColor(0xAA,0xAA,0xAA), align=PP_ALIGN.CENTER)
    add_text(s4, "* To be verified", xl+0.15, 5.65, 3.7, 0.3,
             font_size=7, color=RGBColor(0x77,0x77,0x77), align=PP_ALIGN.CENTER)

add_text(s4, "Publicly available data / To be verified with latest screenshots from creator/manager",
         0.4, 7.2, 12.5, 0.25, font_size=7,
         color=RGBColor(0x66,0x66,0x66))


# ─────────────────────────────────────────────
# SLIDE 5 — Why Big Groove Works
# ─────────────────────────────────────────────
s5 = prs.slides.add_slide(BLANK)
slide_bg(s5, BLACK)
add_accent_bar(s5, GREEN)

add_text(s5, "WHY BIG GROOVE WORKS", 0.4, 0.2, 12.5, 0.4,
         font_size=11, bold=True, color=GREEN)
add_text(s5, "なぜ今、Big Grooveなのか？", 0.4, 0.65, 12.5, 0.65,
         font_size=34, bold=True, color=WHITE)
add_rect(s5, 0.4, 1.35, 4.5, 0.04, fill_color=YELLOW)

reasons = [
    ("⚡  Viral Attention",
     ["リアクション動画・サプライズ演出でSNSバズを量産",
      "視聴者が思わず「シェアしたくなる」瞬間を設計",
      "飲食コンテンツとの親和性が極めて高い"]),
    ("🤝  Real-World Engagement",
     ["現場に入り込み、スタッフ・客とリアルに交流",
      "演技・台本なしの自然なエネルギーで信頼感を醸成",
      "来店者のUGC（二次投稿）を自然に誘発"]),
    ("🍽️  Restaurant Activation",
     ["2,000以上の飲食店で検証済みのコンテンツ手法",
      "料理の魅力を最大限に引き出す表現力",
      "高級店から大衆店まで幅広くマッチ"]),
    ("🌏  Cultural Impact",
     ["ナイジェリア・ボストン・世界を股にかけるバックグラウンド",
      "インバウンド観光客にも強くアピール",
      "グローバルセレブとのコネクションで国際的権威性"]),
]
for i, (title, points) in enumerate(reasons):
    col = i % 2
    row = i // 2
    xl = 0.4 + col * 6.4
    yt = 1.6 + row * 2.9
    add_rect(s5, xl, yt, 6.1, 2.6, fill_color=MID_GRAY)
    add_rect(s5, xl, yt, 6.1, 0.06, fill_color=YELLOW if row==0 else GREEN)
    add_text(s5, title, xl+0.15, yt+0.1, 5.8, 0.45,
             font_size=15, bold=True, color=YELLOW)
    yp = yt + 0.6
    for pt in points:
        add_text(s5, "▸  " + pt, xl+0.15, yp, 5.8, 0.38, font_size=11, color=WHITE)
        yp += 0.4


# ─────────────────────────────────────────────
# SLIDE 6 — Japan Market Fit
# ─────────────────────────────────────────────
s6 = prs.slides.add_slide(BLANK)
slide_bg(s6, BLACK)
add_accent_bar(s6, YELLOW)

add_text(s6, "JAPAN MARKET FIT", 0.4, 0.2, 12.5, 0.4,
         font_size=11, bold=True, color=YELLOW)
add_text(s6, "日本市場との相性", 0.4, 0.65, 12.5, 0.65,
         font_size=34, bold=True, color=WHITE)
add_rect(s6, 0.4, 1.35, 5.0, 0.04, fill_color=GREEN)

add_text(s6,
    "日本の飲食・観光シーンは、世界で最も「映える」コンテンツを生み出すポテンシャルを持っています。\nBig Grooveはその魅力を世界に届ける最強のアンプリファイアーです。",
    0.4, 1.5, 12.5, 0.8, font_size=13, color=RGBColor(0xCC,0xCC,0xCC))

segments = [
    ("🥩  焼肉・BBQ", "煙と炎・肉の迫力はバイラル動画と相性抜群"),
    ("🍣  寿司・高級和食", "繊細な技・美しさをドラマチックに演出"),
    ("🍜  ラーメン・つけ麺", "庶民的で親しみやすく、海外人気も絶大"),
    ("🍺  居酒屋・バー", "群衆のエネルギーをそのまま映像化できる"),
    ("🏯  観光施設・テーマパーク", "日本文化＋エンタメの融合でインバウンド訴求"),
    ("🛍️  食品ブランド・EC", "ブランドコラボ・PR Reelで商品認知を加速"),
]
for i, (title, desc) in enumerate(segments):
    col = i % 3
    row = i // 3
    xl = 0.4 + col * 4.3
    yt = 2.55 + row * 2.15
    add_rect(s6, xl, yt, 4.0, 1.9, fill_color=DARK_GRAY)
    add_rect(s6, xl, yt, 4.0, 0.05, fill_color=YELLOW if col==0 else (GREEN if col==1 else RGBColor(0xFF,0x80,0x00)))
    add_text(s6, title, xl+0.12, yt+0.1, 3.76, 0.4,
             font_size=14, bold=True, color=YELLOW)
    add_text(s6, desc, xl+0.12, yt+0.55, 3.76, 1.1,
             font_size=11, color=WHITE)

add_text(s6, "Publicly available data / To be verified",
         0.4, 7.2, 12.5, 0.25, font_size=7,
         color=RGBColor(0x66,0x66,0x66))


# ─────────────────────────────────────────────
# SLIDE 7 — Campaign Ideas
# ─────────────────────────────────────────────
s7 = prs.slides.add_slide(BLANK)
slide_bg(s7, BLACK)
add_accent_bar(s7, GREEN)

add_text(s7, "CAMPAIGN IDEAS", 0.4, 0.2, 12.5, 0.4,
         font_size=11, bold=True, color=GREEN)
add_text(s7, "提案キャンペーン例", 0.4, 0.65, 12.5, 0.65,
         font_size=34, bold=True, color=WHITE)
add_rect(s7, 0.4, 1.35, 4.5, 0.04, fill_color=YELLOW)

campaigns = [
    ("01", "Restaurant Viral Visit",
     ["サプライズ形式でレストランを訪問",
      "料理・スタッフ・雰囲気をバイラル演出",
      "TikTok / Instagram Reelで公開 → 来店動機を爆発的に創出"]),
    ("02", "Japan Food Tour Series",
     ["東京・大阪・京都など複数都市を巡る動画シリーズ",
      "各都市のB級グルメ〜高級店を対比で魅せる",
      "インバウンド層にも強力にアピール"]),
    ("03", "Brand Collaboration Reel",
     ["特定ブランドの商品・店舗をフィーチャー",
      "Big Grooveのダンス・リアクションで商品価値を増幅",
      "ブランド公式アカウントでの二次利用権も確保"]),
    ("04", "Live Event Appearance",
     ["グランドオープン・周年イベント・フェスへの出演",
      "現場のエネルギーをリアルタイムで発信",
      "集客効果＋SNS拡散の相乗効果を実現"]),
]
for i, (num, title, points) in enumerate(campaigns):
    col = i % 2
    row = i // 2
    xl = 0.4 + col * 6.4
    yt = 1.6 + row * 2.8
    add_rect(s7, xl, yt, 6.1, 2.55, fill_color=MID_GRAY)
    add_rect(s7, xl, yt, 0.55, 2.55, fill_color=YELLOW if row==0 else GREEN)
    add_text(s7, num, xl+0.05, yt+0.8, 0.45, 0.8,
             font_size=24, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
    add_text(s7, title, xl+0.7, yt+0.08, 5.2, 0.45,
             font_size=15, bold=True, color=YELLOW)
    yp = yt + 0.58
    for pt in points:
        add_text(s7, "▸  " + pt, xl+0.7, yp, 5.2, 0.38, font_size=11, color=WHITE)
        yp += 0.42


# ─────────────────────────────────────────────
# SLIDE 8 — Deliverables
# ─────────────────────────────────────────────
s8 = prs.slides.add_slide(BLANK)
slide_bg(s8, BLACK)
add_accent_bar(s8, YELLOW)

add_text(s8, "DELIVERABLES", 0.4, 0.2, 12.5, 0.4,
         font_size=11, bold=True, color=YELLOW)
add_text(s8, "提供コンテンツ・成果物", 0.4, 0.65, 12.5, 0.65,
         font_size=34, bold=True, color=WHITE)
add_rect(s8, 0.4, 1.35, 5.5, 0.04, fill_color=GREEN)

deliverables = [
    ("📱", "Instagram Reel",    "縦型ショート動画（15〜60秒）\nブランドタグ付き投稿",      GREEN),
    ("🎵", "TikTok Video",      "トレンドサウンド活用\nハッシュタグ戦略込み",            YELLOW),
    ("▶️", "YouTube Shorts",    "アーカイブ残存による\n長期的な検索流入効果",            RGBColor(0xFF,0x00,0x00)),
    ("👍", "Facebook Post",     "2M+ページへの投稿\n高エンゲージメント層にリーチ",       RGBColor(0x18,0x77,0xF2)),
    ("🎤", "Live Appearance",   "現地イベント出演\nリアルタイム熱狂を創出",              RGBColor(0xFF,0x80,0x00)),
    ("📂", "二次利用素材",        "クライアント側での広告・SNS再利用\nライセンス条件は要協議", RGBColor(0xCC,0x00,0xCC)),
]
for i, (icon, title, desc, col) in enumerate(deliverables):
    xi = i % 3
    yi = i // 3
    xl = 0.4 + xi * 4.3
    yt = 1.6 + yi * 2.65
    add_rect(s8, xl, yt, 4.0, 2.4, fill_color=DARK_GRAY)
    add_rect(s8, xl, yt, 4.0, 0.06, fill_color=col)
    add_text(s8, icon + "  " + title, xl+0.12, yt+0.12, 3.76, 0.45,
             font_size=14, bold=True, color=col)
    add_text(s8, desc, xl+0.12, yt+0.65, 3.76, 1.5,
             font_size=12, color=WHITE)


# ─────────────────────────────────────────────
# SLIDE 9 — Expected Value
# ─────────────────────────────────────────────
s9 = prs.slides.add_slide(BLANK)
slide_bg(s9, BLACK)
add_accent_bar(s9, GREEN)

add_text(s9, "EXPECTED VALUE", 0.4, 0.2, 12.5, 0.4,
         font_size=11, bold=True, color=GREEN)
add_text(s9, "期待される効果・価値", 0.4, 0.65, 12.5, 0.65,
         font_size=34, bold=True, color=WHITE)
add_rect(s9, 0.4, 1.35, 5.0, 0.04, fill_color=YELLOW)

values = [
    ("🚀", "認知拡大",      "6B+ビューのリーチ基盤を活用し、\nブランド/店舗の認知を国内外に広げる"),
    ("🌏", "海外向け話題化",  "インバウンド層・海外ファンへの\n日本グルメ情報発信を加速"),
    ("📍", "来店動機の創出",  "「Big Grooveが行った店」として\nSNSで話題化し来店意欲を刺激"),
    ("📸", "UGC誘発",      "視聴者・来店客が自発的に再投稿する\nユーザー生成コンテンツの連鎖"),
    ("🌟", "ブランドの国際感", "グローバルインフルエンサーとの協業で\n企業・店舗のブランド格を向上"),
    ("📊", "データ取得",     "コンテンツのインプレッション・リーチ・\nエンゲージメントをレポートとして提供"),
]
for i, (icon, title, desc) in enumerate(values):
    xi = i % 3
    yi = i // 3
    xl = 0.4 + xi * 4.3
    yt = 1.6 + yi * 2.65
    bg_col = MID_GRAY if yi == 0 else DARK_GRAY
    ac_col = YELLOW if xi == 0 else (GREEN if xi == 1 else RGBColor(0xFF,0x80,0x00))
    add_rect(s9, xl, yt, 4.0, 2.4, fill_color=bg_col)
    add_rect(s9, xl, yt+2.34, 4.0, 0.06, fill_color=ac_col)
    add_text(s9, icon, xl+0.15, yt+0.15, 0.6, 0.55, font_size=24)
    add_text(s9, title, xl+0.75, yt+0.18, 3.1, 0.45,
             font_size=15, bold=True, color=ac_col)
    add_text(s9, desc, xl+0.15, yt+0.75, 3.7, 1.45,
             font_size=11, color=WHITE)

add_text(s9, "Publicly available data / To be verified",
         0.4, 7.2, 12.5, 0.25, font_size=7,
         color=RGBColor(0x66,0x66,0x66))


# ─────────────────────────────────────────────
# SLIDE 10 — Next Steps
# ─────────────────────────────────────────────
s10 = prs.slides.add_slide(BLANK)
slide_bg(s10, BLACK)
add_accent_bar(s10, YELLOW)

add_text(s10, "NEXT STEPS", 0.4, 0.2, 12.5, 0.4,
         font_size=11, bold=True, color=YELLOW)
add_text(s10, "今後のステップ・アクション", 0.4, 0.65, 12.5, 0.65,
         font_size=34, bold=True, color=WHITE)
add_rect(s10, 0.4, 1.35, 5.0, 0.04, fill_color=GREEN)

steps = [
    ("STEP 01", "候補店舗・ブランドの選定",
     ["ターゲット業態・エリアの絞り込み",
      "Big Grooveのコンテンツスタイルとの相性確認",
      "予算感・期待効果の仮設定"]),
    ("STEP 02", "最新インサイト・条件の確認",
     ["creator/マネージャーから最新インサイトのスクリーンショット取得",
      "出演費・移動費・制作費の確認",
      "撮影ルール・監修フロー・投稿タイミングの合意"]),
    ("STEP 03", "コンテンツ設計・日程調整",
     ["撮影シナリオ・演出方針の共有",
      "投稿媒体・本数・フォーマット確定",
      "ロケーション下見・許可申請"]),
    ("STEP 04", "契約・実施・レポーティング",
     ["使用権・二次利用範囲の契約締結",
      "撮影・編集・投稿の実施",
      "インプレッション・リーチ・エンゲージメントのレポート提出"]),
]
for i, (step, title, points) in enumerate(steps):
    col = i % 2
    row = i // 2
    xl = 0.4 + col * 6.4
    yt = 1.6 + row * 2.7
    add_rect(s10, xl, yt, 6.1, 2.45, fill_color=MID_GRAY)
    ac = YELLOW if row == 0 else GREEN
    add_rect(s10, xl, yt, 1.0, 2.45, fill_color=ac)
    add_text(s10, step, xl+0.05, yt+0.8, 0.9, 0.7,
             font_size=12, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
    add_text(s10, title, xl+1.1, yt+0.08, 4.8, 0.45,
             font_size=14, bold=True, color=ac)
    yp = yt + 0.58
    for pt in points:
        add_text(s10, "▸  " + pt, xl+1.1, yp, 4.8, 0.38, font_size=11, color=WHITE)
        yp += 0.42

# CTA box
add_rect(s10, 0.4, 7.0, 12.5, 0.45, fill_color=DARK_GRAY)
add_text(s10,
    "Contact:  booking@movewitgroove.com  |  Data: Publicly available / To be verified with latest insights from creator/manager",
    0.6, 7.05, 12.1, 0.35,
    font_size=9, color=RGBColor(0xAA,0xAA,0xAA), align=PP_ALIGN.CENTER)


# ─────────────────────────────────────────────
# Save
# ─────────────────────────────────────────────
output_path = "/home/user/-/BigGroove_Japan_PR_Proposal.pptx"
prs.save(output_path)
print(f"Saved: {output_path}")
