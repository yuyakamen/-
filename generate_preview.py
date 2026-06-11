from PIL import Image, ImageDraw, ImageFont
import os

W = 800
H = 5200
img = Image.new("RGB", (W, H), "#FFFFFF")
draw = ImageDraw.Draw(img)

def hex2rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

PINK      = hex2rgb("F4748A")
PINK_LT   = hex2rgb("FDE8EC")
PINK_PALE = hex2rgb("FFF5F7")
BEIGE     = hex2rgb("F9F3EE")
BEIGE_DK  = hex2rgb("EFE4D8")
ROSE      = hex2rgb("E8527A")
BROWN     = hex2rgb("8B6B5A")
DARK      = hex2rgb("3A2E2A")
GRAY      = hex2rgb("7A7268")
GOLD      = hex2rgb("C9A96E")
WHITE     = (255, 255, 255)

def try_font(size, bold=False):
    paths = [
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJKjp-Regular.otf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]
    for p in paths:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except:
                pass
    return ImageFont.load_default()

f_sm   = try_font(13)
f_md   = try_font(16)
f_lg   = try_font(22, bold=True)
f_xl   = try_font(30, bold=True)
f_xxl  = try_font(42, bold=True)
f_bold = try_font(18, bold=True)

def rounded_rect(draw, xy, radius, fill, outline=None, outline_width=1):
    x0, y0, x1, y1 = xy
    draw.rounded_rectangle([x0, y0, x1, y1], radius=radius, fill=fill, outline=outline, width=outline_width)

def gradient_rect(img, x0, y0, x1, y1, c1, c2, vertical=True):
    region = img.crop((x0, y0, x1, y1))
    w, h = region.size
    grad = Image.new("RGB", (1, h) if vertical else (w, 1))
    for i in range(h if vertical else w):
        t = i / max(h - 1, 1) if vertical else i / max(w - 1, 1)
        r = int(c1[0] + (c2[0]-c1[0])*t)
        g = int(c1[1] + (c2[1]-c1[1])*t)
        b = int(c1[2] + (c2[2]-c1[2])*t)
        if vertical:
            grad.putpixel((0, i), (r, g, b))
        else:
            grad.putpixel((i, 0), (r, g, b))
    grad = grad.resize((w, h), Image.BILINEAR)
    img.paste(grad, (x0, y0))
    return img

y = 0

# ── HEADER ──────────────────────────────────────────
draw.rectangle([0, y, W, y+60], fill=WHITE)
draw.rectangle([0, y+59, W, y+60], fill=PINK_LT)
draw.text((28, y+18), "OfficeColor", fill=PINK, font=f_bold)
rounded_rect(draw, [W-170, y+12, W-20, y+48], 20, fill=PINK)
draw.text((W-158, y+22), "無料登録はこちら", fill=WHITE, font=f_sm)
y += 60

# ── HERO ────────────────────────────────────────────
gradient_rect(img, 0, y, W, y+520, hex2rgb("FFF0F3"), hex2rgb("F9F3EE"))
draw.rectangle([0, y, W, y+520], fill=None)
# badge
rounded_rect(draw, [28, y+32, 320, y+58], 20, fill=PINK_LT, outline=PINK, outline_width=1)
draw.text((42, y+38), "✨ 女性のための事務職専門エージェント", fill=PINK, font=f_sm)
# title
draw.text((28, y+70),  "「自分らしく働きたい」", fill=DARK, font=f_xl)
draw.text((28, y+110), "その気持ちに", fill=DARK, font=f_xl)
draw.text((28+try_font(30,True).getlength("その気持ちに"), y+110), "寄り添う", fill=PINK, font=f_xl)
draw.text((28, y+150), "事務職転職サポート。", fill=DARK, font=f_xl)
# sub
draw.text((28, y+200), "未経験・ブランクOK。残業少なめ・在宅OK", fill=GRAY, font=f_md)
draw.text((28, y+225), "産育休取得実績あり。専任の女性アドバイザーが", fill=GRAY, font=f_md)
draw.text((28, y+250), "あなたに合った事務職を一緒に探します。", fill=GRAY, font=f_md)
# CTA button
gradient_rect(img, 28, y+290, 360, y+340, PINK, ROSE, vertical=False)
draw.rounded_rectangle([28, y+290, 360, y+340], radius=25, fill=None, outline=None)
img2 = img.copy(); d2 = ImageDraw.Draw(img2)
img2.paste(img.crop((28,y+290,360,y+340)), (28,y+290))
rounded_rect(draw, [28, y+290, 360, y+340], 25, fill=PINK)
draw.text((60, y+304), "今すぐ無料で相談する  →", fill=WHITE, font=f_bold)
draw.text((28, y+352), "登録無料・最短30秒・しつこい営業なし", fill=GRAY, font=f_sm)

# stats cards
for i, (num, label) in enumerate([("98%","利用者満足度"),("3,200件+","事務職専門求人"),("女性特化","専任アドバイザー")]):
    cx = 28 + i*230
    rounded_rect(draw, [cx, y+385, cx+210, y+450], 12, fill=WHITE, outline=PINK_LT, outline_width=2)
    draw.text((cx+20, y+398), num, fill=PINK, font=f_lg)
    draw.text((cx+20, y+428), label, fill=GRAY, font=f_sm)

# hero image placeholder
rounded_rect(draw, [W-310, y+30, W-30, y+490], 24, fill=PINK_LT)
draw.text((W-240, y+230), "👩‍💼", fill=DARK, font=try_font(60))
draw.text((W-230, y+310), "あなたらしい", fill=BROWN, font=f_md)
draw.text((W-230, y+335), "  働き方へ", fill=BROWN, font=f_md)
y += 520

# ── WORRIES ─────────────────────────────────────────
draw.rectangle([0, y, W, y+480], fill=BEIGE)
draw.text((W//2 - 120, y+24), "こんなお悩み、ありませんか？", fill=DARK, font=f_lg)
worries = [
    ("😟","残業が多すぎる…","毎日終電で帰宅。自分の時間が全くない。"),
    ("😔","事務職は未経験でも大丈夫？","資格がないと採用されないかも…"),
    ("😰","ブランクがあって不安…","育児・介護でブランク。今さら転職できるか。"),
    ("🤔","在宅・時短勤務できる会社は？","家庭との両立を考えると柔軟な職場を探したい。"),
    ("😣","転職活動の進め方がわからない","履歴書の書き方、面接対策、何から始めるか。"),
    ("💭","本当に転職していいか迷っている","不安と「もっとよくなりたい」気持ちの葛藤。"),
]
for i, (icon, title, desc) in enumerate(worries):
    row, col = divmod(i, 3)
    cx = 20 + col * 258
    cy = y + 70 + row * 190
    rounded_rect(draw, [cx, cy, cx+245, cy+170], 14, fill=WHITE)
    draw.text((cx+14, cy+14), icon, fill=DARK, font=try_font(28))
    draw.text((cx+14, cy+58), title, fill=PINK, font=try_font(13, True))
    # wrap desc
    words = desc
    draw.text((cx+14, cy+82), words[:24], fill=DARK, font=f_sm)
    draw.text((cx+14, cy+102), words[24:], fill=DARK, font=f_sm)
y += 480

# ── FEATURES ────────────────────────────────────────
draw.rectangle([0, y, W, y+560], fill=PINK_PALE)
draw.text((W//2 - 130, y+24), "OfficeColorが選ばれる理由", fill=DARK, font=f_lg)
features = [
    ("👩‍💼","女性専任アドバイザー","同じ女性として共感しながらサポート。"),
    ("🔍","事務職求人3,200件以上","残業少なめ・在宅・時短求人が豊富。"),
    ("📝","書類〜面接まで完全サポート","内定まで二人三脚でサポート。"),
    ("🏠","在宅・時短・産育休実績あり","ライフスタイルに合った職場をご提案。"),
    ("💚","未経験・ブランクOK","毎年多数の方が内定を獲得中。"),
    ("🆓","完全無料・営業なし","ペースはあなた主導で進められます。"),
]
for i, (icon, title, desc) in enumerate(features):
    row, col = divmod(i, 3)
    cx = 20 + col * 258
    cy = y + 72 + row * 230
    rounded_rect(draw, [cx, cy, cx+245, cy+210], 16, fill=WHITE)
    draw.rectangle([cx, cy, cx+245, cy+5], fill=PINK)  # top bar
    draw.text((cx+95, cy+18), icon, fill=DARK, font=try_font(34))
    draw.text((cx+14, cy+74), title, fill=DARK, font=try_font(14, True))
    draw.text((cx+14, cy+100), desc[:22], fill=GRAY, font=f_sm)
    if len(desc) > 22:
        draw.text((cx+14, cy+118), desc[22:], fill=GRAY, font=f_sm)
y += 560

# ── JOB TYPES ───────────────────────────────────────
draw.rectangle([0, y, W, y+340], fill=BEIGE)
draw.text((W//2 - 100, y+22), "取り扱い事務職の種類", fill=DARK, font=f_lg)
jobs = [("💻","一般事務"),("🏥","医療事務"),("📊","経理事務"),("📞","営業事務"),
        ("🏢","総務・人事"),("💊","調剤薬局事務"),("🏦","金融事務"),("🖥️","貿易・英語事務")]
for i, (icon, name) in enumerate(jobs):
    row, col = divmod(i, 4)
    cx = 18 + col * 193
    cy = y + 68 + row * 135
    rounded_rect(draw, [cx, cy, cx+180, cy+118], 14, fill=WHITE, outline=PINK_LT, outline_width=2)
    draw.text((cx+68, cy+14), icon, fill=DARK, font=try_font(28))
    draw.text((cx+20, cy+58), name, fill=DARK, font=try_font(13, True))
    rounded_rect(draw, [cx+24, cy+82, cx+155, cy+106], 12, fill=PINK_LT)
    draw.text((cx+38, cy+87), "未経験OK", fill=PINK, font=f_sm)
y += 340

# ── CTA MID ──────────────────────────────────────────
gradient_rect(img, 0, y, W, y+200, PINK, ROSE)
draw.text((W//2 - 190, y+30), "まずは気軽に話してみませんか？", fill=WHITE, font=f_xl)
draw.text((W//2 - 220, y+80), "「転職するか迷っている」という段階でも大丈夫。", fill=hex2rgb("FFE0E6"), font=f_md)
draw.text((W//2 - 195, y+105), "女性アドバイザーが丁寧にお聞きします。", fill=hex2rgb("FFE0E6"), font=f_md)
rounded_rect(draw, [W//2-170, y+135, W//2+170, y+180], 25, fill=WHITE)
draw.text((W//2-130, y+148), "✨ 無料で相談する（登録30秒）", fill=PINK, font=f_bold)
y += 200

# ── VOICES ───────────────────────────────────────────
draw.rectangle([0, y, W, y+380], fill=WHITE)
draw.text((W//2 - 145, y+22), "転職に成功した女性たちの声", fill=DARK, font=f_lg)
voices = [
    ("😊","Aさん（29歳）","飲食業→一般事務",
     "事務は未経験でしたが、アドバイザー\nさんが丁寧に教えてくれました。今は\n残業ほぼゼロで、趣味の時間もできて\n本当に転職して良かったです！"),
    ("🌸","Bさん（34歳）","育児後→医療事務（時短）",
     "3年のブランクがあり転職は無理と\n思っていましたが「ブランクは強みに\nなる」と言ってもらえ、今は時短勤務\nで子どもの行事にも参加できています。"),
    ("💼","Cさん（26歳）","販売職→営業事務（在宅）",
     "在宅の事務職を探していましたが自分\nでは探し方がわからなくて。エージェ\nントにお任せしたら希望通りほぼフル\nリモートの求人を提案してもらえました！"),
]
for i, (icon, name, info, text) in enumerate(voices):
    cx = 18 + i * 255
    cy = y + 64
    rounded_rect(draw, [cx, cy, cx+242, cy+295], 18, fill=BEIGE)
    draw.text((cx+14, cy+10), "❝", fill=PINK_LT, font=try_font(48))
    rounded_rect(draw, [cx+14, cy+50, cx+62, cy+98], 24, fill=PINK_LT)
    draw.text((cx+18, cy+52), icon, fill=DARK, font=try_font(28))
    draw.text((cx+72, cy+54), name, fill=DARK, font=try_font(13, True))
    draw.text((cx+72, cy+74), info, fill=GRAY, font=f_sm)
    for j, line in enumerate(text.split("\n")):
        draw.text((cx+14, cy+108 + j*21), line, fill=DARK, font=f_sm)
    draw.text((cx+14, cy+268), "★★★★★", fill=GOLD, font=f_md)
y += 380

# ── FLOW ─────────────────────────────────────────────
draw.rectangle([0, y, W, y+600], fill=PINK_PALE)
draw.text((W//2 - 90, y+22), "転職までの流れ", fill=DARK, font=f_lg)
steps = [
    ("1","📋 無料登録（30秒）","フォームに入力するだけで登録完了。費用は一切かかりません。","約30秒"),
    ("2","👩‍💼 女性アドバイザーとのヒアリング","希望・状況・不安を丁寧にお聞きします。オンライン可。","約30〜60分"),
    ("3","🔍 求人提案・応募","希望に合った事務職求人を厳選してご提案します。","数日〜1週間"),
    ("4","📝 書類・面接サポート","履歴書添削・面接対策を徹底サポート。","1〜2週間"),
    ("5","🎉 内定・入社","給与交渉・入社日の調整もサポートします。","最短2週間で内定も"),
]
for i, (num, title, desc, time_) in enumerate(steps):
    cy = y + 68 + i * 104
    # circle
    draw.ellipse([50, cy+4, 98, cy+52], fill=PINK)
    draw.text((66, cy+14), num, fill=WHITE, font=f_bold)
    # line
    if i < 4:
        draw.rectangle([69, cy+52, 79, cy+104+4], fill=PINK_LT)
    # card
    rounded_rect(draw, [118, cy, W-30, cy+88], 14, fill=WHITE)
    draw.text((136, cy+12), title, fill=DARK, font=try_font(14, True))
    draw.text((136, cy+38), desc, fill=GRAY, font=f_sm)
    rounded_rect(draw, [136, cy+60, 136+len(time_)*12+20, cy+82], 12, fill=PINK_LT)
    draw.text((146, cy+63), time_, fill=PINK, font=f_sm)
y += 600

# ── FORM ─────────────────────────────────────────────
draw.rectangle([0, y, W, y+520], fill=BEIGE)
rounded_rect(draw, [100, y+30, W-100, y+490], 22, fill=WHITE)
draw.text((W//2 - 110, y+52), "✨ 無料で転職相談する", fill=DARK, font=f_lg)
draw.text((W//2 - 150, y+90), "30秒で登録完了。女性アドバイザーが丁寧にサポートします。", fill=GRAY, font=f_sm)

fields = ["お名前（必須）","メールアドレス（必須）","電話番号（必須）","現在の状況（必須）","希望の働き方（任意）"]
for i, label in enumerate(fields):
    fy = y + 128 + i * 56
    draw.text((136, fy), label, fill=DARK, font=try_font(12, True))
    rounded_rect(draw, [136, fy+18, W-136, fy+46], 8, fill=hex2rgb("F9F3EE"), outline=BEIGE_DK, outline_width=1)

# submit
gradient_rect(img, 160, y+420, W-160, y+465, PINK, ROSE, vertical=False)
draw.rounded_rectangle([160, y+420, W-160, y+465], radius=25, fill=None)
rounded_rect(draw, [160, y+420, W-160, y+465], 25, fill=PINK)
draw.text((W//2 - 130, y+434), "無料で転職相談を申し込む →", fill=WHITE, font=f_bold)
y += 520

# ── FOOTER ───────────────────────────────────────────
draw.rectangle([0, y, W, y+120], fill=DARK)
draw.text((W//2 - 70, y+18), "OfficeColor", fill=PINK, font=f_bold)
draw.text((W//2 - 135, y+48), "女性のための事務職専門転職エージェント", fill=hex2rgb("AAA09A"), font=f_sm)
draw.text((W//2 - 150, y+72), "プライバシーポリシー ｜ 利用規約 ｜ 会社情報 ｜ お問い合わせ", fill=hex2rgb("888070"), font=f_sm)
draw.text((W//2 - 130, y+96), "© 2024 OfficeColor. All Rights Reserved.", fill=hex2rgb("666060"), font=f_sm)
y += 120

# crop to actual height
final = img.crop((0, 0, W, y))
final.save("/home/user/-/lp_preview.png", quality=95)
print(f"Saved: {W}x{y}px")
