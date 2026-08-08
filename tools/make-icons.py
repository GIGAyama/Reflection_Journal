# アイコン一式を作る（GIGA Standard v5 §2-6 / §3-2 / §3-7）
#
#   icon-192 / icon-512 …… purpose:any。もとの角丸のまま（透明あり）
#   maskable-192 / -512 …… 下地を端まで伸ばし、中身はセーフゾーン（中央80%の円）に収める
#   icon-180 …………………… apple-touch-icon。iOS は透明を黒で埋めるので透明を一切含めない
#
# 中身（本・芽・矢印）は色で切り出す。角丸プレートの輪郭線が縁に沿って残るため、
# 縁から 48px の帯を落としてから外接矩形を採る。
# 下地は元絵の四隅の色を双一次補間した滑らかなグラデーション。プレートごと載せないので継ぎ目は出ない。
from PIL import Image, ImageFilter, ImageDraw
import os, math

SRC, OUT = os.path.join(os.path.dirname(__file__), 'icon-source.png'), os.path.join(os.path.dirname(__file__), '..', 'docs')
os.makedirs(OUT, exist_ok=True)
src = Image.open(SRC).convert('RGBA')
W, H = src.size
px = src.load()

# ── ① 中身のマスク ──
mask = Image.new('L', (W, H), 0)
mp = mask.load()
for y in range(H):
    for x in range(W):
        r, g, b, a = px[x, y]
        if a < 40:
            continue
        bright = (r * 299 + g * 587 + b * 114) // 1000
        if b > 150 or g > r or bright < 120:      # クリーム/白 ・ 緑 ・ 濃い輪郭
            mp[x, y] = 255
mask = mask.filter(ImageFilter.MaxFilter(5)).filter(ImageFilter.MinFilter(5))
d = ImageDraw.Draw(mask)
for box in ([0, 0, W - 1, 47], [0, H - 48, W - 1, H - 1], [0, 0, 47, H - 1], [W - 48, 0, W - 1, H - 1]):
    d.rectangle(box, fill=0)
BBOX = mask.getbbox()
content = src.crop(BBOX)
# 切り出しの縁が硬くならないよう、マスクを少しぼかしてアルファに使う
soft = mask.crop(BBOX).filter(ImageFilter.GaussianBlur(1.2))
content.putalpha(soft)
cw, ch = content.size

# ── ② 下地：元絵の四隅の色から作る滑らかなグラデーション ──
CORNERS = [px[40, 40][:3], px[W - 41, 40][:3], px[40, H - 41][:3], px[W - 41, H - 41][:3]]
def backdrop(size):
    small = Image.new('RGB', (2, 2))
    small.putpixel((0, 0), CORNERS[0]); small.putpixel((1, 0), CORNERS[1])
    small.putpixel((0, 1), CORNERS[2]); small.putpixel((1, 1), CORNERS[3])
    return small.resize((size, size), Image.BICUBIC)

# ── ③ 中身が円の外へどれだけはみ出すかを実際に画素で数える（§3-7） ──
def outside_ratio(size, scale, safe=0.80):
    nw, nh = max(1, round(cw * scale)), max(1, round(ch * scale))
    m = mask.crop(BBOX).resize((nw, nh), Image.LANCZOS)
    canvas = Image.new('L', (size, size), 0)
    canvas.paste(m, ((size - nw) // 2, (size - nh) // 2))
    mpx = canvas.load()
    c, r = size / 2, size * safe / 2
    outside = 0
    for y in range(size):
        for x in range(size):
            if mpx[x, y] > 40 and (x - c) ** 2 + (y - c) ** 2 > r * r:
                outside += 1
    return outside / (size * size) * 100

def fit_scale(size, safe=0.80, limit=0.2):
    """0.2% 以下に収まる、いちばん大きい倍率を探す"""
    lo, hi = 0.2, (size * 1.0) / max(cw, ch)
    for _ in range(12):
        mid = (lo + hi) / 2
        if outside_ratio(size, mid, safe) <= limit:
            lo = mid
        else:
            hi = mid
    return lo

def compose(size, scale):
    bg = backdrop(size).convert('RGBA')
    nw, nh = max(1, round(cw * scale)), max(1, round(ch * scale))
    bg.alpha_composite(content.resize((nw, nh), Image.LANCZOS), ((size - nw) // 2, (size - nh) // 2))
    return bg.convert('RGB')

def save_palette(img, path, budget_kb):
    """予算（KB）に収まる中で、いちばん色数の多い版を選ぶ。
    無条件に軽い版を選ぶと、葉の緑や本の陰影が潰れて別の絵になる。"""
    chosen = None
    for c in (256, 192, 160, 128, 96, 64):
        tmp = path.replace('.png', f'.q{c}.png')
        img.convert('P', palette=Image.ADAPTIVE, colors=c, dither=Image.FLOYDSTEINBERG).save(tmp, optimize=True)
        size_kb = os.path.getsize(tmp) / 1024
        os.remove(tmp)
        if size_kb <= budget_kb:
            chosen = c
            break
    if chosen is None:
        chosen = 64
    img.convert('P', palette=Image.ADAPTIVE, colors=chosen, dither=Image.FLOYDSTEINBERG).save(path, optimize=True)
    return os.path.getsize(path), chosen


print('中身 bbox:', BBOX, f'({cw}x{ch})  四隅色: {CORNERS}')

# purpose:any … もとの絵のまま
for size in (192, 512):
    p = os.path.join(OUT, f'icon-{size}.png')
    src.resize((size, size), Image.LANCZOS).quantize(colors=192, method=Image.FASTOCTREE,
                                                     dither=Image.FLOYDSTEINBERG).save(p, optimize=True)
    print(f'  icon-{size}.png{"":>10} {os.path.getsize(p)/1024:6.1f} KB')

# maskable … 実測して 0.2% 以下に収める
s512 = fit_scale(192) * 1.0        # 192 で探して 512 にも同じ倍率を使う（比率は同じ）
for size in (192, 512):
    img = compose(size, s512 * (size / 192))
    p = os.path.join(OUT, f'icon-maskable-{size}.png')
    _, col = save_palette(img, p, 60 if size == 512 else 30)
    r = outside_ratio(min(size, 192), s512)
    print(f'  icon-maskable-{size}.png {os.path.getsize(p)/1024:6.1f} KB ({col}色)  セーフゾーン外 {r:.3f}%')

# apple-touch-icon … 透明なし。iOS 自身が角を丸めるので下地は端まで
# iOS は角を丸めるので、その内側に収まる大きさにする
p = os.path.join(OUT, 'icon-180.png')
_, col = save_palette(compose(180, (180 * 0.72) / max(cw, ch)), p, 30)
print(f'  icon-180.png{"":>12} {os.path.getsize(p)/1024:6.1f} KB ({col}色)  透明なし')
