"""Genera il logo di Formazioni PZZ (assets/app_icon.ico, app_icon.png, logo_header.png).

Uso: python tools/genera_logo.py  (richiede Pillow)
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"

NAVY = (10, 34, 53)
TEAL = (42, 142, 142)
GOLD = (233, 176, 72)
GOLD_DARK = (196, 135, 38)
PAPER = (255, 255, 255)
LINE = (178, 200, 207)


def _gradient(size, c1, c2):
    """Diagonal gradient from top-left (c1) to bottom-right (c2)."""
    w, h = size
    base = Image.new("RGB", (w, h), c1)
    top = Image.new("RGB", (w, h), c2)
    ramp = Image.linear_gradient("L").resize((w, h))  # 0 at top, 255 at bottom
    mask = Image.blend(ramp, ramp.transpose(Image.Transpose.ROTATE_90), 0.5)
    return Image.composite(top, base, mask)


def draw_logo(px: int, detailed: bool = True) -> Image.Image:
    ss = 4  # supersampling for smooth edges
    S = px * ss
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))

    # Soft drop shadow
    pad = int(S * 0.04)
    radius = int(S * 0.23)
    if detailed:
        shadow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        ImageDraw.Draw(shadow).rounded_rectangle(
            (pad, pad + int(S * 0.02), S - pad, S - pad + int(S * 0.02)),
            radius, fill=(0, 0, 0, 90))
        img = Image.alpha_composite(img, shadow.filter(ImageFilter.GaussianBlur(S * 0.015)))

    # Background tile with gradient
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle((pad, pad, S - pad, S - pad), radius, fill=255)
    grad = _gradient((S, S), TEAL, NAVY).convert("RGBA")
    img.paste(grad, (0, 0), mask)

    d = ImageDraw.Draw(img)

    # Document sheet with folded corner
    x0, y0 = S * 0.27, S * 0.19
    x1, y1 = S * 0.69, S * 0.79
    fold = S * 0.12
    d.polygon([(x0, y0), (x1 - fold, y0), (x1, y0 + fold), (x1, y1), (x0, y1)], fill=PAPER)
    d.polygon([(x1 - fold, y0), (x1 - fold, y0 + fold), (x1, y0 + fold)], fill=LINE)

    # Text lines on the sheet
    lw = max(ss, int(S * 0.035))
    lines = [(0.33, 0.58), (0.33, 0.52), (0.33, 0.47)] if detailed else [(0.33, 0.58), (0.33, 0.47)]
    ys = [0.36, 0.45, 0.54] if detailed else [0.38, 0.50]
    for (lx0, lx1), ly in zip(lines, ys):
        d.rounded_rectangle((S * lx0, S * ly, S * lx1, S * ly + lw), lw // 2, fill=LINE)

    # Gold badge with check mark
    cx, cy, r = S * 0.66, S * 0.70, S * 0.17
    d.ellipse((cx - r - ss * 3, cy - r - ss * 3, cx + r + ss * 3, cy + r + ss * 3), fill=NAVY)
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=GOLD)
    d.ellipse((cx - r * 0.82, cy - r * 0.82, cx + r * 0.82, cy + r * 0.82), outline=GOLD_DARK,
              width=max(1, int(S * 0.008)))
    cw = int(S * 0.045)
    pts = [(cx - r * 0.45, cy + r * 0.02), (cx - r * 0.1, cy + r * 0.38), (cx + r * 0.5, cy - r * 0.35)]
    d.line(pts, fill=PAPER, width=cw, joint="curve")
    for p in (pts[0], pts[2]):
        d.ellipse((p[0] - cw / 2, p[1] - cw / 2, p[0] + cw / 2, p[1] + cw / 2), fill=PAPER)

    return img.resize((px, px), Image.LANCZOS)


def main() -> None:
    ASSETS.mkdir(exist_ok=True)
    big = draw_logo(256)
    big.save(ASSETS / "app_icon.png")
    draw_logo(64).save(ASSETS / "logo_header.png")
    sizes = [16, 20, 24, 32, 40, 48, 64, 128, 256]
    frames = [draw_logo(s, detailed=s >= 32) for s in sizes]
    frames[-1].save(ASSETS / "app_icon.ico", format="ICO",
                    sizes=[(s, s) for s in sizes], append_images=frames[:-1])
    print("Logo generato in", ASSETS)


if __name__ == "__main__":
    main()
