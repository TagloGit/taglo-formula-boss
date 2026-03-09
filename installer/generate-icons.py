"""Generate logo.ico and wizard-banner.bmp from the pixel-art source in resources/."""

from PIL import Image, ImageDraw, ImageFont
import struct, io, os

PALETTE = {
    0: (0, 0, 0, 0),        # transparent
    1: (26, 26, 46, 255),    # #1a1a2e
    4: (255, 215, 0, 255),   # #ffd700
    5: (255, 255, 255, 255), # #ffffff
    49: (153, 27, 27, 255),  # #991b1b
    52: (194, 65, 12, 255),  # #c2410c
    53: (251, 146, 60, 255), # #fb923c
    55: (232, 133, 58, 255), # #e8853a
    56: (217, 117, 38, 255), # #d97526
    57: (184, 94, 26, 255),  # #b85e1a
    58: (154, 78, 21, 255),  # #9a4e15
}

GRID = [
    [0,0,0,0,0,53,53,0,0,0,0,0,0,53,53,0,0,0,0,0],
    [0,0,0,0,0,52,52,1,1,1,1,1,1,52,52,0,0,0,0,0],
    [0,0,1,1,1,1,49,49,49,49,49,49,49,49,1,1,1,1,0,0],
    [0,1,55,55,1,49,49,49,49,49,49,49,49,49,49,1,55,55,1,0],
    [0,1,55,55,1,49,49,49,49,49,49,49,49,49,49,1,55,55,1,0],
    [0,1,55,55,1,49,5,5,49,49,49,49,5,5,49,1,55,55,1,0],
    [0,1,53,53,49,49,5,1,49,49,49,49,1,5,49,49,53,53,1,0],
    [0,1,53,53,49,49,49,49,49,49,49,49,49,49,49,49,53,53,1,0],
    [0,1,53,53,49,49,49,49,49,52,52,49,49,49,49,49,53,53,1,0],
    [0,1,53,53,49,49,49,49,49,49,49,49,49,49,49,49,53,53,1,0],
    [0,1,56,56,1,1,1,1,1,1,1,1,1,1,1,1,56,56,1,0],
    [0,1,56,56,1,5,1,5,1,5,1,5,1,5,5,1,56,56,1,0],
    [0,1,56,56,1,49,49,49,52,52,52,49,49,49,49,1,56,56,1,0],
    [0,1,57,57,1,1,5,1,5,1,5,1,5,1,1,1,57,57,1,0],
    [0,1,57,57,57,49,49,49,49,49,49,49,49,49,49,57,57,57,1,0],
    [0,1,57,57,57,1,1,1,1,1,1,1,1,1,1,57,57,57,1,0],
    [1,55,55,57,57,1,0,0,0,0,0,0,0,0,1,57,57,55,55,1],
    [1,55,53,1,57,57,1,0,0,0,0,0,0,1,57,57,1,53,55,1],
    [1,55,53,1,57,57,1,0,0,0,0,0,0,1,57,57,1,53,55,1],
    [1,1,1,1,1,1,1,0,0,0,0,0,0,1,1,1,1,1,1,1],
]

COLS = len(GRID[0])  # 20
ROWS = len(GRID)     # 20


def render_base() -> Image.Image:
    """Render the pixel art at 1x (20x20)."""
    img = Image.new("RGBA", (COLS, ROWS), (0, 0, 0, 0))
    for y, row in enumerate(GRID):
        for x, code in enumerate(row):
            img.putpixel((x, y), PALETTE[code])
    return img


def make_ico(base: Image.Image, path: str) -> None:
    """Save a multi-resolution ICO (16, 32, 48, 256)."""
    target_sizes = [16, 32, 48, 256]
    # Pillow ICO save picks from the source image at requested sizes,
    # so we create a large image and let it do the downscaling.
    large = base.resize((256, 256), Image.NEAREST)
    large.save(path, format="ICO", sizes=[(s, s) for s in target_sizes])


def make_wizard_banner(base: Image.Image, path: str) -> None:
    """Create InnoSetup wizard sidebar image (164x314 BMP).

    Places the logo centred near the top on a dark-blue gradient background.
    """
    W, H = 164, 314
    banner = Image.new("RGB", (W, H))
    draw = ImageDraw.Draw(banner)

    # Dark-blue gradient background
    for y in range(H):
        t = y / H
        r = int(26 * (1 - t) + 10 * t)
        g = int(26 * (1 - t) + 10 * t)
        b = int(46 * (1 - t) + 30 * t)
        draw.line([(0, y), (W, y)], fill=(r, g, b))

    # Place logo centred, 40px from top, scaled to ~120px
    logo_size = 120
    logo = base.resize((logo_size, logo_size), Image.NEAREST)
    # Composite onto an RGB background matching the gradient at that Y
    logo_x = (W - logo_size) // 2
    logo_y = 40
    # Create RGB version with dark background for transparent pixels
    logo_rgb = Image.new("RGB", (logo_size, logo_size), (26, 26, 46))
    logo_rgb.paste(logo, mask=logo.split()[3])
    banner.paste(logo_rgb, (logo_x, logo_y))

    # Add "Formula Boss" text below logo
    text_y = logo_y + logo_size + 20
    try:
        font = ImageFont.truetype("arial.ttf", 16)
        font_small = ImageFont.truetype("arial.ttf", 12)
    except OSError:
        font = ImageFont.load_default()
        font_small = font

    # "Formula" centred
    bbox = draw.textbbox((0, 0), "Formula", font=font)
    tw = bbox[2] - bbox[0]
    draw.text(((W - tw) // 2, text_y), "Formula", fill=(251, 146, 60), font=font)

    # "Boss" centred below
    bbox = draw.textbbox((0, 0), "Boss", font=font)
    tw = bbox[2] - bbox[0]
    draw.text(((W - tw) // 2, text_y + 22), "Boss", fill=(255, 215, 0), font=font)

    banner.save(path, format="BMP")


def make_wizard_small(base: Image.Image, path: str) -> None:
    """Create InnoSetup small wizard image (55x55 BMP)."""
    logo = base.resize((55, 55), Image.NEAREST)
    logo_rgb = Image.new("RGB", (55, 55), (255, 255, 255))
    logo_rgb.paste(logo, mask=logo.split()[3])
    logo_rgb.save(path, format="BMP")


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    base = render_base()
    make_ico(base, os.path.join(script_dir, "logo.ico"))
    make_wizard_banner(base, os.path.join(script_dir, "wizard-banner.bmp"))
    make_wizard_small(base, os.path.join(script_dir, "wizard-small.bmp"))
    print("Generated: installer/logo.ico, installer/wizard-banner.bmp, installer/wizard-small.bmp")
