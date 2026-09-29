"""
Genere assets/aya.ico (icone du logiciel, de l'EXE et des raccourcis).
A lancer seulement pour changer l'icone :  python packaging/make_icon.py
(necessite Pillow ; le fichier .ico produit est versionne dans le depot).
"""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).resolve().parent.parent / "assets" / "aya.ico"
GREEN = (16, 124, 65)          # vert Excel


def _font(size: int):
    for name in ("DejaVuSans-Bold.ttf", "arialbd.ttf", "Arial Bold.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def draw(size: int = 256) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    r = size // 6
    d.rounded_rectangle((0, 0, size - 1, size - 1), radius=r, fill=GREEN)
    # petite grille facon tableur
    step = size // 8
    for i in range(1, 8):
        d.line((i * step, size * 0.62, i * step, size - step // 2), fill=(255, 255, 255, 70), width=max(1, size // 128))
    d.line((step // 2, size * 0.75, size - step // 2, size * 0.75), fill=(255, 255, 255, 70), width=max(1, size // 128))
    font = _font(int(size * 0.38))
    text = "AYA"
    box = d.textbbox((0, 0), text, font=font)
    x = (size - (box[2] - box[0])) / 2 - box[0]
    y = size * 0.30 - (box[3] - box[1]) / 2 - box[1]
    d.text((x, y), text, font=font, fill="white")
    return img


if __name__ == "__main__":
    OUT.parent.mkdir(parents=True, exist_ok=True)
    draw().save(OUT, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print(f"Icone ecrite : {OUT}")
