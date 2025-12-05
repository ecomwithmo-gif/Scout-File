from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def main() -> None:
    base_dir = Path(__file__).resolve().parent.parent
    assets_dir = base_dir / "assets"
    assets_dir.mkdir(exist_ok=True)

    size = (256, 256)
    img = Image.new("RGBA", size, "#0f172a")
    draw = ImageDraw.Draw(img)

    # Subtle gradient
    for i in range(size[1]):
        blend = int(15 + (59 - 15) * (i / size[1]))
        draw.line([(0, i), (size[0], i)], fill=(59, 130, blend, 255))

    # Rounded square
    padding = 36
    draw.rounded_rectangle(
        (padding, padding, size[0] - padding, size[1] - padding),
        radius=44,
        fill=(15, 23, 42, 220),
    )

    # Centered text
    text = "EF"
    font_size = 120
    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except Exception:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    position = ((size[0] - text_width) / 2, (size[1] - text_height) / 2 - bbox[1])
    draw.text(position, text, fill="#f8fafc", font=font)

    icon_path = assets_dir / "launcher.ico"
    img.save(icon_path, format="ICO")
    print(f"Icon saved to {icon_path}")


if __name__ == "__main__":
    main()

