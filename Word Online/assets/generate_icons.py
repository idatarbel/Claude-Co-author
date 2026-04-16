"""Generate Office Add-in ribbon icons for Claude Co-author.

Run from the repo root:

    python "Word Online/assets/generate_icons.py"

Produces icon-16.png, icon-32.png, icon-80.png in this directory.
"""
import os
from PIL import Image, ImageDraw

# Anthropic-adjacent warm orange + clean white.
ORANGE = (204, 120, 92, 255)   # #CC785C
WHITE  = (255, 255, 255, 255)

HERE = os.path.dirname(os.path.abspath(__file__))


def make_master(size: int) -> Image.Image:
    """Render the icon at a given size, pixel-aligned per size for clean edges."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Rounded square background — Microsoft ribbon convention.
    corner = max(2, size // 6)
    draw.rounded_rectangle([0, 0, size - 1, size - 1], radius=corner, fill=ORANGE)

    # Chat bubble body: a rounded rectangle that leaves room for a tail.
    pad = max(2, size // 5)
    top, left = pad, pad
    right = size - pad
    bottom = size - pad - max(2, size // 8)
    bubble_r = max(2, size // 8)
    draw.rounded_rectangle([left, top, right, bottom], radius=bubble_r, fill=WHITE)

    # Bubble tail pointing down-left, like a speech balloon.
    tail_base_x = left + max(2, size // 6)
    tail_size   = max(2, size // 8)
    draw.polygon([
        (tail_base_x,             bottom - 1),
        (tail_base_x + tail_size, bottom - 1),
        (tail_base_x,             bottom - 1 + tail_size),
    ], fill=WHITE)

    # Three orange dots inside the bubble — "…" ellipsis / message indicator.
    dot_r = max(1, size // 20)
    spacing = max(3, size // 7)
    cy = (top + bottom) // 2
    cx = size // 2
    for dx in (-spacing, 0, spacing):
        draw.ellipse([cx + dx - dot_r, cy - dot_r, cx + dx + dot_r, cy + dot_r], fill=ORANGE)

    return img


def main() -> None:
    # Build a high-resolution master, then also render the small sizes natively
    # (for crisp pixel alignment at 16px).
    targets = [16, 32, 80]
    for size in targets:
        native = make_master(size)
        # For 80px, also downscale a 2x render to smooth anti-aliasing.
        if size == 80:
            big = make_master(size * 4)
            native = big.resize((size, size), Image.LANCZOS)
        out = os.path.join(HERE, f"icon-{size}.png")
        native.save(out, "PNG", optimize=True)
        print(f"Wrote {out} ({size}x{size})")


if __name__ == "__main__":
    main()
