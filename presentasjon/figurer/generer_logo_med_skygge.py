"""Lag versjoner av logoene med drop shadow så de matcher stilistisk."""
import sys
sys.stdout.reconfigure(encoding="utf-8")

from PIL import Image, ImageFilter
import os

LOGO_DIR = "C:/codex/beredskap/logo"
OUT_DIR = "C:/codex/beredskap/presentasjon/figurer"


def med_skygge(input_path, output_path,
               offset_pct=0.012,  # 1.2 % av bildebredden
               shadow_alpha=90,
               blur_pct=0.008,     # 0.8 % av bildebredden
               padding_pct=0.045):  # 4.5 % padding
    """Legg drop shadow på en RGBA-logo."""
    src = Image.open(input_path).convert("RGBA")
    w, h = src.size

    offset_x = int(w * offset_pct)
    offset_y = int(w * offset_pct)
    blur_radius = max(4, int(w * blur_pct))
    padding = max(20, int(w * padding_pct))

    # Lerret med padding rundt — start fullt transparent
    canvas_w = w + 2 * padding
    canvas_h = h + 2 * padding

    # Lag skygge-bilde — fullt størrelse på canvas, transparent bakgrunn
    shadow_layer = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
    # Hent alpha-kanal fra original logo
    alpha = src.split()[3]
    # Lag svart silhuett med valgt opasitet
    silhouette = Image.new("RGBA", src.size, (40, 40, 40, shadow_alpha))
    silhouette.putalpha(alpha.point(lambda p: int(p * shadow_alpha / 255)))
    # Lim inn silhuetten med offset
    shadow_layer.paste(silhouette, (padding + offset_x, padding + offset_y),
                       silhouette)
    # Blur skyggelaget
    shadow_layer = shadow_layer.filter(ImageFilter.GaussianBlur(radius=blur_radius))

    # Lag final-canvas med transparent bakgrunn
    final = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
    # Compose skygge først
    final = Image.alpha_composite(final, shadow_layer)
    # Logo oppå — paste med mask
    final.paste(src, (padding, padding), src)

    final.save(output_path, "PNG")
    print(f"OK — {os.path.basename(output_path)}: {final.size}, blur={blur_radius}, offset=({offset_x},{offset_y})")


if __name__ == "__main__":
    med_skygge(f"{LOGO_DIR}/VAFS.png", f"{OUT_DIR}/VAFS_med_skygge.png")
    med_skygge(f"{LOGO_DIR}/OVFS.png", f"{OUT_DIR}/OVFS_med_skygge.png")
    print("\nLogoer med skygge generert.")
