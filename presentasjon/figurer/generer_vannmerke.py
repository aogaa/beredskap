"""Generer dempede vannmerke-versjoner av frivilligsentralenes logoer."""
import sys
sys.stdout.reconfigure(encoding="utf-8")

from PIL import Image
import os

LOGO_DIR = "C:/codex/beredskap/logo"
OUT_DIR = "C:/codex/beredskap/presentasjon/figurer"

WATERMARK_OPACITY = 0.10  # 10 % opacity — diskret


def dempet_versjon(input_path, output_path, opacity=WATERMARK_OPACITY):
    img = Image.open(input_path).convert("RGBA")
    # Reduser alpha
    alpha = img.split()[3]
    alpha = alpha.point(lambda p: int(p * opacity))
    img.putalpha(alpha)
    # Komposisjoner mot hvit bakgrunn for å få jevn demping
    bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
    composed = Image.alpha_composite(bg, img)
    composed = composed.convert("RGB")
    composed.save(output_path, "PNG")
    print(f"OK — {os.path.basename(output_path)}: opacity {opacity*100:.0f} %")


if __name__ == "__main__":
    dempet_versjon(f"{LOGO_DIR}/VAFS.png", f"{OUT_DIR}/vannmerke_VAFS.png")
    dempet_versjon(f"{LOGO_DIR}/OVFS.png", f"{OUT_DIR}/vannmerke_OVFS.png")
    print("\nVannmerker lagret i", OUT_DIR)
