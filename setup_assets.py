import os
import shutil
from PIL import Image

src = "/Users/heebonpark/.gemini/antigravity/brain/19fc9aa5-1f74-4a5f-bcb6-a91aa906d597/easy_match_logo_final_1771040947191.png"
dst_dir = "assets"
os.makedirs(dst_dir, exist_ok=True)
dst = os.path.join(dst_dir, "logo_header.png")

try:
    if not os.path.exists(src):
        print(f"Source file not found: {src}")
        # Try finding any png in that dir
        d = os.path.dirname(src)
        files = [f for f in os.listdir(d) if f.endswith(".png") and "logo" in f]
        if files:
            src = os.path.join(d, files[0])
            print(f"Using alternative source: {src}")
        else:
            print("No logo found.")
            exit(1)

    img = Image.open(src)
    # Resize to height 80px, keeping aspect ratio
    h = 80
    w = int(img.width * h / img.height)
    img = img.resize((w, h), Image.Resampling.LANCZOS)
    img.save(dst)
    print(f"Saved resized logo to {dst}")
except Exception as e:
    print(f"Error: {e}")
