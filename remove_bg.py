from PIL import Image, ImageFilter, ImageEnhance
import os

def remove_white_bg():
    path = "assets/logo_vo.jpg"
    if not os.path.exists(path):
        print("Logo source (logo_vo.jpg) not found!")
        return

    # Open image
    img = Image.open(path).convert("RGBA")
    
    # 1. Background Mask using Tolerance
    # Since the image is likely a screenshot or downloaded file with compression,
    # pure white (255,255,255) might be (250,250,250) or have noise.
    # We want to remove all light pixels.
    
    datas = img.getdata()
    new_data = []
    
    # Aggressive Threshold for White Background (removes faint shadows/borders)
    threshold = 230 
    
    for item in datas:
        # Check if pixel is "White-ish"
        if item[0] > threshold and item[1] > threshold and item[2] > threshold:
            new_data.append((255, 255, 255, 0)) # Make Transparent
        else:
            new_data.append(item)
            
    img.putdata(new_data)
    
    # 2. Trim Transparent Edges (Auto-Crop)
    bbox = img.getbbox() # Bounding box of non-zero alpha
    if bbox:
        img = img.crop(bbox)
        
    # 3. Enhance Sharpness
    enhancer = ImageEnhance.Sharpness(img)
    img = enhancer.enhance(2.0) # Sharpen significantly
    
    # 4. Resize if too large?
    # Current UI uses 160x80 roughly.
    # Uploaded image `media__1771414335572.png` is 46KB, likely small enough.
    # But if it's huge, shrinking improves apparent sharpness.
    if img.width > 500:
        ratio = 500 / img.width
        new_h = int(img.height * ratio)
        img = img.resize((500, new_h), Image.Resampling.LANCZOS)
        
    img.save("assets/logo.png", "PNG")
    print("Processed logo: Removed white background, cropped, and sharpened.")

if __name__ == "__main__":
    remove_white_bg()
