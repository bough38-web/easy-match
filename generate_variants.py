from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageChops
import os
import random
import math

def get_font(size):
    font_path = "/System/Library/Fonts/Supplemental/Arial Black.ttf"
    if not os.path.exists(font_path):
        font_path = "/System/Library/Fonts/Helvetica.ttc"
    try:
        return ImageFont.truetype(font_path, size, index=0)
    except:
        return ImageFont.load_default()

def draw_poly_style(img, draw, W, H):
    # E (Left) - Warm Colors
    # M (Right) - Cool Colors
    # Overlap in middle
    
    # We will draw "E" and "M" visually using rectangles to simulate the blocky reference
    # Ref: E is boxy. M is boxy.
    
    # Metrics
    m = 20
    h = H - 2*m
    w = (W - 2*m) // 2 + 20 # overlap
    
    # E (Left)
    colors_warm = ["#e74c3c", "#f39c12", "#e67e22", "#d35400", "#c0392b", "#e91e63"]
    # Draw E as blocks
    # Vertical bar
    bw = w // 2
    ex = m
    ey = m
    
    # Mask for E
    mask_e = Image.new('L', (W, H), 0)
    d_e = ImageDraw.Draw(mask_e)
    # E body
    d_e.rectangle([ex, ey, ex + w, ey + h], fill=255) # Placeholder for bounding box of E? No, let's draw text.
    
    # Let's draw the TEXT "E" huge
    font = get_font(180)
    d_e.text((ex+10, ey-20), "E", font=font, fill=255)
    
    # M (Right)
    mask_m = Image.new('L', (W, H), 0)
    d_m = ImageDraw.Draw(mask_m)
    d_m.text((ex + w - 40, ey-20), "M", font=font, fill=255)
    
    # Now fill E with Poly pattern
    # Determine bounds of E
    e_bbox = mask_e.getbbox()
    if e_bbox:
        # Draw random polygons inside E
        for i in range(50):
            x1 = random.randint(e_bbox[0], e_bbox[2])
            y1 = random.randint(e_bbox[1], e_bbox[3])
            x2 = x1 + random.randint(20, 80)
            y2 = y1 + random.randint(20, 80)
            c = random.choice(colors_warm)
            draw.polygon([(x1,y1), (x2, y1), (x1, y2)], fill=c)
            draw.polygon([(x2,y2), (x2, y1), (x1, y2)], fill=random.choice(colors_warm))
            
    # Apply Mask E
    # Actually, simpler: create a colorful texture, then mask it with E.
    tex_warm = Image.new('RGBA', (W, H), (0,0,0,0))
    d_warm = ImageDraw.Draw(tex_warm)
    for i in range(200):
        x = random.randint(0, W//2 + 100)
        y = random.randint(0, H)
        sz = random.randint(20, 60)
        d_warm.regular_polygon((x,y,sz), 3, rotation=random.randint(0,360), fill=random.choice(colors_warm))
    
    # Cut E
    img.paste(tex_warm, (0,0), mask_e)
    
    # Cool Texture
    colors_cool = ["#3498db", "#2980b9", "#1abc9c", "#9b59b6", "#8e44ad", "#2ecc71"]
    tex_cool = Image.new('RGBA', (W, H), (0,0,0,0))
    d_cool = ImageDraw.Draw(tex_cool)
    for i in range(200):
        x = random.randint(W//2 - 50, W)
        y = random.randint(0, H)
        sz = random.randint(20, 60)
        d_cool.regular_polygon((x,y,sz), 3, rotation=random.randint(0,360), fill=random.choice(colors_cool))

    # Cut M
    img.paste(tex_cool, (0,0), mask_m)
    
    # Connector (Knob) - The reference has a knob from E to M
    # Center position
    cx, cy = W//2, H//2
    knob_r = 30
    # Draw Knob Circle in Warm Color
    draw.ellipse([cx-knob_r-10, cy-knob_r, cx+knob_r-10, cy+knob_r], fill="#f39c12", outline="#c0392b", width=2)
    
    # Add text shadow/outline for readability? 
    # The reference has sharp edges.
    
    print("Generated Poly Style")


def draw_gradient_style(img, draw, W, H):
    # Smooth Gradient "EM"
    # Create gradient texture
    tex = Image.new('RGBA', (W, H), (0,0,0,0))
    # Horizontal gradient Pink (#ff758c) to Cyan (#007ade)
    for x in range(W):
        r = int(255 - (x/W)*255)
        g = int(117 + (x/W)*(122-117)) # a bit arbitrary
        b = int(140 + (x/W)*(222-140))
        # Better: Interpolate
        # color1 = (255, 0, 128) # Pink
        # color2 = (0, 200, 255) # Cyan
        r = int(255 + (x/W)*(0-255))
        g = int(0 + (x/W)*(200-0))
        b = int(128 + (x/W)*(255-128))
        
        for y in range(H):
            tex.putpixel((x,y), (r,g,b,255))
            
    # Draw Mask "EM" interlocking
    mask = Image.new('L', (W, H), 0)
    d_mask = ImageDraw.Draw(mask)
    font = get_font(160)
    
    # E
    d_mask.text((20, 20), "E", font=font, fill=255)
    # M - slightly overlapped
    d_mask.text((W/2 - 20, 20), "M", font=font, fill=255)
    
    # Composite
    img.paste(tex, (0,0), mask)
    
    # Add Gloss/Shine?
    draw.ellipse([20, 20, 100, 80], fill=(255,255,255,50))
    print("Generated Gradient Style")

def draw_neon_style(img, draw, W, H):
    # Dark background (actually transparent, but text has glow)
    # E: Cyan Outline, Black Fill
    # M: Magenta Outline, Black Fill
    
    font = get_font(160)
    
    # Glow layer
    glow = Image.new('RGBA', (W, H), (0,0,0,0))
    d_glow = ImageDraw.Draw(glow)
    
    # E Glow
    for off in range(10, 0, -1):
        op = (10-off)*15
        d_glow.text((20, 20), "E", font=font, fill=(0, 255, 255, op), stroke_width=off, stroke_fill=(0, 255, 255, op))
    
    # M Glow
    for off in range(10, 0, -1):
        op = (10-off)*15
        d_glow.text((W/2 - 20, 20), "M", font=font, fill=(255, 0, 255, op), stroke_width=off, stroke_fill=(255, 0, 255, op))
        
    img.alpha_composite(glow)
    
    # Core Text
    draw.text((20, 20), "E", font=font, fill=(0,0,0,200), stroke_width=2, stroke_fill="white")
    draw.text((W/2 - 20, 20), "M", font=font, fill=(0,0,0,200), stroke_width=2, stroke_fill="white")
    
    print("Generated Neon Style")


def generate():
    os.makedirs("assets", exist_ok=True)
    W, H = 400, 200
    
    # Variant 1: Poly
    img1 = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    draw1 = ImageDraw.Draw(img1)
    draw_poly_style(img1, draw1, W, H)
    img1.save("assets/logo_v1.png")
    
    # Variant 2: Gradient
    img2 = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    draw2 = ImageDraw.Draw(img2)
    draw_gradient_style(img2, draw2, W, H)
    img2.save("assets/logo_v2.png")
    
    # Variant 3: Neon
    img3 = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    draw3 = ImageDraw.Draw(img3)
    draw_neon_style(img3, draw3, W, H)
    img3.save("assets/logo_v3.png")

if __name__ == "__main__":
    generate()
