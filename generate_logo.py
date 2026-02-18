from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageChops, ImageOps
import os
import sys

def rounded_rect(draw, box, radius, fill):
    draw.rounded_rectangle(box, radius=radius, fill=fill)

def draw_plastic_piece(img, x, y, s, color_hex, tabs):
    """
    Draws a puzzle piece with "Plastic/Gel" 3D effect.
    tabs: [Top, Right, Bottom, Left] (1=Out, 0=Flat/In)
    """
    # Create high-res mask for shape
    m_scale = 4
    ms = s * m_scale
    mx, my = x * m_scale, y * m_scale
    mw, mh = img.width * m_scale, img.height * m_scale
    
    mask = Image.new('L', (mw, mh), 0)
    d = ImageDraw.Draw(mask)
    
    ts = int(ms * 0.35)
    radius = int(ms * 0.2)
    
    # Draw Shape on Mask
    # Body
    d.rounded_rectangle([mx, my, mx+ms, my+ms], radius=radius, fill=255)
    
    # Tabs
    # Top
    cx, cy = mx + ms//2, my
    if tabs[0] == 1:
        d.rectangle([cx-ts//4, cy-ts//2, cx+ts//4, cy], fill=255)
        d.ellipse([cx-ts//2, cy-ts, cx+ts//2, cy], fill=255)
    # Right
    cx, cy = mx + ms, my + ms//2
    if tabs[1] == 1:
        d.rectangle([cx, cy-ts//4, cx+ts//2, cy+ts//4], fill=255)
        d.ellipse([cx, cy-ts//2, cx+ts, cy+ts//2], fill=255)
    # Bottom
    cx, cy = mx + ms//2, my + ms
    if tabs[2] == 1:
        d.rectangle([cx-ts//4, cy, cx+ts//4, cy+ts//2], fill=255)
        d.ellipse([cx-ts//2, cy, cx+ts//2, cy+ts], fill=255)
    # Left
    cx, cy = mx, my + ms//2
    if tabs[3] == 1:
        d.rectangle([cx-ts//2, cy-ts//4, cx, cy+ts//4], fill=255)
        d.ellipse([cx-ts, cy-ts//2, cx, cy+ts//2], fill=255)

    # Resize mask down for AA
    mask = mask.resize((img.width, img.height), Image.Resampling.LANCZOS)
    
    # Create Base Color Layer
    c_base = Image.new('RGBA', img.size, color_hex)
    
    # Create Highlight/Shadow Layers
    # We use the mask to create bevels via blur/offset
    # Highlight (Top-Left)
    # Shadow (Bottom-Right)
    
    # Extract edge?
    # Simpler: Inner Glow/Shadow
    
    # Bevel Highlight: White, Offset -2, Blurred
    # Bevel Shadow: Black, Offset +2, Blurred
    
    # To do this cleanly on the shape:
    # 1. Composite Base Color clipped by Mask.
    base_shape = Image.new('RGBA', img.size, (0,0,0,0))
    base_shape.paste(c_base, (0,0), mask)
    
    # 2. Add Highlights (Plastic Shine)
    # Large soft white gloss on Top-Left
    gloss = Image.new('RGBA', img.size, (0,0,0,0))
    g_draw = ImageDraw.Draw(gloss)
    # Gloss blobs
    # Adjust coordinates back to normal scale
    gx, gy = x, y
    gs = s
    g_draw.ellipse([gx+10, gy+10, gx+gs//2, gy+gs//2], fill=(255,255,255,100))
    gloss = gloss.filter(ImageFilter.GaussianBlur(10))
    
    # Clip gloss to mask
    base_shape.paste(gloss, (0,0), mask)
    
    # 3. Add Edge Highlight (Bevel)
    # Find edges using mask?
    # Simple trick: Draw White shape offset -1, cut by original mask.
    
    # Paste into main image
    img.alpha_composite(base_shape)


def draw_3d_text(img, text, cx, cy):
    # Yellow/Gold Text with Bevel
    font_path = "/System/Library/Fonts/Supplemental/Arial Rounded Bold.ttf"
    if not os.path.exists(font_path):
        font_path = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
    if not os.path.exists(font_path):
        font_path = "/System/Library/Fonts/Helvetica.ttc"
        
    font_size = 130
    try:
        font = ImageFont.truetype(font_path, font_size)
    except:
        font = ImageFont.load_default()
        
    draw = ImageDraw.Draw(img)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
    tx, ty = cx - tw/2, cy - th/2 - 10 # Slight visual uplift
    
    # Colors
    c_gold = "#f1c40f"
    c_gold_light = "#f9e79f"
    c_gold_dark = "#d35400"
    
    # 1. Drop Shadow
    draw.text((tx+5, ty+5), text, font=font, fill=(0,0,0,100))
    
    # 2. Border/Stroke (Dark Orange)
    draw.text((tx, ty+2), text, font=font, fill=c_gold_dark, stroke_width=4, stroke_fill=c_gold_dark)
    
    # 3. Main Face (Gold)
    draw.text((tx, ty), text, font=font, fill=c_gold)
    
    # 4. Highlight (Top half?)
    # Crude "Bevel" text: Draw slightly shifted lighter text, masked by text?
    # Let's just draw text again at offset -1 with lighter color, but only inside?
    # Hard to do "Inside" without complex masking.
    # Simple highlight: Draw text shifted -2,-2 with white (low alpha)
    # clipped to text mask.
    
    # Create Text Mask
    txt_mask = Image.new('L', img.size, 0)
    ImageDraw.Draw(txt_mask).text((tx, ty), text, font=font, fill=255)
    
    # Create Highlight Layer
    hl = Image.new('RGBA', img.size, (0,0,0,0))
    ImageDraw.Draw(hl).text((tx-3, ty-3), text, font=font, fill=(255,255,255,150))
    
    # Composite Highlight ONLY where Text is
    img.paste(hl, (0,0), txt_mask)


def create_logo():
    W, H = 500, 400 # Square-ish canvas to fit the cross shape
    img = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    
    # Colors (Plastic Bright)
    c_green = "#00e640" # Vivid Green
    c_blue = "#2980b9"  # Vivid Blue
    
    s = 140 # Piece size
    cx, cy = W//2, H//2
    
    # Position: Interlocking
    # Green Top-Left
    gx = cx - s + 25
    gy = cy - s + 25
    
    # Blue Bottom-Right
    bx = cx - 25
    by = cy - 25
    
    # Blue needs to be shifted to fit.
    # Green (0,0) -> Top Tab, Left Tab. Right Side is "In". Bottom is "In".
    # Blue (1,1) -> Bottom Tab, Right Tab. Left Side is "Out" (into Green). Top Side is "Out" (into Green).
    
    # Based on reference:
    # Green Piece: Top-Left. 
    #   Top: Out
    #   Left: Out
    #   Right: In (Concave)
    #   Bottom: In (Concave)
    
    # Blue Piece: Bottom-Right.
    #   Bottom: Out
    #   Right: Out
    #   Left: Out (Convex) -> Fits into Green's Right In.
    #   Top: Out (Convex) -> Fits into Green's Bottom In.
    
    # So both pieces have "Out" tabs on their shared edges?
    # No, Blue has tabs sticking INTO Green.
    
    # Draw Green First (Bottom Layer)
    # Tabs: Top=1, Right=0, Bottom=0, Left=1
    draw_plastic_piece(img, gx, gy, s, c_green, [1, 0, 0, 1])
    
    # Draw Blue Second (Top Layer)
    # Tabs: Top=1 (Into Green), Right=1, Bottom=1, Left=1 (Into Green)
    # Note: To look right, Blue is positioned so its Top/Left tabs overlap Green's body.
    # Green's body is at (gx, gy). Size s.
    # Green's Bottom Edge is at gy+s.
    # Blue's Top Edge is at by.
    # Blue's Top Tab center is at bx+s/2, by.
    # We want Blue's Top Tab to plug into Green's Bottom Edge (center gx+s/2, gy+s).
    # So bx = gx. by = gy+s - (tab_height?). No, standard interlocking is body-to-body offset = s.
    # Let's align exactly:
    bx_real = gx + s * 0.7 # Slight overlap for "Tight" fit
    by_real = gy + s * 0.7
    
    draw_plastic_piece(img, bx_real, by_real, s, c_blue, [1, 1, 1, 1])
    
    # Text on Top
    draw_3d_text(img, "EM", cx, cy)
    
    # Crop to content?
    bbox = img.getbbox()
    if bbox:
        # Add padding
        pad = 20
        bbox = (max(0, bbox[0]-pad), max(0, bbox[1]-pad), min(W, bbox[2]+pad), min(H, bbox[3]+pad))
        img = img.crop(bbox)
    
    os.makedirs("assets", exist_ok=True)
    img.save("assets/logo.png", "PNG")
    print("Generated Plastic 3D Puzzle Logo at assets/logo.png")

if __name__ == "__main__":
    create_logo()
