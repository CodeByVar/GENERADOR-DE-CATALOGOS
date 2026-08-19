from PIL import Image, ImageChops
import os

def autocrop_image(img, tolerance=20):
    if img.mode != 'RGBA':
        img = img.convert('RGBA')
    bg_color = img.getpixel((0, 0))
    bg = Image.new("RGBA", img.size, bg_color)
    diff = ImageChops.difference(img, bg)
    diff_gray = diff.convert("L")
    threshold_img = diff_gray.point(lambda p: 255 if p > tolerance else 0)
    bbox = threshold_img.getbbox()
    if bbox:
        left, top, right, bottom = bbox
        left = max(0, left - 6)
        top = max(0, top - 6)
        right = min(img.width, right + 6)
        bottom = min(img.height, bottom + 6)
        return img.crop((left, top, right, bottom))
    return img

print("Autocrop with tolerance test:")
for f in os.listdir("."):
    if f.lower().startswith("logo") and (f.lower().endswith(".png") or f.lower().endswith(".jpg") or f.lower().endswith(".webp") or f.lower().endswith(".jpeg")):
        try:
            img = Image.open(f)
            cropped = autocrop_image(img)
            print(f"File: {f} | Original: {img.size} | Cropped: {cropped.size} | Aspect: {cropped.width/cropped.height:.2f}")
        except Exception as e:
            print(f"Error: {e}")
