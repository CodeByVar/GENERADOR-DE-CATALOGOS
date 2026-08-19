from PIL import Image

img = Image.open("Logo DonManu.jpg")
pixels = img.load()
width, height = img.size

# Find the most common blue pixel (high blue, lower red)
color_counts = {}
for y in range(height):
    for x in range(width):
        r, g, b = pixels[x, y][:3]
        if b > 100 and r < 100:
            color_counts[(r, g, b)] = color_counts.get((r, g, b), 0) + 1

if color_counts:
    # Get the most common blue color
    most_common = max(color_counts, key=color_counts.get)
    print(f"Most common blue color in Logo DonManu.jpg: {most_common}")
    print(f"Hex: #{most_common[0]:02X}{most_common[1]:02X}{most_common[2]:02X}")
else:
    print("No blue color found in Logo DonManu.jpg")
