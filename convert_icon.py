from PIL import Image
import os

png_path = r"C:/Users/Admin/.gemini/antigravity/brain/edbb7be8-84a6-4def-a8ff-310f96759e6b/app_icon_1770129908048.png"
ico_path = r"c:/Users/Admin/OneDrive/Dokumen/Asyraf/CODING/Inter-SmartPLS/icon.ico"

try:
    img = Image.open(png_path)
    img = img.convert("RGBA")
    
    datas = img.getdata()
    new_data = []
    
    for item in datas:
        # Change all white (also shades of whites)
        # to transparent
        if item[0] > 220 and item[1] > 220 and item[2] > 220:
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(item)

    img.putdata(new_data)
    img.save(ico_path, format='ICO', sizes=[(256, 256)])
    print(f"Successfully created transparent {ico_path}")
except Exception as e:
    print(f"Error: {e}")
