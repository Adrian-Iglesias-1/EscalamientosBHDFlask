import os
from PIL import Image, ImageDraw, ImageFont

def create_icon():
    """Crear icono simple para EscalamientosApp"""
    # Crear imagen 256x256
    img = Image.new('RGBA', (256, 256), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    
    # Fondo circular azul corporativo
    blue_color = (51, 102, 153)  # Azul corporativo
    draw.ellipse([10, 10, 246, 246], fill=blue_color, outline=(255, 255, 255, 200), width=5)
    
    # Texto "EA" en blanco
    text = "EA"
    try:
        # Intentar usar fuente del sistema
        font = ImageFont.truetype("arial.ttf", 80)
    except:
        font = ImageFont.load_default()
    
    # Calcular posición centrada
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    x = (256 - text_width) // 2
    y = (256 - text_height) // 2 - 20
    
    # Dibujar texto
    draw.text((x, y), text, fill=(255, 255, 255), font=font)
    
    # Texto pequeño "ATM"
    try:
        font_small = ImageFont.truetype("arial.ttf", 40)
    except:
        font_small = ImageFont.load_default()
    
    text_small = "ATM"
    bbox_small = draw.textbbox((0, 0), text_small, font=font_small)
    text_width_small = bbox_small[2] - bbox_small[0]
    x_small = (256 - text_width_small) // 2
    y_small = y + text_height + 10
    
    draw.text((x_small, y_small), text_small, fill=(200, 200, 200), font=font_small)
    
    # Guardar en diferentes tamaños para .ico
    sizes = [256, 128, 64, 48, 32, 16]
    images = []
    for size in sizes:
        resized = img.resize((size, size), Image.LANCZOS)
        # Convertir a RGB para .ico
        if size != 256:
            rgb_img = Image.new('RGB', (size, size), (255, 255, 255))
            rgb_img.paste(resized, mask=resized.split()[3] if resized.mode == 'RGBA' else None)
            images.append(rgb_img)
        else:
            images.append(resized.convert('RGB'))
    
    # Guardar icono
    static_dir = os.path.join(os.path.dirname(__file__), 'static')
    os.makedirs(static_dir, exist_ok=True)
    icon_path = os.path.join(static_dir, 'icon.ico')
    
    images[0].save(icon_path, format='ICO', sizes=[(s, s) for s in sizes])
    print(f"Icono creado: {icon_path}")
    return icon_path

if __name__ == '__main__':
    try:
        create_icon()
    except ImportError:
        print("Pillow no instalado. Instalando...")
        import subprocess
        subprocess.run(['pip', 'install', 'Pillow'], check=True)
        create_icon()
