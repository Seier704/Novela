# -*- coding: utf-8 -*-
import os

gen_path = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela\generar_libro.py"
with open(gen_path, "r", encoding="utf-8") as f:
    content = f.read()

target = 'r_div.font.color.rgb = RGBColor(160, 145, 180)'
replacement = '''r_div.font.color.rgb = RGBColor(160, 145, 180)
        
        # Ilustracion del capitulo si existe
        if chap.get('illustration') and os.path.exists(chap['illustration']):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(4)
            r_img = p_img.add_run()
            r_img.add_picture(chap['illustration'], width=Inches(3.2))
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_after = Pt(20)
            r_cap = p_cap.add_run(chap.get('illustration_caption', 'Ilustración del Capítulo'))
            r_cap.font.name = 'Georgia'
            r_cap.font.size = Pt(9.5)
            r_cap.font.italic = True
            r_cap.font.color.rgb = RGBColor(130, 110, 150)'''

if target in content and "add_picture" not in content:
    content = content.replace(target, replacement, 1)
    with open(gen_path, "w", encoding="utf-8") as f:
        f.write(content)
    print("[OK] generar_libro.py actualizado con soporte de ilustraciones.")
else:
    print("[INFO] Ya tiene soporte o no coincide target.")
