import os
import re
from datetime import datetime
from pypdf import PdfReader
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- НАСТРОЙКИ ПУТЕЙ ---
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
automation_folder = os.path.join(desktop, "REDX_Automation")
input_dir = os.path.join(automation_folder, "Incoming_PDFs")
output_dir = os.path.join(automation_folder, "Ready_to_Print")

# Ваши ZIP-коды
MY_ZIPS = ["11223", "11234", "11214", "11229"]

logo_path = os.path.join(automation_folder, "logo.png")
if not os.path.exists(logo_path):
    logo_path = os.path.join(input_dir, "logo.png")

if not os.path.exists(output_dir): os.makedirs(output_dir)

PROCESSED_LETTERS = set()

def get_gender_title(name):
    """Добавляет Mr. или Ms. в зависимости от имени"""
    if name == "Homeowner":
        return "Homeowner"
    
    # Убираем лишние пробелы и берем первое слово (имя)
    first_name = name.split()[0].upper()
    
    # Если имя заканчивается на гласные A, IA, NA - скорее всего женщина
    if first_name.endswith(('A', 'IA', 'NA', 'Y')):
        return f"Ms. {name}"
    else:
        return f"Mr. {name}"

def find_and_fix_name(lines, index):
    """Ищет имя и склеивает разорванные буквы (W + Illa)"""
    potential_block = lines[max(0, index-8):index]
    garbage = ["NEW", "LEAD", "STATUS", "BECAME", "ON", "DATE", "PRINTED", "EXPIRED", "HOMEOWNER"]
    
    for i in range(len(potential_block)-1, -1, -1):
        candidate = potential_block[i].strip()
        if not candidate or re.search(r'\d{4}|-|:', candidate): continue
        if any(word in candidate.upper() for word in garbage): continue
        
        if len(candidate) > 1 and not candidate[0].isdigit():
            # Проверка на разрыв (W + Illa)
            if i > 0:
                prev = potential_block[i-1].strip()
                if len(prev) == 1 and prev.isupper():
                    return (prev + candidate).title()
            return candidate.title()
    return "Homeowner"

def create_letter(doc, name, address):
    # 1. ЛОГОТИП (Центр)
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.add_run().add_picture(logo_path, width=Inches(1.8))

    # Определяем обращение (Mr./Ms.)
    full_title_name = get_gender_title(name)

    # 2. АДРЕСАТ (Слева)
    doc.add_paragraph(f"\n{full_title_name}").paragraph_format.space_after = Pt(0)
    doc.add_paragraph(f"{address}").paragraph_format.space_after = Pt(2)
    
    short_addr = address.split(',')[0]
    doc.add_paragraph(f"Re: {short_addr}").paragraph_format.space_after = Pt(12)
    
    # 3. ДАТА (Справа)
    current_date = datetime.now().strftime("%B %d, %Y") 
    date_para = doc.add_paragraph(current_date)
    date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    doc.add_paragraph(f"Dear {full_title_name},")

    # 4. ТЕКСТ ПИСЬМА
    p = doc.add_paragraph()
    p.add_run(f"\nI noticed your home on {address} was recently taken off the market. ").bold = True
    p.add_run("Are you still considering selling?\n\n"
              "My name is Assiya Zhandossova, a licensed real estate agent. The market is active, "
              "and with low inventory, motivated buyers are still searching.\n\n"
              "I use professional photos, video, and detailed floor plans to create strong interest. "
              "I'd be happy to show you what can be done differently to help your home sell.\n\n"
              "Please call or text me anytime.")

    doc.add_paragraph("Best Regards,")
    
    # 5. КОНТАКТЫ (Плотный блок)
    sign = doc.add_paragraph()
    sign.paragraph_format.space_after = Pt(0)
    sign.add_run("Assiya Zhandossova").bold = True
    sign.add_run(f"\nReal Estate Salesperson | Local Expert")
    sign.add_run(f"\nPride Estates, Inc.")
    sign.add_run(f"\n3215 Quentin Road, Brooklyn, NY 11234")
    sign.add_run(f"\n347-727-3011 (Mobile) | 718-569-2888 (Office)").font.size = Pt(10)
    
    doc.add_page_break()

# --- ЗАПУСК ---
if os.path.exists(input_dir):
    for file_name in [f for f in os.listdir(input_dir) if f.endswith('.pdf')]:
        reader = PdfReader(os.path.join(input_dir, file_name))
        doc = Document()
        found_count = 0
        
        for page in reader.pages:
            text = page.extract_text()
            if not text: continue
            lines = [l.strip() for l in text.split('\n') if l.strip()]
            
            for i, line in enumerate(lines):
                # Фильтр по Бруклину и ZIP-кодам
                if "BROOKLYN, NY" in line.upper() and any(zip_code in line for zip_code in MY_ZIPS):
                    full_addr = f"{lines[i-1]}, {line}"
                    owner = find_and_fix_name(lines, i-1)
                    
                    uid = f"{owner}_{full_addr}".upper()
                    if uid not in PROCESSED_LETTERS:
                        create_letter(doc, owner, full_addr)
                        PROCESSED_LETTERS.add(uid)
                        found_count += 1

        if found_count > 0:
            doc.save(os.path.join(output_dir, f"FINAL_{file_name.replace('.pdf', '.docx')}"))
            print(f"Готово! Создано писем: {found_count}")