from django.shortcuts import render
from django.http import HttpResponse, FileResponse
from django.core.files.storage import FileSystemStorage
from docx import Document
from docx.shared import Inches 
from docx.enum.text import WD_ALIGN_PARAGRAPH 
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Image, Table, TableStyle
from reportlab.lib.units import inch, mm
from reportlab.lib import colors 
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PyPDF2 import PdfReader
import fitz # PyMuPDF
import os
import io
import re
import unicodedata
import shutil 
from collections import defaultdict # Tekrar eden metinleri saymak için eklendi

# =========================================================================
# YENİ EKLENEN KISIM: TÜRKÇE KARAKTER DESTEĞİ İÇİN FONT KAYDI
# =========================================================================
from reportlab.pdfbase.ttfonts import TTFont # Yeni import

# Times New Roman TTF dosyasını ReportLab'e kaydet
# NOT: 'TimesNewRoman.ttf' dosyasının Reportlab'in erişebileceği bir dizinde (tercihen manage.py'nin yanında) olması gerekir.
try:
    pdfmetrics.registerFont(TTFont('TimesNewRoman', 'TimesNewRoman.ttf'))
except Exception as e:
    # Font dosyası bulunamazsa, siyah nokta hatası devam edebilir.
    # Kullanıcıya hata döndürmek yerine, varsayılan Times-Roman'a düşeriz.
    print(f"UYARI: TimesNewRoman.ttf yüklenemedi. Türkçe karakter sorunu oluşabilir. Hata: {e}")

# =========================================================================
# YUKARIDAN İTİBAREN KODUN KALANI DEVAM EDİYOR
# =========================================================================

# --- GENEL TEMİZLEME FONKSİYONU ---

def _clean_document_text_general(all_page_texts: list[str]) -> list[str]:
    """
    Sayfalar arası tekrar eden altbilgi/üstbilgi metinlerini genel bir mantıkla temizler.
    
    Args:
        all_page_texts: Her sayfadan çıkarılan ham metinlerin listesi.
        
    Returns:
        Her sayfanın temizlenmiş metin listesi.
    """
    if not all_page_texts:
        return []

    # 1. Dokümandaki tüm benzersiz kısa metin parçalarını topla
    min_len = 5 
    max_len = 100 
    repeat_threshold = 0.35 
    
    potential_junk = defaultdict(int)
    total_pages = len(all_page_texts)
    if total_pages == 0:
        return []

    for page_text in all_page_texts:
        lines = [re.sub(r'\s+', ' ', line).strip() for line in page_text.split('\n') if line.strip()]
        
        counted_on_page = set()
        for line in lines:
            if min_len <= len(line) <= max_len and len(line.split()) < 20: 
                if line not in counted_on_page:
                    potential_junk[line] += 1
                    counted_on_page.add(line)

    # 2. Tekrar eden (junk) metinleri tespit et
    junk_set = {
        text for text, count in potential_junk.items() 
        if count / total_pages > repeat_threshold 
    }
    
    # 3. Temizlenmiş metinleri oluştur
    cleaned_page_texts = []
    
    for page_text in all_page_texts:
        lines = [line.strip() for line in page_text.split('\n') if line.strip()]
        
        filtered_lines = []
        for line in lines:
            line_normalized = re.sub(r'\s+', ' ', line).strip()
            
            is_junk = False
            for junk in junk_set:
                if junk in line_normalized and len(line_normalized) < 2 * len(junk):
                    is_junk = True
                    break
            
            if line_normalized.isdigit() and len(line_normalized) <= 4:
                is_junk = True
            
            if not is_junk:
                filtered_lines.append(line_normalized)

        cleaned_page_texts.append(' '.join(filtered_lines))
        
    return cleaned_page_texts

# --- METİN ÇIKARMA FONKSİYONLARI ---

def extract_text_from_word(file_path):
    try:
        doc = Document(file_path)
        text_parts = [paragraph.text.strip() for paragraph in doc.paragraphs if paragraph.text.strip()]
        
        all_pages_temp = ['\n'.join(text_parts)] 
        cleaned_page_texts = _clean_document_text_general(all_pages_temp)
        
        return "\n\n".join(cleaned_page_texts) if cleaned_page_texts else ""
    except Exception as e:
        return str(e)

def extract_text_from_powerpoint(file_path):
    try:
        prs = Presentation(file_path)
        all_page_texts = []
        
        for slide in prs.slides:
            slide_text_parts = []
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text:
                    slide_text_parts.append('\n'.join([line.strip() for line in shape.text.split('\n') if line.strip()]))
            
            all_page_texts.append('\n'.join(slide_text_parts))
            
        cleaned_page_texts = _clean_document_text_general(all_page_texts)
                
        return "\n\n".join(cleaned_page_texts)
    except Exception as e:
        return str(e)

def extract_text_from_pdf(file_path):
    try:
        reader = PdfReader(file_path)
        all_page_texts = []
        
        for page in reader.pages:
            content = page.extract_text()
            if content:
                all_page_texts.append(content)
                
        cleaned_page_texts = _clean_document_text_general(all_page_texts)
                    
        return "\n\n".join(cleaned_page_texts)
    except Exception as e:
        return str(e)

# --- DİĞER FONKSİYONLAR (DEĞİŞMEDİ) ---

def normalize_text(raw_text: str) -> str:
    text = unicodedata.normalize('NFC', raw_text or "")
    replacements = {
        '\u00A0': ' ', '\u00AD': '', 
        '•': '-', '◦': '-', '●': '-', '▪': '-', '–': '-', '—': '-', '·': '-', '': '-',
        '“': '"', '”': '"', '‟': '"', '’': "'", '‘': "'",
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    text = re.sub(r'([.,;:!?])(?!\s)', r"\1 ", text)
    text = re.sub(r'(?<=[a-zçğıöşü])(?=[A-ZÇĞİÖŞÜ])', ' ', text)
    text = re.sub(r'(?<=[A-Za-zÇĞİÖŞÜçğıöşü])(?=\d)', ' ', text)
    text = re.sub(r'(?<=\d)(?=[A-Za-zÇĞİÖŞÜçğıöşü])', ' ', text)
    text = re.sub(r'[ \t\x0b\f\r]+', ' ', text)
    text = re.sub(r'\n\s*\n', '\n\n', text)
    return text.strip()

def extract_images_from_powerpoint(file_path):
    try:
        prs = Presentation(file_path)
        images = []
        temp_dir = os.path.join(os.path.dirname(file_path), 'temp_images')
        os.makedirs(temp_dir, exist_ok=True)
        
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    image = shape.image
                    image_bytes = image.blob
                    image_path = os.path.join(temp_dir, f'image_{len(images)}.png')
                    with open(image_path, 'wb') as f:
                        f.write(image_bytes)
                    images.append(image_path)
        return images
    except Exception as e:
        return []

def extract_images_from_word(file_path):
    """Word dosyasından resimleri çıkar"""
    try:
        doc = Document(file_path)
        images = []
        temp_dir = os.path.join(os.path.dirname(file_path), 'temp_images')
        os.makedirs(temp_dir, exist_ok=True)
        
        for rel in doc.part.rels.values():
            if "image" in rel.target_ref:
                try:
                    image_data = rel.target_part.blob
                    image_path = os.path.join(temp_dir, f'word_image_{len(images)}.png')
                    with open(image_path, 'wb') as f:
                        f.write(image_data)
                    images.append(image_path)
                except:
                    continue
        return images
    except Exception as e:
        return []

def extract_images_from_pdf(file_path):
    """PDF dosyasından resimleri çıkar"""
    try:
        doc = fitz.open(file_path)
        images = []
        temp_dir = os.path.join(os.path.dirname(file_path), 'temp_images')
        os.makedirs(temp_dir, exist_ok=True)
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            image_list = page.get_images()
            
            for img_index, img in enumerate(image_list):
                xref = img[0]
                pix = fitz.Pixmap(doc, xref)
                if pix.n - pix.alpha < 4: 
                    image_path = os.path.join(temp_dir, f'pdf_image_{page_num}_{img_index}.png')
                    pix.save(image_path)
                    images.append(image_path)
                pix = None
        doc.close()
        return images
    except Exception as e:
        return []

def empty_page(canvas, doc):
    canvas.saveState()
    canvas.restoreState()

def create_pdf_with_formatting(text, images, buffer, with_images=False):
    """PDF'i istenen akademik/minimal formatta oluşturur."""
    try:
        page_width, page_height = A4
        margin = 10 * mm 
        
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            rightMargin=margin,
            leftMargin=margin,
            topMargin=margin,
            bottomMargin=margin
        )

        usable_width = page_width - (2 * margin)
        img_width = (usable_width / 3) - (3 * mm) 
        img_height = (A4[1] - (2 * margin) - (10 * mm)) / 4 - (3 * mm)

        styles = getSampleStyleSheet()
        # Yazı tipi: Times 12pt, satır aralığı 1.15 (13.8pt), iki yana yaslı
        styles.add(ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            # BURADA TimesNewRoman kullanılır
            fontName='TimesNewRoman',
            fontSize=12,
            leading=13.8,  
            alignment=TA_JUSTIFY, 
            spaceBefore=6,
            spaceAfter=6
        ))
        normal_style = styles['CustomNormal']

        story = []
        
        paragraphs = text.split("\n\n")
        for paragraph in paragraphs:
            if paragraph.strip():
                p = Paragraph(paragraph.strip(), normal_style)
                story.append(p)
                story.append(Spacer(1, 3)) 

        if with_images and images:
            story.append(PageBreak())

            for i in range(0, len(images), 12):
                table_data = [[], [], [], []] 
                
                for row in range(4):
                    for col in range(3):
                        idx = i + row * 3 + col
                        if idx < len(images):
                            img = Image(images[idx], width=img_width, height=img_height)
                            table_data[row].append(img)
                        else:
                            table_data[row].append('')

                table = Table(
                    table_data,
                    colWidths=[img_width + 2*mm] * 3, 
                    rowHeights=[img_height + 2*mm] * 4
                )
                
                table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('TOPPADDING', (0, 0), (-1, -1), 1*mm),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 1*mm),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                    ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
                ]))
                
                story.append(table)
                
                if len(images) > i + 12:
                    story.append(PageBreak())

        doc.build(story, onFirstPage=empty_page, onLaterPages=empty_page) 
        return True
    except Exception as e:
        return str(e)

def create_word_document(text, images, with_images=False):
    """Word belgesi oluşturur ve kenar boşluklarını ayarlar."""
    try:
        doc = Document()
        
        section = doc.sections[0]
        margin_inch = 0.4 
        section.top_margin = Inches(margin_inch)
        section.bottom_margin = Inches(margin_inch)
        section.left_margin = Inches(margin_inch)
        section.right_margin = Inches(margin_inch)
        
        paragraphs = text.split("\n\n")
        for paragraph in paragraphs:
            if paragraph.strip():
                p = doc.add_paragraph(paragraph.strip())
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.style.font.name = 'Times New Roman'
                p.style.font.size = 12
        
        if with_images and images:
            doc.add_page_break()
            
            for i in range(0, len(images), 12):
                table = doc.add_table(rows=4, cols=3)
                table.style = 'Table Grid'
                
                for row in range(4):
                    for col in range(3):
                        idx = i + row * 3 + col
                        if idx < len(images):
                            cell = table.cell(row, col)
                            cell.text = f"Resim {idx + 1}"
                
                if len(images) > i + 12:
                    doc.add_page_break()
        
        return doc
    except Exception as e:
        return str(e)

# --- Django View Fonksiyonu (DEĞİŞMEDİ) ---

def home(request):
    if request.method == 'POST' and request.FILES['document']:
        uploaded_file = request.FILES['document']
        fs = FileSystemStorage()
        name = fs.save(uploaded_file.name, uploaded_file)
        file_path = fs.path(name)
        temp_dir = None
        
        try:
            with_images = request.POST.get('with_images') == 'true'
            output_format = request.POST.get('output_format', 'pdf') 
            text = ""
            images = []
            file_ext = os.path.splitext(name)[1].lower()
            
            if file_ext in ['.doc', '.docx']:
                text = extract_text_from_word(file_path)
                if with_images:
                    temp_dir = os.path.join(fs.location, 'temp_images')
                    os.makedirs(temp_dir, exist_ok=True)
                    images = extract_images_from_word(file_path)
            elif file_ext in ['.ppt', '.pptx']:
                text = extract_text_from_powerpoint(file_path)
                if with_images:
                    temp_dir = os.path.join(fs.location, 'temp_images')
                    os.makedirs(temp_dir, exist_ok=True)
                    images = extract_images_from_powerpoint(file_path)
            elif file_ext in ['.pdf']:
                text = extract_text_from_pdf(file_path)
                if with_images:
                    temp_dir = os.path.join(fs.location, 'temp_images')
                    os.makedirs(temp_dir, exist_ok=True)
                    images = extract_images_from_pdf(file_path)

            text = normalize_text(text)
            
            buffer = io.BytesIO()
            output_name = request.POST.get('output_name', '').strip()
            if not output_name:
                output_name = os.path.splitext(name)[0]
                
            if output_format == 'word':
                doc = create_word_document(text, images, with_images)
                if isinstance(doc, str): 
                    return HttpResponse(f"Word belgesi oluşturulurken hata oluştu: {doc}", status=500)
                
                doc.save(buffer)
                buffer.seek(0)
                
                return FileResponse(
                    buffer,
                    as_attachment=True,
                    filename=f"{output_name}.docx",
                    content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                )
            else: # PDF
                result = create_pdf_with_formatting(text, images, buffer, with_images)
                
                if result is True:
                    buffer.seek(0)
                    return FileResponse(
                        buffer,
                        as_attachment=True,
                        filename=f"{output_name}.pdf",
                        content_type='application/pdf'
                    )
                else:
                    return HttpResponse(f"PDF oluşturulurken hata oluştu: {result}", status=500)
            
        finally:
            fs.delete(name)
            if temp_dir and os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                except:
                    pass
                
    return render(request, 'converter/home.html')
