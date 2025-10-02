from django.shortcuts import render
from django.http import HttpResponse, FileResponse
from django.core.files.storage import FileSystemStorage
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Image, Table, TableStyle
from reportlab.lib.units import inch, mm
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PyPDF2 import PdfReader
import fitz  # PyMuPDF
import os
import io
import re
import unicodedata

def extract_text_from_word(file_path):
    try:
        doc = Document(file_path)
        text = []
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                text.append(paragraph.text.strip())
        return "\n\n".join(text)
    except Exception as e:
        return str(e)

def extract_text_from_powerpoint(file_path):
    try:
        prs = Presentation(file_path)
        text = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text:
                    cleaned = "\n".join([line.strip() for line in shape.text.split("\n") if line.strip()])
                    if cleaned:
                        text.append(cleaned)
        return "\n\n".join(text)
    except Exception as e:
        return str(e)

def extract_text_from_pdf(file_path):
    try:
        reader = PdfReader(file_path)
        text = []
        for page in reader.pages:
            content = page.extract_text()
            if content.strip():
                content_lines = [line.strip() for line in content.split('\n') if line.strip() and not line.strip().isdigit()]
                text.append('\n'.join(content_lines))
        return "\n\n".join(text)
    except Exception as e:
        return str(e)

def normalize_text(raw_text: str) -> str:
    # Unicode normalizasyonu (Türkçe/İngilizce özel karakterler düzgünleşsin)
    text = unicodedata.normalize('NFC', raw_text or "")

    # Yaygın problemli karakterler ve mermiler → ASCII karşılıkları
    replacements = {
        '\u00A0': ' ',   # no-break space
        '\u00AD': '',    # soft hyphen
        '•': '-', '◦': '-', '●': '-', '▪': '-', '–': '-', '—': '-', '·': '-', '': '-',
        '“': '"', '”': '"', '‟': '"', '’': "'", '‘': "'",
    }
    for k, v in replacements.items():
        text = text.replace(k, v)

    # Noktalama sonrası boşluk ekle (eksikse)
    text = re.sub(r'([.,;:!?])(?!\s)', r"\1 ", text)

    # Bitişik İngilizce kelimeleri yumuşak ayır: lower→Upper geçişine boşluk
    text = re.sub(r'(?<=[a-zçğıöşü])(?=[A-ZÇĞİÖŞÜ])', ' ', text)
    # Harf-sayı ve sayı-harf geçişlerinde boşluk ekle
    text = re.sub(r'(?<=[A-Za-zÇĞİÖŞÜçğıöşü])(?=\d)', ' ', text)
    text = re.sub(r'(?<=\d)(?=[A-Za-zÇĞİÖŞÜçğıöşü])', ' ', text)

    # Çoklu boşlukları tek boşluğa indir ve satırları sadeleştir
    text = re.sub(r'[ \t\x0b\f\r]+', ' ', text)
    # Paragrafları korumak için çift yeni satırı standardize et
    text = re.sub(r'\n\s*\n', '\n\n', text)
    text = text.strip()
    return text

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
                if pix.n - pix.alpha < 4:  # GRAY or RGB
                    image_path = os.path.join(temp_dir, f'pdf_image_{page_num}_{img_index}.png')
                    pix.save(image_path)
                    images.append(image_path)
                pix = None
        doc.close()
        return images
    except Exception as e:
        return []
    
def create_pdf_with_formatting(text, images, buffer, with_images=False):
    try:
        # Sayfa biçimi: A4, kenar boşlukları 20 mm
        page_width, page_height = A4
        margin = 20 * mm
        
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            rightMargin=margin,
            leftMargin=margin,
            topMargin=margin,
            bottomMargin=margin
        )

        # Kullanılabilir alan
        usable_width = page_width - (2 * margin)
        usable_height = page_height - (2 * margin)

        # Resim boyutları (sayfa başına 12 görsel: 3x4)
        img_width = (usable_width / 3) - 6  # 3 sütun
        img_height = (usable_height / 4) - 6  # 4 satır

        styles = getSampleStyleSheet()
        # Yazı tipi: Times 12pt, satır aralığı ~1.15, iki yana yaslı
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontName='Times-Roman',
            fontSize=12,
            leading=14,  # ~1.15 * 12
            alignment=TA_JUSTIFY,
            spaceBefore=6,
            spaceAfter=6
        )

        story = []
        
        # Metin
        paragraphs = text.split("\n\n")
        for paragraph in paragraphs:
            if paragraph.strip():
                p = Paragraph(paragraph.strip(), normal_style)
                story.append(p)

        # Resimleri sadece with_images=True ise ekle
        if with_images and images:
            story.append(PageBreak())

            # 12'li gruplar (3x4)
            for i in range(0, len(images), 12):
                chunk = images[i:i + 12]
                table_data = [[], [], [], []]  # 4 satır
                
                # Her satır için 3 resim
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
                    colWidths=[img_width] * 3,
                    rowHeights=[img_height] * 4
                )
                
                # İnce kenarlık ve hafif iç boşluklarla görsel ızgarası
                table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('TOPPADDING', (0, 0), (-1, -1), 3),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                    ('LEFTPADDING', (0, 0), (-1, -1), 3),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                    ('GRID', (0, 0), (-1, -1), 0.25, '#BBBBBB'),
                    ('BOX', (0, 0), (-1, -1), 0.25, '#BBBBBB'),
                ]))
                
                story.append(table)
                
                if len(images) > i + 12:
                    story.append(PageBreak())

        doc.build(story)
        return True
    except Exception as e:
        return str(e)

def create_word_document(text, images, with_images=False):
    """Word belgesi oluştur"""
    try:
        doc = Document()
        
        # Metin ekle
        paragraphs = text.split("\n\n")
        for paragraph in paragraphs:
            if paragraph.strip():
                p = doc.add_paragraph(paragraph.strip())
                p.alignment = 1  # Justify
        
        # Resimleri ekle
        if with_images and images:
            doc.add_page_break()
            
            # 12'li gruplar halinde resimleri ekle
            for i in range(0, len(images), 12):
                chunk = images[i:i + 12]
                
                # 3x4 tablo oluştur
                table = doc.add_table(rows=4, cols=3)
                table.style = 'Table Grid'
                
                for row in range(4):
                    for col in range(3):
                        idx = i + row * 3 + col
                        if idx < len(images):
                            cell = table.cell(row, col)
                            cell.text = f"Resim {idx + 1}"
                            # Resim ekleme burada yapılabilir
                
                if len(images) > i + 12:
                    doc.add_page_break()
        
        return doc
    except Exception as e:
        return str(e)

def home(request):
    if request.method == 'POST' and request.FILES['document']:
        uploaded_file = request.FILES['document']
        fs = FileSystemStorage()
        name = fs.save(uploaded_file.name, uploaded_file)
        file_path = fs.path(name)
        temp_dir = None
        
        try:
            with_images = request.POST.get('with_images') == 'true'
            output_format = request.POST.get('output_format', 'pdf')  # pdf veya word
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

            # Metni normalize et: siyah nokta/mermi karakterleri, NBSP, soft hyphen ve bitişik kelimeler
            text = normalize_text(text)
            
            if output_format == 'word':
                # Word belgesi oluştur
                doc = create_word_document(text, images, with_images)
                if isinstance(doc, str):  # Hata mesajı
                    return HttpResponse(f"Word belgesi oluşturulurken hata oluştu: {doc}", status=500)
                
                # Word belgesini kaydet
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                output_name = request.POST.get('output_name', '').strip()
                if not output_name:
                    output_name = os.path.splitext(name)[0]
                
                return FileResponse(
                    buffer,
                    as_attachment=True,
                    filename=f"{output_name}.docx",
                    content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                )
            else:
                # PDF oluştur
                buffer = io.BytesIO()
                result = create_pdf_with_formatting(text, images, buffer, with_images)
                
                if result is True:
                    output_name = request.POST.get('output_name', '').strip()
                    if not output_name:
                        output_name = os.path.splitext(name)[0]
                    
                    buffer.seek(0)
                    return FileResponse(
                        buffer,
                        as_attachment=True,
                        filename=f"{output_name}.pdf",
                        content_type='application/pdf'
                    )
                else:
                    return HttpResponse(f"PDF oluşturulurken hata oluştu: {result}", status=500)
            
            fs.delete(name)
            
            if images:
                for img_path in images:
                    try:
                        if os.path.exists(img_path):
                            os.remove(img_path)
                    except:
                        pass
                
                if temp_dir and os.path.exists(temp_dir):
                    try:
                        import shutil
                        shutil.rmtree(temp_dir, ignore_errors=True)
                    except:
                        pass
                
        except Exception as e:
            if os.path.exists(file_path):
                fs.delete(name)
            return HttpResponse(f"İşlem sırasında hata oluştu: {str(e)}", status=500)
    
    return render(request, 'converter/home.html')