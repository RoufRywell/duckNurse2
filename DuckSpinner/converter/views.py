from django.shortcuts import render
from django.http import HttpResponse, FileResponse
from django.core.files.storage import FileSystemStorage
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Image, Table, TableStyle
from reportlab.lib.units import inch
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PyPDF2 import PdfReader
import os
import io

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
    
def create_pdf_with_formatting(text, images, buffer, with_images=False):
    try:
        page_width, page_height = letter
        margin = 15
        
        doc = SimpleDocTemplate(
            buffer,
            pagesize=letter,
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
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontSize=12,
            leading=14,
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
                
                table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('TOPPADDING', (0, 0), (-1, -1), 2),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                    ('LEFTPADDING', (0, 0), (-1, -1), 2),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 2),
                ]))
                
                story.append(table)
                
                if len(images) > i + 12:
                    story.append(PageBreak())

        doc.build(story)
        return True
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
            text = ""
            images = []
            file_ext = os.path.splitext(name)[1].lower()
            
            if file_ext in ['.doc', '.docx']:
                text = extract_text_from_word(file_path)
            elif file_ext in ['.ppt', '.pptx']:
                text = extract_text_from_powerpoint(file_path)
                if with_images:
                    temp_dir = os.path.join(fs.location, 'temp_images')
                    os.makedirs(temp_dir, exist_ok=True)
                    images = extract_images_from_powerpoint(file_path)
            elif file_ext in ['.pdf']:
                text = extract_text_from_pdf(file_path)
            
            buffer = io.BytesIO()
            result = create_pdf_with_formatting(text, images, buffer, with_images)
            
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
                
        except Exception as e:
            if os.path.exists(file_path):
                fs.delete(name)
            return HttpResponse(f"İşlem sırasında hata oluştu: {str(e)}", status=500)
    
    return render(request, 'converter/home.html')