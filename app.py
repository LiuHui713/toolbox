"""
Toolbox - 工具箱网站
"""

import io, os, uuid
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf','pptx','ppt','docx','doc','png','jpg','jpeg','mp3','wav','mp4','avi','mov','webm','m4a','ogg','txt','md','csv'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def home():
    return render_template('home.html')


# ==================== 文字提炼 ====================
@app.route('/text')
def text_tool():
    return render_template('text.html')

@app.route('/text/extract', methods=['POST'])
def text_extract():
    """多模态文字提取 → Word 输出"""
    try:
        texts = []
        files = request.files.getlist('files')
        
        for f in files:
            if not f.filename or not allowed_file(f.filename):
                continue
            ext = f.filename.rsplit('.', 1)[1].lower()
            fname = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(f.filename))
            f.save(fname)

            try:
                if ext == 'pdf':
                    import fitz
                    doc = fitz.open(fname)
                    for page in doc:
                        texts.append(page.get_text())
                    doc.close()
                elif ext in ('pptx', 'ppt'):
                    from pptx import Presentation
                    prs = Presentation(fname)
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                texts.append(shape.text_frame.text)
                elif ext in ('docx', 'doc'):
                    import docx
                    doc = docx.Document(fname)
                    texts.extend([p.text for p in doc.paragraphs])
                elif ext in ('png', 'jpg', 'jpeg'):
                    from rapidocr_onnxruntime import RapidOCR
                    ocr = RapidOCR()
                    result, _ = ocr(fname)
                    if result:
                        texts.extend([line[1] for line in result])
                elif ext in ('mp3', 'wav', 'm4a', 'ogg'):
                    import subprocess, tempfile
                    # Use faster-whisper for audio
                    from faster_whisper import WhisperModel
                    model = WhisperModel("base", device="cpu", compute_type="int8")
                    segments, _ = model.transcribe(fname)
                    texts.extend([s.text for s in segments])
                elif ext in ('mp4', 'avi', 'mov', 'webm'):
                    import subprocess, tempfile
                    # Extract audio from video then transcribe
                    audio_path = fname + '.wav'
                    subprocess.run(['ffmpeg', '-i', fname, '-vn', '-acodec', 'pcm_s16le', '-y', audio_path],
                                   capture_output=True, timeout=120)
                    from faster_whisper import WhisperModel
                    model = WhisperModel("base", device="cpu", compute_type="int8")
                    segments, _ = model.transcribe(audio_path)
                    texts.extend([s.text for s in segments])
                    os.remove(audio_path)
                elif ext in ('txt', 'md', 'csv'):
                    with open(fname, 'r', encoding='utf-8', errors='ignore') as fh:
                        texts.append(fh.read())
            except Exception as e:
                texts.append(f"[提取失败: {f.filename} - {str(e)}]")
            finally:
                os.remove(fname)

        if not texts:
            return jsonify({"error": "未能提取到文字"}), 400

        # Generate Word document
        from docx import Document as DocxDoc
        doc = DocxDoc()
        doc.add_heading('文字提取结果', 0)
        for i, t in enumerate(texts):
            if t.strip():
                doc.add_paragraph(t.strip())
        
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         as_attachment=True, download_name='extracted.docx')

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ==================== PPT生成 ====================
@app.route('/ppt')
def ppt_tool():
    return render_template('ppt.html')

@app.route('/ppt/generate', methods=['POST'])
def ppt_generate():
    try:
        content = request.form.get('content', '')
        title = request.form.get('title', '演示文稿')
        if not content.strip():
            return jsonify({"error": "请输入内容"}), 400

        from pptx import Presentation
        from pptx.util import Inches, Pt
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        # Title slide
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = title
        if slide.placeholders[1].has_text_frame:
            slide.placeholders[1].text = "AI 自动生成"

        # Content slides
        paragraphs = [p.strip() for p in content.split('\n\n') if p.strip()]
        for para in paragraphs:
            lines = para.split('\n')
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = lines[0][:80]
            body = slide.placeholders[1].text_frame
            for line in lines[1:8]:
                p = body.add_paragraph()
                p.text = line[:200]
                p.level = 0

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.presentationml.document',
                         as_attachment=True, download_name='generated.pptx')
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ==================== 音频合成 ====================
@app.route('/audio')
def audio_tool():
    return render_template('audio.html')


# ==================== 海报生成 ====================
@app.route('/poster')
def poster_tool():
    return render_template('poster.html')

@app.route('/poster/generate', methods=['POST'])
def poster_generate():
    try:
        text = request.form.get('text', '')
        style = request.form.get('style', 'modern')
        bg_color = request.form.get('bg_color', '#1a1a2e')
        text_color = request.form.get('text_color', '#ffffff')

        from PIL import Image, ImageDraw, ImageFont
        img = Image.new('RGB', (800, 1200), bg_color)
        draw = ImageDraw.Draw(img)

        try:
            font_large = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 72)
            font_small = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 32)
        except:
            font_large = ImageFont.load_default()
            font_small = ImageFont.load_default()

        # Draw text centered
        lines = text.split('\n')
        y = 200
        for i, line in enumerate(lines):
            font = font_large if i == 0 else font_small
            bbox = draw.textbbox((0, 0), line, font=font)
            tw = bbox[2] - bbox[0]
            x = (800 - tw) // 2
            draw.text((x, y), line, fill=text_color, font=font)
            y += bbox[3] - bbox[1] + 40

        buf = io.BytesIO()
        img.save(buf, format='PNG')
        buf.seek(0)
        return send_file(buf, mimetype='image/png', as_attachment=True, download_name='poster.png')
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
