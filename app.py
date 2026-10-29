import sys, os, mimetypes, platform, sqlite3, subprocess, uuid, io, base64, shutil, traceback
from datetime import datetime
from flask import (Flask, render_template, request, redirect, url_for, flash, send_from_directory, abort, jsonify)
from werkzeug.utils import secure_filename
from PIL import Image, ImageEnhance
from pypdf import PdfReader, PdfWriter
import fitz  # PyMuPDF
import qrcode

try:
    import win32print
    import win32ui
    import win32con
    from PIL import ImageWin
    WINDOWS_PRINT_AVAILABLE = True
except ImportError:
    WINDOWS_PRINT_AVAILABLE = False

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
    TEMPLATE_DIR = os.path.join(sys._MEIPASS, 'templates')
else:
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))
    TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')

UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
DB_PATH    = os.path.join(BASE_DIR, 'print_jobs.db')
ALLOWED_EXT = {'pdf', 'png', 'jpg', 'jpeg'}

app = Flask(__name__, template_folder=TEMPLATE_DIR)
app.config['SECRET_KEY']        = os.environ.get('SECRET_KEY', 'dev-key-change-me')
app.config['UPLOAD_FOLDER']     = UPLOAD_DIR
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024
os.makedirs(UPLOAD_DIR, exist_ok=True)

def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("CREATE TABLE IF NOT EXISTS print_jobs (id INTEGER PRIMARY KEY AUTOINCREMENT, stored_name TEXT NOT NULL, original_name TEXT NOT NULL, printer TEXT, copies INTEGER DEFAULT 1, paper_size TEXT, orientation TEXT, duplex INTEGER DEFAULT 0, status TEXT, notes TEXT, created_at TEXT)")
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(print_jobs)")
    columns = [info[1] for info in cursor.fetchall()]
    
    if 'watermark' not in columns: conn.execute("ALTER TABLE print_jobs ADD COLUMN watermark TEXT")
    if 'page_range' not in columns: conn.execute("ALTER TABLE print_jobs ADD COLUMN page_range TEXT")
    if 'final_file' not in columns: conn.execute("ALTER TABLE print_jobs ADD COLUMN final_file TEXT")
    
    conn.commit(); conn.close()

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def allowed_file(name):
    return '.' in name and name.rsplit('.', 1)[1].lower() in ALLOWED_EXT

def list_printers():
    system = platform.system()
    printers = []
    if system == 'Windows':
        try:
            import win32print
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            for p in win32print.EnumPrinters(flags):
                name = p[2]
                if not any(x in name for x in ["OneNote", "PDF", "XPS", "Fax"]):
                    printers.append(name)
        except Exception: pass
    return printers or ['Default']

def process_image_for_print(img_path, settings):
    """اعمال فیلترها، آینه و QR Code روی عکس قبل از چاپ"""
    img = Image.open(img_path)
    
    # آینه
    if settings.get('flip_h'): img = img.transpose(Image.FLIP_LEFT_RIGHT)
    if settings.get('flip_v'): img = img.transpose(Image.FLIP_TOP_BOTTOM)
    
    # پردازش رنگ
    if settings.get('grayscale'):
        img = img.convert('L').convert('RGB')
    
    brightness = float(settings.get('brightness', 1.0))
    contrast = float(settings.get('contrast', 1.0))
    
    if brightness != 1.0:
        img = ImageEnhance.Brightness(img).enhance(brightness)
    if contrast != 1.0:
        img = ImageEnhance.Contrast(img).enhance(contrast)
        
    # درج QR Code
    qr_text = settings.get('qr_text', '').strip()
    if qr_text:
        qr = qrcode.QRCode(box_size=10, border=1)
        qr.add_data(qr_text)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white").convert('RGB')
        
        # تغییر اندازه QR به ۲۰ میلی‌متر (تقریبا 236 پیکسل در 300dpi)
        qr_size = 236
        qr_img = qr_img.resize((qr_size, qr_size), Image.LANCZOS)
        
        # قرار دادن در پایین سمت راست با فاصله ۱۰ میلی‌متر (118 پیکسل)
        margin = 118
        x = img.width - qr_size - margin
        y = img.height - qr_size - margin
        
        if x > 0 and y > 0:
            img.paste(qr_img, (x, y))
            
    return img

def native_print_windows(filepath, printer_name, copies, img_settings=None, pdf_pages=""):
    if not WINDOWS_PRINT_AVAILABLE:
        return False, 'چاپ بومی فقط در ویندوز پشتیبانی می‌شود'
        
    if any(x in printer_name for x in ["OneNote", "PDF", "XPS"]):
        return False, 'لطفاً یک پرینتر فیزیکی انتخاب کنید.'

    try:
        if not printer_name or printer_name == 'Default':
            printer_name = win32print.GetDefaultPrinter()

        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        
        dpi_x = hdc.GetDeviceCaps(win32con.LOGPIXELSX)
        dpi_y = hdc.GetDeviceCaps(win32con.LOGPIXELSY)
        
        printable_width = hdc.GetDeviceCaps(win32con.HORZRES)
        printable_height = hdc.GetDeviceCaps(win32con.VERTRES)
        
        hdc.StartDoc(os.path.basename(filepath))
        
        for _ in range(copies):
            if filepath.lower().endswith('.pdf'):
                doc = fitz.open(filepath)
                
                # مدیریت محدوده صفحات
                pages_to_print = []
                if pdf_pages:
                    for part in pdf_pages.split(','):
                        if '-' in part:
                            start, end = map(int, part.split('-'))
                            pages_to_print.extend(range(start-1, end))
                        else:
                            pages_to_print.append(int(part)-1)
                else:
                    pages_to_print = range(len(doc))
                    
                for page_num in pages_to_print:
                    if 0 <= page_num < len(doc):
                        page = doc.load_page(page_num)
                        pix = page.get_pixmap(dpi=300)
                        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        
                        # اعمال تنظیمات رنگ روی صفحات PDF هم
                        if img_settings:
                            if img_settings.get('grayscale'): img = img.convert('L').convert('RGB')
                            b = float(img_settings.get('brightness', 1.0))
                            c = float(img_settings.get('contrast', 1.0))
                            if b != 1.0: img = ImageEnhance.Brightness(img).enhance(b)
                            if c != 1.0: img = ImageEnhance.Contrast(img).enhance(c)
                        
                        img_ratio = img.width / img.height
                        paper_ratio = printable_width / printable_height
                        
                        if img_ratio > paper_ratio:
                            draw_w = printable_width
                            draw_h = int(printable_width / img_ratio)
                        else:
                            draw_h = printable_height
                            draw_w = int(printable_height * img_ratio)
                            
                        x_pos = (printable_width - draw_w) // 2
                        y_pos = (printable_height - draw_h) // 2
                        
                        hdc.StartPage()
                        dib = ImageWin.Dib(img)
                        dib.draw(hdc.GetHandleOutput(), (x_pos, y_pos, x_pos + draw_w, y_pos + draw_h))
                        hdc.EndPage()
            else:
                img = process_image_for_print(filepath, img_settings)
                if img_settings:
                    w_mm = float(img_settings.get('img_width', 50))
                    h_mm = float(img_settings.get('img_height', 50))
                    top_mm = float(img_settings.get('pos_top', 10))
                    left_mm = float(img_settings.get('pos_left', 10))
                    
                    w_px = int(w_mm * dpi_x / 25.4)
                    h_px = int(h_mm * dpi_y / 25.4)
                    top_px = int(top_mm * dpi_y / 25.4)
                    left_px = int(left_mm * dpi_x / 25.4)
                    
                    img = img.resize((w_px, h_px), Image.LANCZOS)
                    x1, y1 = left_px, top_px
                    x2, y2 = left_px + w_px, top_px + h_px
                else:
                    x1, y1 = 0, 0
                    x2, y2 = printable_width, printable_height
                    
                hdc.StartPage()
                dib = ImageWin.Dib(img)
                dib.draw(hdc.GetHandleOutput(), (x1, y1, x2, y2))
                hdc.EndPage()
                
        hdc.EndDoc()
        hdc.DeleteDC()
        return True, 'چاپ با موفقیت به پرینتر ارسال شد.'
    except Exception as e:
        traceback.print_exc()
        return False, f'خطا در چاپ: {str(e)}'

@app.route('/')
def index():
    conn = db()
    docs = conn.execute('SELECT * FROM print_jobs ORDER BY created_at DESC').fetchall()
    stats = {'total': len(docs), 'sent': len([d for d in docs if d['status'] == 'sent']), 'failed': len([d for d in docs if d['status'] == 'failed'])}
    conn.close()
    return render_template('index.html', docs=docs, stats=stats)

@app.route('/upload', methods=['POST'])
def upload():
    try:
        files = request.files.getlist('files')
        if not files or all(f.filename == '' for f in files):
            flash('فایلی انتخاب نشده است', 'error')
            return redirect(url_for('index'))
        
        saved_paths = []
        first_name = "مجموعه_فایل‌ها.pdf"
        for f in files:
            if f and allowed_file(f.filename):
                original = secure_filename(f.filename)
                stored = f'{uuid.uuid4().hex}.{original.rsplit(".", 1)[1].lower()}'
                path = os.path.join(app.config['UPLOAD_FOLDER'], stored)
                f.save(path)
                saved_paths.append(stored)
                if len(saved_paths) == 1: first_name = original

        if not saved_paths:
            flash('فرمت فایل پشتیبانی نمی‌شود.', 'error')
            return redirect(url_for('index'))

        is_single_image = len(saved_paths) == 1 and saved_paths[0].lower().endswith(('.png', '.jpg', '.jpeg'))
        
        if is_single_image:
            final_pdf_name = saved_paths[0]
        else:
            final_pdf_name = f'{uuid.uuid4().hex}.pdf'
            final_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], final_pdf_name)
            writer = PdfWriter()
            for path in saved_paths:
                full_path = os.path.join(app.config['UPLOAD_FOLDER'], path)
                ext = full_path.rsplit('.', 1)[-1].lower()
                if ext == 'pdf':
                    for page in PdfReader(full_path).pages: writer.add_page(page)
                else:
                    img = Image.open(full_path).convert('RGB')
                    img_pdf = io.BytesIO(); img.save(img_pdf, format='PDF'); img_pdf.seek(0)
                    for page in PdfReader(img_pdf).pages: writer.add_page(page)
            with open(final_pdf_path, "wb") as f: writer.write(f)

        conn = db()
        conn.execute('INSERT INTO print_jobs (stored_name, original_name, status, created_at, final_file) VALUES (?,?,?,?,?)', 
                     ("|".join(saved_paths), first_name if len(saved_paths)==1 else f"{len(saved_paths)} فایل ادغام شده", 'pending', datetime.now().isoformat(), final_pdf_name))
        conn.commit(); conn.close()
        
        flash('فایل با موفقیت آماده شد.', 'success')
        return redirect(url_for('index'))
    except Exception as e:
        flash(f'خطای سیستمی در آپلود: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/get_image_dims/<int:job_id>')
def get_image_dims(job_id):
    conn = db()
    job = conn.execute('SELECT * FROM print_jobs WHERE id=?', (job_id,)).fetchone()
    conn.close()
    if not job: return jsonify({'error': 'not found'}), 404
    
    filename = os.path.basename(job['stored_name'].split('|')[0])
    path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if filename.lower().endswith(('.png', '.jpg', '.jpeg')) and os.path.exists(path):
        try:
            img = Image.open(path)
            dpi = img.info.get('dpi', (96, 96))[0]
            if not dpi or dpi <= 0: dpi = 96
            w_mm = (img.width / dpi) * 25.4
            h_mm = (img.height / dpi) * 25.4
            return jsonify({'width': round(w_mm, 1), 'height': round(h_mm, 1)})
        except Exception:
            pass
    return jsonify({'width': 100, 'height': 100})

@app.route('/print/<int:job_id>', methods=['GET', 'POST'])
def print_job(job_id):
    conn = db()
    job = conn.execute('SELECT * FROM print_jobs WHERE id=?', (job_id,)).fetchone()
    conn.close()
    if not job: abort(404)
    
    original_file = os.path.basename(job['stored_name'].split('|')[0])
    original_path = os.path.join(app.config['UPLOAD_FOLDER'], original_file)
    is_image = original_file.lower().endswith(('.png', '.jpg', '.jpeg'))
    
    printers = list_printers()
    
    if request.method == 'POST':
        try:
            data = request.json
            printer = data.get('printer', '')
            copies = max(1, int(data.get('copies', 1)))
            
            img_settings = {
                'img_width': data.get('img_width', 50),
                'img_height': data.get('img_height', 50),
                'pos_top': data.get('pos_top', 10),
                'pos_left': data.get('pos_left', 10),
                'flip_h': data.get('flip_h', False),
                'flip_v': data.get('flip_v', False),
                'grayscale': data.get('grayscale', False),
                'brightness': data.get('brightness', 1.0),
                'contrast': data.get('contrast', 1.0),
                'qr_text': data.get('qr_text', '')
            }
            
            pdf_pages = data.get('pdf_pages', '')
            
            ok, msg = native_print_windows(original_path, printer, copies, img_settings, pdf_pages)
            
            conn = db()
            conn.execute('UPDATE print_jobs SET printer=?, copies=?, status=?, notes=? WHERE id=?', 
                         (printer, copies, 'sent' if ok else 'failed', msg, job_id))
            conn.commit(); conn.close()
            
            return jsonify({'success': ok, 'message': msg})
        except Exception as e:
            return jsonify({'success': False, 'message': f'خطای سرور: {str(e)}'}), 500

    return render_template('print_settings.html', job=job, printers=printers, is_image=is_image)

@app.route('/raw_file/<int:job_id>')
def raw_file(job_id):
    conn = db()
    job = conn.execute('SELECT * FROM print_jobs WHERE id=?', (job_id,)).fetchone()
    conn.close()
    if not job: abort(404)
    filename = os.path.basename(job['stored_name'].split('|')[0])
    if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
        return send_from_directory(app.config['UPLOAD_FOLDER'], filename)
    return abort(404)

@app.route('/file/<int:job_id>')
def view_file(job_id):
    conn = db()
    job = conn.execute('SELECT * FROM print_jobs WHERE id=?', (job_id,)).fetchone()
    conn.close()
    if not job: abort(404)
    filename = os.path.basename(job['final_file'] or job['stored_name'])
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, mimetype='application/pdf', as_attachment=False, download_name=f"{job['original_name']}.pdf")

@app.route('/delete/<int:job_id>', methods=['POST'])
def delete_job(job_id):
    conn = db()
    job = conn.execute('SELECT * FROM print_jobs WHERE id=?', (job_id,)).fetchone()
    if job:
        try: 
            if job['stored_name']:
                for p in job['stored_name'].split('|'):
                    bname = os.path.basename(p)
                    if os.path.exists(os.path.join(app.config['UPLOAD_FOLDER'], bname)): os.remove(os.path.join(app.config['UPLOAD_FOLDER'], bname))
            if job['final_file']:
                bname = os.path.basename(job['final_file'])
                fpath = os.path.join(app.config['UPLOAD_FOLDER'], bname)
                if os.path.exists(fpath): os.remove(fpath)
        except OSError: pass
        conn.execute('DELETE FROM print_jobs WHERE id=?', (job_id,))
        conn.commit()
    conn.close()
    flash('فایل حذف شد', 'success')
    return redirect(url_for('index'))

init_db()
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
