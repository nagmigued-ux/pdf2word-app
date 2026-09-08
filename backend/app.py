"""
واجهة برمجية (API) بـ Flask للتحويل الثنائي الاتجاه بين PDF و Word.

نقاط النهاية:
- GET  /api/health   : فحص حالة الخادم
- POST /api/inspect  : فحص مسبق لملف PDF (هل يحتوي نصًا حقيقيًا أم أنه ممسوح ضوئيًا؟)
- POST /api/convert  : يستقبل ملف PDF أو Word ويعيد الملف المحوَّل جاهزًا للتنزيل
                       (يُكتشف اتجاه التحويل تلقائيًا من امتداد الملف المرفوع)
"""

import os
import shutil
import tempfile
import time
import threading
from pathlib import Path

# LibreOffice لا يتحمّل عددًا كبيرًا من عمليات التحويل المتزامنة على نفس الخادم؛
# هذا القفل يحدّ من التحويلات المتوازية لتفادي إنهاك الموارد أو تعطّل العملية.
MAX_CONCURRENT_CONVERSIONS = int(os.environ.get("MAX_CONCURRENT_CONVERSIONS", "2"))
CONVERSION_SEMAPHORE = threading.Semaphore(MAX_CONCURRENT_CONVERSIONS)

from flask import Flask, request, jsonify, send_file, after_this_request, send_from_directory
from werkzeug.utils import secure_filename

from converter import (
    convert_pdf_to_docx,
    convert_docx_to_pdf,
    ConversionError,
    ConversionTimeoutError,
    EncryptedPdfError,
    InvalidPdfError,
    NeedsOcrError,
    inspect_pdf,
    new_job_dir,
    resolve_ocr_lang,
    build_ocr_lang_string,
    OCR_LANGUAGE_CHOICES,
)

WORD_EXTENSIONS = (".docx", ".doc")
PDF_EXTENSION = ".pdf"
ALLOWED_EXTENSIONS = (PDF_EXTENSION,) + WORD_EXTENSIONS

WORD_MIMETYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PDF_MIMETYPE = "application/pdf"

MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50 ميجابايت كحد أقصى لكل ملف
BASE_TMP_DIR = os.path.join(tempfile.gettempdir(), "pdf2word_jobs")
JOB_TTL_SECONDS = 60 * 30  # حذف مجلدات العمل بعد 30 دقيقة تلقائيًا
FRONTEND_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "frontend")

os.makedirs(BASE_TMP_DIR, exist_ok=True)

app = Flask(__name__, static_folder=FRONTEND_DIR, static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


@app.get("/")
def index():
    return send_from_directory(FRONTEND_DIR, "index.html")


def cleanup_old_jobs():
    """يزيل مجلدات المهام القديمة لتفادي امتلاء القرص بملفات المستخدمين المؤقتة."""
    now = time.time()
    try:
        for entry in os.scandir(BASE_TMP_DIR):
            if entry.is_dir() and now - entry.stat().st_mtime > JOB_TTL_SECONDS:
                shutil.rmtree(entry.path, ignore_errors=True)
    except FileNotFoundError:
        pass


def schedule_cleanup(job_dir: str, delay: int = 300):
    """يحذف مجلد مهمة محددة بعد إرسال الملف للمستخدم بوقت كافٍ للتنزيل."""

    def _remove():
        shutil.rmtree(job_dir, ignore_errors=True)

    timer = threading.Timer(delay, _remove)
    timer.daemon = True
    timer.start()


SUPPORTED_UI_LANGS = {"ar", "en", "fr", "es", "de"}


def get_ui_lang() -> str:
    lang = (request.form.get("lang") or request.args.get("lang") or "en").lower()
    return lang if lang in SUPPORTED_UI_LANGS else "en"


@app.get("/api/health")
def health():
    return jsonify(status="ok")


@app.get("/api/ocr-languages")
def ocr_languages():
    """قائمة لغات المستند المدعومة صراحةً للتعرف الضوئي (OCR)، لعرضها للمستخدم
    ليختار لغة الملف الممسوح ضوئيًا الفعلية (بغض النظر عن لغة واجهة الموقع)."""
    return jsonify(languages=[{"code": code, "name": name} for code, name in OCR_LANGUAGE_CHOICES])


@app.errorhandler(413)
def handle_too_large(_exc):
    return jsonify(error="الملف أكبر من الحد المسموح به (50 ميغابايت).", code="file_too_large"), 413


@app.post("/api/convert")
def convert():
    """
    نقطة تحويل ثنائية الاتجاه: يُكتشف الاتجاه تلقائيًا من امتداد الملف المرفوع.
    - ملف .pdf   → يُحوَّل إلى Word (.docx)
    - ملف .docx/.doc → يُحوَّل إلى PDF
    """
    cleanup_old_jobs()

    if "file" not in request.files:
        return jsonify(error="لم يتم إرفاق أي ملف.", code="no_file"), 400

    uploaded = request.files["file"]
    if uploaded.filename == "":
        return jsonify(error="لم يتم اختيار ملف.", code="no_file"), 400

    # يُستخرج الامتداد من اسم الملف الأصلي مباشرة (قبل secure_filename)، لأن
    # secure_filename يحذف كل الأحرف غير اللاتينية (كالعربية) — وقد يحذف نقطة
    # الامتداد نفسها، فيصير اسم "ملف.pdf" فارغًا تقريبًا ويُرفض الملف خطأً باعتباره
    # "نوع غير مدعوم" رغم أنه PDF/Word صحيح تمامًا. الحل: نتحقق من الامتداد على
    # الاسم الأصلي، ثم نبني اسمًا آمنًا للحفظ (جذع آمن + الامتداد الصحيح دائمًا).
    original_name = uploaded.filename
    ext = Path(original_name).suffix.lower()

    if ext not in ALLOWED_EXTENSIONS:
        return (
            jsonify(error="الرجاء رفع ملف بصيغة PDF أو Word (DOC/DOCX) فقط.", code="invalid_type"),
            400,
        )

    safe_stem = secure_filename(Path(original_name).stem) or "file"
    filename = f"{safe_stem}{ext}"

    job_dir = new_job_dir(BASE_TMP_DIR)
    input_path = os.path.join(job_dir, filename)
    uploaded.save(input_path)

    is_pdf_input = ext == PDF_EXTENSION

    try:
        with CONVERSION_SEMAPHORE:
            if is_pdf_input:
                use_ocr = request.form.get("use_ocr", "false").lower() == "true"
                # لغة المستند التي اختارها المستخدم صراحةً (قد تختلف عن لغة الواجهة)؛
                # إن لم يحدد شيئًا صالحًا، نعود للتخمين الافتراضي من لغة الواجهة.
                ocr_lang = build_ocr_lang_string(request.form.get("doc_lang")) or resolve_ocr_lang(
                    get_ui_lang()
                )
                output_path = convert_pdf_to_docx(
                    input_path, job_dir, use_ocr=use_ocr, ocr_lang=ocr_lang
                )
                download_ext = ".docx"
                mimetype = WORD_MIMETYPE
            else:
                output_path = convert_docx_to_pdf(input_path, job_dir)
                download_ext = ".pdf"
                mimetype = PDF_MIMETYPE
    except EncryptedPdfError as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify(error=str(exc), code="encrypted"), 422
    except InvalidPdfError as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify(error=str(exc), code="invalid"), 422
    except NeedsOcrError as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify(error=str(exc), code="needs_ocr"), 422
    except ConversionTimeoutError as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify(error=str(exc), code="conversion_timeout"), 422
    except ConversionError as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify(error=str(exc), code="conversion_failed"), 422
    except Exception as exc:  # noqa: BLE001
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify(error="حدث خطأ غير متوقع أثناء التحويل.", code="unknown"), 500

    download_name = Path(filename).stem + download_ext

    @after_this_request
    def _cleanup(response):
        schedule_cleanup(job_dir, delay=60)
        return response

    return send_file(
        output_path,
        as_attachment=True,
        download_name=download_name,
        mimetype=mimetype,
    )


@app.post("/api/inspect")
def inspect():
    """يتحقق سريعًا هل الملف نصي أم صورة ممسوحة ضوئيًا، قبل التحويل الفعلي."""
    if "file" not in request.files:
        return jsonify(error="لم يتم إرفاق أي ملف.", code="no_file"), 400

    uploaded = request.files["file"]
    original_name = uploaded.filename or "file.pdf"
    ext = Path(original_name).suffix.lower()
    if ext != PDF_EXTENSION:
        # الفحص المسبق (OCR) يخص PDF فقط؛ ملفات Word لا تحتاج هذا الفحص إطلاقًا
        return jsonify(page_count=None, has_text_layer=True)

    safe_stem = secure_filename(Path(original_name).stem) or "file"
    filename = f"{safe_stem}{ext}"

    job_dir = new_job_dir(BASE_TMP_DIR)
    pdf_path = os.path.join(job_dir, filename)
    uploaded.save(pdf_path)

    try:
        insight = inspect_pdf(pdf_path)
    except EncryptedPdfError as exc:
        return jsonify(error=str(exc), code="encrypted"), 422
    except InvalidPdfError as exc:
        return jsonify(error=str(exc), code="invalid"), 422
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)

    return jsonify(
        page_count=insight.page_count,
        has_text_layer=insight.has_text_layer,
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port, debug=False)
