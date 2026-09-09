"""
محرك تحويل ثنائي الاتجاه بين PDF و Word (DOCX) - يحافظ على محتوى الملف الأصلي دون أي تغيير.

الاستراتيجية:
1) PDF → Word: نستخدم LibreOffice (headless) لقراءة نص PDF الحقيقي (وليس صورة)
   وإعادة بنائه كمستند Word قابل للتحرير، مع الحفاظ على الفقرات والجداول والصور
   والتنسيق قدر الإمكان، دون تعديل أي حرف من النص.
2) Word → PDF: اتجاه مدعوم أصليًا وبشكل موثوق جدًا في LibreOffice (تصدير PDF من
   Writer)، لذلك لا حاجة لأي حيلة خاصة هنا؛ الناتج مطابق تمامًا لمحتوى ملف Word.
3) الكشف عن ملفات PDF الممسوحة ضوئيًا (صور بلا طبقة نص): في هذه الحالة لا يوجد
   نص حقيقي لاستخراجه، لذلك نعرض خيار OCR (التعرف الضوئي على الحروف) كخطوة
   إضافية اختيارية يقوم بها المستخدم بوعي، لأن أي OCR قد يخطئ في التعرف على بعض
   الأحرف، وهذا مختلف عن "نسخ النص الحقيقي دون تغيير".
"""

import os
import shutil
import signal
import subprocess
import tempfile
import uuid
from dataclasses import dataclass

import pdfplumber
from pypdf import PdfReader
from pypdf.errors import PdfReadError


class ConversionError(Exception):
    """خطأ عام أثناء عملية التحويل، يحمل رسالة مناسبة لعرضها للمستخدم."""


class EncryptedPdfError(ConversionError):
    pass


class InvalidPdfError(ConversionError):
    pass


class NeedsOcrError(ConversionError):
    """يُرفع عندما يكون الملف صورة ممسوحة ضوئيًا بلا نص، ويحتاج تفعيل OCR صراحة."""


class ConversionTimeoutError(ConversionError):
    """يُرفع عندما تستغرق عملية التحويل وقتًا أطول من الحد المسموح به (ملف كبير/معقد جدًا)."""


# مهلتان منفصلتان لمحرّكي PDF→Word (نحاول pdf2docx أولاً، ثم LibreOffice كخطة
# بديلة إن فشل الأول أو أنتج نصًا ناقصًا) بدل مهلة واحدة كبيرة، حتى يصل خطأ
# واضح للمستخدم خلال دقائق معقولة بدل الانتظار إلى ما لا نهاية على ملف معطوب.
PDF2DOCX_TIMEOUT = int(os.environ.get("PDF2DOCX_TIMEOUT", "180"))
LIBREOFFICE_TIMEOUT = int(os.environ.get("LIBREOFFICE_TIMEOUT", "180"))


def _run_with_timeout(func, seconds: int):
    """ينفّذ func() مع مهلة زمنية قصوى (SIGALRM، متاحة على Linux/عمّال gunicorn
    المتزامنين). إن لم تتوفر SIGALRM (مثلاً في بيئة غير POSIX)، يُنفَّذ بلا مهلة."""
    if not seconds or not hasattr(signal, "SIGALRM"):
        return func()

    def _handle_alarm(signum, frame):
        raise ConversionTimeoutError(
            "استغرقت عملية التحويل وقتًا طويلًا جدًا. يمكنك تجربة تفعيل خيار "
            "«التعرف الضوئي على النص (OCR)» كبديل عملي، أو تجربة ملف أصغر."
        )

    previous_handler = signal.signal(signal.SIGALRM, _handle_alarm)
    signal.alarm(seconds)
    try:
        return func()
    finally:
        signal.alarm(0)
        signal.signal(signal.SIGALRM, previous_handler)


# يربط رمز لغة الواجهة (ar/en/fr/es/de) بحزمة لغة Tesseract المناسبة للتعرّف
# الضوئي، وتُستخدم كقيمة افتراضية فقط إن لم يحدد المستخدم لغة المستند صراحةً.
# نضيف eng دائمًا كلغة احتياطية لأن كثيرًا من المستندات تحتوي كلمات أو أرقام
# إنجليزية حتى لو كانت اللغة الأساسية غير ذلك.
OCR_LANG_MAP = {
    "ar": "ara+eng",
    "en": "eng",
    "fr": "fra+eng",
    "es": "spa+eng",
    "de": "deu+eng",
}
DEFAULT_OCR_LANG = "eng"


def resolve_ocr_lang(ui_lang: str) -> str:
    return OCR_LANG_MAP.get((ui_lang or "").lower(), DEFAULT_OCR_LANG)


# قائمة موسّعة بلغات المستند الفعلية (وليس لغة الواجهة) التي يمكن للمستخدم
# اختيارها صراحةً عند استخدام OCR لملف PDF ممسوح ضوئيًا، لأن لغة المستند قد
# تختلف تمامًا عن لغة واجهة الموقع. الصورة (Docker) تُثبّت tesseract-ocr-all
# التي تحتوي كل حزم اللغات هذه وأكثر. كل عنصر: (رمز Tesseract، الاسم بلغته
# الأصلية) — نستخدم الاسم الأصلي بدل الترجمة لتفادي الحاجة لترجمة 40 اسم لغة
# إلى 5 لغات واجهة مختلفة.
OCR_LANGUAGE_CHOICES = [
    ("ara", "العربية"),
    ("eng", "English"),
    ("fra", "Français"),
    ("spa", "Español"),
    ("deu", "Deutsch"),
    ("ita", "Italiano"),
    ("por", "Português"),
    ("rus", "Русский"),
    ("chi_sim", "中文（简体）"),
    ("chi_tra", "中文（繁體）"),
    ("jpn", "日本語"),
    ("kor", "한국어"),
    ("tur", "Türkçe"),
    ("nld", "Nederlands"),
    ("pol", "Polski"),
    ("swe", "Svenska"),
    ("ell", "Ελληνικά"),
    ("heb", "עברית"),
    ("hin", "हिन्दी"),
    ("urd", "اردو"),
    ("fas", "فارسی"),
    ("ind", "Bahasa Indonesia"),
    ("tha", "ไทย"),
    ("vie", "Tiếng Việt"),
    ("ukr", "Українська"),
    ("ces", "Čeština"),
    ("ron", "Română"),
    ("hun", "Magyar"),
    ("bul", "Български"),
    ("srp", "Српски"),
    ("hrv", "Hrvatski"),
    ("slk", "Slovenčina"),
    ("fin", "Suomi"),
    ("dan", "Dansk"),
    ("nor", "Norsk"),
    ("ben", "বাংলা"),
    ("amh", "አማርኛ"),
    ("msa", "Bahasa Melayu"),
    ("aze", "Azərbaycan"),
    ("kat", "ქართული"),
    ("hye", "Հայերեն"),
]
ALLOWED_OCR_CODES = {code for code, _ in OCR_LANGUAGE_CHOICES}


def build_ocr_lang_string(doc_lang_code: str | None) -> str | None:
    """
    يبني معرّف لغة Tesseract (مثل "ara+eng") من رمز لغة مستند اختاره المستخدم
    صراحةً (من OCR_LANGUAGE_CHOICES). يعيد None إن كان الرمز غير معروف/فارغ،
    وعندها يُستخدم resolve_ocr_lang(ui_lang) كقيمة افتراضية بدلاً منه.
    """
    code = (doc_lang_code or "").strip()
    if code not in ALLOWED_OCR_CODES:
        return None
    if code == "eng":
        return "eng"
    return f"{code}+eng"


@dataclass
class PdfInsight:
    page_count: int
    has_text_layer: bool
    is_encrypted: bool


def inspect_pdf(pdf_path: str) -> PdfInsight:
    """يفحص الملف قبل التحويل: عدد الصفحات، هل هو محمي بكلمة مرور، وهل يحتوي نصًا حقيقيًا."""
    try:
        reader = PdfReader(pdf_path)
    except PdfReadError as exc:
        raise InvalidPdfError("الملف تالف أو ليس ملف PDF صالحًا.") from exc
    except Exception as exc:  # noqa: BLE001
        raise InvalidPdfError("تعذّرت قراءة الملف. تأكد من أنه ملف PDF سليم.") from exc

    if reader.is_encrypted:
        # محاولة فتحه بكلمة مرور فارغة (بعض الملفات "محمية" بشكل صوري فقط)
        try:
            if reader.decrypt("") == 0:
                raise EncryptedPdfError(
                    "الملف محمي بكلمة مرور. الرجاء إزالة الحماية أولاً ثم إعادة المحاولة."
                )
        except EncryptedPdfError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise EncryptedPdfError(
                "الملف محمي بكلمة مرور. الرجاء إزالة الحماية أولاً ثم إعادة المحاولة."
            ) from exc

    page_count = len(reader.pages)

    has_text = False
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[: min(3, len(pdf.pages))]:  # يكفي فحص أول 3 صفحات
                text = page.extract_text() or ""
                if text.strip():
                    has_text = True
                    break
    except Exception:  # noqa: BLE001
        # إن تعذر الفحص بـ pdfplumber، نفترض وجود نص ونترك LibreOffice يحاول
        has_text = True

    return PdfInsight(page_count=page_count, has_text_layer=has_text, is_encrypted=False)


def _run_soffice_convert(
    input_path: str,
    out_dir: str,
    target_filter: str,
    output_ext: str,
    infilter: str | None = None,
    timeout: int = 120,
) -> str:
    """
    دالة عامة تشغّل LibreOffice في وضع headless لتحويل ملف من صيغة إلى أخرى.
    مستخدمة في كلا الاتجاهين: PDF → Word و Word → PDF.
    """
    if not shutil.which("soffice"):
        raise ConversionError("محرّك التحويل (LibreOffice) غير مثبت على الخادم.")

    # لكل عملية تحويل نستخدم "ملف تعريف مستخدم" منفصل ومؤقت لتفادي تعارض
    # عمليات LibreOffice المتزامنة على نفس الخادم
    user_profile_dir = tempfile.mkdtemp(prefix="lo_profile_")
    user_install_url = f"file://{user_profile_dir}"

    cmd = [
        "soffice",
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--nodefault",
        "--nofirststartwizard",
        f"-env:UserInstallation={user_install_url}",
        "--convert-to",
        target_filter,
    ]
    if infilter:
        cmd.append(f"--infilter={infilter}")
    cmd.extend(["--outdir", out_dir, input_path])

    try:
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise ConversionTimeoutError(
            "استغرقت عملية التحويل وقتًا طويلًا جدًا. يمكنك تجربة تفعيل خيار "
            "«التعرف الضوئي على النص (OCR)» كبديل عملي، أو تجربة ملف أصغر."
        ) from exc
    finally:
        shutil.rmtree(user_profile_dir, ignore_errors=True)

    if result.returncode != 0:
        stderr = result.stderr.decode("utf-8", errors="ignore")
        raise ConversionError(f"فشل التحويل داخل الخادم: {stderr.strip()[:300]}")

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    produced_path = os.path.join(out_dir, f"{base_name}.{output_ext}")

    if not os.path.exists(produced_path):
        raise ConversionError("لم يتمكن الخادم من إنتاج الملف الناتج. حاول مجددًا.")

    return produced_path


def convert_with_libreoffice(pdf_path: str, out_dir: str, timeout: int = LIBREOFFICE_TIMEOUT) -> str:
    """
    يحوّل PDF إلى DOCX باستخدام LibreOffice في وضع headless.
    يعيد المسار الكامل لملف DOCX الناتج.
    """
    return _run_soffice_convert(
        pdf_path,
        out_dir,
        target_filter="docx:MS Word 2007 XML",
        output_ext="docx",
        # نجبر LibreOffice على فتح الملف بمحرك "Writer" (معالج النصوص) بدل "Draw"،
        # وإلا فسيتعامل مع PDF كرسمة ويفشل تصديره كملف Word
        infilter="writer_pdf_import",
        timeout=timeout,
    )


def convert_docx_to_pdf(docx_path: str, out_dir: str, timeout: int = LIBREOFFICE_TIMEOUT) -> str:
    """
    يحوّل ملف Word (doc/docx) إلى PDF باستخدام LibreOffice. هذا اتجاه مدعوم
    أصليًا وموثوق في LibreOffice (تصدير Writer إلى PDF)، فلا تُغيَّر أي فقرة أو
    جدول أو صورة من محتوى الملف الأصلي.
    """
    return _run_soffice_convert(
        docx_path,
        out_dir,
        target_filter="pdf:writer_pdf_Export",
        output_ext="pdf",
        timeout=timeout,
    )


def convert_with_ocr(pdf_path: str, out_dir: str, lang: str = DEFAULT_OCR_LANG) -> str:
    """
    مسار بديل لملفات PDF الممسوحة ضوئيًا (صور بلا نص): يستخرج نص كل صفحة عبر
    التعرف الضوئي (Tesseract) ثم يبني مستند Word منه. هذه العملية "تعرّف" على
    النص من الصورة ولا "تنسخه" حرفيًا، لذلك تُستخدم فقط عند عدم وجود طبقة نص
    حقيقية في الملف، وبعد موافقة صريحة من المستخدم.
    """
    import pytesseract
    from pdf2image import convert_from_path
    from docx import Document
    from docx.shared import Pt

    images = convert_from_path(pdf_path, dpi=300)
    document = Document()
    style = document.styles["Normal"]
    style.font.size = Pt(11)

    for idx, image in enumerate(images):
        text = pytesseract.image_to_string(image, lang=lang)
        if idx > 0:
            document.add_page_break()
        for line in text.splitlines():
            document.add_paragraph(line)

    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    out_path = os.path.join(out_dir, f"{base_name}.docx")
    document.save(out_path)
    return out_path


def convert_with_pdf2docx(pdf_path: str, out_dir: str) -> str:
    """
    يحوّل PDF إلى DOCX عبر مكتبة pdf2docx: تحلّل محتوى الصفحة مباشرة (نصوصًا
    وجداول وصورًا) وتعيد بناءها كفقرات وجمل Word حقيقية قابلة للتحرير والنسخ،
    بدل الاعتماد فقط على مرشّح استيراد PDF في LibreOffice الذي قد يضع الصفحة
    كاملة كصورة واحدة (بلا نص حقيقي) مع بعض التخطيطات المعقدة أو الخطوط
    المضمّنة غير القياسية — وهذا هو المسار المفضّل لأنه أقرب لـ"نسخ ولصق"
    النص الفعلي دون أي تحويل إلى صور.
    """
    from pdf2docx import Converter

    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    out_path = os.path.join(out_dir, f"{base_name}.docx")

    cv = Converter(pdf_path)
    try:
        cv.convert(out_path)
    finally:
        cv.close()

    if not os.path.exists(out_path):
        raise ConversionError("تعذّر إنشاء ملف Word من هذا الملف عبر محرك النص المباشر.")

    # تصحيح خلل معروف في استخراج النص العربي: بعض الخطوط المضمّنة في PDF
    # (خاصة الناتجة من Word/طابعات PDF عربية) تُشكّل رباط "اللام+الألف" (لا/
    # لأ/لإ/لآ) كرمز خط (glyph) رسم واحد، وبعض مكتبات الاستخراج (ومنها
    # PyMuPDF التي يعتمد عليها pdf2docx) تُخرج حرفَي هذا الرباط بترتيب معكوس
    # (الألف قبل اللام) بدل الترتيب الصحيح. لا يمكن إصلاح هذا نصيًا بشكل أعمى
    # لأن نفس التتابع الحرفي "ألف ثم لام" يرد بشكل صحيح تمامًا في كلمات شائعة
    # جدًا (مثل "مسألة"، "إلى"، "ألف")، لذلك نعتمد فقط على تطابق/تداخل صناديق
    # الأحرف المحيطة (bbox) من ملف PDF نفسه: تطابق الصندوقين يعني يقينًا أن
    # الحرفين استُخرجا من رسمة واحدة فعلية (رباط حقيقي)، وعندها فقط يكون
    # التصحيح آمنًا 100% دون أي التباس مع كلمات أخرى. أي فشل في هذه الخطوة
    # الإضافية لا يُفشل التحويل؛ يبقى ناتج pdf2docx الأصلي كما هو.
    try:
        corrections = _find_arabic_ligature_corrections(pdf_path)
        if corrections:
            _apply_text_corrections_to_docx(out_path, corrections)
    except Exception:  # noqa: BLE001
        pass

    return out_path


# رموز الألف في العربية (عادية/همزة فوق/همزة تحت/مدّة) التي تُشكّل رباطًا
# طباعيًا إلزاميًا مع اللام إذا وردت بعدها مباشرة (لا/لأ/لإ/لآ)
_ARABIC_ALEF_VARIANTS = ("ا", "أ", "إ", "آ")
_ARABIC_LAM = "ل"


def _split_chars_into_words(chars: list) -> list:
    """يقسّم قائمة أحرف span واحد (من PyMuPDF rawdict) إلى كلمات، بالاعتماد
    على الفجوة الأفقية الفعلية بين صناديق الأحرف المتتالية (وليس على وجود
    حرف مسافة صريح، الذي قد لا يُستخرج دائمًا بشكل منفصل)."""
    words = []
    current = []
    prev_bbox = None
    for ch in chars:
        bbox = ch.get("bbox")
        if not bbox:
            continue
        if prev_bbox is not None:
            gap = min(abs(bbox[0] - prev_bbox[2]), abs(bbox[2] - prev_bbox[0]))
            char_width = max(bbox[2] - bbox[0], 1e-3)
            if gap > max(char_width * 1.3, 2.0):
                if current:
                    words.append(current)
                current = []
        current.append(ch)
        prev_bbox = bbox
    if current:
        words.append(current)
    return words


def _bboxes_overlap_strongly(b1, b2, min_ratio: float = 0.3) -> bool:
    """يتحقق أن صندوقي حرفين متداخلان بشكل كبير (وليس مجرد تجاور طبيعي بين
    حرفين متتاليين)، كدليل قوي على أنهما ينتميان لرمز خط (glyph) واحد فعليًا."""
    x0, y0 = max(b1[0], b2[0]), max(b1[1], b2[1])
    x1, y1 = min(b1[2], b2[2]), min(b1[3], b2[3])
    if x1 <= x0 or y1 <= y0:
        return False
    inter_area = (x1 - x0) * (y1 - y0)
    area1 = max((b1[2] - b1[0]) * (b1[3] - b1[1]), 1e-6)
    area2 = max((b2[2] - b2[0]) * (b2[3] - b2[1]), 1e-6)
    return (inter_area / min(area1, area2)) > min_ratio


def _fix_word_ligature_order(word_chars: list) -> tuple[str, bool]:
    """يفحص كلمة واحدة (قائمة أحرف مع صناديقها) بحثًا عن رباط لام+ألف معكوس
    (ألف مباشرة قبل لام بنفس صندوق الرسم تقريبًا)، ويعيد النص بعد التصحيح
    مع علامة تدل إن تم أي تغيير فعلي."""
    letters = [ch.get("c", "") for ch in word_chars]
    changed = False
    i = 0
    while i < len(word_chars) - 1:
        c1, c2 = word_chars[i], word_chars[i + 1]
        if letters[i] in _ARABIC_ALEF_VARIANTS and letters[i + 1] == _ARABIC_LAM:
            b1, b2 = c1.get("bbox"), c2.get("bbox")
            if b1 and b2 and _bboxes_overlap_strongly(b1, b2):
                letters[i], letters[i + 1] = letters[i + 1], letters[i]
                changed = True
        i += 1
    return "".join(letters), changed


def _find_arabic_ligature_corrections(pdf_path: str) -> dict:
    """يفحص ملف PDF على مستوى الحرف (PyMuPDF rawdict) في كل الصفحات، ويبني
    قاموس تصحيحات {الكلمة كما استُخرجت (معطوبة): الكلمة الصحيحة} خاصًا بهذا
    الملف تحديدًا، بالاعتماد حصريًا على تطابق صناديق الأحرف كما هو موضّح في
    _fix_word_ligature_order. يعيد قاموسًا فارغًا إن تعذّر الفحص (مثلاً PyMuPDF
    غير متاح)، دون رفع أي استثناء يوقف التحويل."""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return {}

    corrections: dict = {}
    doc = fitz.open(pdf_path)
    try:
        for page in doc:
            raw = page.get_text("rawdict")
            for block in raw.get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        chars = span.get("chars", [])
                        for word_chars in _split_chars_into_words(chars):
                            if len(word_chars) < 2:
                                continue
                            original = "".join(ch.get("c", "") for ch in word_chars)
                            fixed, changed = _fix_word_ligature_order(word_chars)
                            if changed and original and fixed and original != fixed:
                                corrections[original] = fixed
    finally:
        doc.close()
    return corrections


def _apply_text_corrections_to_docx(docx_path: str, corrections: dict) -> None:
    """يطبّق تصحيحات نصية دقيقة (استبدال كلمات كاملة مؤكَّدة) على كل نصوص
    ملف DOCX (فقرات المتن والجداول ورؤوس/تذييلات الصفحات)، عبر استبدال داخل
    كل "run" على حدة (لا يمسّ التنسيق)، ثم يحفظ الملف في مكانه."""
    if not corrections:
        return

    from docx import Document

    # نبدأ بالكلمات الأطول أولاً لتفادي أن يطابق استبدال قصير جزءًا من كلمة أطول
    ordered = sorted(corrections.items(), key=lambda kv: -len(kv[0]))

    def _fix_paragraph(paragraph):
        for run in paragraph.runs:
            if not run.text:
                continue
            new_text = run.text
            for wrong, right in ordered:
                if wrong in new_text:
                    new_text = new_text.replace(wrong, right)
            if new_text != run.text:
                run.text = new_text

    doc = Document(docx_path)
    for paragraph in doc.paragraphs:
        _fix_paragraph(paragraph)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _fix_paragraph(paragraph)
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            _fix_paragraph(paragraph)
        for paragraph in section.footer.paragraphs:
            _fix_paragraph(paragraph)

    doc.save(docx_path)


def _docx_text_length(docx_path: str) -> int:
    """يحسب إجمالي عدد الأحرف النصية الفعلية داخل ملف DOCX. نمشي على شجرة XML
    الكاملة لجسم المستند (word/document.xml) ونجمع كل عناصر <w:t> بدل الاكتفاء
    بواجهة python-docx العليا (doc.paragraphs/doc.tables)، لأن بعض محركات
    التحويل (مثل مرشّح استيراد PDF في LibreOffice) قد تضع النص داخل "أطر"
    (text frames) مرتبطة بموضع ثابت في الصفحة بدل فقرات المستند العادية،
    وهذه الأطر لا تظهر في doc.paragraphs رغم احتوائها نصًا حقيقيًا فعلاً."""
    from docx import Document

    doc = Document(docx_path)
    total = 0
    for node in doc.element.body.iter():
        if node.tag.endswith("}t") and node.text:
            total += len(node.text)
    return total


def _extract_pdf_text_length(pdf_path: str, max_pages: int = 200) -> int:
    """يحسب إجمالي طول النص الحقيقي المستخرج من ملف PDF المصدر (كل الصفحات
    تقريبًا)، ليُقارَن بطول النص في ملف Word الناتج والتأكد أن التحويل نقل
    النص فعليًا ولم يهمل معظمه (كأن يضع الصفحات كصور بلا نص)."""
    total = 0
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:max_pages]:
                total += len(page.extract_text() or "")
    except Exception:  # noqa: BLE001
        pass
    return total


# نطاقات يونيكود الأساسية للأحرف العربية (وما يشابهها من أبجديات RTL: الفارسية
# والأردية تتقاسم نفس نطاق الحروف العربية الأساسي، لذلك هذا الفحص يخدمها أيضًا).
_RTL_RANGES = (
    (0x0600, 0x06FF),  # العربية
    (0x0750, 0x077F),  # ملحق العربية
    (0x08A0, 0x08FF),  # ملحق العربية الموسّع
    (0xFB50, 0xFDFF),  # أشكال العرض العربية أ
    (0xFE70, 0xFEFF),  # أشكال العرض العربية ب
)


def _is_rtl_char(ch: str) -> bool:
    code = ord(ch)
    return any(lo <= code <= hi for lo, hi in _RTL_RANGES)


def _is_rtl_heavy_pdf(pdf_path: str, max_pages: int = 5, threshold: float = 0.15) -> bool:
    """
    يفحص عيّنة من نص الملف (أول عدّة صفحات) ليقرّر إن كانت غالبية أحرفه
    الأبجدية من نوع عربي/RTL. نستخدم هذا لتفضيل محرك LibreOffice (الذي يعتمد
    داخليًا على poppler لقراءة طبقة نص PDF) على محرك pdf2docx/PyMuPDF، لأن
    الأخير معروف بخلل في ترتيب الكلمات والأرقام داخل السطر الواحد عند التعامل
    مع نصوص عربية RTL في بعض ملفات PDF (يُخرج النص بترتيب بصري/معكوس جزئيًا
    بدل الترتيب المنطقي الصحيح للقراءة)، بعكس poppler الذي يتعامل مع هذه
    الحالة بشكل صحيح وموثوق لهذا النوع من الملفات.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            sample = ""
            for page in pdf.pages[:max_pages]:
                sample += page.extract_text() or ""
    except Exception:  # noqa: BLE001
        return False

    letters = [ch for ch in sample if ch.isalpha()]
    if len(letters) < 20:
        return False

    rtl_count = sum(1 for ch in letters if _is_rtl_char(ch))
    return (rtl_count / len(letters)) >= threshold


def convert_pdf_to_docx(
    pdf_path: str, out_dir: str, use_ocr: bool = False, ocr_lang: str = DEFAULT_OCR_LANG
) -> str:
    """نقطة الدخول الرئيسية: تفحص الملف ثم تختار مسار التحويل المناسب.

    للملفات النصية (غير الممسوحة ضوئيًا)، نجرّب أولاً pdf2docx لأنه ينقل النص
    الحقيقي حرفيًا كفقرات Word قابلة للتحرير (أقرب لـ"نسخ ولصق"). نتحقق أن
    الناتج يحتوي فعلاً نصًا يقارب حجم نص المصدر (وليس صفحات كصور)، وإلا نعود
    لمحرك LibreOffice كخطة بديلة. إن فشل الاثنان أو استغرقا وقتًا طويلًا جدًا،
    يصل خطأ واضح للمستخدم مع اقتراح تفعيل OCR كبديل عملي.
    """
    insight = inspect_pdf(pdf_path)

    if use_ocr:
        return convert_with_ocr(pdf_path, out_dir, lang=ocr_lang)

    if not insight.has_text_layer:
        raise NeedsOcrError(
            "هذا الملف يبدو مسحًا ضوئيًا (صورة) بلا نص قابل للاستخراج. "
            "فعّل خيار «التعرف الضوئي على النص (OCR)» لتحويله."
        )

    source_text_len = _extract_pdf_text_length(pdf_path)

    # للملفات العربية/RTL: PyMuPDF (الذي تعتمد عليه pdf2docx) قد يُخرج بعض
    # الأسطر بترتيب كلمات/أرقام معكوس جزئيًا (خلل معروف في قراءة نصوص RTL من
    # بعض ملفات PDF)، بعكس LibreOffice الذي يعتمد داخليًا على poppler ويقرأ
    # هذا النوع من الملفات بترتيب صحيح وموثوق. لذلك نبدأ بـ LibreOffice
    # مباشرة لهذه الملفات، ونستخدم pdf2docx فقط كخطة بديلة إن فشل LibreOffice.
    if _is_rtl_heavy_pdf(pdf_path):
        try:
            lo_path = _run_with_timeout(
                lambda: convert_with_libreoffice(pdf_path, out_dir, timeout=LIBREOFFICE_TIMEOUT),
                LIBREOFFICE_TIMEOUT,
            )
            produced_len = _docx_text_length(lo_path)
            if source_text_len == 0 or produced_len >= source_text_len * 0.4:
                return lo_path
        except ConversionTimeoutError:
            pass
        except Exception:  # noqa: BLE001
            pass
        # لم ينجح LibreOffice (أو أنتج نصًا ناقصًا) — نجرّب pdf2docx كخطة بديلة
        try:
            return _run_with_timeout(
                lambda: convert_with_pdf2docx(pdf_path, out_dir), PDF2DOCX_TIMEOUT
            )
        except Exception as exc:  # noqa: BLE001
            raise ConversionError(
                "تعذّر تحويل هذا الملف عبر كل محركات التحويل المتاحة."
            ) from exc

    pdf2docx_path = None
    produced_len = 0
    try:
        pdf2docx_path = _run_with_timeout(
            lambda: convert_with_pdf2docx(pdf_path, out_dir), PDF2DOCX_TIMEOUT
        )
        produced_len = _docx_text_length(pdf2docx_path)
    except ConversionTimeoutError:
        pdf2docx_path = None
    except Exception:  # noqa: BLE001
        # أي فشل في المحرك الأول (pdf2docx) لا يجب أن يوقف العملية بالكامل؛
        # ننتقل لمحرك LibreOffice كخطة بديلة موثوقة
        pdf2docx_path = None

    # نقبل ناتج pdf2docx فقط إن كان يحتوي نصًا حقيقيًا يقارب حجم نص المصدر
    # (٪40 كحد أدنى تحوّطًا لاختلافات بسيطة في طريقة استخراج النص)، وإلا فهذا
    # مؤشر أن المحرك لم ينقل النص فعليًا (كأن يضع الصفحات كصور)
    if pdf2docx_path and (source_text_len == 0 or produced_len >= source_text_len * 0.4):
        return pdf2docx_path

    if pdf2docx_path and os.path.exists(pdf2docx_path):
        try:
            os.remove(pdf2docx_path)
        except OSError:
            pass

    return convert_with_libreoffice(pdf_path, out_dir, timeout=LIBREOFFICE_TIMEOUT)


def new_job_dir(base_tmp: str) -> str:
    job_dir = os.path.join(base_tmp, uuid.uuid4().hex)
    os.makedirs(job_dir, exist_ok=True)
    return job_dir
