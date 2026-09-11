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
import re
import shutil
import signal
import subprocess
import tempfile
import uuid
from collections import Counter
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

    # تصحيح ترتيب النص العربي: pdf2docx (المعتمد على PyMuPDF لتحليل التخطيط
    # وبناء الفقرات) قد يُخرج نص بعض الفقرات بترتيب كلمات/حروف معكوس جزئيًا
    # لنصوص RTL، وهذا خلل في منطق pdf2docx الداخلي لإعادة بناء الفقرات نفسه
    # وليس فقط في استخراج الحروف الخام. لذلك لا نحاول تصحيح النص بعد إنتاجه
    # بأنماط جزئية (كرباط حرفين أو كلمة بمفردها)، بل نستبدل نص كل فقرة كاملة
    # بالنص الصحيح المطابق المستخرج مباشرة من نفس ملف PDF عبر pdftotext
    # (الذي تحقّقنا أنه يقرأ نصوص RTL بترتيب صحيح موثوق)، بمطابقة كل فقرة مع
    # المقطع المقابل من النص الصحيح عبر تطابق "كيس الحروف" الكامل (بعد حذف
    # الفراغات) — فإن لم يوجد تطابق مضمون 100% لفقرة ما، تبقى كما أنتجها
    # pdf2docx دون أي لمس (لا تخمين، لا استبدال جزئي محتمل الخطأ). فشل هذه
    # الخطوة كاملة (لأي سبب) لا يُفشل التحويل؛ يبقى ناتج pdf2docx كما هو.
    try:
        _rebuild_pdf2docx_text_order(pdf_path, out_path)
    except Exception:  # noqa: BLE001
        pass

    return out_path


_BIDI_CONTROL_RE = re.compile("[‎‏‪-‮⁦-⁩]")
_WHITESPACE_RE = re.compile(r"\s+")


def _nospace(s: str) -> str:
    return _WHITESPACE_RE.sub("", s)


def _extract_ground_truth_lines(pdf_path: str) -> list:
    """
    يستخرج نص PDF بترتيب قراءة منطقي صحيح عبر pdftotext (poppler)، والذي
    تحقّقنا يدويًا (بمقارنته حرفًا بحرف مع أصل نص المستند) أنه يقرأ نصوص RTL
    العربية في هذا النوع من ملفات PDF بترتيب صحيح 100%، بعكس محرك التحليل
    الداخلي لـ pdf2docx/PyMuPDF. يعيد قائمة أسطر نصية (كل سطر = سطر مرئي واحد
    في الصفحة، بترتيبه الصحيح من أول الصفحة الأولى لآخر الصفحة الأخيرة)، بعد
    حذف رموز التحكم في اتجاه الكتابة (bidi control chars) وتجاهل الأسطر
    الفارغة تمامًا. يعيد قائمة فارغة إن تعذّر ذلك (poppler غير متاح، أو فشل
    الاستخراج لأي سبب)، ولا يرفع أي استثناء.
    """
    if not shutil.which("pdftotext"):
        return []
    try:
        result = subprocess.run(
            ["pdftotext", "-enc", "UTF-8", "-layout", pdf_path, "-"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=60,
            check=False,
        )
    except Exception:  # noqa: BLE001
        return []
    if result.returncode != 0:
        return []

    text = result.stdout.decode("utf-8", errors="ignore")
    text = _BIDI_CONTROL_RE.sub("", text)

    lines = []
    for page_text in text.split("\f"):
        for raw_line in page_text.split("\n"):
            if raw_line.strip():
                lines.append(raw_line)
    return lines


def _find_match_position(flat: str, target_counter: Counter, length: int, exp: int, lo: int, hi: int):
    """
    يبحث عن أول موضع في flat يكون فيه المقطع (بطول length) مطابقًا تمامًا
    (نفس تكرار كل حرف) لـ target_counter، بالبحث أولًا من exp صعودًا حتى hi
    ثم من exp نازلًا حتى lo. يستخدم "نافذة منزلقة" بتحديث تدريجي (إضافة حرف
    جديد وحذف حرف خارج) بدل إعادة بناء العداد من الصفر عند كل موضع، لأداء
    عملي عبر نطاق بحث كبير (قد يبلغ آلاف الأحرف) على مستند من عشرات الصفحات.
    """
    if hi < lo or length <= 0 or lo < 0 or hi + length > len(flat):
        hi = min(hi, len(flat) - length)
        if hi < lo:
            return None

    start = max(lo, min(exp, hi))

    window = Counter(flat[start:start + length])
    pos = start
    while True:
        if window == target_counter:
            return pos
        nxt = pos + 1
        if nxt > hi:
            break
        old_c = flat[pos]
        window[old_c] -= 1
        if window[old_c] == 0:
            del window[old_c]
        new_c = flat[pos + length]
        window[new_c] = window.get(new_c, 0) + 1
        pos = nxt

    pos = start - 1
    if pos < lo:
        return None
    window = Counter(flat[pos:pos + length])
    while True:
        if window == target_counter:
            return pos
        prev_pos = pos - 1
        if prev_pos < lo:
            break
        new_c = flat[prev_pos]
        window[new_c] = window.get(new_c, 0) + 1
        old_c = flat[prev_pos + length]
        window[old_c] -= 1
        if window[old_c] == 0:
            del window[old_c]
        pos = prev_pos
    return None


def _align_docx_to_ground_truth(paragraph_texts: list, truth_lines: list) -> dict:
    """
    يحاذي نصوص فقرات DOCX (كما استخرجها/أعاد بناءها pdf2docx، قد تكون مشوّشة
    الترتيب لبعض الفقرات) مع أسطر النص الصحيحة (truth_lines)، بالاعتماد
    فقط على تطابق "كيس الحروف" الكامل (multiset الحروف بعد حذف كل الفراغات)
    لكل فقرة مقابل نطاق من النص الصحيح حول موضعها المتوقع (بناء على مجموع
    أطوال الفقرات السابقة، بافتراض أن ترتيب الفقرات نفسه يوافق ترتيب القراءة
    الصحيح في الغالب، وأن الخلل محصور في ترتيب الكلمات/الحروف داخل الفقرة).

    يعيد قاموسًا {فهرس الفقرة (في paragraph_texts): النص الصحيح المطابق} لكل
    فقرة وُجد لها تطابق مضمون 100% (تطابق كامل لعدد/نوع كل حرف)، عبر خطوتين:
    مطابقة مباشرة حول الموضع المتوقع، ثم "تعبئة الفراغات" بين فقرتين مطابقتين
    لو كان طول الفراغ بينهما (في النص الصحيح) يساوي تمامًا مجموع أطوال
    الفقرات غير المطابقة الواقعة بينهما (دليل قوي أنها هي نفسها، فقط لم
    تُطابَق مباشرة بسبب فرق طفيف كوجود همزات/أرقام مختلفة الترميز). لا يخمّن
    أي تصحيح لفقرة لم يُوجد لها تطابق مضمون بأي من الطريقتين.
    """
    # نص واحد كبير يجمع كل الأسطر (مفصولة بسطر جديد، كما في الأصل)، مع فهرس
    # يربط كل موضع في "flat" (النص بلا فراغات) بموضعه الحقيقي في هذا النص
    # الكبير المتضمّن الفراغات — لنستطيع استخراج مقطع نصي دقيق (بفراغاته
    # الأصلية الصحيحة بين الكلمات) بدل إعادة أسطر كاملة قد تتجاوز حدود
    # المقطع المطلوب فعليًا (وهو خلل كان يُسبب تكرار نفس النص الكامل لعدة
    # فقرات متتالية تتقاسم سطرًا واحدًا في النص الصحيح).
    spaced_full = "\n".join(line.strip() for line in truth_lines)
    non_space_positions = [i for i, ch in enumerate(spaced_full) if not ch.isspace()]

    flat = "".join(spaced_full[i] for i in non_space_positions)
    n = len(flat)

    def reconstruct(start, end):
        if start >= end or start < 0 or end > len(non_space_positions):
            return ""
        s = non_space_positions[start]
        e = non_space_positions[end - 1] + 1
        return re.sub(r"[ \t]+", " ", spaced_full[s:e]).strip()

    d_lens = [len(_nospace(t)) for t in paragraph_texts]
    d_counters = [Counter(_nospace(t)) for t in paragraph_texts]
    expected = []
    c = 0
    for length in d_lens:
        expected.append(c)
        c += length

    window_radius = 1500
    results = [None] * len(paragraph_texts)
    for di, length in enumerate(d_lens):
        if length == 0:
            continue
        exp = expected[di]
        lo = max(0, exp - window_radius)
        hi = min(n - length, exp + window_radius)
        if hi < lo:
            continue
        found = _find_match_position(flat, d_counters[di], length, exp, lo, hi)
        if found is not None:
            results[di] = (found, found + length)

    total = len(paragraph_texts)
    i = 0
    while i < total:
        if results[i] is not None:
            i += 1
            continue
        j = i
        while j < total and results[j] is None:
            j += 1
        prev_end = results[i - 1][1] if i > 0 else 0
        next_start = results[j][0] if j < total else n
        gap_len = next_start - prev_end
        sum_lens = sum(d_lens[i:j])
        if gap_len == sum_lens and gap_len >= 0:
            cursor = prev_end
            for k in range(i, j):
                length = d_lens[k]
                results[k] = (cursor, cursor + length)
                cursor += length
        i = j + 1

    corrections = {}
    for idx, r in enumerate(results):
        if r is None:
            continue
        corrected = reconstruct(r[0], r[1])
        if corrected and corrected != paragraph_texts[idx]:
            corrections[idx] = corrected
    return corrections


def _apply_paragraph_text_rebuild(docx_path: str, corrections_by_index: dict) -> int:
    """
    يستبدل نص كل فقرة (بفهرسها في doc.paragraphs) بالنص الصحيح المقابل من
    corrections_by_index، بوضعه في "run" واحد فقط (أكبر run موجود أصلًا في
    الفقرة من حيث طول نصه، لإبقاء التنسيق السائد لغالب نص الفقرة)، وتفريغ كل
    الـ runs الأخرى في نفس الفقرة. هذا يعني أن أي تمايز تنسيقي صغير داخل تلك
    الفقرة بعينها (كخط مختلف الحجم لجزء قصير من السطر) قد يُفقد، لكن هذا مقبول
    لأنه محصور بالفقرات التي كان نصها فعليًا مشوّش الترتيب، وأولوية صحة
    المحتوى فيها أعلى من تفصيل تنسيقي بسيط. يعيد عدد الفقرات التي طُبّق عليها
    استبدال فعلي."""
    from docx import Document

    doc = Document(docx_path)
    paragraphs = doc.paragraphs
    applied = 0
    for idx, new_text in corrections_by_index.items():
        if idx < 0 or idx >= len(paragraphs):
            continue
        paragraph = paragraphs[idx]
        runs_with_text = [r for r in paragraph.runs if r.text]
        if not runs_with_text:
            continue
        anchor = max(runs_with_text, key=lambda r: len(r.text))
        anchor.text = new_text
        for r in runs_with_text:
            if r is not anchor:
                r.text = ""
        applied += 1

    if applied:
        doc.save(docx_path)
    return applied


def _rebuild_pdf2docx_text_order(pdf_path: str, docx_path: str) -> int:
    """
    نقطة الدخول: يستخرج النص الصحيح من pdf_path عبر pdftotext، ثم يحاذيه مع
    فقرات docx_path (ناتج pdf2docx) ويستبدل نص كل فقرة وُجد لها تطابق مضمون
    بالنص الصحيح المقابل. يتضمن فحصًا وقائيًا أوليًا: إن اختلف إجمالي عدد
    الأحرف (بعد حذف الفراغات) بين docx وnص PDF الصحيح بأكثر من ٪15 (مؤشر أن
    البنية مختلفة جوهريًا، كوجود جداول/صور معقدة)، يتوقف دون أي تعديل تفاديًا
    لاستبدال غير موثوق. يعيد عدد الفقرات التي عُدّلت فعليًا (0 يعني عدم إجراء
    أي تغيير، وهو أمر آمن تمامًا: يبقى ناتج pdf2docx الأصلي كما هو)."""
    truth_lines = _extract_ground_truth_lines(pdf_path)
    if not truth_lines:
        return 0

    from docx import Document

    doc = Document(docx_path)
    texts = [p.text for p in doc.paragraphs]
    nonempty_idx = [i for i, t in enumerate(texts) if t.strip()]
    nonempty_texts = [texts[i] for i in nonempty_idx]
    if not nonempty_texts:
        return 0

    d_total = sum(len(_nospace(t)) for t in nonempty_texts)
    t_total = sum(len(_nospace(line)) for line in truth_lines)
    if d_total == 0 or t_total == 0:
        return 0
    ratio = d_total / t_total
    if ratio < 0.85 or ratio > 1.15:
        return 0

    local_corrections = _align_docx_to_ground_truth(nonempty_texts, truth_lines)
    if not local_corrections:
        return 0

    corrections_by_index = {nonempty_idx[k]: v for k, v in local_corrections.items()}
    return _apply_paragraph_text_rebuild(docx_path, corrections_by_index)


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


# نطاقات يونيكود الأساسية للأحرف العربية، لتحديد هل سطر معيّن عربي بشكل غالب.
_ARABIC_RANGES = ((0x0600, 0x06FF), (0x0750, 0x077F), (0xFB50, 0xFDFF), (0xFE70, 0xFEFF))


def _is_arabic_letter(ch: str) -> bool:
    if not ch:
        return False
    code = ord(ch)
    return any(lo <= code <= hi for lo, hi in _ARABIC_RANGES)


def _word_left_x(word_chars: list) -> float:
    boxes = [ch.get("bbox") for ch in word_chars if ch.get("bbox")]
    return min(b[0] for b in boxes) if boxes else 0.0


def _find_arabic_word_order_corrections(pdf_path: str) -> dict:
    """
    يفحص ملف PDF على مستوى السطر الواحد (تجميع كل spans المنتمية لنفس
    الـ line في PyMuPDF rawdict)، ويكتشف الأسطر العربية التي استُخرجت بترتيب
    كلمات معكوس جزئيًا عن ترتيب القراءة الصحيح (من اليمين لليسار) — وهو خلل
    معروف في استخراج نص RTL من بعض ملفات PDF عبر PyMuPDF (الذي تعتمد عليه
    pdf2docx)، ويؤدي أحيانًا أيضًا لالتصاق كلمتين متجاورتين بلا مسافة بينهما.

    الاعتماد الحصري هو على المواضع الهندسية الفعلية (bbox) لكل كلمة داخل
    السطر لتحديد ترتيب القراءة الصحيح (الكلمة اليمنى أولاً)، فمقارنته بالترتيب
    الذي استخرجته المكتبة فعليًا. يُبنى قاموس تصحيحات {نص السطر كما استُخرج
    حرفيًا: نص السطر بالترتيب الصحيح}، ولا يُضاف أي سطر لم يتغيّر ترتيبه فعليًا
    (أي أن هذا التصحيح لا يمسّ أي سطر مستخرج بشكل صحيح من الأصل).
    """
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
                    chars = []
                    for span in line.get("spans", []):
                        chars.extend(span.get("chars", []))
                    if len(chars) < 6:
                        continue

                    letters = [ch.get("c", "") for ch in chars if ch.get("c", "").isalpha()]
                    if len(letters) < 6:
                        continue
                    arabic_count = sum(1 for c in letters if _is_arabic_letter(c))
                    if arabic_count < len(letters) * 0.5:
                        continue  # سطر ليس عربيًا بشكل غالب؛ لا نلمسه

                    # نستثني "كلمات" لا تحتوي إلا على مسافات (فجوات تباعد/محاذاة
                    # لا تمثل محتوى فعليًا)، لأنها كانت تُسبب اختلافات صورية في
                    # الترتيب لا علاقة لها بخلل RTL الحقيقي
                    all_words = _split_chars_into_words(chars)
                    words = [
                        w for w in all_words
                        if "".join(ch.get("c", "") for ch in w).strip()
                    ]
                    if len(words) < 2:
                        continue

                    order = sorted(range(len(words)), key=lambda i: -_word_left_x(words[i]))
                    if order == list(range(len(words))):
                        continue  # الترتيب المستخرج صحيح فعليًا؛ لا حاجة لتصحيح

                    original_text = "".join(ch.get("c", "") for ch in chars).strip()
                    if len(original_text) < 8:
                        continue
                    fixed_words = []
                    for i in order:
                        w_text, _ = _fix_word_ligature_order(words[i])
                        fixed_words.append(w_text.strip())
                    corrected_text = " ".join(w for w in fixed_words if w)

                    if original_text and corrected_text and original_text != corrected_text:
                        corrections[original_text] = corrected_text
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

        # بعض التصحيحات (خصوصًا عكس ترتيب كلمات سطر كامل) قد يمتد نصها على
        # أكثر من "run" واحد داخل نفس الفقرة (لو غيّر pdf2docx التنسيق في
        # منتصف السطر)، فلا يجدها التصحيح أعلاه لأنه يفحص كل run على حدة.
        # هنا نتحقق من النص الكامل المدمج للفقرة؛ فإن وجدنا تصحيحًا منطبقًا
        # عليه ولم يُطبَّق بعد، نطبّقه على النص الكامل ونضعه في أول run غير
        # فارغ (مع تفريغ باقي الـ runs)، على حساب فقدان أي تمايز تنسيقي بسيط
        # داخل تلك الفقرة فقط (كحجم خط مختلف لجزء من السطر)، تفاديًا لبقاء
        # النص معكوسًا بالكامل.
        runs_with_text = [r for r in paragraph.runs if r.text]
        if not runs_with_text:
            return
        merged = "".join(r.text for r in runs_with_text)
        new_merged = merged
        for wrong, right in ordered:
            if wrong in new_merged:
                new_merged = new_merged.replace(wrong, right)
        if new_merged != merged:
            runs_with_text[0].text = new_merged
            for r in runs_with_text[1:]:
                r.text = ""

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
