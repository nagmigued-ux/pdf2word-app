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


# يربط رمز لغة الواجهة (ar/en/fr/es/de) بحزمة لغة Tesseract المناسبة للتعرّف
# الضوئي. نضيف eng دائمًا كلغة احتياطية لأن كثيرًا من المستندات تحتوي كلمات أو
# أرقام إنجليزية حتى لو كانت اللغة الأساسية غير ذلك.
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
        raise ConversionError("استغرقت عملية التحويل وقتًا طويلًا جدًا. جرّب ملفًا أصغر.") from exc
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


def convert_with_libreoffice(pdf_path: str, out_dir: str, timeout: int = 120) -> str:
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


def convert_docx_to_pdf(docx_path: str, out_dir: str, timeout: int = 120) -> str:
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


def convert_pdf_to_docx(
    pdf_path: str, out_dir: str, use_ocr: bool = False, ocr_lang: str = DEFAULT_OCR_LANG
) -> str:
    """نقطة الدخول الرئيسية: تفحص الملف ثم تختار مسار التحويل المناسب."""
    insight = inspect_pdf(pdf_path)

    if use_ocr:
        return convert_with_ocr(pdf_path, out_dir, lang=ocr_lang)

    if not insight.has_text_layer:
        raise NeedsOcrError(
            "هذا الملف يبدو مسحًا ضوئيًا (صورة) بلا نص قابل للاستخراج. "
            "فعّل خيار «التعرف الضوئي على النص (OCR)» لتحويله."
        )

    return convert_with_libreoffice(pdf_path, out_dir)


def new_job_dir(base_tmp: str) -> str:
    job_dir = os.path.join(base_tmp, uuid.uuid4().hex)
    os.makedirs(job_dir, exist_ok=True)
    return job_dir
