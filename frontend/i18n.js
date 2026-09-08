/* ترجمات الواجهة: العربية، الإنجليزية، الفرنسية، الإسبانية، الألمانية */

const TRANSLATIONS = {
  ar: {
    dir: "rtl",
    pageTitle: "PDF ⇄ Word — تحويل احترافي",
    brand: 'حوّل <span class="accent">PDF</span> ⇄ <span class="accent">Word</span>',
    tagline: "تحويل احترافي وسريع في الاتجاهين، مع الحفاظ الكامل على محتوى ملفك الأصلي",
    dropzoneTitle: "اسحب وأفلت ملف PDF أو Word هنا",
    or: "أو",
    browseBtn: "اختر ملفًا من جهازك",
    dropzoneHint: "الصيغ المدعومة: PDF, DOC, DOCX — الحد الأقصى لحجم الملف: 50 ميغابايت",
    removeFile: "إزالة الملف",
    ocrNotice:
      "يبدو أن هذا الملف عبارة عن صفحات ممسوحة ضوئيًا (صور) ولا يحتوي على نص حقيقي. يمكنك تفعيل خيار «التعرف الضوئي على النص (OCR)» أدناه، مع العلم أن هذه الطريقة تتعرّف على الحروف من الصورة وقد تحتوي على أخطاء طفيفة بحسب جودة المسح.",
    ocrToggle: "تفعيل التعرف الضوئي على النص (OCR)",
    convertToWordBtn: "تحويل إلى Word",
    convertToPdfBtn: "تحويل إلى PDF",
    progressText: "جارٍ تحويل الملف، الرجاء الانتظار…",
    resultTitle: "تم التحويل بنجاح",
    downloadWordBtn: "تنزيل ملف Word",
    downloadPdfBtn: "تنزيل ملف PDF",
    convertAnother: "تحويل ملف آخر",
    tryAgain: "حاول مرة أخرى",
    feat1Title: "خصوصية تامة",
    feat1Text: "تُحذف ملفاتك تلقائيًا من الخادم بعد التحويل مباشرة.",
    feat2Title: "دقة في المحتوى",
    feat2Text: "لا تغيير في النص أو الجداول أو الصور — فقط تغيير الصيغة، في الاتجاهين.",
    feat3Title: "سريع وسهل",
    feat3Text: "ارفع الملف واحصل على النسخة المحوَّلة جاهزة خلال ثوانٍ.",
    rights: "جميع الحقوق محفوظة",
    err_no_file: "لم يتم اختيار أي ملف.",
    err_invalid_type: "الرجاء رفع ملف بصيغة PDF أو Word (DOC/DOCX) فقط.",
    err_encrypted: "الملف محمي بكلمة مرور. الرجاء إزالة الحماية أولاً ثم إعادة المحاولة.",
    err_invalid: "الملف تالف أو غير صالح. تأكد من سلامته وحاول مجددًا.",
    err_needs_ocr:
      "هذا الملف يبدو مسحًا ضوئيًا (صورة) بلا نص قابل للاستخراج. فعّل خيار «التعرف الضوئي على النص (OCR)» لتحويله.",
    err_conversion_failed: "تعذّر تحويل الملف. حاول مرة أخرى.",
    err_unknown: "حدث خطأ غير متوقع. حاول مرة أخرى.",
    err_file_too_large: "حجم الملف يتجاوز الحد الأقصى المسموح به (50 ميغابايت).",
    err_network: "تعذّر الاتصال بالخادم. تحقق من اتصالك بالإنترنت وحاول مجددًا.",
  },

  en: {
    dir: "ltr",
    pageTitle: "PDF ⇄ Word — Professional Conversion",
    brand: 'Convert <span class="accent">PDF</span> ⇄ <span class="accent">Word</span>',
    tagline: "Fast, professional conversion in both directions — your original content stays 100% unchanged",
    dropzoneTitle: "Drag and drop a PDF or Word file here",
    or: "or",
    browseBtn: "Choose a file from your device",
    dropzoneHint: "Supported formats: PDF, DOC, DOCX — Maximum file size: 50 MB",
    removeFile: "Remove file",
    ocrNotice:
      "This file looks like a scanned document (images) with no real text. You can enable “Optical Character Recognition (OCR)” below — note that OCR recognizes characters from the image and may contain minor errors depending on scan quality.",
    ocrToggle: "Enable text recognition (OCR)",
    convertToWordBtn: "Convert to Word",
    convertToPdfBtn: "Convert to PDF",
    progressText: "Converting your file, please wait…",
    resultTitle: "Conversion completed successfully",
    downloadWordBtn: "Download Word file",
    downloadPdfBtn: "Download PDF file",
    convertAnother: "Convert another file",
    tryAgain: "Try again",
    feat1Title: "Complete privacy",
    feat1Text: "Your files are automatically deleted from the server right after conversion.",
    feat2Title: "Content accuracy",
    feat2Text: "No change to text, tables or images — only the format changes, in either direction.",
    feat3Title: "Fast and easy",
    feat3Text: "Upload your file and get the converted version ready within seconds.",
    rights: "All rights reserved",
    err_no_file: "No file was selected.",
    err_invalid_type: "Please upload a PDF or Word (DOC/DOCX) file only.",
    err_encrypted: "This file is password-protected. Please remove the protection and try again.",
    err_invalid: "The file is corrupted or invalid. Please check it and try again.",
    err_needs_ocr:
      "This file looks like a scanned document (image) with no extractable text. Enable “Optical Character Recognition (OCR)” to convert it.",
    err_conversion_failed: "Could not convert the file. Please try again.",
    err_unknown: "An unexpected error occurred. Please try again.",
    err_file_too_large: "The file exceeds the maximum allowed size (50 MB).",
    err_network: "Could not reach the server. Check your internet connection and try again.",
  },

  fr: {
    dir: "ltr",
    pageTitle: "PDF ⇄ Word — Conversion professionnelle",
    brand: 'Convertir <span class="accent">PDF</span> ⇄ <span class="accent">Word</span>',
    tagline: "Conversion rapide et professionnelle dans les deux sens — le contenu original reste inchangé à 100 %",
    dropzoneTitle: "Glissez-déposez un fichier PDF ou Word ici",
    or: "ou",
    browseBtn: "Choisir un fichier sur votre appareil",
    dropzoneHint: "Formats pris en charge : PDF, DOC, DOCX — Taille maximale : 50 Mo",
    removeFile: "Retirer le fichier",
    ocrNotice:
      "Ce fichier semble être un document numérisé (images) sans texte réel. Vous pouvez activer la « reconnaissance optique de caractères (OCR) » ci-dessous — notez que l'OCR reconnaît les caractères à partir de l'image et peut contenir de légères erreurs selon la qualité du scan.",
    ocrToggle: "Activer la reconnaissance de texte (OCR)",
    convertToWordBtn: "Convertir en Word",
    convertToPdfBtn: "Convertir en PDF",
    progressText: "Conversion de votre fichier en cours, veuillez patienter…",
    resultTitle: "Conversion réussie",
    downloadWordBtn: "Télécharger le fichier Word",
    downloadPdfBtn: "Télécharger le fichier PDF",
    convertAnother: "Convertir un autre fichier",
    tryAgain: "Réessayer",
    feat1Title: "Confidentialité totale",
    feat1Text: "Vos fichiers sont automatiquement supprimés du serveur juste après la conversion.",
    feat2Title: "Fidélité du contenu",
    feat2Text: "Aucun changement dans le texte, les tableaux ou les images — seul le format change, dans les deux sens.",
    feat3Title: "Rapide et simple",
    feat3Text: "Téléversez votre fichier et obtenez la version convertie en quelques secondes.",
    rights: "Tous droits réservés",
    err_no_file: "Aucun fichier sélectionné.",
    err_invalid_type: "Veuillez téléverser uniquement un fichier PDF ou Word (DOC/DOCX).",
    err_encrypted: "Ce fichier est protégé par un mot de passe. Veuillez retirer la protection et réessayer.",
    err_invalid: "Le fichier est corrompu ou invalide. Vérifiez-le et réessayez.",
    err_needs_ocr:
      "Ce fichier semble être un document numérisé (image) sans texte extractible. Activez la « reconnaissance optique de caractères (OCR) » pour le convertir.",
    err_conversion_failed: "Impossible de convertir le fichier. Veuillez réessayer.",
    err_unknown: "Une erreur inattendue s'est produite. Veuillez réessayer.",
    err_file_too_large: "Le fichier dépasse la taille maximale autorisée (50 Mo).",
    err_network: "Impossible de contacter le serveur. Vérifiez votre connexion internet et réessayez.",
  },

  es: {
    dir: "ltr",
    pageTitle: "PDF ⇄ Word — Conversión profesional",
    brand: 'Convertir <span class="accent">PDF</span> ⇄ <span class="accent">Word</span>',
    tagline: "Conversión rápida y profesional en ambos sentidos — el contenido original permanece 100 % intacto",
    dropzoneTitle: "Arrastra y suelta un archivo PDF o Word aquí",
    or: "o",
    browseBtn: "Elegir un archivo de tu dispositivo",
    dropzoneHint: "Formatos admitidos: PDF, DOC, DOCX — Tamaño máximo: 50 MB",
    removeFile: "Quitar archivo",
    ocrNotice:
      "Este archivo parece ser un documento escaneado (imágenes) sin texto real. Puedes activar el «reconocimiento óptico de caracteres (OCR)» a continuación; ten en cuenta que el OCR reconoce los caracteres a partir de la imagen y puede contener pequeños errores según la calidad del escaneo.",
    ocrToggle: "Activar reconocimiento de texto (OCR)",
    convertToWordBtn: "Convertir a Word",
    convertToPdfBtn: "Convertir a PDF",
    progressText: "Convirtiendo tu archivo, por favor espera…",
    resultTitle: "Conversión completada con éxito",
    downloadWordBtn: "Descargar archivo Word",
    downloadPdfBtn: "Descargar archivo PDF",
    convertAnother: "Convertir otro archivo",
    tryAgain: "Intentar de nuevo",
    feat1Title: "Privacidad total",
    feat1Text: "Tus archivos se eliminan automáticamente del servidor justo después de la conversión.",
    feat2Title: "Fidelidad del contenido",
    feat2Text: "Sin cambios en el texto, las tablas o las imágenes — solo cambia el formato, en ambos sentidos.",
    feat3Title: "Rápido y sencillo",
    feat3Text: "Sube tu archivo y obtén la versión convertida en segundos.",
    rights: "Todos los derechos reservados",
    err_no_file: "No se seleccionó ningún archivo.",
    err_invalid_type: "Por favor sube únicamente un archivo PDF o Word (DOC/DOCX).",
    err_encrypted: "Este archivo está protegido con contraseña. Elimina la protección e inténtalo de nuevo.",
    err_invalid: "El archivo está dañado o no es válido. Compruébalo e inténtalo de nuevo.",
    err_needs_ocr:
      "Este archivo parece ser un documento escaneado (imagen) sin texto extraíble. Activa el «reconocimiento óptico de caracteres (OCR)» para convertirlo.",
    err_conversion_failed: "No se pudo convertir el archivo. Inténtalo de nuevo.",
    err_unknown: "Ocurrió un error inesperado. Inténtalo de nuevo.",
    err_file_too_large: "El archivo supera el tamaño máximo permitido (50 MB).",
    err_network: "No se pudo conectar con el servidor. Comprueba tu conexión a internet e inténtalo de nuevo.",
  },

  de: {
    dir: "ltr",
    pageTitle: "PDF ⇄ Word — Professionelle Umwandlung",
    brand: '<span class="accent">PDF</span> ⇄ <span class="accent">Word</span> umwandeln',
    tagline: "Schnelle, professionelle Umwandlung in beide Richtungen — Ihr Originalinhalt bleibt zu 100 % unverändert",
    dropzoneTitle: "PDF- oder Word-Datei hierher ziehen",
    or: "oder",
    browseBtn: "Datei von Ihrem Gerät auswählen",
    dropzoneHint: "Unterstützte Formate: PDF, DOC, DOCX — Maximale Dateigröße: 50 MB",
    removeFile: "Datei entfernen",
    ocrNotice:
      "Diese Datei scheint ein gescanntes Dokument (Bilder) ohne echten Text zu sein. Sie können unten die „optische Zeichenerkennung (OCR)“ aktivieren — beachten Sie, dass OCR Zeichen aus dem Bild erkennt und je nach Scanqualität kleinere Fehler enthalten kann.",
    ocrToggle: "Texterkennung (OCR) aktivieren",
    convertToWordBtn: "In Word umwandeln",
    convertToPdfBtn: "In PDF umwandeln",
    progressText: "Ihre Datei wird umgewandelt, bitte warten…",
    resultTitle: "Umwandlung erfolgreich abgeschlossen",
    downloadWordBtn: "Word-Datei herunterladen",
    downloadPdfBtn: "PDF-Datei herunterladen",
    convertAnother: "Weitere Datei umwandeln",
    tryAgain: "Erneut versuchen",
    feat1Title: "Vollständige Privatsphäre",
    feat1Text: "Ihre Dateien werden direkt nach der Umwandlung automatisch vom Server gelöscht.",
    feat2Title: "Inhaltliche Genauigkeit",
    feat2Text: "Keine Änderung an Text, Tabellen oder Bildern — nur das Format ändert sich, in beide Richtungen.",
    feat3Title: "Schnell und einfach",
    feat3Text: "Datei hochladen und die umgewandelte Version in Sekunden erhalten.",
    rights: "Alle Rechte vorbehalten",
    err_no_file: "Es wurde keine Datei ausgewählt.",
    err_invalid_type: "Bitte laden Sie nur eine PDF- oder Word-Datei (DOC/DOCX) hoch.",
    err_encrypted: "Diese Datei ist passwortgeschützt. Bitte entfernen Sie den Schutz und versuchen Sie es erneut.",
    err_invalid: "Die Datei ist beschädigt oder ungültig. Bitte überprüfen Sie sie und versuchen Sie es erneut.",
    err_needs_ocr:
      "Diese Datei scheint ein gescanntes Dokument (Bild) ohne extrahierbaren Text zu sein. Aktivieren Sie die „optische Zeichenerkennung (OCR)“, um sie umzuwandeln.",
    err_conversion_failed: "Die Datei konnte nicht umgewandelt werden. Bitte versuchen Sie es erneut.",
    err_unknown: "Ein unerwarteter Fehler ist aufgetreten. Bitte versuchen Sie es erneut.",
    err_file_too_large: "Die Datei überschreitet die maximal zulässige Größe (50 MB).",
    err_network: "Der Server konnte nicht erreicht werden. Prüfen Sie Ihre Internetverbindung und versuchen Sie es erneut.",
  },
};

const SUPPORTED_LANGS = Object.keys(TRANSLATIONS);
const DEFAULT_LANG = "en";

function detectInitialLang() {
  try {
    const saved = localStorage.getItem("pdf2word_lang");
    if (saved && SUPPORTED_LANGS.includes(saved)) return saved;
  } catch (e) {
    /* localStorage قد يكون غير متاح؛ نتابع بلا حفظ */
  }
  const browserLangs = navigator.languages || [navigator.language || DEFAULT_LANG];
  for (const bl of browserLangs) {
    const code = bl.slice(0, 2).toLowerCase();
    if (SUPPORTED_LANGS.includes(code)) return code;
  }
  return DEFAULT_LANG;
}

function saveLang(lang) {
  try {
    localStorage.setItem("pdf2word_lang", lang);
  } catch (e) {
    /* تجاهل: بعض المتصفحات تمنع localStorage (وضع خاص مثلاً) */
  }
}
