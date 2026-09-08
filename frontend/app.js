document.getElementById("year").textContent = new Date().getFullYear();

// ---------- عناصر الصفحة ----------
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");
const browseBtn = document.getElementById("browse-btn");
const fileInfo = document.getElementById("file-info");
const fileNameEl = document.getElementById("file-name");
const fileSizeEl = document.getElementById("file-size");
const removeFileBtn = document.getElementById("remove-file");
const convertBtn = document.getElementById("convert-btn");
const ocrWarning = document.getElementById("ocr-warning");
const ocrToggle = document.getElementById("ocr-toggle");
const ocrLangRow = document.getElementById("ocr-lang-row");
const ocrLangSelect = document.getElementById("ocr-lang-select");
const progressBox = document.getElementById("progress-box");
const resultBox = document.getElementById("result-box");
const errorBox = document.getElementById("error-box");
const errorText = document.getElementById("error-text");
const errorOcrRetry = document.getElementById("error-ocr-retry");
const errorOcrLangSelect = document.getElementById("error-ocr-lang-select");
const retryOcrBtn = document.getElementById("retry-ocr-btn");
const downloadLink = document.getElementById("download-link");
const convertAnotherBtn = document.getElementById("convert-another");
const tryAgainBtn = document.getElementById("try-again");

const langFab = document.getElementById("lang-fab");
const langMenu = document.getElementById("lang-menu");
const langOptions = document.querySelectorAll(".lang-option");

const MAX_SIZE = 50 * 1024 * 1024;
let selectedFile = null;
let direction = null; // "pdf2word" | "word2pdf"
let currentLang = DEFAULT_LANG;
let lastErrorCode = null; // لإعادة ترجمة رسالة الخطأ الحالية عند تبديل اللغة

// لغة الواجهة الحالية ليست بالضرورة لغة المستند الممسوح ضوئيًا (OCR)، لذلك
// نعرض للمستخدم قائمة موسّعة بلغات المستند الفعلية (تُجلب من الخادم) ليختار
// منها بنفسه، مع اختيار افتراضي معقول مبني على لغة الواجهة الحالية.
const UI_TO_OCR_DEFAULT = { ar: "ara", en: "eng", fr: "fra", es: "spa", de: "deu" };
let ocrLanguageList = null;

async function fetchOcrLanguages() {
  if (ocrLanguageList) return ocrLanguageList;
  try {
    const res = await fetch("/api/ocr-languages");
    const data = await res.json();
    if (Array.isArray(data.languages) && data.languages.length) {
      ocrLanguageList = data.languages;
    }
  } catch (err) {
    // إن تعذّر الجلب، تبقى القائمة فارغة ويُستخدم تخمين لغة الواجهة كخطة بديلة
  }
  return ocrLanguageList || [];
}

function populateOcrSelect(selectEl, languages) {
  selectEl.innerHTML = "";
  languages.forEach(({ code, name }) => {
    const opt = document.createElement("option");
    opt.value = code;
    opt.textContent = name;
    selectEl.appendChild(opt);
  });
  const preferred = UI_TO_OCR_DEFAULT[currentLang] || "eng";
  if ([...selectEl.options].some((o) => o.value === preferred)) {
    selectEl.value = preferred;
  }
}

async function loadOcrLanguages() {
  const languages = await fetchOcrLanguages();
  if (languages.length) populateOcrSelect(ocrLangSelect, languages);
}

// ---------- الترجمة (i18n) ----------
function t(key) {
  const dict = TRANSLATIONS[currentLang] || TRANSLATIONS[DEFAULT_LANG];
  return dict[key] ?? TRANSLATIONS[DEFAULT_LANG][key] ?? key;
}

function applyLanguage(lang) {
  if (!SUPPORTED_LANGS.includes(lang)) lang = DEFAULT_LANG;
  currentLang = lang;
  const dict = TRANSLATIONS[lang];

  document.documentElement.lang = lang;
  document.documentElement.dir = dict.dir;
  document.title = dict.pageTitle;

  document.querySelectorAll("[data-i18n]").forEach((el) => {
    const key = el.getAttribute("data-i18n");
    el.textContent = t(key);
  });

  document.querySelectorAll("[data-i18n-html]").forEach((el) => {
    const key = el.getAttribute("data-i18n-html");
    el.innerHTML = t(key);
  });

  document.querySelectorAll("[data-i18n-title]").forEach((el) => {
    const key = el.getAttribute("data-i18n-title");
    el.title = t(key);
  });

  langOptions.forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.lang === lang);
  });

  // النصوص التي تعتمد على اتجاه التحويل الحالي (PDF→Word أو Word→PDF)
  updateDirectionTexts();

  // إعادة عرض رسالة الخطأ الحالية باللغة الجديدة إن وُجدت
  if (lastErrorCode && !errorBox.classList.contains("hidden")) {
    errorText.textContent = t(lastErrorCode);
  }

  saveLang(lang);
}

function updateDirectionTexts() {
  if (direction === "word2pdf") {
    convertBtn.textContent = t("convertToPdfBtn");
    downloadLink.textContent = t("downloadPdfBtn");
  } else {
    convertBtn.textContent = t("convertToWordBtn");
    downloadLink.textContent = t("downloadWordBtn");
  }
}

// ---------- زر اللغة العائم ----------
function openLangMenu() {
  langMenu.classList.remove("hidden");
  langFab.setAttribute("aria-expanded", "true");
}
function closeLangMenu() {
  langMenu.classList.add("hidden");
  langFab.setAttribute("aria-expanded", "false");
}

langFab.addEventListener("click", (e) => {
  e.stopPropagation();
  if (langMenu.classList.contains("hidden")) openLangMenu();
  else closeLangMenu();
});

langOptions.forEach((btn) => {
  btn.addEventListener("click", () => {
    applyLanguage(btn.dataset.lang);
    closeLangMenu();
  });
});

document.addEventListener("click", (e) => {
  if (!document.getElementById("lang-widget").contains(e.target)) {
    closeLangMenu();
  }
});

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") closeLangMenu();
});

// ---------- أدوات مساعدة ----------
function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function detectDirection(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".pdf")) return "pdf2word";
  if (name.endsWith(".docx") || name.endsWith(".doc")) return "word2pdf";
  return null;
}

function resetUI() {
  fileInfo.classList.add("hidden");
  progressBox.classList.add("hidden");
  resultBox.classList.add("hidden");
  errorBox.classList.add("hidden");
  errorOcrRetry.classList.add("hidden");
  ocrWarning.classList.add("hidden");
  ocrToggle.checked = false;
  ocrLangRow.hidden = true;
  dropzone.classList.remove("hidden");
  selectedFile = null;
  direction = null;
  lastErrorCode = null;
  fileInput.value = "";
}

ocrToggle.addEventListener("change", () => {
  ocrLangRow.hidden = !ocrToggle.checked;
  if (ocrToggle.checked) loadOcrLanguages();
});

async function showError(code, fallbackMessage) {
  progressBox.classList.add("hidden");
  fileInfo.classList.add("hidden");
  dropzone.classList.add("hidden");
  errorBox.classList.remove("hidden");
  lastErrorCode = code;
  errorText.textContent = code ? t(code) : fallbackMessage || t("err_unknown");

  // عند انتهاء المهلة الزمنية لتحويل ملف PDF نصي، نعرض خيار إعادة المحاولة
  // مباشرة عبر OCR (بدون إعادة رفع الملف)، لأن هذا بديل عملي فوري
  if (code === "err_conversion_timeout" && direction === "pdf2word" && selectedFile) {
    errorOcrRetry.classList.remove("hidden");
    const languages = await fetchOcrLanguages();
    if (languages.length) populateOcrSelect(errorOcrLangSelect, languages);
  } else {
    errorOcrRetry.classList.add("hidden");
  }
}

// ---------- التعامل مع اختيار الملف ----------
async function handleFile(file) {
  if (!file) return;

  const dir = detectDirection(file);
  if (!dir) {
    showError("err_invalid_type");
    return;
  }

  if (file.size > MAX_SIZE) {
    showError("err_file_too_large");
    return;
  }

  selectedFile = file;
  direction = dir;

  dropzone.classList.add("hidden");
  fileInfo.classList.remove("hidden");
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = formatSize(file.size);
  ocrWarning.classList.add("hidden");
  ocrToggle.checked = false;
  ocrLangRow.hidden = true;
  updateDirectionTexts();

  // فحص مسبق: فقط لملفات PDF، للتأكد هل تحتوي نصًا حقيقيًا أم أنها مسح ضوئي
  if (dir === "pdf2word") {
    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("lang", currentLang);
      const res = await fetch("/api/inspect", { method: "POST", body: formData });
      const data = await res.json();
      if (res.ok && data.has_text_layer === false) {
        ocrWarning.classList.remove("hidden");
      }
    } catch (err) {
      // إن فشل الفحص المسبق، لا بأس؛ سيُكتشف الأمر عند التحويل الفعلي
    }
  }
}

browseBtn.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", (e) => {
  handleFile(e.target.files[0]);
});

["dragenter", "dragover"].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });
});

["dragleave", "drop"].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
  });
});

dropzone.addEventListener("drop", (e) => {
  const file = e.dataTransfer.files[0];
  handleFile(file);
});

removeFileBtn.addEventListener("click", resetUI);
tryAgainBtn.addEventListener("click", resetUI);
convertAnotherBtn.addEventListener("click", resetUI);

// ---------- التحويل ----------
async function runConversion({ forceOcr = false } = {}) {
  if (!selectedFile) return;

  fileInfo.classList.add("hidden");
  errorBox.classList.add("hidden");
  progressBox.classList.remove("hidden");

  const formData = new FormData();
  formData.append("file", selectedFile);
  formData.append("lang", currentLang);
  if (direction === "pdf2word") {
    const useOcr = forceOcr || ocrToggle.checked;
    formData.append("use_ocr", useOcr ? "true" : "false");
    if (useOcr) {
      const langVal = forceOcr ? errorOcrLangSelect.value : ocrLangSelect.value;
      if (langVal) formData.append("doc_lang", langVal);
    }
  }

  try {
    const res = await fetch("/api/convert", { method: "POST", body: formData });

    if (!res.ok) {
      let code = "err_conversion_failed";
      try {
        const data = await res.json();
        if (data.code) code = `err_${data.code}`;
      } catch (_) {}
      progressBox.classList.add("hidden");
      showError(code);
      return;
    }

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const baseName = selectedFile.name.replace(/\.(pdf|docx?|)$/i, "");
    const outExt = direction === "word2pdf" ? ".pdf" : ".docx";

    downloadLink.href = url;
    downloadLink.download = `${baseName}${outExt}`;

    progressBox.classList.add("hidden");
    resultBox.classList.remove("hidden");
    updateDirectionTexts();
  } catch (err) {
    progressBox.classList.add("hidden");
    showError("err_network");
  }
}

convertBtn.addEventListener("click", () => runConversion());
retryOcrBtn.addEventListener("click", () => runConversion({ forceOcr: true }));

// ---------- التهيئة الأولية ----------
applyLanguage(detectInitialLang());
