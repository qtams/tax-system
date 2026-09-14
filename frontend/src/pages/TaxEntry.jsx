import { useRef, useState } from "react";
import { Controller, useForm } from "react-hook-form";
import { NumericFormat } from "react-number-format";
import DatePicker from "react-datepicker";
import Swal from "sweetalert2";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import mammoth from "mammoth";
import { createWorker } from "tesseract.js";
import "react-datepicker/dist/react-datepicker.css";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

/* =========================================================
   CONSTANTS
   ========================================================= */

const CURRENT_YEAR = new Date().getFullYear();

const MAX_FILES = 10;
const MAX_FILE_SIZE = 20 * 1024 * 1024;
const MAX_TOTAL_SIZE = 100 * 1024 * 1024;

/*
  TEST MODE:
  Process only the first 15 pages.
*/
const PDF_TEST_PAGE_LIMIT = 15;

/*
  PDF.js scale 3 is approximately 216 DPI.
*/
const PDF_OCR_SCALE = 3;

const ALLOWED_EXTENSIONS = [".pdf", ".docx"];

const ALLOWED_MIME_TYPES = [
  "application/pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
];

const FORM_TYPES = [
  "BIR Form 2307",
  "BIR Form 1701Q",
  "BIR Form 1701",
  "BIR Form 1701A",
  "BIR Form 1601-EQ",
  "BIR Form 2551Q",
  "BIR Form 2550Q",
];

const DEFAULT_VALUES = {
  formType: "BIR Form 1701Q",
  taxpayerName: "",
  tin: "",
  quarter: "",
  year: new Date(CURRENT_YEAR, 0, 1),
  grossSales: "",
  taxableIncome: "",
  taxDue: "",
};

/*
  Exact BIR 2307 regions based on the first 15 pages
  of the sample document.

  All positions are ratios, not fixed pixels.
*/
const FORM_2307_OCR_REGIONS = {
  period: {
    x: 0.2,
    y: 0.132,
    width: 0.69,
    height: 0.04,
  },

  /*
    PART I
    Taxpayer Identification Number
  */
  payeeTin: {
    x: 0.28,
    y: 0.158,
    width: 0.51,
    height: 0.043,
  },

  /*
    PART I
    Payee name
  */
  payeeName: {
    x: 0.025,
    y: 0.187,
    width: 0.95,
    height: 0.045,
  },

  /*
    PART II
    Payor / Withholding Agent TIN
  */
  payorTin: {
    x: 0.28,
    y: 0.268,
    width: 0.51,
    height: 0.043,
  },

  /*
    PART II
    Payor name
  */
  payorName: {
    x: 0.025,
    y: 0.296,
    width: 0.95,
    height: 0.045,
  },

  nature: {
    x: 0.018,
    y: 0.382,
    width: 0.29,
    height: 0.065,
  },

  atc: {
    x: 0.28,
    y: 0.388,
    width: 0.095,
    height: 0.05,
  },

  total: {
    x: 0.7,
    y: 0.388,
    width: 0.17,
    height: 0.05,
  },

  taxWithheld: {
    x: 0.84,
    y: 0.388,
    width: 0.15,
    height: 0.05,
  },
};

/* =========================================================
   SWEET ALERT
   ========================================================= */

const Toast = Swal.mixin({
  toast: true,
  position: "top-end",
  showConfirmButton: false,
  timer: 2500,
  timerProgressBar: true,
});

function openExtractionLoading(totalFiles) {
  void Swal.fire({
    title: "Extracting tax records",

    html: `
      <div style="text-align:center;">
        <div
          data-extraction-status
          style="
            margin-top:8px;
            font-size:14px;
            font-weight:600;
            color:#18181b;
          "
        >
          Preparing documents...
        </div>

        <div
          data-extraction-detail
          style="
            margin-top:7px;
            font-size:12px;
            line-height:1.5;
            color:#71717a;
          "
        >
          0 of ${totalFiles} files processed
        </div>
      </div>
    `,

    allowOutsideClick: false,
    allowEscapeKey: false,
    showConfirmButton: false,

    didOpen: () => {
      Swal.showLoading();
    },
  });
}

function updateExtractionLoading(status, detail = "") {
  const container = Swal.getHtmlContainer();

  if (!container) {
    return;
  }

  const statusElement = container.querySelector("[data-extraction-status]");

  const detailElement = container.querySelector("[data-extraction-detail]");

  if (statusElement) {
    statusElement.textContent = status;
  }

  if (detailElement) {
    detailElement.textContent = detail;
  }
}

/* =========================================================
   INPUT STYLE
   ========================================================= */

const baseInputClass = `
  h-11
  w-full
  rounded-md
  border
  bg-white
  px-3
  text-sm
  text-zinc-900
  outline-none
  transition
  placeholder:text-zinc-400
  disabled:cursor-not-allowed
  disabled:bg-zinc-100
`;

function getInputClass(error) {
  return `${baseInputClass} ${
    error
      ? "border-red-400 focus:border-red-500 focus:ring-2 focus:ring-red-100"
      : "border-zinc-300 hover:border-zinc-400 focus:border-zinc-900 focus:ring-2 focus:ring-zinc-100"
  }`;
}

/* =========================================================
   GENERAL HELPERS
   ========================================================= */

function cleanText(text = "") {
  return String(text)
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\u00a0/g, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function normalizeText(text = "") {
  return cleanText(text).replace(/\n/g, " ").replace(/\s+/g, " ").trim();
}

function cleanExtractedName(value = "") {
  return String(value)
    .replace(/\s+/g, " ")
    .replace(/^[\s:;,.–—|_\-]+/, "")
    .replace(/[\s:;,.–—|_\-]+$/, "")
    .trim();
}

function roundAmount(value) {
  return Math.round(Number(value) * 100) / 100;
}

function parseAmount(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const number = Number(
    String(value)
      .replace(/PHP/gi, "")
      .replace(/₱/g, "")
      .replace(/,/g, "")
      .replace(/[^\d.-]/g, ""),
  );

  return Number.isFinite(number) ? number : null;
}

function getFileExtension(filename) {
  const index = filename.lastIndexOf(".");

  return index >= 0 ? filename.slice(index).toLowerCase() : "";
}

function getFileTypeLabel(file) {
  return getFileExtension(file.name) === ".pdf" ? "PDF" : "DOCX";
}

function getFileKey(file) {
  return `${file.name}-${file.size}-${file.lastModified}`;
}

function isAllowedFile(file) {
  const extension = getFileExtension(file.name);

  return (
    ALLOWED_EXTENSIONS.includes(extension) &&
    (ALLOWED_MIME_TYPES.includes(file.type) || !file.type)
  );
}

function formatFileSize(bytes) {
  if (!bytes) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];

  const index = Math.min(
    Math.floor(Math.log(bytes) / Math.log(1024)),
    units.length - 1,
  );

  return `${(bytes / 1024 ** index).toFixed(
    index === 0 ? 0 : 1,
  )} ${units[index]}`;
}

/* =========================================================
   TIN HELPERS
   ========================================================= */

function formatTin(value = "") {
  const digits = String(value).replace(/\D/g, "").slice(0, 14);

  const groups = [
    digits.slice(0, 3),
    digits.slice(3, 6),
    digits.slice(6, 9),
    digits.slice(9, 14),
  ].filter(Boolean);

  return groups.join("-");
}

function validateTin(value) {
  const length = String(value).replace(/\D/g, "").length;

  if ([9, 12, 14].includes(length)) {
    return true;
  }

  return "TIN must contain 9, 12, or 14 digits.";
}

function isValidExtractedTin(value) {
  const length = String(value || "").replace(/\D/g, "").length;

  return [9, 12, 14].includes(length);
}

function validateTaxpayerName(value) {
  const name = value.trim();

  if (name.length < 2) {
    return "Taxpayer name must contain at least 2 characters.";
  }

  if (name.length > 100) {
    return "Taxpayer name cannot exceed 100 characters.";
  }

  if (!/^[\p{L}\p{M}.'\-\s,]+$/u.test(name)) {
    return "Enter a valid taxpayer name.";
  }

  return true;
}

function validateAmount(value) {
  if (value === "" || value === null || value === undefined) {
    return "This amount is required.";
  }

  const amount = Number(value);

  if (Number.isNaN(amount)) {
    return "Enter a valid amount.";
  }

  if (amount < 0) {
    return "Amount cannot be negative.";
  }

  return true;
}

/* =========================================================
   PERIOD HELPERS
   ========================================================= */

function getQuarterFromMonth(month) {
  const value = Number(month);

  if (!Number.isFinite(value) || value < 1 || value > 12) {
    return "";
  }

  if (value <= 3) {
    return "1st Quarter";
  }

  if (value <= 6) {
    return "2nd Quarter";
  }

  if (value <= 9) {
    return "3rd Quarter";
  }

  return "4th Quarter";
}

function formatDateParts(month, day, year) {
  return `${String(month).padStart(2, "0")}/${String(day).padStart(
    2,
    "0",
  )}/${year}`;
}

/* =========================================================
   BIR 2307 OCR - CHARACTER NORMALIZATION
   ========================================================= */

function normalizeOcrDigitText(value = "") {
  return String(value)
    .toUpperCase()
    .replace(/[OQD]/g, "0")
    .replace(/[IL|!]/g, "1")
    .replace(/Z/g, "2")
    .replace(/S/g, "5")
    .replace(/G/g, "6")
    .replace(/B/g, "8");
}

/* =========================================================
   BIR 2307 OCR - TIN
   ========================================================= */

function parse2307OcrTin(rawText = "") {
  const corrected = normalizeOcrDigitText(rawText);

  /*
    Example:

    297 - 944 - 186 - 00000
  */
  const grouped = corrected.match(
    /(\d{3})\D{0,10}(\d{3})\D{0,10}(\d{3})\D{0,10}(\d{3,5})/,
  );

  if (grouped) {
    let branch = grouped[4];

    /*
      Tesseract sometimes returns:

      000
      instead of
      00000
    */
    if (/^0+$/.test(branch) && branch.length < 5) {
      branch = branch.padEnd(5, "0");
    }

    const digits = `${grouped[1]}${grouped[2]}${grouped[3]}${branch}`;

    if ([12, 14].includes(digits.length)) {
      return formatTin(digits);
    }
  }

  /*
    Fallback if OCR removes
    all separators.
  */
  let digits = corrected.replace(/\D/g, "");

  if (digits.length > 14) {
    digits = digits.slice(0, 14);
  }

  if ([12, 14].includes(digits.length)) {
    return formatTin(digits);
  }

  /*
    Restore missing zeroes
    for branch 00000.
  */
  if (
    digits.length >= 12 &&
    digits.length < 14 &&
    /^0+$/.test(digits.slice(9))
  ) {
    return formatTin(digits.padEnd(14, "0"));
  }

  return "";
}

/* =========================================================
   BIR 2307 OCR - NAMES
   ========================================================= */

function cleanNameLine(value = "") {
  return String(value)
    .replace(/[^\p{L}\p{M}0-9&.,'’()\- ]/gu, " ")
    .replace(/\s+/g, " ")
    .replace(/^[\s.,:;|_\-]+/, "")
    .replace(/[\s.,:;|_\-]+$/, "")
    .trim();
}

function countLetters(value = "") {
  return (String(value).match(/\p{L}/gu) || []).length;
}

function looksLike2307NameLabel(value = "") {
  const text = normalizeText(value);

  return (
    /Payee(?:'s|’s)?\s+Name/i.test(text) ||
    /Payor(?:'s|’s)?\s+Name/i.test(text) ||
    /Last\s+Name/i.test(text) ||
    /First\s+Name/i.test(text) ||
    /Middle\s+Name/i.test(text) ||
    /Registered\s+Name/i.test(text) ||
    /Non[\s-]*Individual/i.test(text) ||
    /For\s+Individual/i.test(text) ||
    /Taxpayer\s+Identification/i.test(text) ||
    /Registered\s+Address/i.test(text) ||
    /Foreign\s+Address/i.test(text) ||
    /ZIP\s+Code/i.test(text) ||
    /\bTIN\b/i.test(text)
  );
}

function strip2307NameLabelFromLine(value = "", type = "payee") {
  const line = cleanNameLine(value);

  if (!line) {
    return "";
  }

  const labelRegex =
    type === "payor" ? /Payor(?:'s|’s)?\s+Name/i : /Payee(?:'s|’s)?\s+Name/i;

  const label = labelRegex.exec(line);

  if (!label) {
    return line;
  }

  const remainder = line.slice(label.index + label[0].length);

  const closingParenthesis = remainder.lastIndexOf(")");

  if (closingParenthesis >= 0) {
    const after = cleanNameLine(remainder.slice(closingParenthesis + 1));

    if (countLetters(after) >= 5) {
      return after;
    }
  }

  return "";
}

function parse2307Name(rawText = "", type = "payee") {
  const lines = cleanText(rawText)
    .split("\n")
    .map(cleanNameLine)
    .filter(Boolean);

  const candidates = [];

  for (const line of lines) {
    const stripped = strip2307NameLabelFromLine(line, type);

    if (
      stripped &&
      countLetters(stripped) >= 5 &&
      !looksLike2307NameLabel(stripped)
    ) {
      candidates.push(stripped);
    }

    if (!looksLike2307NameLabel(line) && countLetters(line) >= 5) {
      candidates.push(line);
    }
  }

  if (!candidates.length) {
    return "";
  }

  const unique = [...new Set(candidates.map(cleanNameLine))];

  const scored = unique.map((line) => {
    const letters = countLetters(line);

    const words = line.split(/\s+/).filter(Boolean).length;

    const uppercase = (line.match(/[A-Z]/g) || []).length;

    const uppercaseRatio = letters ? uppercase / letters : 0;

    return {
      line,

      score: letters + words * 2 + uppercaseRatio * 8,
    };
  });

  scored.sort((a, b) => b.score - a.score);

  return cleanExtractedName(scored[0].line);
}

/* =========================================================
   BIR 2307 OCR - PERIOD
   ========================================================= */

function parse2307OcrPeriod(rawText = "") {
  const normalized = normalizeOcrDigitText(rawText);

  const regex =
    /\b(0?[1-9]|1[0-2])\D{0,10}(0?[1-9]|[12]\d|3[01])\D{0,10}(20\d{2})\b/g;

  const dates = [];

  let match;

  while ((match = regex.exec(normalized)) !== null) {
    dates.push({
      month: Number(match[1]),

      day: Number(match[2]),

      year: Number(match[3]),
    });

    if (dates.length === 2) {
      break;
    }
  }

  if (!dates.length) {
    return {
      from: "",
      to: "",
      quarter: "",
      year: null,
    };
  }

  const from = dates[0];

  const to = dates[1] || dates[0];

  return {
    from: formatDateParts(from.month, from.day, from.year),

    to: formatDateParts(to.month, to.day, to.year),

    quarter: getQuarterFromMonth(to.month),

    year: to.year,
  };
}

/* =========================================================
   BIR 2307 OCR - ATC
   ========================================================= */

function parse2307AtcCell(rawText = "") {
  const text = normalizeOcrDigitText(rawText).replace(/\s+/g, "");

  const full = text.match(/W([IC1L])(\d{3})/);

  if (full) {
    let type = full[1];

    if (type === "1" || type === "L") {
      type = "I";
    }

    return `W${type}${full[2]}`;
  }

  return text.match(/\d{3}/)?.[0] || "";
}

/* =========================================================
   BIR 2307 OCR - MONEY
   ========================================================= */

function parse2307MoneyCell(rawText = "") {
  const corrected = normalizeOcrDigitText(rawText);

  const values =
    corrected.match(/\d{1,3}(?:,\d{3})*(?:\.\d{2})|\d+\.\d{2}/g) || [];

  for (const value of values) {
    const amount = parseAmount(value);

    if (amount !== null && Number.isFinite(amount)) {
      return roundAmount(amount);
    }
  }

  return null;
}

/* =========================================================
   BIR 2307 OCR - NATURE
   ========================================================= */

function parse2307Nature(rawText = "") {
  const text = normalizeText(rawText);

  if (
    (/Gross\s+Income/i.test(text) && /Less\s+than\s+3\s*M/i.test(text)) ||
    (/Non[\s-]*VAT/i.test(text) && /regardless/i.test(text))
  ) {
    return "Gross Income is Less than 3M or Non VAT registered regardless of amount";
  }

  if (/medical\s+practitioners/i.test(text)) {
    return "Payment to medical practitioners through hospital/clinic";
  }

  return cleanExtractedName(text);
}

/* =========================================================
   CANVAS HELPERS
   ========================================================= */

async function renderPdfPageToCanvas(page, scale = PDF_OCR_SCALE) {
  const viewport = page.getViewport({
    scale,
  });

  const canvas = document.createElement("canvas");

  const context = canvas.getContext("2d", {
    willReadFrequently: true,
  });

  if (!context) {
    throw new Error("Unable to create PDF canvas.");
  }

  canvas.width = Math.ceil(viewport.width);

  canvas.height = Math.ceil(viewport.height);

  await page.render({
    canvasContext: context,

    viewport,
  }).promise;

  return canvas;
}

function cropCanvasByRatio(sourceCanvas, region) {
  const sx = Math.max(
    0,

    Math.floor(sourceCanvas.width * region.x),
  );

  const sy = Math.max(
    0,

    Math.floor(sourceCanvas.height * region.y),
  );

  const sw = Math.min(
    sourceCanvas.width - sx,

    Math.ceil(sourceCanvas.width * region.width),
  );

  const sh = Math.min(
    sourceCanvas.height - sy,

    Math.ceil(sourceCanvas.height * region.height),
  );

  const canvas = document.createElement("canvas");

  const context = canvas.getContext("2d", {
    willReadFrequently: true,
  });

  if (!context) {
    throw new Error("Unable to create OCR crop.");
  }

  canvas.width = sw;
  canvas.height = sh;

  context.fillStyle = "#ffffff";

  context.fillRect(0, 0, sw, sh);

  context.drawImage(sourceCanvas, sx, sy, sw, sh, 0, 0, sw, sh);

  return canvas;
}

function upscaleCanvas(sourceCanvas, scale = 2) {
  const canvas = document.createElement("canvas");

  const context = canvas.getContext("2d", {
    willReadFrequently: true,
  });

  if (!context) {
    throw new Error("Unable to enlarge OCR image.");
  }

  canvas.width = Math.max(
    1,

    Math.round(sourceCanvas.width * scale),
  );

  canvas.height = Math.max(
    1,

    Math.round(sourceCanvas.height * scale),
  );

  context.fillStyle = "#ffffff";

  context.fillRect(0, 0, canvas.width, canvas.height);

  context.imageSmoothingEnabled = true;

  context.imageSmoothingQuality = "high";

  context.drawImage(sourceCanvas, 0, 0, canvas.width, canvas.height);

  return canvas;
}

function releaseCanvas(canvas) {
  if (!canvas) {
    return;
  }

  canvas.width = 0;
  canvas.height = 0;
}

/* =========================================================
   OCR SETTINGS
   ========================================================= */

async function setOcrParameters(worker, { psm = 6, whitelist = "" } = {}) {
  try {
    await worker.setParameters({
      tessedit_pageseg_mode: String(psm),

      tessedit_char_whitelist: whitelist,

      preserve_interword_spaces: "1",

      user_defined_dpi: "300",
    });
  } catch (error) {
    console.warn("Unable to apply OCR parameters:", error);
  }
}

/* =========================================================
   OCR A REGION
   ========================================================= */

async function recognize2307Region(
  sourceCanvas,
  region,
  worker,
  { psm = 6, whitelist = "", upscale = 2 } = {},
) {
  let crop = null;

  let enlarged = null;

  try {
    crop = cropCanvasByRatio(sourceCanvas, region);

    enlarged = upscaleCanvas(crop, upscale);

    await setOcrParameters(worker, {
      psm,
      whitelist,
    });

    const result = await worker.recognize(enlarged);

    return result?.data?.text?.trim() || "";
  } finally {
    releaseCanvas(enlarged);

    releaseCanvas(crop);
  }
}

/* =========================================================
   OCR TAXPAYER / PAYOR TIN
   ========================================================= */

async function recognize2307Tin(sourceCanvas, region, worker) {
  /*
    PASS 1:
    numbers only.
  */
  const numericPass = await recognize2307Region(sourceCanvas, region, worker, {
    psm: 6,

    whitelist: "0123456789- ",

    upscale: 3,
  });

  const numericTin = parse2307OcrTin(numericPass);

  if (numericTin) {
    return {
      rawText: numericPass,

      tin: numericTin,
    };
  }

  /*
    PASS 2:
    allow OCR-confusable
    letters.
  */
  const relaxedPass = await recognize2307Region(sourceCanvas, region, worker, {
    psm: 6,

    whitelist: "0123456789OQDILZSBG-| ",

    upscale: 4,
  });

  return {
    rawText: relaxedPass,

    tin: parse2307OcrTin(relaxedPass),
  };
}

/* =========================================================
   OCR NAMES
   ========================================================= */

async function recognize2307Name(sourceCanvas, region, worker, type) {
  const rawText = await recognize2307Region(sourceCanvas, region, worker, {
    psm: 6,

    /*
          Important:
          no whitelist here.
          The field contains letters.
        */
    whitelist: "",

    upscale: 2,
  });

  return {
    rawText,

    name: parse2307Name(rawText, type),
  };
}

/* =========================================================
   STRUCTURED 2307 RESULT
   ========================================================= */

function build2307StructuredData({
  periodText,

  payeeTin,
  payeeName,

  payorTin,
  payorName,

  natureText,

  atcText,

  totalText,

  taxWithheldText,
}) {
  const period = parse2307OcrPeriod(periodText);

  const incomePayment = parse2307MoneyCell(totalText);

  const taxWithheld = parse2307MoneyCell(taxWithheldText);

  let taxRate = null;

  if (incomePayment !== null && incomePayment > 0 && taxWithheld !== null) {
    taxRate = Math.round((taxWithheld / incomePayment) * 100 * 100) / 100;
  }

  return {
    form_type: "BIR Form 2307",

    /*
      PART I - PAYEE
    */
    taxpayer: {
      name: payeeName || "",

      tin: payeeTin || "",
    },

    /*
      PART II - PAYOR
    */
    withholding_agent: {
      name: payorName || "",

      tin: payorTin || "",
    },

    tax_period: {
      quarter: period.quarter,

      year: period.year,

      from: period.from,

      to: period.to,
    },

    withholding: {
      nature_of_income_payment: parse2307Nature(natureText),

      atc: parse2307AtcCell(atcText),

      income_payment: incomePayment,

      tax_rate: taxRate,

      tax_withheld: taxWithheld,
    },

    amounts: {
      gross_sales: incomePayment,

      taxable_income: null,

      tax_due: taxWithheld,
    },
  };
}

/* =========================================================
   IDENTITY FALLBACK
   ========================================================= */

function merge2307WithReference(current, reference) {
  if (!reference || current?.form_type !== "BIR Form 2307") {
    return current;
  }

  return {
    ...current,

    taxpayer: {
      ...current.taxpayer,

      name: current.taxpayer?.name || reference.taxpayer?.name || "",

      tin: current.taxpayer?.tin || reference.taxpayer?.tin || "",
    },

    withholding_agent: {
      ...current.withholding_agent,

      name:
        current.withholding_agent?.name ||
        reference.withholding_agent?.name ||
        "",

      tin:
        current.withholding_agent?.tin ||
        reference.withholding_agent?.tin ||
        "",
    },

    withholding: {
      ...current.withholding,

      nature_of_income_payment:
        current.withholding?.nature_of_income_payment ||
        reference.withholding?.nature_of_income_payment ||
        "",

      atc: current.withholding?.atc || reference.withholding?.atc || "",
    },
  };
}

/* =========================================================
   BIR 2307 PAGE OCR
   ========================================================= */

async function ocr2307Page(page, getOcrWorker, onProgress) {
  const worker = await getOcrWorker();

  let sourceCanvas = null;

  try {
    onProgress?.("Rendering BIR Form 2307");

    sourceCanvas = await renderPdfPageToCanvas(page, PDF_OCR_SCALE);

    /* PERIOD */

    onProgress?.("Reading tax period");

    const periodText = await recognize2307Region(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.period,
      worker,
      {
        psm: 6,

        whitelist: "0123456789/-. ",

        upscale: 2,
      },
    );

    /* TAXPAYER TIN */

    onProgress?.("Reading Taxpayer Identification Number");

    const payeeTinResult = await recognize2307Tin(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.payeeTin,
      worker,
    );

    /* PAYEE NAME */

    onProgress?.("Reading Payee name");

    const payeeNameResult = await recognize2307Name(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.payeeName,
      worker,
      "payee",
    );

    /* PAYOR TIN */

    onProgress?.("Reading Withholding Agent TIN");

    const payorTinResult = await recognize2307Tin(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.payorTin,
      worker,
    );

    /* PAYOR NAME */

    onProgress?.("Reading Payor name");

    const payorNameResult = await recognize2307Name(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.payorName,
      worker,
      "payor",
    );

    /* NATURE */

    onProgress?.("Reading nature of income payment");

    const natureText = await recognize2307Region(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.nature,
      worker,
      {
        psm: 6,

        whitelist: "",

        upscale: 2,
      },
    );

    /* ATC */

    onProgress?.("Reading ATC");

    const atcText = await recognize2307Region(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.atc,
      worker,
      {
        psm: 7,

        whitelist: "0123456789WICL",

        upscale: 3,
      },
    );

    /* INCOME PAYMENT */

    onProgress?.("Reading income payment");

    const totalText = await recognize2307Region(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.total,
      worker,
      {
        psm: 7,

        whitelist: "0123456789,.",

        upscale: 3,
      },
    );

    /* TAX WITHHELD */

    onProgress?.("Reading tax withheld");

    const taxWithheldText = await recognize2307Region(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.taxWithheld,
      worker,
      {
        psm: 7,

        whitelist: "0123456789,.",

        upscale: 3,
      },
    );

    const data = build2307StructuredData({
      periodText,

      payeeTin: payeeTinResult.tin,

      payeeName: payeeNameResult.name,

      payorTin: payorTinResult.tin,

      payorName: payorNameResult.name,

      natureText,

      atcText,

      totalText,

      taxWithheldText,
    });

    /*
      DEBUG
      Open Chrome DevTools > Console.
    */

    console.group(`BIR 2307 OCR - page ${page.pageNumber || ""}`);

    console.log("PAYEE NAME RAW:", payeeNameResult.rawText);

    console.log("PAYEE NAME:", payeeNameResult.name);

    console.log("TAXPAYER TIN RAW:", payeeTinResult.rawText);

    console.log("TAXPAYER TIN:", payeeTinResult.tin);

    console.log("PAYOR NAME RAW:", payorNameResult.rawText);

    console.log("PAYOR NAME:", payorNameResult.name);

    console.log("WITHHOLDING AGENT TIN RAW:", payorTinResult.rawText);

    console.log("WITHHOLDING AGENT TIN:", payorTinResult.tin);

    console.table({
      period: periodText,

      taxpayerName: payeeNameResult.name,

      taxpayerTIN: payeeTinResult.tin,

      withholdingAgentName: payorNameResult.name,

      withholdingAgentTIN: payorTinResult.tin,

      nature: natureText,

      atc: atcText,

      incomePayment: totalText,

      taxWithheld: taxWithheldText,
    });

    console.log("FINAL DATA:", data);

    console.groupEnd();

    return {
      text: cleanText(`
          BIR Form 2307

          ${periodText}

          ${payeeTinResult.tin}

          ${payeeNameResult.name}

          ${payorTinResult.tin}

          ${payorNameResult.name}

          ${natureText}

          ${atcText}

          ${totalText}

          ${taxWithheldText}
        `),

      data,

      method: "ocr-2307-exact-fields",
    };
  } finally {
    releaseCanvas(sourceCanvas);
  }
}

/* =========================================================
   BASIC GENERIC PARSER
   ========================================================= */

function findFormType(text) {
  const match = normalizeText(text).match(
    /\b(2307|1701Q|1701A|1701|1601[\s-]?EQ|2551Q|2550Q)\b/i,
  );

  if (!match) {
    return "";
  }

  let form = match[1].toUpperCase().replace(/\s/g, "");

  if (form === "1601EQ") {
    form = "1601-EQ";
  }

  return `BIR Form ${form}`;
}

function findGenericTin(text) {
  const match = normalizeText(text).match(
    /\b(\d{3})[\s-]+(\d{3})[\s-]+(\d{3})(?:[\s-]+(\d{3,5}))?\b/,
  );

  if (!match) {
    return "";
  }

  return formatTin(`${match[1]}${match[2]}${match[3]}${match[4] || ""}`);
}

function findGenericAmount(text, label) {
  const match = normalizeText(text).match(
    new RegExp(
      `${label}[\\s\\S]{0,120}?(?:PHP|₱)?\\s*([0-9][0-9,]*(?:\\.\\d{1,2})?)`,
      "i",
    ),
  );

  return match?.[1] ? parseAmount(match[1]) : null;
}

function extractGenericTextFields(text) {
  return {
    form_type: findFormType(text),

    taxpayer: {
      name: "",

      tin: findGenericTin(text),
    },

    withholding_agent: {
      name: "",
      tin: "",
    },

    tax_period: {
      quarter: "",
      year: null,
      from: "",
      to: "",
    },

    withholding: {
      nature_of_income_payment: "",

      atc: "",

      income_payment: null,

      tax_rate: null,

      tax_withheld: null,
    },

    amounts: {
      gross_sales: findGenericAmount(text, "Gross\\s+(?:Sales|Receipts)"),

      taxable_income: findGenericAmount(text, "Taxable\\s+Income"),

      tax_due: findGenericAmount(text, "Tax\\s+Due"),
    },
  };
}

/* =========================================================
   QUALITY SCORE
   ========================================================= */

function scoreExtractedData(data) {
  let score = 0;

  if (data.form_type) {
    score += 2;
  }

  if (data.taxpayer?.name) {
    score += 2;
  }

  if (data.taxpayer?.tin) {
    score += 2;
  }

  if (data.withholding_agent?.name) {
    score += 2;
  }

  if (data.withholding_agent?.tin) {
    score += 2;
  }

  if (data.tax_period?.from) {
    score += 1;
  }

  if (data.tax_period?.to) {
    score += 1;
  }

  if (data.tax_period?.year) {
    score += 1;
  }

  if (data.withholding?.atc) {
    score += 1;
  }

  if (data.withholding?.income_payment !== null) {
    score += 2;
  }

  if (data.withholding?.tax_withheld !== null) {
    score += 2;
  }

  if (data.amounts?.gross_sales !== null) {
    score += 1;
  }

  if (data.amounts?.taxable_income !== null) {
    score += 1;
  }

  if (data.amounts?.tax_due !== null) {
    score += 1;
  }

  return score;
}

function hasUsefulData(data) {
  return scoreExtractedData(data) >= 4;
}

/* =========================================================
   SAFE PDF CLEANUP
   ========================================================= */

async function destroyPdfSafely(loadingTask, pdf) {
  try {
    if (typeof pdf?.cleanup === "function") {
      pdf.cleanup();
    }
  } catch (error) {
    console.warn("PDF cleanup warning:", error);
  }

  try {
    /*
      Preferred cleanup.

      This avoids the earlier:
      pdf.destroy is not a function
      failure.
    */
    if (typeof loadingTask?.destroy === "function") {
      await loadingTask.destroy();

      return;
    }

    if (typeof pdf?.destroy === "function") {
      await pdf.destroy();
    }
  } catch (error) {
    /*
      Cleanup errors must not
      fail extracted data.
    */
    console.warn("PDF destroy warning:", error);
  }
}

/* =========================================================
   PDF EXTRACTION
   ========================================================= */

async function extractPdfPages(file, { getOcrWorker, onProgress }) {
  const arrayBuffer = await file.arrayBuffer();

  const loadingTask = pdfjsLib.getDocument({
    data: arrayBuffer,
  });

  let pdf = null;

  const pages = [];

  let identityReference = null;

  try {
    pdf = await loadingTask.promise;

    const pageLimit = Math.min(pdf.numPages, PDF_TEST_PAGE_LIMIT);

    for (let pageNumber = 1; pageNumber <= pageLimit; pageNumber += 1) {
      onProgress?.(`Reading page ${pageNumber} of ${pageLimit}`);

      const page = await pdf.getPage(pageNumber);

      try {
        const textContent = await page.getTextContent();

        const directText = textContent.items
          .map((item) => item.str)
          .join(" ")
          .trim();

        const formType = findFormType(directText);

        /*
          Sample filename contains 2307.
          The text-layer fallback is also checked.
        */
        const use2307Ocr =
          /2307/i.test(file.name) || formType === "BIR Form 2307";

        let text = directText;

        let data = null;

        let method = "text";

        if (use2307Ocr) {
          const result = await ocr2307Page(
            page,

            getOcrWorker,

            (message) => {
              onProgress?.(`${message} - page ${pageNumber} of ${pageLimit}`);
            },
          );

          text = result.text;

          data = result.data;

          method = result.method;

          /*
            Save first complete
            identity.

            The first 15 pages use the
            same Payee/Payor.
          */
          if (
            !identityReference &&
            data.taxpayer?.name &&
            isValidExtractedTin(data.taxpayer?.tin) &&
            data.withholding_agent?.name &&
            isValidExtractedTin(data.withholding_agent?.tin)
          ) {
            identityReference = data;
          }

          /*
            If OCR misses a name/TIN on
            a later page, reuse the
            already confirmed identity.
          */
          if (identityReference && pageNumber > 1) {
            data = merge2307WithReference(data, identityReference);
          }
        } else {
          data = extractGenericTextFields(directText);
        }

        pages.push({
          pageNumber,

          text: cleanText(text),

          method,

          /*
            Keep structured OCR.
            Do not parse the OCR text
            again later.
          */
          preParsedData: data,

          sourceTotalPages: pdf.numPages,

          processedPageCount: pageLimit,
        });
      } finally {
        try {
          if (typeof page?.cleanup === "function") {
            page.cleanup();
          }
        } catch (error) {
          console.warn(`Page ${pageNumber} cleanup warning:`, error);
        }
      }
    }
  } finally {
    await destroyPdfSafely(loadingTask, pdf);
  }

  return pages;
}

/* =========================================================
   DOCX
   ========================================================= */

function splitDocxTextIntoRecords(rawText) {
  const text = cleanText(rawText);

  if (!text) {
    return [];
  }

  const regex =
    /\b(?:BIR\s+Form(?:\s+No\.?)?\s*)?(?:2307|1701Q|1701A|1701|1601[\s-]?EQ|2551Q|2550Q)\b/gi;

  const indexes = [];

  let match;

  while ((match = regex.exec(text)) !== null) {
    const last = indexes[indexes.length - 1];

    if (last === undefined || match.index - last > 500) {
      indexes.push(match.index);
    }
  }

  if (indexes.length <= 1) {
    return [
      {
        pageNumber: 1,

        text,

        method: "docx",
      },
    ];
  }

  if (indexes[0] > 0) {
    indexes[0] = 0;
  }

  return indexes
    .map((start, index) => {
      const end = index + 1 < indexes.length ? indexes[index + 1] : text.length;

      return cleanText(text.slice(start, end));
    })
    .filter(Boolean)
    .map((recordText, index) => ({
      pageNumber: index + 1,

      text: recordText,

      method: "docx",
    }));
}

async function extractDocxPages(file, { onProgress }) {
  onProgress?.("Reading Word document");

  const arrayBuffer = await file.arrayBuffer();

  const result = await mammoth.extractRawText({
    arrayBuffer,
  });

  const all = splitDocxTextIntoRecords(result.value || "");

  const limited = all.slice(0, PDF_TEST_PAGE_LIMIT);

  return limited.map((record) => ({
    ...record,

    preParsedData: extractGenericTextFields(record.text),

    sourceTotalPages: all.length,

    processedPageCount: limited.length,
  }));
}

/* =========================================================
   DOCUMENT DISPATCHER
   ========================================================= */

async function extractDocumentPages(file, helpers) {
  const extension = getFileExtension(file.name);

  if (extension === ".pdf") {
    return extractPdfPages(file, helpers);
  }

  if (extension === ".docx") {
    return extractDocxPages(file, helpers);
  }

  throw new Error("Unsupported document type.");
}

/* =========================================================
   EXCEL ROWS
   ========================================================= */

function getExcelRows(records) {
  return records.map(({ data }) => ({
    "Form Type": data.form_type || "",

    "Taxpayer / Payee Name": data.taxpayer?.name || "",

    /*
        PART I
      */
    "Taxpayer Identification Number (TIN)": data.taxpayer?.tin || "",

    "Payor / Withholding Agent": data.withholding_agent?.name || "",

    /*
        PART II
      */
    "Withholding Agent Identification Number (TIN)":
      data.withholding_agent?.tin || "",

    Quarter: data.tax_period?.quarter || "",

    "Tax Year": data.tax_period?.year || "",

    "Period From": data.tax_period?.from || "",

    "Period To": data.tax_period?.to || "",

    "Nature of Income Payment":
      data.withholding?.nature_of_income_payment || "",

    ATC: data.withholding?.atc || "",

    "Income Payment": data.withholding?.income_payment ?? "",

    "Tax Rate (%)": data.withholding?.tax_rate ?? "",

    "Tax Withheld": data.withholding?.tax_withheld ?? "",

    "Gross Sales": data.amounts?.gross_sales ?? "",

    "Taxable Income": data.amounts?.taxable_income ?? "",

    "Tax Due": data.amounts?.tax_due ?? "",
  }));
}

/* =========================================================
   EXCEL EXPORT
   ========================================================= */

function exportRecordsToExcel(records) {
  const rows = getExcelRows(records);

  const worksheet = XLSX.utils.json_to_sheet(rows);

  worksheet["!cols"] = [
    { wch: 18 },
    { wch: 34 },

    /*
      Taxpayer TIN
    */
    { wch: 34 },

    { wch: 34 },

    /*
      Withholding Agent TIN
    */
    { wch: 40 },

    { wch: 16 },
    { wch: 12 },
    { wch: 15 },
    { wch: 15 },
    { wch: 55 },
    { wch: 12 },
    { wch: 18 },
    { wch: 15 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
  ];

  /*
    L = Income Payment
    M = Tax Rate
    N = Tax Withheld
    O = Gross Sales
    P = Taxable Income
    Q = Tax Due
  */

  for (let row = 2; row <= rows.length + 1; row += 1) {
    ["L", "N", "O", "P", "Q"].forEach((column) => {
      const cell = worksheet[`${column}${row}`];

      if (cell && typeof cell.v === "number") {
        cell.z = "#,##0.00";
      }
    });

    const rateCell = worksheet[`M${row}`];

    if (rateCell && typeof rateCell.v === "number") {
      rateCell.z = "0.00";
    }
  }

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, worksheet, "Tax Records");

  const now = new Date();

  const date = [
    now.getFullYear(),

    String(now.getMonth() + 1).padStart(2, "0"),

    String(now.getDate()).padStart(2, "0"),
  ].join("-");

  XLSX.writeFile(workbook, `tax-records-${date}.xlsx`);
}

/* =========================================================
   MAIN COMPONENT
   ========================================================= */

export default function TaxEntry() {
  const fileInputRef = useRef(null);

  const [entryMode, setEntryMode] = useState("manual");

  const [files, setFiles] = useState([]);

  const [fileError, setFileError] = useState("");

  const [isDragging, setIsDragging] = useState(false);

  const [isExtracting, setIsExtracting] = useState(false);

  const [fileStatuses, setFileStatuses] = useState({});

  const [extractedRecords, setExtractedRecords] = useState([]);

  const {
    register,

    control,

    handleSubmit,

    reset,

    formState: { errors, isSubmitting, isDirty },
  } = useForm({
    defaultValues: DEFAULT_VALUES,

    mode: "onBlur",
  });

  const uploadMode = entryMode === "upload";

  const totalFileSize = files.reduce((total, file) => total + file.size, 0);

  /* =======================================================
     FILE STATUS
     ======================================================= */

  function updateFileStatus(file, status, message = "") {
    setFileStatuses((current) => ({
      ...current,

      [getFileKey(file)]: {
        status,

        message,
      },
    }));
  }

  /* =======================================================
     EXTRACT DOCUMENTS
     ======================================================= */

  async function extractDocuments(documents) {
    if (!documents.length || isExtracting) {
      return;
    }

    setIsExtracting(true);

    openExtractionLoading(documents.length);

    let ocrWorker = null;

    async function getOcrWorker() {
      if (!ocrWorker) {
        updateExtractionLoading(
          "Starting OCR engine",

          "Tesseract.js is loading once and will be reused for all test pages.",
        );

        ocrWorker = await createWorker("eng");
      }

      return ocrWorker;
    }

    const extractedNow = [];

    const processedKeys = new Set(documents.map(getFileKey));

    let reviewPageCount = 0;

    let failedFileCount = 0;

    let totalPageCount = 0;

    try {
      for (let fileIndex = 0; fileIndex < documents.length; fileIndex += 1) {
        const file = documents[fileIndex];

        const fileKey = getFileKey(file);

        try {
          updateFileStatus(file, "reading", "Reading document");

          const pages = await extractDocumentPages(file, {
            getOcrWorker,

            onProgress: (message) => {
              updateFileStatus(file, "reading", message);

              updateExtractionLoading(
                message,

                `${file.name} • File ${fileIndex + 1} of ${documents.length}`,
              );
            },
          });

          totalPageCount += pages.length;

          const sourceTotalPages = pages[0]?.sourceTotalPages ?? pages.length;

          let fileRecordCount = 0;

          let fileReviewCount = 0;

          for (const page of pages) {
            /*
              IMPORTANT:
              use preParsedData.

              This keeps the accurate
              field-by-field OCR result.
            */
            const data =
              page.preParsedData || extractGenericTextFields(page.text);

            console.debug(
              `[Tax extraction] Page ${page.pageNumber}`,

              {
                method: page.method,

                data,
              },
            );

            if (!hasUsefulData(data)) {
              fileReviewCount += 1;

              reviewPageCount += 1;

              continue;
            }

            extractedNow.push({
              fileKey,

              pageNumber: page.pageNumber,

              data,
            });

            fileRecordCount += 1;
          }

          const limitText =
            sourceTotalPages > pages.length
              ? ` • testing first ${pages.length} of ${sourceTotalPages} pages`
              : "";

          if (!fileRecordCount) {
            updateFileStatus(
              file,
              "review",

              `No usable records found${limitText}`,
            );
          } else if (fileReviewCount) {
            updateFileStatus(
              file,
              "partial",

              `${fileRecordCount} extracted • ${fileReviewCount} need review${limitText}`,
            );
          } else {
            updateFileStatus(
              file,
              "extracted",

              `${fileRecordCount} record${
                fileRecordCount === 1 ? "" : "s"
              } extracted${limitText}`,
            );
          }
        } catch (error) {
          console.error(
            `Unable to read ${file.name}`,

            error,
          );

          failedFileCount += 1;

          updateFileStatus(
            file,
            "failed",

            error?.message || "Unable to read document",
          );
        }
      }

      setExtractedRecords((current) => [
        ...current.filter((record) => !processedKeys.has(record.fileKey)),

        ...extractedNow,
      ]);
    } finally {
      if (ocrWorker) {
        try {
          await ocrWorker.terminate();
        } catch (error) {
          console.warn(
            "OCR worker termination warning:",

            error,
          );
        }
      }

      setIsExtracting(false);

      Swal.close();
    }

    if (!extractedNow.length) {
      await Swal.fire({
        icon: "warning",

        title: "No tax records extracted",

        text:
          failedFileCount > 0
            ? `${failedFileCount} file(s) could not be processed. Check the browser console for the exact error.`
            : "The documents were read, but no supported tax information was detected.",

        confirmButtonColor: "#18181b",
      });

      return;
    }

    await Swal.fire({
      icon: reviewPageCount || failedFileCount ? "warning" : "success",

      title: "Extraction complete",

      html: `
        <div style="font-size:14px;line-height:1.7;color:#52525b;">
          <div>
            <strong>${extractedNow.length}</strong>
            record${extractedNow.length === 1 ? "" : "s"} extracted.
          </div>

          <div>
            <strong>${totalPageCount}</strong>
            page${totalPageCount === 1 ? "" : "s"} processed.
          </div>

          ${
            reviewPageCount
              ? `
                <div>
                  <strong>${reviewPageCount}</strong>
                  page(s) need review.
                </div>
              `
              : ""
          }

          ${
            failedFileCount
              ? `
                <div>
                  <strong>${failedFileCount}</strong>
                  file(s) failed.
                </div>
              `
              : ""
          }

          <div style="margin-top:8px;font-size:12px;color:#71717a;">
            Test mode processes a maximum of ${PDF_TEST_PAGE_LIMIT} pages per document.
          </div>
        </div>
      `,

      confirmButtonColor: "#18181b",
    });
  }

  /* =======================================================
     VALIDATE FILES
     ======================================================= */

  function validateAndAddFiles(incomingFiles) {
    const incoming = Array.from(incomingFiles);

    if (!incoming.length) {
      return;
    }

    setFileError("");

    const existingKeys = new Set(files.map(getFileKey));

    const nextFiles = [...files];

    const accepted = [];

    const rejected = [];

    for (const file of incoming) {
      if (!isAllowedFile(file)) {
        rejected.push(`${file.name}: only PDF and DOCX are supported`);

        continue;
      }

      if (file.size > MAX_FILE_SIZE) {
        rejected.push(`${file.name}: exceeds 20 MB`);

        continue;
      }

      if (existingKeys.has(getFileKey(file))) {
        rejected.push(`${file.name}: already selected`);

        continue;
      }

      if (nextFiles.length >= MAX_FILES) {
        rejected.push(`${file.name}: maximum ${MAX_FILES} files`);

        continue;
      }

      const total =
        nextFiles.reduce((sum, selectedFile) => sum + selectedFile.size, 0) +
        file.size;

      if (total > MAX_TOTAL_SIZE) {
        rejected.push(`${file.name}: total upload exceeds 100 MB`);

        continue;
      }

      existingKeys.add(getFileKey(file));

      nextFiles.push(file);

      accepted.push(file);
    }

    setFiles(nextFiles);

    if (accepted.length) {
      void extractDocuments(accepted);
    }

    if (rejected.length) {
      setFileError(rejected.join(". "));

      Toast.fire({
        icon: "warning",

        title:
          rejected.length === 1
            ? "1 file was not added"
            : `${rejected.length} files were not added`,
      });
    }
  }

  function handleFileInput(event) {
    validateAndAddFiles(event.target.files);

    event.target.value = "";
  }

  function handleDrop(event) {
    event.preventDefault();

    event.stopPropagation();

    setIsDragging(false);

    validateAndAddFiles(event.dataTransfer.files);
  }

  /* =======================================================
     REMOVE FILE
     ======================================================= */

  function removeFile(fileToRemove) {
    const fileKey = getFileKey(fileToRemove);

    setFiles((current) =>
      current.filter((file) => getFileKey(file) !== fileKey),
    );

    setFileStatuses((current) => {
      const next = {
        ...current,
      };

      delete next[fileKey];

      return next;
    });

    setExtractedRecords((current) =>
      current.filter((record) => record.fileKey !== fileKey),
    );
  }

  function clearFiles() {
    setFiles([]);

    setFileStatuses({});

    setExtractedRecords([]);

    setFileError("");
  }

  /* =======================================================
     SAVE UPLOADED EXCEL
     ======================================================= */

  async function handleSaveUploadedExcel() {
    if (!extractedRecords.length) {
      Toast.fire({
        icon: "warning",

        title: "No extracted records to save",
      });

      return;
    }

    const confirmation = await Swal.fire({
      title: "Save extracted records?",

      text: `${extractedRecords.length} tax record${
        extractedRecords.length === 1 ? "" : "s"
      } will be exported into one Excel workbook.`,

      icon: "question",

      showCancelButton: true,

      reverseButtons: true,

      confirmButtonText: "Save to Excel",

      cancelButtonText: "Cancel",

      confirmButtonColor: "#18181b",

      cancelButtonColor: "#71717a",
    });

    if (!confirmation.isConfirmed) {
      return;
    }

    exportRecordsToExcel(extractedRecords);

    Toast.fire({
      icon: "success",

      title: `${extractedRecords.length} record${
        extractedRecords.length === 1 ? "" : "s"
      } saved to Excel`,
    });
  }

  /* =======================================================
     MANUAL SUBMIT
     ======================================================= */

  async function onManualSubmit(data) {
    const confirmation = await Swal.fire({
      title: "Save tax record?",

      text: "This tax record will be exported to Excel.",

      icon: "question",

      showCancelButton: true,

      reverseButtons: true,

      confirmButtonText: "Save to Excel",

      cancelButtonText: "Review Again",

      confirmButtonColor: "#18181b",

      cancelButtonColor: "#71717a",
    });

    if (!confirmation.isConfirmed) {
      return;
    }

    const record = {
      fileKey: "manual",

      pageNumber: null,

      data: {
        form_type: data.formType,

        taxpayer: {
          name: data.taxpayerName.trim().replace(/\s+/g, " "),

          tin: data.tin,
        },

        withholding_agent: {
          name: "",
          tin: "",
        },

        tax_period: {
          quarter: data.quarter,

          year: String(data.year.getFullYear()),

          from: "",

          to: "",
        },

        withholding: {
          nature_of_income_payment: "",

          atc: "",

          income_payment: null,

          tax_rate: null,

          tax_withheld: null,
        },

        amounts: {
          gross_sales: Number(data.grossSales),

          taxable_income: Number(data.taxableIncome),

          tax_due: Number(data.taxDue),
        },
      },
    };

    exportRecordsToExcel([record]);

    Toast.fire({
      icon: "success",

      title: "Tax record saved to Excel",
    });
  }

  function onInvalid() {
    Toast.fire({
      icon: "error",

      title: "Please check the highlighted fields",
    });
  }

  async function handleClearManual() {
    if (!isDirty) {
      reset(DEFAULT_VALUES);

      return;
    }

    const confirmation = await Swal.fire({
      title: "Clear form?",

      text: "All entered information will be removed.",

      icon: "warning",

      showCancelButton: true,

      reverseButtons: true,

      confirmButtonText: "Clear Form",

      cancelButtonText: "Keep Editing",

      confirmButtonColor: "#18181b",

      cancelButtonColor: "#71717a",
    });

    if (confirmation.isConfirmed) {
      reset(DEFAULT_VALUES);
    }
  }

  /* =======================================================
     JSX
     ======================================================= */

  return (
    <main className="min-h-screen bg-zinc-50 px-4 py-6 sm:px-6 md:py-10">
      <div className="mx-auto max-w-5xl">
        {/* HEADER */}

        <header className="mb-7">
          <p className="text-sm font-medium text-zinc-500">Tax Management</p>

          <h1 className="mt-1 text-2xl font-semibold tracking-tight text-zinc-950 sm:text-3xl">
            Add Tax Record
          </h1>

          <p className="mt-2 max-w-2xl text-sm leading-6 text-zinc-600">
            Enter tax information manually or extract multiple tax records from
            PDF and Word documents.
          </p>
        </header>

        {/* MODE */}

        <div className="mb-6 grid gap-3 sm:grid-cols-2">
          <ModeButton
            active={entryMode === "manual"}
            title="Enter Manually"
            description="Enter one tax record manually and save it to Excel."
            onClick={() => setEntryMode("manual")}
          />

          <ModeButton
            active={entryMode === "upload"}
            title="Upload Documents"
            description={`Upload PDF or DOCX files. Test mode processes only the first ${PDF_TEST_PAGE_LIMIT} pages.`}
            onClick={() => setEntryMode("upload")}
          />
        </div>

        {/* UPLOAD */}

        {uploadMode ? (
          <section className="overflow-hidden rounded-xl border border-zinc-200 bg-white shadow-sm">
            <div className="border-b border-zinc-200 p-5 sm:p-6 md:p-7">
              <h2 className="text-base font-semibold text-zinc-950">
                Upload Tax Documents
              </h2>

              <p className="mt-1 text-sm leading-5 text-zinc-500">
                BIR Form 2307 scans Payee name, Taxpayer TIN, Payor name,
                Withholding Agent TIN, period, ATC, income payment, and tax
                withheld separately for better accuracy.
              </p>
            </div>

            <div className="p-5 sm:p-6 md:p-7">
              <input
                ref={fileInputRef}
                type="file"
                multiple
                accept=".pdf,.docx,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                onChange={handleFileInput}
                className="hidden"
              />

              {/* DROPZONE */}

              <div
                role="button"
                tabIndex={0}
                onClick={() => {
                  if (!isExtracting) {
                    fileInputRef.current?.click();
                  }
                }}
                onKeyDown={(event) => {
                  if (
                    !isExtracting &&
                    (event.key === "Enter" || event.key === " ")
                  ) {
                    event.preventDefault();

                    fileInputRef.current?.click();
                  }
                }}
                onDragEnter={(event) => {
                  event.preventDefault();

                  event.stopPropagation();

                  if (!isExtracting) {
                    setIsDragging(true);
                  }
                }}
                onDragOver={(event) => {
                  event.preventDefault();

                  event.stopPropagation();
                }}
                onDragLeave={(event) => {
                  event.preventDefault();

                  event.stopPropagation();

                  setIsDragging(false);
                }}
                onDrop={(event) => {
                  if (isExtracting) {
                    event.preventDefault();

                    return;
                  }

                  handleDrop(event);
                }}
                className={`
                  flex
                  min-h-56
                  flex-col
                  items-center
                  justify-center
                  rounded-lg
                  border-2
                  border-dashed
                  px-6
                  py-8
                  text-center
                  outline-none
                  transition

                  ${
                    isExtracting
                      ? "cursor-not-allowed border-zinc-200 bg-zinc-50 opacity-60"
                      : "cursor-pointer focus:ring-2 focus:ring-zinc-900 focus:ring-offset-2"
                  }

                  ${
                    isDragging
                      ? "border-zinc-900 bg-zinc-100"
                      : "border-zinc-300 bg-zinc-50 hover:border-zinc-400 hover:bg-zinc-100/70"
                  }
                `}
              >
                {isExtracting ? (
                  <>
                    <span className="h-8 w-8 animate-spin rounded-full border-2 border-zinc-300 border-t-zinc-900" />

                    <p className="mt-4 text-sm font-semibold text-zinc-900">
                      Extracting documents...
                    </p>

                    <p className="mt-1 max-w-md text-xs leading-5 text-zinc-500">
                      OCR is reading the BIR fields individually.
                    </p>
                  </>
                ) : (
                  <>
                    <div className="flex h-11 w-11 items-center justify-center rounded-lg border border-zinc-200 bg-white text-zinc-600 shadow-sm">
                      <UploadIcon />
                    </div>

                    <p className="mt-4 text-sm font-semibold text-zinc-900">
                      Drop documents here
                    </p>

                    <p className="mt-1 text-sm text-zinc-500">
                      or{" "}
                      <span className="font-medium text-zinc-900 underline underline-offset-2">
                        browse files
                      </span>
                    </p>

                    <div className="mt-4 flex flex-wrap justify-center gap-x-2 gap-y-1 text-xs text-zinc-500">
                      <span>PDF, DOCX</span>

                      <span>•</span>

                      <span>20 MB each</span>

                      <span>•</span>

                      <span>100 MB total</span>

                      <span>•</span>

                      <span>First {PDF_TEST_PAGE_LIMIT} pages</span>
                    </div>
                  </>
                )}
              </div>

              {/* ERROR */}

              {fileError && (
                <div className="mt-4 rounded-md border border-red-200 bg-red-50 px-4 py-3 text-sm leading-5 text-red-700">
                  {fileError}
                </div>
              )}

              {/* FILE LIST */}

              {files.length > 0 && (
                <div className="mt-5 overflow-hidden rounded-lg border border-zinc-200">
                  <div className="flex flex-col gap-3 border-b border-zinc-200 bg-zinc-50 px-4 py-3 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                      <p className="text-sm font-medium text-zinc-900">
                        Documents
                      </p>

                      <p className="mt-0.5 text-xs text-zinc-500">
                        {files.length} file
                        {files.length === 1 ? "" : "s"} •{" "}
                        {formatFileSize(totalFileSize)}
                      </p>
                    </div>

                    <button
                      type="button"
                      disabled={isExtracting}
                      onClick={clearFiles}
                      className="w-fit text-sm font-medium text-zinc-500 transition hover:text-red-600 disabled:opacity-40"
                    >
                      Remove all
                    </button>
                  </div>

                  <ul className="divide-y divide-zinc-200">
                    {files.map((file) => (
                      <FileItem
                        key={getFileKey(file)}
                        file={file}
                        state={fileStatuses[getFileKey(file)]}
                        disabled={isExtracting}
                        onRemove={() => removeFile(file)}
                      />
                    ))}
                  </ul>
                </div>
              )}

              {/* SUMMARY */}

              {files.length > 0 && (
                <div className="mt-6 rounded-lg border border-zinc-200 bg-zinc-50 p-4">
                  <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                      <p className="text-sm font-semibold text-zinc-900">
                        {extractedRecords.length} extracted record
                        {extractedRecords.length === 1 ? "" : "s"}
                      </p>

                      <p className="mt-1 text-xs leading-5 text-zinc-500">
                        Taxpayer TIN and Withholding Agent TIN are exported as
                        separate columns.
                      </p>
                    </div>

                    <div className="flex flex-col-reverse gap-3 sm:flex-row">
                      <button
                        type="button"
                        onClick={() => extractDocuments(files)}
                        disabled={isExtracting}
                        className="inline-flex h-10 items-center justify-center rounded-md border border-zinc-300 bg-white px-4 text-sm font-medium text-zinc-700 transition hover:bg-zinc-100 disabled:opacity-50"
                      >
                        Extract Again
                      </button>

                      <button
                        type="button"
                        onClick={handleSaveUploadedExcel}
                        disabled={isExtracting || extractedRecords.length === 0}
                        className="inline-flex h-10 items-center justify-center gap-2 rounded-md bg-zinc-900 px-5 text-sm font-medium text-white transition hover:bg-zinc-800 disabled:opacity-50"
                      >
                        <ExcelIcon />
                        Save to Excel
                      </button>
                    </div>
                  </div>
                </div>
              )}
            </div>
          </section>
        ) : (
          /* MANUAL */

          <form
            noValidate
            onSubmit={handleSubmit(onManualSubmit, onInvalid)}
            className="overflow-hidden rounded-xl border border-zinc-200 bg-white shadow-sm"
          >
            <FormSection
              number="01"
              title="Form Information"
              description="Select the BIR tax form."
            >
              <div className="grid gap-5 md:grid-cols-2">
                <Field
                  id="formType"
                  label="BIR Form"
                  required
                  error={errors.formType?.message}
                >
                  <select
                    id="formType"
                    className={getInputClass(errors.formType)}
                    {...register("formType", {
                      required: "BIR form is required.",
                    })}
                  >
                    {FORM_TYPES.map((form) => (
                      <option key={form} value={form}>
                        {form}
                      </option>
                    ))}
                  </select>
                </Field>
              </div>
            </FormSection>

            <FormSection
              number="02"
              title="Taxpayer Information"
              description="Enter the registered taxpayer information."
            >
              <div className="grid gap-5 md:grid-cols-2">
                <Field
                  id="taxpayerName"
                  label="Taxpayer Name"
                  required
                  error={errors.taxpayerName?.message}
                >
                  <input
                    id="taxpayerName"
                    type="text"
                    placeholder="Juan Dela Cruz"
                    className={getInputClass(errors.taxpayerName)}
                    {...register("taxpayerName", {
                      required: "Taxpayer name is required.",

                      validate: validateTaxpayerName,
                    })}
                  />
                </Field>

                <Field
                  id="tin"
                  label="TIN"
                  required
                  hint="Supports 9, 12, or 14 digits."
                  error={errors.tin?.message}
                >
                  <Controller
                    name="tin"
                    control={control}
                    rules={{
                      required: "TIN is required.",

                      validate: validateTin,
                    }}
                    render={({ field }) => (
                      <input
                        id="tin"
                        type="text"
                        inputMode="numeric"
                        maxLength={17}
                        placeholder="123-456-789-00000"
                        name={field.name}
                        ref={field.ref}
                        value={field.value}
                        onBlur={field.onBlur}
                        onChange={(event) =>
                          field.onChange(formatTin(event.target.value))
                        }
                        className={getInputClass(errors.tin)}
                      />
                    )}
                  />
                </Field>
              </div>
            </FormSection>

            <FormSection
              number="03"
              title="Tax Period"
              description="Enter the applicable tax period."
            >
              <div className="grid gap-5 md:grid-cols-2">
                <Field
                  id="quarter"
                  label="Quarter"
                  required
                  error={errors.quarter?.message}
                >
                  <select
                    id="quarter"
                    className={getInputClass(errors.quarter)}
                    {...register("quarter", {
                      required: "Quarter is required.",
                    })}
                  >
                    <option value="">Select quarter</option>

                    <option value="1st Quarter">1st Quarter</option>

                    <option value="2nd Quarter">2nd Quarter</option>

                    <option value="3rd Quarter">3rd Quarter</option>

                    <option value="4th Quarter">4th Quarter</option>
                  </select>
                </Field>

                <Field
                  id="year"
                  label="Tax Year"
                  required
                  error={errors.year?.message}
                >
                  <Controller
                    name="year"
                    control={control}
                    rules={{
                      required: "Tax year is required.",

                      validate: (value) => {
                        if (!(value instanceof Date)) {
                          return "Select a valid year.";
                        }

                        const year = value.getFullYear();

                        if (year < 2000 || year > CURRENT_YEAR) {
                          return `Year must be between 2000 and ${CURRENT_YEAR}.`;
                        }

                        return true;
                      },
                    }}
                    render={({ field }) => (
                      <DatePicker
                        id="year"
                        selected={field.value}
                        onChange={field.onChange}
                        onBlur={field.onBlur}
                        showYearPicker
                        dateFormat="yyyy"
                        minDate={new Date(2000, 0, 1)}
                        maxDate={new Date(CURRENT_YEAR, 11, 31)}
                        wrapperClassName="w-full"
                        className={getInputClass(errors.year)}
                        readOnly
                      />
                    )}
                  />
                </Field>
              </div>
            </FormSection>

            <FormSection
              number="04"
              title="Tax Amounts"
              description="Enter the financial amounts."
              last
            >
              <div className="grid gap-5 md:grid-cols-2">
                <CurrencyField
                  name="grossSales"
                  label="Gross Sales"
                  control={control}
                  error={errors.grossSales?.message}
                />

                <CurrencyField
                  name="taxableIncome"
                  label="Taxable Income"
                  control={control}
                  error={errors.taxableIncome?.message}
                />

                <CurrencyField
                  name="taxDue"
                  label="Tax Due"
                  control={control}
                  error={errors.taxDue?.message}
                />
              </div>
            </FormSection>

            <footer className="flex flex-col-reverse gap-3 border-t border-zinc-200 bg-zinc-50 px-5 py-4 sm:flex-row sm:items-center sm:justify-between md:px-7">
              <p className="text-xs text-zinc-500">
                Fields marked with{" "}
                <span className="font-semibold text-red-600">*</span> are
                required.
              </p>

              <div className="flex flex-col-reverse gap-3 sm:flex-row">
                <button
                  type="button"
                  onClick={handleClearManual}
                  disabled={isSubmitting}
                  className="inline-flex h-10 items-center justify-center rounded-md border border-zinc-300 bg-white px-4 text-sm font-medium text-zinc-700 transition hover:bg-zinc-100 disabled:opacity-50"
                >
                  Clear Form
                </button>

                <button
                  type="submit"
                  disabled={isSubmitting}
                  className="inline-flex h-10 items-center justify-center gap-2 rounded-md bg-zinc-900 px-5 text-sm font-medium text-white transition hover:bg-zinc-800 disabled:opacity-60"
                >
                  <ExcelIcon />
                  Save to Excel
                </button>
              </div>
            </footer>
          </form>
        )}
      </div>
    </main>
  );
}

/* =========================================================
   MODE BUTTON
   ========================================================= */

function ModeButton({ active, title, description, onClick }) {
  return (
    <button
      type="button"
      onClick={onClick}
      aria-pressed={active}
      className={`rounded-lg border p-4 text-left outline-none transition focus:ring-2 focus:ring-zinc-900 focus:ring-offset-2 ${
        active
          ? "border-zinc-900 bg-white ring-1 ring-zinc-900"
          : "border-zinc-200 bg-white hover:border-zinc-300 hover:bg-zinc-50"
      }`}
    >
      <div className="flex items-start justify-between gap-4">
        <div>
          <p className="text-sm font-semibold text-zinc-900">{title}</p>

          <p className="mt-1 text-sm leading-5 text-zinc-500">{description}</p>
        </div>

        <span
          className={`mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded-full border ${
            active ? "border-zinc-900" : "border-zinc-300"
          }`}
        >
          {active && <span className="h-2.5 w-2.5 rounded-full bg-zinc-900" />}
        </span>
      </div>
    </button>
  );
}

/* =========================================================
   FORM SECTION
   ========================================================= */

function FormSection({ number, title, description, children, last = false }) {
  return (
    <section
      className={`p-5 sm:p-6 md:p-7 ${last ? "" : "border-b border-zinc-200"}`}
    >
      <div className="mb-6 flex items-start gap-4">
        <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-md bg-zinc-100 text-xs font-semibold text-zinc-600">
          {number}
        </span>

        <div>
          <h2 className="text-base font-semibold text-zinc-950">{title}</h2>

          <p className="mt-1 text-sm leading-5 text-zinc-500">{description}</p>
        </div>
      </div>

      {children}
    </section>
  );
}

/* =========================================================
   FIELD
   ========================================================= */

function Field({ id, label, required, hint, error, children }) {
  return (
    <div className="min-w-0">
      <label
        htmlFor={id}
        className="mb-2 block text-sm font-medium text-zinc-700"
      >
        {label}

        {required && <span className="ml-1 text-red-600">*</span>}
      </label>

      {children}

      {error ? (
        <p role="alert" className="mt-1.5 text-xs font-medium text-red-600">
          {error}
        </p>
      ) : hint ? (
        <p className="mt-1.5 text-xs leading-5 text-zinc-500">{hint}</p>
      ) : null}
    </div>
  );
}

/* =========================================================
   CURRENCY FIELD
   ========================================================= */

function CurrencyField({ name, label, control, error }) {
  return (
    <Field id={name} label={label} required error={error}>
      <Controller
        name={name}
        control={control}
        rules={{
          validate: validateAmount,
        }}
        render={({ field }) => (
          <NumericFormat
            id={name}
            name={field.name}
            value={field.value}
            getInputRef={field.ref}
            onBlur={field.onBlur}
            onValueChange={(values) => {
              field.onChange(values.value);
            }}
            prefix="₱ "
            thousandSeparator=","
            decimalSeparator="."
            decimalScale={2}
            fixedDecimalScale
            allowNegative={false}
            allowLeadingZeros={false}
            inputMode="decimal"
            placeholder="₱ 0.00"
            className={getInputClass(error)}
          />
        )}
      />
    </Field>
  );
}

/* =========================================================
   FILE ITEM
   ========================================================= */

function FileItem({ file, state, disabled, onRemove }) {
  return (
    <li className="flex items-center gap-3 bg-white px-4 py-3.5">
      <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-md border border-zinc-200 bg-zinc-50 text-zinc-600">
        <DocumentIcon />
      </div>

      <div className="min-w-0 flex-1">
        <div className="flex min-w-0 flex-wrap items-center gap-2">
          <p
            className="min-w-0 truncate text-sm font-medium text-zinc-900"
            title={file.name}
          >
            {file.name}
          </p>

          <span className="shrink-0 rounded border border-zinc-200 bg-zinc-50 px-1.5 py-0.5 text-[10px] font-semibold uppercase text-zinc-500">
            {getFileTypeLabel(file)}
          </span>

          <FileStatus state={state} />
        </div>

        <p className="mt-1 text-xs text-zinc-500">
          {formatFileSize(file.size)}

          {state?.message ? ` • ${state.message}` : ""}
        </p>
      </div>

      <button
        type="button"
        onClick={onRemove}
        disabled={disabled}
        aria-label={`Remove ${file.name}`}
        className="flex h-9 w-9 shrink-0 items-center justify-center rounded-md text-zinc-400 transition hover:bg-red-50 hover:text-red-600 disabled:opacity-30"
      >
        <TrashIcon />
      </button>
    </li>
  );
}

/* =========================================================
   FILE STATUS
   ========================================================= */

function FileStatus({ state }) {
  if (!state) {
    return null;
  }

  const styles = {
    reading: "bg-zinc-100 text-zinc-600",

    extracted: "bg-emerald-50 text-emerald-700",

    partial: "bg-amber-50 text-amber-700",

    review: "bg-amber-50 text-amber-700",

    failed: "bg-red-50 text-red-700",
  };

  const labels = {
    reading: "Reading",

    extracted: "Extracted",

    partial: "Partial",

    review: "Review",

    failed: "Failed",
  };

  return (
    <span
      className={`rounded-full px-2 py-0.5 text-[10px] font-semibold ${
        styles[state.status] || styles.reading
      }`}
    >
      {labels[state.status] || state.status}
    </span>
  );
}

/* =========================================================
   ICONS
   ========================================================= */

function UploadIcon() {
  return (
    <svg
      width="20"
      height="20"
      viewBox="0 0 24 24"
      fill="none"
      aria-hidden="true"
    >
      <path
        d="M12 16V4M12 4L7.5 8.5M12 4L16.5 8.5M5 15.5V18C5 19.1 5.9 20 7 20H17C18.1 20 19 19.1 19 18V15.5"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function DocumentIcon() {
  return (
    <svg
      width="19"
      height="19"
      viewBox="0 0 24 24"
      fill="none"
      aria-hidden="true"
    >
      <path
        d="M14 3H7C5.9 3 5 3.9 5 5V19C5 20.1 5.9 21 7 21H17C18.1 21 19 20.1 19 19V8L14 3Z"
        stroke="currentColor"
        strokeWidth="1.7"
        strokeLinejoin="round"
      />

      <path
        d="M14 3V8H19"
        stroke="currentColor"
        strokeWidth="1.7"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function TrashIcon() {
  return (
    <svg
      width="17"
      height="17"
      viewBox="0 0 24 24"
      fill="none"
      aria-hidden="true"
    >
      <path
        d="M4 7H20M10 11V17M14 11V17M6 7L7 20H17L18 7M9 7V4H15V7"
        stroke="currentColor"
        strokeWidth="1.7"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function ExcelIcon() {
  return (
    <svg
      width="17"
      height="17"
      viewBox="0 0 24 24"
      fill="none"
      aria-hidden="true"
    >
      <path
        d="M14 3H7C5.9 3 5 3.9 5 5V19C5 20.1 5.9 21 7 21H17C18.1 21 19 20.1 19 19V8L14 3Z"
        stroke="currentColor"
        strokeWidth="1.7"
        strokeLinejoin="round"
      />

      <path
        d="M14 3V8H19"
        stroke="currentColor"
        strokeWidth="1.7"
        strokeLinejoin="round"
      />

      <path
        d="M9 12L15 17M15 12L9 17"
        stroke="currentColor"
        strokeWidth="1.7"
        strokeLinecap="round"
      />
    </svg>
  );
}
