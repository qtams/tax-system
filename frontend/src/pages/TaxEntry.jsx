import { useRef, useState } from "react";
import Swal from "sweetalert2";
import * as XLSX from "xlsx";

import * as pdfjsLib from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import mammoth from "mammoth";
import { createWorker } from "tesseract.js";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

/* =========================================================
   CONSTANTS
   ========================================================= */

const CURRENT_YEAR = new Date().getFullYear();

const MAX_FILES = 10;
const MAX_FILE_SIZE = 20 * 1024 * 1024;
const MAX_TOTAL_SIZE = 100 * 1024 * 1024;

/*
  TEST MODE

  Process only page 1 from each uploaded PDF while testing.
  A 44-page PDF will scan ONLY page 1.
*/
const DOCUMENT_PAGE_LIMIT = 1;

/*
  Render PDF pages at 4x for OCR accuracy.
*/
const PDF_OCR_SCALE = 4;

/*
  Identity fields on BIR 2307 are physically small. After cropping them
  from the rendered page, enlarge the crop again before sending it to
  Tesseract. This is especially important for names and the first TIN group.
*/
const PDF_IDENTITY_UPSCALE = 4;
const PDF_TIN_GROUP_UPSCALE = 5;

/*
  A smaller render is enough to identify the BIR form number
  before deciding which OCR strategy to use.
*/
const PDF_CLASSIFIER_SCALE = 2;

const PDF_TEXT_MIN_LENGTH = 80;

const PDF_TEXT_MIN_WORDS = 8;

const PDF_TEXT_MIN_ALNUM_RATIO = 0.5;

const PDF_HEADER_OCR_REGION = {
  x: 0,
  y: 0,
  width: 1,
  height: 0.24,
};

const ALLOWED_EXTENSIONS = [".pdf", ".docx"];

const ALLOWED_MIME_TYPES = [
  "application/pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
];

/*
  These percentages are based on the actual BIR 2307
  layout in the uploaded file.

  They are relative to the rendered PDF page, therefore
  they continue to work when the PDF is rendered at a
  different resolution.

  PERIOD:
  only the "For the Period" row.

  INFORMATION:
  Payee + Payor information.

  TABLE:
  income table + total row.
*/
const FORM_2307_TIN_GROUPS = [
  { x: 0.0, y: 0.12, width: 0.19, height: 0.76, digits: 3 },
  { x: 0.22, y: 0.12, width: 0.2, height: 0.76, digits: 3 },
  { x: 0.45, y: 0.12, width: 0.2, height: 0.76, digits: 3 },
  { x: 0.68, y: 0.12, width: 0.3, height: 0.76, digits: 5 },
];

const FORM_2307_OCR_REGIONS = {
  period: {
    x: 0.03,
    y: 0.135,
    width: 0.94,
    height: 0.035,
  },

  /*
    Tight date-value crops for the standard 2307 template.

    The wider period crop above remains as a fallback for photographed /
    shifted pages. Reading From and To independently prevents row numbers,
    labels and border artifacts from being concatenated into a fake date.
  */
  fromDate: {
    x: 0.235,
    y: 0.141,
    width: 0.17,
    height: 0.022,
  },

  toDate: {
    x: 0.62,
    y: 0.141,
    width: 0.175,
    height: 0.022,
  },

  /*
    Calibrated against the ACTUAL uploaded 2307-Q2-2026 PDF.

    Important: these four identity crops target only the VALUE rows.
    The previous coordinates were landing on the labels / section headers,
    which is why Payee Name, Payor Name and Payor TIN were blank or the
    wrong TIN was being assigned to the Payee.

    Ratios remain resolution-independent because they are relative to the
    rendered PDF page.
  */
  information: {
    x: 0.025,
    y: 0.155,
    width: 0.95,
    height: 0.22,
  },

  // Page 1 actual row: 297 - 944 - 186 - 00000
  payeeTin: {
    x: 0.3255,
    y: 0.1774,
    width: 0.3627,
    height: 0.0167,
  },

  // Page 1 actual value: CASTILLO LEX LYCURGUS NORONA
  payeeName: {
    x: 0.05,
    y: 0.195,
    width: 0.9,
    height: 0.023,
  },

  // Page 1 actual row: 236 - 809 - 045 - 00000
  payorTin: {
    x: 0.3255,
    y: 0.2869,
    width: 0.3627,
    height: 0.0167,
  },

  // Page 1 actual value: THE MEDICAL CITY SOUTH LUZON
  payorName: {
    x: 0.05,
    y: 0.304,
    width: 0.9,
    height: 0.023,
  },

  table: {
    x: 0.025,
    y: 0.365,
    width: 0.95,
    height: 0.205,
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
   GENERAL HELPERS
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
   TEXT HELPERS
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

function parseAmount(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const normalized = String(value)
    .replace(/PHP/gi, "")
    .replace(/₱/g, "")
    .replace(/,/g, "")
    .replace(/[^\d.-]/g, "");

  const number = Number(normalized);

  return Number.isFinite(number) ? number : null;
}

function cleanExtractedName(value = "") {
  return String(value)
    .replace(/\s+/g, " ")
    .replace(/^[\s:;,.–—|_-]+/, "")
    .replace(/[\s:;,.–—|_-]+$/, "")
    .trim();
}

function roundAmount(value) {
  return Math.round(Number(value) * 100) / 100;
}

/* =========================================================
   MONEY EXTRACTION
   ========================================================= */

function findAmount(text, labels, maxDistance = 120) {
  for (const label of labels) {
    const regex = new RegExp(
      `${label}[\\s\\S]{0,${maxDistance}}?(?:PHP|₱)?\\s*([0-9][0-9,]*(?:\\.\\d{1,2})?)`,
      "i",
    );

    const match = text.match(regex);

    if (match?.[1]) {
      const amount = parseAmount(match[1]);

      if (amount !== null) {
        return amount;
      }
    }
  }

  return null;
}

function extractMoneyValues(text) {
  const values = [];

  /*
    Require either:

    5,700.00
    570.00
    15,273.00

    This prevents ATC "151" from becoming
    a financial value.
  */
  const regex = /\b(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}\b/g;

  let match;

  while ((match = regex.exec(text)) !== null) {
    const amount = parseAmount(match[0]);

    if (amount !== null && Number.isFinite(amount) && amount >= 0) {
      values.push(roundAmount(amount));
    }
  }

  return values;
}

function getMoneyFrequency(values) {
  const map = new Map();

  for (const value of values) {
    if (!Number.isFinite(value)) {
      continue;
    }

    const rounded = roundAmount(value);

    const key = rounded.toFixed(2);

    const existing = map.get(key);

    if (existing) {
      existing.count += 1;
    } else {
      map.set(key, {
        value: rounded,
        count: 1,
      });
    }
  }

  return [...map.values()];
}

/* =========================================================
   FORM TYPE
   ========================================================= */

function looksLike2307(text) {
  const normalized = normalizeText(text);

  const hasPayee = /Payee(?:'s|’s)?\s+Name|Payee\s+Information/i.test(
    normalized,
  );

  const hasPayor = /Payor(?:'s|’s)?\s+Name|Payor\s+Information/i.test(
    normalized,
  );

  const hasWithholding =
    /Tax\s+Withheld|Expanded\s+Withholding\s+Tax|Income\s+Payments/i.test(
      normalized,
    );

  return hasPayee && hasPayor && hasWithholding;
}

function findFormType(text) {
  const normalized = normalizeText(text);

  const match = normalized.match(
    /\b(2307|1701Q|1701A|1701|1601[\s-]?EQ|2551Q|2550Q)\b/i,
  );

  if (match) {
    let form = match[1].toUpperCase().replace(/\s/g, "");

    if (form === "1601EQ") {
      form = "1601-EQ";
    }

    return `BIR Form ${form}`;
  }

  /*
    The region OCR intentionally does not OCR
    the entire 2307 header.

    Detect 2307 using its unique Payee/Payor layout.
  */
  if (looksLike2307(normalized)) {
    return "BIR Form 2307";
  }

  return "";
}

/* =========================================================
   OCR DATE HELPERS
   ========================================================= */

function normalizeOcrDateToken(value = "") {
  return String(value)
    .toUpperCase()
    .replace(/O/g, "0")
    .replace(/Q/g, "0")
    .replace(/A/g, "4")
    .replace(/[IL|]/g, "1")
    .replace(/Z/g, "2")
    .replace(/S/g, "5")
    .replace(/B/g, "8")
    .replace(/G/g, "6")
    .replace(/\D/g, "");
}

function extractDatesFromOcrLines(rawText) {
  const lines = cleanText(rawText)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  const dates = [];

  for (const line of lines) {
    const yearMatch = line.match(/\b((?:19|20)\d{2})\b/);

    if (!yearMatch) {
      continue;
    }

    const year = Number(yearMatch[1]);

    const beforeYear = line.slice(0, yearMatch.index);

    const rawTokens = beforeYear.match(/[0-9OAQILSBZG|]{1,3}/gi) || [];

    const numericTokens = rawTokens
      .map(normalizeOcrDateToken)
      .filter((value) => value && value.length <= 2 && Number(value) >= 0);

    if (numericTokens.length < 2) {
      continue;
    }

    const month = Number(numericTokens[numericTokens.length - 2]);

    const day = Number(numericTokens[numericTokens.length - 1]);

    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      dates.push({
        month,
        day,
        year,
      });
    }
  }

  return dates;
}

/* =========================================================
   TIN EXTRACTION
   ========================================================= */

function normalizeOcrDigitToken(value = "") {
  return String(value)
    .toUpperCase()
    .replace(/[OQD]/g, "0")
    .replace(/[IL|]/g, "1")
    .replace(/Z/g, "2")
    .replace(/S/g, "5")
    .replace(/G/g, "6")
    .replace(/B/g, "8")
    .replace(/\D/g, "");
}

function buildTinFromGroups(groups) {
  if (!Array.isArray(groups)) {
    return "";
  }

  if (groups.length < 3) {
    return "";
  }

  let first = normalizeOcrDigitToken(groups[0]);

  let second = normalizeOcrDigitToken(groups[1]);

  let third = normalizeOcrDigitToken(groups[2]);

  let branch = normalizeOcrDigitToken(groups[3] || "");

  if (first.length !== 3 || second.length !== 3 || third.length !== 3) {
    return "";
  }

  /*
    In the uploaded BIR 2307 documents, the
    branch field is 5 digits.

    OCR occasionally reads:
    00000 -> 000

    For a branch consisting entirely of zeroes,
    safely restore the omitted zeroes.
  */
  if (branch.length >= 1 && branch.length < 5 && /^0+$/.test(branch)) {
    branch = branch.padEnd(5, "0");
  }

  if (branch && branch.length > 5) {
    branch = branch.slice(0, 5);
  }

  const digits = `${first}${second}${third}${branch}`;

  if (![9, 12, 14].includes(digits.length)) {
    return "";
  }

  return formatTin(digits);
}

function findTinGroupsInText(text) {
  const normalized = String(text);

  /*
    First try ordinary numeric groups.

    Actual OCR from the sample produces:
    297 944 M186 Mi 00000

    Numeric groups are still:
    297
    944
    186
    00000
  */
  const numericGroups = normalized.match(/\d{3,5}/g) || [];

  for (let index = 0; index <= numericGroups.length - 4; index += 1) {
    const group = numericGroups.slice(index, index + 4);

    if (
      group[0].length === 3 &&
      group[1].length === 3 &&
      group[2].length === 3 &&
      group[3].length >= 3 &&
      group[3].length <= 5
    ) {
      const tin = buildTinFromGroups(group);

      if (tin) {
        return tin;
      }
    }
  }

  /*
    Fallback for OCR using letters that look
    like digits.
  */
  const permissiveRegex =
    /([0-9OQDIiLl|SBZG]{3})[^0-9OQDIiLl|SBZG]{0,6}([0-9OQDIiLl|SBZG]{3})[^0-9OQDIiLl|SBZG]{0,6}([0-9OQDIiLl|SBZG]{3})[^0-9OQDIiLl|SBZG]{0,8}([0-9OQDIiLl|SBZG]{3,5})/g;

  let match;

  while ((match = permissiveRegex.exec(normalized)) !== null) {
    const tin = buildTinFromGroups([match[1], match[2], match[3], match[4]]);

    if (tin) {
      return tin;
    }
  }

  return "";
}

function findTinsNearLabels(rawText) {
  const lines = cleanText(rawText)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  const results = [];

  for (let index = 0; index < lines.length; index += 1) {
    if (
      !/Taxpayer\s+Identification\s+Number|Identification\s+Number\s*\(TIN\)|Number\s*\(TIN\)/i.test(
        lines[index],
      )
    ) {
      continue;
    }

    const window = lines.slice(index, index + 4).join(" ");

    const tin = findTinGroupsInText(window);

    if (tin && !results.includes(tin)) {
      results.push(tin);
    }
  }

  return results;
}

function findAllTins(text) {
  const targeted = findTinsNearLabels(text);

  if (targeted.length >= 2) {
    return targeted;
  }

  const values = [...targeted];

  const normalized = normalizeText(text);

  /*
    Normal 2307 TIN format.
  */
  const genericRegex =
    /\b(\d{3})\s*[-–—| ]+\s*(\d{3})\s*[-–—| ]+\s*(\d{3})\s*[-–—| ]+\s*(\d{3,5})\b/g;

  let match;

  while ((match = genericRegex.exec(normalized)) !== null) {
    const tin = buildTinFromGroups([match[1], match[2], match[3], match[4]]);

    if (tin && !values.includes(tin)) {
      values.push(tin);
    }
  }

  return values;
}

function findTin(text) {
  return findAllTins(text)[0] || "";
}

/* =========================================================
   GENERIC TAXPAYER
   ========================================================= */

function findTaxpayerName(text) {
  const patterns = [
    /Taxpayer(?:'s)?\s+Name[\s\S]{0,40}?[:\-]?\s*([A-Z][A-Z ,.'&’/-]{2,100}?)(?=\s+(?:Registered|Address|TIN|RDO|Tax|Email|Contact|Quarter|Year|Citizenship|Civil)\b|$)/i,

    /Name\s*\(Last\s*Name,\s*First\s*Name,\s*Middle\s*Name\)[\s\S]{0,40}?([A-Z][A-Z ,.'&’/-]{2,100}?)(?=\s+(?:Registered|Address|TIN|RDO|Tax|Email|Contact)\b|$)/i,

    /Taxpayer\/Filer[\s\S]{0,40}?Name[\s\S]{0,30}?[:\-]?\s*([A-Z][A-Z ,.'&’/-]{2,100}?)(?=\s+(?:Registered|Address|TIN|RDO|Tax|Email|Contact)\b|$)/i,
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);

    if (match?.[1]) {
      return cleanExtractedName(match[1]);
    }
  }

  return "";
}

/* =========================================================
   PERIOD / YEAR
   ========================================================= */

function getQuarterFromMonth(month) {
  const monthNumber = Number(month);

  if (!Number.isFinite(monthNumber) || monthNumber < 1 || monthNumber > 12) {
    return "";
  }

  if (monthNumber <= 3) {
    return "1st Quarter";
  }

  if (monthNumber <= 6) {
    return "2nd Quarter";
  }

  if (monthNumber <= 9) {
    return "3rd Quarter";
  }

  return "4th Quarter";
}

function findQuarter(text) {
  const direct = text.match(/\b(1st|2nd|3rd|4th)\s+Quarter\b/i);

  if (direct?.[1]) {
    const number = direct[1][0];

    const suffix = {
      1: "st",
      2: "nd",
      3: "rd",
      4: "th",
    };

    return `${number}${suffix[number]} Quarter`;
  }

  const numbered = text.match(/\b(?:Quarter|Qtr)[\s:.-]*([1-4])\b/i);

  if (numbered?.[1]) {
    const number = numbered[1];

    const suffix = {
      1: "st",
      2: "nd",
      3: "rd",
      4: "th",
    };

    return `${number}${suffix[number]} Quarter`;
  }

  return "";
}

function findYear(text) {
  const patterns = [
    /Tax\s*Year[\s:.-]*((?:19|20)\d{2})/i,

    /For\s+the\s+Year[\s\S]{0,30}?((?:19|20)\d{2})/i,

    /Calendar\s+Year[\s:.-]*((?:19|20)\d{2})/i,

    /Year[\s:.-]+((?:19|20)\d{2})/i,
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);

    if (match?.[1]) {
      const year = Number(match[1]);

      if (year >= 2000 && year <= CURRENT_YEAR) {
        return year;
      }
    }
  }

  return null;
}

function formatDateParts(month, day, year) {
  const mm = String(month).padStart(2, "0");

  const dd = String(day).padStart(2, "0");

  return `${mm}/${dd}/${year}`;
}

/* =========================================================
   FORM 2307 - PERIOD
   ========================================================= */

function find2307Period(text) {
  /*
    First use the actual OCR lines.

    This is more reliable for boxed dates such as:

    From | 04 | 01 | 2026
    To   | 04 | 30 | 2026
  */
  const lineDates = extractDatesFromOcrLines(text);

  if (lineDates.length >= 2) {
    const first = lineDates[0];

    const second = lineDates[1];

    return {
      from: formatDateParts(first.month, first.day, first.year),

      to: formatDateParts(second.month, second.day, second.year),

      quarter: getQuarterFromMonth(second.month),

      year: second.year,
    };
  }

  /*
    Fallback for OCR that puts the dates
    on one continuous line.
  */
  const normalized = normalizeText(text);

  const periodIndex = normalized.search(/For\s+the\s+Period/i);

  const searchText =
    periodIndex >= 0
      ? normalized.slice(periodIndex, periodIndex + 350)
      : normalized;

  const dateRegex =
    /\b(0?[1-9]|1[0-2])\D{0,6}(0?[1-9]|[12]\d|3[01])\D{0,6}((?:19|20)\d{2})\b/g;

  const dates = [];

  let match;

  while ((match = dateRegex.exec(searchText)) !== null) {
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

  const first = dates[0];

  const second = dates[1] || dates[0];

  return {
    from: formatDateParts(first.month, first.day, first.year),

    to: formatDateParts(second.month, second.day, second.year),

    quarter: getQuarterFromMonth(second.month),

    year: second.year,
  };
}

/* =========================================================
   FORM 2307 - NAME EXTRACTION
   ========================================================= */

function isLikelyPersonOrCompanyName(value) {
  const candidate = cleanExtractedName(value);

  if (!candidate || candidate.length < 3 || candidate.length > 120) {
    return false;
  }

  if (
    /Registered\s+Address|Foreign\s+Address|ZIP\s+Code|Taxpayer\s+Identification|TIN\b|Part\s+[IVX]+|For\s+Individual|Non-Individual/i.test(
      candidate,
    )
  ) {
    return false;
  }

  const letters = candidate.replace(/[^A-Za-z]/g, "");

  return letters.length >= 3;
}

function extract2307NameFromFieldCrop(rawText) {
  const lines = cleanText(rawText)
    .split("\n")
    .map((line) =>
      cleanExtractedName(
        line
          .replace(/^[^A-Za-z]+/, "")
          .replace(/[^A-Za-z0-9 .,'&’()\/-]+$/g, ""),
      ),
    )
    .filter(Boolean)
    .filter((line) => isLikelyPersonOrCompanyName(line))
    .filter(
      (line) =>
        !/Payee|Payor|Taxpayer|Identification|Registered|Address|TIN|ZIP|Part\s+[IVX]+/i.test(
          line,
        ),
    );

  if (!lines.length) {
    return "";
  }

  /*
    A tight field crop should normally produce one line. If OCR creates
    multiple fragments, prefer the candidate containing the most letters.
  */
  return [...lines].sort((a, b) => {
    const aLetters = a.replace(/[^A-Za-z]/g, "").length;
    const bLetters = b.replace(/[^A-Za-z]/g, "").length;

    return bLetters - aLetters;
  })[0];
}

function extract2307TinFromFieldCrop(rawText) {
  const direct = findTinGroupsInText(rawText);

  if (direct) {
    return direct;
  }

  /*
    OCR sometimes turns a zero into O/Q or a one into I/L. Build the TIN
    from permissive groups as a second pass.
  */
  const groups = String(rawText).match(/[0-9OQDIiLl|SBZG]{3,5}/g) || [];

  for (let index = 0; index <= groups.length - 4; index += 1) {
    const tin = buildTinFromGroups(groups.slice(index, index + 4));

    if (tin) {
      return tin;
    }
  }

  return "";
}

function merge2307IdentityOverrides(data, overrides) {
  if (!overrides) {
    return data;
  }

  const payeeName = isLikelyPersonOrCompanyName(overrides.payeeName)
    ? cleanExtractedName(overrides.payeeName)
    : "";

  const payorName = isLikelyPersonOrCompanyName(overrides.payorName)
    ? cleanExtractedName(overrides.payorName)
    : "";

  const payeeTin = isValidExtractedTin(overrides.payeeTin)
    ? overrides.payeeTin
    : "";

  const payorTin = isValidExtractedTin(overrides.payorTin)
    ? overrides.payorTin
    : "";

  return {
    ...data,
    form_type: data?.form_type || "BIR Form 2307",
    taxpayer: {
      ...data?.taxpayer,
      name: payeeName || data?.taxpayer?.name || "",
      tin: payeeTin || data?.taxpayer?.tin || "",
    },
    withholding_agent: {
      ...data?.withholding_agent,
      name: payorName || data?.withholding_agent?.name || "",
      tin: payorTin || data?.withholding_agent?.tin || "",
    },
  };
}

function find2307NameByLines(rawText, type) {
  const lines = cleanText(rawText)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  const labelRegex =
    type === "payee"
      ? /Payee(?:'s|’s)?\s*.*Name/i
      : /Payor(?:'s|’s)?\s*.*Name/i;

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];

    if (!labelRegex.test(line)) {
      continue;
    }

    let sameLineCandidate = "";

    /*
      The printed BIR label contains a long
      parenthetical description.

      Anything after the final ")" may be the
      actual taxpayer name.
    */
    const closingParenthesis = line.lastIndexOf(")");

    if (closingParenthesis >= 0) {
      sameLineCandidate = line.slice(closingParenthesis + 1);
    } else {
      sameLineCandidate = line.replace(labelRegex, "");
    }

    sameLineCandidate = cleanExtractedName(sameLineCandidate);

    if (isLikelyPersonOrCompanyName(sameLineCandidate)) {
      return sameLineCandidate;
    }

    /*
      Tesseract usually returns the actual
      name on the next line.
    */
    for (let offset = 1; offset <= 3; offset += 1) {
      const nextLine = lines[index + offset];

      if (!nextLine) {
        continue;
      }

      if (
        /Registered\s+Address|Foreign\s+Address|Part\s+[IVX]+/i.test(nextLine)
      ) {
        break;
      }

      const candidate = cleanExtractedName(
        nextLine.replace(/^[|_[\]{}:;,.–—-]+/, ""),
      );

      if (isLikelyPersonOrCompanyName(candidate)) {
        return candidate;
      }
    }
  }

  return "";
}

/* =========================================================
   FORM 2307 - PAYEE
   ========================================================= */

function find2307PayeeName(text) {
  const lineValue = find2307NameByLines(text, "payee");

  if (lineValue) {
    return lineValue;
  }

  const normalized = normalizeText(text);

  const patterns = [
    /Payee(?:'s|’s)?\s+Name(?:\s*\([^)]*\))?[\s:.-]*([A-Z][A-Z0-9 ,.'&’/-]{2,120}?)(?=\s+(?:4\s+)?Registered\s+Address|\s+Registered\s+Address|\s+Taxpayer\s+Identification|\s+TIN\b|\s+Part\s+II)/i,

    /Payee\s+Name[\s:.-]*([A-Z][A-Z0-9 ,.'&’/-]{2,120}?)(?=\s+Registered|\s+Address|\s+TIN\b)/i,
  ];

  for (const pattern of patterns) {
    const match = normalized.match(pattern);

    if (match?.[1]) {
      return cleanExtractedName(match[1]);
    }
  }

  return "";
}

/* =========================================================
   FORM 2307 - PAYOR
   ========================================================= */

function find2307PayorName(text) {
  const lineValue = find2307NameByLines(text, "payor");

  if (lineValue) {
    return lineValue;
  }

  const normalized = normalizeText(text);

  const patterns = [
    /Payor(?:'s|’s)?\s+Name(?:\s*\([^)]*\))?[\s:.-]*([A-Z][A-Z0-9 ,.'&’/-]{2,120}?)(?=\s+(?:8\s+)?Registered\s+Address|\s+Registered\s+Address|\s+Taxpayer\s+Identification|\s+TIN\b|\s+Part\s+III)/i,

    /Payor\s+Name[\s:.-]*([A-Z][A-Z0-9 ,.'&’/-]{2,120}?)(?=\s+Registered|\s+Address|\s+TIN\b)/i,
  ];

  for (const pattern of patterns) {
    const match = normalized.match(pattern);

    if (match?.[1]) {
      return cleanExtractedName(match[1]);
    }
  }

  return "";
}

/* =========================================================
   FORM 2307 - ATC
   ========================================================= */

function find2307Atc(text) {
  const normalized = normalizeText(text);

  /*
    Standard ATCs such as:
    WI150
    WI151
    WC010
  */
  const standardMatch = normalized.match(/\bW\s*([IC1L])\s*[-]?\s*(\d{3})\b/i);

  if (standardMatch) {
    let type = standardMatch[1].toUpperCase();

    if (type === "1" || type === "L") {
      type = "I";
    }

    return `W${type}${standardMatch[2]}`;
  }

  /*
    IMPORTANT:

    The uploaded first 15 pages visibly contain
    bare ATC "151", not "WI151".

    Find a 3-digit code immediately before a
    monetary value.
  */
  const beforeMoney = normalized.match(
    /\b([1-9]\d{2})\b\s*[|_\-–—]*\s*(?=\d{1,3}(?:,\d{3})*\.\d{2})/,
  );

  if (beforeMoney?.[1]) {
    return beforeMoney[1];
  }

  /*
    Final fallback inside the 2307 table region.
  */
  const bare = normalized.match(/\b(1[0-9]{2}|[2-9][0-9]{2})\b/);

  return bare?.[1] || "";
}

/* =========================================================
   FORM 2307 - NATURE OF PAYMENT
   ========================================================= */

function find2307NatureOfIncome(text) {
  const normalized = normalizeText(text);

  /*
    Exact wording used in pages 1-15.
  */
  if (
    /Gross\s+Income[\s\S]{0,50}Less\s+than\s+3\s*M/i.test(normalized) &&
    /Non[\s-]*VAT[\s\S]{0,80}regardless[\s\S]{0,30}amount/i.test(normalized)
  ) {
    return "Gross Income is Less than 3M or Non VAT registered regardless of amount";
  }

  if (/Payment\s+to\s+medical\s+practitioners/i.test(normalized)) {
    return "Payment to medical practitioners through hospital/clinic";
  }

  const knownPatterns = [
    /Gross\s+Income\s+(?:is\s+)?Less\s+Than\s+3M/i,

    /Gross\s+Income\s+(?:is\s+)?3M\s+and\s+Above/i,

    /Non[\s-]*VAT\s+Registered[\s\S]{0,60}?Amount/i,

    /Professional\s+Fees?/i,

    /Rental/i,

    /Commission/i,
  ];

  for (const pattern of knownPatterns) {
    const match = normalized.match(pattern);

    if (match?.[0]) {
      return cleanExtractedName(match[0]);
    }
  }

  return "";
}

/* =========================================================
   FORM 2307 - FINANCIAL DATA
   ========================================================= */

function find2307Amounts(text) {
  const normalized = normalizeText(text);

  const allValues = extractMoneyValues(normalized);

  const positiveValues = allValues.filter((value) => value > 0);

  if (!positiveValues.length) {
    return {
      incomePayment: null,
      taxRate: null,
      taxWithheld: null,
    };
  }

  /*
    A BIR 2307 normally repeats the correct amount:

    - monthly amount
    - row total
    - bottom total

    OCR may make one isolated mistake.

    Example from page 7:
      OCR first row: 9,454.00
      actual total:  9,451.00
      repeated total: 9,451.00

    Using frequency instead of simply Math.max()
    rejects that OCR mistake.
  */
  const frequencies = getMoneyFrequency(positiveValues);

  const incomeCandidates = [...frequencies].sort((a, b) => {
    if (b.count !== a.count) {
      return b.count - a.count;
    }

    return b.value - a.value;
  });

  let incomePayment = incomeCandidates[0]?.value ?? null;

  /*
    If the highest-frequency candidate is clearly
    a small tax value and a much larger value has
    nearly the same repetition count, prefer the
    larger financial amount.
  */
  if (incomePayment !== null && incomeCandidates.length > 1) {
    const first = incomeCandidates[0];

    const second = incomeCandidates[1];

    if (second.value > first.value * 2 && second.count >= first.count - 1) {
      incomePayment = second.value;
    }
  }

  let taxWithheld = null;

  if (incomePayment !== null) {
    const taxCandidates = frequencies
      .filter((entry) => entry.value > 0 && entry.value < incomePayment)
      .sort((a, b) => {
        if (b.count !== a.count) {
          return b.count - a.count;
        }

        /*
            If frequency is tied, prefer the
            larger candidate rather than a tiny
            OCR artifact.
          */
        return b.value - a.value;
      });

    /*
      Prefer candidates yielding a plausible
      withholding rate.
    */
    const plausible = taxCandidates.find((entry) => {
      const rate = (entry.value / incomePayment) * 100;

      return rate >= 0.1 && rate <= 35;
    });

    taxWithheld = plausible?.value ?? taxCandidates[0]?.value ?? null;
  }

  let taxRate = null;

  if (incomePayment !== null && incomePayment > 0 && taxWithheld !== null) {
    taxRate = Math.round((taxWithheld / incomePayment) * 100 * 10000) / 10000;
  }

  return {
    incomePayment: incomePayment !== null ? roundAmount(incomePayment) : null,

    taxRate,

    taxWithheld: taxWithheld !== null ? roundAmount(taxWithheld) : null,
  };
}

/* =========================================================
   FORM 2307 EXTRACTION
   ========================================================= */

function extract2307Fields(rawText) {
  const text = cleanText(rawText);

  const tins = findAllTins(text);

  const period = find2307Period(text);

  const withholding = find2307Amounts(text);

  return {
    form_type: "BIR Form 2307",

    taxpayer: {
      /*
        Form 2307 taxpayer = Payee
      */
      name: find2307PayeeName(text),

      tin: tins[0] || "",
    },

    withholding_agent: {
      /*
        Form 2307 withholding agent = Payor
      */
      name: find2307PayorName(text),

      tin: tins[1] || "",
    },

    tax_period: {
      quarter: period.quarter,

      year: period.year,

      from: period.from,

      to: period.to,
    },

    withholding: {
      nature_of_income_payment: find2307NatureOfIncome(text),

      atc: find2307Atc(text),

      income_payment: withholding.incomePayment,

      tax_rate: withholding.taxRate,

      tax_withheld: withholding.taxWithheld,
    },

    amounts: {
      /*
        Keep existing Excel compatibility.

        Gross Sales = Income Payment
        Tax Due = Tax Withheld
      */
      gross_sales: withholding.incomePayment,

      taxable_income: null,

      tax_due: withholding.taxWithheld,

      total_amount_payable: null,
    },
  };
}

/* =========================================================
   GENERIC FORM EXTRACTION
   ========================================================= */

function extractGenericTaxFields(rawText) {
  const text = normalizeText(rawText);

  return {
    form_type: findFormType(text),

    taxpayer: {
      name: findTaxpayerName(text),

      tin: findTin(text),
    },

    withholding_agent: {
      name: "",
      tin: "",
    },

    tax_period: {
      quarter: findQuarter(text),

      year: findYear(text),

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
      gross_sales: findAmount(text, [
        "Gross\\s+Sales",
        "Gross\\s+Receipts",
        "\\bSales\\b",
      ]),

      taxable_income: findAmount(text, [
        "Taxable\\s+Income",
        "Net\\s+Taxable\\s+Income",
      ]),

      tax_due: findAmount(text, ["Tax\\s+Due", "Tax\\s+Still\\s+Due"]),

      total_amount_payable: findAmount(text, [
        "Total\\s+Amount\\s+Payable",
        "Aggregate\\s+Amount\\s+Payable",
        "Total\\s+Amount\\s+Still\\s+Due",
        "Tax\\s+Payable",
        "Amount\\s+Payable",
      ]),
    },
  };
}

/* =========================================================
   MAIN TAX PARSER
   ========================================================= */

function extractTaxFields(rawText) {
  const formType = findFormType(rawText);

  if (formType === "BIR Form 2307" || looksLike2307(rawText)) {
    return extract2307Fields(rawText);
  }

  return extractGenericTaxFields(rawText);
}

/* =========================================================
   EXTRACTION QUALITY
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

  if (data.tax_period?.quarter) {
    score += 1;
  }

  if (data.tax_period?.year) {
    score += 1;
  }

  if (data.tax_period?.from) {
    score += 1;
  }

  if (data.tax_period?.to) {
    score += 1;
  }

  if (
    data.amounts?.gross_sales !== null &&
    data.amounts?.gross_sales !== undefined
  ) {
    score += 1;
  }

  if (
    data.amounts?.taxable_income !== null &&
    data.amounts?.taxable_income !== undefined
  ) {
    score += 1;
  }

  if (data.amounts?.tax_due !== null && data.amounts?.tax_due !== undefined) {
    score += 1;
  }

  if (
    data.amounts?.total_amount_payable !== null &&
    data.amounts?.total_amount_payable !== undefined
  ) {
    score += 1;
  }

  if (data.withholding_agent?.name) {
    score += 2;
  }

  if (data.withholding_agent?.tin) {
    score += 2;
  }

  if (data.withholding?.atc) {
    score += 1;
  }

  if (data.withholding?.nature_of_income_payment) {
    score += 1;
  }

  if (
    data.withholding?.income_payment !== null &&
    data.withholding?.income_payment !== undefined
  ) {
    score += 2;
  }

  if (
    data.withholding?.tax_rate !== null &&
    data.withholding?.tax_rate !== undefined
  ) {
    score += 1;
  }

  if (
    data.withholding?.tax_withheld !== null &&
    data.withholding?.tax_withheld !== undefined
  ) {
    score += 2;
  }

  return score;
}

function hasUsefulData(data) {
  return scoreExtractedData(data) >= 4;
}

function isHighConfidence2307(data) {
  if (data?.form_type !== "BIR Form 2307") {
    return false;
  }

  if (!data.taxpayer?.name || !isValidExtractedTin(data.taxpayer?.tin)) {
    return false;
  }

  if (
    !data.withholding_agent?.name ||
    !isValidExtractedTin(data.withholding_agent?.tin)
  ) {
    return false;
  }

  if (
    !data.tax_period?.from ||
    !data.tax_period?.to ||
    !data.tax_period?.year
  ) {
    return false;
  }

  if (!data.withholding?.atc) {
    return false;
  }

  if (!data.withholding?.nature_of_income_payment) {
    return false;
  }

  if (
    data.withholding?.income_payment === null ||
    data.withholding?.income_payment === undefined
  ) {
    return false;
  }

  if (
    data.withholding?.tax_withheld === null ||
    data.withholding?.tax_withheld === undefined
  ) {
    return false;
  }

  return true;
}

/* =========================================================
   PDF NATIVE TEXT / OCR DECISION
   ========================================================= */

function getPdfTextItemMetrics(item) {
  const transform = Array.isArray(item?.transform) ? item.transform : [];

  const x = Number(transform[4]) || 0;

  const y = Number(transform[5]) || 0;

  const width = Math.max(0, Number(item?.width) || 0);

  const height = Math.max(
    0,
    Math.abs(Number(item?.height) || 0),
    Math.abs(Number(transform[3]) || 0),
  );

  return {
    text: String(item?.str || ""),
    x,
    y,
    width,
    height,
    hasEOL: Boolean(item?.hasEOL),
  };
}

function arePdfItemsOnSameLine(first, second) {
  const tolerance = Math.max(
    2,
    Math.min(8, Math.max(first.height || 0, second.height || 0) * 0.55),
  );

  return Math.abs(first.y - second.y) <= tolerance;
}

function joinPdfLineItems(items) {
  if (!items.length) {
    return "";
  }

  const sorted = [...items].sort((a, b) => a.x - b.x);

  let output = "";

  let previous = null;

  for (const item of sorted) {
    const value = item.text;

    if (!value) {
      continue;
    }

    if (!previous) {
      output = value;

      previous = item;

      continue;
    }

    const previousRight = previous.x + previous.width;

    const gap = item.x - previousRight;

    const previousCharacterWidth =
      previous.text.length > 0 ? previous.width / previous.text.length : 0;

    const currentCharacterWidth =
      item.text.length > 0 ? item.width / item.text.length : 0;

    const approximateCharacterWidth = Math.max(
      1,
      previousCharacterWidth,
      currentCharacterWidth,
    );

    const needsSpace =
      !/\s$/.test(output) &&
      !/^\s/.test(value) &&
      gap > Math.max(0.75, approximateCharacterWidth * 0.28);

    output += `${needsSpace ? " " : ""}${value}`;

    previous = item;
  }

  return output.trim();
}

function extractPdfTextWithLayout(textContent) {
  const items = (textContent?.items || [])
    .map(getPdfTextItemMetrics)
    .filter((item) => item.text.trim());

  if (!items.length) {
    return "";
  }

  const sorted = [...items].sort((a, b) => {
    const verticalDifference = b.y - a.y;

    if (Math.abs(verticalDifference) > 0.01) {
      return verticalDifference;
    }

    return a.x - b.x;
  });

  const lines = [];

  let currentLine = [];

  for (const item of sorted) {
    const reference = currentLine[0];

    if (!reference || arePdfItemsOnSameLine(reference, item)) {
      currentLine.push(item);
    } else {
      lines.push(currentLine);

      currentLine = [item];
    }
  }

  if (currentLine.length) {
    lines.push(currentLine);
  }

  return cleanText(lines.map(joinPdfLineItems).filter(Boolean).join("\n"));
}

function getPdfTextQuality(text) {
  const value = String(text || "");

  const compact = value.replace(/\s/g, "");

  const words = value.match(/[\p{L}\p{N}][\p{L}\p{N}.'’/&-]*/gu) || [];

  const alphanumeric = value.match(/[\p{L}\p{N}]/gu) || [];

  const replacementCharacters = value.match(/[�□■]/g) || [];

  const alphanumericRatio =
    compact.length > 0 ? alphanumeric.length / compact.length : 0;

  return {
    characterCount: compact.length,
    wordCount: words.length,
    alphanumericRatio,
    replacementCharacterCount: replacementCharacters.length,
  };
}

function hasUsableNativePdfText(text) {
  const quality = getPdfTextQuality(text);

  if (quality.characterCount < PDF_TEXT_MIN_LENGTH) {
    return false;
  }

  if (quality.wordCount < PDF_TEXT_MIN_WORDS) {
    return false;
  }

  if (quality.alphanumericRatio < PDF_TEXT_MIN_ALNUM_RATIO) {
    return false;
  }

  if (
    quality.replacementCharacterCount > 0 &&
    quality.replacementCharacterCount / quality.characterCount > 0.02
  ) {
    return false;
  }

  return true;
}

function shouldOcrPdfText(text) {
  if (!hasUsableNativePdfText(text)) {
    return true;
  }

  const data = extractTaxFields(text);

  if (data.form_type === "BIR Form 2307") {
    return !isHighConfidence2307(data);
  }

  /*
    Even when the PDF has a valid text layer, OCR can still help if
    the tax parser cannot recover enough of the expected fields.
  */
  return scoreExtractedData(data) < 5;
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
  const sourceWidth = sourceCanvas.width;

  const sourceHeight = sourceCanvas.height;

  const sx = Math.max(0, Math.floor(sourceWidth * region.x));

  const sy = Math.max(0, Math.floor(sourceHeight * region.y));

  const sw = Math.min(sourceWidth - sx, Math.ceil(sourceWidth * region.width));

  const sh = Math.min(
    sourceHeight - sy,
    Math.ceil(sourceHeight * region.height),
  );

  const canvas = document.createElement("canvas");

  canvas.width = sw;
  canvas.height = sh;

  const context = canvas.getContext("2d", {
    willReadFrequently: true,
  });

  if (!context) {
    throw new Error("Unable to create OCR crop.");
  }

  context.fillStyle = "#ffffff";

  context.fillRect(0, 0, sw, sh);

  context.drawImage(sourceCanvas, sx, sy, sw, sh, 0, 0, sw, sh);

  return canvas;
}

function upscaleCanvas(
  sourceCanvas,
  scale = PDF_IDENTITY_UPSCALE,
  padding = 18,
) {
  const canvas = document.createElement("canvas");

  const targetWidth = Math.max(1, Math.round(sourceCanvas.width * scale));
  const targetHeight = Math.max(1, Math.round(sourceCanvas.height * scale));

  canvas.width = targetWidth + padding * 2;
  canvas.height = targetHeight + padding * 2;

  const context = canvas.getContext("2d", {
    willReadFrequently: true,
  });

  if (!context) {
    throw new Error("Unable to enlarge OCR crop.");
  }

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);

  /*
    Keep edges crisp. Tesseract performs better on the small BIR box fonts
    when the original pixels are enlarged without browser interpolation.
  */
  context.imageSmoothingEnabled = false;
  context.drawImage(
    sourceCanvas,
    0,
    0,
    sourceCanvas.width,
    sourceCanvas.height,
    padding,
    padding,
    targetWidth,
    targetHeight,
  );

  return canvas;
}

function preprocessOcrCanvas(sourceCanvas, threshold = 205) {
  const canvas = document.createElement("canvas");

  canvas.width = sourceCanvas.width;

  canvas.height = sourceCanvas.height;

  const context = canvas.getContext("2d", {
    willReadFrequently: true,
  });

  if (!context) {
    throw new Error("Unable to preprocess OCR image.");
  }

  context.drawImage(sourceCanvas, 0, 0);

  const imageData = context.getImageData(0, 0, canvas.width, canvas.height);

  const pixels = imageData.data;

  for (let index = 0; index < pixels.length; index += 4) {
    const red = pixels[index];

    const green = pixels[index + 1];

    const blue = pixels[index + 2];

    const gray = red * 0.299 + green * 0.587 + blue * 0.114;

    /*
      Increase contrast before applying
      the black/white threshold.
    */
    const contrast = (gray - 128) * 1.2 + 128;

    const value = contrast < threshold ? 0 : 255;

    pixels[index] = value;
    pixels[index + 1] = value;
    pixels[index + 2] = value;
    pixels[index + 3] = 255;
  }

  context.putImageData(imageData, 0, 0);

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
   OCR HELPERS
   ========================================================= */

const OCR_GENERAL_WHITELIST =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 .,:;'-/&()[]_%";

async function setOcrParameters(worker, mode, options = {}) {
  /*
    Tesseract worker parameters persist between recognize() calls.

    Never reset tessedit_char_whitelist to an empty string: in browser
    Tesseract that can effectively leave the next recognition pass with no
    usable characters. Instead, explicitly restore a broad text whitelist
    after numeric-only TIN OCR.
  */
  const parameters = {
    tessedit_pageseg_mode: String(mode),
    preserve_interword_spaces: "1",
    user_defined_dpi: "300",
    tessedit_char_whitelist: options.whitelist ?? OCR_GENERAL_WHITELIST,
    classify_bln_numeric_mode: options.numericMode ? "1" : "0",
  };

  try {
    await worker.setParameters(parameters);
  } catch (error) {
    console.warn("Unable to apply all OCR parameters:", error);
  }
}

async function setOcrPsm(worker, mode) {
  await setOcrParameters(worker, mode);
}

async function recognizeCanvas(worker, canvas) {
  const result = await worker.recognize(canvas);

  return result?.data?.text || "";
}

/*
  BIR 2307 identity extraction
  ----------------------------

  The values are inside small bordered boxes. Reading the whole row at once
  is fragile, so names are enlarged and TINs are split into the four printed
  groups (3-3-3-5) before OCR. This removes most interference from the form
  separators and prevents a valid-but-wrong whole-row OCR value from winning.
*/
function chooseMostFrequent(values) {
  const counts = new Map();

  for (const value of values.filter(Boolean)) {
    counts.set(value, (counts.get(value) || 0) + 1);
  }

  return (
    [...counts.entries()].sort((a, b) => {
      if (b[1] !== a[1]) {
        return b[1] - a[1];
      }

      return values.indexOf(a[0]) - values.indexOf(b[0]);
    })[0]?.[0] || ""
  );
}

function score2307NameCandidate(value = "") {
  const candidate = cleanExtractedName(value)
    .replace(/^\d+\s*/, "")
    .replace(/\s+\d+\s*$/, "")
    .trim();

  if (!isLikelyPersonOrCompanyName(candidate)) {
    return -Infinity;
  }

  const letters = candidate.replace(/[^A-Za-z]/g, "").length;
  const digits = candidate.replace(/\D/g, "").length;
  const words = candidate.split(/\s+/).filter(Boolean).length;

  /*
    Real 2307 names are predominantly letters and normally contain multiple
    words. Penalize numeric / border artifacts heavily.
  */
  return letters + words * 20 - digits * 25;
}

async function recognize2307NameField(worker, rawCanvas) {
  let enlargedCanvas = null;
  let thresholdLow = null;
  let thresholdHigh = null;

  const attempts = [];

  try {
    /*
      The source crop is already the value row only. Keep the ORIGINAL crop
      as the first OCR variant because it preserves the thin BIR characters.
    */
    enlargedCanvas = upscaleCanvas(rawCanvas, 3, 18);
    thresholdLow = preprocessOcrCanvas(enlargedCanvas, 190);
    thresholdHigh = preprocessOcrCanvas(enlargedCanvas, 215);

    const variants = [rawCanvas, enlargedCanvas, thresholdLow, thresholdHigh];
    const whitelist =
      "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz .,'&/-()";

    for (const psm of [7, 13, 11, 6]) {
      await setOcrParameters(worker, psm, {
        whitelist,
        numericMode: false,
      });

      for (const variant of variants) {
        const rawText = await recognizeCanvas(worker, variant);
        const value = extract2307NameFromFieldCrop(rawText)
          .replace(/^\d+\s*/, "")
          .replace(/\s+\d+\s*$/, "")
          .trim();

        attempts.push({ psm, rawText, value });
      }
    }

    const candidates = [...new Set(attempts.map((attempt) => attempt.value))]
      .filter(Boolean)
      .sort((a, b) => score2307NameCandidate(b) - score2307NameCandidate(a));

    const best = candidates[0] || "";
    const bestAttempt = attempts.find((attempt) => attempt.value === best);

    return {
      value: best,
      rawText: bestAttempt?.rawText || attempts[0]?.rawText || "",
      attempts,
    };
  } finally {
    releaseCanvas(thresholdLow);
    releaseCanvas(thresholdHigh);
    releaseCanvas(enlargedCanvas);
  }
}

function normalize2307TinCandidate(rawText = "") {
  const digits = String(rawText).replace(/\D/g, "");

  if (digits.length === 14) {
    return formatTin(digits);
  }

  /*
    On the actual uploaded PDF, Tesseract reliably reads the Payor TIN as:
      2368090450000
    i.e. the 5-zero branch is collapsed to 4 zeroes.

    The first 9 digits are the base TIN. If the remaining branch consists
    only of zeroes, safely restore it to the printed 5-digit branch field.
  */
  if (digits.length >= 10 && digits.length < 14) {
    const base = digits.slice(0, 9);
    const branch = digits.slice(9);

    if (base.length === 9 && branch.length >= 1 && /^0+$/.test(branch)) {
      return formatTin(`${base}${branch.padEnd(5, "0")}`);
    }
  }

  /*
    Occasionally a border / row number is recognized as an extra leading
    digit. Try every 14-digit window and prefer one ending in the standard
    five-digit branch field.
  */
  if (digits.length > 14) {
    const candidates = [];

    for (let index = 0; index <= digits.length - 14; index += 1) {
      const candidate = digits.slice(index, index + 14);
      const branch = candidate.slice(9);

      if (/^\d{9}\d{5}$/.test(candidate)) {
        candidates.push({
          value: formatTin(candidate),
          zeroBranch: /^0{5}$/.test(branch),
        });
      }
    }

    const preferred = candidates.find((candidate) => candidate.zeroBranch);

    return preferred?.value || candidates[0]?.value || "";
  }

  return "";
}

function getMostFrequentCandidate(values = []) {
  const cleaned = values.filter(Boolean);

  if (!cleaned.length) {
    return "";
  }

  const counts = new Map();

  for (const value of cleaned) {
    counts.set(value, (counts.get(value) || 0) + 1);
  }

  return (
    [...counts.entries()].sort((a, b) => {
      if (b[1] !== a[1]) {
        return b[1] - a[1];
      }

      return cleaned.indexOf(a[0]) - cleaned.indexOf(b[0]);
    })[0]?.[0] || ""
  );
}

function resolve2307Tin(dedicatedValue, broadValues = []) {
  const broad = broadValues.filter((value) => isValidExtractedTin(value));
  const dedicated = isValidExtractedTin(dedicatedValue) ? dedicatedValue : "";

  /*
    Prefer agreement from broad Payee/Payor OCR over one isolated field OCR.

    This directly protects values such as 297 being misread as 207 because
    the isolated crop contains a thin box line. If two broad passes agree,
    they beat one conflicting dedicated result.
  */
  return getMostFrequentCandidate([...broad, dedicated]);
}

function resolve2307Name(dedicatedValue, broadValues = []) {
  const candidates = [...broadValues, dedicatedValue]
    .map((value) => cleanExtractedName(value || ""))
    .filter((value) => isLikelyPersonOrCompanyName(value));

  if (!candidates.length) {
    return "";
  }

  const counts = new Map();

  for (const candidate of candidates) {
    const key = candidate.toUpperCase();
    const existing = counts.get(key) || { value: candidate, count: 0 };
    existing.count += 1;
    counts.set(key, existing);
  }

  return (
    [...counts.values()].sort((a, b) => {
      if (b.count !== a.count) {
        return b.count - a.count;
      }

      return score2307NameCandidate(b.value) - score2307NameCandidate(a.value);
    })[0]?.value || ""
  );
}

function get2307IdentityEvidence(text = "") {
  const tins = findAllTins(text).filter((value) => isValidExtractedTin(value));

  /*
    IMPORTANT:

    Never assign a single broadly-recognized TIN to Payee just because it is
    the first TIN found. On this exact 2307 sample, broad OCR can see only the
    Payor TIN (236-809-045-00000). The previous implementation then placed it
    into payeeTin because tins[0] was assumed to be the Payee.

    Broad OCR is only allowed to provide positional TIN evidence when BOTH
    Payee and Payor TINs were found. Otherwise the dedicated field crops are
    the source of truth and broad OCR remains only supporting evidence.
  */
  const hasBothRoleTins = tins.length >= 2;

  return {
    payeeTin: hasBothRoleTins ? tins[0] : "",
    payorTin: hasBothRoleTins ? tins[1] : "",
    payeeName: find2307PayeeName(text),
    payorName: find2307PayorName(text),
  };
}

async function recognize2307TinField(worker, rawCanvas) {
  let enlargedCanvas = null;
  let thresholdLow = null;
  let thresholdHigh = null;

  const attempts = [];

  async function runTinAttempt(canvas, psm, variant, weight) {
    await setOcrParameters(worker, psm, {
      whitelist: "0123456789",
      numericMode: true,
    });

    const rawText = await recognizeCanvas(worker, canvas);
    const value = normalize2307TinCandidate(rawText);

    const attempt = {
      psm,
      variant,
      weight,
      rawText,
      digits: String(rawText || "").replace(/\D/g, ""),
      value,
    };

    attempts.push(attempt);

    return attempt;
  }

  try {
    /*
      FIRST PASS: read the untouched field crop as a single line.

      The page-1 TIN images are already extremely clean. Hard thresholding can
      turn a thin 9 into 0/2 or add a border as a leading 1. Previously all
      OCR variants had equal voting power, so several damaged threshold passes
      could outvote the correct raw image.

      If PSM 7 and PSM 6 agree on the RAW crop, return that value immediately.
      This is the strongest evidence for this standard BIR 2307 page.
    */
    const rawPsm7 = await runTinAttempt(rawCanvas, 7, "raw", 12);
    const rawPsm6 = await runTinAttempt(rawCanvas, 6, "raw", 12);

    if (
      rawPsm7.value &&
      rawPsm6.value &&
      rawPsm7.value === rawPsm6.value &&
      isValidExtractedTin(rawPsm7.value)
    ) {
      return {
        value: rawPsm7.value,
        rawText: rawPsm7.rawText || rawPsm6.rawText,
        method: "raw-consensus-psm7-psm6",
        groups: [],
        attempts,
      };
    }

    const rawPsm13 = await runTinAttempt(rawCanvas, 13, "raw", 8);

    const rawConsensus = getMostFrequentCandidate(
      [rawPsm7.value, rawPsm6.value, rawPsm13.value].filter((value) =>
        isValidExtractedTin(value),
      ),
    );

    const rawConsensusCount = [
      rawPsm7.value,
      rawPsm6.value,
      rawPsm13.value,
    ].filter((value) => value === rawConsensus).length;

    if (rawConsensus && rawConsensusCount >= 2) {
      const winningAttempt = attempts.find(
        (attempt) => attempt.value === rawConsensus,
      );

      return {
        value: rawConsensus,
        rawText: winningAttempt?.rawText || rawPsm7.rawText || "",
        method: "raw-consensus",
        groups: [],
        attempts,
      };
    }

    /*
      FALLBACK PASSES:
      - enlarged crop gets medium weight
      - thresholded crops get low weight

      Threshold images are useful when the scan is faint, but they must never
      overpower a readable original crop merely because there are more of them.
    */
    enlargedCanvas = upscaleCanvas(rawCanvas, 3, 18);
    thresholdLow = preprocessOcrCanvas(enlargedCanvas, 185);
    thresholdHigh = preprocessOcrCanvas(enlargedCanvas, 215);

    const variants = [
      { name: "enlarged", canvas: enlargedCanvas, weight: 6 },
      { name: "threshold-185", canvas: thresholdLow, weight: 2 },
      { name: "threshold-215", canvas: thresholdHigh, weight: 2 },
    ];

    for (const psm of [7, 6, 13, 11]) {
      const psmWeight = psm === 7 || psm === 6 ? 3 : psm === 13 ? 2 : 1;

      for (const variant of variants) {
        await runTinAttempt(
          variant.canvas,
          psm,
          variant.name,
          variant.weight * psmWeight,
        );
      }
    }

    const scores = new Map();

    for (const attempt of attempts) {
      if (!isValidExtractedTin(attempt.value)) {
        continue;
      }

      scores.set(
        attempt.value,
        (scores.get(attempt.value) || 0) + Number(attempt.weight || 1),
      );
    }

    const value =
      [...scores.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || "";

    const winningAttempt = attempts.find((attempt) => attempt.value === value);

    return {
      value,
      rawText: winningAttempt?.rawText || attempts[0]?.rawText || "",
      method: value ? "weighted-whole-row" : "failed",
      groups: [],
      attempts,
    };
  } finally {
    /* Restore text mode before the next OCR call on this reused worker. */
    await setOcrParameters(worker, 11, {
      whitelist: OCR_GENERAL_WHITELIST,
      numericMode: false,
    });

    releaseCanvas(thresholdLow);
    releaseCanvas(thresholdHigh);
    releaseCanvas(enlargedCanvas);
  }
}

function isValidCalendarDate(month, day, year) {
  if (
    !Number.isInteger(month) ||
    !Number.isInteger(day) ||
    !Number.isInteger(year) ||
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > 31 ||
    year < 2000 ||
    year > CURRENT_YEAR + 1
  ) {
    return false;
  }

  const date = new Date(year, month - 1, day);

  return (
    date.getFullYear() === year &&
    date.getMonth() === month - 1 &&
    date.getDate() === day
  );
}

function parseSingleDateFromDigitString(rawText) {
  const digits = String(rawText || "").replace(/\D/g, "");

  if (digits.length < 8) {
    return null;
  }

  /*
    OCR commonly adds a border as a leading 1. Scan every 8-digit window
    instead of assuming the first eight characters are MMDDYYYY.
  */
  for (let start = 0; start <= digits.length - 8; start += 1) {
    const chunk = digits.slice(start, start + 8);
    const month = Number(chunk.slice(0, 2));
    const day = Number(chunk.slice(2, 4));
    const year = Number(chunk.slice(4, 8));

    if (isValidCalendarDate(month, day, year)) {
      return {
        month,
        day,
        year,
        formatted: formatDateParts(month, day, year),
      };
    }
  }

  return null;
}

async function recognize2307SingleDateField(worker, rawCanvas) {
  let enlargedCanvas = null;
  let thresholdLow = null;
  let thresholdHigh = null;
  const attempts = [];

  try {
    enlargedCanvas = upscaleCanvas(rawCanvas, 3, 12);
    thresholdLow = preprocessOcrCanvas(enlargedCanvas, 185);
    thresholdHigh = preprocessOcrCanvas(enlargedCanvas, 220);

    const variants = [rawCanvas, enlargedCanvas, thresholdLow, thresholdHigh];

    for (const psm of [7, 13, 6, 11]) {
      await setOcrParameters(worker, psm, {
        whitelist: "0123456789",
        numericMode: true,
      });

      for (const variant of variants) {
        const rawText = await recognizeCanvas(worker, variant);
        const value = parseSingleDateFromDigitString(rawText);

        attempts.push({ psm, rawText, value });
      }
    }

    const formattedValues = attempts
      .map((attempt) => attempt.value?.formatted || "")
      .filter(Boolean);

    const winningFormatted = getMostFrequentCandidate(formattedValues);
    const winningAttempt = attempts.find(
      (attempt) => attempt.value?.formatted === winningFormatted,
    );
    const voteCount = winningFormatted
      ? formattedValues.filter((value) => value === winningFormatted).length
      : 0;

    return {
      value: winningAttempt?.value || null,
      voteCount,
      rawText: winningAttempt?.rawText || attempts[0]?.rawText || "",
      attempts,
    };
  } finally {
    await setOcrParameters(worker, 11, {
      whitelist: OCR_GENERAL_WHITELIST,
      numericMode: false,
    });

    releaseCanvas(thresholdLow);
    releaseCanvas(thresholdHigh);
    releaseCanvas(enlargedCanvas);
  }
}

function buildPeriodFromSingleDates(fromDate, toDate) {
  if (!fromDate || !toDate) {
    return null;
  }

  const from = new Date(fromDate.year, fromDate.month - 1, fromDate.day);
  const to = new Date(toDate.year, toDate.month - 1, toDate.day);

  if (to < from) {
    return null;
  }

  if (fromDate.year !== toDate.year) {
    return null;
  }

  return {
    from: fromDate.formatted,
    to: toDate.formatted,
    quarter: getQuarterFromMonth(toDate.month),
    year: toDate.year,
  };
}

function parsePeriodFromDigitString(rawText) {
  const digits = String(rawText || "").replace(/\D/g, "");

  if (digits.length < 16) {
    return null;
  }

  /* Try every 16-digit window and keep one containing two valid dates. */
  for (let start = 0; start <= digits.length - 16; start += 1) {
    const chunk = digits.slice(start, start + 16);

    const fromMonth = Number(chunk.slice(0, 2));
    const fromDay = Number(chunk.slice(2, 4));
    const fromYear = Number(chunk.slice(4, 8));
    const toMonth = Number(chunk.slice(8, 10));
    const toDay = Number(chunk.slice(10, 12));
    const toYear = Number(chunk.slice(12, 16));

    const validFrom =
      fromMonth >= 1 &&
      fromMonth <= 12 &&
      fromDay >= 1 &&
      fromDay <= 31 &&
      fromYear >= 2000 &&
      fromYear <= CURRENT_YEAR + 1;

    const validTo =
      toMonth >= 1 &&
      toMonth <= 12 &&
      toDay >= 1 &&
      toDay <= 31 &&
      toYear >= 2000 &&
      toYear <= CURRENT_YEAR + 1;

    if (validFrom && validTo) {
      return {
        from: formatDateParts(fromMonth, fromDay, fromYear),
        to: formatDateParts(toMonth, toDay, toYear),
        quarter: getQuarterFromMonth(toMonth),
        year: toYear,
      };
    }
  }

  return null;
}

async function recognize2307PeriodField(worker, rawCanvas) {
  let enlargedCanvas = null;
  let thresholdCanvas = null;
  const attempts = [];

  try {
    enlargedCanvas = upscaleCanvas(rawCanvas, 3, 18);
    thresholdCanvas = preprocessOcrCanvas(enlargedCanvas, 205);

    for (const psm of [6, 11, 7]) {
      await setOcrParameters(worker, psm, {
        whitelist: "0123456789 ",
        numericMode: true,
      });

      for (const variant of [enlargedCanvas, thresholdCanvas]) {
        const rawText = await recognizeCanvas(worker, variant);
        const value = parsePeriodFromDigitString(rawText);

        attempts.push({ psm, rawText, value });

        if (value) {
          return { value, rawText, attempts };
        }
      }
    }

    return {
      value: null,
      rawText: attempts[0]?.rawText || "",
      attempts,
    };
  } finally {
    releaseCanvas(thresholdCanvas);
    releaseCanvas(enlargedCanvas);
  }
}

function merge2307PeriodOverride(data, periodOverride) {
  if (!periodOverride || data?.form_type !== "BIR Form 2307") {
    return data;
  }

  return {
    ...data,
    tax_period: {
      ...data.tax_period,
      ...periodOverride,
    },
  };
}

async function classifyPdfPageForOcr(page, getOcrWorker, onProgress) {
  const worker = await getOcrWorker();

  let sourceCanvas = null;

  let headerCrop = null;

  let preparedHeader = null;

  try {
    onProgress?.("Identifying BIR form");

    sourceCanvas = await renderPdfPageToCanvas(page, PDF_CLASSIFIER_SCALE);

    headerCrop = cropCanvasByRatio(sourceCanvas, PDF_HEADER_OCR_REGION);

    preparedHeader = preprocessOcrCanvas(headerCrop, 215);

    await setOcrPsm(worker, 11);

    const text = await recognizeCanvas(worker, preparedHeader);

    const cleanedText = cleanText(text);

    return {
      text: cleanedText,
      formType: findFormType(cleanedText),
    };
  } finally {
    releaseCanvas(preparedHeader);

    releaseCanvas(headerCrop);

    releaseCanvas(sourceCanvas);
  }
}

/* =========================================================
   SMART 2307 OCR
   ========================================================= */

async function ocr2307PageRegions(page, getOcrWorker, onProgress) {
  const worker = await getOcrWorker();

  let sourceCanvas = null;

  let periodCrop = null;
  let fromDateCrop = null;
  let toDateCrop = null;
  let informationCrop = null;
  let payeeTinCrop = null;
  let payeeNameCrop = null;
  let payorTinCrop = null;
  let payorNameCrop = null;
  let tableCrop = null;

  let preparedPeriod = null;
  let preparedInformation = null;
  let preparedTable = null;
  let preparedFullPage = null;

  let identityOverrides = null;
  let periodOverride = null;

  try {
    onProgress?.("Rendering tax form");

    sourceCanvas = await renderPdfPageToCanvas(page);

    /* --------------------------------
       PERIOD
       -------------------------------- */

    onProgress?.("Reading tax period");

    periodCrop = cropCanvasByRatio(sourceCanvas, FORM_2307_OCR_REGIONS.period);
    fromDateCrop = cropCanvasByRatio(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.fromDate,
    );
    toDateCrop = cropCanvasByRatio(sourceCanvas, FORM_2307_OCR_REGIONS.toDate);
    preparedPeriod = preprocessOcrCanvas(periodCrop, 205);

    /*
      First read From and To independently. This avoids the old bug where
      every digit in the whole row was concatenated and a valid-looking but
      wrong 16-digit window (for example 04/04/2018 -> 01/01/2026) won.
    */
    const fromDateResult = await recognize2307SingleDateField(
      worker,
      fromDateCrop,
    );
    const toDateResult = await recognize2307SingleDateField(worker, toDateCrop);

    periodOverride =
      fromDateResult.voteCount >= 2 && toDateResult.voteCount >= 2
        ? buildPeriodFromSingleDates(fromDateResult.value, toDateResult.value)
        : null;

    /*
      Keep two general-text OCR passes as evidence. They are especially useful
      on photographed pages where the fixed date crops may be shifted.
    */
    await setOcrPsm(worker, 13);
    const periodTextPsm13 = await recognizeCanvas(worker, preparedPeriod);

    await setOcrPsm(worker, 11);
    const periodTextPsm11 = await recognizeCanvas(worker, preparedPeriod);

    const periodText = cleanText(`${periodTextPsm13}\n${periodTextPsm11}`);

    if (!periodOverride) {
      const periodCandidates = [periodTextPsm13, periodTextPsm11]
        .map((text) => find2307Period(text))
        .filter((value) => value?.from && value?.to && value?.year);

      if (periodCandidates.length) {
        const grouped = new Map();

        for (const candidate of periodCandidates) {
          const key = `${candidate.from}|${candidate.to}`;
          const current = grouped.get(key) || { value: candidate, count: 0 };
          current.count += 1;
          grouped.set(key, current);
        }

        periodOverride =
          [...grouped.values()].sort((a, b) => b.count - a.count)[0]?.value ||
          null;
      }
    }

    /*
      Last fallback: retain the old full-period numeric OCR, but only use it
      when the safer single-date and text strategies both fail.
    */
    let legacyPeriodResult = null;

    if (!periodOverride) {
      legacyPeriodResult = await recognize2307PeriodField(worker, periodCrop);
      periodOverride = legacyPeriodResult.value;
    }

    console.debug("[2307 OCR period]", {
      parsed: periodOverride,
      fromDate: {
        raw: cleanText(fromDateResult.rawText),
        parsed: fromDateResult.value,
      },
      toDate: {
        raw: cleanText(toDateResult.rawText),
        parsed: toDateResult.value,
      },
      periodText: cleanText(periodText),
      legacy: legacyPeriodResult
        ? {
            raw: cleanText(legacyPeriodResult.rawText),
            parsed: legacyPeriodResult.value,
          }
        : null,
    });

    /* --------------------------------
       BROAD PAYEE / PAYOR EVIDENCE
       -------------------------------- */

    onProgress?.("Reading Payee and Payor information");

    informationCrop = cropCanvasByRatio(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.information,
    );
    preparedInformation = preprocessOcrCanvas(informationCrop, 205);

    await setOcrPsm(worker, 11);
    const informationTextPsm11 = await recognizeCanvas(
      worker,
      preparedInformation,
    );

    await setOcrPsm(worker, 6);
    const informationTextPsm6 = await recognizeCanvas(
      worker,
      preparedInformation,
    );

    const informationText = cleanText(
      `${informationTextPsm11}\n${informationTextPsm6}`,
    );

    const broadIdentityEvidence = [
      get2307IdentityEvidence(informationTextPsm11),
      get2307IdentityEvidence(informationTextPsm6),
    ];

    /* --------------------------------
       DEDICATED IDENTITY FIELDS
       -------------------------------- */

    onProgress?.("Reading Payee name and TIN");

    payeeTinCrop = cropCanvasByRatio(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.payeeTin,
    );
    payeeNameCrop = cropCanvasByRatio(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.payeeName,
    );
    payorTinCrop = cropCanvasByRatio(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.payorTin,
    );
    payorNameCrop = cropCanvasByRatio(
      sourceCanvas,
      FORM_2307_OCR_REGIONS.payorName,
    );

    const payeeTinResult = await recognize2307TinField(worker, payeeTinCrop);
    const payeeNameResult = await recognize2307NameField(worker, payeeNameCrop);

    onProgress?.("Reading Payor name and TIN");

    const payorTinResult = await recognize2307TinField(worker, payorTinCrop);
    const payorNameResult = await recognize2307NameField(worker, payorNameCrop);

    /*
      Do not blindly trust the isolated crop. The old version accepted a
      syntactically valid but incorrect value such as 207-944-186-00000 and
      overwrote the correct broad OCR result 297-944-186-00000.

      Use broad PSM 11 + PSM 6 evidence first and the dedicated crop as a
      third vote / fallback.
    */
    const broadPayeeTins = broadIdentityEvidence
      .map((item) => item.payeeTin)
      .filter((value) => isValidExtractedTin(value));
    const broadPayorTins = broadIdentityEvidence
      .map((item) => item.payorTin)
      .filter((value) => isValidExtractedTin(value));
    const broadPayeeNames = broadIdentityEvidence
      .map((item) => item.payeeName)
      .filter((value) => isLikelyPersonOrCompanyName(value));
    const broadPayorNames = broadIdentityEvidence
      .map((item) => item.payorName)
      .filter((value) => isLikelyPersonOrCompanyName(value));

    /*
      A fixed-position crop is trusted only when the broader information
      region also sees the same kind of field. This is important for the
      photographed pages later in the sample PDF: their form position shifts,
      so a fixed crop can land on an unrelated number and still look like a
      syntactically valid TIN.
    */
    /*
      Dedicated crops are authoritative for the standard page-1 test form.

      The previous code accidentally DISCARDED a correct dedicated field when
      broad OCR failed to see the same field. That is why names became blank.
      It also let one broad Payor TIN be mislabeled as the Payee TIN.

      resolve2307Tin / resolve2307Name already support an empty broad array,
      so always pass the dedicated result and use broad OCR only as a vote when
      role-safe broad evidence is actually available.
    */
    const payeeTin = resolve2307Tin(payeeTinResult.value, broadPayeeTins);

    const payorTin = resolve2307Tin(payorTinResult.value, broadPayorTins);

    const payeeName = resolve2307Name(payeeNameResult.value, broadPayeeNames);

    const payorName = resolve2307Name(payorNameResult.value, broadPayorNames);

    identityOverrides = {
      payeeName,
      payeeTin,
      payorName,
      payorTin,
    };

    console.debug("[2307 OCR identity fields]", {
      broad: broadIdentityEvidence,
      dedicated: {
        payeeName: payeeNameResult.value,
        payeeTin: payeeTinResult.value,
        payorName: payorNameResult.value,
        payorTin: payorTinResult.value,
      },
      resolved: identityOverrides,
      diagnostics: {
        payeeTinAttempts: payeeTinResult.attempts?.map((attempt) => ({
          psm: attempt.psm,
          text: cleanText(attempt.rawText),
          digits: attempt.digits,
          value: attempt.value,
        })),
        payorTinAttempts: payorTinResult.attempts?.map((attempt) => ({
          psm: attempt.psm,
          text: cleanText(attempt.rawText),
          digits: attempt.digits,
          value: attempt.value,
        })),
      },
    });

    const identityText = cleanText(`
      Taxpayer Identification Number (TIN)
      ${payeeTin || payeeTinResult.rawText}
      Payee's Name
      ${payeeName || payeeNameResult.rawText}

      Taxpayer Identification Number (TIN)
      ${payorTin || payorTinResult.rawText}
      Payor's Name
      ${payorName || payorNameResult.rawText}
    `);

    /* --------------------------------
       INCOME TABLE
       -------------------------------- */

    onProgress?.("Reading income and withholding row");

    tableCrop = cropCanvasByRatio(sourceCanvas, FORM_2307_OCR_REGIONS.table);
    preparedTable = preprocessOcrCanvas(tableCrop, 205);

    await setOcrPsm(worker, 11);
    const tableTextPsm11 = await recognizeCanvas(worker, preparedTable);

    await setOcrPsm(worker, 6);
    const tableTextPsm6 = await recognizeCanvas(worker, preparedTable);

    const tableText = cleanText(`${tableTextPsm11}\n${tableTextPsm6}`);

    const periodEvidenceText = periodOverride
      ? `For the Period From ${periodOverride.from} To ${periodOverride.to}`
      : periodText;

    const regionText = cleanText(`
      BIR Form 2307

      ${periodEvidenceText}

      ${identityText}

      ${informationText}

      Part III - Details of Monthly Income Payments and Taxes Withheld

      ${tableText}
    `);

    let regionData = extractTaxFields(regionText);
    regionData = merge2307IdentityOverrides(regionData, identityOverrides);
    regionData = merge2307PeriodOverride(regionData, periodOverride);

    const hasReliableRegionalIdentity =
      Boolean(payeeName) &&
      isValidExtractedTin(payeeTin) &&
      Boolean(payorName) &&
      isValidExtractedTin(payorTin);

    if (hasReliableRegionalIdentity && isHighConfidence2307(regionData)) {
      return {
        text: regionText,
        method: "ocr-regions-consensus",
        identityOverrides,
        periodOverride,
      };
    }

    /* --------------------------------
       FALLBACK VERIFICATION
       -------------------------------- */

    onProgress?.("Verifying difficult fields");

    preparedFullPage = preprocessOcrCanvas(sourceCanvas, 210);

    const fullPageTexts = [];

    for (const psm of [11, 6]) {
      await setOcrPsm(worker, psm);

      const thresholdText = await recognizeCanvas(worker, preparedFullPage);
      const rawText = await recognizeCanvas(worker, sourceCanvas);

      fullPageTexts.push(cleanText(thresholdText));
      fullPageTexts.push(cleanText(rawText));
    }

    const candidates = [
      regionText,
      ...fullPageTexts,
      ...fullPageTexts.map((text) => cleanText(`${regionText}\n${text}`)),
    ];

    const best = chooseBestText(candidates);

    return {
      text: best.text || regionText,
      method: "ocr-regions-consensus+verification",
      identityOverrides,
      periodOverride,
    };
  } finally {
    releaseCanvas(preparedPeriod);
    releaseCanvas(preparedInformation);
    releaseCanvas(preparedTable);
    releaseCanvas(preparedFullPage);

    releaseCanvas(periodCrop);
    releaseCanvas(fromDateCrop);
    releaseCanvas(toDateCrop);
    releaseCanvas(informationCrop);
    releaseCanvas(payeeTinCrop);
    releaseCanvas(payeeNameCrop);
    releaseCanvas(payorTinCrop);
    releaseCanvas(payorNameCrop);
    releaseCanvas(tableCrop);
    releaseCanvas(sourceCanvas);
  }
}

/* =========================================================
   GENERIC FULL PAGE OCR FALLBACK
   ========================================================= */

async function ocrPdfPage(page, getOcrWorker) {
  const worker = await getOcrWorker();

  let canvas = null;
  let prepared = null;

  try {
    canvas = await renderPdfPageToCanvas(page);

    prepared = preprocessOcrCanvas(canvas, 210);

    await setOcrPsm(worker, 11);

    return await recognizeCanvas(worker, prepared);
  } finally {
    releaseCanvas(prepared);

    releaseCanvas(canvas);
  }
}

/* =========================================================
   CHOOSE BEST OCR TEXT
   ========================================================= */

function chooseBestText(candidates) {
  let best = {
    text: "",
    score: -1,
    data: null,
  };

  for (const candidate of candidates) {
    if (!candidate?.trim()) {
      continue;
    }

    const data = extractTaxFields(candidate);

    const score = scoreExtractedData(data);

    if (score > best.score) {
      best = {
        text: candidate,
        score,
        data,
      };
    }
  }

  return best;
}

/* =========================================================
   DOCUMENT REFERENCE FALLBACK
   ========================================================= */

function merge2307WithReference(currentData) {
  /*
    Do not copy values from another page.

    This PDF contains different withholding agents, ATCs and payment details
    on later pages. Reusing page 1 can silently create believable but wrong
    records. Missing OCR data is safer than fabricated cross-page data.
  */
  return currentData;
}

/* =========================================================
   PDF PAGE-BY-PAGE EXTRACTION
   ========================================================= */

async function extractPdfPages(file, { getOcrWorker, onProgress }) {
  const arrayBuffer = await file.arrayBuffer();

  const loadingTask = pdfjsLib.getDocument({
    data: arrayBuffer,
  });

  const pdf = await loadingTask.promise;

  const pages = [];

  const pageLimit = Math.min(pdf.numPages, DOCUMENT_PAGE_LIMIT);

  try {
    for (let pageNumber = 1; pageNumber <= pageLimit; pageNumber += 1) {
      onProgress?.(`Reading page ${pageNumber} of ${pageLimit}`);

      const page = await pdf.getPage(pageNumber);

      try {
        const textContent = await page.getTextContent({
          disableNormalization: false,
        });

        /*
          Keep visual line order instead of flattening every PDF text item.
          This helps TIN groups, dates, labels, and table values stay related.
        */
        const directText = extractPdfTextWithLayout(textContent);

        const directData = extractTaxFields(directText);

        let finalText = directText;

        let extractionMethod = "text-layout";

        let identityOverrides = null;
        let periodOverride = null;

        if (shouldOcrPdfText(directText)) {
          onProgress?.(`Analyzing page ${pageNumber} of ${pageLimit}`);

          let detectedFormType = directData.form_type;

          let classifierText = "";

          if (!detectedFormType) {
            const classification = await classifyPdfPageForOcr(
              page,
              getOcrWorker,
              (message) => {
                onProgress?.(`${message} - page ${pageNumber} of ${pageLimit}`);
              },
            );

            detectedFormType = classification.formType;

            classifierText = classification.text;
          }

          if (
            detectedFormType === "BIR Form 2307" ||
            looksLike2307(directText) ||
            looksLike2307(classifierText)
          ) {
            onProgress?.(
              `Scanning BIR Form 2307 - page ${pageNumber} of ${pageLimit}`,
            );

            const specialized = await ocr2307PageRegions(
              page,
              getOcrWorker,
              (message) => {
                onProgress?.(`${message} - page ${pageNumber} of ${pageLimit}`);
              },
            );

            const candidates = [
              directText,
              specialized.text,
              cleanText(`${classifierText}\n${specialized.text}`),
            ];

            const best = chooseBestText(candidates);

            finalText = best.text || specialized.text || directText;

            extractionMethod = specialized.method;

            identityOverrides = specialized.identityOverrides || null;
            periodOverride = specialized.periodOverride || null;
          } else {
            onProgress?.(`Scanning page ${pageNumber} of ${pageLimit}`);

            const ocrText = await ocrPdfPage(page, getOcrWorker);

            const candidates = [
              directText,
              ocrText,
              cleanText(`${directText}\n${ocrText}`),
            ];

            const best = chooseBestText(candidates);

            finalText = best.text || ocrText || directText;

            extractionMethod = "ocr-full-page";
          }
        }

        let data = extractTaxFields(finalText);

        if (identityOverrides) {
          data = merge2307IdentityOverrides(data, identityOverrides);
        }

        if (periodOverride) {
          data = merge2307PeriodOverride(data, periodOverride);
        }

        pages.push({
          pageNumber,
          text: cleanText(finalText),
          method: extractionMethod,
          preParsedData: data,
          sourceTotalPages: pdf.numPages,
          processedPageCount: pageLimit,
        });
      } finally {
        if (typeof page?.cleanup === "function") {
          try {
            page.cleanup();
          } catch (error) {
            console.warn("PDF page cleanup warning:", error);
          }
        }
      }
    }
  } finally {
    // Some pdfjs-dist versions do not expose pdf.destroy().
    // Destroy the loading task when available and treat cleanup failures as non-fatal.
    if (typeof pdf?.cleanup === "function") {
      try {
        pdf.cleanup();
      } catch (error) {
        console.warn("PDF cleanup warning:", error);
      }
    }

    if (typeof loadingTask?.destroy === "function") {
      try {
        await loadingTask.destroy();
      } catch (error) {
        console.warn("PDF loading task destroy warning:", error);
      }
    }
  }

  return pages;
}

/* =========================================================
   DOCX RECORD SPLITTING
   ========================================================= */

function splitDocxTextIntoRecords(rawText) {
  const text = cleanText(rawText);

  if (!text) {
    return [];
  }

  const formRegex =
    /\b(?:BIR\s+Form(?:\s+No\.?)?\s*)?(?:2307|1701Q|1701A|1701|1601[\s-]?EQ|2551Q|2550Q)\b/gi;

  const indexes = [];

  let match;

  while ((match = formRegex.exec(text)) !== null) {
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

  const records = [];

  if (indexes[0] > 0) {
    indexes[0] = 0;
  }

  for (let index = 0; index < indexes.length; index += 1) {
    const start = indexes[index];

    const end = index + 1 < indexes.length ? indexes[index + 1] : text.length;

    const recordText = cleanText(text.slice(start, end));

    if (recordText) {
      records.push({
        pageNumber: records.length + 1,

        text: recordText,

        method: "docx",
      });
    }
  }

  return records;
}

/* =========================================================
   DOCX EXTRACTION
   ========================================================= */

async function extractDocxPages(file, { onProgress }) {
  onProgress?.("Reading Word document");

  const arrayBuffer = await file.arrayBuffer();

  const result = await mammoth.extractRawText({
    arrayBuffer,
  });

  onProgress?.("Analyzing Word tax forms");

  const allRecords = splitDocxTextIntoRecords(result.value || "");

  const limited = allRecords.slice(0, DOCUMENT_PAGE_LIMIT);

  return limited.map((record) => ({
    ...record,

    sourceTotalPages: allRecords.length,

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
   EXTRACTION DEBUG
   ========================================================= */

function showExtractedDataInNetwork(records) {
  console.groupCollapsed(
    `[Tax extraction] ${records.length} extracted record(s)`,
  );
  console.log(records);
  console.groupEnd();
}

/* =========================================================
   EXCEL
   ========================================================= */

function getExcelRows(records) {
  return records.map((record) => {
    const data = record.data;

    return {
      "Form Type": data.form_type || "",

      "Taxpayer / Payee Name": data.taxpayer?.name || "",

      TIN: data.taxpayer?.tin || "",

      "Payor / Withholding Agent": data.withholding_agent?.name || "",

      "Payor TIN": data.withholding_agent?.tin || "",

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

      "Total Amount Payable": data.amounts?.total_amount_payable ?? "",
    };
  });
}

function exportRecordsToExcel(records) {
  const rows = getExcelRows(records);

  const worksheet = XLSX.utils.json_to_sheet(rows);

  worksheet["!cols"] = [
    { wch: 18 },
    { wch: 32 },
    { wch: 22 },
    { wch: 32 },
    { wch: 22 },
    { wch: 16 },
    { wch: 12 },
    { wch: 14 },
    { wch: 14 },
    { wch: 50 },
    { wch: 12 },
    { wch: 18 },
    { wch: 14 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 24 },
  ];

  for (let row = 2; row <= rows.length + 1; row += 1) {
    ["L", "N", "O", "P", "Q", "R"].forEach((column) => {
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

  const [files, setFiles] = useState([]);

  const [fileError, setFileError] = useState("");

  const [isDragging, setIsDragging] = useState(false);

  const [isExtracting, setIsExtracting] = useState(false);

  const [fileStatuses, setFileStatuses] = useState({});

  const [extractedRecords, setExtractedRecords] = useState([]);

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
     EXTRACTION
     ======================================================= */

  async function extractDocuments(documents) {
    if (!documents.length) {
      Toast.fire({
        icon: "warning",

        title: "Select at least one document",
      });

      return;
    }

    if (isExtracting) {
      return;
    }

    setIsExtracting(true);

    openExtractionLoading(documents.length);

    let ocrWorker = null;

    async function getOcrWorker() {
      if (!ocrWorker) {
        updateExtractionLoading(
          "Starting OCR engine",
          "OCR is loading once. Test mode processes page 1 only.",
        );

        ocrWorker = await createWorker("eng");
      }

      return ocrWorker;
    }

    const extractedNow = [];

    const processedFileKeys = new Set(documents.map(getFileKey));

    let reviewPageCount = 0;

    let failedFileCount = 0;

    let totalPageCount = 0;

    try {
      for (let fileIndex = 0; fileIndex < documents.length; fileIndex += 1) {
        const file = documents[fileIndex];

        const fileKey = getFileKey(file);

        try {
          updateFileStatus(file, "reading", "Reading document");

          updateExtractionLoading(
            `Reading ${file.name}`,
            `File ${fileIndex + 1} of ${documents.length}`,
          );

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

          for (let pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
            const page = pages[pageIndex];

            updateExtractionLoading(
              `Checking page ${page.pageNumber} of ${pages.length}`,
              `${file.name} • File ${fileIndex + 1} of ${documents.length}`,
            );

            if (!page.text?.trim()) {
              fileReviewCount += 1;

              reviewPageCount += 1;

              continue;
            }

            /*
              PDF extraction can already provide
              parsed data after applying the
              first-page identity fallback.

              DOCX continues to use the normal
              parser.
            */
            const data = page.preParsedData || extractTaxFields(page.text);

            console.debug(`[Tax extraction] Page ${page.pageNumber}`, {
              method: page.method,

              data,
            });

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

          const testLimitMessage =
            sourceTotalPages > pages.length
              ? ` • testing first ${pages.length} of ${sourceTotalPages} pages`
              : "";

          if (fileRecordCount === 0) {
            updateFileStatus(
              file,
              "review",
              `No usable tax records found in ${pages.length} processed page${
                pages.length === 1 ? "" : "s"
              }${testLimitMessage}`,
            );
          } else if (fileReviewCount > 0) {
            updateFileStatus(
              file,
              "partial",
              `${fileRecordCount} record${
                fileRecordCount === 1 ? "" : "s"
              } extracted • ${fileReviewCount} page${
                fileReviewCount === 1 ? "" : "s"
              } need review${testLimitMessage}`,
            );
          } else {
            updateFileStatus(
              file,
              "extracted",
              `${fileRecordCount} record${
                fileRecordCount === 1 ? "" : "s"
              } extracted from ${pages.length} page${
                pages.length === 1 ? "" : "s"
              }${testLimitMessage}`,
            );
          }
        } catch (error) {
          const errorMessage =
            error?.message ||
            error?.cause?.message ||
            String(error) ||
            "Unable to read document";

          console.error(`Unable to read ${file.name}`, error);
          console.error("Full extraction error:", {
            name: error?.name,
            message: errorMessage,
            stack: error?.stack,
            cause: error?.cause,
          });

          failedFileCount += 1;

          updateFileStatus(file, "failed", errorMessage);
        }
      }

      setExtractedRecords((current) => {
        const untouched = current.filter(
          (record) => !processedFileKeys.has(record.fileKey),
        );

        const nextRecords = [...untouched, ...extractedNow];

        showExtractedDataInNetwork(nextRecords.map((record) => record.data));

        return nextRecords;
      });
    } finally {
      if (ocrWorker) {
        await ocrWorker.terminate();
      }

      setIsExtracting(false);

      Swal.close();
    }

    /* =====================================================
       FINAL RESULT
       ===================================================== */

    if (extractedNow.length === 0) {
      await Swal.fire({
        icon: "warning",

        title: "No tax records extracted",

        text:
          failedFileCount > 0
            ? `${failedFileCount} file(s) could not be processed. Check the file row or browser console for the exact error.`
            : "The documents were read, but no supported tax information was detected.",

        confirmButtonColor: "#18181b",
      });

      return;
    }

    if (reviewPageCount > 0 || failedFileCount > 0) {
      await Swal.fire({
        icon: "warning",

        title: "Extraction completed",

        html: `
          <div style="font-size:14px;line-height:1.7;color:#52525b;">
            <div>
              <strong>${extractedNow.length}</strong>
              tax record${extractedNow.length === 1 ? "" : "s"} extracted.
            </div>

            <div>
              <strong>${totalPageCount}</strong>
              page${totalPageCount === 1 ? "" : "s"} processed.
            </div>

            ${
              reviewPageCount > 0
                ? `
                  <div>
                    <strong>${reviewPageCount}</strong>
                    page${reviewPageCount === 1 ? "" : "s"} need review.
                  </div>
                `
                : ""
            }

            ${
              failedFileCount > 0
                ? `
                  <div>
                    <strong>${failedFileCount}</strong>
                    file${failedFileCount === 1 ? "" : "s"} failed.
                  </div>
                `
                : ""
            }

            <div style="margin-top:8px;font-size:12px;color:#71717a;">
              Test mode processed page 1 only.
            </div>
          </div>
        `,

        confirmButtonText: "Continue",

        confirmButtonColor: "#18181b",
      });

      return;
    }

    await Swal.fire({
      icon: "success",

      title: "Extraction complete",

      html: `
        <div style="font-size:14px;line-height:1.7;color:#52525b;">
          <div>
            <strong>${extractedNow.length}</strong>
            tax record${
              extractedNow.length === 1 ? "" : "s"
            } extracted successfully.
          </div>

          <div style="margin-top:6px;font-size:12px;color:#71717a;">
            Test mode processed page 1 only.
          </div>
        </div>
      `,

      confirmButtonColor: "#18181b",
    });
  }

  /* =======================================================
     FILE VALIDATION
     ======================================================= */

  function validateAndAddFiles(incomingFiles) {
    const incoming = Array.from(incomingFiles);

    if (!incoming.length) {
      return;
    }

    setFileError("");

    const existingKeys = new Set(files.map(getFileKey));

    const nextFiles = [...files];

    const acceptedFiles = [];

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

      const nextTotalSize =
        nextFiles.reduce(
          (total, selectedFile) => total + selectedFile.size,
          0,
        ) + file.size;

      if (nextTotalSize > MAX_TOTAL_SIZE) {
        rejected.push(`${file.name}: total upload exceeds 100 MB`);

        continue;
      }

      existingKeys.add(getFileKey(file));

      nextFiles.push(file);

      acceptedFiles.push(file);
    }

    setFiles(nextFiles);

    if (acceptedFiles.length > 0) {
      void extractDocuments(acceptedFiles);
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

    setExtractedRecords((current) => {
      const nextRecords = current.filter(
        (record) => record.fileKey !== fileKey,
      );

      showExtractedDataInNetwork(nextRecords.map((record) => record.data));

      return nextRecords;
    });

    setFileError("");
  }

  function clearFiles() {
    setFiles([]);

    setFileStatuses({});

    setExtractedRecords([]);

    setFileError("");

    showExtractedDataInNetwork([]);
  }

  /* =======================================================
     SAVE EXCEL
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

    void Swal.fire({
      title: "Generating Excel",

      text: "Preparing your tax records...",

      allowOutsideClick: false,

      allowEscapeKey: false,

      showConfirmButton: false,

      didOpen: () => {
        Swal.showLoading();
      },
    });

    await new Promise((resolve) => setTimeout(resolve, 100));

    try {
      exportRecordsToExcel(extractedRecords);
    } finally {
      Swal.close();
    }

    Toast.fire({
      icon: "success",

      title: `${extractedRecords.length} record${
        extractedRecords.length === 1 ? "" : "s"
      } saved to one Excel file`,
    });
  }

  /* =======================================================
     JSX
     ======================================================= */

  return (
    <main className="min-h-screen bg-zinc-50 px-4 py-6 sm:px-6 md:py-10">
      <div className="mx-auto max-w-5xl">
        <header className="mb-7">
          <p className="text-sm font-medium text-zinc-500">Tax Management</p>

          <h1 className="mt-1 text-2xl font-semibold tracking-tight text-zinc-950 sm:text-3xl">
            Extract Tax Records
          </h1>

          <p className="mt-2 max-w-2xl text-sm leading-6 text-zinc-600">
            Upload PDF or Word tax documents. During testing, only page 1 is
            processed, while the extractor still cross-checks native text with
            OCR evidence.
          </p>
        </header>

        <section className="overflow-hidden rounded-xl border border-zinc-200 bg-white shadow-sm">
          <div className="border-b border-zinc-200 p-5 sm:p-6 md:p-7">
            <div className="flex flex-col gap-2 sm:flex-row sm:items-start sm:justify-between">
              <div>
                <h2 className="text-base font-semibold text-zinc-950">
                  Upload Tax Documents
                </h2>

                <p className="mt-1 max-w-3xl text-sm leading-5 text-zinc-500">
                  BIR Form 2307 uses targeted multi-pass OCR for dates,
                  Payee/Payor identity, and withholding details, with full-page
                  verification when a page is difficult to read.
                </p>
              </div>

              <span className="w-fit shrink-0 rounded-full bg-amber-50 px-3 py-1 text-xs font-semibold text-amber-700">
                Page 1 test mode
              </span>
            </div>
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
                flex min-h-56 flex-col items-center justify-center rounded-lg
                border-2 border-dashed px-6 py-8 text-center outline-none transition
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
                    Extracting document...
                  </p>

                  <p className="mt-1 max-w-md text-xs leading-5 text-zinc-500">
                    Reading page 1 and cross-checking OCR results.
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
                    <span>Only page 1 is scanned</span>
                  </div>
                </>
              )}
            </div>

            {fileError && (
              <div
                role="alert"
                className="mt-4 rounded-md border border-red-200 bg-red-50 px-4 py-3 text-sm leading-5 text-red-700"
              >
                {fileError}
              </div>
            )}

            {files.length > 0 && (
              <div className="mt-5 overflow-hidden rounded-lg border border-zinc-200">
                <div className="flex flex-col gap-3 border-b border-zinc-200 bg-zinc-50 px-4 py-3 sm:flex-row sm:items-center sm:justify-between">
                  <div>
                    <p className="text-sm font-medium text-zinc-900">
                      Documents
                    </p>

                    <p className="mt-0.5 text-xs text-zinc-500">
                      {files.length} file{files.length === 1 ? "" : "s"} •{" "}
                      {formatFileSize(totalFileSize)} • page 1 only
                    </p>
                  </div>

                  <button
                    type="button"
                    disabled={isExtracting}
                    onClick={clearFiles}
                    className="w-fit text-sm font-medium text-zinc-500 transition hover:text-red-600 disabled:cursor-not-allowed disabled:opacity-40"
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

            {files.length > 0 && (
              <div className="mt-6 rounded-lg border border-zinc-200 bg-zinc-50 p-4">
                <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
                  <div>
                    <p className="text-sm font-semibold text-zinc-900">
                      {extractedRecords.length} extracted record
                      {extractedRecords.length === 1 ? "" : "s"}
                    </p>

                    <p className="mt-1 text-xs leading-5 text-zinc-500">
                      Only page 1 is processed during testing. Low-confidence
                      data is marked for review instead of being silently
                      accepted.
                    </p>
                  </div>

                  <div className="flex flex-col-reverse gap-3 sm:flex-row">
                    <button
                      type="button"
                      onClick={() => extractDocuments(files)}
                      disabled={isExtracting}
                      className="inline-flex h-10 items-center justify-center rounded-md border border-zinc-300 bg-white px-4 text-sm font-medium text-zinc-700 transition hover:bg-zinc-100 disabled:cursor-not-allowed disabled:opacity-50"
                    >
                      Extract Again
                    </button>

                    <button
                      type="button"
                      onClick={handleSaveUploadedExcel}
                      disabled={isExtracting || extractedRecords.length === 0}
                      className="inline-flex h-10 items-center justify-center gap-2 rounded-md bg-zinc-900 px-5 text-sm font-medium text-white transition hover:bg-zinc-800 disabled:cursor-not-allowed disabled:opacity-50"
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
      </div>
    </main>
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
        title="Remove document"
        className="flex h-9 w-9 shrink-0 items-center justify-center rounded-md text-zinc-400 transition hover:bg-red-50 hover:text-red-600 focus:outline-none focus:ring-2 focus:ring-red-200 disabled:cursor-not-allowed disabled:opacity-30"
      >
        <TrashIcon />
      </button>
    </li>
  );
}

/* =========================================================
   STATUS
   ========================================================= */

function FileStatus({ state }) {
  if (!state) {
    return null;
  }

  if (state.status === "reading") {
    return (
      <span className="rounded-full bg-zinc-100 px-2 py-0.5 text-[10px] font-semibold text-zinc-600">
        Reading
      </span>
    );
  }

  if (state.status === "extracted") {
    return (
      <span className="rounded-full bg-emerald-50 px-2 py-0.5 text-[10px] font-semibold text-emerald-700">
        Extracted
      </span>
    );
  }

  if (state.status === "partial") {
    return (
      <span className="rounded-full bg-amber-50 px-2 py-0.5 text-[10px] font-semibold text-amber-700">
        Partial
      </span>
    );
  }

  if (state.status === "review") {
    return (
      <span className="rounded-full bg-amber-50 px-2 py-0.5 text-[10px] font-semibold text-amber-700">
        Review
      </span>
    );
  }

  if (state.status === "failed") {
    return (
      <span className="rounded-full bg-red-50 px-2 py-0.5 text-[10px] font-semibold text-red-700">
        Failed
      </span>
    );
  }

  return null;
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
