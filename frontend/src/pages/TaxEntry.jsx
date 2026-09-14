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
  TESTING ONLY

  Only the first 15 pages / records are analyzed.
*/
const TEST_EXTRACT_LIMIT = 15;

/*
  We no longer OCR the entire BIR 2307 page.

  These positions correspond to the fixed BIR 2307 layout
  you provided.

  x, y, width and height are percentages of the PDF page.

  Some extra margin is intentionally included because
  scanned PDFs may be shifted slightly.
*/
const BIR_2307_REGIONS = {
  tin: {
    x: 0.18,
    y: 0.125,
    width: 0.65,
    height: 0.062,
  },

  name: {
    x: 0.01,
    y: 0.165,
    width: 0.98,
    height: 0.06,
  },

  paymentTable: {
    x: 0.01,
    y: 0.36,
    width: 0.98,
    height: 0.27,
  },
};

/*
  Render at a moderate scale.

  The small cropped areas are enlarged before OCR,
  which is faster than rendering the entire page huge.
*/
const PDF_RENDER_SCALE = 1.5;

const REGION_UPSCALE = 2;

/* =========================================================
   FILE TYPES
   ========================================================= */

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
  totalAmountPayable: "",
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
    title: "Analyzing BIR 2307",
    html: `
      <div style="padding-top:4px;text-align:center;">
        <div
          id="extract-status"
          style="
            font-size:14px;
            font-weight:600;
            color:#18181b;
          "
        >
          Preparing document...
        </div>

        <div
          id="extract-detail"
          style="
            margin-top:8px;
            font-size:12px;
            line-height:1.6;
            color:#71717a;
          "
        >
          Testing first ${TEST_EXTRACT_LIMIT} pages only
        </div>

        <div
          style="
            margin-top:12px;
            font-size:11px;
            color:#a1a1aa;
          "
        >
          Only Payee TIN, Payee Name and Payment Table are scanned.
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

  const statusElement = container.querySelector("#extract-status");

  const detailElement = container.querySelector("#extract-detail");

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
   GENERIC HELPERS
   ========================================================= */

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
   TEXT CLEANING
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

function cleanName(value = "") {
  return String(value)
    .replace(/\s+/g, " ")
    .replace(/^[\s:;,.–—-]+/, "")
    .replace(/[\s:;,.–—-]+$/, "")
    .trim();
}

/* =========================================================
   TIN
   ========================================================= */

function formatTin(value = "") {
  const digits = String(value).replace(/\D/g, "").slice(0, 14);

  if (digits.length <= 3) {
    return digits;
  }

  if (digits.length <= 6) {
    return `${digits.slice(0, 3)}-${digits.slice(3)}`;
  }

  if (digits.length <= 9) {
    return `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6)}`;
  }

  return `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(
    6,
    9,
  )}-${digits.slice(9)}`;
}

function validateTin(value) {
  const length = String(value).replace(/\D/g, "").length;

  if (length === 9 || length === 12 || length === 14) {
    return true;
  }

  return "TIN must contain 9, 12, or 14 digits.";
}

/*
  Specifically designed for the BIR 2307 TIN row.

  Supports:

  297 - 944 - 186 - 00000
  297-944-186-00000
  297 944 186 00000
  29794418600000
*/
function parseTinFromText(text) {
  const normalized = normalizeText(text);

  const grouped = normalized.match(
    /\b(\d{3})\s*[- ]\s*(\d{3})\s*[- ]\s*(\d{3})\s*[- ]\s*(\d{3,5})\b/,
  );

  if (grouped) {
    return formatTin(`${grouped[1]}${grouped[2]}${grouped[3]}${grouped[4]}`);
  }

  const compact = normalized.match(/\b(\d{12,14})\b/);

  if (compact?.[1]) {
    return formatTin(compact[1]);
  }

  /*
    OCR fallback.

    Ignore single digit row labels such as "2".
  */
  const numericParts =
    normalized.match(/\d+/g)?.filter((part) => part.length >= 2) || [];

  for (let index = 0; index < numericParts.length; index += 1) {
    const candidate = numericParts.slice(index, index + 4).join("");

    if (candidate.length === 12 || candidate.length === 14) {
      return formatTin(candidate);
    }
  }

  return "";
}

/* =========================================================
   PAYEE NAME
   ========================================================= */

function parsePayeeNameFromText(text) {
  const raw = cleanText(text);

  if (!raw) {
    return "";
  }

  /*
    First try a direct labeled extraction.
  */
  const normalized = normalizeText(raw);

  const directPatterns = [
    /Payee(?:'s|’s)?\s+Name(?:\s*\([^)]*\))?\s*[:\-]?\s*([A-Z][A-Z ,.'&’/-]{4,120}?)(?=\s+(?:Registered|Address|TIN|Taxpayer|Foreign)\b|$)/i,

    /Payee\s+Name\s*[:\-]?\s*([A-Z][A-Z ,.'&’/-]{4,120}?)(?=\s+(?:Registered|Address|TIN|Taxpayer|Foreign)\b|$)/i,
  ];

  for (const pattern of directPatterns) {
    const match = normalized.match(pattern);

    if (match?.[1] && match[1].trim()) {
      return cleanName(match[1]);
    }
  }

  /*
    The crop usually has only:
      - the row label
      - the actual Payee name

    Remove all known header/label text and choose
    the strongest remaining text line.
  */
  const ignored =
    /PAYEE|NAME|LAST NAME|FIRST NAME|MIDDLE NAME|INDIVIDUAL|REGISTERED NAME|NON-INDIVIDUAL|REGISTERED ADDRESS|FOREIGN ADDRESS|TAXPAYER|IDENTIFICATION|TIN|ZIP CODE/i;

  const candidates = raw
    .split("\n")
    .map((line) => cleanName(line))
    .filter(Boolean)
    .filter((line) => !ignored.test(line))
    .filter((line) => {
      const letters = (line.match(/[A-Za-z]/g) || []).length;

      return letters >= 5 && line.length <= 120;
    });

  if (!candidates.length) {
    return "";
  }

  /*
    Prefer a name written mostly in capital letters.
  */
  const scored = candidates.map((line) => {
    const letters = line.replace(/[^A-Za-z]/g, "");

    const uppercase = letters.replace(/[^A-Z]/g, "").length;

    const uppercaseRatio = letters.length ? uppercase / letters.length : 0;

    return {
      line,

      score: uppercaseRatio * 100 + letters.length,
    };
  });

  scored.sort((a, b) => b.score - a.score);

  return scored[0].line;
}

/* =========================================================
   MONEY
   ========================================================= */

function parseAmount(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const normalized = String(value)
    .replace(/₱/g, "")
    .replace(/PHP/gi, "")
    .replace(/,/g, "")
    .replace(/[^\d.-]/g, "");

  const number = Number(normalized);

  return Number.isFinite(number) ? number : null;
}

function getMoneyValues(text) {
  const matches =
    String(text).match(/(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}/g) || [];

  return matches.map(parseAmount).filter((value) => value !== null);
}

/* =========================================================
   2307 NATURE OF PAYMENT
   ========================================================= */

function parseNatureOfPayment(text) {
  const normalized = normalizeText(text);

  const patterns = [
    /Gross\s+Income\s+(?:is\s+)?Less\s+Than\s+3M(?:\s+or)?(?:\s+Non[\s-]*VAT\s+registered\s+regardless\s+of\s+amount)?/i,

    /Non[\s-]*VAT\s+registered\s+regardless\s+of\s+amount/i,

    /Gross\s+Income\s+(?:is\s+)?3M\s+and\s+Above/i,

    /Professional\s+Fees?/i,

    /Rental/i,

    /Commission/i,
  ];

  for (const pattern of patterns) {
    const match = normalized.match(pattern);

    if (match?.[0]) {
      return cleanName(match[0]);
    }
  }

  return "";
}

/* =========================================================
   2307 ATC
   ========================================================= */

function parseAtcFromLine(line) {
  if (!line) {
    return "";
  }

  /*
    Only search before the first monetary amount.

    This prevents 700 from 5,700.00
    being mistaken as the ATC.
  */
  const firstAmount = line.search(/(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}/);

  const prefix = firstAmount >= 0 ? line.slice(0, firstAmount) : line;

  /*
    Examples OCR may return:

      WI151
      W I 151
      I151
      151
  */
  const candidates = prefix.match(/\b(?:W\s*)?[A-Z]{0,2}\s*\d{3}\b/gi) || [];

  if (!candidates.length) {
    return "";
  }

  const value = candidates[candidates.length - 1]
    .replace(/\s+/g, "")
    .toUpperCase();

  return value;
}

/* =========================================================
   PAYMENT TABLE PARSER
   ========================================================= */

function parse2307PaymentTable(text) {
  const cleaned = cleanText(text);

  const lines = cleaned
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  let dataLine = "";

  let dataLineIndex = -1;

  /*
    Find the first actual row containing monetary values.

    We prefer a row containing the expected nature text,
    otherwise use the first line with at least two amounts.
  */
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];

    const amounts = getMoneyValues(line);

    if (
      amounts.length >= 2 &&
      /Gross|Income|VAT|Professional|Rental|Commission/i.test(line)
    ) {
      dataLine = line;
      dataLineIndex = index;

      break;
    }
  }

  if (!dataLine) {
    for (let index = 0; index < lines.length; index += 1) {
      const amounts = getMoneyValues(lines[index]);

      if (amounts.length >= 2) {
        dataLine = lines[index];

        dataLineIndex = index;

        break;
      }
    }
  }

  let nature = parseNatureOfPayment(cleaned);

  /*
    If the "Non VAT..." part wrapped onto the next row,
    combine it.
  */
  if (dataLineIndex >= 0 && lines[dataLineIndex + 1]) {
    const nextLine = lines[dataLineIndex + 1];

    if (/Non[\s-]*VAT|registered regardless/i.test(nextLine)) {
      nature = parseNatureOfPayment(`${dataLine} ${nextLine}`) || nature;
    }
  }

  const atc = parseAtcFromLine(dataLine);

  const rowAmounts = getMoneyValues(dataLine);

  let firstMonth = null;
  let secondMonth = null;
  let thirdMonth = null;
  let total = null;
  let taxWithheld = null;

  /*
    BIR 2307 table:

      Month 1
      Month 2
      Month 3
      Total
      Tax Withheld

    Depending on blank cells, OCR may only return
    the populated numbers.

    Example from your sample:

      5,700.00   5,700.00   570.00

    becomes:

      Month 1  = 5,700
      Total    = 5,700
      Withheld = 570
  */

  if (rowAmounts.length >= 5) {
    const lastFive = rowAmounts.slice(-5);

    [firstMonth, secondMonth, thirdMonth, total, taxWithheld] = lastFive;
  } else if (rowAmounts.length === 4) {
    firstMonth = rowAmounts[0];

    secondMonth = rowAmounts[1];

    total = rowAmounts[2];

    taxWithheld = rowAmounts[3];
  } else if (rowAmounts.length === 3) {
    firstMonth = rowAmounts[0];

    total = rowAmounts[1];

    taxWithheld = rowAmounts[2];
  } else if (rowAmounts.length === 2) {
    total = rowAmounts[0];

    taxWithheld = rowAmounts[1];
  }

  /*
    Look for the Total row as a fallback.
  */
  const totalLine = lines.find(
    (line) => /^Total\b/i.test(line) && getMoneyValues(line).length >= 1,
  );

  if (totalLine) {
    const values = getMoneyValues(totalLine);

    if (values.length >= 2) {
      total = total ?? values[values.length - 2];

      taxWithheld = taxWithheld ?? values[values.length - 1];
    }
  }

  /*
    Final fallback based on all values in the crop.
  */
  const allAmounts = getMoneyValues(cleaned);

  const positive = allAmounts.filter((value) => value > 0);

  if (total === null && positive.length) {
    total = Math.max(...positive);
  }

  if (taxWithheld === null && total !== null) {
    const smaller = positive
      .filter((value) => value < total)
      .sort((a, b) => a - b);

    if (smaller.length) {
      taxWithheld = smaller[0];
    }
  }

  return {
    nature_of_income_payment: nature,

    atc,

    first_month: firstMonth,

    second_month: secondMonth,

    third_month: thirdMonth,

    total_income_payment: total,

    tax_withheld: taxWithheld,
  };
}

/* =========================================================
   BIR 2307 DATA
   ========================================================= */

function create2307Data({
  nativeText = "",
  tinText = "",
  nameText = "",
  tableText = "",
}) {
  const tin = parseTinFromText(tinText) || parseTinFromText(nativeText);

  const name =
    parsePayeeNameFromText(nameText) || parsePayeeNameFromText(nativeText);

  const tableFromCrop = parse2307PaymentTable(tableText);

  const tableFromNative = parse2307PaymentTable(nativeText);

  return {
    form_type: "BIR Form 2307",

    taxpayer: {
      name,

      tin,
    },

    withholding: {
      nature_of_income_payment:
        tableFromCrop.nature_of_income_payment ||
        tableFromNative.nature_of_income_payment,

      atc: tableFromCrop.atc || tableFromNative.atc,

      first_month: tableFromCrop.first_month ?? tableFromNative.first_month,

      second_month: tableFromCrop.second_month ?? tableFromNative.second_month,

      third_month: tableFromCrop.third_month ?? tableFromNative.third_month,

      total_income_payment:
        tableFromCrop.total_income_payment ??
        tableFromNative.total_income_payment,

      tax_withheld: tableFromCrop.tax_withheld ?? tableFromNative.tax_withheld,
    },
  };
}

function hasValue(value) {
  return !(value === "" || value === null || value === undefined);
}

function hasAny2307Data(data) {
  return Boolean(
    data.taxpayer?.name ||
    data.taxpayer?.tin ||
    data.withholding?.nature_of_income_payment ||
    data.withholding?.atc ||
    hasValue(data.withholding?.first_month) ||
    hasValue(data.withholding?.total_income_payment) ||
    hasValue(data.withholding?.tax_withheld),
  );
}

function isComplete2307(data) {
  return Boolean(
    data.taxpayer?.name &&
    data.taxpayer?.tin &&
    hasValue(data.withholding?.total_income_payment) &&
    hasValue(data.withholding?.tax_withheld),
  );
}

function tableNeedsOcr(data) {
  return !(
    hasValue(data.withholding?.total_income_payment) &&
    hasValue(data.withholding?.tax_withheld)
  );
}

/* =========================================================
   PDF TEXT LAYER
   ========================================================= */

async function extractNativePdfText(page) {
  const textContent = await page.getTextContent();

  const items = textContent.items
    .filter((item) => item.str && item.str.trim())
    .map((item) => ({
      text: item.str.trim(),

      x: item.transform?.[4] ?? 0,

      y: item.transform?.[5] ?? 0,
    }));

  const lines = [];

  const yTolerance = 3;

  const sorted = [...items].sort((a, b) => {
    const yDifference = b.y - a.y;

    if (Math.abs(yDifference) > yTolerance) {
      return yDifference;
    }

    return a.x - b.x;
  });

  sorted.forEach((item) => {
    let line = lines.find((entry) => Math.abs(entry.y - item.y) <= yTolerance);

    if (!line) {
      line = {
        y: item.y,
        items: [],
      };

      lines.push(line);
    }

    line.items.push(item);
  });

  return lines
    .sort((a, b) => b.y - a.y)
    .map((line) =>
      line.items
        .sort((a, b) => a.x - b.x)
        .map((item) => item.text)
        .join(" "),
    )
    .join("\n");
}

/* =========================================================
   PDF CANVAS
   ========================================================= */

async function renderPageCanvas(page) {
  const viewport = page.getViewport({
    scale: PDF_RENDER_SCALE,
  });

  const canvas = document.createElement("canvas");

  canvas.width = Math.ceil(viewport.width);

  canvas.height = Math.ceil(viewport.height);

  const context = canvas.getContext("2d", {
    willReadFrequently: true,
  });

  if (!context) {
    throw new Error("Unable to create PDF canvas.");
  }

  /*
    White background helps OCR
    on scanned PDF pages.
  */
  context.fillStyle = "#ffffff";

  context.fillRect(0, 0, canvas.width, canvas.height);

  await page.render({
    canvasContext: context,

    viewport,
  }).promise;

  return canvas;
}

function cropCanvas(sourceCanvas, region) {
  const sourceX = Math.floor(sourceCanvas.width * region.x);

  const sourceY = Math.floor(sourceCanvas.height * region.y);

  const sourceWidth = Math.floor(sourceCanvas.width * region.width);

  const sourceHeight = Math.floor(sourceCanvas.height * region.height);

  const canvas = document.createElement("canvas");

  canvas.width = Math.max(1, Math.floor(sourceWidth * REGION_UPSCALE));

  canvas.height = Math.max(1, Math.floor(sourceHeight * REGION_UPSCALE));

  const context = canvas.getContext("2d", {
    willReadFrequently: true,
  });

  if (!context) {
    throw new Error("Unable to create cropped OCR canvas.");
  }

  context.fillStyle = "#ffffff";

  context.fillRect(0, 0, canvas.width, canvas.height);

  context.imageSmoothingEnabled = true;

  context.drawImage(
    sourceCanvas,

    sourceX,
    sourceY,
    sourceWidth,
    sourceHeight,

    0,
    0,
    canvas.width,
    canvas.height,
  );

  return canvas;
}

/* =========================================================
   TARGETED OCR
   ========================================================= */

async function recognizeRegion(worker, canvas, pageSegMode = "6") {
  /*
    PSM 6 = assume one block of text.

    We intentionally do not use a hard character whitelist
    because that sometimes removes letters from names.
  */
  if (typeof worker.setParameters === "function") {
    try {
      await worker.setParameters({
        tessedit_pageseg_mode: pageSegMode,

        preserve_interword_spaces: "1",
      });
    } catch (error) {
      console.warn("OCR parameter warning:", error);
    }
  }

  const result = await worker.recognize(canvas);

  return cleanText(result.data.text || "");
}

/* =========================================================
   PDF 2307 EXTRACTION
   ========================================================= */

async function extractPdf2307Records(
  file,
  { getOcrWorker, remainingLimit, onProgress },
) {
  const arrayBuffer = await file.arrayBuffer();

  const loadingTask = pdfjsLib.getDocument({
    data: arrayBuffer,
  });

  const pdf = await loadingTask.promise;

  const records = [];

  const pagesToRead = Math.min(
    pdf.numPages,
    remainingLimit,
    TEST_EXTRACT_LIMIT,
  );

  try {
    for (let pageNumber = 1; pageNumber <= pagesToRead; pageNumber += 1) {
      onProgress?.(
        `Reading page ${pageNumber} of ${pagesToRead}`,
        "Checking PDF text first...",
      );

      const page = await pdf.getPage(pageNumber);

      try {
        /*
          STEP 1

          Fast native PDF text.
        */
        let nativeText = "";

        try {
          nativeText = await extractNativePdfText(page);
        } catch (error) {
          console.warn(
            `Unable to read PDF text layer on page ${pageNumber}`,
            error,
          );
        }

        let data = create2307Data({
          nativeText,
        });

        /*
          STEP 2

          Only render the page if one of the
          three important areas needs OCR.
        */

        const needsTin = !data.taxpayer?.tin;

        const needsName = !data.taxpayer?.name;

        const needsTable = tableNeedsOcr(data);

        let tinText = "";
        let nameText = "";
        let tableText = "";

        if (needsTin || needsName || needsTable) {
          onProgress?.(
            `Scanning page ${pageNumber} of ${pagesToRead}`,
            "Only required BIR 2307 areas are being scanned.",
          );

          const fullCanvas = await renderPageCanvas(page);

          try {
            const worker = await getOcrWorker();

            /*
              PAYEE TIN
            */
            if (needsTin) {
              onProgress?.(
                `Page ${pageNumber}: Payee TIN`,
                "Scanning TIN area...",
              );

              const tinCanvas = cropCanvas(fullCanvas, BIR_2307_REGIONS.tin);

              try {
                tinText = await recognizeRegion(worker, tinCanvas, "7");
              } finally {
                tinCanvas.width = 0;

                tinCanvas.height = 0;
              }
            }

            /*
              PAYEE NAME
            */
            if (needsName) {
              onProgress?.(
                `Page ${pageNumber}: Payee Name`,
                "Scanning name area...",
              );

              const nameCanvas = cropCanvas(fullCanvas, BIR_2307_REGIONS.name);

              try {
                nameText = await recognizeRegion(worker, nameCanvas, "6");
              } finally {
                nameCanvas.width = 0;

                nameCanvas.height = 0;
              }
            }

            /*
              PAYMENT TABLE
            */
            if (needsTable) {
              onProgress?.(
                `Page ${pageNumber}: Payment Table`,
                "Scanning Part III amounts...",
              );

              const tableCanvas = cropCanvas(
                fullCanvas,
                BIR_2307_REGIONS.paymentTable,
              );

              try {
                tableText = await recognizeRegion(worker, tableCanvas, "6");
              } finally {
                tableCanvas.width = 0;

                tableCanvas.height = 0;
              }
            }
          } finally {
            fullCanvas.width = 0;

            fullCanvas.height = 0;
          }

          /*
            Merge the cropped OCR with any native
            PDF text we already found.
          */
          data = create2307Data({
            nativeText,
            tinText,
            nameText,
            tableText,
          });
        }

        /*
          IMPORTANT:

          Do not throw away a page just because
          one field is missing.

          This prevents "some data is gone".

          Complete and partial records are both kept.
        */
        if (hasAny2307Data(data)) {
          records.push({
            pageNumber,

            data,

            status: isComplete2307(data) ? "complete" : "partial",
          });
        }
      } finally {
        if (typeof page.cleanup === "function") {
          try {
            page.cleanup();
          } catch (error) {
            console.warn("Page cleanup warning:", error);
          }
        }
      }
    }
  } finally {
    /*
      Do NOT use:

      pdf.destroy()

      This fixes your previous:
      "pdf.destroy is not a function"
      error.
    */

    if (typeof pdf.cleanup === "function") {
      try {
        pdf.cleanup();
      } catch (error) {
        console.warn("PDF cleanup warning:", error);
      }
    }

    if (typeof loadingTask.destroy === "function") {
      try {
        await loadingTask.destroy();
      } catch (error) {
        console.warn("PDF loading-task cleanup warning:", error);
      }
    }
  }

  return records;
}

/* =========================================================
   DOCX EXTRACTION
   ========================================================= */

function splitDocx2307Records(text) {
  const cleaned = cleanText(text);

  if (!cleaned) {
    return [];
  }

  const matches = [];

  const regex = /\b(?:BIR\s+Form\s*)?2307\b/gi;

  let match;

  while ((match = regex.exec(cleaned)) !== null) {
    const previous = matches[matches.length - 1];

    if (previous === undefined || match.index - previous > 500) {
      matches.push(match.index);
    }
  }

  if (matches.length <= 1) {
    return [cleaned];
  }

  const records = [];

  for (
    let index = 0;
    index < matches.length && records.length < TEST_EXTRACT_LIMIT;
    index += 1
  ) {
    const start = matches[index];

    const end =
      index + 1 < matches.length ? matches[index + 1] : cleaned.length;

    const section = cleanText(cleaned.slice(start, end));

    if (section) {
      records.push(section);
    }
  }

  return records;
}

async function extractDocx2307Records(file, { remainingLimit, onProgress }) {
  onProgress?.("Reading Word document", "Extracting BIR 2307 records...");

  const arrayBuffer = await file.arrayBuffer();

  const result = await mammoth.extractRawText({
    arrayBuffer,
  });

  const sections = splitDocx2307Records(result.value || "").slice(
    0,
    remainingLimit,
  );

  return sections
    .map((section, index) => {
      const data = create2307Data({
        nativeText: section,
      });

      if (!hasAny2307Data(data)) {
        return null;
      }

      return {
        pageNumber: index + 1,

        data,

        status: isComplete2307(data) ? "complete" : "partial",
      };
    })
    .filter(Boolean);
}

/* =========================================================
   DOCUMENT DISPATCHER
   ========================================================= */

async function extractDocumentRecords(file, helpers) {
  const extension = getFileExtension(file.name);

  if (extension === ".pdf") {
    return extractPdf2307Records(file, helpers);
  }

  if (extension === ".docx") {
    return extractDocx2307Records(file, helpers);
  }

  throw new Error("Unsupported document type.");
}

/* =========================================================
   DEBUG NETWORK
   ========================================================= */

function showExtractedDataInNetwork(records) {
  fetch("/__debug/extracted-tax-data", {
    method: "POST",

    headers: {
      "Content-Type": "application/json",
    },

    body: JSON.stringify(records),
  }).catch(() => {
    /*
      404 is expected if there is
      no debug backend route.
    */
  });
}

/* =========================================================
   EXCEL
   ========================================================= */

function getExcelRows(records) {
  return records.map((record) => {
    const data = record.data;

    return {
      "Form Type": data.form_type || "",

      "Payee Name": data.taxpayer?.name || "",

      "Payee TIN": data.taxpayer?.tin || "",

      "Nature of Income Payment":
        data.withholding?.nature_of_income_payment || "",

      ATC: data.withholding?.atc || "",

      "1st Month of Quarter": data.withholding?.first_month ?? "",

      "2nd Month of Quarter": data.withholding?.second_month ?? "",

      "3rd Month of Quarter": data.withholding?.third_month ?? "",

      "Total Income Payment": data.withholding?.total_income_payment ?? "",

      "Tax Withheld for Quarter": data.withholding?.tax_withheld ?? "",
    };
  });
}

function exportRecordsToExcel(records) {
  const rows = getExcelRows(records);

  const worksheet = XLSX.utils.json_to_sheet(rows);

  worksheet["!cols"] = [
    { wch: 18 },
    { wch: 34 },
    { wch: 24 },
    { wch: 45 },
    { wch: 14 },
    { wch: 20 },
    { wch: 20 },
    { wch: 20 },
    { wch: 22 },
    { wch: 24 },
  ];

  /*
    Money columns:

    F = Month 1
    G = Month 2
    H = Month 3
    I = Total
    J = Tax Withheld
  */
  for (let row = 2; row <= rows.length + 1; row += 1) {
    ["F", "G", "H", "I", "J"].forEach((column) => {
      const cell = worksheet[`${column}${row}`];

      if (cell && typeof cell.v === "number") {
        cell.z = "#,##0.00";
      }
    });
  }

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, worksheet, "BIR 2307");

  const now = new Date();

  const date = [
    now.getFullYear(),

    String(now.getMonth() + 1).padStart(2, "0"),

    String(now.getDate()).padStart(2, "0"),
  ].join("-");

  XLSX.writeFile(workbook, `bir-2307-test-${date}.xlsx`);
}

/* =========================================================
   MANUAL VALIDATION
   ========================================================= */

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
     EXTRACT
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
          "Starting OCR",
          "OCR is loaded only when PDF text is incomplete.",
        );

        ocrWorker = await createWorker("eng");
      }

      return ocrWorker;
    }

    const extractedNow = [];

    const processedKeys = new Set(documents.map(getFileKey));

    let completeCount = 0;

    let partialCount = 0;

    let failedCount = 0;

    try {
      for (
        let fileIndex = 0;
        fileIndex < documents.length &&
        extractedNow.length < TEST_EXTRACT_LIMIT;
        fileIndex += 1
      ) {
        const file = documents[fileIndex];

        const fileKey = getFileKey(file);

        try {
          updateFileStatus(file, "reading", "Analyzing BIR 2307");

          const remaining = TEST_EXTRACT_LIMIT - extractedNow.length;

          const records = await extractDocumentRecords(file, {
            getOcrWorker,

            remainingLimit: remaining,

            onProgress: (status, detail) => {
              updateFileStatus(file, "reading", status);

              updateExtractionLoading(status, `${file.name} • ${detail || ""}`);
            },
          });

          let fileComplete = 0;

          let filePartial = 0;

          for (const record of records) {
            if (extractedNow.length >= TEST_EXTRACT_LIMIT) {
              break;
            }

            extractedNow.push({
              fileKey,

              pageNumber: record.pageNumber,

              data: record.data,

              status: record.status,
            });

            if (record.status === "complete") {
              completeCount += 1;

              fileComplete += 1;
            } else {
              partialCount += 1;

              filePartial += 1;
            }

            updateExtractionLoading(
              `${extractedNow.length} of ${TEST_EXTRACT_LIMIT} test records`,
              `${file.name} • Page ${record.pageNumber}`,
            );
          }

          if (fileComplete > 0 && filePartial === 0) {
            updateFileStatus(
              file,
              "extracted",
              `${fileComplete} complete record${fileComplete === 1 ? "" : "s"}`,
            );
          } else if (fileComplete > 0 || filePartial > 0) {
            updateFileStatus(
              file,
              "partial",
              `${fileComplete} complete • ${filePartial} partial`,
            );
          } else {
            updateFileStatus(file, "review", "No target 2307 data detected");
          }
        } catch (error) {
          console.error(`Unable to analyze ${file.name}`, error);

          failedCount += 1;

          updateFileStatus(
            file,
            "failed",
            error.message || "Unable to analyze document",
          );
        }
      }

      setExtractedRecords((current) => {
        const untouched = current.filter(
          (record) => !processedKeys.has(record.fileKey),
        );

        const next = [...untouched, ...extractedNow].slice(
          0,
          TEST_EXTRACT_LIMIT,
        );

        /*
            Only extracted data is sent.
          */
        showExtractedDataInNetwork(next.map((record) => record.data));

        return next;
      });
    } finally {
      if (ocrWorker) {
        try {
          await ocrWorker.terminate();
        } catch (error) {
          console.warn("OCR cleanup warning:", error);
        }
      }

      setIsExtracting(false);

      Swal.close();
    }

    if (!extractedNow.length) {
      await Swal.fire({
        icon: "warning",

        title: "No BIR 2307 data found",

        text: "The target areas were scanned but no usable Payee Name, TIN, or payment data was detected.",

        confirmButtonColor: "#18181b",
      });

      return;
    }

    await Swal.fire({
      icon: partialCount || failedCount ? "warning" : "success",

      title: "Extraction completed",

      html: `
        <div style="
          font-size:14px;
          line-height:1.8;
          color:#52525b;
        ">
          <div>
            <strong>${extractedNow.length}</strong>
            record${extractedNow.length === 1 ? "" : "s"} kept
          </div>

          <div>
            <strong>${completeCount}</strong>
            complete
          </div>

          <div>
            <strong>${partialCount}</strong>
            partial
          </div>

          ${
            failedCount
              ? `
                <div>
                  <strong>${failedCount}</strong>
                  file failure${failedCount === 1 ? "" : "s"}
                </div>
              `
              : ""
          }

          <div style="
            margin-top:8px;
            font-size:12px;
            color:#71717a;
          ">
            Partial rows are preserved instead of being deleted.
          </div>
        </div>
      `,

      confirmButtonColor: "#18181b",
    });
  }

  /* =======================================================
     ADD FILES
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

      const nextTotal =
        nextFiles.reduce(
          (total, selectedFile) => total + selectedFile.size,
          0,
        ) + file.size;

      if (nextTotal > MAX_TOTAL_SIZE) {
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

    setExtractedRecords((current) => {
      const next = current.filter((record) => record.fileKey !== fileKey);

      showExtractedDataInNetwork(next.map((record) => record.data));

      return next;
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
      title: "Generate BIR 2307 Excel?",

      text: `${extractedRecords.length} extracted record${
        extractedRecords.length === 1 ? "" : "s"
      } will be exported.`,

      icon: "question",

      showCancelButton: true,

      reverseButtons: true,

      confirmButtonText: "Generate Excel",

      cancelButtonText: "Cancel",

      confirmButtonColor: "#18181b",

      cancelButtonColor: "#71717a",
    });

    if (!confirmation.isConfirmed) {
      return;
    }

    void Swal.fire({
      title: "Generating Excel",

      text: "Preparing extracted BIR 2307 records...",

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

      title: `${extractedRecords.length} records saved`,
    });
  }

  /* =======================================================
     MANUAL
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

      data: {
        form_type: data.formType,

        taxpayer: {
          name: data.taxpayerName.trim().replace(/\s+/g, " "),

          tin: data.tin,
        },

        withholding: {
          nature_of_income_payment: "",

          atc: "",

          first_month: Number(data.grossSales),

          second_month: null,

          third_month: null,

          total_income_payment: Number(data.grossSales),

          tax_withheld: Number(data.taxDue),
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

    if (!confirmation.isConfirmed) {
      return;
    }

    reset(DEFAULT_VALUES);

    Toast.fire({
      icon: "success",

      title: "Form cleared",
    });
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
            Enter a record manually or extract BIR Form 2307 records from PDF
            and Word documents.
          </p>
        </header>

        {/* TEST NOTICE */}

        <div className="mb-5 rounded-lg border border-amber-200 bg-amber-50 px-4 py-3">
          <p className="text-sm font-semibold text-amber-900">
            BIR 2307 Test Mode
          </p>

          <p className="mt-1 text-xs leading-5 text-amber-700">
            Only the first {TEST_EXTRACT_LIMIT} pages are analyzed. OCR only
            scans Payee TIN, Payee Name and the Part III Payment Table.
          </p>
        </div>

        {/* MODE */}

        <div className="mb-6 grid gap-3 sm:grid-cols-2">
          <ModeButton
            active={entryMode === "manual"}
            title="Enter Manually"
            description="Enter one tax record manually."
            onClick={() => setEntryMode("manual")}
          />

          <ModeButton
            active={entryMode === "upload"}
            title="Upload BIR 2307"
            description="Targeted extraction for the BIR 2307 layout."
            onClick={() => setEntryMode("upload")}
          />
        </div>

        {uploadMode ? (
          /* =================================================
             UPLOAD
             ================================================= */

          <section className="overflow-hidden rounded-xl border border-zinc-200 bg-white shadow-sm">
            <div className="border-b border-zinc-200 p-5 sm:p-6 md:p-7">
              <h2 className="text-base font-semibold text-zinc-950">
                Upload BIR 2307 Documents
              </h2>

              <p className="mt-1 text-sm leading-5 text-zinc-500">
                Each PDF page is treated as one possible 2307 record. Partial
                records are preserved instead of discarded.
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

              {/* DROP ZONE */}

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
                      Analyzing BIR 2307...
                    </p>

                    <p className="mt-1 max-w-md text-xs leading-5 text-zinc-500">
                      Only the required areas are being scanned.
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

                      <span>First {TEST_EXTRACT_LIMIT} pages</span>
                    </div>
                  </>
                )}
              </div>

              {/* ERROR */}

              {fileError && (
                <div
                  role="alert"
                  className="mt-4 rounded-md border border-red-200 bg-red-50 px-4 py-3 text-sm leading-5 text-red-700"
                >
                  {fileError}
                </div>
              )}

              {/* FILES */}

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
                        Missing fields no longer cause the entire row to
                        disappear.
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
        ) : (
          /* =================================================
             MANUAL
             ================================================= */

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
                    autoComplete="name"
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
                  hint="Supports 9, 12 and 14 digit TINs."
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
                        autoComplete="off"
                        maxLength={18}
                        placeholder="297-944-186-00000"
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

                <CurrencyField
                  name="totalAmountPayable"
                  label="Total Amount Payable"
                  control={control}
                  error={errors.totalAmountPayable?.message}
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
                  className="inline-flex h-10 items-center justify-center rounded-md border border-zinc-300 bg-white px-4 text-sm font-medium text-zinc-700 transition hover:bg-zinc-100 disabled:cursor-not-allowed disabled:opacity-50"
                >
                  Clear Form
                </button>

                <button
                  type="submit"
                  disabled={isSubmitting}
                  className="inline-flex h-10 items-center justify-center gap-2 rounded-md bg-zinc-900 px-5 text-sm font-medium text-white transition hover:bg-zinc-800 disabled:cursor-not-allowed disabled:opacity-60"
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
   CURRENCY
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
      <span className="rounded-full bg-blue-50 px-2 py-0.5 text-[10px] font-semibold text-blue-700">
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
