/**
 * POS File Checker — script.js
 *
 * Pure vanilla JavaScript. No frameworks, no build tools.
 * All business logic ported directly from the VBA macros in
 * POSFileChecker_v1.xlsm. Runs entirely in the browser;
 * files are never sent to any server.
 *
 * Sections:
 *   1.  Line item definitions   (mirrors Column A+B of spreadsheet)
 *   2.  File decoder            (UTF-8 → windows-1252 → iso-8859-1 fallback)
 *   3.  File parser             (.001 fixed-width format → structured data)
 *   4.  Calculator              (11 admin-expected values, VBA formulas)
 *   5.  Validator               (orchestrates pipeline, builds result object)
 *   6.  Formatters              (currency, date display)
 *   7.  PDF export              (via jsPDF + autoTable from CDN)
 *   8.  Excel export            (via SheetJS from CDN)
 *   9.  UI rendering            (drop zone, table, summary, filters)
 *   10. Event wiring            (drag-drop, file input, reset, export)
 */

'use strict';

/* ============================================================
   1. LINE ITEM DEFINITIONS
   Static metadata for all 65 .001 file positions.
   decodeMode: 'currency' → raw int ÷ 100 = PHP amount
               'integer'  → plain integer, no division
               'text'     → raw string
   isValidated: true → Column E has Pass/Failed check
   section: 'vat' (lines 1-34) | 'nonvat' (lines 35-65)
   ============================================================ */
const LINE_ITEMS = [
  // ── Section A: VAT Sales (lines 1–34) ──────────────────────────────────
  { lineItem:  1, definition: 'Tenant Code',                              decodeMode: 'text',     isValidated: false, section: 'vat' },
  { lineItem:  2, definition: 'POS Terminal Number',                      decodeMode: 'text',     isValidated: false, section: 'vat' },
  { lineItem:  3, definition: 'Date (mmddyyyy)',                          decodeMode: 'text',     isValidated: false, section: 'vat' },
  { lineItem:  4, definition: 'Old Accumulated Sales',                    decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem:  5, definition: 'New Accumulated Sales',                    decodeMode: 'currency', isValidated: true,  section: 'vat' },
  { lineItem:  6, definition: 'Total Gross Amount',                       decodeMode: 'currency', isValidated: true,  section: 'vat' },
  { lineItem:  7, definition: 'Total Deductions',                         decodeMode: 'currency', isValidated: true,  section: 'vat' },
  { lineItem:  8, definition: 'Total Promo Sales Amount',                 decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem:  9, definition: 'Total Discount',                           decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 10, definition: 'Total Refund Amount',                      decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 11, definition: 'Total Returned Items Amount',              decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 12, definition: 'Total Other Taxes',                        decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 13, definition: 'Total Service Charge Amount',              decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 14, definition: 'Total Adjustment Discount',                decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 15, definition: 'Total Void Amount',                        decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 16, definition: 'Total Discount Cards',                     decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 17, definition: 'Total Delivery Charges',                   decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 18, definition: 'Total Gift Certificates/Checks Redeemed', decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 19, definition: 'Store Specific Discount 1 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 20, definition: 'Store Specific Discount 2 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 21, definition: 'Store Specific Discount 3 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 22, definition: 'Store Specific Discount 4 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 23, definition: 'Store Specific Discount 5 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 24, definition: 'Total of all Non-Approved Store Discounts',decodeMode: 'currency', isValidated: true,  section: 'vat' },
  { lineItem: 25, definition: 'Store Specific Discount 1 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 26, definition: 'Store Specific Discount 2 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 27, definition: 'Store Specific Discount 3 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 28, definition: 'Store Specific Discount 4 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 29, definition: 'Store Specific Discount 5 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 30, definition: 'Total VAT/Tax Amount',                     decodeMode: 'currency', isValidated: true,  section: 'vat' },
  { lineItem: 31, definition: 'Total Net Sales Amount',                   decodeMode: 'currency', isValidated: true,  section: 'vat' },
  // Lines 32–34: plain integers — VBA branch `If i >= 31 And i <= 34`
  { lineItem: 32, definition: 'Total Cover Count',                        decodeMode: 'integer',  isValidated: false, section: 'vat' },
  { lineItem: 33, definition: 'Control Number',                           decodeMode: 'integer',  isValidated: false, section: 'vat' },
  { lineItem: 34, definition: 'Total Number of Sales Transactions',       decodeMode: 'integer',  isValidated: false, section: 'vat' },
  // ── Section B: Non-VAT Sales (lines 35–65) ─────────────────────────────
  { lineItem: 35, definition: 'Sales Type',                               decodeMode: 'text',     isValidated: false, section: 'nonvat' },
  { lineItem: 36, definition: 'Amount',                                   decodeMode: 'currency', isValidated: true,  section: 'nonvat' },
  { lineItem: 37, definition: 'Old Accumulated Sales',                    decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 38, definition: 'New Accumulated Sales',                    decodeMode: 'currency', isValidated: true,  section: 'nonvat' },
  { lineItem: 39, definition: 'Total Gross Amount',                       decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 40, definition: 'Total Deductions',                         decodeMode: 'currency', isValidated: true,  section: 'nonvat' },
  { lineItem: 41, definition: 'Total Promo Sales Amount',                 decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 42, definition: 'Senior Citizen Discount / PWD Discount',   decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 43, definition: 'Total Refund Amount',                      decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 44, definition: 'Total Returned Items Amount',              decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 45, definition: 'Total Other Taxes',                        decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 46, definition: 'Total Service Charge Amount',              decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 47, definition: 'Total Adjustment Discount',                decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 48, definition: 'Total Void Amount',                        decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 49, definition: 'Total Discount Cards',                     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 50, definition: 'Total Delivery Charges',                   decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 51, definition: 'Total Gift Certificates/Checks Redeemed', decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 52, definition: 'Store Specific Discount 1 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 53, definition: 'Store Specific Discount 2 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 54, definition: 'Store Specific Discount 3 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 55, definition: 'Store Specific Discount 4 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 56, definition: 'Store Specific Discount 5 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  /**
   * Line 57 — KNOWN ANOMALY:
   * In the original VBA, totalNASD_NV (sum of lines 58–62) is written to
   * Column D of this row WITHOUT the standard ÷100 division applied to every
   * other currency field. This is replicated faithfully. The field is not
   * validated (no Pass/Failed check). The UI surfaces a ⚠️ warning tooltip.
   */
  { lineItem: 57, definition: 'Total of all Non-Approved Store Discounts',decodeMode: 'currency', isValidated: false, section: 'nonvat', knownAnomaly: 'Known anomaly from original VBA workbook: this admin value was computed without the standard ÷100 division applied to all other currency fields. This field is not validated. Displayed for reference only.' },
  { lineItem: 58, definition: 'Store Specific Discount 1 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 59, definition: 'Store Specific Discount 2 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 60, definition: 'Store Specific Discount 3 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 61, definition: 'Store Specific Discount 4 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 62, definition: 'Store Specific Discount 5 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 63, definition: 'Total VAT/Tax Amount',                     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 64, definition: 'Total Net Sales Amount',                   decodeMode: 'currency', isValidated: true,  section: 'nonvat' },
  { lineItem: 65, definition: 'Grand Total Net Sales',                    decodeMode: 'currency', isValidated: true,  section: 'nonvat' },
];

/* ============================================================
   2. FILE DECODER
   Attempts UTF-8 first (strict), falls back to windows-1252
   (covers ANSI/Latin-1 legacy POS files), then iso-8859-1.
   ============================================================ */

/**
 * Decode an ArrayBuffer to a string.
 * @param {ArrayBuffer} buffer
 * @returns {string}
 */
function decodeFileBuffer(buffer) {
  // UTF-8 with fatal=true so we detect encoding errors
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(buffer);
  } catch (_) { /* fall through */ }
  // Windows-1252 covers all ANSI POS files used in the Philippines
  try {
    return new TextDecoder('windows-1252', { fatal: false }).decode(buffer);
  } catch (_) { /* fall through */ }
  // Last resort
  return new TextDecoder('iso-8859-1', { fatal: false }).decode(buffer);
}

/* ============================================================
   3. FILE PARSER
   Parses the .001 fixed-width format into structured line objects
   and populates the accumulator integers needed by the calculator.

   File format per line:
     chars 0-1  : line number (e.g. "01", "34")
     chars 2-13 : 12-char value field (right-padded)
     then CRLF or LF

   Currency decoding: raw integer string → parseFloat → ÷ 100
   Integer decoding:  raw integer string → parseInt  (no division)
   Text decoding:     raw string as-is (trimmed)
   ============================================================ */

/** Round to 2 decimal places — matches VBA Format("#,###,##0.00") */
function round2(n) {
  return Math.round(n * 100) / 100;
}

/**
 * Parse raw .001 file content into lines and accumulators.
 * @param {string} content  Decoded file text
 * @returns {{ lines: Array, acc: Object }}
 */
function parseFile(content) {
  // Normalise line endings: CRLF → LF, lone CR → LF, then split
  const rawLines = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');

  // Strip trailing empty lines (VBA Split on vbCrLf may produce one)
  while (rawLines.length > 0 && rawLines[rawLines.length - 1].trim() === '') {
    rawLines.pop();
  }

  if (rawLines.length < 65) {
    throw new Error(
      `File appears incomplete — expected at least 65 lines, found ${rawLines.length}. ` +
      `Please verify this is a valid POS .001 remittance file.`
    );
  }

  // ── Accumulators (raw integer cents, before ÷100) ──────────────────────
  let oldACCSales      = 0;  // i=3  → line 4
  let totGrossAmount   = 0;  // i=5  → line 6
  let totalDeductions  = 0;  // i=6..21  → lines 7-22
  let totalNASD        = 0;  // i=23..27 → lines 24-28
  let newAccNV         = 0;  // i=36 → line 37 (Old Acc Sales Non-VAT)
  let totGrossNV       = 0;  // i=38 → line 39
  let totalDeductionsNV = 0; // i=40..55 → lines 41-56
  let totalNASD_NV     = 0;  // i=57..61 → lines 58-62

  const lines = [];

  for (let i = 0; i < 65; i++) {
    const lineItem = i + 1;
    const raw = (rawLines[i] || '').substring(2, 14).trim(); // chars 2-13
    const meta = LINE_ITEMS[i]; // LINE_ITEMS is 0-indexed, i corresponds to lineItem-1

    let tenantValue;
    if (meta.decodeMode === 'text') {
      tenantValue = raw;
    } else if (meta.decodeMode === 'integer') {
      tenantValue = parseInt(raw, 10) || 0;
    } else {
      // currency: raw integer ÷ 100
      tenantValue = (parseFloat(raw) || 0) / 100;
    }

    lines.push({ lineItem, rawValue: raw, tenantValue, meta });

    // ── Populate accumulators ───────────────────────────────────────────
    const rawNum = parseFloat(raw) || 0;

    if (i === 3)  oldACCSales     = rawNum;
    if (i === 5)  totGrossAmount  = rawNum;
    if (i >= 7  && i <= 21) totalDeductions  += rawNum;
    if (i >= 23 && i <= 27) totalNASD        += rawNum;
    if (i === 36) newAccNV        = rawNum;
    if (i === 38) totGrossNV      = rawNum;
    if (i >= 40 && i <= 55) totalDeductionsNV += rawNum;
    if (i >= 57 && i <= 61) totalNASD_NV     += rawNum;
  }

  return {
    lines,
    acc: { oldACCSales, totGrossAmount, totalDeductions, totalNASD,
           newAccNV, totGrossNV, totalDeductionsNV, totalNASD_NV },
  };
}

/* ============================================================
   4. CALCULATOR
   Computes all 11 admin-expected Column D values from the raw
   accumulator integers. Faithfully ports the VBA sub
   ReadTextFileAndWriteToColumns from POSFileChecker_v1.xlsm.

   Philippine VAT back-calculation:
     VAT = (GrossAmount − Deductions) × 12 / 112
     Net = (GrossAmount − Deductions) / 1.12

   All inputs are raw integer cents (×100).
   All outputs are rounded to 2dp before returning.
   ============================================================ */

/**
 * @param {Object} acc  Parsed accumulators (raw integer cents)
 * @returns {Object}    Map of lineItem → admin expected value
 */
function computeAdminValues(acc) {
  const { oldACCSales, totGrossAmount, totalDeductions, totalNASD,
          newAccNV, totGrossNV, totalDeductionsNV } = acc;

  // Core VAT intermediates (still in raw cents)
  const vatOnNetSales   = (totGrossAmount - totalDeductions) * 0.12 / 1.12;
  const netSalesExclVAT = totGrossAmount - totalDeductions - vatOnNetSales;
  //   ≡ (totGrossAmount - totalDeductions) / 1.12

  // Core Non-VAT intermediate (raw cents)
  const netSalesNonVAT = totGrossNV - totalDeductionsNV;

  return {
    /**
     * Line 5 — New Accumulated Sales (VAT)
     * = OldAcc/100 + NetSalesExclVAT/100
     */
    5:  round2(oldACCSales / 100 + netSalesExclVAT / 100),

    /**
     * Line 6 — Total Gross Amount (VAT)
     * = totGrossAmount / 100
     * Note: VBA formula strips and re-adds VAT which cancels out —
     * this is a structural identity check; it always equals the
     * tenant's reported gross if the file parses correctly.
     */
    6:  round2(totGrossAmount / 100),

    /**
     * Line 7 — Total Deductions (VAT)
     * = sum of file lines 7–22 / 100
     */
    7:  round2(totalDeductions / 100),

    /**
     * Line 24 — Total Non-Approved Store Discounts (VAT)
     * = sum of file lines 25–29 / 100
     */
    24: round2(totalNASD / 100),

    /**
     * Line 30 — Total VAT/Tax Amount
     * = (Gross − Deductions) × 12/112 / 100
     */
    30: round2(vatOnNetSales / 100),

    /**
     * Line 31 — Total Net Sales Amount (VAT)
     * = (Gross − Deductions) / 1.12 / 100
     */
    31: round2(netSalesExclVAT / 100),

    /**
     * Line 36 — Amount (Non-VAT section header)
     * Same formula as Line 31 — the Non-VAT "Amount" entry is
     * expected to equal the VAT section's net sales.
     */
    36: round2(netSalesExclVAT / 100),

    /**
     * Line 38 — New Accumulated Sales (Non-VAT)
     * = OldAccNV/100 + (GrossNV − DeductionsNV)/100
     */
    38: round2(newAccNV / 100 + netSalesNonVAT / 100),

    /**
     * Line 40 — Total Deductions (Non-VAT)
     * = sum of file lines 41–56 / 100
     */
    40: round2(totalDeductionsNV / 100),

    /**
     * Line 64 — Total Net Sales Amount (Non-VAT)
     * = (GrossNV − DeductionsNV) / 100
     */
    64: round2(netSalesNonVAT / 100),

    /**
     * Line 65 — Grand Total Net Sales
     * = VAT net sales + Non-VAT net sales
     */
    65: round2((netSalesExclVAT + netSalesNonVAT) / 100),
  };
}

/* ============================================================
   5. VALIDATOR
   Orchestrates the full pipeline and builds the result object.
   ============================================================ */

/**
 * Process an ArrayBuffer from a .001 file upload.
 * @param {ArrayBuffer} buffer
 * @param {string} filename
 * @returns {Object} ValidationResult
 */
function processFile(buffer, filename) {
  const content = decodeFileBuffer(buffer);
  const { lines, acc } = parseFile(content);
  const adminValues = computeAdminValues(acc);

  const rows = LINE_ITEMS.map(meta => {
    const parsed = lines.find(l => l.lineItem === meta.lineItem);
    const tenantValue = parsed ? parsed.tenantValue : null;
    const adminValue  = meta.isValidated ? (adminValues[meta.lineItem] ?? null) : null;

    let status = null;
    if (meta.isValidated && tenantValue !== null && adminValue !== null) {
      // Round both to 2dp before comparing — per stakeholder decision.
      const tRounded = round2(typeof tenantValue === 'number' ? tenantValue : 0);
      const aRounded = round2(adminValue);
      status = (tRounded === aRounded) ? 'Pass' : 'Failed';
    }

    return {
      lineItem:   meta.lineItem,
      definition: meta.definition,
      tenantValue,
      adminValue,
      status,
      isValidated: meta.isValidated,
      section:    meta.section,
      knownAnomaly: meta.knownAnomaly || null,
    };
  });

  const validated  = rows.filter(r => r.isValidated);
  const passed     = validated.filter(r => r.status === 'Pass').length;
  const failed     = validated.filter(r => r.status === 'Failed').length;

  return {
    filename,
    processedAt: new Date(),
    tenantCode:     String(lines[0]?.tenantValue || '').trim() || 'Unknown',
    terminalNumber: String(lines[1]?.tenantValue || '').trim() || 'Unknown',
    posDate:        String(lines[2]?.tenantValue || '').trim() || 'Unknown',
    rows,
    summary: { totalChecked: validated.length, passed, failed, allPassed: failed === 0 },
  };
}

/* ============================================================
   6. FORMATTERS
   ============================================================ */

/**
 * Format a numeric value as PHP currency (2dp, locale-aware).
 * @param {number|string|null} value
 * @returns {string}
 */
function fmtCurrency(value) {
  if (value === null || value === undefined) return '—';
  if (typeof value === 'string') return value;
  return value.toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/**
 * Format a POS date string (mmddyyyy) for display.
 * @param {string} raw
 * @returns {string}
 */
function fmtDate(raw) {
  if (!raw || raw === 'Unknown') return raw;
  if (/^\d{8}$/.test(raw)) {
    return `${raw.slice(0,2)}/${raw.slice(2,4)}/${raw.slice(4)}`;
  }
  return raw;
}

/**
 * Format a cell value for display in the table.
 * @param {number|string|null} value
 * @param {string} decodeMode
 * @returns {string}
 */
function fmtCellValue(value, decodeMode) {
  if (value === null || value === undefined) return '—';
  if (decodeMode === 'text' || decodeMode === 'integer') return String(value);
  return fmtCurrency(value);
}

/* ============================================================
   7. PDF EXPORT
   Uses jsPDF + jsPDF-AutoTable loaded from CDN in index.html.
   Generates a landscape A4 report matching the spreadsheet
   column layout (A–E).
   ============================================================ */

/**
 * Export the validation result as a PDF and trigger download.
 * @param {Object} result  ValidationResult
 */
async function exportToPDF(result) {
  // jsPDF and autoTable are loaded globally via <script> in index.html
  if (typeof window.jspdf === 'undefined') {
    alert('PDF export library is loading. Please try again in a moment.');
    return;
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

  // ── Header ─────────────────────────────────────────────────────────────
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(14);
  doc.setTextColor(15, 28, 46);
  doc.text('POS File Validation Report', 14, 16);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.setTextColor(80, 100, 120);
  doc.text(`File: ${result.filename}`, 14, 23);
  doc.text(`Tenant Code: ${result.tenantCode}`, 14, 28);
  doc.text(`Terminal No.: ${result.terminalNumber}`, 75, 28);
  doc.text(`POS Date: ${fmtDate(result.posDate)}`, 145, 28);
  doc.text(`Processed: ${result.processedAt.toLocaleString('en-PH')}`, 14, 33);

  // ── Summary pills ──────────────────────────────────────────────────────
  const { passed, failed, totalChecked } = result.summary;
  const sy = 38;
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);

  doc.setFillColor(0, 200, 150);
  doc.roundedRect(14, sy, 38, 8, 2, 2, 'F');
  doc.setTextColor(255, 255, 255);
  doc.text(`\u2713 ${passed} Passed`, 33, sy + 5.5, { align: 'center' });

  doc.setFillColor(229, 62, 62);
  doc.roundedRect(56, sy, 38, 8, 2, 2, 'F');
  doc.text(`\u2717 ${failed} Failed`, 75, sy + 5.5, { align: 'center' });

  doc.setFillColor(42, 78, 122);
  doc.roundedRect(98, sy, 46, 8, 2, 2, 'F');
  doc.text(`${totalChecked} Checks Total`, 121, sy + 5.5, { align: 'center' });

  // ── Table ──────────────────────────────────────────────────────────────
  const tableData = result.rows.map(row => [
    String(row.lineItem),
    row.definition,
    row.tenantValue !== null ? fmtCellValue(row.tenantValue, LINE_ITEMS[row.lineItem - 1].decodeMode) : '—',
    row.adminValue  !== null ? fmtCurrency(row.adminValue)  : '—',
    row.status || '—',
  ]);

  doc.autoTable({
    startY: sy + 12,
    head: [['#', 'Line Item Definition', 'Tenant File Value (C)', 'Admin Expected (D)', 'Status (E)']],
    body: tableData,
    styles: { font: 'helvetica', fontSize: 7.5, cellPadding: 2.5, textColor: [30, 50, 70] },
    headStyles: { fillColor: [15, 28, 46], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
    columnStyles: {
      0: { cellWidth: 10, halign: 'center' },
      1: { cellWidth: 92 },
      2: { cellWidth: 42, halign: 'right', font: 'courier' },
      3: { cellWidth: 42, halign: 'right', font: 'courier' },
      4: { cellWidth: 22, halign: 'center' },
    },
    didDrawCell(data) {
      if (data.column.index === 4 && data.section === 'body') {
        if (data.cell.raw === 'Pass')   doc.setTextColor(0, 160, 100);
        else if (data.cell.raw === 'Failed') doc.setTextColor(220, 50, 50);
        else doc.setTextColor(120, 140, 160);
      }
    },
    alternateRowStyles: { fillColor: [245, 248, 252] },
    margin: { left: 14, right: 14 },
  });

  // ── Footer ─────────────────────────────────────────────────────────────
  const pages = doc.internal.getNumberOfPages();
  for (let p = 1; p <= pages; p++) {
    doc.setPage(p);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(140, 150, 160);
    doc.text(
      `POS File Checker  |  Page ${p} of ${pages}`,
      doc.internal.pageSize.width / 2,
      doc.internal.pageSize.height - 6,
      { align: 'center' }
    );
  }

  doc.save(`validation_${result.filename.replace(/[^a-zA-Z0-9._-]/g, '_')}.pdf`);
}

/* ============================================================
   8. EXCEL EXPORT
   Uses SheetJS (XLSX) loaded from CDN in index.html.
   Produces a two-sheet workbook: Summary + Validation data,
   mirroring the original spreadsheet's column structure.
   ============================================================ */

/**
 * Export the validation result as an .xlsx file and trigger download.
 * @param {Object} result  ValidationResult
 */
function exportToExcel(result) {
  if (typeof window.XLSX === 'undefined') {
    alert('Excel export library is loading. Please try again in a moment.');
    return;
  }
  const XLSX = window.XLSX;
  const wb = XLSX.utils.book_new();

  // ── Summary sheet ───────────────────────────────────────────────────────
  const summaryData = [
    ['POS File Validation Report'],
    [],
    ['File',            result.filename],
    ['Tenant Code',     result.tenantCode],
    ['Terminal Number', result.terminalNumber],
    ['POS Date',        fmtDate(result.posDate)],
    ['Processed At',    result.processedAt.toLocaleString('en-PH')],
    [],
    ['Checks Passed',   result.summary.passed],
    ['Checks Failed',   result.summary.failed],
    ['Total Checks',    result.summary.totalChecked],
    ['Overall Result',  result.summary.allPassed ? 'ALL PASS' : `${result.summary.failed} FAILED`],
  ];
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  summarySheet['!cols'] = [{ wch: 20 }, { wch: 42 }];
  XLSX.utils.book_append_sheet(wb, summarySheet, 'Summary');

  // ── Validation sheet (mirrors Columns A–E of the spreadsheet) ────────────
  const headers = [
    'Line Item #',
    'Line Item Definition',
    'Tenant File Value',   // Column C
    'Admin Expected Value',// Column D
    'Status',              // Column E
  ];
  const dataRows = result.rows.map(row => [
    row.lineItem,
    row.definition,
    row.tenantValue !== null
      ? (typeof row.tenantValue === 'number' ? row.tenantValue : String(row.tenantValue))
      : null,
    row.adminValue !== null ? row.adminValue : null,
    row.status || '',
  ]);
  const valSheet = XLSX.utils.aoa_to_sheet([headers, ...dataRows]);
  valSheet['!cols'] = [{ wch: 12 }, { wch: 46 }, { wch: 20 }, { wch: 22 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, valSheet, 'Validation');

  XLSX.writeFile(wb, `validation_${result.filename.replace(/[^a-zA-Z0-9._-]/g, '_')}.xlsx`);
}

/* ============================================================
   9. UI RENDERING
   ============================================================ */

/**
 * Application views:
 *   'upload'  — drop zone (always the starting state)
 *   'batch'   — multi-file summary table
 *   'detail'  — single-file 65-row validation detail (existing results-view)
 *
 * Single-file uploads skip 'batch' and go straight to 'detail'.
 * From 'detail', the Back button returns to 'batch' if the session
 * was a multi-file upload, or resets to 'upload' if it was single.
 */

/** Currently active filter in the detail view */
let activeFilter = 'all';

/** The file currently shown in the detail view (ValidationResult | null) */
let currentResult = null;

/**
 * All results from the most recent batch upload.
 * Length === 1 for single-file uploads.
 * @type {Array<{result: Object|null, error: string|null, filename: string}>}
 */
let batchResults = [];

// ── DOM references ────────────────────────────────────────────────────────
const uploadView       = document.getElementById('upload-view');
const batchView        = document.getElementById('batch-view');
const resultsView      = document.getElementById('results-view');
const dropZone         = document.getElementById('drop-zone');
const fileInput        = document.getElementById('file-input');
const dropTextPrimary  = document.getElementById('drop-text-primary');
const dropTextSecondary = document.getElementById('drop-text-secondary');
const dropIconBox      = document.getElementById('drop-icon-box');
const errorBanner      = document.getElementById('error-banner');
const errorMsg         = document.getElementById('error-msg');

// Detail view elements
const resultBanner    = document.getElementById('result-banner');
const metaPills       = document.getElementById('meta-pills');
const statPassed      = document.getElementById('stat-passed');
const statFailed      = document.getElementById('stat-failed');
const statTotal       = document.getElementById('stat-total');
const filterTabs      = document.querySelectorAll('.filter-tab');
const tabCountAll     = document.getElementById('tab-count-all');
const tabCountVal     = document.getElementById('tab-count-val');
const tabCountFail    = document.getElementById('tab-count-fail');
const tableBody       = document.getElementById('table-body');
const btnExportPDF    = document.getElementById('btn-export-pdf');
const btnExportExcel  = document.getElementById('btn-export-excel');
const resultsSubtitle = document.getElementById('results-subtitle');
const btnBackToBatch  = document.getElementById('btn-back-to-batch');

// Batch view elements
const batchTableBody  = document.getElementById('batch-table-body');
const batchErrors     = document.getElementById('batch-errors');
const batchSubtitle   = document.getElementById('batch-subtitle');
const batchPassCount  = document.getElementById('batch-pass-count');
const batchFailCount  = document.getElementById('batch-fail-count');
const batchTotalCount = document.getElementById('batch-total-count');
const batchStatFail   = document.getElementById('batch-stat-fail');

// ── View switching ────────────────────────────────────────────────────────

function showView(name) {
  uploadView.style.display  = (name === 'upload')  ? '' : 'none';
  batchView.classList.toggle('visible',  name === 'batch');
  resultsView.classList.toggle('visible', name === 'detail');
}

/** Show an error message. */
function showError(msg) {
  errorMsg.textContent = msg;
  errorBanner.classList.add('visible');
}

function hideError() {
  errorBanner.classList.remove('visible');
}

/**
 * Set the drop zone into processing state.
 * When count > 1, show "Processing N files…"
 */
function setProcessing(on, count) {
  if (on) {
    dropZone.classList.add('processing');
    dropIconBox.innerHTML = '<div class="drop-spinner"></div>';
    dropTextPrimary.textContent = count > 1
      ? `Processing ${count} files\u2026`
      : 'Processing file\u2026';
    dropTextSecondary.textContent = 'Parsing and running validation checks';
  } else {
    dropZone.classList.remove('processing');
    dropIconBox.innerHTML = iconUpload();
    dropTextPrimary.innerHTML = 'Drop your <span>.001&nbsp;&ndash;&nbsp;.999</span> files here';
    dropTextSecondary.textContent = 'or browse to select';
  }
}

// ── Detail view rendering (single file) ──────────────────────────────────

/** Render the detail view for one ValidationResult. */
function renderDetail(result) {
  const { passed, failed, totalChecked, allPassed } = result.summary;

  // Overall banner
  if (allPassed) {
    resultBanner.className = 'result-banner all-pass';
    resultBanner.innerHTML = `${iconCheckCircle()}<span>All ${totalChecked} checks passed — file is valid.</span>`;
  } else {
    resultBanner.className = 'result-banner has-fail';
    resultBanner.innerHTML = `${iconXCircle()}<span>${failed} of ${totalChecked} checks failed — review highlighted rows below.</span>`;
  }

  // Meta pills
  metaPills.innerHTML = [
    metaPill('File',     result.filename,              true),
    metaPill('Tenant',   result.tenantCode,             false),
    metaPill('Terminal', result.terminalNumber,         false),
    metaPill('Date',     fmtDate(result.posDate),       false),
  ].join('');

  // Stat cards
  statPassed.textContent = passed;
  statFailed.textContent = failed;
  statTotal.textContent  = totalChecked;

  // Filter tab counts
  tabCountAll.textContent  = result.rows.length;
  tabCountVal.textContent  = result.rows.filter(r => r.isValidated).length;
  tabCountFail.textContent = failed;
  tabCountFail.className   = 'tab-count' + (failed > 0 ? ' fail-count' : '');

  resultsSubtitle.textContent =
    `65 line items · 11 validated fields · processed ${result.processedAt.toLocaleTimeString('en-PH')}`;
}

/** Create a meta pill HTML string. */
function metaPill(label, value, mono) {
  return `<span class="meta-pill">
    <span class="label">${esc(label)}:</span>
    <span class="value${mono ? ' mono' : ''}">${esc(value)}</span>
  </span>`;
}

/** Render the 65-row validation table for the current filter. */
function renderTable() {
  if (!currentResult) return;

  const rows = currentResult.rows.filter(row => {
    if (activeFilter === 'validated') return row.isValidated;
    if (activeFilter === 'failed')    return row.status === 'Failed';
    return true;
  });

  if (rows.length === 0) {
    tableBody.innerHTML = `<tr><td colspan="5" class="table-empty">No rows match the selected filter.</td></tr>`;
    return;
  }

  let html = '';
  let prevSection = null;

  rows.forEach(row => {
    if (activeFilter === 'all' && row.section === 'nonvat' && prevSection === 'vat') {
      html += `<tr class="section-sep">
        <td colspan="5">
          <div class="section-sep-inner">
            <div class="section-sep-line"></div>
            <span class="section-sep-label">Non-VAT Sales Section</span>
            <div class="section-sep-line"></div>
          </div>
        </td>
      </tr>`;
    }
    prevSection = row.section;

    const stripe = row.status === 'Pass'   ? 'stripe-pass'
                 : row.status === 'Failed' ? 'stripe-fail'
                 : 'stripe-none';

    const tMeta    = LINE_ITEMS[row.lineItem - 1];
    const tDisplay = row.tenantValue !== null ? fmtCellValue(row.tenantValue, tMeta.decodeMode) : '—';
    const aDisplay = row.adminValue  !== null ? fmtCurrency(row.adminValue) : '—';
    const tMuted   = tDisplay === '—';
    const aMuted   = aDisplay === '—';

    let statusCell;
    if (row.status === 'Pass') {
      statusCell = `<span class="badge pass">Pass</span>`;
    } else if (row.status === 'Failed') {
      statusCell = `<button class="btn-expand" id="expand-${row.lineItem}" aria-expanded="false" aria-controls="detail-${row.lineItem}" onclick="toggleDetail(${row.lineItem})">
        <span class="badge fail">Failed</span>
        ${iconChevronDown()}
      </button>`;
    } else {
      statusCell = `<span style="color:var(--slate-muted);font-size:.75rem">—</span>`;
    }

    let anomalyBadge = '';
    if (row.knownAnomaly) {
      anomalyBadge = `<span class="anomaly-wrap">
        <button class="anomaly-btn" tabindex="0" aria-label="Known anomaly">${iconWarning()}</button>
        <div class="anomaly-tip"><p class="anomaly-tip-title">Known Anomaly</p>${esc(row.knownAnomaly)}</div>
      </span>`;
    }

    html += `<tr class="${stripe}" id="row-${row.lineItem}">
      <td class="td-num">${row.lineItem}</td>
      <td class="td-def${row.isValidated ? ' validated' : ''}">
        <div class="td-def-inner">${esc(row.definition)}${anomalyBadge}</div>
      </td>
      <td class="td-val${row.isValidated ? ' validated' : tMuted ? ' muted' : ''}">${tDisplay}</td>
      <td class="td-val${aMuted ? ' muted' : ''}">${aDisplay}</td>
      <td class="td-status">${statusCell}</td>
    </tr>`;

    if (row.status === 'Failed') {
      const tenantNum = typeof row.tenantValue === 'number' ? row.tenantValue : 0;
      const adminNum  = typeof row.adminValue  === 'number' ? row.adminValue  : 0;
      html += `<tr class="detail-row" id="detail-${row.lineItem}">
        <td colspan="5">
          <div class="detail-inner">
            <p class="detail-title">Discrepancy Detail</p>
            <div class="detail-cards">
              <div class="detail-card">
                <div class="detail-card-label">Tenant reported</div>
                <div class="detail-card-value">${fmtCellValue(row.tenantValue, tMeta.decodeMode)}</div>
              </div>
              <div class="detail-card">
                <div class="detail-card-label">Admin expected</div>
                <div class="detail-card-value">${fmtCurrency(row.adminValue)}</div>
              </div>
            </div>
            <p class="detail-diff">Difference: <span>${fmtCurrency(tenantNum - adminNum)}</span></p>
          </div>
        </td>
      </tr>`;
    }
  });

  tableBody.innerHTML = html;
}

/** Toggle a Failed row's discrepancy tray. */
function toggleDetail(lineItem) {
  const btn    = document.getElementById(`expand-${lineItem}`);
  const detail = document.getElementById(`detail-${lineItem}`);
  if (!btn || !detail) return;
  const open = detail.classList.toggle('open');
  btn.classList.toggle('open', open);
  btn.setAttribute('aria-expanded', String(open));
}
window.toggleDetail = toggleDetail;

/** Update which filter tab is active and re-render the detail table. */
function updateFilterTabs() {
  filterTabs.forEach(tab => {
    tab.classList.toggle('active', tab.dataset.filter === activeFilter);
  });
  renderTable();
}

// ── Batch view rendering ──────────────────────────────────────────────────

/**
 * Build a summary row for the batch table.
 * @param {Object} result  ValidationResult
 * @param {number} index   Index into batchResults (used by the View Detail button)
 */
function batchRow(result, index) {
  const { passed, failed, totalChecked, allPassed } = result.summary;

  // Key totals: Net Sales VAT (line 31) and Grand Total (line 65)
  const row31 = result.rows.find(r => r.lineItem === 31);
  const row65 = result.rows.find(r => r.lineItem === 65);
  const netSalesVAT  = row31 ? fmtCurrency(row31.tenantValue) : '—';
  const grandTotal   = row65 ? fmtCurrency(row65.tenantValue) : '—';

  const statusBadge = allPassed
    ? `<span class="badge pass">All Pass</span>`
    : `<span class="badge fail">${failed} Failed</span>`;

  const checksCell = `<span class="td-checks">
    <span class="pass-n">${passed}</span>/<span class="fail-n${failed > 0 ? '' : '" style="color:var(--slate-muted)'}">${failed}</span>
    <span style="color:var(--slate-muted);font-size:.7rem"> of ${totalChecked}</span>
  </span>`;

  const stripe = allPassed ? 'stripe-pass' : 'stripe-fail';

  return `<tr class="batch-row ${stripe}" onclick="openDetailFromBatch(${index})">
    <td class="td-num" style="padding-left:1.125rem">${index + 1}</td>
    <td><span class="td-filename" title="${esc(result.filename)}">${esc(result.filename)}</span></td>
    <td><span style="font-size:.8125rem">${esc(result.tenantCode)}</span></td>
    <td><span style="font-size:.8125rem">${esc(result.terminalNumber)}</span></td>
    <td><span style="font-size:.8125rem">${esc(fmtDate(result.posDate))}</span></td>
    <td class="td-val">${netSalesVAT}</td>
    <td class="td-val">${grandTotal}</td>
    <td class="td-status">${checksCell}</td>
    <td class="td-status">${statusBadge}</td>
    <td class="td-status">
      <button class="btn-detail" onclick="event.stopPropagation(); openDetailFromBatch(${index})">
        View ${iconChevronRight()}
      </button>
    </td>
  </tr>`;
}

/** Render the batch summary view from batchResults. */
function renderBatchView() {
  const successful = batchResults.filter(b => b.result !== null);
  const errored    = batchResults.filter(b => b.result === null);

  const allPassed  = successful.filter(b => b.result.summary.allPassed).length;
  const anyFailed  = successful.filter(b => !b.result.summary.allPassed).length;

  // Header stats (count of files, not individual checks)
  batchPassCount.textContent  = allPassed;
  batchFailCount.textContent  = anyFailed;
  batchTotalCount.textContent = batchResults.length;
  batchStatFail.className     = 'batch-stat fail' + (anyFailed > 0 ? '' : ' ' + 'batch-stat-zero');

  batchSubtitle.textContent =
    `${batchResults.length} file${batchResults.length !== 1 ? 's' : ''} processed · ` +
    `${allPassed} all-pass · ${anyFailed} with failures` +
    (errored.length > 0 ? ` · ${errored.length} could not be parsed` : '');

  // Build table rows (successful files only)
  batchTableBody.innerHTML = successful.length > 0
    ? successful.map((b, i) => batchRow(b.result, batchResults.indexOf(b))).join('')
    : `<tr><td colspan="10" class="table-empty">No files were successfully parsed.</td></tr>`;

  // Build error list
  batchErrors.innerHTML = errored.map(b => `
    <div class="batch-error-item">
      <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
      <div>
        <p class="batch-error-file">${esc(b.filename)}</p>
        <p class="batch-error-msg">${esc(b.error)}</p>
      </div>
    </div>`
  ).join('');
}

/**
 * Open the detail view for a specific file in the current batch.
 * The Back button will return to the batch view.
 * @param {number} index  Index into batchResults
 */
function openDetailFromBatch(index) {
  const entry = batchResults[index];
  if (!entry || !entry.result) return;

  currentResult = entry.result;
  activeFilter  = 'all';
  updateFilterTabs();
  renderDetail(currentResult);

  // Show Back button; hide Reset button
  btnBackToBatch.style.display = '';
  document.getElementById('btn-reset').style.display = 'none';

  showView('detail');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}
window.openDetailFromBatch = openDetailFromBatch;

// ── Entry points ──────────────────────────────────────────────────────────

/**
 * Show the detail view for a single-file result.
 * Reset button visible; Back button hidden.
 */
function showSingleResult(result) {
  currentResult = result;
  activeFilter  = 'all';
  updateFilterTabs();
  renderDetail(result);
  btnBackToBatch.style.display = 'none';
  document.getElementById('btn-reset').style.display = '';
  showView('detail');
  hideError();
}

/** Return to the batch view from the detail view. */
function goBackToBatch() {
  currentResult = null;
  showView('batch');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

/** Reset everything back to the upload view. */
function showUpload() {
  currentResult = null;
  batchResults  = [];
  showView('upload');
  setProcessing(false, 0);
  hideError();
  fileInput.value = '';
}

/* ============================================================
   10. EVENT WIRING
   ============================================================ */

/**
 * Return true if the filename has a valid POS numeric extension.
 * Accepts exactly 3 decimal digits (.001 – .999).
 * @param {string} filename
 * @returns {boolean}
 */
function isValidPosExtension(filename) {
  const dot = filename.lastIndexOf('.');
  if (dot === -1) return false;
  const ext = filename.slice(dot + 1);
  return /^\d{3}$/.test(ext);
}

/**
 * Process an array of File objects.
 * - Single valid file  → validate → detail view directly
 * - Multiple files     → validate each → batch summary view
 * - Any invalid ext    → showError for that file; skip it
 */
async function handleFiles(files) {
  if (!files || files.length === 0) return;

  // Partition into valid / invalid by extension
  const valid   = Array.from(files).filter(f => isValidPosExtension(f.name));
  const invalid = Array.from(files).filter(f => !isValidPosExtension(f.name));

  if (valid.length === 0) {
    // All files rejected
    const names = invalid.slice(0, 3).map(f => {
      const dot = f.name.lastIndexOf('.');
      return dot !== -1 ? f.name.slice(dot) : '(no extension)';
    }).join(', ');
    const more = invalid.length > 3 ? ` and ${invalid.length - 3} more` : '';
    showError(
      `No valid POS files found. Extensions must be 3 digits (.001 – .999). ` +
      `Received: ${names}${more}`
    );
    return;
  }

  // Warn about any invalid files mixed in with valid ones (non-blocking)
  if (invalid.length > 0) {
    const names = invalid.map(f => f.name).join(', ');
    showError(
      `${invalid.length} file${invalid.length > 1 ? 's were' : ' was'} skipped ` +
      `(invalid extension): ${names}`
    );
    // Don't return — continue with the valid files
  } else {
    hideError();
  }

  setProcessing(true, valid.length);

  // Read all files as ArrayBuffers concurrently
  const buffers = await Promise.all(valid.map(f => f.arrayBuffer()));

  // Process each file. We use a small setTimeout so the spinner paints first,
  // then process synchronously (all CPU work stays off the main thread queue).
  await new Promise(resolve => setTimeout(resolve, 60));

  batchResults = valid.map((file, i) => {
    try {
      const result = processFile(buffers[i], file.name);
      return { filename: file.name, result, error: null };
    } catch (err) {
      return { filename: file.name, result: null, error: err.message || 'Parse error' };
    }
  });

  setProcessing(false, 0);

  const successful = batchResults.filter(b => b.result !== null);

  if (successful.length === 0) {
    // All files failed to parse
    showView('upload');
    showError(batchResults[0].error || 'All files failed to parse.');
    return;
  }

  if (valid.length === 1 && successful.length === 1) {
    // Single valid file — go straight to detail view (preserves existing UX)
    showSingleResult(successful[0].result);
  } else {
    // Multiple files — show batch summary
    renderBatchView();
    showView('batch');
  }
}

// ── Drop zone ─────────────────────────────────────────────────────────────
dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  handleFiles(e.dataTransfer?.files);
});
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('keydown', e => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
});

// ── File input ────────────────────────────────────────────────────────────
fileInput.addEventListener('change', () => {
  handleFiles(fileInput.files);
});

// ── Error banner close ────────────────────────────────────────────────────
document.getElementById('error-close').addEventListener('click', hideError);

// ── Reset / Back buttons ──────────────────────────────────────────────────
document.getElementById('btn-reset').addEventListener('click', showUpload);
btnBackToBatch.addEventListener('click', goBackToBatch);
document.getElementById('btn-batch-reset').addEventListener('click', showUpload);

// ── Filter tabs ───────────────────────────────────────────────────────────
filterTabs.forEach(tab => {
  tab.addEventListener('click', () => {
    activeFilter = tab.dataset.filter;
    updateFilterTabs();
  });
});

// ── Export buttons ────────────────────────────────────────────────────────
btnExportPDF.addEventListener('click', async () => {
  if (!currentResult) return;
  btnExportPDF.disabled = true;
  try   { await exportToPDF(currentResult); }
  finally { btnExportPDF.disabled = false; }
});

btnExportExcel.addEventListener('click', () => {
  if (!currentResult) return;
  btnExportExcel.disabled = true;
  try   { exportToExcel(currentResult); }
  finally { btnExportExcel.disabled = false; }
});

/* ============================================================
   INLINE SVG HELPERS
   ============================================================ */
function iconUpload() {
  return `<svg viewBox="0 0 24 24"><polyline points="16 16 12 12 8 16"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>`;
}
function iconCheckCircle() {
  return `<svg viewBox="0 0 24 24" style="width:20px;height:20px;stroke:currentColor;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;flex-shrink:0"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>`;
}
function iconXCircle() {
  return `<svg viewBox="0 0 24 24" style="width:20px;height:20px;stroke:currentColor;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;flex-shrink:0"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>`;
}
function iconChevronDown() {
  return `<svg viewBox="0 0 24 24"><polyline points="6 9 12 15 18 9"/></svg>`;
}
function iconChevronRight() {
  return `<svg viewBox="0 0 24 24"><polyline points="9 18 15 12 9 6"/></svg>`;
}
function iconWarning() {
  return `<svg viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>`;
}

/** HTML-escape a string for safe insertion. */
function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── Initialise ────────────────────────────────────────────────────────────
dropIconBox.innerHTML = iconUpload();
