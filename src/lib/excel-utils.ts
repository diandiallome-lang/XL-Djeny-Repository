
import * as XLSX from "xlsx";

export interface ExcelMetadata {
  sheets: string[];
  columns: number;
  rows: number;
}

export interface FormulaPreview {
  column: string;
  formula: string;
}

export const getExcelMetadata = async (file: File): Promise<ExcelMetadata> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1");

        resolve({
          sheets: workbook.SheetNames,
          columns: range.e.c - range.s.c + 1,
          rows: range.e.r - range.s.r + 1,
        });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

/**
 * Extracts the formula preview from a template file (row 2 = formula row).
 */
export const getTemplateFormulas = async (file: File): Promise<FormulaPreview[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array", cellFormula: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
        const result: FormulaPreview[] = [];

        for (let c = range.s.c; c <= range.e.c; c++) {
          const cellAddr = XLSX.utils.encode_cell({ r: 1, c }); // Row 2 (0-based index 1)
          const cell = sheet[cellAddr];
          if (cell?.f) {
            result.push({
              column: XLSX.utils.encode_col(c),
              formula: `=${cell.f}`,
            });
          }
        }
        resolve(result);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

/**
 * Adjusts row references in an Excel formula.
 *
 * Only RELATIVE row references are shifted; absolute references ($) are preserved.
 *
 * Examples (fromRow=2, toRow=5, offset=+3):
 *   "A2+B2"          → "A5+B5"
 *   "SUM(A2:A10)"    → "SUM(A5:A13)"
 *   "VLOOKUP(A2,$C$1:$D$100,2,0)" → "VLOOKUP(A5,$C$1:$D$100,2,0)"
 */
function adjustFormulaRow(formula: string, fromExcelRow: number, toExcelRow: number): string {
  const offset = toExcelRow - fromExcelRow;
  if (offset === 0) return formula;

  // Robust regex to avoid matching function names (e.g., LOG10) or word endings instead of cell references.
  // Group 1: Prefix character (start of string or any non-letter character)
  // Group 2: Column (1-3 letters, optional $)
  // Group 3: Optional absolute row marker ($)
  // Group 4: Row digits
  const regex = /(^|[^A-Za-z])(\$?[A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_.\(])/g;

  return formula.replace(regex, (match, prefix, col, dollar, row) => {
    if (dollar === "$") return match; // Absolute row
    const newRow = parseInt(row, 10) + offset;
    if (newRow < 1) return match; // Invalid row result
    return `${prefix}${col}${newRow}`;
  });
}

const readWorkbook = (file: File): Promise<XLSX.WorkBook> =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      resolve(XLSX.read(data, { type: "array", cellFormula: true }));
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });

/**
 * Applies the formulas from a template file to a raw data file.
 */
export const applyTemplateToData = async (
  templateFile: File,
  dataFile: File
): Promise<Blob> => {
  const templateWb = await readWorkbook(templateFile);
  const dataWb = await readWorkbook(dataFile);

  const newWb = XLSX.utils.book_new();

  // Process each sheet present in the template
  for (const sheetName of templateWb.SheetNames) {
    const templateSheet = templateWb.Sheets[sheetName];

    // Match by sheet name; fall back to first sheet of data file
    const dataSheetName = dataWb.SheetNames.includes(sheetName)
      ? sheetName
      : dataWb.SheetNames[0];
    const dataSheet = dataWb.Sheets[dataSheetName];

    if (!templateSheet || !dataSheet) continue;

    const templateRange = XLSX.utils.decode_range(templateSheet["!ref"] || "A1");
    const dataRange = XLSX.utils.decode_range(dataSheet["!ref"] || "A1");

    // Row 2 of the template (0-based index 1) is the formula pattern row
    const FORMULA_ROW_IDX = 1;
    const FORMULA_EXCEL_ROW = 2; // 1-based row number as seen in Excel

    // Collect formula cells keyed by column index to preserve formatting
    const formulaCells: Record<number, XLSX.CellObject> = {};
    for (let c = templateRange.s.c; c <= templateRange.e.c; c++) {
      const cell = templateSheet[XLSX.utils.encode_cell({ r: FORMULA_ROW_IDX, c })];
      if (cell && cell.f) {
        formulaCells[c] = cell;
      }
    }

    const outputSheet: XLSX.WorkSheet = {};

    // ── Row 1: copy headers from template ──────────────────────────────────
    for (let c = templateRange.s.c; c <= templateRange.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c });
      const cell = templateSheet[addr];
      if (cell) outputSheet[addr] = { ...cell };
    }

    // Determine max column width for the output
    const maxCol = Math.max(templateRange.e.c, dataRange.e.c);
    let lastOutputRow = 0;

    // ── Rows 2+: data rows from data file (skip its header at dr=0) ────────
    for (let dr = 1; dr <= dataRange.e.r; dr++) {
      const outputRowIdx = dr; // Same 0-based index: data row 1 → output index 1 (Excel row 2)
      const outputExcelRow = outputRowIdx + 1; // 1-based Excel row number

      for (let c = 0; c <= maxCol; c++) {
        const outputAddr = XLSX.utils.encode_cell({ r: outputRowIdx, c });

        if (formulaCells[c] !== undefined) {
          // Column has a formula in the template → apply with row adjustment
          // and preserve formatting styles from the original cell.
          const tCell = formulaCells[c];
          if (tCell.f) {
            const adjusted = adjustFormulaRow(
              tCell.f,
              FORMULA_EXCEL_ROW,
              outputExcelRow
            );
            outputSheet[outputAddr] = { ...tCell, f: adjusted, v: undefined };
          }
        } else {
          // Column is a plain data column → copy value from data file
          const dataAddr = XLSX.utils.encode_cell({ r: dr, c });
          const dataCell = dataSheet[dataAddr];
          if (dataCell !== undefined) {
            outputSheet[outputAddr] = { ...dataCell };
          }
        }
      }

      lastOutputRow = outputRowIdx;
    }

    // Set the used range of the output sheet
    outputSheet["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: lastOutputRow, c: maxCol },
    });

    // Preserve column widths from template if available
    if (templateSheet["!cols"]) {
      outputSheet["!cols"] = templateSheet["!cols"];
    }

    XLSX.utils.book_append_sheet(newWb, outputSheet, sheetName);
  }

  const out = XLSX.write(newWb, { type: "array", bookType: "xlsx" });
  return new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
};
