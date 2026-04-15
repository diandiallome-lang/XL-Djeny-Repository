
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
        
        // Prioritize the 'Traitement' sheet; otherwise fall back to the first sheet.
        const targetSheetName = wb.SheetNames.includes("Traitement") 
          ? "Traitement" 
          : wb.SheetNames[0];
        
        const sheet = wb.Sheets[targetSheetName];
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
  console.log("Starting mirroring application...");
  const templateWb = await readWorkbook(templateFile);
  const dataWb = await readWorkbook(dataFile);
  console.log("Workbooks loaded.");

  // 1. Identify sources
  const templateTraitName = templateWb.SheetNames.includes("Traitement") ? "Traitement" : templateWb.SheetNames[0];
  const templateTraitSheet = templateWb.Sheets[templateTraitName];
  
  const dataAnchorName = dataWb.SheetNames[0];
  const dataAnchorSheet = dataWb.Sheets[dataAnchorName];

  if (!templateTraitSheet || !dataAnchorSheet) {
    throw new Error("Missing required sheets for processing.");
  }

  // 2. Determine Bounds
  const templateRange = XLSX.utils.decode_range(templateTraitSheet["!ref"] || "A1");
  
  // Calculate true data range (ignoring empty trailing rows/cols if any)
  let actualDataMaxRow = 0;
  Object.keys(dataAnchorSheet).forEach(key => {
    if (key[0] === '!') return;
    const cell = XLSX.utils.decode_cell(key);
    if (cell.r > actualDataMaxRow) actualDataMaxRow = cell.r;
  });
  
  const FORMULA_ROW_IDX = 1; // Pattern row (Row 2 in Excel)
  const FORMULA_EXCEL_ROW = 2;

  // 3. Create the Output Workbook
  const newWb = XLSX.utils.book_new();
  
  // 4. Copy all original sheets from the data file first
  for (const sn of dataWb.SheetNames) {
    if (sn === "Traitement") continue; // We will handle Traitement separately
    XLSX.utils.book_append_sheet(newWb, dataWb.Sheets[sn], sn);
  }

  // 5. Generate the NEW 'Traitement' sheet
  const outTraitSheet: XLSX.WorkSheet = {};
  
  // Copy Column widths from template
  if (templateTraitSheet["!cols"]) outTraitSheet["!cols"] = [...templateTraitSheet["!cols"]];

  // Phase A: Copy Headers (Row 1) from template
  for (let c = templateRange.s.c; c <= templateRange.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    const cell = templateTraitSheet[addr];
    if (cell) outTraitSheet[addr] = { ...cell };
  }

  // Phase B: Collect Pattern Cells from Template Row 2
  const patternCells: Record<number, XLSX.CellObject> = {};
  for (let c = templateRange.s.c; c <= templateRange.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: FORMULA_ROW_IDX, c });
    const cell = templateTraitSheet[addr];
    if (cell) patternCells[c] = cell;
  }

  // Phase C: Generate Rows 2 to actualDataMaxRow + 1
  for (let dr = 0; dr <= actualDataMaxRow - 0; dr++) {
    // We want to generate a row for each row in the anchor data sheet.
    // Excel Row 1 is header. Data starts row 2.
    // If dr=1 (2nd row of data), output index=1.
    const outputRowIdx = dr + 1; 
    const outputExcelRow = outputRowIdx + 1;

    for (let c = templateRange.s.c; c <= templateRange.e.c; c++) {
      const pCell = patternCells[c];
      if (!pCell) continue;

      const outputAddr = XLSX.utils.encode_cell({ r: outputRowIdx, c });

      if (pCell.f) {
        // Apply formula mirroring
        const adjusted = adjustFormulaRow(pCell.f, FORMULA_EXCEL_ROW, outputExcelRow);
        outTraitSheet[outputAddr] = { ...pCell, f: adjusted, v: undefined };
      } else {
        // Copy static value/style
        outTraitSheet[outputAddr] = { ...pCell };
      }
    }
  }

  // Set reference range
  outTraitSheet["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: actualDataMaxRow + 1, c: templateRange.e.c }
  });

  XLSX.utils.book_append_sheet(newWb, outTraitSheet, "Traitement");

  console.log("Writing output workbook with mirrored 'Traitement' sheet...");
  const out = XLSX.write(newWb, { type: "array", bookType: "xlsx" });
  console.log("Finished.");
  
  return new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
};
