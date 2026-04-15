
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
 * Finds the sheet name in the current workbook that best matches a reference.
 * Ignores spaces, underscores, and case. Also handles minor typos (O vs D, etc.)
 */
function findBestSheetMatch(target: string, availableSheets: string[]): string {
  const normalize = (s: string) => s.replace(/[^A-Z0-9]/gi, "").toLowerCase();
  const normalizedTarget = normalize(target);
  
  // 1. Exact normalized match (ignores spaces/underscores)
  const exactMatch = availableSheets.find(s => normalize(s) === normalizedTarget);
  if (exactMatch) return exactMatch;

  // 2. Similarity match for typos (e.g. ROC vs RDC)
  // We check if the target is a substring or has very high similarity
  const similarityMatch = availableSheets.find(s => {
    const n = normalize(s);
    if (n.length !== normalizedTarget.length) return false;
    let diffs = 0;
    for (let i = 0; i < n.length; i++) {
      if (n[i] !== normalizedTarget[i]) diffs++;
    }
    return diffs <= 1; // Allow 1 character typo
  });

  return similarityMatch || target;
}

/**
 * Adjusts row references in an Excel formula and normalizes sheet names.
 */
function adjustFormulaRow(
  formula: string,
  fromExcelRow: number,
  toExcelRow: number,
  availableSheets: string[]
): string {
  let processed = formula;

  // 1. Handle the "Spill Operator" transformation.
  // SheetJS often reads 'A2#' as 'ANCHORARRAY(A2)'. We must convert it back.
  processed = processed.replace(/(?:_xlfn\._xlws\.)?ANCHORARRAY\(([^)]+)\)/g, "$1#");

  // 2. Normalize Sheet Names with typo tolerance
  const sheetRegex = /'([^']+)'!|([A-Za-z0-9._]+)!/g;
  processed = processed.replace(sheetRegex, (match, quoted, unquoted) => {
    const rawName = quoted || unquoted;
    const bestMatch = findBestSheetMatch(rawName, availableSheets);
    return bestMatch.includes(" ") ? `'${bestMatch}'!` : `${bestMatch}!`;
  });

  // 3. Adjust Row References
  const offset = toExcelRow - fromExcelRow;
  if (offset === 0) return processed;

  const rowRegex = /(^|[^A-Za-z])(\$?[A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_.\(])/g;

  return processed.replace(rowRegex, (match, prefix, col, dollar, row) => {
    if (dollar === "$") return match;
    const newRow = parseInt(row, 10) + offset;
    if (newRow < 1) return match;
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

  const templateTraitName = templateWb.SheetNames.includes("Traitement") ? "Traitement" : templateWb.SheetNames[0];
  const templateTraitSheet = templateWb.Sheets[templateTraitName];
  
  const dataAnchorName = dataWb.SheetNames[0];
  const dataAnchorSheet = dataWb.Sheets[dataAnchorName];

  if (!templateTraitSheet || !dataAnchorSheet) {
    throw new Error("Missing required sheets for processing.");
  }

  const templateRange = XLSX.utils.decode_range(templateTraitSheet["!ref"] || "A1");
  
  let actualDataMaxRow = 0;
  Object.keys(dataAnchorSheet).forEach(key => {
    if (key[0] === '!') return;
    const cell = XLSX.utils.decode_cell(key);
    if (cell.r > actualDataMaxRow) actualDataMaxRow = cell.r;
  });
  
  const FORMULA_ROW_IDX = 1; 
  const FORMULA_EXCEL_ROW = 2;

  const newWb = XLSX.utils.book_new();
  
  for (const sn of dataWb.SheetNames) {
    if (sn === "Traitement") continue;
    XLSX.utils.book_append_sheet(newWb, dataWb.Sheets[sn], sn);
  }

  const outTraitSheet: XLSX.WorkSheet = {};
  if (templateTraitSheet["!cols"]) outTraitSheet["!cols"] = [...templateTraitSheet["!cols"]];

  // Headers
  for (let c = templateRange.s.c; c <= templateRange.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    const cell = templateTraitSheet[addr];
    if (cell) outTraitSheet[addr] = { ...cell };
  }

  // Pattern Cells
  const patternCells: Record<number, XLSX.CellObject> = {};
  for (let c = templateRange.s.c; c <= templateRange.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: FORMULA_ROW_IDX, c });
    const cell = templateTraitSheet[addr];
    if (cell) patternCells[c] = cell;
  }

  const availableSheets = dataWb.SheetNames;

  // Phase C: Generate Rows
  for (let dr = 0; dr <= actualDataMaxRow; dr++) {
    const outputRowIdx = dr + 1; 
    const outputExcelRow = outputRowIdx + 1;

    for (let c = templateRange.s.c; c <= templateRange.e.c; c++) {
      const pCell = patternCells[c];
      if (!pCell) continue;

      const outputAddr = XLSX.utils.encode_cell({ r: outputRowIdx, c });

      if (pCell.f) {
        try {
          const adjusted = adjustFormulaRow(pCell.f, FORMULA_EXCEL_ROW, outputExcelRow, availableSheets);
          
          const newCell: XLSX.CellObject = { ...pCell, f: adjusted, v: undefined };
          
          // CRITICAL: If the formula is a Dynamic Array (FILTER, UNIQUE, VSTACK, etc.)
          // we must mark it as an Array Formula with a range (even if 1-cell) 
          // to prevent internal library mangling.
          const isModern = /FILTER|UNIQUE|VSTACK|SORT|SEQUENCE/i.test(adjusted);
          if (isModern) {
             // We mark it as an array formula for this specific cell.
             (newCell as any).F = outputAddr; 
          }

          outTraitSheet[outputAddr] = newCell;
        } catch (err) {
          outTraitSheet[outputAddr] = { ...pCell, f: pCell.f, v: undefined };
        }
      } else {
        outTraitSheet[outputAddr] = { ...pCell };
      }
    }
  }

  outTraitSheet["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
     e: { r: actualDataMaxRow + 1, c: templateRange.e.c }
  });

  XLSX.utils.book_append_sheet(newWb, outTraitSheet, "Traitement");

  const out = XLSX.write(newWb, { type: "array", bookType: "xlsx" });
  return new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
};
