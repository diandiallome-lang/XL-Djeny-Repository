
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
 *
 * IMPORTANT — what this function must NOT do:
 *  - Do NOT remove _xlfn. or _xlws. prefixes. SheetJS uses them internally
 *    to identify modern functions (UNIQUE, FILTER, XLOOKUP, VSTACK…). Removing
 *    them makes Excel unable to resolve the function name → #NOM? / #NAME? error.
 *  - Do NOT transform ANCHORARRAY(...) into the # spill-range notation. SheetJS
 *    v0.18 cannot serialize the # operator back to valid OOXML XML, which causes
 *    Excel's repair dialog ("enregistrements supprimés / formule illisible").
 *    SheetJS knows how to write ANCHORARRAY back correctly — leave it alone.
 */
function adjustFormulaRow(
  formula: string,
  fromExcelRow: number,
  toExcelRow: number,
  availableSheets: string[]
): string {
  let processed = formula;

  // 1. Normalize argument separators.
  //    Proper .xlsx XML must use ',' but some files saved by non-English Excel
  //    incorrectly store ';'. We normalise to avoid parse errors.
  processed = processed.replace(/;/g, ",");

  // 2. Normalize sheet names referenced in the formula (typo tolerance).
  // Handle optional spaces before the ! (e.g. 'Sheet' !A1)
  const sheetRegex = /'([^']+)'\s*!|([A-Za-z0-9._]+)\s*!/g;
  processed = processed.replace(sheetRegex, (match, quoted, unquoted) => {
    const rawName = quoted || unquoted;
    const bestMatch = findBestSheetMatch(rawName, availableSheets);
    // Standardize to the correct ! notation without extra spaces
    return bestMatch.includes(" ") ? `'${bestMatch}'!` : `${bestMatch}!`;
  });

  // 3. Adjust relative row references.
  const offset = toExcelRow - fromExcelRow;
  if (offset === 0) return processed;

  // Matches: optional leading non-letter, column letters, optional $, row digits.
  // Negative lookahead prevents matching numbers inside function names.
  const rowRegex = /(^|[^A-Za-z])(\$?[A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_.(])/g;

  return processed.replace(rowRegex, (match, prefix, col, dollar, row) => {
    if (dollar === "$") return match; // Absolute row — keep as-is
    const newRow = parseInt(row, 10) + offset;
    if (newRow < 1) return match; // Guard against going above row 1
    return `${prefix}${col}${newRow}`;
  });
}

const readWorkbook = (file: File): Promise<XLSX.WorkBook> =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      // cellNF: true preserves number formatting (z code)
      resolve(XLSX.read(data, { type: "array", cellFormula: true, cellNF: true }));
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });

/**
 * Applies the formulas from a template file to a raw data file using a "Template-First" strategy.
 * This preserves the modern Excel metadata (Dynamic Array flags) for FILTER and VSTACK functions.
 */
export const applyTemplateToData = async (
  templateFile: File,
  dataFile: File
): Promise<Blob> => {
  console.log("Starting Template-First mirroring...");
  const templateWb = await readWorkbook(templateFile);
  const dataWb = await readWorkbook(dataFile);

  // 1. Identify the 'Traitement' sheet in the template.
  const traitSheetName = templateWb.SheetNames.find(sn => sn.toLowerCase().includes("trait")) || templateWb.SheetNames[0];
  const traitSheet = templateWb.Sheets[traitSheetName];
  
  if (!traitSheet) throw new Error("Could not find Traitement sheet in template.");

  // 2. Load the Data Anchor to determine how many rows to generate.
  const dataAnchorName = dataWb.SheetNames[0];
  const dataAnchorSheet = dataWb.Sheets[dataAnchorName];
  if (!dataAnchorSheet) throw new Error("Data file is empty or invalid.");

  let actualDataMaxRow = 0;
  Object.keys(dataAnchorSheet).forEach(key => {
    if (key[0] === '!') return;
    const cell = XLSX.utils.decode_cell(key);
    if (cell.r > actualDataMaxRow) actualDataMaxRow = cell.r;
  });

  // 3. Prepare the Result Workbook. 
  const resultWb = XLSX.utils.book_new();
  
  // A. First, append all sheets from the DATA file.
  for (const sn of dataWb.SheetNames) {
    if (sn === traitSheetName) continue; // Avoid name collision
    XLSX.utils.book_append_sheet(resultWb, dataWb.Sheets[sn], sn);
  }

  // B. Then, append the 'Traitement' sheet from the TEMPLATE.
  XLSX.utils.book_append_sheet(resultWb, traitSheet, traitSheetName);

  // 4. Mirroring Logic on the 'Traitement' sheet.
  const workTraitSheet = resultWb.Sheets[traitSheetName];
  const templateRange = XLSX.utils.decode_range(workTraitSheet["!ref"] || "A1");
  const FORMULA_ROW_IDX = 1; // Pattern row (Row 2)
  const FORMULA_EXCEL_ROW = 2;
  const availableSheets = dataWb.SheetNames;

  // Collect Pattern Cells from Template Row 2
  const patternCells: Record<number, XLSX.CellObject> = {};
  for (let c = templateRange.s.c; c <= templateRange.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: FORMULA_ROW_IDX, c });
    const cell = workTraitSheet[addr];
    if (cell) patternCells[c] = { ...cell };
  }

  // Generate additional rows
  for (let dr = 0; dr <= actualDataMaxRow; dr++) {
    const outputRowIdx = dr + 1; // Row 2, 3, 4...
    const outputExcelRow = outputRowIdx + 1;

    for (let c = templateRange.s.c; c <= templateRange.e.c; c++) {
      const pCell = patternCells[c];
      if (!pCell) continue;

      const outputAddr = XLSX.utils.encode_cell({ r: outputRowIdx, c });

      if (pCell.f) {
        try {
          const adjusted = adjustFormulaRow(pCell.f, FORMULA_EXCEL_ROW, outputExcelRow, availableSheets);
          // Preserve type (t) and format (z)
          const newCell: XLSX.CellObject = { f: adjusted, t: pCell.t, z: pCell.z };
          workTraitSheet[outputAddr] = newCell;
        } catch (err) {
          workTraitSheet[outputAddr] = { f: pCell.f, t: pCell.t, z: pCell.z };
        }
      } else {
        // Copy static cell with type and format
        workTraitSheet[outputAddr] = { ...pCell };
      }
    }
  }

  // Update sheet bounds
  workTraitSheet["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: actualDataMaxRow + 1, c: templateRange.e.c }
  });

  console.log("Writing Template-First output workbook...");
  const out = XLSX.write(resultWb, { type: "array", bookType: "xlsx" });
  return new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
};
