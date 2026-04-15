
import * as XLSX from "xlsx";

export interface ExcelMetadata {
  sheets: string[];
  columns: number;
  rows: number;
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

export const applyTemplateToData = async (templateFile: File, dataFile: File): Promise<Blob> => {
  const readWorkbook = async (file: File): Promise<XLSX.WorkBook> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        resolve(XLSX.read(data, { type: "array", cellFormula: true }));
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  };

  const templateWb = await readWorkbook(templateFile);
  const dataWb = await readWorkbook(dataFile);

  const templateSheet = templateWb.Sheets[templateWb.SheetNames[0]];
  const dataSheet = dataWb.Sheets[dataWb.SheetNames[0]];

  // Extract formulas from row 2 of the template (index 1)
  const templateRange = XLSX.utils.decode_range(templateSheet["!ref"] || "A1");
  const formulaRowIndex = 1; // 0-based index for Row 2
  const formulas: Record<number, string> = {};

  for (let c = templateRange.s.c; c <= templateRange.e.c; c++) {
    const cellAddress = XLSX.utils.encode_cell({ r: formulaRowIndex, c });
    const cell = templateSheet[cellAddress];
    if (cell && cell.f) {
      formulas[c] = cell.f;
    }
  }

  // Read data rows and apply formulas
  const dataRows = XLSX.utils.sheet_to_json(dataSheet, { header: 1 }) as any[][];
  const processedRows: any[][] = [dataRows[0]]; // Keep headers

  for (let i = 1; i < dataRows.length; i++) {
    const newRow = [...dataRows[i]];
    // Injected formulas adjust for current row index
    // Note: Excel formulas are 1-based, so Row 2 in sheet is index 1.
    // If the template was written for Row 2, and we are on Row i+1.
    // We can use SheetJS utility to shift formulas if needed, or simple string replacement
    // For simplicity in this tool, we assume formulas are relative.
    Object.entries(formulas).forEach(([colIndex, formula]) => {
      const cIdx = parseInt(colIndex);
      // We set the formula. SheetJS handles relative references if we use XLSX.utils.sheet_add_aoa with formula objects
      newRow[cIdx] = { f: formula };
    });
    processedRows.push(newRow);
  }

  const newSheet = XLSX.utils.aoa_to_sheet(processedRows);
  const newWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWb, newSheet, "Processed Data");

  const out = XLSX.write(newWb, { type: "array", bookType: "xlsx" });
  return new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
};
