import ExcelJS from "exceljs";

export const SEMESTER_OPTIONS = ["2025H", "2026V", "2026H"];

export const COLUMNS = [
  { header: "PersonID", key: "personId", width: 18 },
  { header: "Fornavn", key: "fornavn", width: 20 },
  { header: "Etternavn", key: "etternavn", width: 22 },
  { header: "Sist.bet.sem.avg", key: "sistBetSemAvg", width: 18 },
  { header: "Fritatt.sem.avg", key: "fritattSemAvg", width: 18 },
  { header: "Epost", key: "epost", width: 30 },
  { header: "Prefiks", key: "prefiks", width: 14 },
  { header: "Mobilnummer", key: "mobilnummer", width: 18 }
];

const EXTRA_READY_ROWS = 100;
const HEADER_ROW = 6;
const FIRST_DATA_ROW = HEADER_ROW + 1;
const LAST_READY_ROW = FIRST_DATA_ROW + EXTRA_READY_ROWS - 1;
const STUDENTS_SHEET_NAME = "Students";
const COLUMN_INDEX_BY_HEADER = Object.fromEntries(
  COLUMNS.map(({ header }, index) => [header, index + 1])
);

function applyCellBorders(cell, color) {
  cell.border = {
    top: { style: "thin", color: { argb: color } },
    left: { style: "thin", color: { argb: color } },
    bottom: { style: "thin", color: { argb: color } },
    right: { style: "thin", color: { argb: color } }
  };
}

function addMetadataFields(worksheet, schoolName, selectedSemesters) {
  worksheet.getRow(1).height = 18;
  worksheet.getRow(3).height = 24;
  worksheet.getRow(4).height = 28;
  worksheet.getRow(5).height = 18;

  worksheet.getCell("A3").value = "Institusjon:";
  worksheet.getCell("A3").font = {
    bold: true,
    color: { argb: "FF123B63" },
    size: 11
  };
  worksheet.getCell("A3").alignment = { vertical: "middle" };
  worksheet.getCell("A3").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF5F8FB" }
  };
  applyCellBorders(worksheet.getCell("A3"), "FFD5DDE7");

  worksheet.getCell("B3").value = schoolName;
  worksheet.getCell("B3").font = {
    bold: true,
    size: 13,
    color: { argb: "FF132033" }
  };
  worksheet.getCell("B3").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFEAF2F8" }
  };
  worksheet.getCell("B3").alignment = { vertical: "middle", horizontal: "left" };
  applyCellBorders(worksheet.getCell("B3"), "FFD5DDE7");
  worksheet.getCell("B3").protection = { locked: true };

  worksheet.getCell("A4").value = "Semestre:";
  worksheet.getCell("A4").font = {
    bold: true,
    color: { argb: "FF123B63" },
    size: 11
  };
  worksheet.getCell("A4").alignment = { vertical: "middle" };
  worksheet.getCell("A4").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF5F8FB" }
  };
  applyCellBorders(worksheet.getCell("A4"), "FFD5DDE7");

  worksheet.getCell("B4").value = selectedSemesters.join(", ");
  worksheet.getCell("B4").font = {
    bold: true,
    size: 12,
    color: { argb: "FF132033" }
  };
  worksheet.getCell("B4").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFEAF2F8" }
  };
  worksheet.getCell("B4").alignment = { vertical: "middle", horizontal: "left", wrapText: true };
  applyCellBorders(worksheet.getCell("B4"), "FFD5DDE7");
  worksheet.getCell("B4").protection = { locked: true };
}

function configureHeaderRow(worksheet) {
  worksheet.getRow(HEADER_ROW).height = 24;
  worksheet.getRow(HEADER_ROW).eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF123B63" }
    };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    applyCellBorders(cell, "FFD5DDE7");
  });
}

function buildSemesterValidationFormula(selectedSemesters) {
  return `"${selectedSemesters.join(",")}"`;
}

function configureDataSheet(worksheet, schoolName, selectedSemesters) {
  worksheet.columns = COLUMNS;

  addMetadataFields(worksheet, schoolName, selectedSemesters);
  worksheet.spliceRows(1, 1);

  COLUMNS.forEach(({ header }, index) => {
    const columnLetter = worksheet.getColumn(index + 1).letter;
    worksheet.getCell(`${columnLetter}${HEADER_ROW}`).value = header;
  });
  configureHeaderRow(worksheet);

  for (let rowNumber = FIRST_DATA_ROW; rowNumber <= LAST_READY_ROW; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);

    COLUMNS.forEach((_, index) => {
      const cell = row.getCell(index + 1);
      cell.alignment = { vertical: "middle" };
      applyCellBorders(cell, "FFE4EAF1");
      cell.protection = { locked: false };
    });

    if (rowNumber % 2 === 1) {
      COLUMNS.forEach((_, index) => {
        const cell = row.getCell(index + 1);
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF8FAFC" }
        };
      });
    }

    worksheet.getCell(`A${rowNumber}`).dataValidation = {
      type: "whole",
      operator: "greaterThanOrEqual",
      allowBlank: false,
      showErrorMessage: true,
      errorStyle: "error",
      errorTitle: "Invalid PersonID",
      error: "PersonID is required and must contain whole numbers only.",
      formulae: [0]
    };

    worksheet.getCell(`B${rowNumber}`).dataValidation = {
      type: "custom",
      allowBlank: false,
      showErrorMessage: true,
      errorStyle: "error",
      errorTitle: "Missing Fornavn",
      error: "Fornavn is required.",
      formulae: [`LEN(TRIM(B${rowNumber}))>0`]
    };

    worksheet.getCell(`C${rowNumber}`).dataValidation = {
      type: "custom",
      allowBlank: false,
      showErrorMessage: true,
      errorStyle: "error",
      errorTitle: "Missing Etternavn",
      error: "Etternavn is required.",
      formulae: [`LEN(TRIM(C${rowNumber}))>0`]
    };

    worksheet.getCell(`D${rowNumber}`).dataValidation = {
      type: "list",
      allowBlank: false,
      showErrorMessage: true,
      errorStyle: "error",
      errorTitle: "Invalid semester",
      error: "Sist.bet.sem.avg is required and must match one of the allowed semester values.",
      formulae: [buildSemesterValidationFormula(selectedSemesters)]
    };

    worksheet.getCell(`F${rowNumber}`).dataValidation = {
      type: "custom",
      allowBlank: false,
      showErrorMessage: true,
      errorStyle: "error",
      errorTitle: "Invalid email",
      error: "Epost is required and must be a valid email address.",
      formulae: [
        `AND(LEN(TRIM(F${rowNumber}))>0,ISNUMBER(SEARCH("@",F${rowNumber})),ISNUMBER(SEARCH(".",F${rowNumber},SEARCH("@",F${rowNumber})+2)),LEN(F${rowNumber})-LEN(SUBSTITUTE(F${rowNumber},"@",""))=1)`
      ]
    };

    worksheet.getCell(`G${rowNumber}`).dataValidation = {
      type: "custom",
      allowBlank: false,
      showErrorMessage: true,
      errorStyle: "error",
      errorTitle: "Missing Prefiks",
      error: "Prefiks is required.",
      formulae: [`LEN(TRIM(G${rowNumber}))>0`]
    };

    worksheet.getCell(`H${rowNumber}`).dataValidation = {
      type: "custom",
      allowBlank: false,
      showErrorMessage: true,
      errorStyle: "error",
      errorTitle: "Invalid phone number",
      error: "Mobilnummer is required and must be an 8-digit Norwegian mobile number without country code.",
      formulae: [
        `AND(LEN(H${rowNumber})=8,ISNUMBER(H${rowNumber}),OR(LEFT(H${rowNumber},1)="4",LEFT(H${rowNumber},1)="9"))`
      ]
    };
  }

  worksheet.autoFilter = `A${HEADER_ROW}:H${LAST_READY_ROW}`;
  worksheet.views = [{ state: "frozen", ySplit: HEADER_ROW }];
}

function configureInstructionSheet(worksheet, selectedSemesters) {
  worksheet.columns = [{ width: 110 }];

  const instructionRows = [
    "Instructions",
    "This workbook contains a template for importing data about paid semester fees.",
    "",
    "Validation rules:",
    "1. All columns are mandatory except Fritatt.sem.avg.",
    "2. PersonID: This field is required and only accepts whole numbers. Letters and other characters are not allowed.",
    `3. Sist.bet.sem.avg: This field is required and may only contain one of the semester values selected for this download: ${selectedSemesters.join(", ")}.`,
    "4. Epost: This field is required and must look like a valid email address and contain one @ sign and a period after the @ sign.",
    "5. Mobilnummer: This field is required and must be a Norwegian mobile number with 8 digits, without a country code, and it must start with 4 or 9.",
    "",
    "Other notes:",
    "1. The table contains 100 blank rows ready for data entry."
  ];

  instructionRows.forEach((text, index) => {
    const cell = worksheet.getCell(`A${index + 1}`);
    cell.value = text;
    cell.alignment = { vertical: "top", wrapText: true };
    cell.font = { bold: false };
  });

  worksheet.getCell("A1").font = {
    bold: true,
    size: 16,
    color: { argb: "FF123B63" }
  };
  worksheet.getCell("A4").font = { bold: true };
  worksheet.getCell("A10").font = { bold: true };
}

function normalizeText(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}

function normalizeNumberLikeText(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return typeof value === "number" ? String(value) : String(value).trim();
}

function cellDisplayValue(cell) {
  const { value } = cell;

  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "object") {
    if ("text" in value && value.text) {
      return String(value.text).trim();
    }

    if ("result" in value && value.result !== undefined && value.result !== null) {
      return String(value.result).trim();
    }

    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text).join("").trim();
    }
  }

  return String(value).trim();
}

function hasMeaningfulCellValue(cell) {
  const { value } = cell;

  if (value === null || value === undefined) {
    return false;
  }

  if (typeof value === "string") {
    return value.trim().length > 0;
  }

  if (typeof value === "number" || typeof value === "boolean") {
    return true;
  }

  if (value instanceof Date) {
    return true;
  }

  if (typeof value === "object") {
    if ("text" in value) {
      return String(value.text || "").trim().length > 0;
    }

    if ("result" in value) {
      return value.result !== null && value.result !== undefined && String(value.result).trim() !== "";
    }

    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText.some((part) => String(part.text || "").trim().length > 0);
    }

    if ("formula" in value) {
      return false;
    }
  }

  return String(value).trim().length > 0;
}

function rowHasUserInput(row) {
  return COLUMNS.some((_, index) => hasMeaningfulCellValue(row.getCell(index + 1)));
}

function isRowEmpty(row) {
  return !rowHasUserInput(row);
}

function findHeaderRow(worksheet) {
  for (let rowNumber = 1; rowNumber <= Math.min(worksheet.rowCount, 20); rowNumber += 1) {
    const matches = COLUMNS.every(({ header }, index) => {
      return cellDisplayValue(worksheet.getRow(rowNumber).getCell(index + 1)) === header;
    });

    if (matches) {
      return rowNumber;
    }
  }

  return null;
}

function parseSelectedSemesters(worksheet) {
  const rawSemesters = cellDisplayValue(worksheet.getCell("B3"));

  if (!rawSemesters) {
    return [];
  }

  return rawSemesters
    .split(",")
    .map((semester) => semester.trim())
    .filter((semester) => SEMESTER_OPTIONS.includes(semester));
}

function validatePersonId(value) {
  return /^\d+$/.test(normalizeNumberLikeText(value));
}

function validateRequiredText(value) {
  return normalizeText(value).length > 0;
}

function validateSemester(value, allowedSemesters) {
  const normalized = normalizeText(value);
  return normalized.length > 0 && allowedSemesters.includes(normalized);
}

function validateEmail(value) {
  const normalized = normalizeText(value);

  if (!normalized) {
    return false;
  }

  const parts = normalized.split("@");
  if (parts.length !== 2) {
    return false;
  }

  const [localPart, domainPart] = parts;
  if (!localPart || !domainPart) {
    return false;
  }

  const dotIndex = domainPart.indexOf(".");
  return dotIndex > 0 && dotIndex < domainPart.length - 1;
}

function validatePhoneNumber(value) {
  return /^[49]\d{7}$/.test(normalizeNumberLikeText(value));
}

function applyValidationHighlight(cell, message) {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFE3E3" }
  };
  cell.font = {
    ...(cell.font || {}),
    color: { argb: "FF991B1B" }
  };
  cell.border = {
    top: { style: "medium", color: { argb: "FFDC2626" } },
    left: { style: "medium", color: { argb: "FFDC2626" } },
    bottom: { style: "medium", color: { argb: "FFDC2626" } },
    right: { style: "medium", color: { argb: "FFDC2626" } }
  };
}

function cloneCellValue(value) {
  if (value === null || value === undefined) {
    return value;
  }

  if (typeof structuredClone === "function") {
    return structuredClone(value);
  }

  if (typeof value === "object") {
    return JSON.parse(JSON.stringify(value));
  }

  return value;
}

function copyCell(sourceCell, targetCell) {
  targetCell.value = cloneCellValue(sourceCell.value);
  targetCell.style = cloneCellValue(sourceCell.style || {});
  targetCell.numFmt = sourceCell.numFmt;
  targetCell.alignment = cloneCellValue(sourceCell.alignment || undefined);
  targetCell.font = cloneCellValue(sourceCell.font || undefined);
  targetCell.fill = cloneCellValue(sourceCell.fill || undefined);
  targetCell.border = cloneCellValue(sourceCell.border || undefined);
  targetCell.protection = cloneCellValue(sourceCell.protection || undefined);
  targetCell.dataValidation = cloneCellValue(sourceCell.dataValidation || undefined);
  targetCell.note = undefined;
  targetCell._comment = undefined;

  if (targetCell.model && targetCell.model.comment) {
    delete targetCell.model.comment;
  }
}

function copyWorksheetRow(sourceWorksheet, targetWorksheet, sourceRowNumber, targetRowNumber) {
  const sourceRow = sourceWorksheet.getRow(sourceRowNumber);
  const targetRow = targetWorksheet.getRow(targetRowNumber);

  targetRow.height = sourceRow.height;

  for (let columnNumber = 1; columnNumber <= sourceWorksheet.columnCount; columnNumber += 1) {
    copyCell(sourceRow.getCell(columnNumber), targetRow.getCell(columnNumber));
  }
}

function clearWorksheetNotes(worksheet) {
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);

    for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
      const cell = row.getCell(columnNumber);
      cell.note = undefined;
      cell._comment = undefined;

      if (cell.model && cell.model.comment) {
        delete cell.model.comment;
      }
    }
  }

  worksheet.comments = [];
}

function copyWorksheetColumns(sourceWorksheet, targetWorksheet) {
  for (let columnNumber = 1; columnNumber <= sourceWorksheet.columnCount; columnNumber += 1) {
    const sourceColumn = sourceWorksheet.getColumn(columnNumber);
    const targetColumn = targetWorksheet.getColumn(columnNumber);

    targetColumn.width = sourceColumn.width;
    targetColumn.hidden = sourceColumn.hidden;
  }
}

export async function buildWorkbook(schoolName, selectedSemesters) {
  const workbook = new ExcelJS.Workbook();
  const studentsSheet = workbook.addWorksheet(STUDENTS_SHEET_NAME);
  const instructionSheet = workbook.addWorksheet("Instructions");

  configureDataSheet(studentsSheet, schoolName, selectedSemesters);
  configureInstructionSheet(instructionSheet, selectedSemesters);
  await studentsSheet.protect("", {
    selectLockedCells: false,
    selectUnlockedCells: true,
    formatCells: false,
    formatColumns: false,
    formatRows: false,
    insertColumns: false,
    insertRows: true,
    deleteColumns: false,
    deleteRows: false,
    sort: true,
    autoFilter: true
  });

  return workbook;
}

export async function validateWorkbookBuffer(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.getWorksheet(STUDENTS_SHEET_NAME) || workbook.worksheets[0];

  if (!worksheet) {
    return {
      isValid: false,
      summary: { rowCount: 0, errorCount: 1, warningCount: 0 },
      semesters: [],
      warnings: [],
      errors: [
        {
          rowNumber: null,
          column: null,
          message: "The workbook does not contain a worksheet to validate."
        }
      ]
    };
  }

  const headerRow = findHeaderRow(worksheet);

  if (!headerRow) {
    return {
      isValid: false,
      summary: { rowCount: 0, errorCount: 1, warningCount: 0 },
      semesters: [],
      warnings: [],
      errors: [
        {
          rowNumber: null,
          column: null,
          message: "Could not find the expected template headers in the uploaded workbook."
        }
      ]
    };
  }

  const metadataSemesters = parseSelectedSemesters(worksheet);
  const allowedSemesters = metadataSemesters.length > 0 ? metadataSemesters : SEMESTER_OPTIONS;
  const warnings = [];
  const errors = [];
  let rowCount = 0;

  const populatedRows = [];

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber > headerRow) {
      populatedRows.push({ row, rowNumber });
    }
  });

  for (const { row, rowNumber } of populatedRows) {

    if (!rowHasUserInput(row)) {
      continue;
    }

    rowCount += 1;

    if (!validatePersonId(row.getCell(1).value)) {
      errors.push({
        rowNumber,
        column: "PersonID",
        message: "PersonID is required and must contain whole numbers only."
      });
    }

    if (!validateRequiredText(row.getCell(2).value)) {
      errors.push({
        rowNumber,
        column: "Fornavn",
        message: "Fornavn is required."
      });
    }

    if (!validateRequiredText(row.getCell(3).value)) {
      errors.push({
        rowNumber,
        column: "Etternavn",
        message: "Etternavn is required."
      });
    }

    if (!validateSemester(row.getCell(4).value, allowedSemesters)) {
      errors.push({
        rowNumber,
        column: "Sist.bet.sem.avg",
        message: `Sist.bet.sem.avg is required and must match one of: ${allowedSemesters.join(", ")}.`
      });
    }

    if (!validateEmail(row.getCell(6).value)) {
      errors.push({
        rowNumber,
        column: "Epost",
        message: "Epost is required and must be a valid email address."
      });
    }

    if (!validateRequiredText(row.getCell(7).value)) {
      errors.push({
        rowNumber,
        column: "Prefiks",
        message: "Prefiks is required."
      });
    }

    if (!validatePhoneNumber(row.getCell(8).value)) {
      errors.push({
        rowNumber,
        column: "Mobilnummer",
        message: "Mobilnummer is required and must be an 8-digit Norwegian mobile number without country code."
      });
    }
  }

  if (rowCount === 0) {
    errors.push({
      rowNumber: null,
      column: null,
      message: "The uploaded workbook does not contain any data rows in the table."
    });
  }

  return {
    isValid: errors.length === 0,
    summary: {
      rowCount,
      errorCount: errors.length,
      warningCount: warnings.length
    },
    semesters: allowedSemesters,
    warnings,
    errors
  };
}

export async function buildHighlightedWorkbookBuffer(buffer, errors) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.getWorksheet(STUDENTS_SHEET_NAME) || workbook.worksheets[0];

  if (!worksheet) {
    return null;
  }

  const cleanedWorkbook = new ExcelJS.Workbook();
  const rowNumberMap = new Map();

  workbook.worksheets.forEach((sourceWorksheet) => {
    const targetWorksheet = cleanedWorkbook.addWorksheet(sourceWorksheet.name);
    copyWorksheetColumns(sourceWorksheet, targetWorksheet);
    targetWorksheet.views = cloneCellValue(sourceWorksheet.views || []);
    targetWorksheet.autoFilter = cloneCellValue(sourceWorksheet.autoFilter || undefined);
    targetWorksheet.properties = cloneCellValue(sourceWorksheet.properties || {});
    targetWorksheet.pageSetup = cloneCellValue(sourceWorksheet.pageSetup || {});
    targetWorksheet.state = sourceWorksheet.state;

    if (sourceWorksheet.id !== worksheet.id) {
      for (let rowNumber = 1; rowNumber <= sourceWorksheet.rowCount; rowNumber += 1) {
        copyWorksheetRow(sourceWorksheet, targetWorksheet, rowNumber, rowNumber);
      }
      return;
    }

    let nextTargetRowNumber = 1;

    for (let rowNumber = 1; rowNumber <= sourceWorksheet.rowCount; rowNumber += 1) {
      const sourceRow = sourceWorksheet.getRow(rowNumber);
      const shouldKeepRow =
        rowNumber <= HEADER_ROW || !isRowEmpty(sourceRow);

      if (!shouldKeepRow) {
        continue;
      }

      copyWorksheetRow(sourceWorksheet, targetWorksheet, rowNumber, nextTargetRowNumber);
      rowNumberMap.set(rowNumber, nextTargetRowNumber);
      nextTargetRowNumber += 1;
    }
  });

  const cleanedWorksheet =
    cleanedWorkbook.getWorksheet(STUDENTS_SHEET_NAME) || cleanedWorkbook.worksheets[0];

  if (!cleanedWorksheet) {
    return null;
  }

  cleanedWorkbook.worksheets.forEach((sheet) => {
    clearWorksheetNotes(sheet);
  });

  errors.forEach((item) => {
    if (!item.rowNumber || !item.column) {
      return;
    }

    const columnIndex = COLUMN_INDEX_BY_HEADER[item.column];

    if (!columnIndex) {
      return;
    }

    const shiftedRowNumber = rowNumberMap.get(item.rowNumber);

    if (!shiftedRowNumber) {
      return;
    }

    const shiftedRow = cleanedWorksheet.getRow(shiftedRowNumber);

    if (isRowEmpty(shiftedRow)) {
      return;
    }

    const cell = shiftedRow.getCell(columnIndex);
    applyValidationHighlight(cell, item.message);
  });

  return cleanedWorkbook.xlsx.writeBuffer();
}
