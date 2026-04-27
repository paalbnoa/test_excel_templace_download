import ExcelJS from "exceljs";
import { FIELD_RULES } from "./validation-rules.js";

export const SEMESTER_OPTIONS = ["2025H", "2026V", "2026H"];

export const COLUMNS = [
  { header: "PersonID", key: "personId", width: 18 },
  { header: "Fornavn", key: "fornavn", width: 20 },
  { header: "Etternavn", key: "etternavn", width: 22 },
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
const STUDENTS_TABLE_NAME = "StudentsTable";
const TEST_FIRST_NAMES = [
  "Anna",
  "Erik",
  "Maja",
  "Oskar",
  "Ingrid",
  "Emil",
  "Sara",
  "Henrik",
  "Nora",
  "Lars"
];
const TEST_LAST_NAMES = [
  "Hansen",
  "Johansen",
  "Berg",
  "Dahl",
  "Lunde",
  "Nilsen",
  "Aasen",
  "Lie",
  "Moen",
  "Solberg"
];
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

  worksheet.getCell("A3").value = "Institution short name:";
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
  worksheet.getCell("B3").protection = { locked: true };

  worksheet.getCell("A4").value = "Semester:";
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
  });

}

function addRowsButton(worksheet) {
  const buttonCell = worksheet.getCell("E3");
  buttonCell.value = "Add 100 new empty rows";
  buttonCell.font = {
    bold: true,
    color: { argb: "FFFFFFFF" },
    size: 11
  };
  buttonCell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1F7A5A" }
  };
  buttonCell.alignment = {
    vertical: "middle",
    horizontal: "center"
  };
  buttonCell.protection = { locked: true };

  worksheet.getRow(3).height = Math.max(worksheet.getRow(3).height || 0, 28);
}

function addStudentsTable(worksheet) {
  worksheet.addTable({
    name: STUDENTS_TABLE_NAME,
    displayName: STUDENTS_TABLE_NAME,
    ref: `A${HEADER_ROW}`,
    headerRow: true,
    totalsRow: false,
    style: {
      theme: null,
      showRowStripes: false
    },
    columns: COLUMNS.map(({ header }) => ({
      name: header,
      filterButton: true
    })),
    rows: Array.from({ length: EXTRA_READY_ROWS }, () => COLUMNS.map(() => ""))
  });
}

function buildDataValidation(fieldRule, cellReference) {
  return {
    type: fieldRule.excel.type,
    allowBlank: fieldRule.excel.allowBlank,
    showErrorMessage: true,
    errorStyle: "error",
    errorTitle: fieldRule.errorTitle,
    error: fieldRule.errorMessage,
    formulae: [fieldRule.excel.formula(cellReference)]
  };
}

function configureDataRowCells(worksheet, rowNumber) {
  const row = worksheet.getRow(rowNumber);

  COLUMNS.forEach((_, index) => {
    const cell = row.getCell(index + 1);
    cell.alignment = { vertical: "middle" };
    applyCellBorders(cell, "FFE4EAF1");
    cell.protection = { locked: false };

    if (index === 5) {
      cell.numFmt = "@";
    }
  });

  return row;
}

function applyColumnValidations(worksheet) {
  const columnValidations = [
    { rule: FIELD_RULES.PersonID, col: "A" },
    { rule: FIELD_RULES.Fornavn, col: "B" },
    { rule: FIELD_RULES.Etternavn, col: "C" },
    { rule: FIELD_RULES["Fritatt.sem.avg"], col: "D" },
    { rule: FIELD_RULES.Epost, col: "E" },
    { rule: FIELD_RULES.Prefiks, col: "F" },
    { rule: FIELD_RULES.Mobilnummer, col: "G" }
  ];

  columnValidations.forEach(({ rule, col }) => {
    const firstCell = `${col}${FIRST_DATA_ROW}`;
    const range = `${firstCell}:${col}${LAST_READY_ROW}`;
    worksheet.dataValidations.add(range, buildDataValidation(rule, firstCell));
  });
}

function buildTestRow(index) {
  const rowNumber = index + 1;
  const personId = 10000000000 + rowNumber;
  const mobileBase = 40000000 + rowNumber;
  const fornavn = TEST_FIRST_NAMES[index % TEST_FIRST_NAMES.length];
  const etternavn = TEST_LAST_NAMES[index % TEST_LAST_NAMES.length];

  return {
    personId,
    fornavn,
    etternavn,
    fritattSemAvg: rowNumber % 2 === 0 ? "Ja" : "Nei",
    epost: `test${rowNumber}@example.no`,
    prefiks: "+47",
    mobilnummer: String(mobileBase)
  };
}

function populateTestRow(row, index) {
  const sample = buildTestRow(index);

  row.getCell(1).value = sample.personId;
  row.getCell(2).value = sample.fornavn;
  row.getCell(3).value = sample.etternavn;
  row.getCell(4).value = sample.fritattSemAvg;
  row.getCell(5).value = sample.epost;
  row.getCell(6).value = sample.prefiks;
  row.getCell(7).value = sample.mobilnummer;
}

function pickRandomItems(items, count) {
  const shuffled = [...items];

  for (let index = shuffled.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(Math.random() * (index + 1));
    const currentValue = shuffled[index];
    shuffled[index] = shuffled[swapIndex];
    shuffled[swapIndex] = currentValue;
  }

  return shuffled.slice(0, count);
}

function applyRandomTestIssues(worksheet) {
  const rowNumbers = Array.from(
    { length: EXTRA_READY_ROWS },
    (_, index) => FIRST_DATA_ROW + index
  );
  const selectedRows = pickRandomItems(rowNumbers, 6);
  const invalidValueMutations = [
    (row) => {
      row.getCell(1).value = "ABC123";
    },
    (row) => {
      row.getCell(2).value = "Anna3";
    },
    (row) => {
      row.getCell(4).value = "Kanskje";
    },
    (row) => {
      row.getCell(5).value = "invalid-email";
    },
    (row) => {
      row.getCell(6).value = "0047";
    },
    (row) => {
      row.getCell(7).value = "12A45678";
    }
  ];
  const emptyValueMutations = [
    (row) => {
      row.getCell(1).value = "";
    },
    (row) => {
      row.getCell(2).value = "";
    },
    (row) => {
      row.getCell(3).value = "";
    },
    (row) => {
      row.getCell(5).value = "";
    },
    (row) => {
      row.getCell(6).value = "";
    },
    (row) => {
      row.getCell(7).value = "";
    }
  ];

  pickRandomItems(invalidValueMutations, 3).forEach((mutation, index) => {
    mutation(worksheet.getRow(selectedRows[index]));
  });

  pickRandomItems(emptyValueMutations, 3).forEach((mutation, index) => {
    mutation(worksheet.getRow(selectedRows[index + 3]));
  });
}

function configureDataSheet(worksheet, schoolName, selectedSemesters, options = {}) {
  worksheet.columns = COLUMNS;

  addMetadataFields(worksheet, schoolName, selectedSemesters);
  worksheet.spliceRows(1, 1);
  addRowsButton(worksheet);
  addStudentsTable(worksheet);
  configureHeaderRow(worksheet);

  for (let rowNumber = FIRST_DATA_ROW; rowNumber <= LAST_READY_ROW; rowNumber += 1) {
    const row = configureDataRowCells(worksheet, rowNumber);

    if (options.includeTestData) {
      populateTestRow(row, rowNumber - FIRST_DATA_ROW);
    }
  }

  applyColumnValidations(worksheet);

  if (options.includeTestData && options.includeRandomErrors) {
    applyRandomTestIssues(worksheet);
  }

  const lastCol = String.fromCharCode(64 + COLUMNS.length);
  worksheet.addConditionalFormatting({
    ref: `A${FIRST_DATA_ROW}:${lastCol}${FIRST_DATA_ROW + 10000}`,
    rules: [
      {
        type: "expression",
        formulae: [`=MOD(ROW(),2)=1`],
        style: {
          fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } }
        },
        priority: 1
      },
    ]
  });

  worksheet.views = [{ state: "frozen", ySplit: HEADER_ROW, showGridLines: false }];
}

function configureInstructionSheet(worksheet, selectedSemesters, options = {}) {
  worksheet.columns = [{ width: 110 }];

  const instructionRows = [
    "Instructions",
    "This workbook contains a template for importing data about paid semester fees.",
    "",
    "Validation rules:",
    "1. All columns are mandatory except Fritatt.sem.avg.",
    `2. ${FIELD_RULES.PersonID.instruction}`,
    `3. Semester: The selected semester for this workbook is shown above the table: ${selectedSemesters.join(", ")}.`,
    `4. ${FIELD_RULES.Fornavn.instruction}`,
    `5. ${FIELD_RULES.Etternavn.instruction}`,
    `6. ${FIELD_RULES["Fritatt.sem.avg"].instruction}`,
    `7. ${FIELD_RULES.Epost.instruction}`,
    `8. ${FIELD_RULES.Prefiks.instruction}`,
    `9. ${FIELD_RULES.Mobilnummer.instruction}`,
    "",
    "Other notes:",
    "1. The table contains 100 blank rows ready for data entry."
  ];

  if (options.includeTestData) {
    instructionRows.push("2. This download includes 100 rows of sample data.");
  }

  if (options.includeTestData && options.includeRandomErrors) {
    instructionRows.push(
      "3. A few random rows intentionally contain invalid values or blank required cells for testing."
    );
  }

  worksheet.views = [{ showGridLines: false }];

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
  worksheet.getCell("A15").font = { bold: true };
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
  const rawSemesters = cellDisplayValue(worksheet.getCell("B4"));

  if (!rawSemesters) {
    return [];
  }

  return rawSemesters
    .split(",")
    .map((semester) => semester.trim())
    .filter((semester) => SEMESTER_OPTIONS.includes(semester));
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

export function getHeaderNotes() {
  return COLUMNS.map(({ header }, index) => {
    const rule = FIELD_RULES[header];
    return rule ? {
      cellRef: `${String.fromCharCode(65 + index)}${HEADER_ROW}`,
      text: rule.instruction,
      row: HEADER_ROW - 1,
      col: index
    } : null;
  }).filter(Boolean);
}

export async function buildWorkbook(schoolName, selectedSemesters, options = {}) {
  const workbook = new ExcelJS.Workbook();
  const studentsSheet = workbook.addWorksheet(STUDENTS_SHEET_NAME);
  const instructionSheet = workbook.addWorksheet("Instructions");

  configureDataSheet(studentsSheet, schoolName, selectedSemesters, options);
  configureInstructionSheet(instructionSheet, selectedSemesters, options);
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

    if (!FIELD_RULES.PersonID.validate(row.getCell(1).value)) {
      errors.push({
        rowNumber,
        column: "PersonID",
        message: FIELD_RULES.PersonID.errorMessage
      });
    }

    if (!FIELD_RULES.Fornavn.validate(row.getCell(2).value)) {
      errors.push({
        rowNumber,
        column: "Fornavn",
        message: FIELD_RULES.Fornavn.errorMessage
      });
    }

    if (!FIELD_RULES.Etternavn.validate(row.getCell(3).value)) {
      errors.push({
        rowNumber,
        column: "Etternavn",
        message: FIELD_RULES.Etternavn.errorMessage
      });
    }

    if (!FIELD_RULES["Fritatt.sem.avg"].validate(row.getCell(4).value)) {
      errors.push({
        rowNumber,
        column: "Fritatt.sem.avg",
        message: FIELD_RULES["Fritatt.sem.avg"].errorMessage
      });
    }

    if (!FIELD_RULES.Epost.validate(row.getCell(5).value)) {
      errors.push({
        rowNumber,
        column: "Epost",
        message: FIELD_RULES.Epost.errorMessage
      });
    }

    if (!FIELD_RULES.Prefiks.validate(row.getCell(6).value)) {
      errors.push({
        rowNumber,
        column: "Prefiks",
        message: FIELD_RULES.Prefiks.errorMessage
      });
    }

    if (!FIELD_RULES.Mobilnummer.validate(row.getCell(7).value)) {
      errors.push({
        rowNumber,
        column: "Mobilnummer",
        message: FIELD_RULES.Mobilnummer.errorMessage
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
