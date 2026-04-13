import ExcelJS from "exceljs";

const SEMESTER_OPTIONS = ["2025H", "2026V", "2026H"];
const COLUMNS = [
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

async function buildWorkbook(schoolName, selectedSemesters) {
  const workbook = new ExcelJS.Workbook();
  const studentsSheet = workbook.addWorksheet("Students");
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

export async function POST(request) {
  try {
    const { schoolName, semesters } = await request.json();

    if (!schoolName || typeof schoolName !== "string" || !schoolName.trim()) {
      return Response.json(
        { error: "A valid institution name is required." },
        { status: 400 }
      );
    }

    if (
      !Array.isArray(semesters) ||
      semesters.length === 0 ||
      semesters.some((semester) => !SEMESTER_OPTIONS.includes(semester))
    ) {
      return Response.json(
        { error: "At least one valid semester is required." },
        { status: 400 }
      );
    }

    const workbook = await buildWorkbook(schoolName.trim(), semesters);
    const buffer = await workbook.xlsx.writeBuffer();
    const filename = `${schoolName.trim().replace(/[^a-z0-9]+/gi, "-").toLowerCase()}-template.xlsx`;

    return new Response(buffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filename}"`
      }
    });
  } catch {
    return Response.json(
      { error: "Unable to generate template." },
      { status: 500 }
    );
  }
}
