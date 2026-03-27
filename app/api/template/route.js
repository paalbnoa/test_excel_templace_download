import ExcelJS from "exceljs";

const HEADERS = ["Name", "email", "fnr", "Phone number", "Semester fee code"];

const DUMMY_ROWS = [
  ["Emma Hansen", "emma.hansen@example.com", "12039812345", "+47 901 12 301", "SF-1001"],
  ["Liam Johansen", "liam.johansen@example.com", "23049712345", "+47 901 12 302", "SF-1002"],
  ["Olivia Larsen", "olivia.larsen@example.com", "14089612345", "+47 901 12 303", "SF-1003"],
  ["Noah Berg", "noah.berg@example.com", "05119512345", "+47 901 12 304", "SF-1004"],
  ["Sofie Nilsen", "sofie.nilsen@example.com", "16029412345", "+47 901 12 305", "SF-1005"],
  ["Jakob Andersen", "jakob.andersen@example.com", "27039312345", "+47 901 12 306", "SF-1006"],
  ["Ella Pedersen", "ella.pedersen@example.com", "18059212345", "+47 901 12 307", "SF-1007"],
  ["William Solberg", "william.solberg@example.com", "09069112345", "+47 901 12 308", "SF-1008"],
  ["Nora Kristiansen", "nora.kristiansen@example.com", "20079012345", "+47 901 12 309", "SF-1009"],
  ["Theodor Eriksen", "theodor.eriksen@example.com", "11088912345", "+47 901 12 310", "SF-1010"],
  ["Leah Dahl", "leah.dahl@example.com", "22018812345", "+47 901 12 311", "SF-1011"],
  ["Henrik Hauge", "henrik.hauge@example.com", "13028712345", "+47 901 12 312", "SF-1012"],
  ["Maja Strand", "maja.strand@example.com", "24038612345", "+47 901 12 313", "SF-1013"],
  ["Magnus Moe", "magnus.moe@example.com", "15048512345", "+47 901 12 314", "SF-1014"],
  ["Ingrid Lie", "ingrid.lie@example.com", "26058412345", "+47 901 12 315", "SF-1015"],
  ["Lucas Eide", "lucas.eide@example.com", "17068312345", "+47 901 12 316", "SF-1016"],
  ["Sara Halvorsen", "sara.halvorsen@example.com", "28078212345", "+47 901 12 317", "SF-1017"],
  ["Aksel Sunde", "aksel.sunde@example.com", "19088112345", "+47 901 12 318", "SF-1018"],
  ["Amalie Bakke", "amalie.bakke@example.com", "30098012345", "+47 901 12 319", "SF-1019"],
  ["Benjamin Lund", "benjamin.lund@example.com", "21117912345", "+47 901 12 320", "SF-1020"]
];

const EXTRA_READY_ROWS = 100;
const HEADER_ROW = 6;
const FIRST_DATA_ROW = HEADER_ROW + 1;
const LAST_READY_ROW = FIRST_DATA_ROW + DUMMY_ROWS.length + EXTRA_READY_ROWS - 1;

function applyCellBorders(cell, color) {
  cell.border = {
    top: { style: "thin", color: { argb: color } },
    left: { style: "thin", color: { argb: color } },
    bottom: { style: "thin", color: { argb: color } },
    right: { style: "thin", color: { argb: color } }
  };
}

function addSchoolField(worksheet, schoolName) {
  worksheet.getRow(1).height = 18;
  worksheet.getRow(3).height = 24;
  worksheet.getRow(4).height = 28;
  worksheet.getRow(5).height = 18;

  worksheet.getCell("A3").value = "School name:";
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

function configureDataSheet(worksheet, schoolName) {
  worksheet.columns = [
    { header: HEADERS[0], key: "name", width: 24 },
    { header: HEADERS[1], key: "email", width: 30 },
    { header: HEADERS[2], key: "fnr", width: 18 },
    { header: HEADERS[3], key: "phoneNumber", width: 20 },
    { header: HEADERS[4], key: "semesterFeeCode", width: 20 }
  ];

  addSchoolField(worksheet, schoolName);
  worksheet.spliceRows(1, 1);

  worksheet.getCell(`A${HEADER_ROW}`).value = HEADERS[0];
  worksheet.getCell(`B${HEADER_ROW}`).value = HEADERS[1];
  worksheet.getCell(`C${HEADER_ROW}`).value = HEADERS[2];
  worksheet.getCell(`D${HEADER_ROW}`).value = HEADERS[3];
  worksheet.getCell(`E${HEADER_ROW}`).value = HEADERS[4];
  configureHeaderRow(worksheet);

  const rows = [
    ...DUMMY_ROWS,
    ...Array.from({ length: EXTRA_READY_ROWS }, () => ["", "", "", "", ""])
  ];

  rows.forEach((row, index) => {
    const rowNumber = FIRST_DATA_ROW + index;
    worksheet.getCell(`A${rowNumber}`).value = row[0];
    worksheet.getCell(`B${rowNumber}`).value = row[1];
    worksheet.getCell(`C${rowNumber}`).value = row[2];
    worksheet.getCell(`D${rowNumber}`).value = row[3];
    worksheet.getCell(`E${rowNumber}`).value = row[4];
  });

  for (let rowNumber = FIRST_DATA_ROW; rowNumber <= LAST_READY_ROW; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);

    row.eachCell((cell) => {
      cell.alignment = { vertical: "middle" };
      applyCellBorders(cell, "FFE4EAF1");
      cell.protection = { locked: false };
    });

    if (rowNumber % 2 === 1) {
      row.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF8FAFC" }
        };
      });
    }

    worksheet.getCell(`B${rowNumber}`).dataValidation = {
      type: "custom",
      allowBlank: rowNumber > FIRST_DATA_ROW + DUMMY_ROWS.length - 1,
      showErrorMessage: true,
      errorStyle: "error",
      errorTitle: "Invalid email",
      error: "Please enter a valid email address.",
      formulae: [
        `OR(B${rowNumber}="",AND(ISNUMBER(SEARCH("@",B${rowNumber})),ISNUMBER(SEARCH(".",B${rowNumber},SEARCH("@",B${rowNumber})+2)),LEN(B${rowNumber})-LEN(SUBSTITUTE(B${rowNumber},"@",""))=1))`
      ]
    };
  }

  worksheet.autoFilter = `A${HEADER_ROW}:E${LAST_READY_ROW}`;
  worksheet.views = [{ state: "frozen", ySplit: HEADER_ROW }];
}

async function buildWorkbook(schoolName) {
  const workbook = new ExcelJS.Workbook();
  const studentsSheet = workbook.addWorksheet("Students");

  configureDataSheet(studentsSheet, schoolName);

  await studentsSheet.protect("template-lock", {
    selectLockedCells: true,
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
    const { schoolName } = await request.json();

    if (!schoolName || typeof schoolName !== "string" || !schoolName.trim()) {
      return Response.json(
        { error: "A valid school name is required." },
        { status: 400 }
      );
    }

    const workbook = await buildWorkbook(schoolName.trim());
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
