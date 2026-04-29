import { buildWorkbook, SEMESTER_OPTIONS } from "../../../lib/template";
import {
  addMacroButtonToWorkbookBuffer,
  MACRO_ENABLED_CONTENT_TYPE,
  MACRO_ENABLED_EXTENSION
} from "../../../lib/macro-workbook";

const STANDARD_WORKBOOK_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const STANDARD_WORKBOOK_EXTENSION = "xlsx";

export async function POST(request) {
  try {
    const { schoolName, semesters, includeTestData, includeRandomErrors, includeMacros, numRows } =
      await request.json();

    if (!schoolName || typeof schoolName !== "string" || !schoolName.trim()) {
      return Response.json(
        { error: "A valid institution short name is required." },
        { status: 400 }
      );
    }

    if (
      !Array.isArray(semesters) ||
      semesters.length !== 1 ||
      semesters.some((semester) => !SEMESTER_OPTIONS.includes(semester))
    ) {
      return Response.json(
        { error: "Exactly one valid semester is required." },
        { status: 400 }
      );
    }

    const parsedNumRows = Number.isInteger(numRows) && numRows > 0 ? numRows : 1000;
    const shouldIncludeMacros = Boolean(includeMacros);
    const workbook = await buildWorkbook(schoolName.trim(), semesters, {
      includeTestData: Boolean(includeTestData),
      includeRandomErrors: Boolean(includeTestData) && Boolean(includeRandomErrors),
      includeMacros: shouldIncludeMacros,
      numRows: parsedNumRows
    });
    const workbookBuffer = await workbook.xlsx.writeBuffer();
    const buffer = shouldIncludeMacros
      ? await addMacroButtonToWorkbookBuffer(workbookBuffer)
      : workbookBuffer;
    const extension = shouldIncludeMacros
      ? MACRO_ENABLED_EXTENSION
      : STANDARD_WORKBOOK_EXTENSION;
    const contentType = shouldIncludeMacros
      ? MACRO_ENABLED_CONTENT_TYPE
      : STANDARD_WORKBOOK_CONTENT_TYPE;
    const filename = `${schoolName.trim().replace(/[^a-z0-9]+/gi, "-").toLowerCase()}-template.${extension}`;

    return new Response(buffer, {
      status: 200,
      headers: {
        "Content-Type": contentType,
        "Content-Disposition": `attachment; filename="${filename}"`
      }
    });
  } catch (error) {
    console.error("Template generation failed", error);

    return Response.json(
      { error: "Unable to generate template." },
      { status: 500 }
    );
  }
}
