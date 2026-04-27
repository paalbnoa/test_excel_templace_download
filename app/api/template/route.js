import { buildWorkbook, getHeaderNotes, SEMESTER_OPTIONS } from "../../../lib/template";
import {
  addMacroButtonToWorkbookBuffer,
  MACRO_ENABLED_CONTENT_TYPE,
  MACRO_ENABLED_EXTENSION
} from "../../../lib/macro-workbook";

export async function POST(request) {
  try {
    const { schoolName, semesters, includeTestData, includeRandomErrors } =
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

    const workbook = await buildWorkbook(schoolName.trim(), semesters, {
      includeTestData: Boolean(includeTestData),
      includeRandomErrors: Boolean(includeTestData) && Boolean(includeRandomErrors)
    });
    const workbookBuffer = await workbook.xlsx.writeBuffer();
    const buffer = await addMacroButtonToWorkbookBuffer(workbookBuffer, { notes: getHeaderNotes() });
    const filename = `${schoolName.trim().replace(/[^a-z0-9]+/gi, "-").toLowerCase()}-template.${MACRO_ENABLED_EXTENSION}`;

    return new Response(buffer, {
      status: 200,
      headers: {
        "Content-Type": MACRO_ENABLED_CONTENT_TYPE,
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
