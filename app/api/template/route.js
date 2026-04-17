import { buildWorkbook, SEMESTER_OPTIONS } from "../../../lib/template";

export async function POST(request) {
  try {
    const { schoolName, semesters, includeTestData } = await request.json();

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
      includeTestData: Boolean(includeTestData)
    });
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
