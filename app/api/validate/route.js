import {
  buildHighlightedWorkbookBuffer,
  validateWorkbookBuffer
} from "../../../lib/template";

export async function POST(request) {
  try {
    const formData = await request.formData();
    const uploadedFile = formData.get("file");

    if (!(uploadedFile instanceof File)) {
      return Response.json(
        { error: "Please upload an Excel file to validate." },
        { status: 400 }
      );
    }

    const normalizedFileName = uploadedFile.name.toLowerCase();
    const isExcelFile =
      normalizedFileName.endsWith(".xlsx") ||
      normalizedFileName.endsWith(".xlsm") ||
      uploadedFile.type ===
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      uploadedFile.type === "application/vnd.ms-excel.sheet.macroEnabled.12";

    if (!isExcelFile) {
      return Response.json(
        { error: "Only .xlsx and .xlsm files are supported." },
        { status: 400 }
      );
    }

    const arrayBuffer = await uploadedFile.arrayBuffer();
    const result = await validateWorkbookBuffer(arrayBuffer);
    const hasRowLevelErrors =
      result.summary.rowCount > 0 &&
      result.errors?.some((item) => item.rowNumber && item.column);

    if (hasRowLevelErrors) {
      const highlightedBuffer = await buildHighlightedWorkbookBuffer(
        arrayBuffer,
        result.errors
      );

      if (highlightedBuffer) {
        result.highlightedWorkbook = Buffer.from(highlightedBuffer).toString("base64");
      }
    }

    return Response.json({ ...result, source: "local" }, { status: 200 });
  } catch {
    return Response.json(
      { error: "Unable to validate the uploaded workbook." },
      { status: 500 }
    );
  }
}
