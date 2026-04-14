import { validateWorkbookBuffer } from "../../../lib/template";

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

    const isExcelFile =
      uploadedFile.name.toLowerCase().endsWith(".xlsx") ||
      uploadedFile.type ===
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    if (!isExcelFile) {
      return Response.json(
        { error: "Only .xlsx files are supported." },
        { status: 400 }
      );
    }

    const arrayBuffer = await uploadedFile.arrayBuffer();
    const result = await validateWorkbookBuffer(arrayBuffer);

    return Response.json(result, { status: 200 });
  } catch {
    return Response.json(
      { error: "Unable to validate the uploaded workbook." },
      { status: 500 }
    );
  }
}
