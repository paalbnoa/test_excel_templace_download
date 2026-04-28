const SL_BASE_URL = process.env.SL_BASE_URL || "http://localhost:9065";
const SL_API_KEY = process.env.SL_API_KEY || "123";

const VALIDATE_URL = `${SL_BASE_URL}/api/v1/excel/validate`;
const VALIDATE_DETAILS_URL = `${SL_BASE_URL}/api/v1/excel/validate/details`;

function buildMultipart(uploadedFile, buffer) {
  const formData = new FormData();
  const blob = new Blob([buffer], {
    type: uploadedFile.type || "application/octet-stream"
  });
  formData.append("semesterFeesMultipartFile", blob, uploadedFile.name);
  return formData;
}

function parseRowFromMessage(message) {
  const match = message?.match(/rowNumber=(\d+)/);
  return match ? Number(match[1]) : null;
}

function parseErrorsFromMessage(message) {
  const match = message?.match(/errors=\[(.*)\]$/);
  return match ? match[1] : message;
}

async function fetchDetails(uploadedFile, buffer) {
  const response = await fetch(VALIDATE_DETAILS_URL, {
    method: "POST",
    headers: { "X-API-KEY": SL_API_KEY },
    body: buildMultipart(uploadedFile, buffer)
  });

  if (response.status === 204) {
    return [];
  }

  let payload = null;
  try {
    payload = await response.json();
  } catch {
    return [];
  }

  const rawErrors = payload?.errors || [];
  return rawErrors
    .map((entry) => ({
      rowNumber: parseRowFromMessage(entry.message),
      message: parseErrorsFromMessage(entry.message)
    }))
    .filter((item) => item.message);
}

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

    const arrayBuffer = await uploadedFile.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const validateResponse = await fetch(VALIDATE_URL, {
      method: "POST",
      headers: { "X-API-KEY": SL_API_KEY },
      body: buildMultipart(uploadedFile, buffer)
    });

    if (validateResponse.status !== 200 && validateResponse.status !== 204) {
      const text = await validateResponse.text();
      return Response.json(
        {
          error: `SL validation failed with status ${validateResponse.status}.`,
          details: text
        },
        { status: 502 }
      );
    }

    const errorCountHeader = validateResponse.headers.get("X-Validation-Error-Count");
    const errorCount = errorCountHeader ? Number(errorCountHeader) : 0;

    if (validateResponse.status === 204 || errorCount === 0) {
      return Response.json(
        {
          isValid: true,
          summary: { rowCount: 0, errorCount: 0, warningCount: 0 },
          semesters: [],
          warnings: [],
          errors: [],
          source: "sl"
        },
        { status: 200 }
      );
    }

    const annotatedArrayBuffer = await validateResponse.arrayBuffer();
    const highlightedBase64 = Buffer.from(annotatedArrayBuffer).toString("base64");

    const detailedErrors = await fetchDetails(uploadedFile, buffer);

    return Response.json(
      {
        isValid: false,
        summary: {
          rowCount: 0,
          errorCount,
          warningCount: 0
        },
        semesters: [],
        warnings: [],
        errors: detailedErrors,
        highlightedWorkbook: highlightedBase64,
        source: "sl"
      },
      { status: 200 }
    );
  } catch (error) {
    return Response.json(
      {
        error: "Unable to validate the workbook via Service Layer.",
        details: error?.message
      },
      { status: 500 }
    );
  }
}
