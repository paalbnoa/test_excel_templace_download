# What This Application Does

This application is a small portal for institutions that need to prepare and validate an Excel file before sending semester-fee related student data to SiO.

## Main Flow

1. The user opens the portal.
2. The portal is divided into two main areas:
   - Left side: download a template
   - Right side: validate a completed workbook

## Download Template

1. The user enters the institution name.
2. The user selects one or more allowed semesters from the preset list: `2025H`, `2026V`, `2026H`.
3. When the user clicks `Download Template`, the browser sends the request to `/api/template`.
4. The server generates a new `.xlsx` workbook using `ExcelJS`.
5. The workbook includes:
   - A `Students` sheet
   - An `Instructions` sheet
   - The institution name in the sheet metadata
   - The selected semesters in the sheet metadata
   - 100 ready-to-fill table rows
6. The `Students` sheet contains these columns:
   - `PersonID`
   - `Fornavn`
   - `Etternavn`
   - `Sist.bet.sem.avg`
   - `Fritatt.sem.avg`
   - `Epost`
   - `Prefiks`
   - `Mobilnummer`
7. Excel validation rules are built into the template so the sheet itself helps prevent bad input:
   - `PersonID` must be whole numbers
   - `Fornavn` is required
   - `Etternavn` is required
   - `Sist.bet.sem.avg` must match one of the selected semesters
   - `Epost` must look like a valid email address
   - `Prefiks` is required
   - `Mobilnummer` must be an 8-digit Norwegian mobile number starting with `4` or `9`
8. The generated workbook is returned to the browser and downloaded as an `.xlsx` file.

## Validate Workbook

1. After filling in the Excel file, the user clicks `Validate Excel`.
2. The user uploads the completed `.xlsx` workbook.
3. The browser sends the file to `/api/validate`.
4. The server opens the workbook and checks:
   - That a worksheet exists
   - That the expected headers are present
   - That only rows with actual data are validated
   - That each populated row follows the required rules
5. If the uploaded file contains no actual data rows, that is treated as a validation error.
6. The portal shows a validation summary on screen:
   - Whether validation passed or failed
   - How many rows were checked
   - How many warnings and errors were found

## Error Output

1. If validation fails, the portal does not show the full row-by-row error list directly on the page.
2. Instead, it provides a link to download a text file with the full validation details.
3. If the workbook contains real data rows and row-level validation errors, the portal also provides a second link to download an annotated Excel file.
4. The annotated Excel file:
   - Removes fully empty table rows
   - Highlights only the specific cells that contain validation errors
   - Does not highlight valid cells
   - Does not include Excel comments/markers

## What Is Involved Technically

- Frontend page/UI: `app/page.js`
- Template download API: `app/api/template/route.js`
- Validation upload API: `app/api/validate/route.js`
- Excel generation, validation, and annotated workbook logic: `lib/template.js`

In short: the app helps an institution download a controlled Excel template, fill it in, validate the uploaded data, and if needed download both a detailed error report and an Excel copy with invalid cells highlighted.
