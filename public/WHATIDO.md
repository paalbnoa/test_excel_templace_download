# What This Application Does

This application is a small web portal for creating and checking an Excel spreadsheet used to submit student semester-fee data.

## How It Works

1. The user opens the portal and enters the name of an institution.
2. The user selects one or more allowed semesters from the preset list: `2025H`, `2026V`, `2026H`.
3. When the user clicks `Download Template`, the browser sends that information to the server at `/api/template`.
4. The server generates a new `.xlsx` workbook using `ExcelJS`.
5. The workbook includes:
   - A `Students` sheet
   - An `Instructions` sheet
   - The institution name written into the sheet
   - The selected semesters written into the sheet
   - 100 ready-to-fill data rows
6. The `Students` sheet contains these columns:
   - `PersonID`
   - `Fornavn`
   - `Etternavn`
   - `Sist.bet.sem.avg`
   - `Fritatt.sem.avg`
   - `Epost`
   - `Prefiks`
   - `Mobilnummer`
7. Excel validation rules are added directly into the template so users get spreadsheet-level checks while typing:
   - `PersonID` must be whole numbers
   - `Fornavn` is required
   - `Etternavn` is required
   - `Sist.bet.sem.avg` must match one of the selected semesters
   - `Epost` must look like a valid email
   - `Prefiks` is required
   - `Mobilnummer` must be an 8-digit Norwegian mobile number starting with `4` or `9`
8. The finished workbook is returned to the browser and downloaded as an `.xlsx` file.

## Validation Flow

1. After filling in the Excel file, the user comes back to the portal and clicks `Validate Excel`.
2. The user uploads the completed `.xlsx` file.
3. The browser sends the file to `/api/validate`.
4. The server opens the workbook and checks:
   - That a valid worksheet exists
   - That the expected column headers are present
   - That each non-empty row follows the required rules
5. The server returns a result with:
   - Whether the workbook passed
   - How many rows were checked
   - Warnings
   - Row-by-row error messages
6. The portal shows the results on screen so the user can fix issues before sending the file onward.

## What Is Involved Technically

- Frontend page/UI: `app/page.js`
- Template download API: `app/api/template/route.js`
- Validation upload API: `app/api/validate/route.js`
- Excel generation and validation logic: `lib/template.js`

In short: the app helps a school generate a standardized Excel template, fill it with student data, and validate that data before submission.
