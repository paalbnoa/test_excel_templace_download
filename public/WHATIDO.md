# What This Application Does

This application is a small portal for institutions that need to prepare and validate an Excel file before sending semester-fee related student data to SiO.

## Main Flow

1. The user opens the portal.
2. The portal is divided into three sections arranged in a left-to-right flow:
   - Section 1: download a template
   - Section 2: validate a completed workbook
   - Section 3: send the validated workbook to SiO by email

## Download Template

1. The user enters the institution short name.
2. The user selects exactly one semester from the preset list: `2025H`, `2026V`, `2026H`.
3. The user can optionally choose to include test data in the Excel file before downloading it.
4. The user can also choose to include a few random validation errors and blank required cells in that test data for easier testing.
5. When the user clicks `Download Template`, the browser sends the request to `/api/template`.
6. The server generates a new `.xlsm` (macro-enabled) workbook using `ExcelJS`, then post-processes the raw binary using `JSZip` to inject a VBA macro and a clickable button.
7. The workbook includes:
   - A `Students` sheet
   - An `Instructions` sheet
   - The institution short name in the sheet metadata
   - The selected semester in the sheet metadata
   - 100 ready-to-fill table rows with validation rules, borders, and alternating row colors
   - A clickable macro button in cell E3 labelled `Add 100 new empty rows`; clicking it runs the embedded VBA macro (`Add100Rows`) which extends the table by 100 rows, and the conditional formatting automatically applies the correct alternating colors to the new rows
   - Gridlines hidden on both sheets; borders are shown only on the editable data cells
   - Optional sample/test data if the user selected that option
8. If test data is included, the workbook is prefilled with 100 rows of sample data that matches the validation rules unless the optional random-error setting intentionally changes a few cells.
9. The `Students` sheet contains these columns:
   - `PersonID`
   - `Fornavn`
   - `Etternavn`
   - `Fritatt.sem.avg`
   - `Epost`
   - `Prefiks`
   - `Mobilnummer`
8. The selected semester is shown above the table in the workbook metadata and is not part of the table itself.
9. Excel validation rules are built into the template so the sheet itself helps prevent bad input:
   - `PersonID` must contain exactly 11 digits
   - `Fornavn` is required and must contain letters only
   - `Etternavn` is required and must contain letters only, with no special characters
   - `Fritatt.sem.avg` is a text field and may be blank, or contain only `Ja` or `Nei`
   - `Epost` must look like a valid email address
   - `Prefiks` must start with `+`, followed by digits only, with no spaces, and `00` is not allowed
   - `Mobilnummer` must contain exactly 8 digits with no spaces
10. The generated workbook is returned to the browser and downloaded as an `.xlsm` file.

## Validate Workbook

1. After filling in the Excel file, the user clicks `Validate Excel`.
2. The user uploads the completed `.xlsx` or `.xlsm` workbook.
3. The browser sends the file to `/api/validate`.
4. The server opens the workbook and checks:
   - That a worksheet exists
   - That the expected headers are present
   - That the semester metadata can be read from the workbook
   - That only rows with actual data are validated
   - That each populated row follows the required rules for names, exemption value, email, prefix, phone number, and `PersonID`
5. If the uploaded file contains no actual data rows, that is treated as a validation error.
6. The portal shows a validation summary on screen:
   - Whether validation passed or failed
   - How many rows were checked
   - How many warnings and errors were found

## Send To SiO

1. After the workbook validates successfully, the portal instructs the user to send the Excel file to SiO.
2. The file should be attached to an email sent to `semester_fees@sio.no`.
3. This is presented as the third and final step in the portal flow.

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
- VBA macro injection and macro-enabled workbook assembly: `lib/macro-workbook.js`

In short: the app helps an institution download a controlled Excel template, fill it in, validate the uploaded data, and then send the validated workbook to SiO. If validation fails, the app also provides a detailed error report and, when relevant, an Excel copy with invalid cells highlighted.
