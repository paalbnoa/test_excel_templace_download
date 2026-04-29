"use client";

import { useEffect, useRef, useState } from "react";

const SEMESTER_OPTIONS = ["2025H", "2026V", "2026H"];

const TEXT = {
  en: {
    languageLabel: "Choose language",
    logoAlt: "Studentsamskipnaden SiO",
    workflowLabel: "Portal workflow",
    heroTitle: "Download an Excel template for your institution.",
    intro:
      "Enter the institution short name, choose a semester, and generate an Excel template to use to send student data to SiO.",
    downloadTitle: "1. Download template",
    downloadText:
      "Enter the institution short name and choose the semester value that should be allowed in the Excel template.",
    institutionLabel: "Institution short name",
    institutionPlaceholder: "Example: BI",
    semesterLabel: "Semester fee",
    semesterPlaceholder: "Select semester",
    numRowsLabel: "Number of rows",
    testingTitle: "For testing purposes",
    includeTestData: "Include 100 rows of test data",
    includeRandomErrors: "Include random errors",
    includeMacros: "Include macros",
    preparingFile: "Preparing file...",
    downloadButton: "Download Template",
    validateTitle: "2. Validate workbook",
    validateIntro:
      "Before you send your Excel with data to SiO please validate the contents and correct any errors discovered.",
    validating: "Validating...",
    validateNoSl: "Validate Excel no SL",
    validatingSl: "Validating via SL...",
    validateSl: "Validate Excel using SL",
    selectedFile: (fileName) => `Selected file: ${fileName}`,
    validationPassed: "Validation passed",
    validationIssues: "Validation found issues",
    validatedViaSl: "Validated via Service Layer (semester fees).",
    checkedRows: (count) => `Checked ${count} data row${count === 1 ? "" : "s"}.`,
    warnings: "Warnings",
    validationFoundErrors: (count) =>
      `Validation found ${count} error${count === 1 ? "" : "s"}.`,
    openDetails: "Open full validation details",
    downloadHighlighted: "Download Excel file with highlighted errors",
    noErrors: "No validation errors were found in the uploaded workbook.",
    sendTitle: "3. Send to SiO",
    sendText: "Upload the Excel to SiO by using the Automation tool.",
    openAutomation: "Open automation tool",
    apiStampAlt: "API stamp",
    whatido: "For devs: step-by-step description of what this application does",
    detailsTitle: "Validation details",
    source: "Source",
    sourceSl: "Service Layer (semester fees)",
    sourceLocal: "Local",
    validationStatus: "Validation status",
    passed: "Passed",
    failed: "Failed",
    rowsChecked: "Rows checked",
    errorCount: "Error count",
    warningCount: "Warning count",
    allowedSemesters: "Allowed semesters",
    none: "None",
    errors: "Errors",
    row: "Row",
    workbook: "Workbook"
  },
  no: {
    languageLabel: "Velg språk",
    logoAlt: "Studentsamskipnaden SiO",
    workflowLabel: "Portalflyt",
    heroTitle: "Last ned en Excel-mal for institusjonen din.",
    intro:
      "Skriv inn institusjonens kortnavn, velg semester, og generer en Excel-mal som brukes til å sende studentdata til SiO.",
    downloadTitle: "1. Last ned mal",
    downloadText:
      "Skriv inn institusjonens kortnavn og velg semesterverdien som skal være tillatt i Excel-malen.",
    institutionLabel: "Institusjonens kortnavn",
    institutionPlaceholder: "Eksempel: BI",
    semesterLabel: "Semesteravgift",
    semesterPlaceholder: "Velg semester",
    numRowsLabel: "Antall rader",
    testingTitle: "For testing",
    includeTestData: "Inkluder 100 rader med testdata",
    includeRandomErrors: "Inkluder tilfeldige feil",
    includeMacros: "Inkluder makroer",
    preparingFile: "Klargjør fil...",
    downloadButton: "Last ned mal",
    validateTitle: "2. Valider arbeidsbok",
    validateIntro:
      "Før du sender Excel-filen med data til SiO, bør du validere innholdet og rette eventuelle feil som blir funnet.",
    validating: "Validerer...",
    validateNoSl: "Valider Excel uten SL",
    validatingSl: "Validerer med SL...",
    validateSl: "Valider Excel med SL",
    selectedFile: (fileName) => `Valgt fil: ${fileName}`,
    validationPassed: "Validering bestått",
    validationIssues: "Valideringen fant feil",
    validatedViaSl: "Validert via Service Layer (semesteravgifter).",
    checkedRows: (count) => `Kontrollerte ${count} datarad${count === 1 ? "" : "er"}.`,
    warnings: "Advarsler",
    validationFoundErrors: (count) =>
      `Valideringen fant ${count} feil.`,
    openDetails: "Åpne full valideringsrapport",
    downloadHighlighted: "Last ned Excel-fil med markerte feil",
    noErrors: "Ingen valideringsfeil ble funnet i den opplastede arbeidsboken.",
    sendTitle: "3. Send til SiO",
    sendText: "Last opp Excel-filen til SiO ved å bruke automasjonsverktøyet.",
    openAutomation: "Åpne automasjonsverktøy",
    apiStampAlt: "API-stempel",
    whatido: "Les en trinnvis beskrivelse av hva applikasjonen gjør",
    detailsTitle: "Valideringsdetaljer",
    source: "Kilde",
    sourceSl: "Service Layer (semesteravgifter)",
    sourceLocal: "Lokal",
    validationStatus: "Valideringsstatus",
    passed: "Bestått",
    failed: "Feilet",
    rowsChecked: "Rader kontrollert",
    errorCount: "Antall feil",
    warningCount: "Antall advarsler",
    allowedSemesters: "Tillatte semestre",
    none: "Ingen",
    errors: "Feil",
    row: "Rad",
    workbook: "Arbeidsbok"
  }
};

const MESSAGE_TRANSLATIONS_NO = {
  "Please enter the institution short name before downloading the template.":
    "Skriv inn institusjonens kortnavn før du laster ned malen.",
  "Please select a semester before downloading the template.":
    "Velg semester før du laster ned malen.",
  "The template could not be generated.": "Malen kunne ikke genereres.",
  "Something went wrong while creating the file.":
    "Noe gikk galt under opprettelsen av filen.",
  "The workbook could not be validated.": "Arbeidsboken kunne ikke valideres.",
  "Something went wrong while validating the workbook.":
    "Noe gikk galt under valideringen av arbeidsboken.",
  "Please upload an Excel file to validate.": "Last opp en Excel-fil som skal valideres.",
  "Only .xlsx and .xlsm files are supported.": "Kun .xlsx- og .xlsm-filer støttes.",
  "Unable to validate the uploaded workbook.":
    "Kunne ikke validere den opplastede arbeidsboken.",
  "Unable to validate the workbook via Service Layer.":
    "Kunne ikke validere arbeidsboken via Service Layer.",
  "A valid institution short name is required.":
    "Et gyldig kortnavn for institusjonen er påkrevd.",
  "Exactly one valid semester is required.": "Nøyaktig ett gyldig semester er påkrevd.",
  "Unable to generate template.": "Kunne ikke generere malen.",
  "The workbook does not contain a worksheet to validate.":
    "Arbeidsboken inneholder ikke et regneark som kan valideres.",
  "Could not find the expected template headers in the uploaded workbook.":
    "Fant ikke de forventede maloverskriftene i den opplastede arbeidsboken.",
  "The uploaded workbook does not contain any data rows in the table.":
    "Den opplastede arbeidsboken inneholder ingen datarader i tabellen.",
  "PersonID is required and must be a valid 11-digit Norwegian personal number with a real birth date, valid individual number, and valid MOD11 check digits.":
    "PersonID er påkrevd og må være et gyldig norsk fødselsnummer med 11 sifre, ekte fødselsdato, gyldig individnummer og gyldige MOD11-kontrollsifre.",
  "PersonID is required and must be a valid 11-digit Norwegian personal number. The person must be younger than 100 years old. The first 6 digits must be a real birth date, followed by a valid individual number for the birth year and two valid MOD11 check digits.":
    "PersonID er påkrevd og må være et gyldig norsk fødselsnummer med 11 sifre. Personen må være yngre enn 100 år. De første 6 sifrene må være en ekte fødselsdato, etterfulgt av et gyldig individnummer for fødselsåret og to gyldige MOD11-kontrollsifre.",
  "Fornavn is required and must contain letters only.":
    "Fornavn er påkrevd og kan kun inneholde bokstaver.",
  "Etternavn is required and must contain letters only. No special characters are allowed.":
    "Etternavn er påkrevd og kan kun inneholde bokstaver. Spesialtegn er ikke tillatt.",
  'Fritatt.sem.avg may only be blank, "Ja", or "Nei".':
    'Fritatt.sem.avg kan kun være tom, "Ja" eller "Nei".',
  "Epost is required and must be a valid email address.":
    "Epost er påkrevd og må være en gyldig e-postadresse.",
  'Prefiks is required and must start with "+", followed by digits only. "00" is not allowed and spaces are not allowed.':
    'Prefiks er påkrevd og må starte med "+", etterfulgt av kun sifre. "00" er ikke tillatt, og mellomrom er ikke tillatt.',
  "Mobilnummer is required and must contain exactly 8 digits with no spaces.":
    "Mobilnummer er påkrevd og må inneholde nøyaktig 8 sifre uten mellomrom."
};

function translateMessage(message, language) {
  if (!message || language === "en") {
    return message;
  }

  const statusMatch = message.match(/^SL validation failed with status (\d+)\.$/);
  if (statusMatch) {
    return `SL-validering feilet med status ${statusMatch[1]}.`;
  }

  return MESSAGE_TRANSLATIONS_NO[message] || message;
}

function GlobeIcon() {
  return (
    <svg
      viewBox="0 0 64 64"
      aria-hidden="true"
      className="language-icon"
      focusable="false"
    >
      <rect width="64" height="64" rx="32" />
      <circle cx="32" cy="32" r="19" />
      <path d="M13 32h38" />
      <path d="M32 13c6 5 9 11 9 19s-3 14-9 19c-6-5-9-11-9-19s3-14 9-19Z" />
      <path d="M17 23c4 2 9 3 15 3s11-1 15-3" />
      <path d="M17 41c4-2 9-3 15-3s11 1 15 3" />
    </svg>
  );
}

function LanguageToggle({ language, onChange, labels }) {
  return (
    <div className="language-toggle" aria-label={labels.languageLabel}>
      <GlobeIcon />
      <div className="language-options">
        <button
          type="button"
          className={`language-option ${language === "no" ? "language-option-active" : ""}`}
          onClick={() => onChange("no")}
          aria-pressed={language === "no"}
        >
          No
        </button>
        <span className="language-divider" aria-hidden="true">
          |
        </span>
        <button
          type="button"
          className={`language-option ${language === "en" ? "language-option-active" : ""}`}
          onClick={() => onChange("en")}
          aria-pressed={language === "en"}
        >
          En
        </button>
      </div>
    </div>
  );
}

function buildValidationDetailsText(validationResult, labels, language) {
  const isSl = validationResult.source === "sl";
  const lines = [
    labels.detailsTitle,
    "",
    `${labels.source}: ${isSl ? labels.sourceSl : labels.sourceLocal}`,
    `${labels.validationStatus}: ${validationResult.isValid ? labels.passed : labels.failed}`,
    ...(isSl
      ? []
      : [`${labels.rowsChecked}: ${validationResult.summary.rowCount}`]),
    `${labels.errorCount}: ${validationResult.summary.errorCount}`,
    `${labels.warningCount}: ${validationResult.summary.warningCount}`,
    ...(isSl
      ? []
      : [`${labels.allowedSemesters}: ${validationResult.semesters.join(", ") || labels.none}`]),
    ""
  ];

  if (validationResult.warnings?.length) {
    lines.push(`${labels.warnings}:`);
    validationResult.warnings.forEach((warning) => {
      lines.push(`- ${translateMessage(warning, language)}`);
    });
    lines.push("");
  }

  if (validationResult.errors?.length) {
    lines.push(`${labels.errors}:`);
    validationResult.errors.forEach((item) => {
      const rowLabel = item.rowNumber ? `${labels.row} ${item.rowNumber}` : labels.workbook;
      const columnLabel = item.column ? ` (${item.column})` : "";
      lines.push(`- ${rowLabel}${columnLabel}: ${translateMessage(item.message, language)}`);
    });
  }

  return lines.join("\n");
}

function buildHighlightedWorkbookHref(validationResult) {
  if (!validationResult.highlightedWorkbook) {
    return "";
  }

  return `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${validationResult.highlightedWorkbook}`;
}

export default function HomePage() {
  const [language, setLanguage] = useState("en");
  const [schoolName, setSchoolName] = useState("");
  const [selectedSemester, setSelectedSemester] = useState("");
  const [numRows, setNumRows] = useState(1000);
  const [includeTestData, setIncludeTestData] = useState(false);
  const [includeRandomErrors, setIncludeRandomErrors] = useState(false);
  const [includeMacros, setIncludeMacros] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [error, setError] = useState("");
  const [selectedFileName, setSelectedFileName] = useState("");
  const [isValidating, setIsValidating] = useState(false);
  const [validationError, setValidationError] = useState("");
  const [validationResult, setValidationResult] = useState(null);
  const [validationMode, setValidationMode] = useState("local");
  const fileInputRef = useRef(null);
  const slFileInputRef = useRef(null);
  const labels = TEXT[language];

  useEffect(() => {
    const storedLanguage = window.localStorage.getItem("portal-language");
    if (storedLanguage === "no" || storedLanguage === "en") {
      setLanguage(storedLanguage);
    }
  }, []);

  useEffect(() => {
    document.documentElement.lang = language;
    window.localStorage.setItem("portal-language", language);
  }, [language]);

  async function handleDownload() {
    const trimmedSchoolName = schoolName.trim();

    if (!trimmedSchoolName) {
      setError("Please enter the institution short name before downloading the template.");
      return;
    }

    if (!selectedSemester) {
      setError("Please select a semester before downloading the template.");
      return;
    }

    setError("");
    setIsDownloading(true);

    try {
      const response = await fetch("/api/template", {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          schoolName: trimmedSchoolName,
          semesters: [selectedSemester],
          includeTestData,
          includeRandomErrors,
          includeMacros,
          numRows
        })
      });

      if (!response.ok) {
        throw new Error("The template could not be generated.");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      const safeName = trimmedSchoolName.toLowerCase().replace(/[^a-z0-9]+/g, "-");
      const extension = includeMacros ? "xlsm" : "xlsx";

      link.href = url;
      link.download = `${safeName || "school"}-template.${extension}`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (downloadError) {
      setError(downloadError.message || "Something went wrong while creating the file.");
    } finally {
      setIsDownloading(false);
    }
  }

  async function runValidation(uploadedFile, endpoint) {
    const formData = new FormData();
    formData.append("file", uploadedFile);

    setIsValidating(true);

    try {
      const response = await fetch(endpoint, {
        method: "POST",
        body: formData
      });
      const result = await response.json();

      if (!response.ok) {
        throw new Error(result.error || "The workbook could not be validated.");
      }

      setValidationResult(result);
    } catch (uploadError) {
      setValidationError(
        uploadError.message || "Something went wrong while validating the workbook."
      );
    } finally {
      setIsValidating(false);
    }
  }

  async function handleFileChange(event) {
    const uploadedFile = event.target.files?.[0];

    setValidationResult(null);
    setValidationError("");
    setSelectedFileName(uploadedFile?.name || "");

    if (!uploadedFile) {
      return;
    }

    setValidationMode("local");
    await runValidation(uploadedFile, "/api/validate");
  }

  async function handleSlFileChange(event) {
    const uploadedFile = event.target.files?.[0];

    setValidationResult(null);
    setValidationError("");
    setSelectedFileName(uploadedFile?.name || "");

    if (!uploadedFile) {
      return;
    }

    setValidationMode("sl");
    await runValidation(uploadedFile, "/api/validate-sl");
  }

  function handleValidateClick() {
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
      fileInputRef.current.click();
    }
  }

  function handleValidateSlClick() {
    if (slFileInputRef.current) {
      slFileInputRef.current.value = "";
      slFileInputRef.current.click();
    }
  }

  function handleOpenValidationDetails() {
    if (!validationResult) {
      return;
    }

    const detailsBlob = new Blob([buildValidationDetailsText(validationResult, labels, language)], {
      type: "text/plain;charset=utf-8"
    });
    const detailsUrl = window.URL.createObjectURL(detailsBlob);

    window.open(detailsUrl, "_blank", "noopener,noreferrer");
    window.setTimeout(() => {
      window.URL.revokeObjectURL(detailsUrl);
    }, 60000);
  }

  function handleOpenAutomationTool() {
    if (!validationResult?.isValid) {
      return;
    }

    window.open(
      "http://localhost:3001/dashboard/semester-fees/hackathon",
      "_blank",
      "noopener,noreferrer"
    );
  }

  return (
    <main className="page-shell">
      <header className="top-banner" aria-label="SiO">
        <img
          src="/sio_logo.avif"
          alt={labels.logoAlt}
          className="top-banner-logo"
          width="280"
          height="56"
        />
        <LanguageToggle language={language} onChange={setLanguage} labels={labels} />
      </header>

      <section className="hero-card">
        <div className="brand-bar">
          <div className="brand-mark">
            <div className="brand-copy">
              <div className="hero-heading-row">
                <h1>{labels.heroTitle}</h1>
                <img
                  src="/stamp.png"
                  alt=""
                  aria-hidden="true"
                  className="hero-heading-stamp"
                  width="180"
                  height="180"
                />
              </div>
              <p className="intro-text">
                {labels.intro}
              </p>
            </div>
          </div>
        </div>

        <div className="portal-sections" aria-label={labels.workflowLabel}>
          <section className="form-panel">
            <div className="panel-header">
              <h2 className="panel-title">{labels.downloadTitle}</h2>
              <p className="panel-text">
                {labels.downloadText}
              </p>
            </div>

            <label className="field-label" htmlFor="schoolName">
              {labels.institutionLabel}
            </label>
            <input
              id="schoolName"
              name="schoolName"
              type="text"
              className="school-input"
              placeholder={labels.institutionPlaceholder}
              value={schoolName}
              onChange={(event) => {
                setSchoolName(event.target.value);
                setError("");
              }}
            />

            <div className="semester-group">
              <label className="field-label" htmlFor="semester">
                {labels.semesterLabel}
              </label>
              <select
                id="semester"
                name="semester"
                className="semester-select"
                value={selectedSemester}
                onChange={(event) => {
                  setSelectedSemester(event.target.value);
                  setError("");
                }}
              >
                <option value="">{labels.semesterPlaceholder}</option>
                {SEMESTER_OPTIONS.map((semester) => (
                  <option key={semester} value={semester}>
                    {semester}
                  </option>
                ))}
              </select>
            </div>

            <div className="semester-group">
              <label className="field-label" htmlFor="numRows">
                {labels.numRowsLabel}
              </label>
              <input
                type="number"
                id="numRows"
                name="numRows"
                className="semester-select"
                min={1}
                value={numRows}
                onChange={(event) => setNumRows(Number(event.target.value))}
              />
            </div>

            <div className="download-option-group">
              <p className="download-option-title">{labels.testingTitle}</p>
              <label className="download-option">
                <input
                  type="checkbox"
                  checked={includeTestData}
                  onChange={(event) => {
                    const isChecked = event.target.checked;
                    setIncludeTestData(isChecked);

                    if (!isChecked) {
                      setIncludeRandomErrors(false);
                    }
                  }}
                />
                <span>{labels.includeTestData}</span>
              </label>
              <label className="download-option">
                <input
                  type="checkbox"
                  checked={includeRandomErrors}
                  disabled={!includeTestData}
                  onChange={(event) => setIncludeRandomErrors(event.target.checked)}
                />
                <span>{labels.includeRandomErrors}</span>
              </label>

              <label className="download-option">
                <input
                  type="checkbox"
                  checked={includeMacros}
                  onChange={(event) => setIncludeMacros(event.target.checked)}
                />
                <span>{labels.includeMacros}</span>
              </label>
            </div>

            <button
              type="button"
              className="download-button"
              onClick={handleDownload}
              disabled={isDownloading}
            >
              {isDownloading ? labels.preparingFile : labels.downloadButton}
            </button>

            {error ? <p className="error-text">{translateMessage(error, language)}</p> : null}
          </section>

          <div className="flow-arrow" aria-hidden="true">
            <span className="flow-arrow-line" />
            <span className="flow-arrow-head" />
          </div>

          <section className="form-panel validation-panel">
            <div className="panel-header">
              <h2 className="panel-title">{labels.validateTitle}</h2>
            </div>

            <p className="validation-intro">
              {labels.validateIntro}
            </p>

            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12"
              className="file-input"
              onChange={handleFileChange}
            />
            <input
              ref={slFileInputRef}
              type="file"
              accept=".xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12"
              className="file-input"
              onChange={handleSlFileChange}
            />

            <button
              type="button"
              className="validate-button"
              onClick={handleValidateClick}
              disabled={isValidating}
            >
              {isValidating && validationMode === "local" ? labels.validating : labels.validateNoSl}
            </button>

            <button
              type="button"
              className="validate-button validate-button-sl"
              onClick={handleValidateSlClick}
              disabled={isValidating}
            >
              {isValidating && validationMode === "sl"
                ? labels.validatingSl
                : labels.validateSl}
            </button>

            {selectedFileName ? (
              <p className="selected-file">{labels.selectedFile(selectedFileName)}</p>
            ) : null}

            {validationError ? (
              <p className="error-text">{translateMessage(validationError, language)}</p>
            ) : null}

            {validationResult ? (
              <section
                className={`validation-results ${
                  validationResult.isValid ? "validation-results-success" : "validation-results-error"
                }`}
                aria-live="polite"
              >
                <h2 className={validationResult.isValid ? "" : "validation-error-heading"}>
                  {validationResult.isValid
                    ? labels.validationPassed
                    : labels.validationIssues}
                </h2>
                {validationResult.source === "sl" ? (
                  <p className="validation-summary">
                    {labels.validatedViaSl}
                  </p>
                ) : (
                  <p className="validation-summary">
                    {labels.checkedRows(validationResult.summary.rowCount)}
                  </p>
                )}

                {validationResult.warnings?.length ? (
                  <div className="validation-block">
                    <h3>{labels.warnings}</h3>
                    <ul className="validation-list">
                      {validationResult.warnings.map((warning) => (
                        <li key={warning}>{translateMessage(warning, language)}</li>
                      ))}
                    </ul>
                  </div>
                ) : null}

                {validationResult.errors?.length ? (
                  <div className="validation-block">
                    <p className="validation-compact-summary validation-error-summary">
                      {labels.validationFoundErrors(validationResult.summary.errorCount)}
                    </p>
                    <button
                      type="button"
                      className="validation-details-link"
                      onClick={handleOpenValidationDetails}
                    >
                      {labels.openDetails}
                    </button>
                    {validationResult.highlightedWorkbook ? (
                      <a
                        className="validation-details-link highlighted-workbook-link"
                        href={buildHighlightedWorkbookHref(validationResult)}
                        download={`${
                          selectedFileName.replace(/\.xlsx$/i, "") || "validation"
                        }-highlighted-errors.xlsx`}
                      >
                        {labels.downloadHighlighted}
                      </a>
                    ) : null}
                  </div>
                ) : (
                  <p className="success-text">
                    {labels.noErrors}
                  </p>
                )}
              </section>
            ) : null}
          </section>

          <div className="flow-arrow" aria-hidden="true">
            <span className="flow-arrow-line" />
            <span className="flow-arrow-head" />
          </div>

          <section className="form-panel send-panel">
            <div className="panel-header">
              <h2 className="panel-title">{labels.sendTitle}</h2>
              <p className="panel-text">
                {labels.sendText}
              </p>
              <button
                type="button"
                className="automation-button"
                onClick={handleOpenAutomationTool}
                disabled={!validationResult?.isValid}
              >
                {labels.openAutomation}
              </button>
            </div>
            <img
              src="/API stamp.png"
              alt={labels.apiStampAlt}
              className="send-api-stamp"
              width="1536"
              height="1024"
            />
          </section>
        </div>

        <a className="whatido-link" href="/WHATIDO.md" target="_blank" rel="noreferrer">
          {labels.whatido}
        </a>
      </section>
    </main>
  );
}
