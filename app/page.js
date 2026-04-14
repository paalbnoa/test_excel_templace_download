"use client";

import { useRef, useState } from "react";

const SEMESTER_OPTIONS = ["2025H", "2026V", "2026H"];

function buildValidationDetailsText(validationResult) {
  const lines = [
    "Validation details",
    "",
    `Validation status: ${validationResult.isValid ? "Passed" : "Failed"}`,
    `Rows checked: ${validationResult.summary.rowCount}`,
    `Error count: ${validationResult.summary.errorCount}`,
    `Warning count: ${validationResult.summary.warningCount}`,
    `Allowed semesters: ${validationResult.semesters.join(", ") || "None"}`,
    ""
  ];

  if (validationResult.warnings?.length) {
    lines.push("Warnings:");
    validationResult.warnings.forEach((warning) => {
      lines.push(`- ${warning}`);
    });
    lines.push("");
  }

  if (validationResult.errors?.length) {
    lines.push("Errors:");
    validationResult.errors.forEach((item) => {
      const rowLabel = item.rowNumber ? `Row ${item.rowNumber}` : "Workbook";
      const columnLabel = item.column ? ` (${item.column})` : "";
      lines.push(`- ${rowLabel}${columnLabel}: ${item.message}`);
    });
  }

  return lines.join("\n");
}

function buildValidationDetailsHref(validationResult) {
  return `data:text/plain;charset=utf-8,${encodeURIComponent(
    buildValidationDetailsText(validationResult)
  )}`;
}

function buildHighlightedWorkbookHref(validationResult) {
  if (!validationResult.highlightedWorkbook) {
    return "";
  }

  return `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${validationResult.highlightedWorkbook}`;
}

export default function HomePage() {
  const [schoolName, setSchoolName] = useState("");
  const [selectedSemesters, setSelectedSemesters] = useState([]);
  const [includeTestData, setIncludeTestData] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [error, setError] = useState("");
  const [selectedFileName, setSelectedFileName] = useState("");
  const [isValidating, setIsValidating] = useState(false);
  const [validationError, setValidationError] = useState("");
  const [validationResult, setValidationResult] = useState(null);
  const fileInputRef = useRef(null);

  function handleSemesterToggle(semester) {
    setSelectedSemesters((currentSemesters) =>
      currentSemesters.includes(semester)
        ? currentSemesters.filter((value) => value !== semester)
        : [...currentSemesters, semester]
    );
  }

  async function handleDownload() {
    const trimmedSchoolName = schoolName.trim();

    if (!trimmedSchoolName) {
      setError("Please enter the name of an institution before downloading the template.");
      return;
    }

    if (selectedSemesters.length === 0) {
      setError("Please select at least one semester before downloading the template.");
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
          semesters: selectedSemesters,
          includeTestData
        })
      });

      if (!response.ok) {
        throw new Error("The template could not be generated.");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      const safeName = trimmedSchoolName.toLowerCase().replace(/[^a-z0-9]+/g, "-");

      link.href = url;
      link.download = `${safeName || "school"}-template.xlsx`;
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

  async function handleFileChange(event) {
    const uploadedFile = event.target.files?.[0];

    setValidationResult(null);
    setValidationError("");
    setSelectedFileName(uploadedFile?.name || "");

    if (!uploadedFile) {
      return;
    }

    const formData = new FormData();
    formData.append("file", uploadedFile);

    setIsValidating(true);

    try {
      const response = await fetch("/api/validate", {
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

  function handleValidateClick() {
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
      fileInputRef.current.click();
    }
  }

  return (
    <main className="page-shell">
      <section className="hero-card">
        <div className="brand-bar">
          <div className="brand-mark">
            <img
              src="/siologo.png"
              alt="Studentsamskipnaden SiO"
              className="brand-logo"
              width="280"
              height="56"
            />

            <div className="brand-copy">
              <div className="eyebrow">Semester fee template portal</div>
              <h1>Create an Excel template for your institution.</h1>
              <p className="intro-text">
                Enter the institution name, choose semester values, and generate an
                Excel template to use to send student data to SiO.
              </p>
            </div>
          </div>
        </div>

        <div className="portal-sections">
          <section className="form-panel">
            <div className="panel-header">
              <h2 className="panel-title">1. Download template</h2>
              <p className="panel-text">
                Enter the institution details and choose the semester values that
                should be allowed in the Excel template.
              </p>
            </div>

            <label className="field-label" htmlFor="schoolName">
              Name of institution
            </label>
            <input
              id="schoolName"
              name="schoolName"
              type="text"
              className="school-input"
              placeholder="Example: University of Oslo"
              value={schoolName}
              onChange={(event) => setSchoolName(event.target.value)}
            />

            <div className="semester-group">
              <p className="field-label">Semester(s)</p>
              <div className="semester-options">
                {SEMESTER_OPTIONS.map((semester) => (
                  <label key={semester} className="semester-option">
                    <input
                      type="checkbox"
                      checked={selectedSemesters.includes(semester)}
                      onChange={() => handleSemesterToggle(semester)}
                    />
                    <span>{semester}</span>
                  </label>
                ))}
              </div>
            </div>

            <div className="download-option-group">
              <label className="download-option">
                <input
                  type="checkbox"
                  checked={includeTestData}
                  onChange={(event) => setIncludeTestData(event.target.checked)}
                />
                <span>Include 100 rows of test data</span>
              </label>
            </div>

            <button
              type="button"
              className="download-button"
              onClick={handleDownload}
              disabled={isDownloading}
            >
              {isDownloading ? "Preparing file..." : "Download Template"}
            </button>

            {error ? <p className="error-text">{error}</p> : null}
          </section>

          <section className="form-panel validation-panel">
            <div className="panel-header">
              <h2 className="panel-title">2. Validate workbook</h2>
            </div>

            <p className="validation-intro">
              Before you send your Excel with data to SiO please validate the contents
              and correct any errors discovered.
            </p>

            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              className="file-input"
              onChange={handleFileChange}
            />

            <button
              type="button"
              className="validate-button"
              onClick={handleValidateClick}
              disabled={isValidating}
            >
              {isValidating ? "Validating..." : "Validate Excel"}
            </button>

            {selectedFileName ? (
              <p className="selected-file">Selected file: {selectedFileName}</p>
            ) : null}

            {validationError ? <p className="error-text">{validationError}</p> : null}

            {validationResult ? (
              <section className="validation-results" aria-live="polite">
                <h2>
                  {validationResult.isValid
                    ? "Validation passed"
                    : "Validation found issues"}
                </h2>
                <p className="validation-summary">
                  Checked {validationResult.summary.rowCount} data row
                  {validationResult.summary.rowCount === 1 ? "" : "s"} using semester
                  values {validationResult.semesters.join(", ")}.
                </p>

                {validationResult.warnings?.length ? (
                  <div className="validation-block">
                    <h3>Warnings</h3>
                    <ul className="validation-list">
                      {validationResult.warnings.map((warning) => (
                        <li key={warning}>{warning}</li>
                      ))}
                    </ul>
                  </div>
                ) : null}

                {validationResult.errors?.length ? (
                  <div className="validation-block">
                    <p className="validation-compact-summary">
                      Validation found {validationResult.summary.errorCount} error
                      {validationResult.summary.errorCount === 1 ? "" : "s"}.
                    </p>
                    <a
                      className="validation-details-link"
                      href={buildValidationDetailsHref(validationResult)}
                      download={`${selectedFileName.replace(/\.xlsx$/i, "") || "validation"}-details.txt`}
                    >
                      Download full validation details
                    </a>
                    {validationResult.highlightedWorkbook ? (
                      <a
                        className="validation-details-link highlighted-workbook-link"
                        href={buildHighlightedWorkbookHref(validationResult)}
                        download={`${
                          selectedFileName.replace(/\.xlsx$/i, "") || "validation"
                        }-highlighted-errors.xlsx`}
                      >
                        Download Excel file with highlighted errors
                      </a>
                    ) : null}
                  </div>
                ) : (
                  <p className="success-text">
                    No validation errors were found in the uploaded workbook.
                  </p>
                )}
              </section>
            ) : null}
          </section>
        </div>

        <a className="whatido-link" href="/WHATIDO.md" target="_blank" rel="noreferrer">
          Read a step-by-step description of what this application does
        </a>
      </section>
    </main>
  );
}
