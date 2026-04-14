"use client";

import { useRef, useState } from "react";

const SEMESTER_OPTIONS = ["2025H", "2026V", "2026H"];

export default function HomePage() {
  const [schoolName, setSchoolName] = useState("");
  const [selectedSemesters, setSelectedSemesters] = useState([]);
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
          semesters: selectedSemesters
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
        <div className="eyebrow">Semester fee template portal</div>
        <h1>Download a template for your institution.</h1>
        <p className="intro-text">
          Enter the institution name and choose one or more semesters to generate a
          polished spreadsheet template. The downloaded file includes validation for
          PersonID, Epost, and the allowed semester values.
        </p>

        <div className="form-panel">
          <label className="field-label" htmlFor="schoolName">
            Name of institution
          </label>
          <input
            id="schoolName"
            name="schoolName"
            type="text"
            className="school-input"
            placeholder="Example: Northbridge Business School"
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

          <button
            type="button"
            className="download-button"
            onClick={handleDownload}
            disabled={isDownloading}
          >
            {isDownloading ? "Preparing file..." : "Download Template"}
          </button>

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

          {error ? <p className="error-text">{error}</p> : null}
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
                  <h3>Errors</h3>
                  <ul className="validation-list">
                    {validationResult.errors.map((item, index) => (
                      <li key={`${item.rowNumber ?? "general"}-${item.column ?? "file"}-${index}`}>
                        {item.rowNumber ? `Row ${item.rowNumber}` : "Workbook"}
                        {item.column ? ` (${item.column})` : ""}: {item.message}
                      </li>
                    ))}
                  </ul>
                </div>
              ) : (
                <p className="success-text">
                  No validation errors were found in the uploaded workbook.
                </p>
              )}
            </section>
          ) : null}
        </div>

        <a className="whatido-link" href="/WHATIDO.md" target="_blank" rel="noreferrer">
          Read a step-by-step description of what this application does
        </a>
      </section>
    </main>
  );
}
