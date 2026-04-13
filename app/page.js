"use client";

import { useState } from "react";

const SEMESTER_OPTIONS = ["2025H", "2026V", "2026H"];

export default function HomePage() {
  const [schoolName, setSchoolName] = useState("");
  const [selectedSemesters, setSelectedSemesters] = useState([]);
  const [isDownloading, setIsDownloading] = useState(false);
  const [error, setError] = useState("");

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

  return (
    <main className="page-shell">
      <section className="hero-card">
        <div className="eyebrow">Semester fee template generator</div>
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
            {isDownloading ? "Preparing file..." : "download template"}
          </button>

          {error ? <p className="error-text">{error}</p> : null}
        </div>
      </section>
    </main>
  );
}
