"use client";

import { useState } from "react";

export default function HomePage() {
  const [schoolName, setSchoolName] = useState("");
  const [isDownloading, setIsDownloading] = useState(false);
  const [error, setError] = useState("");

  async function handleDownload() {
    const trimmedSchoolName = schoolName.trim();

    if (!trimmedSchoolName) {
      setError("Please enter the name of a school before downloading the template.");
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
        body: JSON.stringify({ schoolName: trimmedSchoolName })
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
        <h1>Download a template for your school.</h1>
        <p className="intro-text">
          Enter the school name below to generate a polished spreadsheet with 20 rows
          of sample data. The educational institution column is prefilled and protected,
          while the email column includes Excel validation.
        </p>

        <div className="form-panel">
          <label className="field-label" htmlFor="schoolName">
            Name of school
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
