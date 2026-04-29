# Exam Seating Planner

This is a static web app for generating an editable exam seating plan from an Excel sheet and exporting the final layout to PDF.

## Features

- Import Excel or CSV files directly in the browser.
- Support both one-candidate-per-row spreadsheets and workbook-style room plans with candidate lists, rows, and columns in each row.
- Map columns for candidate number, school name, syllabus code, component code, syllabus title, and workbook-style room plan headers.
- Generate a seating plan with draggable seats and draggable room fixtures.
- Edit candidate numbers and plan headings inline after generation.
- Export the completed layout to PDF, including multi-page exports for stacked room views.
- Keep saved room layouts in browser storage for up to 30 days.
- Deploy on GitHub Pages with no paid hosting.

## Expected spreadsheet data

The app supports either of these input styles:

- One candidate per row with columns like `School Name`, `Exam Series`, `Room Number`, `Paper`, `Syllabus Code`, `Component Code`, and `Candidate Number`.
- One plan per row with columns like `School Name`, `Exam Title`, `Room`, `Syllabus Title`, `Syllabus Code`, `Component Code`, `Candidates`, `Rows`, and `Cols`.

The app tries to auto-detect these headers, but you can change the mapping manually from the control panel.

## Sample file

The repository includes [seating_plan_template.xlsx](seating_plan_template.xlsx) as the bundled sample workbook used by the in-app sample download link.

## Local use

Because this is a static app, you can either:

- Open `index.html` directly in a browser.
- Serve the folder with any simple static server if you prefer.

## GitHub Pages deployment

1. Push this repository to GitHub.
2. Make sure the default branch is `main`, or update `.github/workflows/deploy.yml` if you publish from another branch.
3. In the repository settings, open `Pages` and set the source to `GitHub Actions`.
4. Push to `main` and the workflow in `.github/workflows/deploy.yml` will publish the site.
5. After the first successful run, GitHub Pages will provide the live site URL in the workflow summary and Pages settings.

## Notes

- The PDF export captures the generated plan exactly as arranged on screen.
- For best results, finish all drag and text edits before exporting.
- Loaded room layouts are also cached in the browser for 30 days, but the JSON export is still the reliable backup/share format.
- The app uses browser-side libraries from CDNs for Excel parsing and PDF export, so an internet connection is required when loading the page unless you later vendor those libraries locally.
