# IT New Hire Dashboard

A full-featured internal tool for IT teams to onboard new employees — track credentials, system access, and setup status across every hire.

## Features

- **Single Entry** — fill in and save one hire at a time with live search/autocomplete to load existing records
- **Bulk CSV Import** — upload a CSV of new hires with inline preview, row-level validation, and error highlighting
- **Record View** — filterable/sortable table of all hires with column-sticky layout, stats bar, and export
- **Batch Actions** — select multiple records to bulk-send welcome emails, assign M365 profiles, mark station IDs, or delete
- **IT Setup Email** — generate a pre-filled IT email with credentials and open it directly in your mail client
- **Dark Mode** — full Fluent UI dark theme toggle

## Tech Stack

- Vanilla HTML / CSS / JavaScript (no framework, no build step)
- Microsoft Fluent UI design tokens (CSS variables)
- SharePoint REST API for data persistence
- CSV parsing, export, and template download built-in

## Screenshots

> Add screenshots here after deploying

## Setup

1. Deploy `index.html` and `app.js` to your SharePoint Site Assets
2. Update the `SITE_URL` constant in `app.js` with your tenant URL
3. Ensure the SharePoint list `ITNewHires` exists with the required columns (see list schema in comments at top of `app.js`)
