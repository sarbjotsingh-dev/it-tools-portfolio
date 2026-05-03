# New Hire Dashboard

A comprehensive multi-region new hire management dashboard with credential tracking, bulk import, analytics, and automated email workflows.

## Features

- Track all new hire credentials (M365, Five9, ILOAN, FusionID) in one place
- **Analytics bar** — total hires, this week, pending M365 setup, pending welcome emails
- **Department breakdown** chips with visual bars
- Bulk CSV import with per-row validation and inline editing before save
- Welcome email generator — pre-fills credentials, opens in email client
- LM Profile CSV export for Lead Manager system integration
- Filter by name, department, position, hire date
- Batch actions: M365 profile, Unifi account, assign manager, send email, assign station
- Dark mode + mobile-responsive

## Tech Stack

- Vanilla HTML / CSS / JavaScript
- SharePoint REST API
- Chart.js for analytics
- Fluent UI design tokens (CSS variables)

## Regional Variants

This repo includes a Region-IN (Region-IN) variant with region-specific fields and column mappings.

## Setup

1. Deploy `index.html` and `app.js` to SharePoint Site Assets
2. Update `SITE_URL` in `app.js`
3. Required SharePoint list: `NewHires` with columns matching the CSV template headers
