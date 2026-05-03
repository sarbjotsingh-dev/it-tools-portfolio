# Headset Tracker

An equipment inventory management app for tracking headset assignments across employees and office locations.

## Features

- Assign headsets to employees with date and location tracking
- View current inventory status — available, assigned, or in repair
- Bulk CSV import for initial stock upload
- Export current inventory to CSV
- Filter by location, status, or employee name
- Dark mode support

## Tech Stack

- Vanilla HTML / CSS / JavaScript
- SharePoint REST API for data storage
- Fluent UI design tokens

## Setup

1. Deploy `index.html` and `app.js` to SharePoint Site Assets
2. Update `SITE_URL` in `app.js`
3. Create a `HeadsetInventory` SharePoint list with columns: `Title` (employee name), `SerialNumber`, `Model`, `Status`, `AssignedDate`, `Location`
