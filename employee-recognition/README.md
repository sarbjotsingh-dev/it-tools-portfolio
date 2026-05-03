# Employee Recognition App ("High Five")

A peer-to-peer employee recognition system where staff can send appreciation messages ("High Fives") to colleagues, visible company-wide.

## Features

- Send a "High Five" to any employee with a custom message and category (Teamwork, Innovation, etc.)
- Company-wide recognition feed showing recent High Fives
- Leaderboard — most recognized employees this month
- Region-specific variants (Canada, India) with localized employee lists
- SharePoint-backed persistence

## Tech Stack

- Vanilla JavaScript
- SharePoint REST API
- PnP JS helper library

## Files

- `app.js` — core logic (US/Philippines variant)
- Additional regional variants available for other office locations

## Setup

1. Deploy JS files to SharePoint Site Assets
2. Update `SITE_URL` in each file
3. Required SharePoint list: `HighFives` — columns: `Sender`, `Recipient`, `Message`, `Category`, `Date`
