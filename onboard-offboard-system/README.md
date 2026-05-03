# Onboard / Offboard System

A unified employee lifecycle management portal for IT teams — handles both new hire onboarding and departure offboarding workflows in a single interface.

## Features

- **Onboarding flow** — create M365 accounts, assign licenses, set up systems, send welcome emails
- **Offboarding flow** — revoke access, disable accounts, collect equipment, archive data
- **Task checklists** — per-employee status tracking with checklist items per department
- **Audit trail** — timestamped activity log per employee
- **Search** — find any employee by name or username instantly
- **Dark mode** + mobile-responsive layout

## Tech Stack

- Vanilla HTML / CSS / JavaScript
- SharePoint REST API
- Fluent UI design system

## Setup

1. Deploy `index.html` and `app.js` to SharePoint Site Assets
2. Update `SITE_URL` in `app.js` with your tenant URL
3. Required SharePoint lists: `OnboardingRecords`, `OffboardingRecords`, `EmployeeDirectory`
