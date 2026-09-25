# Techno Club — Business Management & Operations Platform

A web-based business management platform developed for Techno Club to centralize customer management, attendance, POS, invoicing, inventory, financial records, shift reconciliation, room bookings, feedback, and operational KPI reporting.

## Overview

The system was designed as an internal operations platform for a technology-driven coworking and business workspace. It combines a Google Apps Script backend with an HTML/JavaScript web interface and Google Sheets as the operational data layer.

## Core Modules

- **Dashboard** — operational and financial KPIs, revenue trends, expenses, active sessions, inventory alerts, and feedback metrics.
- **Clients / CRM** — client records, phone-based unique identification, account types, lead sources, and hourly balances.
- **Attendance** — check-in/check-out, session duration, charging calculations, and package-hour deductions.
- **Point of Sale** — services/products, quantities, pricing, and transaction processing.
- **Invoices** — invoice creation, invoice items, discounts, taxes, payment methods, and stock updates.
- **Services & Stock** — service/product catalog, cost and selling prices, capacity/stock, and low-stock thresholds.
- **Room Bookings** — room availability validation, pricing, discounts, booking status, and Google Calendar synchronization.
- **Financial Ledger** — POS revenue, manual inflows, expenses, gross income, net yield, and category-level expense analysis.
- **Shift Reports** — expected-vs-reported cash reconciliation and variance status.
- **Feedback** — customer ratings, messages, and feedback status management.
- **Follow-ups** — customer follow-up tracking and inactive-client monitoring.
- **Room Reports** — booking counts, booked hours, revenue, discounts, and occupancy metrics.

## Architecture

```text
Google Apps Script Backend
          │
          ├── CRM / Clients
          ├── Attendance
          ├── POS & Invoices
          ├── Services & Stock
          ├── Finance & Expenses
          ├── Shift Reconciliation
          ├── Feedback
          └── Room Booking / Calendar Sync
                    │
                    ▼
             Google Sheets
                    │
                    ▼
          HTML / JavaScript Web App
                    │
                    ├── Dashboard KPIs
                    ├── Charts
                    └── Operational Interface
```

## Technology Stack

- Google Apps Script
- JavaScript
- HTML5
- Tailwind CSS
- Chart.js
- Font Awesome
- Google Sheets
- Google Calendar

## Engineering Highlights

### CRM data integrity

The client module uses the phone number as the primary client identifier and includes a duplicate-registration guard before inserting a new client.

### Automated attendance and charging

Check-in and check-out operations calculate session duration and charges. Hourly-package balances can also be deducted based on session duration.

### POS and inventory integration

Invoice creation calculates subtotal, discount, tax, and total values. Product/cafe items can automatically reduce the corresponding stock quantity.

### Financial aggregation

The backend aggregates POS revenue, manual inflows, expenses, gross income, net yield, monthly metrics, daily revenue, and expense categories for management reporting.

### Shift reconciliation

The system calculates expected daily revenue from recorded invoices and inflows, compares it with reported cash, subtracts expenses where applicable, and returns a reconciliation status such as Balanced, Surplus, or Deficit.

### Room booking and calendar validation

Room bookings include availability checks against Google Calendar before creating a reservation. The booking module also calculates room pricing, discounts, final amounts, and stores booking information for reporting.

## Public Repository Note

This repository is a public showcase of the system architecture and implementation approach.

For security and privacy, environment-specific Google Calendar IDs are replaced with placeholders in the public `Code.gs` file. Configure your own calendar IDs before deploying the booking module.

Do not commit:

- Customer records
- Real phone numbers
- Private Google Sheet IDs
- Credentials or API keys
- Private Calendar IDs if your organization's policy requires them to remain private
- Production exports or financial records

## Deployment

This project is designed for Google Apps Script.

1. Create a Google Apps Script project.
2. Add `Code.gs` and `index.html`.
3. Connect the script to the required Google Sheets data source.
4. Configure your own Google Calendar IDs in `ROOM_CALENDARS`.
5. Review and authorize the required Apps Script services.
6. Deploy the project as a web app according to your organization's access requirements.

## Project Context

**Organization:** Techno Club  
**Project Type:** Business Management & Operations Platform  
**Primary Focus:** CRM, POS, Automation, Operations, Financial Tracking, Booking Management, and KPI Reporting

## Author

**Ebraheem Alnaqer**  
Computer Science & AI | Technology & Data | Business & Operations

- LinkedIn: https://linkedin.com/in/ebrahemalnaqer
- GitHub: https://github.com/ebrahemalnaqer
