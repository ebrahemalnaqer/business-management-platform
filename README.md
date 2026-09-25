# Business Management & Operations Platform

## Overview

A web-based business management platform designed to centralize and streamline day-to-day business operations through an integrated system.

The platform combines customer management, attendance tracking, point-of-sale operations, invoicing, inventory management, financial records, room bookings, shift reconciliation, feedback, and operational analytics within a single application.

---

## Core Modules

### Customer Relationship Management

- Customer profiles and account management
- Customer identification and contact records
- Account types and hourly balances
- Duplicate customer validation

### Attendance Management

- Check-in and check-out
- Active session tracking
- Session duration calculation
- Automated usage charging
- Hourly package deduction

### Point of Sale

- Service and product management
- Product transactions
- Invoice creation
- Invoice items
- Discounts and tax calculation
- Inventory updates

### Inventory Management

- Service and product catalog
- Stock tracking
- Low-stock monitoring
- Automatic stock updates after sales

### Financial Management

- Revenue tracking
- Manual inflows
- Expense management
- Gross income calculation
- Net yield analysis
- Monthly financial metrics
- Daily revenue analysis
- Expense categorization

### Room Booking Management

- Room reservation management
- Booking validation
- Schedule management
- Operational room reporting

### Shift Management

- Shift opening and closing
- Expected revenue calculation
- Reported cash reconciliation
- Surplus / deficit detection

### Feedback & Customer Follow-up

- Customer feedback collection
- Average feedback monitoring
- Customer follow-up workflows

---

## Management Dashboard

The platform provides an operational dashboard containing key business indicators such as:

- Total customers
- Active sessions
- Gross income
- Net yield
- Invoice count
- Average feedback
- Low-stock alerts
- Daily revenue
- Expense categories
- Monthly performance metrics

---

## System Architecture

```text
                    Business Management Platform
                              │
             ┌────────────────┴────────────────┐
             │                                 │
        Frontend Web App                  Backend Engine
             │                                 │
             │                          Google Apps Script
             │                                 │
             └───────────────┬─────────────────┘
                             │
                        Data Layer
                             │
                       Google Sheets
                             │
        ┌────────────┬───────┼────────┬────────────┐
        │            │       │        │            │
       CRM       Attendance  POS    Finance    Bookings
        │            │       │        │            │
        └────────────┴───────┴────────┴────────────┘
                             │
                     KPI & Analytics
