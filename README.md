# Automated Monthly Reporting Monitoring System

## Overview

This project automates the monitoring of monthly reporting data across multiple Google Sheets and sends reminder emails if reports are incomplete.

It uses a centralized control sheet to manage multiple projects dynamically.

---

## Key Features

* Multi-project monitoring from a central control sheet
* Dynamic month mapping based on project start month
* Automated detection of incomplete data
* Email reminders sent to respective PULs
* Status tracking (email sent / not sent)
* Monthly execution using time-based triggers

---

## How It Works

1. The script runs on the 1st of every month
2. It checks the **previous month’s data**
3. Reads project details from the control sheet
4. Adjusts range dynamically based on:

   * Project start month
   * Standard reporting structure
5. Verifies if all cells are filled
6. If incomplete:

   * Sends email reminder
   * Updates status in control sheet

---

## Data Structure

### Control Sheet Columns:

* Project Name
* Sheet Link
* Sheet Name
* PUL Email
* Start Month
* Status (Active/Inactive)
* Last Checked
* Email Sent

---

## Month-to-Range Mapping

| Month     | Range    |
| --------- | -------- |
| April     | J5:K32   |
| May       | Q5:R32   |
| June      | X5:Y32   |
| July      | AH5:AI32 |
| August    | AO5:AP32 |
| September | AV5:AW32 |
| October   | BF5:BG32 |
| November  | BM5:BN32 |
| December  | BT5:BU32 |
| January   | CD5:CE32 |
| February  | CK5:CL32 |
| March     | CR5:CS32 |

---

## Tech Stack

* Google Apps Script (JavaScript)
* Google Sheets
* Gmail Automation

---

## Impact

* Eliminated manual follow-ups
* Improved reporting compliance
* Centralized monitoring system
* Scalable across multiple projects

---

## Automation

* Trigger: Monthly (1st of every month)
* Function: `checkAllProjects()`

---

## Author

Neetu Kharb

---

## Notes

This system was implemented in a real NGO environment to improve reporting efficiency and accountability.
