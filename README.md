<div align="center">

# 🚀 Excel MC Data Cleaner

**An Excel VBA automation macro for cleaning carrier dispatch data, validating emails, extracting states from addresses, and formatting MC records automatically.**

Documented · MIT licensed · Maintained

![VBA](https://img.shields.io/badge/VBA-Latest-555555?style=for-the-badge)
[![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)](LICENSE)
[![PRs](https://img.shields.io/badge/PRs-welcome-brightgreen?style=for-the-badge)](CONTRIBUTING.md)

[Features](#-features) · [Quick Start](#-quick-start) · [Screenshots](#-screenshots) · [Contributing](CONTRIBUTING.md)

</div>

---

## 🖼 Screenshots

![Excel MC Data Cleaner screenshot](docs/screenshots/placeholder.svg)

*Replace `docs/screenshots/placeholder.svg` with real app screenshots.*

---

## 🐍 Contribution graph


<picture>
  <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/mafzalkalwardev/excel-mc-data-cleaner/output/snake-dark.svg" />
  <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/mafzalkalwardev/excel-mc-data-cleaner/output/snake.svg" />
  <img alt="Contribution snake" src="https://raw.githubusercontent.com/mafzalkalwardev/excel-mc-data-cleaner/output/snake.svg" />
</picture>


---

\# Excel MC Data Cleaner



An Excel VBA automation macro designed for cleaning and formatting carrier dispatch data automatically.



The macro validates emails, extracts states from addresses, formats records into structured tables, and prepares carrier data for dispatch workflows.



\## Features



\- Automatic email validation

\- State extraction from addresses

\- Carrier data cleaning

\- Excel table formatting

\- Automated worksheet generation

\- Phone and email processing

\- Dispatch workflow support

\- VBA automation



\## Tech Stack



\- Excel VBA

\- Microsoft Excel

\- Data Cleaning Automation



\## Features Overview



\### Email Validation



Rows with invalid or missing emails are automatically skipped.



\### State Extraction



Automatically extracts U.S. state abbreviations from addresses.



\### Automatic Table Formatting



Creates a formatted Excel table with cleaned records.



\### Dispatch Data Processing



Useful for:

\- carrier databases

\- dispatch workflows

\- CRM cleanup

\- lead processing



\## Output Columns



The macro generates:



| MC Number | Legal Name | States | Address | Phone | Email |

|-----------|------------|--------|---------|-------|-------|



\## Project Structure



```text

excel-mc-data-cleaner/

│

├── mc\_formatting\_macro.bas

├── README.md

└── .gitignore

```



\## How to Use



1\. Open Excel

2\. Press `ALT + F11`

3\. Import the `.bas` file

4\. Run:



```vba

MC\_formatting

```



\## Author



Muhammad Afzal Kalwar



GitHub:

@mafzalkalwardev

