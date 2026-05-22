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

