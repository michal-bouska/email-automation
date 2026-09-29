# Email Automation Project

## Description

This project allows sending emails based on rules defined within Google Sheets. It also enables parsing payments from Fio Bank.

## Features

- Send emails based on conditions specified in Google Sheets.
- Parse and process transactions from Fio Bank API.
- Generate and insert QR codes into email content.

## Setup

1. Clone the repository.
2. Use the Google Sheets template.
3. Set up your Google Apps Script environment. 
4. Set the `FIO_API_TOKEN` variable in Script Properties to the token value from the Fio API.  
5. Optionally set `VARIABLE_SYMBOL_PREFIX_REGEX` in Script Properties to a regex for the prefix of variable symbols that belong to this project (e.g. `7240[0-9]` or `72405|72406`). It is anchored to the start of the symbol automatically. When it is empty, transactions are matched by exact equality with `VARIABLE_SYMBOL_KEY` (default `72405`).
6. Optionally set `REALIZATION_DOCUMENT_ID` in Script Properties to the ID of the spreadsheet holding the realization sheets (defaults to the active spreadsheet). All realization sheets live in this single document; the plan selects them by the `Sheet Name` column.

## Triggers

Mail sending and fio synchronization run on independent time-based triggers:

- `createMailTrigger()` / `deleteMailTrigger()` — hourly trigger for `processEmailsMonitored`.
- `createFioTrigger()` / `deleteFioTrigger()` — 5-minute trigger for `runUpdateMonitored`.

Run the create functions once from the Apps Script editor; each can be installed or removed without affecting the other.

## Logging

Besides the console, the mail script writes to two sheets in the active spreadsheet
(both are created automatically with a header row on first use):

- `log` — general events: `timestamp`, `name`, `severity`, `values` (JSON), `message`.
  Every run writes `mail_run_start` and `mail_run_end` (with the remaining daily email
  quota and the sent/error counts); every per-campaign or per-recipient failure is
  written as an `ERROR` row as it happens.
- `email_log` — one row per send attempt: `timestamp`, `email_template`, `recipient`,
  `quota_before`, `quota_after`, `status` (`sent` or `ERROR: …`), `payload` (JSON with
  the template data, addressing and row reference; bodies are not stored).

Use `logEvent_(name, severity, values, message)` to add your own events and
`checkMailQuota()` to print the remaining daily quota from the editor.
`MailApp.getRemainingDailyQuota()` needs the `script.send_mail` OAuth scope,
which is listed in `appsscript.json`; re-authorize the script after deploying.

## Usage

- Define your email rules and conditions in the Google Sheets.
- Run the script to send emails and process transactions.
- The "Column Sent" cell receives a timestamp on success or `ERROR: <message> at <time>` on failure.
  The script never changes column widths or existing cell formats.

# Fio
Set lock
```angular2html
curl https://fioapi.fio.cz/v1/rest/set-last-date/{$TOKEN}/2025-01-01/
```

## Acknowledgements

Special thanks to Martin Hawksey for the Mail Merger script, which served as the foundation for this project.