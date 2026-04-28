<div align="center">

<!-- Banner image placeholder — replace with actual banner export -->
<!-- ![UAL Operations Parser Banner](docs/banner.png) -->

# UAL Operations Parser

**Forensic-grade M365 Unified Audit Log parser for Business Email Compromise investigations**

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-blue?style=flat-square)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-lightgrey?style=flat-square)]()
[![DFIR](https://img.shields.io/badge/Use%20Case-DFIR%20%7C%20BEC-E95555?style=flat-square)]()
[![Author](https://img.shields.io/badge/Author-Yuvi%20Kapoor-2EAD7A?style=flat-square)](https://linkedin.com/in/yuvi-kapoor-5a38521a5)

</div>

---

## Overview

**UAL Operations Parser** is a desktop DFIR tool built for extracting forensic intelligence from Microsoft 365 Unified Audit Log exports during Business Email Compromise engagements. It replaces the manual VBA-in-Excel workflow — drop in a raw UAL CSV or XLSX and receive a clean, categorised, investigator-ready output in one click.

Each row's `AuditData` JSON blob is decoded and routed to an operation-specific parser that extracts the fields that matter: email subjects, folder paths, inbox rule conditions, SharePoint file names, Internet Message IDs, and App/OAuth access context — all of which are buried inside nested JSON in the raw export.

Built by [Yuvi Kapoor](https://linkedin.com/in/yuvi-kapoor-5a38521a5)

---

## Features

- **UAL ingestion** — load raw M365 UAL exports directly from Microsoft Purview (CSV or XLSX)
- **Operation-specific parsing** — 9 dedicated parsers extract the right fields for each operation category rather than applying a generic approach across all rows
- **AuditData JSON decoding** — extracts Subject, ParentFolder Path, InternetMessageId, SourceFileName, ObjectId, rule Name/conditions, Actor UPN, and App/OAuth context from nested JSON blobs
- **AppAccessContext detection** — identifies OAuth app-token access records that produce no mail content and labels them correctly rather than flagging false warnings
- **Timestamp normalisation** — converts all M365 timestamp variants (ISO 8601, AU locale DD/MM/YYYY, US locale, Hawk/PowerShell exports) to a consistent `dd MMM yyyy HH:MM:SS` format; falls back to `AuditData.CreationTime` when the top-level column has been truncated by Excel
- **Structured error output** — three-tier error system (CRITICAL / WARNING / INFO) flags JSON parse failures, schema mismatches, and empty required fields; error rows are flagged in output rather than silently dropped
- **XLSX export** — styled, freeze-paned output with a Summary sheet, category counts, and a dedicated `⚠ Parse Errors` sheet; CreationTime written as a proper Excel datetime with custom `DD MMM YYYY HH:MM:SS` number format
- **CSV export** — clean UTF-8 CSV with a companion `_ERRORS.csv` when parse errors exist
- **HTML investigation report** — self-contained report auto-generated alongside every export; includes stat cards, inbox rule detail table, delete detail table, operation breakdown, top users, top source IPs, and a CRITICAL / HIGH / MEDIUM / LOW severity classification
- **Operation filter** — checkboxes to scope the export to specific operation categories before parsing
- **Parsed preview** — first 200 rows rendered inline with colour-coded category rows and error highlighting
- **Threaded processing** — parse and export run in background threads; the UI stays responsive throughout

---

## Operation Coverage

| Category | Operations |
|---|---|
| **Email — Create / Send / SendAs** | `Create`, `Send`, `SendAs`, `SendOnBehalf`, `MoveToFolder`, `Copy` |
| **Delete — Hard / Soft / Purge** | `HardDelete`, `SoftDelete`, `MoveToDeletedItems`, `RecoverDeletedMessages`, `PurgedMessages` |
| **Mail Items Accessed** | `MailItemsAccessed` |
| **File / Folder / SharePoint** | `FileModified`, `FileAccessed`, `FileDownloaded`, `FileUploadedPartial`, `FileModifiedExtended`, `SharingLinkUsed`, `PermissionLevelAdded`, `SharingInheritanceBroken`, `FolderRenamed`, and 30+ additional SharePoint / OneDrive operations |
| **Inbox Rules** | `New-InboxRule`, `Set-InboxRule`, `Remove-InboxRule`, `UpdateInboxRules`, `Set-SweepRule` |
| **Teams** | `TeamsSessionStarted`, `MessageCreatedHasLink`, `MessageEditedHasLink`, `ReactedToMessage`, `MemberAdded`, `AppInstalled`, and 20+ additional Teams operations |
| **Sign-In / Logon** | `UserLoggedIn`, `UserLoginFailed`, `PasswordLogonInitial`, `PasswordLogonFailed`, `SignInEvent` |
| **Admin / Config** | `Set-Mailbox`, `Add-MailboxPermission`, `Set-MailboxAutoReplyConfiguration`, `New-TransportRule`, `SiteCollectionCreated`, `Get-QuarantineMessage`, and 20+ additional admin cmdlets |

---

## Parsed Output Columns

| Column | Description |
|---|---|
| `CreationTime` | Timestamp — normalised to `dd MMM yyyy HH:MM:SS` from any M365 format |
| `Operation` | Raw operation name |
| `OperationCategory` | Classifier: `email_send`, `delete`, `mail_access`, `file_folder`, `inbox_rule`, `teams`, `sign_in`, `admin`, `generic` |
| `UserId` | UPN of the acting user |
| `ClientIP` | Source IP address |
| `Workload` | Exchange / SharePoint / MicrosoftTeams / AzureActiveDirectory |
| `ResultStatus` | Succeeded / Failed |
| `Subject` | Email subject, filename, rule name, or actor — context-dependent |
| `Path/Folder` | Mailbox folder path, SharePoint URL, or ObjectId |
| `InternetMessageId` | For `MailItemsAccessed` and delete operations — correlate via Message Trace |
| `Attachments` | Attachment names for email send operations |
| `Extra` | Rule conditions, AppId, ErrorNumber, protocol strings, aggregate record flags |
| `RawAuditData` | Original `AuditData` JSON preserved for reference |

---

## Error System

Every row produces structured parse diagnostics alongside the output. Errors are written to the `⚠ Parse Errors` sheet (XLSX) or `_ERRORS.csv` (CSV).

| Severity | Meaning |
|---|---|
| `CRITICAL` | `AuditData` could not be JSON-decoded, or the parser raised an exception. The row is still written with a `[JSON PARSE FAILURE]` marker in the Subject column so it is immediately visible. |
| `WARNING` | `AuditData` decoded successfully but a required field for the operation type is empty — indicates an unhandled schema variant. |
| `INFO` | Non-fatal notice — unknown operation routed to generic parser, or empty `AuditData` object on a synthetic record. |

---

## Downloads

> No installer required — unzip and run.

| Platform | Requirements |
|---|---|
| **Windows 10/11** | Python 3.10+ |
| **macOS 12+** (Intel & Apple Silicon) | Python 3.10+ from python.org |

---

## Installation

### 1. Install dependencies

```bash
pip install customtkinter openpyxl
```

### 2. Export a UAL from Microsoft Purview

1. Go to [Microsoft Purview compliance portal](https://compliance.microsoft.com) → **Audit**
2. Set your date range, users, and operations, then click **Search**
3. Once complete → **Export** → **Download all results**
4. Load the resulting `.csv` into UAL Operations Parser

📖 [Full export instructions — Microsoft documentation](https://learn.microsoft.com/en-us/purview/audit-log-export-records)

### 3. Run

**Windows**
```bat
launch.bat
```

**macOS / Linux**
```bash
chmod +x launch.sh && ./launch.sh
```

**Manual**
```bash
python ual_parser.py
```

The launchers auto-create a `.venv`, install dependencies, and launch the application.

---

## Dependencies

| Package | Purpose |
|---|---|
| `customtkinter` | Modern themed GUI framework |
| `openpyxl` | XLSX read/write with full formatting support |
| `Pillow` | Logo and icon rendering *(optional — icon falls back gracefully if absent)* |

---

## Usage

### Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| `Enter` | Run Parse |
| `Ctrl+E` | Export output |
| `Ctrl+O` | Open input file |
| `Ctrl+L` | Clear activity log |

### Platform Notes

**Windows** — works out of the box with standard Python from [python.org](https://www.python.org/downloads/).

**macOS** — install Python from [python.org](https://www.python.org/downloads/), not Homebrew. The python.org installer bundles a more reliable version of Tk. If the app opens behind other windows on first launch, click the icon in the Dock.

### Notes on MailItemsAccessed

`MailItemsAccessed` rows populate `InternetMessageId` and folder path from `AffectedItems` but not Subject — this is a limitation of the UAL data itself, not this tool. To recover email subjects, run a **Message Trace** in the Exchange Admin Centre and join on `InternetMessageId`. Aggregate sync records (where `ExternalAccess` is present but `AffectedItems` is empty) are labelled `AGGREGATE_RECORD` in the `Extra` column.

### Input Format Compatibility

| Source | Format | Notes |
|---|---|---|
| Microsoft Purview portal export | `.csv` (ISO 8601 timestamps) | Fully supported |
| Excel-resaved Purview export | `.csv` or `.xlsx` (AU/US locale timestamps) | Timestamps auto-normalised; seconds recovered from `AuditData.CreationTime` when truncated by Excel |
| PowerShell `Search-UnifiedAuditLog` | `.csv` (`Operations` column, US locale) | Column names normalised automatically |
| Hawk forensic tool | `.csv` (US 12-hour AM/PM timestamps) | Timestamps auto-normalised |

---

## HTML Report

Every export automatically generates a companion `_report.html` file. The report is self-contained and opens directly in any browser — no server or internet connection required.

Sections included:
- Summary stat cards for all 8 operation categories
- **Inbox Rule Operations** table — every rule event with timestamp, user, rule name, and extracted conditions (`ForwardTo`, `DeleteMessage`, etc.) — flagged CRITICAL
- **Delete Operations** table — hard and soft deletes with subjects and folder paths — flagged HIGH
- Operation breakdown with CRITICAL / HIGH / MEDIUM / LOW severity badges
- Category and result status breakdowns
- Top users by event count
- Top source IP addresses

> ⚠️ HTML reports contain personal information derived from audit logs. Distribute only to authorised personnel and handle in accordance with your engagement data handling policy.

---

## Data & Privacy

### What leaves your workstation

| Component | External contact | Data sent |
|---|---|---|
| UAL parsing | None — fully local | Nothing |
| XLSX / CSV export | None | Nothing |
| HTML report | None (opens in local browser) | Nothing unless you send the file |

### Privacy considerations

UAL exports contain personal information about every user whose activity was logged — names, email addresses, IP addresses, session identifiers, file access records, and authentication events.

| Jurisdiction | Instrument | Key consideration |
|---|---|---|
| Australia | Privacy Act 1988 / APPs | Collect only what is necessary; destroy when no longer required |
| European Union | GDPR | Data minimisation and purpose limitation apply |
| United Kingdom | UK GDPR / DPA 2018 | Same as EU GDPR in practice |
| United States (healthcare) | HIPAA | UAL logs from healthcare environments may touch PHI-adjacent systems |

> This is general guidance, not legal advice. Consult your firm's legal team for jurisdiction-specific obligations.

---

## File Structure

```
ual_parser/
├── ual_parser.py      # Main application — parsing engine + GUI
├── launch.bat         # Windows launcher (auto-venv + dependency install)
├── launch.sh          # macOS / Linux launcher
└── README.md
```

---

## Author

**Yuvi Kapoor**

Specialising in ransomware and BEC incident response engagements.

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Yuvi%20Kapoor-0A66C2?style=flat-square&logo=linkedin)](https://linkedin.com/in/yuvi-kapoor-5a38521a5)

---

## Licence

MIT — see [LICENSE](LICENSE)

---

<div align="center">

Built for the DFIR community

</div>
