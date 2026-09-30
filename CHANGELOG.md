# Changelog

## 2.5.2

Audit / production hardening release.

- Restrict SQL submitted by the tool to a single SELECT statement.
- Reject DML/DDL, CALL, EXEC, SET, transaction commands, SQL comments, semicolons, SELECT INTO, and FOR UPDATE.
- Restrict automatic SQL Anywhere DSN probing to likely IX Messaging DSNs.
- Require approval before probing nonstandard SQL Anywhere DSNs.
- Add CSV formula-injection protection.
- Add optional PII redaction for CSV exports.
- Prefix exports with CONFIDENTIAL or REDACTED classifications.
- Move fallback exports away from shared C:\Temp to a user-scoped location.
- Add confirmation for retained-log searches longer than 31 days.
- Add a performance guard for expensive message-loop scans on very large MESSAGES tables.
- Clarify that the tool enforces SELECT-only SQL while effective database permissions are defined by the configured SQL Anywhere account.
- Add least-privilege and PowerShell transcription notes.
- Preserve PowerShell 5.1 compatibility.

## 2.5.1

PSScriptAnalyzer code-quality cleanup.

- Avoid assignment to PowerShell automatic variables.
- Fix null-comparison ordering.
- Remove unused Graph-audit variables/parameters.
- Replace empty catch blocks with Write-Verbose diagnostics.
- Keep the DataTable return workaround explicit for Windows PowerShell 5.1.

## 2.5.0

Added IX Messaging System Health Check (Option 14).

- Windows/resource checks.
- Core IXM services.
- IIS/application pools.
- SQL/database health indicators.
- license expiration.
- current-day VServer activity.
- VPIM, TiffConverter, Mutare, and TCP checks.
