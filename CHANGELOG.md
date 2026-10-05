# Changelog

## 2.5.3

HA / MobiLink diagnostics and Option 14 parity release.

- Preserve the original broad Option 14 health-check intent while adding HA-specific diagnostics.
- Detect local IX Messaging role from installed services and refine topology from `DBA.LocationNodes`.
- Add role-aware checks for `MobiLink - Consolidated` and `SQL Anywhere - MobiLink Remote`.
- Report service state, startup type, Log On As identity, PID, process start time, and executable path.
- Correlate Service Control Manager startup, dependency, timeout, unexpected-termination, and credential/logon failures.
- Discover `Mobiclient.log` across documented and alternate UC paths.
- Use `Completed processing of download stream` as the documented file-sync success marker and report synchronization age.
- Add HEALTHY / WARNING / FAILED / UNKNOWN HA summary logic.
- Add Avaya Messaging 11.0 SP2 initial-sync guidance and the release-specific 10-day Primary-to-Consolidated recovery-window note.
- Restore explicit DBWatcher and UCArchiver checks from the original health check.
- Restore compact running/stopped UC-service and SQL Anywhere service summaries.
- Restore the last 20 DTMF buffer entries.
- Make ActiveSubscriptions guidance Primary-Consolidated-aware, including expected 1900-era upload timestamps for applicable topology entries.
- Treat confirmed single-server deployments as HA/MobiLink NOT APPLICABLE.
- Fix Windows PowerShell 5.1 `Sort-Object` syntax in the first 2.5.3 build.
- Fix Windows PowerShell 5.1 generic-list `Argument types do not match` handling in HA service/event collections.

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
- License expiration.
- Current-day VServer activity.
- VPIM, TiffConverter, Mutare, and TCP checks.
