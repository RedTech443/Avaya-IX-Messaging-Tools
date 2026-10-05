# Avaya IX Messaging Tools

Private troubleshooting and health-check toolkit for Avaya IX Messaging.

## Current release

**Version 2.5.3**

Current stable script:

- `IXM-Tools.ps1`

Versioned release:

- `IXM-Tools-v2.5.3.ps1`

Audit / verification files:

- `PSScriptAnalyzerSettings.psd1`
- `PSScriptAnalyzerSettings-v2.5.3.psd1`
- `IXM-Tools-v2.5.3.ps1.sha256.txt`

Documentation:

- `docs/User_Guide_v2.5.3.md`
- Previous guides/releases remain in the repository for reference.

## Requirements

- Windows PowerShell 5.1 or newer
- Avaya IX Messaging on Windows Server
- Read access to the required IX Messaging log directories
- SQL Anywhere ODBC DSN access for database-backed reports
- Administrator rights are not inherently required; use the least-privileged Windows account that has the required read visibility

Typical IX Messaging paths:

- VServer logs: `X:\UC\logs\VServer`
- DBCOM logs: `X:\UC\logs\DBCOM`
- CSE Graph logs: `X:\UC\logs\uccse\CSE`
- Database: `X:\UC\db\eeam21.db`

The script automatically searches fixed drives when the standard locations are not present.

## Menu

1. Mailbox lifecycle (voicemail deposits + MWI + clear events)
2. Voicemails left for one mailbox
3. Voicemail / MWI clear history for one mailbox
4. All voicemails left for all mailboxes
5. All mailbox MWI clear events
6. Voicemail summary by mailbox
7. Search voicemails by caller ID
8. Raw MWI history for one mailbox
9. Log coverage / diagnostics
10. Extension Graph / email sync history
11. Export mailboxes / email addresses
12. Current mailbox status / health
13. Graph / Exchange mailbox failure audit
14. IX Messaging system health check + HA / MobiLink

## Highlights

### Voicemail and MWI

Correlates IX Messaging STATUS logs with RVSIP/SIP logs to reconstruct voicemail deposits and MWI activity. MWI OFF states are not treated as proof that a voicemail was deleted.

### Graph / Exchange diagnostics

Option 10 correlates IXM message IDs with DBCOM activity and confirms synchronization when a non-empty external `SyncID` is recorded.

Option 13 scans CSE Graph error logs, groups failures by extension, joins them to mailbox data, and separates persistent mailbox/folder lookup failures, isolated/transient failures, Graph/API timeouts, and authentication/permission problems.

### Mailbox database reports

Options 11 and 12 use SQL Anywhere ODBC to report mailbox configuration and current state, including tutorial/setup status, lock state, PIN attempts, Inbox counts, MWI state, and Graph sync timestamps.

### System health check

Option 14 preserves the original broad IX Messaging health-check intent: Windows resources, SQL Anywhere, DBWatcher, UCArchiver, running/stopped UC services, IIS/application pools, topology, database indicators, licensing, current-day activity, DTMF, VPIM events, TiffConverter exceptions, Mutare IIS activity, and TCP connection statistics. Version 2.5.3 adds role-aware HA / MobiLink diagnostics without replacing those original checks.

## v2.5.3 HA / MobiLink diagnostics

Version 2.5.3 expands Option 14 while preserving the original health-check coverage.

- Detects Single Server, Primary/Secondary Voice, and Consolidated roles using installed services and `DBA.LocationNodes` where available
- Checks role-specific `MobiLink - Consolidated`, `SQL Anywhere - MobiLink Remote`, SQL Anywhere database, DBWatcher, and related services
- Reports service state, startup mode, Log On As account, PID, process start time, and executable path
- Correlates recent Service Control Manager startup, dependency, timeout, unexpected-termination, and service-account/logon failures
- Discovers `Mobiclient.log` under documented and alternate UC paths
- Uses `Completed processing of download stream` as the documented successful file-sync marker
- Reports synchronization age and recent MobiLink failure-pattern lines
- Produces an HA status of HEALTHY / WARNING / FAILED / UNKNOWN without assuming that a running service proves successful synchronization
- Restores explicit DBWatcher and UCArchiver checks, a compact running/stopped UC-service overview, broad SQL Anywhere service status, and the last 20 DTMF buffer entries
- Treats `vw_ml_ActiveSubscriptions` as a Primary Consolidated diagnostic and retains topology guidance for expected 1900-era upload timestamps
- Treats HA/MobiLink as NOT APPLICABLE on a confirmed single-server deployment
- Includes Avaya Messaging 11.0 SP2 HA guidance for initial synchronization and the release-specific 10-day Primary-to-Consolidated recovery window

The 30-minute sync-freshness threshold used by the tool is a diagnostic threshold, not an Avaya support-policy threshold.

## v2.5.2 audit hardening

Version 2.5.2 adds several production-safety controls:

- SQL submitted by the tool is restricted to one `SELECT` statement
- DML/DDL, `CALL`, `EXEC`, `SET`, transactions, SQL comments, semicolons, `SELECT INTO`, and `FOR UPDATE` are rejected
- automatic SQL Anywhere probing is limited to likely IX Messaging DSNs; nonstandard DSNs require approval
- CSV formula-injection protection
- optional CSV PII redaction
- `CONFIDENTIAL_` / `REDACTED_` export naming
- user-scoped fallback export directory instead of a shared `C:\Temp` folder
- confirmation before very large retained-log scans
- guardrail before expensive message-loop searches on very large message tables
- clearer distinction between application-enforced SELECT-only SQL and the permissions granted to the configured database account

For the strongest audit posture, use a SQL Anywhere identity that itself has SELECT-only permissions and Authenticode-sign the script with the organization's code-signing certificate.

## Safety

The utility is intended as non-destructive diagnostic software. It does not normally:

- restart or stop Windows services
- modify IIS configuration
- modify mailbox records
- change subscriber configuration
- directly manipulate MWI database state
- modify firewall rules
- modify the Windows registry
- delete IX Messaging logs
- download software from the Internet
- make external Graph or Mutare API calls

Graph and Mutare diagnostics read existing IX Messaging/CSE/DBCOM/IIS logs.

## CSV exports

Most reports can be exported to CSV. The tool warns when customer data may be exported and offers optional redaction.

When the Save As dialog is unavailable, the fallback location is user-scoped under Local AppData (or Documents/TEMP as a final fallback), rather than `C:\Temp`.

## PSScriptAnalyzer

Run the audit with:

```powershell
Invoke-ScriptAnalyzer `
    -Path ".\IXM-Tools-v2.5.3.ps1" `
    -Settings ".\PSScriptAnalyzerSettings.psd1"
```

The settings suppress intentional interactive-UI/naming rules while leaving correctness/security-related warnings enabled.

## SHA-256

The published checksum for `IXM-Tools-v2.5.3.ps1` is stored in:

`IXM-Tools-v2.5.3.ps1.sha256.txt`

## Notes

Results are limited by available log retention. Missing or rotated logs are not evidence that an event never occurred.

Health-check thresholds are diagnostic thresholds used by this utility and are not Avaya support policy.

This repository is private and intended for internal troubleshooting/development use.
