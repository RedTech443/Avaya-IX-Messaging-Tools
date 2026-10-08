# Avaya IX Messaging Tools

Read-only PowerShell troubleshooting and health-check toolkit for Avaya IX Messaging. Analyze voicemail deposits, caller activity, abandoned and silent/no-message scenarios, MWI, Microsoft Graph / Exchange integration, SQL Anywhere mailbox state, and high-availability MobiLink synchronization.

## Current release

**Version 2.5.6** — current version on `main`.

**Download:** [GitHub Releases](https://github.com/RedTech443/Avaya-IX-Messaging-Tools/releases) (look for `IXM-Tools-v2.5.6.ps1` in the release assets when the publishing workflow completes). The [current source script](IXM-Tools.ps1) is always available on `main`.

Repository contents:

- `IXM-Tools.ps1` — current PowerShell tool
- `README.md` — project overview
- `CHANGELOG.md` — release history
- `docs/User_Guide_v2.5.3.md` — previous stable guide (Option 15 not yet documented)

Older versioned scripts, checksums, analyzer settings, and older guides are intentionally not retained in the repository.

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
15. Inbound calls / abandoned voicemail analysis

## Option 15 — Inbound call and abandoned voicemail analysis (v2.5.6)

Option 15 reconstructs inbound call sessions from VServer `STATUS#YYYYMMDD.Log` records and helps answer what happened when a caller reached a mailbox but left no voicemail.

- Filter by caller ID and mailbox/extension; choose all calls or unsuccessful/uncertain calls.
- Show start time, duration, channel, recording attempts, rejected-too-short attempts, outcome and confidence.
- Classify greeting disconnects, calls with no recording, repeated too-short recordings, saved voicemails, and uncertain/incomplete sessions.
- Cross-check saved-voicemail evidence against the existing deposit parser using caller ID, mailbox and the call time window; normalize common US +1 caller-number formats.
- Summarize call outcomes and optionally export detailed evidence to CSV.
- Keep multiple recording attempts within the same call session instead of counting them as separate calls.

**Interpretation and limits:** A `RECORDING_TOO_SHORT` event identifies a recording attempt rejected by IX Messaging; it is not an audio-level measurement. A recording's elapsed duration does not independently establish what the caller said, nor can STATUS logs prove silence, media dead air or spam intent. Saved voicemail correlation relies on matching caller, mailbox, and time to the separate deposit parser, rather than a universally available channel-specific storage transaction. Investigate ambiguous results against the raw logs.

**Log discovery:** The application scans the local server's selected VServer log root. A Voice Server using `X:\\UC\\logs\\VServer` is not guaranteed to contain the same call records as another server using `E:\\UC\\logs\\VServer`. In HA environments, each server may have a different subset of calls. Only retained, uncompressed date-named STATUS logs matching the search period can be analyzed. One day of displayed coverage means only one matching date was discovered there, not necessarily a system-wide one-day retention policy.

### Getting started

1. Download the versioned `.ps1` file from [Releases](https://github.com/RedTech443/Avaya-IX-Messaging-Tools/releases) or use the current [source](IXM-Tools.ps1).
2. Run in Windows PowerShell 5.1 with read access to the appropriate IX Messaging logs.
3. Choose menu **15**, select a date range, and optionally enter the caller ID or destination mailbox.
4. Review the summary, then export a CSV if needed for a customer troubleshooting record.

The tool is intended for read-only investigation. Validate new call-flow classifications against representative customer logs before relying on them for incident conclusions.

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

## Notes

Results are limited by available log retention. Missing or rotated logs are not evidence that an event never occurred.

Health-check thresholds are diagnostic thresholds used by this utility and are not Avaya support policy.

This repository contains the current script, README, changelog, user guide, and a GitHub Actions workflow for publishing a versioned PowerShell asset to Releases. The user guide currently documents v2.5.3 and has not yet been updated for Option 15.
