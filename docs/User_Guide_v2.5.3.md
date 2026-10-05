# Avaya IX Messaging - Mailbox Activity & MWI Log Search

**User Guide | Script Version 2.5.3**

## Purpose

Troubleshoot IX Messaging voicemail deposits, MWI activity, Graph/email synchronization, mailbox configuration, Graph/Exchange failures, and overall IX Messaging system health.

Primary log path:

`X:\UC\logs\VServer`

Related paths:

- DBCOM: `<drive>:\UC\logs\DBCOM`
- CSE Graph errors: `<drive>:\UC\logs\uccse\CSE`

Platform:

- Windows PowerShell 5.1 or newer

Important: the tool reports only what can be proven from retained IX Messaging logs or the local IX Messaging database. Missing logs are not proof that an event did not occur.

---

## 1. Overview

The Mailbox Activity & MWI Log Search tool is an interactive PowerShell utility designed to simplify common IX Messaging troubleshooting tasks. It correlates voicemail deposits, MWI notifications, Graph/email synchronization, mailbox configuration, Graph/Exchange error activity, and system-health data into readable reports.

The current menu contains:

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
0. Exit

---

## 2. Starting the Script

Open Windows PowerShell as an account that can read the IX Messaging log directories and run the script from the directory where it is stored.

Example:

```powershell
PS C:\scripts> .\IXM-Tools-v2.5.3.ps1
```

### Automatic VServer log detection

The normal VServer log location is:

`X:\UC\logs\VServer`

If it is not present, the script searches local fixed drives for:

`\UC\logs\VServer`

If one location is found, it is selected automatically. If multiple locations are found, the script asks which one to use.

Manual entries may be a drive letter or a full path, for example:

```text
F:
H:
F:\UC\logs\VServer
```

A drive letter is automatically expanded to the standard VServer log path.

Database-backed options 11 and 12 remain usable even if VServer logs are unavailable. Options 13 and 14 can locate or use other IX Messaging data separately.

### Date-range selection

Most log-based options allow:

- Current day
- Last X days
- Custom date range

The tool shows retained-log coverage before the search where applicable. In v2.5.2, searches spanning more than 31 days require explicit confirmation to reduce unnecessary production disk I/O.

---

## 3. Menu Quick Reference

| Option | Function | Primary source | Scope | Purpose |
|---|---|---|---|---|
| 1 | Mailbox lifecycle | STATUS + RVSIP/SIP | One mailbox | Combined voicemail and MWI timeline |
| 2 | Voicemails left for one mailbox | STATUS | One mailbox | Successfully stored voicemail deposits |
| 3 | Voicemail / MWI clear history | RVSIP/SIP | One mailbox | MWI OFF / unread-clear events |
| 4 | All voicemails | STATUS | All mailboxes | All successful deposits in range |
| 5 | All MWI clear events | RVSIP/SIP | All mailboxes | All MWI OFF / clear events |
| 6 | Voicemail summary | STATUS | All mailboxes | Counts, last deposit, callers |
| 7 | Search by caller ID | STATUS | Caller number | Deposits from matching caller |
| 8 | Raw MWI history | RVSIP/SIP | One mailbox | All MWI ON/OFF events |
| 9 | Log coverage / diagnostics | STATUS/RVSIP/SIP | Date range | Available files and retention coverage |
| 10 | Graph / email sync history | STATUS + DBCOM | One mailbox | Message-to-Graph synchronization |
| 11 | Mailbox/email export | SQL Anywhere DB | All mailboxes | Name, extension, feature group, email |
| 12 | Current mailbox status / health | SQL Anywhere DB | One/all mailboxes | Current tutorial, MWI, Inbox and sync state |
| 13 | Graph / Exchange mailbox failure audit | CSE + SQL Anywhere DB | Affected mailboxes | Persistent/transient Graph mailbox failures |
| 14 | IX Messaging system health check + HA / MobiLink | Windows + IIS + SQL + logs | Server/system | Read-only overall health review with role-aware HA synchronization diagnostics |

---

## 4. Detailed Menu Options

### Option 1 - Mailbox lifecycle

Creates a chronological timeline for one mailbox by combining successfully stored voicemail deposits with MWI ON/OFF activity.

**Prompts:** mailbox/extension and date range.

**Data sources:** STATUS plus RVSIP/SIP.

Typical events include:

- VOICEMAIL SAVED
- MWI ON
- UNREAD CLEARED / MWI OFF
- EMPTY STATE / MWI OFF

Use this first when a user says a voicemail arrived but the lamp behaved unexpectedly.

### Option 2 - Voicemails left for one mailbox

Shows successfully stored voicemail deposits for one mailbox.

A voicemail is treated as saved only when the script finds the correlated successful message-add markers.

Typical fields include:

- date/time
- mailbox
- IXM mailbox ID
- caller ID
- caller name
- duration
- message file/GUID
- result
- source log

### Option 3 - Voicemail / MWI clear history

Finds MWI OFF events and interprets the message-summary counts.

**UNREAD CLEARED** means IX Messaging reported zero new messages while old/read messages may still remain.

**EMPTY STATE** means zero new and zero old voice messages.

Neither state proves that a specific voicemail was deleted.

### Option 4 - All voicemails

Lists every successfully stored voicemail deposit found in the selected range.

Useful for broad incident reviews or when the destination mailbox is unknown.

### Option 5 - All mailbox MWI clear events

Lists MWI OFF / clear activity for all mailboxes in the selected range.

Useful for determining whether an MWI condition affected one user or many users.

### Option 6 - Voicemail summary by mailbox

Summarizes successfully stored voicemail deposits by mailbox.

Typical fields:

- mailbox
- IXM mailbox ID
- voicemail count
- last deposit time
- unique caller IDs

### Option 7 - Search voicemails by caller ID

Searches successfully stored deposits by full or partial caller-ID value.

### Option 8 - Raw MWI history

Displays the underlying MWI ON/OFF signaling for one mailbox without reducing it to only clear events.

Typical fields:

- date/time
- extension
- MWI state
- new/old voice counts
- acknowledgement
- source/log file

### Option 9 - Log coverage / diagnostics

Displays which STATUS, RVSIP, and SIP files are actually available for the selected period.

This is important before interpreting a negative search. Missing logs mean the tool cannot prove that an event did not occur.

### Option 10 - Extension Graph / email sync history

Correlates voicemail messages with IX Messaging DBCOM activity.

Relevant logs include:

```text
<drive>:\UC\logs\DBCOM\EEAM_EEAMHELPER#YYYYMMDD.log
<drive>:\UC\logs\DBCOM\EEAM_TSECMGR#YYYYMMDD.log
```

A message is **CONFIRMED** when `InternalUpdateSyncStatusOfMessage` contains a non-empty external `SyncID`.

**FAILED** is used only when an explicit message-linked failure exists in retained logs.

**NOT CONFIRMED** means no confirming SyncID was found and is not automatically treated as a failure.

### Option 11 - Export mailboxes / email addresses

Reads mailbox configuration directly from the local IX Messaging SQL Anywhere database.

The tool restricts database statements to a single SELECT query. Effective database permissions are still controlled by the configured SQL Anywhere DSN/account.

Typical output:

- name
- extension
- feature group
- email address

The preferred email field is `MAILBOX.IMAPNAME`.

A populated email address does not by itself prove that Graph synchronization is enabled or healthy.

### Option 12 - Current mailbox status / health

Reads current mailbox state from SQL Anywhere.

Typical checks include:

- mailbox ID
- subscriber name
- extension
- feature group
- configured email
- tutorial state
- lock state
- failed PIN attempt count
- Inbox unread/read voice counts
- current MWI state
- MWI update time
- last Inbox sync
- last Calendar sync
- Graph/IMAP configuration

#### Tutorial state

`TUTORIAL=True` is reported as:

```text
Initial Tutorial: PENDING
Setup Status: INITIAL SETUP NOT COMPLETED
```

This does not claim that the subscriber has never entered the mailbox.

#### MWI consistency

Examples:

```text
Unread > 0 + MWI ON  = OK
Unread = 0 + MWI OFF = OK
Unread > 0 + MWI OFF = POSSIBLE MWI MISMATCH
Unread = 0 + MWI ON  = POSSIBLE STUCK MWI
```

These are troubleshooting indicators, not proof of a defect.

### Option 13 - Graph / Exchange mailbox failure audit

Scans CSE Graph error files:

```text
<drive>:\UC\logs\uccse\CSE\ERR.SESGRFM.YYYY-MM-DDTHH.csv
```

The audit extracts failing mailbox/extension sessions and correlates them with SQL mailbox data.

Typical fields include:

- extension
- subscriber name
- email address
- feature group
- Graph/IMAP configured
- last Inbox sync
- error code
- operation
- occurrence count
- approximate rate/hour
- first seen
- last seen
- priority
- assessment
- rationale
- recommendation

#### Graph mailbox/folder lookup failures

A stack path involving:

```text
Microsoft.Graph.MailFolderRequest.GetAsync
UC.CSE.MSGR.MailService.GetFolderId
```

is classified as a Graph mailbox/folder lookup failure.

The script deliberately does **not** label every `NullReferenceException` as an invalid email address.

#### Persistent failure

Repeated high-rate lookup failures with no successful Inbox sync are treated as persistent and can be shown as:

```text
PERSISTENT GRAPH MAILBOX FAILURE - VERIFY EXCHANGE/M365
```

#### Isolated / transient failure

One or two lookup failures with a previous successful Inbox sync are treated as:

```text
ISOLATED / TRANSIENT GRAPH FAILURE - MONITOR
```

This prevents a single Graph error from being treated the same as a continuously failing mailbox.

#### Timeouts

`ServiceException:timeout` is reported separately as a Graph/API timeout. A timeout alone is not evidence that the mailbox is invalid.

#### Authentication / permissions

Authentication or authorization errors are classified separately and direct troubleshooting toward the Graph application configuration and permissions.

### Option 14 - IX Messaging system health check + HA / MobiLink

Performs a read-only health review of the local IX Messaging server.

It does not restart services, modify IIS, change mailbox configuration, or intentionally submit database modification statements. Database statements are restricted by the tool to a single SELECT query.

#### System checks

- hostname
- Windows version
- uptime
- UC installation/version discovery
- UC root
- CPU
- physical memory
- fixed-drive free space
- SQL database file size

#### Services

- SQL Anywhere
- DBWatcher
- UCArchiver
- services stuck in StartPending
- local Voice Server role
- running/stopped UC services
- Dialogic where installed
- Nuance where installed
- RealSpeak where installed
- IIS World Wide Web service
- IIS application pools

Stopped UC services are shown for review because required services depend on system role and installed features.

#### Database checks

- LocationNodes
- server roles
- invalid `ServerType=0` entries
- task counts by task type
- Messages table count
- looping notification subject indicators
- ActiveSubscriptions view where applicable
- WebLM expiration date

#### Current-day VServer activity

The health check currently counts:

- `IDMS!` answered-call markers
- `MessageAdd succeeded` voicemail deposits
- `you have` TUI prompt occurrences
- `msg too short` occurrences
- DTMF buffer entries

The `you have` value is a prompt-occurrence estimate, not a unique-user login count.

#### Additional checks

- MobiLink/Mobiclient log where applicable
- recent UC VPIMServer terminations
- TiffConverter endpoint exceptions
- Mutare activity in the current IIS log
- TCP connection totals and state counts

#### HA / MobiLink diagnostics added in 2.5.3

Option 14 first preserves the original general server-health checks and then adds role-aware HA analysis.

The tool can identify or refine:

- Single Server
- Primary Voice
- Secondary Voice
- Primary Consolidated
- Secondary Consolidated
- Dedicated CSE / Report roles where present in LocationNodes

For HA systems it checks the role-appropriate MobiLink services. Avaya Messaging 11.0 SP2 documents:

- `MobiLink - Consolidated` on the Consolidated server
- `SQL Anywhere - MobiLink Remote` on Primary and Secondary Voice servers
- `DB\Logs\Mobiclient.log` as the file-sync log
- `Completed processing of download stream` as the successful synchronization-completion marker

The tool also reviews service startup mode and Log On As identities and searches recent Service Control Manager events for service-account/logon failures, startup failures, dependency failures, timeouts, and unexpected termination evidence.

A MobiLink service being in Running state does **not** by itself make HA healthy. The health summary also considers recent synchronization evidence from Mobiclient.log.

The script uses a 30-minute synchronization-freshness threshold as a diagnostic heuristic. This is not an Avaya support-policy threshold.

For Avaya Messaging 11.0 SP2, the HA installation guide warns that newly installed Primary/Secondary systems should not be logged into before the initial full synchronization is complete. The same HA chapter documents a release-specific 10-day recovery window for Primary-to-Consolidated synchronization loss before the system may revert to Demo Mode.

On a confirmed single-server deployment, HA/MobiLink checks are reported as NOT APPLICABLE rather than as missing-service failures.

#### Original Option 14 parity retained in 2.5.3

The 2.5.3 health check keeps the original diagnostic intent, including:

- DBWatcher and UCArchiver status
- broad SQL Anywhere service status
- compact running/stopped UC-service review
- IIS and application-pool status
- LocationNodes and role inventory
- Tasks and Messages table counts
- notification-loop indicators
- WebLM expiration
- current-day call, voicemail, TUI-prompt, short-message, and DTMF activity
- the last 20 DTMF buffer entries
- VPIM termination history
- TiffConverter and Mutare checks
- TCP state counts

`vw_ml_ActiveSubscriptions` is treated as a Primary Consolidated diagnostic. Applicable topology entries such as `ml_remote_consol_0`, CSE, Web, and Report can legitimately show an old/1900-era last-upload value and should be interpreted in topology context.


Health results are categorized as:

- OK
- WARNING
- ATTENTION
- INFO
- NOT APPLICABLE

Diagnostic thresholds are tool logic and should not be interpreted as Avaya support policy.

---

## 5. CSV Export

Most reporting options display results first and then prompt:

```text
Export these results to CSV? [y/N]
```

Enter `Y` or `Yes` to save the report.

If a Save As dialog is unavailable, the fallback directory is:

`%LOCALAPPDATA%\IXM-Tools\Reports` (with Documents/TEMP used only as later fallbacks when required)

---

## 6. Common Result Meanings

| Result | Meaning |
|---|---|
| VOICEMAIL SAVED | A successful IX Messaging message-add sequence was found |
| MWI ON | IX Messaging sent a message-summary notification showing new/waiting voicemail |
| UNREAD CLEARED / MWI OFF | Zero new voice messages were reported; read messages may remain |
| EMPTY STATE / MWI OFF | Zero new and zero old voice messages were reported |
| Acknowledged = True | Corresponding SIP acknowledgement was found |
| Graph CONFIRMED | A non-empty external SyncID was assigned |
| Graph FAILED | Explicit message-linked sync failure found |
| Graph NOT CONFIRMED | No confirming SyncID found in retained logs |
| PERSISTENT GRAPH MAILBOX FAILURE | Repeated mailbox/folder lookup failures requiring investigation |
| ISOLATED / TRANSIENT GRAPH FAILURE | Low-count failure with previous sync history; monitor |
| GRAPH/API TIMEOUT | Service/network/API timeout; does not prove mailbox invalid |

---

## 7. Suggested Troubleshooting Workflows

### MWI lamp problem

Start with Option 1.

Use Option 8 for raw signaling.

Use Option 3 for MWI OFF / clear events.

### Missing voicemail

Use Option 2.

If the mailbox is uncertain, use Option 4 or Option 7.

### Determine whether Graph/email synchronization occurred

Use Option 10.

A CONFIRMED result means an external SyncID was assigned.

If NOT CONFIRMED, verify retained-log coverage before concluding failure.

### Find Graph/Exchange accounts causing repeated failures

Use Option 13.

Focus first on persistent failures with blank or stale Inbox sync history.

Do not treat an isolated NullReferenceException as proof that an Exchange account is invalid.

### Review current mailbox configuration and MWI state

Use Option 12.

### Review overall server health

Use Option 14.

This is useful for quickly reviewing Windows resources, services, IIS, SQL health indicators, licensing, current-day activity, and selected application logs.

### Need a mailbox/email directory

Use Option 11.

### Search returned no results

Use Option 9 to verify that the required logs exist for the period being investigated.

---

## 8. Important Limitations

- Results are limited to retained server logs.
- Missing/rotated logs can prevent historical reconstruction.
- MWI OFF does not by itself identify how or why a message was read, moved, or deleted.
- Graph CONFIRMED proves an IX Messaging synchronization identifier was assigned; it does not prove that the recipient opened or read an email.
- An email address in `IMAPNAME` does not prove Graph is enabled or healthy.
- A `NullReferenceException` does not by itself prove that an Exchange mailbox is invalid.
- Option 14 thresholds are diagnostic heuristics and are not Avaya support policy.
- SQL database access used by the reporting functions is intentionally read-only.

---

## 9. Quick Start

1. Copy `IXM-Tools-v2.5.3.ps1` to the IX Messaging server.
2. Open Windows PowerShell.
3. Change to the script directory.
4. Run:

```powershell
.\IXM-Tools-v2.5.3.ps1
```

5. Confirm the detected log directory.
6. Select the menu option that matches the issue.
7. Enter mailbox/caller/date information where requested.
8. Review the on-screen findings.
9. Export to CSV when a report is needed.

---

End of Guide.


---

## 10. Version 2.5.2 Audit / Production Safety Controls

Version 2.5.2 adds additional safeguards intended for use on customer production systems.

### SQL statement restrictions

Database access generated by the tool is restricted to a single `SELECT` statement.

The SQL guard rejects:

- INSERT
- UPDATE
- DELETE
- ALTER
- DROP
- CREATE
- TRUNCATE
- MERGE
- CALL
- EXEC / EXECUTE
- SET
- transaction commands
- SELECT INTO
- FOR UPDATE
- SQL comments
- semicolon-separated or multiple statements

This is defense in depth inside the application. For the strongest security posture, the SQL Anywhere account associated with the DSN should also be granted SELECT-only permissions.

### SQL Anywhere DSN probing

The tool no longer automatically attempts connections to every SQL Anywhere System DSN.

Automatic probing is limited to DSN/database names that resemble IX Messaging. If only nonstandard SQL Anywhere DSNs are found, the operator must explicitly approve probing them.

### CSV export protection

CSV exports may contain customer information such as names, extensions, email addresses, caller IDs, mailbox IDs, message filenames, and diagnostic details.

Version 2.5.2:

- warns before exporting customer data;
- offers an optional redacted export mode;
- prefixes normal export filenames with `CONFIDENTIAL_`;
- prefixes redacted exports with `REDACTED_`;
- neutralizes spreadsheet-formula characters to reduce CSV formula-injection risk;
- uses a user-scoped fallback export directory instead of a shared `C:\Temp` path.

Reports should be protected and removed according to the customer's data-retention requirements.

### Production performance guardrails

Large retained-log searches spanning more than 31 days require explicit confirmation.

The health check also avoids automatically running expensive full-table message-loop subject scans when the `MESSAGES` table exceeds the configured safety threshold. The operator must approve those extended scans.

### Least privilege and PowerShell transcription

Administrator rights are not inherently required by the script. Run it under the least-privileged account that can read the necessary IX Messaging logs, service/IIS state, and SQL Anywhere DSN.

In environments where PowerShell transcription is enabled, console output may be recorded and can contain customer information.

### PSScriptAnalyzer

A companion `PSScriptAnalyzerSettings.psd1` file is included with the repository.

Example:

```powershell
Invoke-ScriptAnalyzer `
    -Path ".\IXM-Tools-v2.5.3.ps1" `
    -Settings ".\PSScriptAnalyzerSettings.psd1"
```

The settings suppress intentional interactive UI/naming conventions while leaving correctness and security-oriented rules enabled.

### Code signing

Authenticode signing is a deployment control and is not embedded in the script. For audited customer deployments, sign the release with the organization's approved PowerShell code-signing certificate and retain the published SHA-256 checksum with the release.


---

## 11. Version 2.5.3 HA / MobiLink Release Notes

Version 2.5.3 extends Option 14 with HA/MobiLink troubleshooting while preserving the original general health-check coverage. It adds role-aware service validation, service-account and SCM-event correlation, Mobiclient synchronization evidence, single-server handling, and PowerShell 5.1 compatibility fixes discovered during field testing.
