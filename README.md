# Avaya IX Messaging Tools

Private troubleshooting toolkit for Avaya IX Messaging.

## Current release

**Version 2.5.0**

Main script:

- `vm-check-v2.5.0.ps1`

Documentation:

- `docs/User_Guide_v2.5.0.md`

## Requirements

- Windows PowerShell 5.1 or newer
- Run on an Avaya IX Messaging server, or from a system that can access the required IXM logs/database
- Read access to IX Messaging log directories
- SQL Anywhere ODBC DSN access for database-backed options

Typical IX Messaging paths:

- VServer logs: `X:\UC\logs\VServer`
- DBCOM logs: `X:\UC\logs\DBCOM`
- CSE Graph logs: `X:\UC\logs\uccse\CSE`

The script automatically searches fixed drives when the standard paths are not present.

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
14. IX Messaging system health check

## Highlights

### Voicemail and MWI

The tool correlates IX Messaging STATUS logs with RVSIP/SIP logs to reconstruct voicemail deposits and MWI activity. It distinguishes MWI OFF states from actual voicemail deletion; an MWI clear event alone is not treated as proof that a message was deleted.

### Graph / email synchronization

Option 10 correlates IXM message IDs with DBCOM activity. A message is shown as **CONFIRMED** only when `InternalUpdateSyncStatusOfMessage` contains a non-empty external `SyncID`.

### Mailbox database reports

Options 11 and 12 use read-only SQL Anywhere ODBC queries. The tool validates the IX Messaging schema and does not modify mailbox data.

### Graph / Exchange failure audit

Option 13 scans `ERR.SESGRFM.YYYY-MM-DDTHH.csv` CSE logs, groups failures by extension, and joins them to IX Messaging mailbox data. It distinguishes persistent mailbox/folder lookup failures, transient failures, Graph/API timeouts, authentication/permission failures, and stale CSE references.

### System health check

Option 14 performs a read-only IX Messaging health review including Windows resources, core services, IIS/application pools, database inventory, message-loop indicators, licensing, current-day VServer activity, VPIM events, TiffConverter exceptions, Mutare IIS activity, and TCP connection counts.

## Safety

Database-backed reports are intentionally read-only. The tool does not:

- restart services
- modify IIS
- update IX Messaging database records
- change mailbox configuration
- directly manipulate MWI database state

## CSV exports

Most reports can be exported to CSV. The default fallback export location is:

`C:\Temp\IXM-Reports`

## Notes

Results are limited by available log retention. Missing or rotated logs are not evidence that an event never occurred.

This repository is private and intended for internal troubleshooting/development use.
