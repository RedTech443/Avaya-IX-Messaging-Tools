#requires -version 5.1

<#
.SYNOPSIS
    Avaya IX Messaging - Mailbox Activity & MWI Log Search

.DESCRIPTION
    Interactive troubleshooting tool for Avaya IX Messaging VServer logs.

    Capabilities:
      - Combined mailbox lifecycle timeline
      - Successfully deposited voicemail search for one mailbox
      - Voicemail/MWI clear history for one mailbox
      - All successfully deposited voicemails for a date range
      - All MWI clear events across all mailboxes
      - Voicemail summary by mailbox
      - Search voicemail deposits by caller ID
      - Raw MWI ON/OFF history for one mailbox
      - Extension Graph/email sync history from DBCOM
      - Last confirmed SyncID and explicit sync failures by extension
      - Mailbox directory export (name, extension, feature group, email address)
      - Current mailbox status / health from the IX Messaging SQL database
      - Graph / Exchange mailbox failure audit from CSE ERR.SESGRFM logs
      - Graph failure correlation to mailbox name, extension, configured email,
        feature group, Graph/IMAP setting, and last successful Inbox sync
      - IX Messaging system health check covering Windows, services, IIS,
        database inventory, replication, licensing, message-loop indicators,
        current-day VServer activity, VPIM events, fax exceptions, Mutare IIS
        activity, and TCP connection counts
      - Automatic SQL Anywhere / IX Messaging DSN discovery for database-backed reports
      - Log coverage / diagnostics
      - Date prompts show retained-log coverage and maximum searchable lookback
      - Optional CSV export

    Voicemail deposits are identified from STATUS logs by correlating:
      INMSGSTART / INMSGEND
      EEAM.GetPlayTime
      FastMessageAdd
      XEEAM_MessageAdd succeeded
      MbxNo / MbxID
      msgrec.FileName

    MWI activity is identified from RVSIP outbound SIP NOTIFY messages when
    available, with SIP SetMWI as a fallback. MWI OFF events are classified
    as mailbox/unread clear events from the message-summary counts. The tool
    does not claim a user deleted a message unless the logs explicitly prove it.

    Graph/email sync history is correlated from:
      DBCOM\EEAM_EEAMHELPER#YYYYMMDD.log
      DBCOM\EEAM_TSECMGR#YYYYMMDD.log

    A populated SyncID on InternalUpdateSyncStatusOfMessage is treated as
    CONFIRMED external synchronization. A missing SyncID is NOT CONFIRMED
    unless an explicit message-linked failure is present in the retained logs.

    Graph / Exchange failure auditing reads:
      uccse\CSE\ERR.SESGRFM.YYYY-MM-DDTHH.csv

    The audit does not call an email address "invalid" solely because IX Messaging
    logged a NullReferenceException. It reports the Graph operation that failed
    when it can be identified and recommends verification of the Exchange/M365
    mailbox and Graph access. The tool submits SELECT-only SQL; effective database permissions are defined by the configured SQL Anywhere account.

    Version 2.5.1 includes a PSScriptAnalyzer code-quality cleanup. It avoids
    assignments to PowerShell automatic variables, removes unused Graph-audit
    inputs, uses null-safe comparisons, makes intentionally ignored exceptions
    visible through Write-Verbose, and keeps the DataTable return workaround
    explicit for Windows PowerShell 5.1.

    Version 2.5.2 adds audit-hardening controls:
      - SQL is restricted to a single SELECT statement with no comments,
        semicolons, CALL/EXEC/SET/DDL/DML, SELECT INTO, or transaction commands.
      - Only likely IX Messaging SQL Anywhere DSNs are connected to by default;
        probing nonstandard SQL Anywhere DSNs requires explicit approval.
      - CSV export neutralizes spreadsheet-formula injection, offers optional
        PII redaction, labels default filenames CONFIDENTIAL, and uses a
        user-scoped fallback export directory instead of C:\Temp.
      - Large retained-log searches require explicit confirmation.
      - Expensive full-table message-loop health scans are skipped by default
        on very large message tables unless explicitly approved.
      - Database wording now distinguishes application-enforced SELECT-only SQL
        from the permissions granted to the configured SQL Anywhere account.

    Version 2.5.3 expands the system health check with role-aware IX Messaging
    HA / MobiLink diagnostics:
      - Detects Consolidated versus Voice-server roles without hard-coded hostnames
        and refines Primary/Secondary role from DBA.LocationNodes when available.
      - Verifies role-specific MobiLink, SQL Anywhere, DBWatcher, and related UC
        services with state, startup mode, logon identity, PID, process start time,
        executable path, and service-account correlation.
      - Reviews recent Service Control Manager failures, including logon/password,
        startup, dependency, timeout, and unexpected-termination evidence.
      - Discovers Mobiclient.log across documented and alternate UC paths, reports
        recent successful download-stream markers, sync age, and recent errors.
      - Produces HEALTHY / WARNING / FAILED / UNKNOWN HA status without equating a
        running MobiLink service by itself with successful synchronization.
      - Incorporates Avaya Messaging 11.0 SP2 HA guidance: page 203's initial
        synchronization warning and documented success marker, plus the HA chapter's
        10-day Primary-to-Consolidated synchronization recovery window. These are
        presented as operational guidance and do not change the tool's diagnostic
        30-minute sync-freshness threshold.
      - Fixes the Windows PowerShell 5.1 Sort-Object syntax used by the service-account
        summary in the first 2.5.3 build.
      - Fixes Windows PowerShell 5.1 generic-list array conversion in the HA service
        inventory path (prevents runtime 'Argument types do not match').
      - Restores original health-check parity for DBWatcher/UCArchiver, broad SQL
        Anywhere and UC-service status review, and the last 20 DTMF buffer entries.
      - Keeps ActiveSubscriptions as a Primary Consolidated diagnostic and restores
        the original guidance that some CSE/Web/Report/consolidated-remote upload
        timestamps can legitimately remain at 1900-01-01 depending on topology.
      - Refines topology before role-specific service checks so a true single-server
        installation is not incorrectly failed for missing HA/MobiLink components.

    Authenticode signing is a deployment control and is not added automatically.
    For audited customer deployments, sign the release with the organization's
    code-signing certificate and use a SQL Anywhere identity limited to SELECT.

.NOTES
    Default log location:
        X:\UC\logs\VServer

    Designed for Windows PowerShell 5.1 and newer.
#>

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------

$ToolVersion = '2.5.5'
$LogRoot = 'X:\UC\logs\VServer'

# CSV fallback is intentionally user-scoped rather than a shared C:\Temp path.
$LocalAppData = [Environment]::GetFolderPath('LocalApplicationData')
if (-not [string]::IsNullOrWhiteSpace($LocalAppData)) {
    $ExportRoot = Join-Path $LocalAppData 'IXM-Tools\Reports'
}
else {
    $Documents = [Environment]::GetFolderPath('MyDocuments')
    if (-not [string]::IsNullOrWhiteSpace($Documents)) {
        $ExportRoot = Join-Path $Documents 'IXM-Reports'
    }
    else {
        $ExportRoot = Join-Path $env:TEMP 'IXM-Reports'
    }
}

$script:LargeSearchWarningDays = 31
$script:HealthMessageLoopScanThreshold = 100000

$script:DepositCache = @{}
$script:MWICache = @{}
$script:AllMWICache = @{}
$script:GraphSyncCache = @{}
$script:LastExportDirectory = $null
$script:CseGraphRoot = $null

# -----------------------------------------------------------------------------
# Console helpers
# -----------------------------------------------------------------------------

function Write-Section {
    param([Parameter(Mandatory)][string]$Title)

    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host ('  ' + $Title) -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host ''
}

function Pause-Tool {
    Write-Host ''
    [void](Read-Host 'Press Enter to continue')
}

function Read-NumericValue {
    param([Parameter(Mandatory)][string]$Prompt)

    do {
        $Value = (Read-Host $Prompt).Trim()
        if ($Value -notmatch '^\d+$') {
            Write-Host 'Please enter digits only.' -ForegroundColor Yellow
        }
    } until ($Value -match '^\d+$')

    return $Value
}

function Confirm-IxmLargeSearchRange {
    param(
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate,
        [string]$Context = 'log search',
        [int]$WarnAfterDays = $script:LargeSearchWarningDays
    )

    $Days = [int](($EndDate.Date - $StartDate.Date).TotalDays + 1)
    if ($Days -le $WarnAfterDays) {
        return
    }

    Write-Host ''
    Write-Host ('PRODUCTION LOAD WARNING: {0} spans {1} day(s).' -f $Context,$Days) -ForegroundColor Yellow
    Write-Host 'Large retained-log scans can create significant disk I/O on a busy IX Messaging server.' -ForegroundColor Yellow
    Write-Host ('The normal warning threshold for this tool is {0} day(s).' -f $WarnAfterDays) -ForegroundColor Yellow

    $Approval = (Read-Host 'Type CONTINUE to run this large search, or press Enter to cancel').Trim()
    if ($Approval -cne 'CONTINUE') {
        throw [System.OperationCanceledException]::new('Large search canceled by operator.')
    }
}

function Read-DateRange {
    param(
        [ValidateSet('STATUS','MWI','LIFECYCLE','ALL')]
        [string]$CoverageMode = 'ALL'
    )

    $Coverage = Get-SearchCoverage -Mode $CoverageMode

    Write-Host ''
    if ($Coverage.HasData) {
        Write-Host ('Available {0} log coverage:' -f $Coverage.Label) -ForegroundColor Cyan
        Write-Host ('  Oldest retained date  : {0}' -f $Coverage.Oldest.ToString('MM/dd/yyyy'))
        Write-Host ('  Newest retained date  : {0}' -f $Coverage.Newest.ToString('MM/dd/yyyy'))
        Write-Host ('  Dated log days        : {0}' -f $Coverage.DateCount)
        Write-Host ('  Calendar span         : {0} day(s)' -f $Coverage.CalendarSpanDays)
        Write-Host ('  Max lookback from today: {0} day(s)' -f $Coverage.MaxLookbackDays) -ForegroundColor Green

        if ($Coverage.MissingDates.Count -gt 0) {
            $Preview = @($Coverage.MissingDates | Select-Object -First 8 | ForEach-Object { $_.ToString('MM/dd/yyyy') })
            $Suffix = ''
            if ($Coverage.MissingDates.Count -gt 8) {
                $Suffix = (' ... +{0} more' -f ($Coverage.MissingDates.Count - 8))
            }
            Write-Host ('  Missing dated log days: {0}{1}' -f ($Preview -join ', '),$Suffix) -ForegroundColor Yellow
        }

        if (-not $Coverage.HasToday) {
            Write-Host ('  WARNING: No matching dated log file is present for today ({0}).' -f (Get-Date).Date.ToString('MM/dd/yyyy')) -ForegroundColor Yellow
        }
    }
    else {
        Write-Host ('No dated {0} logs were found. Date searches will still run, but coverage cannot be estimated.' -f $Coverage.Label) -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host 'Search period:' -ForegroundColor Cyan
    Write-Host '  1. Current day'
    if ($Coverage.HasData) {
        Write-Host ('  2. Last X days (maximum useful lookback: {0})' -f $Coverage.MaxLookbackDays)
    }
    else {
        Write-Host '  2. Last X days'
    }
    Write-Host '  3. Custom date range'
    Write-Host ''

    $Choice = ''
    while ($Choice -notin @('1','2','3')) {
        $Choice = (Read-Host 'Select 1, 2, or 3').Trim()
    }

    $Today = (Get-Date).Date

    if ($Choice -eq '1') {
        if ($Coverage.HasData -and -not $Coverage.HasToday) {
            Write-Host ('WARNING: No matching dated log file exists for today. Newest available date is {0}.' -f $Coverage.Newest.ToString('MM/dd/yyyy')) -ForegroundColor Yellow
        }

        $Result = New-Object PSObject -Property @{
            Start = $Today
            End = $Today
            Description = 'today'
        }
        return $Result
    }

    if ($Choice -eq '2') {
        $DefaultDays = 7
        if ($Coverage.HasData -and $Coverage.MaxLookbackDays -lt $DefaultDays) {
            $DefaultDays = $Coverage.MaxLookbackDays
        }
        if ($DefaultDays -lt 1) { $DefaultDays = 1 }

        $Days = 0
        while ($Days -le 0) {
            $Text = (Read-Host ('How many days? [Default: {0}]' -f $DefaultDays)).Trim()

            if ([string]::IsNullOrWhiteSpace($Text)) {
                $Days = $DefaultDays
            }
            elseif ($Text -match '^\d+$') {
                $Days = [int]$Text
                if ($Days -le 0) {
                    Write-Host 'Enter a number greater than zero.' -ForegroundColor Yellow
                }
                elseif ($Coverage.HasData -and $Days -gt $Coverage.MaxLookbackDays) {
                    Write-Host ('Only {0} day(s) can be searched from today based on the oldest retained {1} log date ({2}).' -f $Coverage.MaxLookbackDays,$Coverage.Label,$Coverage.Oldest.ToString('MM/dd/yyyy')) -ForegroundColor Yellow
                    $Days = 0
                }
            }
            else {
                $Days = 0
                Write-Host 'Enter a number greater than zero.' -ForegroundColor Yellow
            }
        }

        $SelectedStart = $Today.AddDays(-($Days - 1))
        Confirm-IxmLargeSearchRange -StartDate $SelectedStart -EndDate $Today -Context ('{0} retained-log search' -f $Coverage.Label)

        $Result = New-Object PSObject -Property @{
            Start = $SelectedStart
            End = $Today
            Description = ('last {0} day(s)' -f $Days)
        }
        return $Result
    }

    $StartDate = $null
    while ($null -eq $StartDate) {
        $StartText = (Read-Host 'Start date (MM/dd/yyyy or yyyy-MM-dd)').Trim()
        try {
            $StartDate = ([datetime]$StartText).Date
        }
        catch {
            $StartDate = $null
            Write-Host 'Invalid start date.' -ForegroundColor Yellow
        }
    }

    $EndDate = $null
    while ($null -eq $EndDate) {
        $EndText = (Read-Host 'End date (MM/dd/yyyy or yyyy-MM-dd)').Trim()
        try {
            $CandidateEnd = ([datetime]$EndText).Date
            if ($CandidateEnd -lt $StartDate) {
                Write-Host 'End date must be on or after the start date.' -ForegroundColor Yellow
            }
            else {
                $EndDate = $CandidateEnd
            }
        }
        catch {
            $EndDate = $null
            Write-Host 'Invalid end date.' -ForegroundColor Yellow
        }
    }

    if ($Coverage.HasData -and ($StartDate -lt $Coverage.Oldest -or $EndDate -gt $Coverage.Newest)) {
        Write-Host ('WARNING: Selected range extends beyond retained {0} log dates ({1} through {2}). Results will only reflect available logs.' -f $Coverage.Label,$Coverage.Oldest.ToString('MM/dd/yyyy'),$Coverage.Newest.ToString('MM/dd/yyyy')) -ForegroundColor Yellow
    }

    Confirm-IxmLargeSearchRange -StartDate $StartDate -EndDate $EndDate -Context ('{0} retained-log search' -f $Coverage.Label)

    $Result = New-Object PSObject -Property @{
        Start = $StartDate
        End = $EndDate
        Description = ('{0} through {1}' -f $StartDate.ToString('MM/dd/yyyy'), $EndDate.ToString('MM/dd/yyyy'))
    }
    return $Result
}

# -----------------------------------------------------------------------------
# File helpers
# -----------------------------------------------------------------------------

function New-SharedReader {
    param([Parameter(Mandatory)][string]$Path)

    $Share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $Stream = New-Object System.IO.FileStream -ArgumentList @(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        $Share
    )

    $Reader = New-Object System.IO.StreamReader -ArgumentList $Stream

    return [pscustomobject]@{
        Stream = $Stream
        Reader = $Reader
    }
}

function Get-DatedLogs {
    param(
        [Parameter(Mandatory)][ValidateSet('STATUS','RVSIP','SIP')][string]$Type,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    $Pattern = '^{0}#(?<Date>\d{{8}})\.log$' -f [regex]::Escape($Type)
    $Results = New-Object System.Collections.Generic.List[object]

    foreach ($File in Get-ChildItem -LiteralPath $LogRoot -File -ErrorAction SilentlyContinue) {
        if ($File.Name -match $Pattern) {
            try {
                $Date = [datetime]::ParseExact(
                    $Matches.Date,
                    'yyyyMMdd',
                    [System.Globalization.CultureInfo]::InvariantCulture
                ).Date

                if ($Date -ge $StartDate.Date -and $Date -le $EndDate.Date) {
                    $Results.Add([pscustomobject]@{
                        Type = $Type
                        Date = $Date
                        Name = $File.Name
                        Path = $File.FullName
                        Length = $File.Length
                        LastWriteTime = $File.LastWriteTime
                    })
                }
            }
            catch {
                Write-Verbose ('Ignoring malformed dated filename: {0}' -f $_.Exception.Message)
            }
        }
    }

    return @($Results | Sort-Object Date,Name)
}

function Get-SearchCoverage {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('STATUS','MWI','LIFECYCLE','ALL')]
        [string]$Mode
    )

    $MinDate = [datetime]'1900-01-01'
    $MaxDate = [datetime]'2999-12-31'

    $StatusDates = @(
        Get-DatedLogs -Type STATUS -StartDate $MinDate -EndDate $MaxDate |
        Where-Object { $_.Length -gt 0 } |
        ForEach-Object { $_.Date.Date } |
        Sort-Object -Unique
    )

    $MWILogObjects = @()
    $MWILogObjects += @(Get-DatedLogs -Type RVSIP -StartDate $MinDate -EndDate $MaxDate)
    $MWILogObjects += @(Get-DatedLogs -Type SIP -StartDate $MinDate -EndDate $MaxDate)

    $MWIDates = @(
        $MWILogObjects |
        Where-Object { $_.Length -gt 0 } |
        ForEach-Object { $_.Date.Date } |
        Sort-Object -Unique
    )

    switch ($Mode) {
        'STATUS' {
            $Dates = @($StatusDates)
            $Label = 'STATUS / voicemail-deposit'
        }
        'MWI' {
            $Dates = @($MWIDates)
            $Label = 'RVSIP/SIP MWI'
        }
        'LIFECYCLE' {
            $Dates = @(
                $StatusDates |
                Where-Object { $MWIDates -contains $_ } |
                Sort-Object -Unique
            )
            $Label = 'full lifecycle (STATUS + MWI)'
        }
        default {
            $AllDates = @()
            $AllDates += @($StatusDates)
            $AllDates += @($MWIDates)
            $Dates = @($AllDates | Sort-Object -Unique)
            $Label = 'STATUS/RVSIP/SIP'
        }
    }

    if ($Dates.Count -eq 0) {
        return [pscustomobject]@{
            Mode = $Mode
            Label = $Label
            HasData = $false
            Oldest = $null
            Newest = $null
            DateCount = 0
            CalendarSpanDays = 0
            MaxLookbackDays = 0
            HasToday = $false
            MissingDates = @()
        }
    }

    $Oldest = ($Dates | Sort-Object | Select-Object -First 1).Date
    $Newest = ($Dates | Sort-Object | Select-Object -Last 1).Date
    $Today = (Get-Date).Date
    $CalendarSpanDays = [int](($Newest - $Oldest).TotalDays) + 1
    $MaxLookbackDays = [int](($Today - $Oldest).TotalDays) + 1
    if ($MaxLookbackDays -lt 1) { $MaxLookbackDays = 1 }

    $MissingDates = @()
    $Cursor = $Oldest
    while ($Cursor -le $Newest) {
        if ($Dates -notcontains $Cursor) {
            $MissingDates += $Cursor
        }
        $Cursor = $Cursor.AddDays(1)
    }

    return [pscustomobject]@{
        Mode = $Mode
        Label = $Label
        HasData = $true
        Oldest = $Oldest
        Newest = $Newest
        DateCount = $Dates.Count
        CalendarSpanDays = $CalendarSpanDays
        MaxLookbackDays = $MaxLookbackDays
        HasToday = ($Dates -contains $Today)
        MissingDates = @($MissingDates)
    }
}

function Get-LineTime {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Line,
        [Parameter(Mandatory)][datetime]$FileDate
    )

    if ([string]::IsNullOrWhiteSpace($Line)) { return $null }

    if ($Line -match '(?<!\d)(?<Time>\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?)') {
        $Formats = @('HH:mm:ss.fff','HH:mm:ss.ff','HH:mm:ss.f','HH:mm:ss')
        foreach ($Format in $Formats) {
            $Parsed = [datetime]::MinValue
            if ([datetime]::TryParseExact(
                $Matches.Time,
                $Format,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None,
                [ref]$Parsed
            )) {
                return $FileDate.Date.Add($Parsed.TimeOfDay)
            }
        }
    }

    return $null
}

function Get-ChannelFromLine {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) { return $null }

    if ($Line -match '\[CH:\s*(?<Channel>\d+)\]') {
        return [int]$Matches.Channel
    }

    return $null
}

# -----------------------------------------------------------------------------
# STATUS log voicemail deposit parser
# -----------------------------------------------------------------------------

function New-DepositState {
    return [ordered]@{
        StartTime = $null
        EndTime = $null
        SavedTime = $null
        Mailbox = $null
        MailboxID = $null
        CallerID = $null
        CallerName = $null
        DurationMs = $null
        MessageFile = $null
        MessageAddReturn = $null
        MessageAddSucceeded = $false
        Added = $false
    }
}

function Parse-InMsgXml {
    param([Parameter(Mandatory)][string]$Text)

    $Result = [ordered]@{
        Command = $null
        MailboxID = $null
        Channel = $null
        CallerID = $null
        CallerName = $null
        Timestamp = $null
    }

    if ($Text -match '<CMD>(INMSGSTART|INMSGEND)</CMD>') {
        $Result.Command = $Matches[1]
    }
    if ($Text -match '<MBXID>(\d+)</MBXID>') {
        $Result.MailboxID = $Matches[1]
    }
    if ($Text -match '<CHAN>(\d+)</CHAN>') {
        $Result.Channel = [int]$Matches[1]
    }
    if ($Text -match '<CALLERID>(.*?)</CALLERID>') {
        $Result.CallerID = $Matches[1]
    }
    if ($Text -match '<CALLERIDNAME>(.*?)</CALLERIDNAME>') {
        $Result.CallerName = ($Matches[1] -replace '\s+',' ').Trim()
    }
    if ($Text -match '<TIMESTAMP>(\d{14})</TIMESTAMP>') {
        try {
            $Result.Timestamp = [datetime]::ParseExact(
                $Matches[1],
                'yyyyMMddHHmmss',
                [System.Globalization.CultureInfo]::InvariantCulture
            )
        }
        catch {
            $Result.Timestamp = $null
        }
    }

    return [pscustomobject]$Result
}

function Get-VoicemailDeposits {
    param(
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    $CacheKey = '{0:yyyyMMdd}-{1:yyyyMMdd}' -f $StartDate,$EndDate
    if ($script:DepositCache.ContainsKey($CacheKey)) {
        return @($script:DepositCache[$CacheKey])
    }

    $Logs = @(Get-DatedLogs -Type STATUS -StartDate $StartDate -EndDate $EndDate)
    $Results = New-Object System.Collections.Generic.List[object]

    foreach ($Log in $Logs) {
        if ($Log.Length -le 0) { continue }

        Write-Host ('Parsing {0}...' -f $Log.Name) -ForegroundColor DarkGray

        $States = @{}
        $LastFastAddChannel = $null
        $XmlBuffer = $null

        $Handle = $null
        try {
            $Handle = New-SharedReader -Path $Log.Path
            $Reader = $Handle.Reader

            while ($null -ne ($Line = $Reader.ReadLine())) {
                # -----------------------------------------------------------------
                # INMSGSTART / INMSGEND XML may span display/log lines.
                # -----------------------------------------------------------------
                if ($null -ne $XmlBuffer) {
                    $XmlBuffer += ' ' + $Line.Trim()
                    if ($XmlBuffer -match '</TIMESTAMP>') {
                        $Xml = Parse-InMsgXml -Text $XmlBuffer
                        $XmlBuffer = $null

                        if ($null -ne $Xml.Channel) {
                            $Channel = [int]$Xml.Channel
                            if (-not $States.ContainsKey($Channel) -or $Xml.Command -eq 'INMSGSTART') {
                                $States[$Channel] = New-DepositState
                            }

                            $State = $States[$Channel]
                            if ($Xml.MailboxID) { $State.MailboxID = $Xml.MailboxID }
                            if ($Xml.CallerID) { $State.CallerID = $Xml.CallerID }
                            if ($Xml.CallerName) { $State.CallerName = $Xml.CallerName }

                            if ($Xml.Command -eq 'INMSGSTART') {
                                $State.StartTime = $Xml.Timestamp
                            }
                            elseif ($Xml.Command -eq 'INMSGEND') {
                                $State.EndTime = $Xml.Timestamp
                            }
                        }
                    }
                    continue
                }

                if ($Line -match '<CMD>INMSG(?:START|END)</CMD>') {
                    $XmlBuffer = $Line
                    if ($XmlBuffer -match '</TIMESTAMP>') {
                        $Xml = Parse-InMsgXml -Text $XmlBuffer
                        $XmlBuffer = $null

                        if ($null -ne $Xml.Channel) {
                            $Channel = [int]$Xml.Channel
                            if (-not $States.ContainsKey($Channel) -or $Xml.Command -eq 'INMSGSTART') {
                                $States[$Channel] = New-DepositState
                            }

                            $State = $States[$Channel]
                            if ($Xml.MailboxID) { $State.MailboxID = $Xml.MailboxID }
                            if ($Xml.CallerID) { $State.CallerID = $Xml.CallerID }
                            if ($Xml.CallerName) { $State.CallerName = $Xml.CallerName }

                            if ($Xml.Command -eq 'INMSGSTART') {
                                $State.StartTime = $Xml.Timestamp
                            }
                            elseif ($Xml.Command -eq 'INMSGEND') {
                                $State.EndTime = $Xml.Timestamp
                            }
                        }
                    }
                    continue
                }

                $Channel = Get-ChannelFromLine -Line $Line

                # Mailbox number appears in the recording states.
                if ($null -ne $Channel -and $Line -match 'Recording Menu (?:Mailbox|Mbx)\s+(?<Mailbox>\d+)') {
                    if (-not $States.ContainsKey($Channel)) { $States[$Channel] = New-DepositState }
                    $States[$Channel].Mailbox = $Matches.Mailbox
                }

                # Recording duration reported in milliseconds.
                if ($null -ne $Channel -and $Line -match 'EEAM\.GetPlayTime\s*=\s*(?<Ms>\d+)') {
                    if (-not $States.ContainsKey($Channel)) { $States[$Channel] = New-DepositState }
                    $States[$Channel].DurationMs = [int64]$Matches.Ms
                }

                # MessageAdd start supplies channel and caller information.
                if ($Line -match '\[F:FastMessageAdd\]\s+start,.*?Channel:\s*(?<Channel>\d+),\s*CallerIDNumber:\s*(?<Caller>.*?),\s*CallerIDName:\s*(?<Name>.*)$') {
                    $Channel = [int]$Matches.Channel
                    if (-not $States.ContainsKey($Channel)) { $States[$Channel] = New-DepositState }
                    $States[$Channel].CallerID = $Matches.Caller.Trim()
                    $States[$Channel].CallerName = $Matches.Name.Trim()
                    $LastFastAddChannel = $Channel
                }

                if ($Line -match 'EEAMHelper MessageAdd returned\s*=\s*(?<Code>-?\d+)') {
                    if ($null -ne $LastFastAddChannel -and $States.ContainsKey($LastFastAddChannel)) {
                        $States[$LastFastAddChannel].MessageAddReturn = [int]$Matches.Code
                    }
                }

                # This is our authoritative successful message-store marker.
                if ($Line -match 'XEEAM_MessageAdd succeeded') {
                    if ($null -ne $LastFastAddChannel -and $States.ContainsKey($LastFastAddChannel)) {
                        $State = $States[$LastFastAddChannel]
                        $State.MessageAddSucceeded = $true
                        $State.SavedTime = Get-LineTime -Line $Line -FileDate $Log.Date
                    }
                }

                # Final mailbox number / mailbox ID correlation.
                if ($null -ne $Channel -and $Line -match 'MbxNo\s*=\s*(?<Mailbox>\d+),\s*MbxID\s*=\s*(?<MailboxID>\d+)') {
                    if (-not $States.ContainsKey($Channel)) { $States[$Channel] = New-DepositState }
                    $States[$Channel].Mailbox = $Matches.Mailbox
                    $States[$Channel].MailboxID = $Matches.MailboxID
                }

                # Permanent voicemail GUID / filename.
                if ($null -ne $Channel -and $Line -match 'msgrec\.FileName\s*=\s*(?<File>[^\s]+)') {
                    if (-not $States.ContainsKey($Channel)) { $States[$Channel] = New-DepositState }
                    $State = $States[$Channel]
                    $State.MessageFile = $Matches.File.Trim()

                    if ($State.MessageAddSucceeded -and -not $State.Added -and $State.Mailbox -and $null -ne $State.StartTime) {
                        $Saved = $State.SavedTime
                        if ($null -eq $Saved) { $Saved = $State.EndTime }
                        if ($null -eq $Saved) { $Saved = Get-LineTime -Line $Line -FileDate $Log.Date }

                        $DurationSec = $null
                        if ($null -ne $State.DurationMs) {
                            $DurationSec = [math]::Round(($State.DurationMs / 1000.0), 1)
                        }

                        $Results.Add([pscustomobject]@{
                            EventTime = $Saved
                            Date = if ($Saved) { $Saved.ToString('MM/dd/yyyy') } else { $Log.Date.ToString('MM/dd/yyyy') }
                            Time = if ($Saved) { $Saved.ToString('HH:mm:ss') } else { '' }
                            Mailbox = [string]$State.Mailbox
                            'IXM-ID' = [string]$State.MailboxID
                            CallerID = [string]$State.CallerID
                            CallerName = [string]$State.CallerName
                            DurationSec = $DurationSec
                            MessageFile = [string]$State.MessageFile
                            Result = 'SAVED'
                            Log = $Log.Name
                        })
                        $State.Added = $true
                    }
                }
            }
        }
        catch {
            Write-Host ('Unable to parse {0}: {1}' -f $Log.Name,$_.Exception.Message) -ForegroundColor Red
        }
        finally {
            if ($null -ne $Handle) {
                if ($null -ne $Handle.Reader) {
                    $Handle.Reader.Close()
                    $Handle.Reader.Dispose()
                }
                if ($null -ne $Handle.Stream) {
                    $Handle.Stream.Close()
                    $Handle.Stream.Dispose()
                }
            }
        }
    }

    $Sorted = @($Results | Sort-Object EventTime,Mailbox)
    $script:DepositCache[$CacheKey] = $Sorted
    return $Sorted
}

# -----------------------------------------------------------------------------
# MWI parser
# -----------------------------------------------------------------------------

function Parse-VoiceMessage {
    param([string]$Line)

    $Result = [ordered]@{
        NewVoice = $null
        OldVoice = $null
        UrgentNew = $null
        UrgentOld = $null
    }

    if ($Line -and $Line -match 'Voice-Message:\s*(\d+)\s*/\s*(\d+)(?:\s*\((\d+)\s*/\s*(\d+)\))?') {
        $Result.NewVoice = [int]$Matches[1]
        $Result.OldVoice = [int]$Matches[2]
        if ($Matches[3] -ne '') { $Result.UrgentNew = [int]$Matches[3] }
        if ($Matches[4] -ne '') { $Result.UrgentOld = [int]$Matches[4] }
    }

    return [pscustomobject]$Result
}

function Get-PreferredMWILogs {
    param(
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    $RVSIP = @(Get-DatedLogs -Type RVSIP -StartDate $StartDate -EndDate $EndDate)
    $SIP = @(Get-DatedLogs -Type SIP -StartDate $StartDate -EndDate $EndDate)
    $AllDates = @()
    $AllDates += @($RVSIP | ForEach-Object { $_.Date })
    $AllDates += @($SIP | ForEach-Object { $_.Date })
    $Dates = @($AllDates | Sort-Object -Unique)
    $Selected = New-Object System.Collections.Generic.List[object]

    foreach ($Date in $Dates) {
        $Rv = $RVSIP | Where-Object { $_.Date -eq $Date -and $_.Length -gt 0 } | Select-Object -First 1
        if ($Rv) {
            $Selected.Add($Rv)
            continue
        }

        $Sp = $SIP | Where-Object { $_.Date -eq $Date -and $_.Length -gt 0 } | Select-Object -First 1
        if ($Sp) { $Selected.Add($Sp) }
    }

    return @($Selected | Sort-Object Date)
}

function Get-AllMWIEvents {
    param(
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    $CacheKey = '{0:yyyyMMdd}-{1:yyyyMMdd}' -f $StartDate,$EndDate
    if ($script:AllMWICache.ContainsKey($CacheKey)) {
        return @($script:AllMWICache[$CacheKey])
    }

    $Logs = @(Get-PreferredMWILogs -StartDate $StartDate -EndDate $EndDate)
    $Results = New-Object System.Collections.Generic.List[object]

    foreach ($Log in $Logs) {
        Write-Host ('Searching {0} for MWI activity...' -f $Log.Name) -ForegroundColor DarkGray
        $Handle = $null

        try {
            $Handle = New-SharedReader -Path $Log.Path
            $Reader = $Handle.Reader

            while ($null -ne ($Line = $Reader.ReadLine())) {
                if ($Log.Type -eq 'RVSIP') {
                    if ($Line -match '-->\s+NOTIFY\s+sip:(?<Extension>\d+)@') {
                        $Extension = $Matches.Extension
                        $Block = New-Object System.Collections.Generic.List[string]
                        $Block.Add($Line)

                        for ($i=0; $i -lt 60 -and -not $Reader.EndOfStream; $i++) {
                            $Next = $Reader.ReadLine()
                            if ($null -eq $Next) { break }
                            $Block.Add($Next)
                        }

                        $IsMWI = $false
                        $State = $null
                        $VoiceLine = $null
                        $Acknowledged = $false

                        foreach ($B in $Block) {
                            if ($B -match 'Event:\s*message-summary' -or $B -match 'X-TOL-Call-Reason:\s*MWI') {
                                $IsMWI = $true
                            }
                            if ($B -match 'Messages-Waiting:\s*(yes|no)') {
                                $State = $Matches[1].ToLower()
                            }
                            if ($B -match 'Voice-Message:') {
                                $VoiceLine = $B
                            }
                            if ($B -match '<--\s+SIP/2\.0\s+200\s+OK') {
                                $Acknowledged = $true
                            }
                        }

                        if ($IsMWI -and $State) {
                            $When = Get-LineTime -Line $Line -FileDate $Log.Date
                            $Voice = Parse-VoiceMessage -Line $VoiceLine
                            $Results.Add([pscustomobject]@{
                                EventTime = $When
                                Date = if ($When) { $When.ToString('MM/dd/yyyy') } else { $Log.Date.ToString('MM/dd/yyyy') }
                                Time = if ($When) { $When.ToString('HH:mm:ss.fff') } else { '' }
                                Extension = [string]$Extension
                                MWI = if ($State -eq 'yes') { 'ON' } else { 'OFF' }
                                NewVoice = $Voice.NewVoice
                                OldVoice = $Voice.OldVoice
                                UrgentNew = $Voice.UrgentNew
                                UrgentOld = $Voice.UrgentOld
                                'IXM-ID' = ''
                                Acknowledged = $Acknowledged
                                Source = 'RVSIP'
                                Log = $Log.Name
                            })
                        }
                    }
                }
                elseif ($Log.Type -eq 'SIP') {
                    if ($Line -match '\[IF:SetMWI\].*?/\s*[0-9A-Fa-f]+:(?<Extension>\d+)\s*/') {
                        $Extension = $Matches.Extension
                        $IXMID = ''
                        $Unread = $null
                        $State = $null
                        $VoiceLine = $null
                        $Acknowledged = $false

                        if ($Line -match '<MAILBOXID>(.*?)</MAILBOXID>') { $IXMID = $Matches[1] }
                        if ($Line -match '<VOICENUU>(\d+)</VOICENUU>') { $Unread = [int]$Matches[1] }

                        for ($i=0; $i -lt 40 -and -not $Reader.EndOfStream; $i++) {
                            $Next = $Reader.ReadLine()
                            if ($null -eq $Next) { break }
                            if ($Next -match 'Messages-Waiting:\s*(yes|no)') { $State = $Matches[1].ToLower() }
                            if ($Next -match 'Voice-Message:') { $VoiceLine = $Next }
                            if ($Next -match 'RESPONSE_SUCCESSFUL_RECVD\s+METHOD:\s*NOTIFY') { $Acknowledged = $true }
                        }

                        if (-not $State -and $null -ne $Unread) {
                            $State = if ($Unread -gt 0) { 'yes' } else { 'no' }
                        }

                        if ($State) {
                            $When = Get-LineTime -Line $Line -FileDate $Log.Date
                            $Voice = Parse-VoiceMessage -Line $VoiceLine
                            $Results.Add([pscustomobject]@{
                                EventTime = $When
                                Date = if ($When) { $When.ToString('MM/dd/yyyy') } else { $Log.Date.ToString('MM/dd/yyyy') }
                                Time = if ($When) { $When.ToString('HH:mm:ss.fff') } else { '' }
                                Extension = [string]$Extension
                                MWI = if ($State -eq 'yes') { 'ON' } else { 'OFF' }
                                NewVoice = $Voice.NewVoice
                                OldVoice = $Voice.OldVoice
                                UrgentNew = $Voice.UrgentNew
                                UrgentOld = $Voice.UrgentOld
                                'IXM-ID' = $IXMID
                                Acknowledged = $Acknowledged
                                Source = 'SIP'
                                Log = $Log.Name
                            })
                        }
                    }
                }
            }
        }
        catch {
            Write-Host ('Unable to search {0}: {1}' -f $Log.Name,$_.Exception.Message) -ForegroundColor Red
        }
        finally {
            if ($null -ne $Handle) {
                if ($null -ne $Handle.Reader) {
                    $Handle.Reader.Close()
                    $Handle.Reader.Dispose()
                }
                if ($null -ne $Handle.Stream) {
                    $Handle.Stream.Close()
                    $Handle.Stream.Dispose()
                }
            }
        }
    }

    $Sorted = @($Results | Sort-Object EventTime,Extension)
    $script:AllMWICache[$CacheKey] = $Sorted
    return $Sorted
}

function Get-MWIEvents {
    param(
        [Parameter(Mandatory)][string]$Extension,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    $CacheKey = '{0}|{1:yyyyMMdd}-{2:yyyyMMdd}' -f $Extension,$StartDate,$EndDate
    if ($script:MWICache.ContainsKey($CacheKey)) {
        return @($script:MWICache[$CacheKey])
    }

    $Events = @(
        Get-AllMWIEvents -StartDate $StartDate -EndDate $EndDate |
        Where-Object { $_.Extension -eq $Extension }
    )

    $script:MWICache[$CacheKey] = $Events
    return $Events
}

function Convert-ToClearEvent {
    param([Parameter(Mandatory)]$MWIEvent)

    $ClearType = 'MWI OFF'
    $Meaning = 'IX Messaging sent MWI OFF'

    if ($null -ne $MWIEvent.NewVoice -and $MWIEvent.NewVoice -eq 0) {
        if ($null -ne $MWIEvent.OldVoice) {
            if ($MWIEvent.OldVoice -gt 0) {
                $ClearType = 'UNREAD CLEARED'
                $Meaning = ('MWI OFF; no new voice messages; {0} old/read voice message(s) remain' -f $MWIEvent.OldVoice)
            }
            elseif ($MWIEvent.OldVoice -eq 0) {
                $ClearType = 'EMPTY STATE'
                $Meaning = 'MWI OFF; mailbox reports 0 new / 0 old voice messages'
            }
            else {
                $ClearType = 'UNREAD CLEARED'
                $Meaning = 'MWI OFF; no unread voice messages'
            }
        }
        else {
            $ClearType = 'UNREAD CLEARED'
            $Meaning = 'MWI OFF; no unread voice messages'
        }
    }

    return [pscustomobject]@{
        EventTime = $MWIEvent.EventTime
        Date = $MWIEvent.Date
        Time = $MWIEvent.Time
        Mailbox = $MWIEvent.Extension
        'IXM-ID' = $MWIEvent.'IXM-ID'
        NewVoice = $MWIEvent.NewVoice
        OldVoice = $MWIEvent.OldVoice
        ClearType = $ClearType
        Meaning = $Meaning
        Acknowledged = $MWIEvent.Acknowledged
        Source = $MWIEvent.Source
        Log = $MWIEvent.Log
    }
}

function Get-MailboxClearEvents {
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    $Results = New-Object System.Collections.Generic.List[object]
    foreach ($MwiEvent in @(Get-MWIEvents -Extension $Mailbox -StartDate $StartDate -EndDate $EndDate)) {
        if ($MwiEvent.MWI -eq 'OFF') {
            $Results.Add((Convert-ToClearEvent -MWIEvent $MwiEvent))
        }
    }
    return @($Results | Sort-Object EventTime)
}

function Get-AllClearEvents {
    param(
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    $Results = New-Object System.Collections.Generic.List[object]
    foreach ($MwiEvent in @(Get-AllMWIEvents -StartDate $StartDate -EndDate $EndDate)) {
        if ($MwiEvent.MWI -eq 'OFF') {
            $Results.Add((Convert-ToClearEvent -MWIEvent $MwiEvent))
        }
    }
    return @($Results | Sort-Object EventTime,Mailbox)
}

# -----------------------------------------------------------------------------
# Display / export helpers
# -----------------------------------------------------------------------------

function Protect-IxmCsvCell {
    param($Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -isnot [string]) {
        return $Value
    }

    $Text = [string]$Value
    $Trimmed = $Text.TrimStart()

    # Excel and similar spreadsheet applications can evaluate CSV cells that
    # begin with formula characters. Prefix an apostrophe so customer-derived
    # text is treated as literal data.
    if ($Trimmed -match '^[=+\-@]' -or
        ($Text.Length -gt 0 -and ([int][char]$Text[0] -eq 9 -or [int][char]$Text[0] -eq 13))) {
        return ("'" + $Text)
    }

    return $Text
}

function ConvertTo-IxmCsvSafeRecord {
    param(
        [Parameter(Mandatory)]$InputObject,
        [switch]$RedactSensitive
    )

    $SensitiveProperties = @(
        'FirstName',
        'LastName',
        'Name',
        'CallerID',
        'CallerName',
        'EmailAddress',
        'MessageFile',
        'SyncID',
        'IMAPUIDS',
        'LastSourceFile',
        'SourceLog'
    )

    $Record = [ordered]@{}

    foreach ($Property in $InputObject.PSObject.Properties) {
        if ($Property.Name -eq 'EventTime') {
            continue
        }

        $Value = $Property.Value

        if ($RedactSensitive -and
            $SensitiveProperties -contains $Property.Name -and
            $null -ne $Value -and
            -not [string]::IsNullOrWhiteSpace([string]$Value)) {
            $Value = '[REDACTED]'
        }

        $Record[$Property.Name] = Protect-IxmCsvCell -Value $Value
    }

    return [pscustomobject]$Record
}

function Export-ResultSet {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Data,
        [Parameter(Mandatory)][string]$BaseName
    )

    if (@($Data).Count -eq 0) { return }

    Write-Host ''
    Write-Host 'CSV exports can contain customer data such as names, email addresses, caller IDs, mailbox identifiers, and diagnostic details.' -ForegroundColor Yellow
    Write-Host 'Save reports only to an approved secured location and follow the customer retention policy.' -ForegroundColor Yellow

    $Answer = (Read-Host 'Export these results to CSV? [y/N]').Trim()
    if ($Answer -notmatch '^(?i)y(?:es)?$') { return }

    $RedactAnswer = (Read-Host 'Redact names, email addresses, caller IDs, message filenames, and external sync IDs in the CSV? [y/N]').Trim()
    $RedactSensitive = ($RedactAnswer -match '^(?i)y(?:es)?$')

    $SafeExportData = @(
        $Data | ForEach-Object {
            ConvertTo-IxmCsvSafeRecord -InputObject $_ -RedactSensitive:$RedactSensitive
        }
    )

    $SafeName = $BaseName -replace '[^A-Za-z0-9_.-]','_'
    $Classification = if ($RedactSensitive) { 'REDACTED' } else { 'CONFIDENTIAL' }
    $DefaultFileName = ('{0}_{1}_{2}.csv' -f $Classification,$SafeName,(Get-Date -Format 'yyyyMMdd_HHmmss'))
    $Path = $null
    $Dialog = $null

    try {
        # Use the standard Windows Save As dialog when available.
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

        $Dialog = New-Object System.Windows.Forms.SaveFileDialog
        $Dialog.Title = 'Save IX Messaging CSV Report'
        $Dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
        $Dialog.DefaultExt = 'csv'
        $Dialog.AddExtension = $true
        $Dialog.OverwritePrompt = $true
        $Dialog.RestoreDirectory = $true
        $Dialog.FileName = $DefaultFileName

        if ($script:LastExportDirectory -and (Test-Path -LiteralPath $script:LastExportDirectory)) {
            $Dialog.InitialDirectory = $script:LastExportDirectory
        }
        else {
            $Documents = [Environment]::GetFolderPath('MyDocuments')
            if ($Documents -and (Test-Path -LiteralPath $Documents)) {
                $Dialog.InitialDirectory = $Documents
            }
        }

        $DialogResult = $Dialog.ShowDialog()
        if ($DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Host 'CSV export canceled.' -ForegroundColor DarkGray
            return
        }

        $Path = $Dialog.FileName
        $script:LastExportDirectory = Split-Path -Parent $Path
    }
    catch {
        # If the GUI dialog cannot be opened, use a user-scoped fallback
        # directory rather than a shared C:\Temp location.
        Write-Host ('Save As dialog unavailable ({0}). Using protected user-scoped export folder.' -f $_.Exception.Message) -ForegroundColor Yellow

        if (-not (Test-Path -LiteralPath $ExportRoot)) {
            New-Item -ItemType Directory -Path $ExportRoot -Force | Out-Null
        }

        $Path = Join-Path $ExportRoot $DefaultFileName
    }
    finally {
        if ($null -ne $Dialog) {
            $Dialog.Dispose()
        }
    }

    try {
        $SafeExportData | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
        Write-Host ('Saved: {0}' -f $Path) -ForegroundColor Green
        if (-not $RedactSensitive) {
            Write-Host 'Classification: CONFIDENTIAL - CUSTOMER DATA. Protect and delete according to customer policy.' -ForegroundColor Yellow
        }
        else {
            Write-Host 'Export mode: REDACTED. Formula-injection protection was also applied.' -ForegroundColor Green
        }
    }
    catch {
        Write-Host ('CSV export failed: {0}' -f $_.Exception.Message) -ForegroundColor Red
    }
}

function Show-Deposits {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Data)

    if (@($Data).Count -eq 0) {
        Write-Host 'No successfully stored voicemail deposits were found in the available STATUS logs.' -ForegroundColor Yellow
        return
    }

    $Data |
        Select-Object Date,Time,Mailbox,'IXM-ID',CallerID,CallerName,DurationSec,MessageFile,Result,Log |
        Format-Table -AutoSize | Out-Host

    Write-Host ''
    Write-Host ('Successful voicemail deposits: {0}' -f @($Data).Count) -ForegroundColor Green
}

function Show-MWIEvents {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Data)

    if (@($Data).Count -eq 0) {
        Write-Host 'No MWI ON/OFF events were found in the available SIP/RVSIP logs.' -ForegroundColor Yellow
        return
    }

    $Data |
        Select-Object Date,Time,Extension,MWI,NewVoice,OldVoice,'IXM-ID',Acknowledged,Source,Log |
        Format-Table -AutoSize | Out-Host
}

function Show-ClearEvents {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Data)

    if (@($Data).Count -eq 0) {
        Write-Host 'No MWI OFF / voicemail-clear events were found in the available logs.' -ForegroundColor Yellow
        return
    }

    $Data |
        Select-Object Date,Time,Mailbox,'IXM-ID',NewVoice,OldVoice,ClearType,Meaning,Acknowledged,Source,Log |
        Format-Table -Wrap -AutoSize | Out-Host

    Write-Host ''
    Write-Host ('Clear / MWI OFF events: {0}' -f @($Data).Count) -ForegroundColor Green
    Write-Host 'Note: UNREAD CLEARED means IXM reported 0 new messages and sent MWI OFF. EMPTY STATE means IXM reported 0 new / 0 old; neither alone proves a specific delete action.' -ForegroundColor DarkGray
}

function Show-MailboxTimeline {
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Deposits,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$MWI
    )

    $Timeline = New-Object System.Collections.Generic.List[object]

    foreach ($D in $Deposits) {
        $Caller = $D.CallerID
        if ($D.CallerName) { $Caller = ('{0} ({1})' -f $D.CallerID,$D.CallerName) }
        $Detail = 'From {0}; {1}s; IXM-ID {2}; file {3}' -f $Caller,$D.DurationSec,$D.'IXM-ID',$D.MessageFile

        $Timeline.Add([pscustomobject]@{
            EventTime = $D.EventTime
            Date = $D.Date
            Time = $D.Time
            Event = 'VOICEMAIL SAVED'
            Detail = $Detail
            Source = $D.Log
        })
    }

    foreach ($M in $MWI) {
        $EventName = ('MWI ' + $M.MWI)
        if ($M.MWI -eq 'OFF' -and $null -ne $M.NewVoice -and $M.NewVoice -eq 0) {
            if ($null -ne $M.OldVoice -and $M.OldVoice -gt 0) {
                $EventName = 'UNREAD CLEARED / MWI OFF'
            }
            elseif ($null -ne $M.OldVoice -and $M.OldVoice -eq 0) {
                $EventName = 'EMPTY STATE / MWI OFF'
            }
            else {
                $EventName = 'MWI OFF'
            }
        }

        $AckText = if ($M.Acknowledged) { 'yes' } else { 'not confirmed' }
        $Detail = 'MWI {0}; New={1}; Old={2}; acknowledged={3}; source={4}' -f $M.MWI,$M.NewVoice,$M.OldVoice,$AckText,$M.Source
        $Timeline.Add([pscustomobject]@{
            EventTime = $M.EventTime
            Date = $M.Date
            Time = $M.Time
            Event = $EventName
            Detail = $Detail
            Source = $M.Log
        })
    }

    $Sorted = @($Timeline | Sort-Object EventTime)

    if ($Sorted.Count -eq 0) {
        Write-Host ('No mailbox activity was found for {0} in the available logs.' -f $Mailbox) -ForegroundColor Yellow
        return @()
    }

    $Sorted | Select-Object Date,Time,Event,Detail,Source | Format-Table -Wrap -AutoSize | Out-Host

    $LastDeposit = @($Deposits | Sort-Object EventTime | Select-Object -Last 1)
    if ($LastDeposit.Count -gt 0) {
        $LastD = $LastDeposit[0]
        $LaterOff = @($MWI | Where-Object { $_.MWI -eq 'OFF' -and $_.EventTime -gt $LastD.EventTime } | Sort-Object EventTime | Select-Object -First 1)
        Write-Host ''
        if ($LaterOff.Count -gt 0) {
            $C = Convert-ToClearEvent -MWIEvent $LaterOff[0]
            Write-Host ('Latest deposited voicemail was followed by MWI OFF at {0} {1}.' -f $C.Date,$C.Time) -ForegroundColor Green
            Write-Host ('  {0}; acknowledged={1}' -f $C.Meaning,$C.Acknowledged) -ForegroundColor Green
        }
        else {
            Write-Host 'No later MWI OFF event was found after the latest voicemail deposit in the selected logs.' -ForegroundColor Yellow
        }
    }

    return $Sorted
}

function Show-LogCoverage {
    param(
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    $Rows = New-Object System.Collections.Generic.List[object]
    foreach ($Type in @('STATUS','RVSIP','SIP')) {
        foreach ($Log in @(Get-DatedLogs -Type $Type -StartDate $StartDate -EndDate $EndDate)) {
            $Rows.Add([pscustomobject]@{
                Date = $Log.Date.ToString('MM/dd/yyyy')
                Type = $Type
                File = $Log.Name
                SizeMB = [math]::Round($Log.Length / 1MB, 2)
                LastWrite = $Log.LastWriteTime
            })
        }
    }

    if ($Rows.Count -eq 0) {
        Write-Host 'No STATUS, SIP, or RVSIP logs were found for this date range.' -ForegroundColor Yellow
        return
    }

    $Rows | Sort-Object Date,Type | Format-Table -AutoSize | Out-Host

    Write-Host ''
    Write-Host 'Important:' -ForegroundColor Yellow
    Write-Host '  A day with no STATUS log cannot be treated as proof that no voicemail was left.' -ForegroundColor Yellow
    Write-Host '  Results only describe activity visible in the retained log files above.' -ForegroundColor Yellow
}


# -----------------------------------------------------------------------------
# DBCOM Graph / email synchronization parser
# -----------------------------------------------------------------------------

function Get-DBComRoot {
    $Parent = Split-Path -Path $LogRoot -Parent
    return (Join-Path $Parent 'DBCOM')
}

function Get-DBComDatedLogs {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('EEAMHELPER','TSECMGR')]
        [string]$Type,

        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    $DBComRoot = Get-DBComRoot
    if (-not (Test-Path -LiteralPath $DBComRoot)) {
        return @()
    }

    switch ($Type) {
        'EEAMHELPER' { $Prefix = 'EEAM_EEAMHELPER' }
        'TSECMGR'    { $Prefix = 'EEAM_TSECMGR' }
    }

    $Pattern = '^{0}#(?<Date>\d{{8}})\.log$' -f [regex]::Escape($Prefix)
    $Results = New-Object System.Collections.Generic.List[object]

    foreach ($File in Get-ChildItem -LiteralPath $DBComRoot -File -ErrorAction SilentlyContinue) {
        if ($File.Name -match $Pattern) {
            try {
                $Date = [datetime]::ParseExact(
                    $Matches.Date,
                    'yyyyMMdd',
                    [System.Globalization.CultureInfo]::InvariantCulture
                ).Date

                if ($Date -ge $StartDate.Date -and $Date -le $EndDate.Date) {
                    $Results.Add([pscustomobject]@{
                        Type = $Type
                        Date = $Date
                        Name = $File.Name
                        Path = $File.FullName
                        Length = $File.Length
                        LastWriteTime = $File.LastWriteTime
                    })
                }
            }
            catch {
                Write-Verbose ('Ignoring malformed dated filename: {0}' -f $_.Exception.Message)
            }
        }
    }

    return @($Results | Sort-Object Date,Name)
}

function Resolve-ExtensionMailboxIds {
    param(
        [Parameter(Mandatory)][string]$Extension,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    $Ids = New-Object System.Collections.Generic.List[string]

    # Strongest correlation: successful deposits already parsed by this tool.
    $Deposits = @(
        Get-VoicemailDeposits -StartDate $StartDate -EndDate $EndDate |
        Where-Object { $_.Mailbox -eq $Extension -and $_.'IXM-ID' }
    )

    foreach ($Deposit in $Deposits) {
        $Id = [string]$Deposit.'IXM-ID'
        if ($Id -and -not $Ids.Contains($Id)) {
            $Ids.Add($Id)
        }
    }

    # Also inspect STATUS logs for direct extension -> internal mailbox-ID
    # mappings so the lookup can work without a successful deposit in the range.
    $StatusLogs = @(Get-DatedLogs -Type STATUS -StartDate $StartDate -EndDate $EndDate)

    foreach ($Log in $StatusLogs) {
        if ($Log.Length -le 0) { continue }

        $Handle = $null
        $PendingExtensionLookup = $false
        $PendingLines = 0

        try {
            $Handle = New-SharedReader -Path $Log.Path
            $Reader = $Handle.Reader

            while ($null -ne ($Line = $Reader.ReadLine())) {
                if ($Line -match ('MbxNo\s*=\s*{0},\s*MbxID\s*=\s*(?<Id>\d+)' -f [regex]::Escape($Extension))) {
                    $Id = $Matches.Id
                    if ($Id -and -not $Ids.Contains($Id)) {
                        $Ids.Add($Id)
                    }
                }

                if ($Line -match ('GetmailboxIDFromExtension\].*Begin,\s*Extension:\s*{0}(?:\D|$)' -f [regex]::Escape($Extension))) {
                    $PendingExtensionLookup = $true
                    $PendingLines = 0
                    continue
                }

                if ($PendingExtensionLookup) {
                    $PendingLines++

                    if ($Line -match 'GetmailboxIDFromExtension\].*End,\s*MboxId:\s*(?<Id>\d+)') {
                        $Id = $Matches.Id
                        if ($Id -and $Id -ne '0' -and -not $Ids.Contains($Id)) {
                            $Ids.Add($Id)
                        }
                        $PendingExtensionLookup = $false
                        continue
                    }

                    if ($PendingLines -gt 12) {
                        $PendingExtensionLookup = $false
                    }
                }
            }
        }
        finally {
            if ($null -ne $Handle) {
                if ($null -ne $Handle.Reader) {
                    $Handle.Reader.Close()
                    $Handle.Reader.Dispose()
                }
                if ($null -ne $Handle.Stream) {
                    $Handle.Stream.Close()
                    $Handle.Stream.Dispose()
                }
            }
        }
    }

    return @($Ids | Sort-Object -Unique)
}

function Get-GraphMessageAdds {
    param(
        [Parameter(Mandatory)][string[]]$MailboxIds,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    $Results = New-Object System.Collections.Generic.List[object]
    $Logs = @(Get-DBComDatedLogs -Type EEAMHELPER -StartDate $StartDate -EndDate $EndDate)

    foreach ($Log in $Logs) {
        if ($Log.Length -le 0) { continue }

        Write-Host ('Parsing {0} for IXM message IDs...' -f $Log.Name) -ForegroundColor DarkGray
        $Handle = $null

        try {
            $Handle = New-SharedReader -Path $Log.Path
            $Reader = $Handle.Reader

            while ($null -ne ($Line = $Reader.ReadLine())) {
                if ($Line -notmatch '\[F:\s*MessageAddInternal\]') { continue }

                if ($Line -match '\[M:\s*(?<Mbx>\d+)\]\s*\[FLD:\s*(?<Fld>\d+)\]\s*\[MSG:\s*(?<Msg>\d+)\]\s*Message has been added,\s*retval:\s*(?<Ret>-?\d+)') {
                    $Mbx = $Matches.Mbx
                    if ($MailboxIds -notcontains $Mbx) { continue }

                    $EventTime = Get-LineTime -Line $Line -FileDate $Log.Date

                    $Results.Add([pscustomobject]@{
                        EventTime = $EventTime
                        MailboxID = $Mbx
                        FolderID = $Matches.Fld
                        MessageID = $Matches.Msg
                        AddReturn = [int]$Matches.Ret
                        AddLog = $Log.Name
                    })
                }
            }
        }
        finally {
            if ($null -ne $Handle) {
                if ($null -ne $Handle.Reader) {
                    $Handle.Reader.Close()
                    $Handle.Reader.Dispose()
                }
                if ($null -ne $Handle.Stream) {
                    $Handle.Stream.Close()
                    $Handle.Stream.Dispose()
                }
            }
        }
    }

    return @($Results | Sort-Object EventTime,MessageID)
}

function Get-GraphSyncUpdates {
    param(
        [Parameter(Mandatory)][string[]]$MessageIds,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    $Updates = @{}
    $Failures = @{}
    $Logs = @(Get-DBComDatedLogs -Type TSECMGR -StartDate $StartDate -EndDate $EndDate)

    foreach ($Log in $Logs) {
        if ($Log.Length -le 0) { continue }

        Write-Host ('Parsing {0} for Graph/email sync results...' -f $Log.Name) -ForegroundColor DarkGray
        $Handle = $null

        try {
            $Handle = New-SharedReader -Path $Log.Path
            $Reader = $Handle.Reader
            $Pending = $null

            while ($null -ne ($Line = $Reader.ReadLine())) {
                if ($null -ne $Pending) {
                    if ($Line -notmatch '^\s*\d{2}:\d{2}:\d{2}(?:\.\d+)?\s' -and
                        $Line -notmatch '^\s*\[TID:' -and
                        $Line -match '^\s*AAMk') {

                        $Continuation = $Line.Trim()
                        if ($Continuation -match '^(?<Sync>AAMk.*),\s*IMAPUIDS:\s*(?<Uid>.*)$') {
                            $Pending.SyncID = $Matches.Sync.Trim()
                            $Pending.IMAPUIDS = $Matches.Uid.Trim()
                        }
                        else {
                            $Pending.SyncID = $Continuation.TrimEnd(',')
                        }

                        $Updates[$Pending.MessageID] = [pscustomobject]$Pending
                        $Pending = $null
                        continue
                    }
                    else {
                        $Updates[$Pending.MessageID] = [pscustomobject]$Pending
                        $Pending = $null
                    }
                }

                if ($Line -match '\[F:\s*InternalUpdateSyncStatusOfMessage\]\s*\[MSG:\s*(?<Msg>\d+)\]\s*SyncStatus:\s*(?<Status>-?\d+),\s*IMAPID:\s*(?<Imap>-?\d+),\s*SyncID:\s*(?<Rest>.*)$') {
                    $Msg = $Matches.Msg
                    if ($MessageIds -notcontains $Msg) { continue }

                    $Status = [int]$Matches.Status
                    $Imap = $Matches.Imap
                    $Rest = $Matches.Rest
                    $SyncID = ''
                    $IMAPUIDS = ''

                    if ($Rest -match '^(?<Sync>.*),\s*IMAPUIDS:\s*(?<Uid>.*)$') {
                        $SyncID = $Matches.Sync.Trim()
                        $IMAPUIDS = $Matches.Uid.Trim()
                    }
                    else {
                        $SyncID = $Rest.Trim()
                    }

                    $Record = [ordered]@{
                        EventTime = Get-LineTime -Line $Line -FileDate $Log.Date
                        MessageID = $Msg
                        SyncStatus = $Status
                        IMAPID = $Imap
                        SyncID = $SyncID
                        IMAPUIDS = $IMAPUIDS
                        SyncLog = $Log.Name
                    }

                    if ([string]::IsNullOrWhiteSpace($SyncID)) {
                        $Pending = $Record
                    }
                    else {
                        $Updates[$Msg] = [pscustomobject]$Record
                    }

                    continue
                }

                if ($Line -match '\[MSG:\s*(?<Msg>\d+)\]') {
                    $FailureMsg = $Matches.Msg

                    if ($Line -match '(?i)sync' -and
                        $Line -match '(?i)(fail|error|exception|timeout)' -and
                        $MessageIds -contains $FailureMsg) {

                        $Failures[$FailureMsg] = [pscustomobject]@{
                            EventTime = Get-LineTime -Line $Line -FileDate $Log.Date
                            MessageID = $FailureMsg
                            Text = $Line.Trim()
                            Log = $Log.Name
                        }
                    }
                }
            }

            if ($null -ne $Pending) {
                $Updates[$Pending.MessageID] = [pscustomobject]$Pending
            }
        }
        finally {
            if ($null -ne $Handle) {
                if ($null -ne $Handle.Reader) {
                    $Handle.Reader.Close()
                    $Handle.Reader.Dispose()
                }
                if ($null -ne $Handle.Stream) {
                    $Handle.Stream.Close()
                    $Handle.Stream.Dispose()
                }
            }
        }
    }

    return [pscustomobject]@{
        Updates = $Updates
        Failures = $Failures
    }
}

function Get-ExtensionGraphSyncHistory {
    param(
        [Parameter(Mandatory)][string]$Extension,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )

    $CacheKey = '{0}-{1:yyyyMMdd}-{2:yyyyMMdd}' -f $Extension,$StartDate,$EndDate
    if ($script:GraphSyncCache.ContainsKey($CacheKey)) {
        return @($script:GraphSyncCache[$CacheKey])
    }

    $MailboxIds = @(Resolve-ExtensionMailboxIds -Extension $Extension -StartDate $StartDate -EndDate $EndDate)
    if ($MailboxIds.Count -eq 0) {
        return @()
    }

    Write-Host ('Resolved extension {0} to IXM mailbox ID(s): {1}' -f $Extension,($MailboxIds -join ', ')) -ForegroundColor DarkGray

    $Messages = @(Get-GraphMessageAdds -MailboxIds $MailboxIds -StartDate $StartDate -EndDate $EndDate)
    if ($Messages.Count -eq 0) {
        $script:GraphSyncCache[$CacheKey] = @()
        return @()
    }

    $MessageIds = @($Messages | ForEach-Object { [string]$_.MessageID } | Sort-Object -Unique)
    $Sync = Get-GraphSyncUpdates -MessageIds $MessageIds -StartDate $StartDate -EndDate $EndDate

    $Rows = foreach ($Message in $Messages) {
        $MsgId = [string]$Message.MessageID
        $Update = $null
        $Failure = $null

        if ($Sync.Updates.ContainsKey($MsgId)) {
            $Update = $Sync.Updates[$MsgId]
        }
        if ($Sync.Failures.ContainsKey($MsgId)) {
            $Failure = $Sync.Failures[$MsgId]
        }

        $State = 'NOT CONFIRMED'
        if ($null -ne $Failure) {
            $State = 'FAILED'
        }
        elseif ($null -ne $Update -and -not [string]::IsNullOrWhiteSpace([string]$Update.SyncID)) {
            $State = 'CONFIRMED'
        }

        $DelaySec = $null
        if ($null -ne $Update -and $null -ne $Message.EventTime -and $null -ne $Update.EventTime) {
            $DelaySec = [math]::Round(($Update.EventTime - $Message.EventTime).TotalSeconds,1)
        }

        [pscustomobject]@{
            Extension = $Extension
            'IXM-ID' = $Message.MailboxID
            MessageID = $MsgId
            MessageTime = if ($Message.EventTime) { $Message.EventTime.ToString('MM/dd/yyyy HH:mm:ss.fff') } else { '' }
            Result = $State
            SyncTime = if ($Update -and $Update.EventTime) { $Update.EventTime.ToString('MM/dd/yyyy HH:mm:ss.fff') } else { '' }
            SyncStatus = if ($Update) { $Update.SyncStatus } else { $null }
            SyncDelaySec = $DelaySec
            SyncID = if ($Update) { [string]$Update.SyncID } else { '' }
            IMAPUIDS = if ($Update) { [string]$Update.IMAPUIDS } else { '' }
            Failure = if ($Failure) { [string]$Failure.Text } else { '' }
            HelperLog = $Message.AddLog
            SyncLog = if ($Update) { $Update.SyncLog } elseif ($Failure) { $Failure.Log } else { '' }
        }
    }

    $Rows = @($Rows | Sort-Object MessageTime)
    $script:GraphSyncCache[$CacheKey] = $Rows
    return $Rows
}

function Invoke-ExtensionGraphSyncHistory {
    $Extension = Read-NumericValue -Prompt 'Extension / mailbox'
    $Range = Read-DateRange -CoverageMode STATUS

    Write-Section ("Graph / Email Sync History - Extension {0} - {1}" -f $Extension,$Range.Description)

    $DBComRoot = Get-DBComRoot
    if (-not (Test-Path -LiteralPath $DBComRoot)) {
        Write-Host ('DBCOM log directory not found: {0}' -f $DBComRoot) -ForegroundColor Red
        return
    }

    $Rows = @(Get-ExtensionGraphSyncHistory -Extension $Extension -StartDate $Range.Start -EndDate $Range.End)

    if ($Rows.Count -eq 0) {
        Write-Host 'No correlatable voicemail/Graph sync records were found for this extension in the selected retained logs.' -ForegroundColor Yellow
        Write-Host 'This is not proof that the extension never synchronized; the required STATUS/DBCOM logs may not be retained.' -ForegroundColor Yellow
        return
    }

    $Display = @(
        $Rows |
        Select-Object MessageTime,MessageID,Result,SyncTime,SyncStatus,SyncDelaySec,
            @{Name='SyncID';Expression={
                $Value = [string]$_.SyncID
                if ($Value.Length -gt 34) { $Value.Substring(0,34) + '...' } else { $Value }
            }}
    )

    $Display | Format-Table -AutoSize | Out-Host

    $Confirmed = @($Rows | Where-Object { $_.Result -eq 'CONFIRMED' })
    $Failed = @($Rows | Where-Object { $_.Result -eq 'FAILED' })
    $Unconfirmed = @($Rows | Where-Object { $_.Result -eq 'NOT CONFIRMED' })

    Write-Host ''
    Write-Host ('Confirmed syncs  : {0}' -f $Confirmed.Count)
    Write-Host ('Explicit failures: {0}' -f $Failed.Count)
    Write-Host ('Not confirmed    : {0}' -f $Unconfirmed.Count)

    if ($Confirmed.Count -gt 0) {
        $LastGood = $Confirmed | Sort-Object SyncTime | Select-Object -Last 1
        Write-Host ''
        Write-Host ('Last confirmed sync: {0}' -f $LastGood.SyncTime) -ForegroundColor Green
        Write-Host ('  Message ID : {0}' -f $LastGood.MessageID)
        Write-Host ('  SyncStatus : {0}' -f $LastGood.SyncStatus)
        Write-Host ('  Delay      : {0} sec' -f $LastGood.SyncDelaySec)
        Write-Host ('  SyncID     : {0}' -f $LastGood.SyncID)
        if ($LastGood.IMAPUIDS) {
            Write-Host ('  IMAPUIDS   : {0}' -f $LastGood.IMAPUIDS)
        }
    }

    if ($Failed.Count -gt 0) {
        $LastFail = $Failed | Sort-Object MessageTime | Select-Object -Last 1
        Write-Host ''
        Write-Host ('Last explicit sync failure: {0}' -f $LastFail.MessageTime) -ForegroundColor Red
        Write-Host ('  Message ID : {0}' -f $LastFail.MessageID)
        Write-Host ('  {0}' -f $LastFail.Failure)
    }

    if ($Unconfirmed.Count -gt 0) {
        $LastUnconfirmed = $Unconfirmed | Sort-Object MessageTime | Select-Object -Last 1
        Write-Host ''
        Write-Host ('Most recent message without a confirmed SyncID: {0}' -f $LastUnconfirmed.MessageTime) -ForegroundColor Yellow
        Write-Host ('  Message ID : {0}' -f $LastUnconfirmed.MessageID)
        Write-Host '  This is reported as NOT CONFIRMED, not FAILED, unless an explicit message-linked failure is present.' -ForegroundColor Yellow
    }

    Export-ResultSet -Data $Rows -BaseName ("extension_{0}_graph_sync" -f $Extension)
}



# -----------------------------------------------------------------------------
# IX Messaging database mailbox / email export
# -----------------------------------------------------------------------------

function Get-SystemSqlAnywhereDsns {
    $Found = New-Object System.Collections.Generic.List[object]

    foreach ($Platform in @('64-bit','32-bit')) {
        try {
            $Dsns = @(Get-OdbcDsn -DsnType System -Platform $Platform -ErrorAction Stop)
        }
        catch {
            continue
        }

        foreach ($Dsn in $Dsns) {
            if ($Dsn.DriverName -notmatch 'SQL Anywhere') {
                continue
            }

            $Attrs = @{}
            try {
                foreach ($Attr in $Dsn.Attribute.GetEnumerator()) {
                    $Attrs[[string]$Attr.Key] = [string]$Attr.Value
                }
            }
            catch {
                Write-Verbose ('Optional DSN display attributes are unavailable: {0}' -f $_.Exception.Message)
            }

            $Found.Add([pscustomobject]@{
                Name         = [string]$Dsn.Name
                DriverName   = [string]$Dsn.DriverName
                Platform     = $Platform
                ServerName   = if ($Attrs.ContainsKey('ServerName')) { $Attrs['ServerName'] } else { '' }
                DatabaseName = if ($Attrs.ContainsKey('DatabaseName')) { $Attrs['DatabaseName'] } else { '' }
            })
        }
    }

    # Prefer the first occurrence (64-bit is enumerated first) when identical
    # logical DSNs exist in both ODBC architectures.
    return @(
        $Found |
        Group-Object Name,DatabaseName |
        ForEach-Object { $_.Group | Select-Object -First 1 }
    )
}

function Open-IxmDsnConnection {
    param([Parameter(Mandatory)][string]$Name)

    $Connection = New-Object System.Data.Odbc.OdbcConnection
    $Connection.ConnectionString = ('DSN={0}' -f $Name)
    $Connection.Open()
    return $Connection
}

function Test-IxmMailboxSchema {
    param([Parameter(Mandatory)]$Connection)

    try {
        $Tables = $Connection.GetSchema('Tables')

        $HasMailbox = @(
            $Tables | Where-Object {
                $_.TABLE_NAME -eq 'MAILBOX' -and
                ($_.TABLE_SCHEM -eq 'DBA' -or [string]::IsNullOrWhiteSpace([string]$_.TABLE_SCHEM))
            }
        ).Count -gt 0

        $HasFGroup = @(
            $Tables | Where-Object {
                $_.TABLE_NAME -eq 'FGROUP' -and
                ($_.TABLE_SCHEM -eq 'DBA' -or [string]::IsNullOrWhiteSpace([string]$_.TABLE_SCHEM))
            }
        ).Count -gt 0

        if (-not ($HasMailbox -and $HasFGroup)) {
            return $false
        }

        $Columns = $Connection.GetSchema('Columns')

        $MailboxColumns = @(
            $Columns |
            Where-Object { $_.TABLE_NAME -eq 'MAILBOX' } |
            ForEach-Object { [string]$_.COLUMN_NAME }
        )

        $FGroupColumns = @(
            $Columns |
            Where-Object { $_.TABLE_NAME -eq 'FGROUP' } |
            ForEach-Object { [string]$_.COLUMN_NAME }
        )

        foreach ($Required in @('FGROUPID','MBXNUMBER','FIRSTNAME','LASTNAME')) {
            if ($MailboxColumns -notcontains $Required) {
                return $false
            }
        }

        foreach ($Required in @('FGROUPID','FGNAME')) {
            if ($FGroupColumns -notcontains $Required) {
                return $false
            }
        }

        return $true
    }
    catch {
        return $false
    }
}

function Select-IxmDatabaseDsn {
    $AllCandidates = @(Get-SystemSqlAnywhereDsns)

    if ($AllCandidates.Count -eq 0) {
        throw 'No SQL Anywhere System DSNs were found on this server.'
    }

    # Do not connect to every SQL Anywhere DSN on the host. Restrict automatic
    # probing to conventional IX Messaging DSN/database names.
    $Candidates = @(
        $AllCandidates | Where-Object {
            $_.Name -match '(?i)^UC.*SQLANY$' -or
            $_.DatabaseName -match '(?i)^EEAM\d*$'
        }
    )

    if ($Candidates.Count -eq 0) {
        Write-Host ''
        Write-Host 'No conventional IX Messaging SQL Anywhere DSN name/database was detected.' -ForegroundColor Yellow
        Write-Host 'Nonstandard SQL Anywhere DSNs will NOT be connected to automatically.' -ForegroundColor Yellow
        Write-Host ''
        foreach ($Candidate in $AllCandidates) {
            Write-Host ('  {0}  DB={1}  Server={2}' -f $Candidate.Name,$Candidate.DatabaseName,$Candidate.ServerName) -ForegroundColor DarkGray
        }

        $Approval = (Read-Host 'Probe these nonstandard SQL Anywhere DSNs for the IX Messaging schema? [y/N]').Trim()
        if ($Approval -notmatch '^(?i)y(?:es)?$') {
            throw [System.OperationCanceledException]::new('Nonstandard SQL Anywhere DSN probing was not approved.')
        }

        $Candidates = $AllCandidates
    }

    $Valid = New-Object System.Collections.Generic.List[object]

    foreach ($Candidate in $Candidates) {
        $DbText = if ($Candidate.DatabaseName) { $Candidate.DatabaseName } else { '?' }
        $ServerText = if ($Candidate.ServerName) { $Candidate.ServerName } else { '?' }

        Write-Host (
            'Testing DSN {0} ({1}, DB={2}, Server={3})...' -f
            $Candidate.Name,$Candidate.Platform,$DbText,$ServerText
        ) -ForegroundColor DarkGray

        $Connection = $null
        try {
            $Connection = Open-IxmDsnConnection -Name $Candidate.Name

            if (Test-IxmMailboxSchema -Connection $Connection) {
                $Valid.Add($Candidate)
                Write-Host '  IX Messaging mailbox schema found.' -ForegroundColor Green
            }
            else {
                Write-Host '  Connected, but the expected MAILBOX/FGROUP schema was not found.' -ForegroundColor DarkGray
            }
        }
        catch {
            Write-Host ('  Connection failed: {0}' -f $_.Exception.Message) -ForegroundColor DarkGray
        }
        finally {
            if ($null -ne $Connection) {
                try { $Connection.Close() } catch { Write-Verbose ('ODBC connection Close() cleanup failed: {0}' -f $_.Exception.Message) }
                try { $Connection.Dispose() } catch { Write-Verbose ('ODBC connection Dispose() cleanup failed: {0}' -f $_.Exception.Message) }
            }
        }
    }

    if ($Valid.Count -eq 0) {
        throw 'No SQL Anywhere DSN with the expected IX Messaging MAILBOX/FGROUP schema was found.'
    }

    if ($Valid.Count -eq 1) {
        return $Valid[0]
    }

    $Preferred = @(
        $Valid | Where-Object {
            $_.Name -match '^UC\d*_SQLANY$|^UC.*SQLANY$' -or
            $_.DatabaseName -match '^EEAM'
        }
    )

    if ($Preferred.Count -eq 1) {
        return $Preferred[0]
    }

    Write-Host ''
    Write-Host 'Multiple IX Messaging database DSNs were found:' -ForegroundColor Yellow

    for ($i = 0; $i -lt $Valid.Count; $i++) {
        $Item = $Valid[$i]
        Write-Host (
            '  {0}. {1}  DB={2}  Server={3}  Platform={4}' -f
            ($i + 1),$Item.Name,$Item.DatabaseName,$Item.ServerName,$Item.Platform
        )
    }

    $Selection = 0
    do {
        $Text = (Read-Host ('Select 1-{0}' -f $Valid.Count)).Trim()
        $Parsed = [int]::TryParse($Text,[ref]$Selection)
    } until ($Parsed -and $Selection -ge 1 -and $Selection -le $Valid.Count)

    return $Valid[$Selection - 1]
}

function Test-IxmColumn {
    param(
        [Parameter(Mandatory)]$Columns,
        [Parameter(Mandatory)][string]$Table,
        [Parameter(Mandatory)][string]$Column
    )

    return @(
        $Columns | Where-Object {
            $_.TABLE_NAME -eq $Table -and $_.COLUMN_NAME -eq $Column
        }
    ).Count -gt 0
}

function Test-IxmSelectOnlySql {
    param([Parameter(Mandatory)][string]$Sql)

    $Statement = $Sql.Trim()

    if ([string]::IsNullOrWhiteSpace($Statement)) {
        throw 'SQL statement is empty.'
    }

    # This utility does not accept arbitrary SQL. All internal database access
    # must be one plain SELECT statement.
    if ($Statement -notmatch '(?is)^SELECT(?:\s|$)') {
        throw 'Only a single SELECT statement is permitted.'
    }

    # Disallow statement separators and SQL comments. This blocks appending a
    # second statement or hiding a disallowed operation after a comment.
    if ($Statement -match ';' -or $Statement -match '(?s)/\*|\*/|--') {
        throw 'SQL statement separators and comments are not permitted.'
    }

    $Disallowed = '(?i)\b(INSERT|UPDATE|DELETE|ALTER|DROP|CREATE|TRUNCATE|MERGE|GRANT|REVOKE|CALL|EXEC|EXECUTE|SET|BEGIN|DECLARE|DO|LOAD|UNLOAD|OUTPUT|INTO|LOCK|COMMIT|ROLLBACK|SAVEPOINT|TRIGGER|PROCEDURE)\b'
    if ($Statement -match $Disallowed) {
        throw ('Disallowed SQL token detected: {0}' -f $Matches[1])
    }

    if ($Statement -match '(?i)\bFOR\s+UPDATE\b') {
        throw 'SELECT FOR UPDATE is not permitted.'
    }
}

function Invoke-IxmReadOnlyQuery {
    param(
        [Parameter(Mandatory)]$Connection,
        [Parameter(Mandatory)][string]$Sql
    )

    Test-IxmSelectOnlySql -Sql $Sql

    $Command = $Connection.CreateCommand()
    $Command.CommandText = $Sql
    $Command.CommandTimeout = 120

    $Adapter = New-Object System.Data.Odbc.OdbcDataAdapter $Command
    $Table = New-Object System.Data.DataTable
    [void]$Adapter.Fill($Table)

    # DataTable implements IEnumerable. Prevent PowerShell from silently
    # converting the table into DataRow objects when returning it.
    Write-Output -InputObject $Table -NoEnumerate
}

function Get-IxmMailboxDirectory {
    param([Parameter(Mandatory)]$Connection)

    $Columns = $Connection.GetSchema('Columns')

    $HasImapName = Test-IxmColumn -Columns $Columns -Table 'MAILBOX' -Column 'IMAPNAME'
    $HasUserName = Test-IxmColumn -Columns $Columns -Table 'MAILBOX' -Column 'USERNAME'

    if ($HasImapName) {
        $EmailExpression = 'm.IMAPNAME AS EmailAddress'
        $EmailSource = 'MAILBOX.IMAPNAME'
    }
    else {
        # Keep the report usable on schemas that do not expose IMAPNAME.
        # USERNAME is a fallback only; it may not be a Graph mailbox address.
        $EmailExpression = if ($HasUserName) {
            'm.USERNAME AS EmailAddress'
        }
        else {
            "CAST(NULL AS VARCHAR(1)) AS EmailAddress"
        }

        $EmailSource = if ($HasUserName) { 'MAILBOX.USERNAME (fallback)' } else { 'not available' }
    }

    $Sql = @"
SELECT
    m.MBXNUMBER AS Extension,
    m.FIRSTNAME AS FirstName,
    m.LASTNAME AS LastName,
    f.FGNAME AS FeatureGroup,
    $EmailExpression
FROM DBA.MAILBOX m
LEFT JOIN DBA.FGROUP f
    ON m.FGROUPID = f.FGROUPID
WHERE m.MBXNUMBER IS NOT NULL
  AND m.MBXNUMBER <> ''
ORDER BY m.MBXNUMBER
"@

    $Table = Invoke-IxmReadOnlyQuery -Connection $Connection -Sql $Sql

    if ($Table -is [System.Data.DataTable]) {
        $SourceRows = @($Table.Rows)
    }
    elseif ($Table -is [System.Data.DataRow]) {
        $SourceRows = @($Table)
    }
    elseif ($Table -is [System.Array]) {
        $SourceRows = @($Table)
    }
    else {
        throw ('Unexpected SQL query result type: {0}' -f $Table.GetType().FullName)
    }

    $Rows = foreach ($Row in $SourceRows) {
        $First = if ($Row.IsNull('FirstName')) { '' } else { [string]$Row.FirstName }
        $Last = if ($Row.IsNull('LastName')) { '' } else { [string]$Row.LastName }
        $Email = if ($Row.IsNull('EmailAddress')) { '' } else { [string]$Row.EmailAddress }

        $Name = (($First.Trim() + ' ' + $Last.Trim()).Trim())

        [pscustomobject]@{
            Name = $Name
            Extension = if ($Row.IsNull('Extension')) { '' } else { [string]$Row.Extension }
            FeatureGroup = if ($Row.IsNull('FeatureGroup')) { '' } else { [string]$Row.FeatureGroup }
            EmailAddress = $Email.Trim()
        }
    }

    return [pscustomobject]@{
        Rows = @($Rows | Sort-Object Extension)
        EmailSource = $EmailSource
    }
}

function Invoke-MailboxDirectoryExport {
    Write-Section 'Mailbox Directory / Email Export'

    Write-Host 'This tool submits SELECT-only database queries. Effective database permissions are controlled by the configured SQL Anywhere DSN/account.' -ForegroundColor Green
    Write-Host 'Searching installed SQL Anywhere System DSNs...' -ForegroundColor DarkGray
    Write-Host ''

    $Selected = Select-IxmDatabaseDsn

    Write-Host ''
    Write-Host ('Using DSN     : {0}' -f $Selected.Name) -ForegroundColor Cyan
    Write-Host ('Database      : {0}' -f $(if ($Selected.DatabaseName) { $Selected.DatabaseName } else { '(not reported by ODBC)' }))
    Write-Host ('Server        : {0}' -f $(if ($Selected.ServerName) { $Selected.ServerName } else { '(not reported by ODBC)' }))
    Write-Host ('ODBC platform : {0}' -f $Selected.Platform)
    Write-Host ''

    $Connection = $null

    try {
        $Connection = Open-IxmDsnConnection -Name $Selected.Name

        if (-not (Test-IxmMailboxSchema -Connection $Connection)) {
            throw "DSN '$($Selected.Name)' connected, but the expected IX Messaging MAILBOX/FGROUP schema was not found."
        }

        Write-Host 'Reading IX Messaging mailbox configuration...' -ForegroundColor DarkGray
        $Directory = Get-IxmMailboxDirectory -Connection $Connection
        $Rows = @($Directory.Rows)

        if ($Rows.Count -eq 0) {
            Write-Host 'No mailbox records were returned.' -ForegroundColor Yellow
            return
        }

        Write-Host ''
        Write-Host ('Mailbox / Email Results ({0} mailbox(es))' -f $Rows.Count) -ForegroundColor Cyan
        Write-Host ''

        $Rows |
            Select-Object Name,Extension,FeatureGroup,EmailAddress |
            Format-Table -AutoSize -Wrap |
            Out-Host

        $WithEmail = @($Rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.EmailAddress) }).Count
        $WithoutEmail = $Rows.Count - $WithEmail

        Write-Host ''
        Write-Host ('Total mailboxes       : {0}' -f $Rows.Count)
        Write-Host ('With email address    : {0}' -f $WithEmail)
        Write-Host ('Without email address : {0}' -f $WithoutEmail)
        Write-Host ('Email source          : {0}' -f $Directory.EmailSource) -ForegroundColor DarkGray

        if ($Directory.EmailSource -ne 'MAILBOX.IMAPNAME') {
            Write-Host 'WARNING: IMAPNAME is not available on this schema; EmailAddress is using a fallback field.' -ForegroundColor Yellow
        }
        else {
            Write-Host 'Note: A populated IMAPNAME is a configured mailbox/account address; by itself it does not prove Graph sync is enabled or healthy.' -ForegroundColor DarkGray
        }

        # Reuse the tool's standard Save As / CSV workflow.
        Export-ResultSet -Data $Rows -BaseName 'IXM-Mailboxes'
    }
    finally {
        if ($null -ne $Connection) {
            try { $Connection.Close() } catch { Write-Verbose ('ODBC connection Close() cleanup failed: {0}' -f $_.Exception.Message) }
            try { $Connection.Dispose() } catch { Write-Verbose ('ODBC connection Dispose() cleanup failed: {0}' -f $_.Exception.Message) }
        }
    }
}


# -----------------------------------------------------------------------------
# IX Messaging database current mailbox status / health
# -----------------------------------------------------------------------------

function Test-IxmTable {
    param(
        [Parameter(Mandatory)]$Tables,
        [Parameter(Mandatory)][string]$Table
    )

    return @(
        $Tables | Where-Object {
            $_.TABLE_NAME -eq $Table -and
            ($_.TABLE_SCHEM -eq 'DBA' -or [string]::IsNullOrWhiteSpace([string]$_.TABLE_SCHEM))
        }
    ).Count -gt 0
}

function Get-IxmResultRows {
    param([Parameter(Mandatory)]$Result)

    if ($Result -is [System.Data.DataTable]) {
        # DataTable.Select() returns a real DataRow[] and avoids PowerShell 5.1
        # collection-adapter issues seen with @($Result.Rows).
        return $Result.Select()
    }
    elseif ($Result -is [System.Data.DataRow]) {
        return ,$Result
    }
    elseif ($Result -is [System.Array]) {
        return $Result
    }

    throw ('Unexpected SQL query result type: {0}' -f $Result.GetType().FullName)
}

function Convert-IxmBooleanText {
    param(
        $Value,
        [string]$TrueText = 'Yes',
        [string]$FalseText = 'No',
        [string]$UnknownText = 'Not available'
    )

    if ($null -eq $Value -or $Value -is [System.DBNull]) {
        return $UnknownText
    }

    try {
        if ([System.Convert]::ToBoolean($Value)) {
            return $TrueText
        }

        return $FalseText
    }
    catch {
        return $UnknownText
    }
}

function Convert-IxmDateTimeText {
    param($Value)

    if ($null -eq $Value -or $Value -is [System.DBNull]) {
        return ''
    }

    if ($Value -is [datetime]) {
        return ([datetime]$Value).ToString('MM/dd/yyyy hh:mm:ss tt')
    }

    $Text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    foreach ($Format in @('yyyyMMddHHmmss','yyyyMMddHHmm','yyyyMMdd')) {
        $Parsed = [datetime]::MinValue

        if ([datetime]::TryParseExact(
            $Text,
            $Format,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None,
            [ref]$Parsed
        )) {
            if ($Format -eq 'yyyyMMdd') {
                return $Parsed.ToString('MM/dd/yyyy')
            }

            return $Parsed.ToString('MM/dd/yyyy hh:mm:ss tt')
        }
    }

    return $Text
}

function Get-IxmMailboxHealth {
    param(
        [Parameter(Mandatory)]$Connection,
        [Parameter(Mandatory)][string]$Extension
    )

    # Extension comes from Read-NumericValue in the caller, but validate again
    # before embedding it in the read-only SQL statement.
    if ($Extension -notmatch '^\d+$') {
        throw 'Mailbox / extension must contain digits only.'
    }

    $Tables = $Connection.GetSchema('Tables')
    $Columns = $Connection.GetSchema('Columns')

    # ---------------------------------------------------------------------
    # Base mailbox record
    # ---------------------------------------------------------------------
    $HasTutorial = Test-IxmColumn -Columns $Columns -Table 'MAILBOX' -Column 'TUTORIAL'
    $HasLocked = Test-IxmColumn -Columns $Columns -Table 'MAILBOX' -Column 'MbxLocked'
    $HasLockedTime = Test-IxmColumn -Columns $Columns -Table 'MAILBOX' -Column 'MbxLockedTime'
    $HasAttempts = Test-IxmColumn -Columns $Columns -Table 'MAILBOX' -Column 'MbxPassNumAttempts'
    $HasImapName = Test-IxmColumn -Columns $Columns -Table 'MAILBOX' -Column 'IMAPNAME'
    $HasUserName = Test-IxmColumn -Columns $Columns -Table 'MAILBOX' -Column 'USERNAME'
    $HasUseImap = Test-IxmColumn -Columns $Columns -Table 'MAILBOX' -Column 'MbxUseMbxIMAP'

    $TutorialExpr = if ($HasTutorial) {
        'm.TUTORIAL AS Tutorial'
    }
    else {
        'CAST(NULL AS INTEGER) AS Tutorial'
    }

    $LockedExpr = if ($HasLocked) {
        'm.MbxLocked AS MailboxLocked'
    }
    else {
        'CAST(NULL AS INTEGER) AS MailboxLocked'
    }

    $LockedTimeExpr = if ($HasLockedTime) {
        'm.MbxLockedTime AS MailboxLockedTime'
    }
    else {
        "CAST(NULL AS VARCHAR(1)) AS MailboxLockedTime"
    }

    $AttemptsExpr = if ($HasAttempts) {
        'm.MbxPassNumAttempts AS FailedPINAttempts'
    }
    else {
        'CAST(NULL AS INTEGER) AS FailedPINAttempts'
    }

    if ($HasImapName) {
        $EmailExpr = 'm.IMAPNAME AS EmailAddress'
        $EmailSource = 'MAILBOX.IMAPNAME'
    }
    elseif ($HasUserName) {
        $EmailExpr = 'm.USERNAME AS EmailAddress'
        $EmailSource = 'MAILBOX.USERNAME (fallback)'
    }
    else {
        $EmailExpr = "CAST(NULL AS VARCHAR(1)) AS EmailAddress"
        $EmailSource = 'not available'
    }

    $UseImapExpr = if ($HasUseImap) {
        'm.MbxUseMbxIMAP AS GraphIMAPConfigured'
    }
    else {
        'CAST(NULL AS INTEGER) AS GraphIMAPConfigured'
    }

    $MailboxSql = @"
SELECT
    m.MBXID,
    m.MBXNUMBER AS Extension,
    m.FIRSTNAME AS FirstName,
    m.LASTNAME AS LastName,
    f.FGNAME AS FeatureGroup,
    $TutorialExpr,
    $LockedExpr,
    $LockedTimeExpr,
    $AttemptsExpr,
    $EmailExpr,
    $UseImapExpr
FROM DBA.MAILBOX m
LEFT JOIN DBA.FGROUP f
    ON m.FGROUPID = f.FGROUPID
WHERE m.MBXNUMBER = '$Extension'
"@

    $MailboxResult = Invoke-IxmReadOnlyQuery -Connection $Connection -Sql $MailboxSql
    $MailboxRows = @(Get-IxmResultRows -Result $MailboxResult)

    if ($MailboxRows.Count -eq 0) {
        return $null
    }

    $Mailbox = $MailboxRows[0]
    $MailboxId = [int]$Mailbox.MBXID

    $FirstName = if ($Mailbox.IsNull('FirstName')) { '' } else { [string]$Mailbox.FirstName }
    $LastName = if ($Mailbox.IsNull('LastName')) { '' } else { [string]$Mailbox.LastName }
    $Name = (($FirstName.Trim() + ' ' + $LastName.Trim()).Trim())

    $FeatureGroup = if ($Mailbox.IsNull('FeatureGroup')) { '' } else { [string]$Mailbox.FeatureGroup }
    $EmailAddress = if ($Mailbox.IsNull('EmailAddress')) { '' } else { ([string]$Mailbox.EmailAddress).Trim() }

    $TutorialRaw = if ($Mailbox.IsNull('Tutorial')) { $null } else { $Mailbox.Tutorial }
    $LockedRaw = if ($Mailbox.IsNull('MailboxLocked')) { $null } else { $Mailbox.MailboxLocked }
    $FailedAttempts = if ($Mailbox.IsNull('FailedPINAttempts')) { $null } else { [int]$Mailbox.FailedPINAttempts }
    $LockedTime = if ($Mailbox.IsNull('MailboxLockedTime')) { '' } else { Convert-IxmDateTimeText -Value $Mailbox.MailboxLockedTime }
    $GraphRaw = if ($Mailbox.IsNull('GraphIMAPConfigured')) { $null } else { $Mailbox.GraphIMAPConfigured }

    $TutorialState = 'Not available'
    $SetupStatus = 'Not available'
    $SetupWarning = $false

    if ($null -ne $TutorialRaw) {
        if ([System.Convert]::ToBoolean($TutorialRaw)) {
            $TutorialState = 'PENDING'
            $SetupStatus = 'INITIAL SETUP NOT COMPLETED'
            $SetupWarning = $true
        }
        else {
            $TutorialState = 'Completed / Disabled'
            $SetupStatus = 'No tutorial pending'
        }
    }

    $MailboxLockedText = Convert-IxmBooleanText -Value $LockedRaw -TrueText 'Yes' -FalseText 'No'
    $GraphConfiguredText = Convert-IxmBooleanText -Value $GraphRaw -TrueText 'Yes' -FalseText 'No'

    # ---------------------------------------------------------------------
    # Current Inbox voice-message counters
    # ---------------------------------------------------------------------
    $InboxCountsAvailable = $false
    $InboxFolderId = $null
    $VoiceNormalUnread = $null
    $VoiceUrgentUnread = $null
    $VoiceNormalRead = $null
    $VoiceUrgentRead = $null
    $VoiceUnread = $null
    $VoiceRead = $null
    $VoiceTotal = $null

    $HasFolders = Test-IxmTable -Tables $Tables -Table 'FOLDERS'
    $HasFolderCounts = Test-IxmTable -Tables $Tables -Table 'FolderMsgsCount'

    $RequiredFolderColumns = @('FOLDERID','FOLDERNAME','FOLDERTYPE','MBXID')
    $RequiredCountColumns = @(
        'FOLDERID',
        'FldMsgsVoiceNotUrgUnread',
        'FldMsgsVoiceUrgUnread',
        'FldMsgsVoiceNotUrgRead',
        'FldMsgsVoiceUrgRead'
    )

    $FolderColumnsOk = $HasFolders
    foreach ($Column in $RequiredFolderColumns) {
        if (-not (Test-IxmColumn -Columns $Columns -Table 'FOLDERS' -Column $Column)) {
            $FolderColumnsOk = $false
        }
    }

    $CountColumnsOk = $HasFolderCounts
    foreach ($Column in $RequiredCountColumns) {
        if (-not (Test-IxmColumn -Columns $Columns -Table 'FolderMsgsCount' -Column $Column)) {
            $CountColumnsOk = $false
        }
    }

    if ($FolderColumnsOk -and $CountColumnsOk) {
        $CountsSql = @"
SELECT
    f.FOLDERID,
    f.FOLDERNAME,
    f.FOLDERTYPE,
    COALESCE(c.FldMsgsVoiceNotUrgUnread,0) AS VoiceNormalUnread,
    COALESCE(c.FldMsgsVoiceUrgUnread,0)    AS VoiceUrgentUnread,
    COALESCE(c.FldMsgsVoiceNotUrgRead,0)  AS VoiceNormalRead,
    COALESCE(c.FldMsgsVoiceUrgRead,0)     AS VoiceUrgentRead
FROM DBA.FOLDERS f
LEFT JOIN DBA.FolderMsgsCount c
    ON f.FOLDERID = c.FOLDERID
WHERE f.MBXID = $MailboxId
ORDER BY f.FOLDERID
"@

        $CountsResult = Invoke-IxmReadOnlyQuery -Connection $Connection -Sql $CountsSql
        $FolderRows = @(Get-IxmResultRows -Result $CountsResult)

        # FOLDERTYPE 2 was validated as Inbox on both lab and production.
        $Inbox = @(
            $FolderRows |
            Where-Object {
                $TypeIsInbox = $false
                $NameIsInbox = $false

                if (-not $_.IsNull('FOLDERTYPE')) {
                    $TypeIsInbox = ([int]$_.FOLDERTYPE -eq 2)
                }

                if (-not $_.IsNull('FOLDERNAME')) {
                    $NameIsInbox = ([string]$_.FOLDERNAME -eq 'Inbox')
                }

                $TypeIsInbox -or $NameIsInbox
            } |
            Select-Object -First 1
        )

        if ($Inbox.Count -gt 0) {
            $InboxRow = $Inbox[0]
            $InboxCountsAvailable = $true
            $InboxFolderId = [int]$InboxRow.FOLDERID
            $VoiceNormalUnread = [int]$InboxRow.VoiceNormalUnread
            $VoiceUrgentUnread = [int]$InboxRow.VoiceUrgentUnread
            $VoiceNormalRead = [int]$InboxRow.VoiceNormalRead
            $VoiceUrgentRead = [int]$InboxRow.VoiceUrgentRead
            $VoiceUnread = $VoiceNormalUnread + $VoiceUrgentUnread
            $VoiceRead = $VoiceNormalRead + $VoiceUrgentRead
            $VoiceTotal = $VoiceUnread + $VoiceRead
        }
    }

    # ---------------------------------------------------------------------
    # Current MWI and synchronization timestamps
    # ---------------------------------------------------------------------
    $MwiAvailable = $false
    $MwiRaw = $null
    $MwiState = 'Not available'
    $MwiUpdateTime = ''
    $LastInboxSync = ''
    $LastCalendarSync = ''

    $HasMbxMwi = Test-IxmTable -Tables $Tables -Table 'MbxMWI'
    $HasMwiStatus = Test-IxmColumn -Columns $Columns -Table 'MbxMWI' -Column 'MSGMWISTATUS'
    $HasMwiUpdate = Test-IxmColumn -Columns $Columns -Table 'MbxMWI' -Column 'MWIUpdateTime'
    $HasLastInboxSync = Test-IxmColumn -Columns $Columns -Table 'MbxMWI' -Column 'LastSyncInboxDateTime'
    $HasLastCalendarSync = Test-IxmColumn -Columns $Columns -Table 'MbxMWI' -Column 'LastCalendarSyncDateTime'

    if ($HasMbxMwi -and $HasMwiStatus) {
        $MwiUpdateExpr = if ($HasMwiUpdate) {
            'MWIUpdateTime'
        }
        else {
            'CAST(NULL AS TIMESTAMP) AS MWIUpdateTime'
        }

        $InboxSyncExpr = if ($HasLastInboxSync) {
            'LastSyncInboxDateTime'
        }
        else {
            "CAST(NULL AS VARCHAR(1)) AS LastSyncInboxDateTime"
        }

        $CalendarSyncExpr = if ($HasLastCalendarSync) {
            'LastCalendarSyncDateTime'
        }
        else {
            "CAST(NULL AS VARCHAR(1)) AS LastCalendarSyncDateTime"
        }

        $MwiSql = @"
SELECT
    MSGMWISTATUS,
    $MwiUpdateExpr,
    $InboxSyncExpr,
    $CalendarSyncExpr
FROM DBA.MbxMWI
WHERE MBXID = $MailboxId
"@

        $MwiResult = Invoke-IxmReadOnlyQuery -Connection $Connection -Sql $MwiSql
        $MwiRows = @(Get-IxmResultRows -Result $MwiResult)

        if ($MwiRows.Count -gt 0) {
            $Mwi = $MwiRows[0]

            if (-not $Mwi.IsNull('MSGMWISTATUS')) {
                $MwiAvailable = $true
                $MwiRaw = $Mwi.MSGMWISTATUS
                if ([System.Convert]::ToBoolean($MwiRaw)) {
                    $MwiState = 'ON'
                }
                else {
                    $MwiState = 'OFF'
                }
            }

            if (-not $Mwi.IsNull('MWIUpdateTime')) {
                $MwiUpdateTime = Convert-IxmDateTimeText -Value $Mwi.MWIUpdateTime
            }

            if (-not $Mwi.IsNull('LastSyncInboxDateTime')) {
                $LastInboxSync = Convert-IxmDateTimeText -Value $Mwi.LastSyncInboxDateTime
            }

            if (-not $Mwi.IsNull('LastCalendarSyncDateTime')) {
                $LastCalendarSync = Convert-IxmDateTimeText -Value $Mwi.LastCalendarSyncDateTime
            }
        }
    }

    # ---------------------------------------------------------------------
    # Compare current unread Inbox count to current MWI state.
    # ---------------------------------------------------------------------
    $MwiCheck = 'Not available'
    $MwiCheckDetail = 'Current Inbox message counts and/or MWI state are not available.'

    if ($InboxCountsAvailable -and $MwiAvailable) {
        $MwiOn = [System.Convert]::ToBoolean($MwiRaw)

        if ($VoiceUnread -gt 0 -and $MwiOn) {
            $MwiCheck = 'OK'
            $MwiCheckDetail = 'Unread Inbox voicemail is present and MWI is ON.'
        }
        elseif ($VoiceUnread -eq 0 -and -not $MwiOn) {
            $MwiCheck = 'OK'
            $MwiCheckDetail = 'No unread Inbox voicemail is present and MWI is OFF.'
        }
        elseif ($VoiceUnread -gt 0 -and -not $MwiOn) {
            $MwiCheck = 'POSSIBLE MWI MISMATCH'
            $MwiCheckDetail = 'Unread Inbox voicemail is present, but IX Messaging currently reports MWI OFF.'
        }
        elseif ($VoiceUnread -eq 0 -and $MwiOn) {
            $MwiCheck = 'POSSIBLE STUCK MWI'
            $MwiCheckDetail = 'No unread Inbox voicemail is present, but IX Messaging currently reports MWI ON.'
        }
    }

    return [pscustomobject]@{
        Name = $Name
        Extension = [string]$Mailbox.Extension
        MailboxID = $MailboxId
        FeatureGroup = $FeatureGroup
        EmailAddress = $EmailAddress
        EmailSource = $EmailSource

        TutorialState = $TutorialState
        SetupStatus = $SetupStatus
        SetupWarning = $SetupWarning
        MailboxLocked = $MailboxLockedText
        MailboxLockedTime = $LockedTime
        FailedPINAttempts = $FailedAttempts

        InboxCountsAvailable = $InboxCountsAvailable
        InboxFolderID = $InboxFolderId
        InboxUnread = $VoiceUnread
        InboxRead = $VoiceRead
        UrgentUnread = $VoiceUrgentUnread
        UrgentRead = $VoiceUrgentRead
        TotalInboxVoice = $VoiceTotal

        MWIState = $MwiState
        MWIUpdateTime = $MwiUpdateTime
        MWIMessageCheck = $MwiCheck
        MWIMessageCheckDetail = $MwiCheckDetail

        GraphIMAPConfigured = $GraphConfiguredText
        LastInboxSync = $LastInboxSync
        LastCalendarSync = $LastCalendarSync
    }
}

function Show-IxmMailboxHealth {
    param([Parameter(Mandatory)]$Health)

    Write-Host 'MAILBOX' -ForegroundColor Cyan
    Write-Host ('-' * 78)
    Write-Host ('{0,-24}: {1}' -f 'Name',$Health.Name)
    Write-Host ('{0,-24}: {1}' -f 'Extension',$Health.Extension)
    Write-Host ('{0,-24}: {1}' -f 'IXM Mailbox ID',$Health.MailboxID)
    Write-Host ('{0,-24}: {1}' -f 'Feature Group',$Health.FeatureGroup)
    Write-Host ('{0,-24}: {1}' -f 'Email Address',$Health.EmailAddress)

    Write-Host ''
    Write-Host 'MAILBOX SETUP / SECURITY' -ForegroundColor Cyan
    Write-Host ('-' * 78)

    if ($Health.SetupWarning) {
        Write-Host ('{0,-24}: {1}' -f 'Initial Tutorial',$Health.TutorialState) -ForegroundColor Yellow
        Write-Host ('{0,-24}: {1}' -f 'Setup Status',$Health.SetupStatus) -ForegroundColor Yellow
    }
    else {
        Write-Host ('{0,-24}: {1}' -f 'Initial Tutorial',$Health.TutorialState)
        Write-Host ('{0,-24}: {1}' -f 'Setup Status',$Health.SetupStatus)
    }

    $LockColor = if ($Health.MailboxLocked -eq 'Yes') { 'Yellow' } else { 'Gray' }
    Write-Host ('{0,-24}: {1}' -f 'Mailbox Locked',$Health.MailboxLocked) -ForegroundColor $LockColor

    if (-not [string]::IsNullOrWhiteSpace([string]$Health.MailboxLockedTime)) {
        Write-Host ('{0,-24}: {1}' -f 'Mailbox Locked Time',$Health.MailboxLockedTime)
    }

    $AttemptsText = if ($null -eq $Health.FailedPINAttempts) { 'Not available' } else { [string]$Health.FailedPINAttempts }
    Write-Host ('{0,-24}: {1}' -f 'Failed PIN Attempts',$AttemptsText)

    Write-Host ''
    Write-Host 'VOICE MESSAGES - INBOX' -ForegroundColor Cyan
    Write-Host ('-' * 78)

    if ($Health.InboxCountsAvailable) {
        Write-Host ('{0,-24}: {1}' -f 'Unread',$Health.InboxUnread)
        Write-Host ('{0,-24}: {1}' -f 'Read',$Health.InboxRead)
        Write-Host ('{0,-24}: {1}' -f 'Urgent Unread',$Health.UrgentUnread)
        Write-Host ('{0,-24}: {1}' -f 'Urgent Read',$Health.UrgentRead)
        Write-Host ('{0,-24}: {1}' -f 'Total',$Health.TotalInboxVoice)
    }
    else {
        Write-Host 'Current Inbox voice-message counters are not available on this database/schema.' -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host 'MWI' -ForegroundColor Cyan
    Write-Host ('-' * 78)

    $MwiColor = if ($Health.MWIState -eq 'ON') { 'Yellow' } elseif ($Health.MWIState -eq 'OFF') { 'Green' } else { 'Gray' }
    Write-Host ('{0,-24}: {1}' -f 'Current MWI State',$Health.MWIState) -ForegroundColor $MwiColor
    Write-Host ('{0,-24}: {1}' -f 'MWI Last Updated',$(if ($Health.MWIUpdateTime) { $Health.MWIUpdateTime } else { 'Not available' }))

    Write-Host ''
    Write-Host 'MWI / MESSAGE CHECK' -ForegroundColor Cyan
    Write-Host ('-' * 78)
    Write-Host ('{0,-24}: {1}' -f 'Unread Inbox Messages',$(if ($null -ne $Health.InboxUnread) { $Health.InboxUnread } else { 'Not available' }))
    Write-Host ('{0,-24}: {1}' -f 'MWI State',$Health.MWIState)

    if ($Health.MWIMessageCheck -eq 'OK') {
        Write-Host ('{0,-24}: {1}' -f 'Status',$Health.MWIMessageCheck) -ForegroundColor Green
        Write-Host ('  {0}' -f $Health.MWIMessageCheckDetail) -ForegroundColor DarkGray
    }
    elseif ($Health.MWIMessageCheck -match '^POSSIBLE') {
        Write-Host ('{0,-24}: {1}' -f 'Status',$Health.MWIMessageCheck) -ForegroundColor Yellow
        Write-Host ('  WARNING: {0}' -f $Health.MWIMessageCheckDetail) -ForegroundColor Yellow
    }
    else {
        Write-Host ('{0,-24}: {1}' -f 'Status',$Health.MWIMessageCheck)
        Write-Host ('  {0}' -f $Health.MWIMessageCheckDetail) -ForegroundColor DarkGray
    }

    Write-Host ''
    Write-Host 'EMAIL / GRAPH' -ForegroundColor Cyan
    Write-Host ('-' * 78)
    Write-Host ('{0,-24}: {1}' -f 'Graph / IMAP Configured',$Health.GraphIMAPConfigured)
    Write-Host ('{0,-24}: {1}' -f 'Email Address',$Health.EmailAddress)
    Write-Host ('{0,-24}: {1}' -f 'Last Inbox Sync',$(if ($Health.LastInboxSync) { $Health.LastInboxSync } else { 'Not available' }))
    Write-Host ('{0,-24}: {1}' -f 'Last Calendar Sync',$(if ($Health.LastCalendarSync) { $Health.LastCalendarSync } else { 'Not available' }))

    Write-Host ''
    Write-Host ('Email source: {0}' -f $Health.EmailSource) -ForegroundColor DarkGray
    Write-Host 'Message totals above count the Inbox only; Sent, Deleted Items, and other folders are intentionally excluded.' -ForegroundColor DarkGray

    if ($Health.SetupWarning) {
        Write-Host ''
        Write-Host 'WARNING: The subscriber initial mailbox tutorial is still pending.' -ForegroundColor Yellow
        Write-Host 'This indicates that initial mailbox setup has not been completed; it is not presented as proof that the subscriber has never entered the mailbox.' -ForegroundColor Yellow
    }
}

function Convert-IxmMailboxHealthToExportRow {
    param([Parameter(Mandatory)]$Health)

    return [pscustomobject]@{
        Name = $Health.Name
        Extension = $Health.Extension
        MailboxID = $Health.MailboxID
        FeatureGroup = $Health.FeatureGroup
        EmailAddress = $Health.EmailAddress
        InitialTutorial = $Health.TutorialState
        SetupStatus = $Health.SetupStatus
        MailboxLocked = $Health.MailboxLocked
        MailboxLockedTime = $Health.MailboxLockedTime
        FailedPINAttempts = $Health.FailedPINAttempts
        InboxUnread = $Health.InboxUnread
        InboxRead = $Health.InboxRead
        UrgentUnread = $Health.UrgentUnread
        UrgentRead = $Health.UrgentRead
        TotalInboxVoice = $Health.TotalInboxVoice
        CurrentMWI = $Health.MWIState
        MWIUpdateTime = $Health.MWIUpdateTime
        MWIMessageCheck = $Health.MWIMessageCheck
        GraphIMAPConfigured = $Health.GraphIMAPConfigured
        LastInboxSync = $Health.LastInboxSync
        LastCalendarSync = $Health.LastCalendarSync
    }
}

function Invoke-CurrentMailboxHealth {
    $Extension = Read-NumericValue -Prompt 'Mailbox / extension'

    Write-Section ("Current Mailbox Status / Health - {0}" -f $Extension)

    Write-Host 'This tool submits SELECT-only database queries. Effective database permissions are controlled by the configured SQL Anywhere DSN/account.' -ForegroundColor Green
    Write-Host 'Searching installed SQL Anywhere System DSNs...' -ForegroundColor DarkGray
    Write-Host ''

    $Selected = Select-IxmDatabaseDsn

    Write-Host ''
    Write-Host ('Using DSN     : {0}' -f $Selected.Name) -ForegroundColor Cyan
    Write-Host ('Database      : {0}' -f $(if ($Selected.DatabaseName) { $Selected.DatabaseName } else { '(not reported by ODBC)' }))
    Write-Host ('Server        : {0}' -f $(if ($Selected.ServerName) { $Selected.ServerName } else { '(not reported by ODBC)' }))
    Write-Host ('ODBC platform : {0}' -f $Selected.Platform)
    Write-Host ''

    $Connection = $null

    try {
        $Connection = Open-IxmDsnConnection -Name $Selected.Name

        if (-not (Test-IxmMailboxSchema -Connection $Connection)) {
            throw "DSN '$($Selected.Name)' connected, but the expected IX Messaging MAILBOX/FGROUP schema was not found."
        }

        Write-Host 'Reading current IX Messaging mailbox state...' -ForegroundColor DarkGray
        Write-Host ''

        $Health = Get-IxmMailboxHealth -Connection $Connection -Extension $Extension

        if ($null -eq $Health) {
            Write-Host ('Mailbox / extension {0} was not found in DBA.MAILBOX.' -f $Extension) -ForegroundColor Yellow
            return
        }

        Show-IxmMailboxHealth -Health $Health

        $ExportRow = Convert-IxmMailboxHealthToExportRow -Health $Health
        Export-ResultSet -Data @($ExportRow) -BaseName ("mailbox_{0}_current_health" -f $Extension)
    }
    finally {
        if ($null -ne $Connection) {
            try { $Connection.Close() } catch { Write-Verbose ('ODBC connection Close() cleanup failed: {0}' -f $_.Exception.Message) }
            try { $Connection.Dispose() } catch { Write-Verbose ('ODBC connection Dispose() cleanup failed: {0}' -f $_.Exception.Message) }
        }
    }
}


# -----------------------------------------------------------------------------
# Graph / Exchange mailbox failure audit
# -----------------------------------------------------------------------------

function Get-CseGraphRootFromLogRoot {
    if ([string]::IsNullOrWhiteSpace([string]$LogRoot)) {
        return $null
    }

    try {
        $LogsRoot = Split-Path -Path $LogRoot -Parent
        if ([string]::IsNullOrWhiteSpace([string]$LogsRoot)) {
            return $null
        }

        return (Join-Path $LogsRoot 'uccse\CSE')
    }
    catch {
        return $null
    }
}

function Resolve-CseGraphRoot {
    if ($script:CseGraphRoot -and (Test-Path -LiteralPath $script:CseGraphRoot)) {
        return $script:CseGraphRoot
    }

    $Derived = Get-CseGraphRootFromLogRoot
    if ($Derived -and (Test-Path -LiteralPath $Derived)) {
        $script:CseGraphRoot = $Derived
        return $Derived
    }

    Write-Host ''
    Write-Host 'Graph/CSE error log directory was not found from the current VServer path.' -ForegroundColor Yellow
    Write-Host 'Checking local fixed drives for \UC\logs\uccse\CSE...' -ForegroundColor DarkGray

    $Detected = New-Object System.Collections.Generic.List[string]

    try {
        $Drives = @(
            Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop |
            Select-Object -ExpandProperty DeviceID
        )
    }
    catch {
        $Drives = @(
            Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Root.TrimEnd('\') }
        )
    }

    foreach ($Drive in @($Drives | Sort-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace([string]$Drive)) { continue }

        $Candidate = Join-Path ($Drive.TrimEnd('\') + '\') 'UC\logs\uccse\CSE'
        if (Test-Path -LiteralPath $Candidate) {
            if (-not $Detected.Contains($Candidate)) {
                $Detected.Add($Candidate)
            }
        }
    }

    if ($Detected.Count -eq 1) {
        $script:CseGraphRoot = $Detected[0]
        Write-Host ('Detected Graph/CSE logs: {0}' -f $script:CseGraphRoot) -ForegroundColor Green
        return $script:CseGraphRoot
    }

    if ($Detected.Count -gt 1) {
        Write-Host ''
        Write-Host 'Multiple Graph/CSE log directories were detected:' -ForegroundColor Yellow

        for ($i = 0; $i -lt $Detected.Count; $i++) {
            Write-Host ('  {0}. {1}' -f ($i + 1),$Detected[$i])
        }

        do {
            $Selection = (Read-Host ('Select 1-{0}, or press Enter to enter a different location' -f $Detected.Count)).Trim()

            if ([string]::IsNullOrWhiteSpace($Selection)) {
                break
            }

            $Number = 0
            if ([int]::TryParse($Selection,[ref]$Number) -and
                $Number -ge 1 -and
                $Number -le $Detected.Count) {

                $script:CseGraphRoot = $Detected[$Number - 1]
                return $script:CseGraphRoot
            }

            Write-Host 'Invalid selection.' -ForegroundColor Yellow
        } while ($true)
    }

    Write-Host ''
    Write-Host 'Enter the IX Messaging drive letter or full CSE directory.' -ForegroundColor Cyan
    Write-Host 'Examples:' -ForegroundColor DarkGray
    Write-Host '  X:' -ForegroundColor DarkGray
    Write-Host '  X:\UC\logs\uccse\CSE' -ForegroundColor DarkGray

    while ($true) {
        $Entered = (Read-Host 'IX Messaging drive or CSE log directory').Trim()

        if ([string]::IsNullOrWhiteSpace($Entered)) {
            throw 'Graph/CSE log directory is required for the Graph / email failure audit.'
        }

        $Candidate = $Entered

        if ($Entered -match '^(?<Drive>[A-Za-z])(?::)?(?:\\)?$') {
            $Candidate = ('{0}:\UC\logs\uccse\CSE' -f $Matches.Drive.ToUpper())
        }

        if (Test-Path -LiteralPath $Candidate) {
            $script:CseGraphRoot = $Candidate
            Write-Host ('Using Graph/CSE logs: {0}' -f $script:CseGraphRoot) -ForegroundColor Green
            return $script:CseGraphRoot
        }

        Write-Host ('Directory not found: {0}' -f $Candidate) -ForegroundColor Yellow
    }
}

function Get-CseGraphErrorFileInventory {
    param([Parameter(Mandatory)][string]$CseRoot)

    if (-not (Test-Path -LiteralPath $CseRoot)) {
        return @()
    }

    $Pattern = '^ERR\.SESGRFM\.(?<Date>\d{4}-\d{2}-\d{2})T(?<Hour>\d{2})\.csv$'
    $Results = New-Object System.Collections.Generic.List[object]

    foreach ($File in @(Get-ChildItem -LiteralPath $CseRoot -File -ErrorAction SilentlyContinue)) {
        if ($File.Name -notmatch $Pattern) {
            continue
        }

        $StampText = ('{0} {1}' -f $Matches.Date,$Matches.Hour)
        $Stamp = [datetime]::MinValue

        if (-not [datetime]::TryParseExact(
            $StampText,
            'yyyy-MM-dd HH',
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None,
            [ref]$Stamp
        )) {
            continue
        }

        $Results.Add([pscustomobject]@{
            Path = $File.FullName
            Name = $File.Name
            FileHour = $Stamp
            Length = $File.Length
            LastWriteTime = $File.LastWriteTime
        })
    }

    return @($Results | Sort-Object FileHour,Name)
}

function Read-CseGraphAuditRange {
    param([Parameter(Mandatory)][object[]]$Inventory)

    if (@($Inventory).Count -eq 0) {
        throw 'No ERR.SESGRFM hourly error logs were found.'
    }

    $Oldest = ($Inventory | Sort-Object FileHour | Select-Object -First 1).FileHour
    $Newest = ($Inventory | Sort-Object FileHour | Select-Object -Last 1).FileHour

    Write-Host ''
    Write-Host 'Available Graph/CSE error-log coverage:' -ForegroundColor Cyan
    Write-Host ('  Oldest hourly file : {0}' -f $Oldest.ToString('MM/dd/yyyy HH:00'))
    Write-Host ('  Newest hourly file : {0}' -f $Newest.ToString('MM/dd/yyyy HH:00'))
    Write-Host ('  Hourly files       : {0}' -f @($Inventory).Count)
    Write-Host ''
    Write-Host 'Audit period:' -ForegroundColor Cyan
    Write-Host '  1. Current day'
    Write-Host '  2. Last X days'
    Write-Host '  3. All retained Graph/CSE error logs'
    Write-Host '  4. Custom date range'
    Write-Host ''

    $Choice = ''
    while ($Choice -notin @('1','2','3','4')) {
        $Choice = (Read-Host 'Select 1, 2, 3, or 4').Trim()
    }

    $Today = (Get-Date).Date

    if ($Choice -eq '1') {
        return [pscustomobject]@{
            Start = $Today
            End = $Today
            Description = 'current day'
        }
    }

    if ($Choice -eq '2') {
        $Days = 0
        while ($Days -le 0) {
            $Text = (Read-Host 'How many days? [Default: 1]').Trim()
            if ([string]::IsNullOrWhiteSpace($Text)) {
                $Days = 1
            }
            else {
                [void][int]::TryParse($Text,[ref]$Days)
            }

            if ($Days -le 0) {
                Write-Host 'Enter a positive number of days.' -ForegroundColor Yellow
            }
        }

        $SelectedStart = $Today.AddDays(-($Days - 1))
        Confirm-IxmLargeSearchRange -StartDate $SelectedStart -EndDate $Today -Context 'Graph/CSE error-log audit'

        return [pscustomobject]@{
            Start = $SelectedStart
            End = $Today
            Description = ('last {0} day(s)' -f $Days)
        }
    }

    if ($Choice -eq '3') {
        Confirm-IxmLargeSearchRange -StartDate $Oldest.Date -EndDate $Newest.Date -Context 'all retained Graph/CSE error logs'

        return [pscustomobject]@{
            Start = $Oldest.Date
            End = $Newest.Date
            Description = 'all retained Graph/CSE error logs'
        }
    }

    $StartDate = $null
    while ($null -eq $StartDate) {
        $Text = (Read-Host 'Start date (MM/DD/YYYY)').Trim()
        $Parsed = [datetime]::MinValue

        if ([datetime]::TryParse($Text,[ref]$Parsed)) {
            $StartDate = $Parsed.Date
        }
        else {
            Write-Host 'Invalid date.' -ForegroundColor Yellow
        }
    }

    $EndDate = $null
    while ($null -eq $EndDate) {
        $Text = (Read-Host 'End date (MM/DD/YYYY)').Trim()
        $Parsed = [datetime]::MinValue

        if ([datetime]::TryParse($Text,[ref]$Parsed)) {
            $EndDate = $Parsed.Date
        }
        else {
            Write-Host 'Invalid date.' -ForegroundColor Yellow
        }
    }

    if ($EndDate -lt $StartDate) {
        $Temp = $StartDate
        $StartDate = $EndDate
        $EndDate = $Temp
    }

    Confirm-IxmLargeSearchRange -StartDate $StartDate -EndDate $EndDate -Context 'Graph/CSE error-log audit'

    return [pscustomobject]@{
        Start = $StartDate
        End = $EndDate
        Description = ('{0} through {1}' -f $StartDate.ToString('MM/dd/yyyy'),$EndDate.ToString('MM/dd/yyyy'))
    }
}

function Get-CseEventTime {
    param(
        [Parameter(Mandatory)][datetime]$FileHour,
        [Parameter(Mandatory)][string]$Line
    )

    if ($Line -match '\[D:(?<Hour>\d{2})!(?<Minute>\d{2})\|(?<Second>\d{2})(?:\.(?<Milli>\d{1,3}))?\]') {
        $Hour = [int]$Matches.Hour
        $Minute = [int]$Matches.Minute
        $Second = [int]$Matches.Second
        $MilliText = [string]$Matches.Milli

        $Milli = 0
        if (-not [string]::IsNullOrWhiteSpace($MilliText)) {
            $MilliText = $MilliText.PadRight(3,'0')
            if ($MilliText.Length -gt 3) {
                $MilliText = $MilliText.Substring(0,3)
            }
            $Milli = [int]$MilliText
        }

        try {
            return $FileHour.Date.AddHours($Hour).AddMinutes($Minute).AddSeconds($Second).AddMilliseconds($Milli)
        }
        catch {
            return $FileHour
        }
    }

    return $FileHour
}

function Get-CseGraphErrorInfo {
    param(
        [Parameter(Mandatory)][string]$BlockText
    )

    $ErrorCode = 'GRAPH/CSE Error'
    $Operation = 'Graph/CSE operation'
    $Classification = 'GRAPH/CSE ERROR - REVIEW'
    $Recommendation = 'Review the CSE error and verify the configured Exchange/M365 mailbox and Graph access.'

    if ($BlockText -match '(?i)(Request_ResourceNotFound|ResourceNotFound|ItemNotFound|MailboxNotEnabled)') {
        $ErrorCode = 'MSG:ResourceNotFound'
    }
    elseif ($BlockText -match '(?i)ServiceException' -and $BlockText -match '(?i)timeout|timed\s*out') {
        $ErrorCode = 'MSG:ServiceException:timeout'
    }
    elseif ($BlockText -match '(?i)NullReferenceException') {
        $ErrorCode = 'SYS:NullReferenceException'
    }
    elseif ($BlockText -match '(?i)ServiceException') {
        $ErrorCode = 'MSG:ServiceException'
    }
    elseif ($BlockText -match '(?i)(?<Type>[A-Za-z0-9_.]+Exception)') {
        $ErrorCode = $Matches.Type
    }

    $HasFolderLookup = (
        $BlockText -match '(?i)UC\.CSE\.MSGR\.MailService\.<GetFolderId>' -or
        $BlockText -match '(?i)Microsoft\.Graph\.MailFolderRequest\.<GetAsync>'
    )

    if ($HasFolderLookup) {
        $Operation = 'Graph mailbox/folder lookup'
    }
    elseif ($BlockText -match '(?i)MailFolder') {
        $Operation = 'Graph mail-folder operation'
    }
    elseif ($BlockText -match '(?i)Microsoft\.Graph') {
        $Operation = 'Microsoft Graph request'
    }

    if ($BlockText -match '(?i)(Request_ResourceNotFound|ResourceNotFound|ItemNotFound|MailboxNotEnabled|mailbox\s+not\s+found|user\s+not\s+found)') {
        $Classification = 'GRAPH USER / MAILBOX NOT FOUND'
        $Recommendation = 'Verify the Exchange/M365 user and mailbox still exist and are mailbox-enabled. Remove or correct stale IX Messaging Graph configuration if the user has left the organization.'
    }
    elseif ($BlockText -match '(?i)(unauthorized|forbidden|authentication|authorization|insufficient privileges|access denied|InvalidAuthenticationToken)') {
        $Classification = 'GRAPH AUTH / PERMISSION FAILURE'
        $Recommendation = 'Verify the IX Messaging Graph application credentials, permissions, tenant configuration, and mailbox access.'
    }
    elseif ($BlockText -match '(?i)ServiceException' -and $BlockText -match '(?i)timeout|timed\s*out') {
        $Classification = 'GRAPH/API TIMEOUT'
        $Recommendation = 'Check Graph/API connectivity, service responsiveness, throttling, proxy/firewall conditions, and whether the problem is transient.'
    }
    elseif ($BlockText -match '(?i)NullReferenceException' -and $HasFolderLookup) {
        $Classification = 'GRAPH MAILBOX/FOLDER LOOKUP FAILED'
        $Recommendation = 'Verify the configured Exchange/M365 mailbox exists, is mailbox-enabled, and is accessible to the IX Messaging Graph application. If the user has left, remove or update the stale email/Graph configuration.'
    }
    elseif ($BlockText -match '(?i)NullReferenceException') {
        $Classification = 'GRAPH/CSE NULL REFERENCE - CHECK ACCOUNT'
        $Recommendation = 'Verify the configured Exchange/M365 mailbox and Graph access. The exception alone does not prove the email address is invalid.'
    }
    elseif ($BlockText -match '(?i)ServiceException') {
        $Classification = 'GRAPH SERVICE FAILURE'
        $Recommendation = 'Review the Graph service error and verify mailbox availability, Graph access, and service health.'
    }

    return [pscustomobject]@{
        ErrorCode = $ErrorCode
        Operation = $Operation
        Classification = $Classification
        Recommendation = $Recommendation
    }
}

function Get-CseGraphFailureEvents {
    param(
        [Parameter(Mandatory)][object[]]$Files
    )

    $Events = New-Object System.Collections.Generic.List[object]
    $SessionPattern = '\[S:(?<Server>\d+)-(?<Extension>\d+)-(?<Folder>[^\]]+)\]'
    $ErrorPattern = '(?i)(NullReferenceException|ServiceException|timeout|ResourceNotFound|ItemNotFound|MailboxNotEnabled|Request_ResourceNotFound|[A-Za-z0-9_.]+Exception)'

    $FileNumber = 0

    foreach ($File in @($Files | Sort-Object FileHour)) {
        $FileNumber++
        Write-Progress `
            -Activity 'Scanning Graph/CSE error logs' `
            -Status ('{0} ({1}/{2})' -f $File.Name,$FileNumber,@($Files).Count) `
            -PercentComplete (($FileNumber / [double]@($Files).Count) * 100)

        $Lines = @(Get-Content -LiteralPath $File.Path -ErrorAction SilentlyContinue)
        if ($Lines.Count -eq 0) {
            continue
        }

        $Starts = New-Object System.Collections.Generic.List[int]

        for ($i = 0; $i -lt $Lines.Count; $i++) {
            $Line = [string]$Lines[$i]

            if ($Line -match $SessionPattern -and $Line -match $ErrorPattern) {
                $Starts.Add($i)
            }
        }

        if ($Starts.Count -eq 0) {
            continue
        }

        for ($s = 0; $s -lt $Starts.Count; $s++) {
            $StartIndex = $Starts[$s]

            if ($s -lt ($Starts.Count - 1)) {
                $EndIndex = $Starts[$s + 1] - 1
            }
            else {
                $EndIndex = [math]::Min($Lines.Count - 1,$StartIndex + 80)
            }

            $FirstLine = [string]$Lines[$StartIndex]

            # Re-run the session match so $Matches contains this event's session.
            if ($FirstLine -notmatch $SessionPattern) {
                continue
            }

            $Extension = [string]$Matches.Extension
            $Folder = [string]$Matches.Folder

            if ($EndIndex -gt $StartIndex) {
                $BlockLines = @($Lines[$StartIndex..$EndIndex])
            }
            else {
                $BlockLines = @($Lines[$StartIndex])
            }

            $BlockText = ($BlockLines -join "`n")
            $Info = Get-CseGraphErrorInfo -BlockText $BlockText
            $EventTime = Get-CseEventTime -FileHour $File.FileHour -Line $FirstLine

            $Events.Add([pscustomobject]@{
                EventTime = $EventTime
                Extension = $Extension
                Folder = $Folder
                ErrorCode = $Info.ErrorCode
                Operation = $Info.Operation
                Classification = $Info.Classification
                Recommendation = $Info.Recommendation
                SourceFile = $File.Name
                SourcePath = $File.Path
            })
        }
    }

    Write-Progress -Activity 'Scanning Graph/CSE error logs' -Completed
    return @($Events | Sort-Object EventTime,Extension)
}

function Get-CseAuditCoverageWindow {
    param([Parameter(Mandatory)][object[]]$Files)

    if (@($Files).Count -eq 0) {
        return $null
    }

    $Start = ($Files | Sort-Object FileHour | Select-Object -First 1).FileHour
    $LastHour = ($Files | Sort-Object FileHour | Select-Object -Last 1).FileHour
    $End = $LastHour.AddHours(1)
    $Now = Get-Date

    if ($End -gt $Now -and $LastHour.Date -eq $Now.Date -and $LastHour.Hour -eq $Now.Hour) {
        $End = $Now
    }

    if ($End -le $Start) {
        $End = $Start.AddMinutes(1)
    }

    return [pscustomobject]@{
        Start = $Start
        End = $End
        Hours = [math]::Max(($End - $Start).TotalHours,(1.0 / 60.0))
    }
}

function Get-IxmGraphAuditDirectory {
    param([Parameter(Mandatory)]$Connection)

    $Tables = $Connection.GetSchema('Tables')
    $Columns = $Connection.GetSchema('Columns')

    $HasImapName = Test-IxmColumn -Columns $Columns -Table 'MAILBOX' -Column 'IMAPNAME'
    $HasUserName = Test-IxmColumn -Columns $Columns -Table 'MAILBOX' -Column 'USERNAME'
    $HasUseImap = Test-IxmColumn -Columns $Columns -Table 'MAILBOX' -Column 'MbxUseMbxIMAP'
    $HasMbxMwi = Test-IxmTable -Tables $Tables -Table 'MbxMWI'
    $HasLastInboxSync = Test-IxmColumn -Columns $Columns -Table 'MbxMWI' -Column 'LastSyncInboxDateTime'

    if ($HasImapName) {
        $EmailExpr = 'm.IMAPNAME AS EmailAddress'
    }
    elseif ($HasUserName) {
        $EmailExpr = 'm.USERNAME AS EmailAddress'
    }
    else {
        $EmailExpr = "CAST(NULL AS VARCHAR(1)) AS EmailAddress"
    }

    $UseImapExpr = if ($HasUseImap) {
        'm.MbxUseMbxIMAP AS GraphIMAPConfigured'
    }
    else {
        'CAST(NULL AS INTEGER) AS GraphIMAPConfigured'
    }

    if ($HasMbxMwi -and $HasLastInboxSync) {
        $JoinExpr = 'LEFT JOIN DBA.MbxMWI w ON m.MBXID = w.MBXID'
        $LastSyncExpr = 'w.LastSyncInboxDateTime AS LastSyncInboxDateTime'
    }
    else {
        $JoinExpr = ''
        $LastSyncExpr = "CAST(NULL AS VARCHAR(1)) AS LastSyncInboxDateTime"
    }

    $Sql = @"
SELECT
    m.MBXID,
    m.MBXNUMBER AS Extension,
    m.FIRSTNAME AS FirstName,
    m.LASTNAME AS LastName,
    f.FGNAME AS FeatureGroup,
    $EmailExpr,
    $UseImapExpr,
    $LastSyncExpr
FROM DBA.MAILBOX m
LEFT JOIN DBA.FGROUP f
    ON m.FGROUPID = f.FGROUPID
$JoinExpr
WHERE m.MBXNUMBER IS NOT NULL
  AND m.MBXNUMBER <> ''
ORDER BY m.MBXNUMBER
"@

    $Result = Invoke-IxmReadOnlyQuery -Connection $Connection -Sql $Sql
    $Rows = @(Get-IxmResultRows -Result $Result)

    $Output = New-Object System.Collections.Generic.List[object]

    foreach ($Row in $Rows) {
        $First = if ($Row.IsNull('FirstName')) { '' } else { [string]$Row.FirstName }
        $Last = if ($Row.IsNull('LastName')) { '' } else { [string]$Row.LastName }
        $Name = (($First.Trim() + ' ' + $Last.Trim()).Trim())

        $GraphText = 'Not available'
        if (-not $Row.IsNull('GraphIMAPConfigured')) {
            if ([System.Convert]::ToBoolean($Row.GraphIMAPConfigured)) {
                $GraphText = 'Yes'
            }
            else {
                $GraphText = 'No'
            }
        }

        $LastSync = ''
        $LastSyncDateTime = $null

        if (-not $Row.IsNull('LastSyncInboxDateTime')) {
            $LastSyncValue = $Row.LastSyncInboxDateTime
            $LastSync = Convert-IxmDateTimeText -Value $LastSyncValue
            $LastSyncDateTime = Convert-GraphAuditDateTime -Value $LastSyncValue
        }

        $Output.Add([pscustomobject]@{
            MailboxID = if ($Row.IsNull('MBXID')) { $null } else { [int]$Row.MBXID }
            Extension = if ($Row.IsNull('Extension')) { '' } else { [string]$Row.Extension }
            Name = $Name
            FeatureGroup = if ($Row.IsNull('FeatureGroup')) { '' } else { [string]$Row.FeatureGroup }
            EmailAddress = if ($Row.IsNull('EmailAddress')) { '' } else { ([string]$Row.EmailAddress).Trim() }
            GraphIMAPConfigured = $GraphText
            LastInboxSync = $LastSync
            LastInboxSyncDateTime = $LastSyncDateTime
        })
    }

    return $Output.ToArray()
}

function Convert-GraphAuditDateTime {
    param($Value)

    if ($null -eq $Value -or $Value -is [System.DBNull]) {
        return $null
    }

    if ($Value -is [datetime]) {
        return [datetime]$Value
    }

    $Text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $Formats = @(
        'yyyyMMddHHmmss',
        'yyyyMMddHHmm',
        'MM/dd/yyyy hh:mm:ss tt',
        'MM/dd/yyyy h:mm:ss tt',
        'MM/dd/yyyy HH:mm:ss'
    )

    foreach ($Format in $Formats) {
        $Parsed = [datetime]::MinValue

        if ([datetime]::TryParseExact(
            $Text,
            $Format,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None,
            [ref]$Parsed
        )) {
            return $Parsed
        }
    }

    $General = [datetime]::MinValue
    if ([datetime]::TryParse($Text,[ref]$General)) {
        return $General
    }

    return $null
}

function Get-GraphFailureAssessment {
    param(
        [Parameter(Mandatory)]$Failure,
        $Mailbox,
        [Parameter(Mandatory)][int]$Occurrences,
        [Parameter(Mandatory)][double]$RatePerHour,
        [Parameter(Mandatory)][datetime]$FirstSeen,
        [Parameter(Mandatory)][datetime]$LastSeen,
        [bool]$DatabaseAvailable = $true
    )

    $Priority = 'REVIEW'
    $Assessment = 'REVIEW GRAPH / EXCHANGE ACCOUNT'
    $Rationale = ''
    $Recommendation = $Failure.Recommendation

    # "Persistent" is intended to represent the high-rate polling pattern seen
    # in the Avaya CSE logs, not merely several failures spread over many days.
    $IsPersistent = (
        ($Occurrences -ge 5 -and $RatePerHour -ge 0.50) -or
        $Occurrences -ge 10
    )

    $IsIsolated = ($Occurrences -le 2)

    if (-not $DatabaseAvailable) {
        return [pscustomobject]@{
            Priority = 'INFO'
            Assessment = 'DATABASE CORRELATION UNAVAILABLE'
            Rationale = 'CSE failures were found, but mailbox/email configuration could not be read from IX Messaging.'
            Recommendation = 'Resolve database correlation before deciding whether the Exchange/M365 account itself needs attention.'
        }
    }

    if ($null -eq $Mailbox) {
        return [pscustomobject]@{
            Priority = 'INVESTIGATE'
            Assessment = 'STALE CSE SESSION / NOT IN IXM DATABASE'
            Rationale = 'The extension appears in CSE Graph error logs but was not found in the current IX Messaging MAILBOX table.'
            Recommendation = 'Check for stale CSE/Graph state and confirm that the mailbox was intentionally removed from IX Messaging.'
        }
    }

    if ($Mailbox.GraphIMAPConfigured -eq 'No') {
        return [pscustomobject]@{
            Priority = 'INVESTIGATE'
            Assessment = 'GRAPH DISABLED IN IXM / CHECK STALE ACTIVITY'
            Rationale = 'CSE Graph activity exists even though the mailbox is currently marked as not using Graph/IMAP.'
            Recommendation = 'Check for stale CSE synchronization state or a recent configuration change.'
        }
    }

    $LastSync = $null
    if ($null -ne $Mailbox.LastInboxSyncDateTime) {
        $LastSync = $Mailbox.LastInboxSyncDateTime
    }

    $HasSuccessfulSync = ($null -ne $LastSync)
    $SyncAfterLastFailure = ($HasSuccessfulSync -and $LastSync -ge $LastSeen)
    $SyncDuringWindow = ($HasSuccessfulSync -and $LastSync -ge $FirstSeen -and $LastSync -le $LastSeen)

    # Explicit Graph "not found" style responses are stronger evidence than a
    # generic NullReferenceException and should remain actionable even if there
    # was an older successful synchronization.
    if ($Failure.Classification -eq 'GRAPH USER / MAILBOX NOT FOUND') {
        return [pscustomobject]@{
            Priority = 'VERIFY'
            Assessment = 'VERIFY / REMOVE INVALID EXCHANGE MAILBOX'
            Rationale = 'The CSE error explicitly indicates that the Graph user/mailbox/resource was not found.'
            Recommendation = 'Verify that the Exchange/M365 user and mailbox still exist and are mailbox-enabled. Remove or correct stale IX Messaging Graph configuration if appropriate.'
        }
    }

    if ($Failure.Classification -eq 'GRAPH AUTH / PERMISSION FAILURE') {
        return [pscustomobject]@{
            Priority = 'INVESTIGATE'
            Assessment = 'CHECK GRAPH APP PERMISSIONS'
            Rationale = 'The Graph request failed with an authentication, authorization, or permission-related error.'
            Recommendation = 'Verify IX Messaging Graph credentials, tenant configuration, permissions, and mailbox access.'
        }
    }

    if ($Failure.Classification -eq 'GRAPH/API TIMEOUT') {
        if ($IsPersistent) {
            return [pscustomobject]@{
                Priority = 'INVESTIGATE'
                Assessment = 'REPEATED GRAPH/API TIMEOUT - CHECK CONNECTIVITY'
                Rationale = ('{0} timeout failure(s) were seen at approximately {1:N2}/hour in the selected window.' -f $Occurrences,$RatePerHour)
                Recommendation = 'Check Graph/API connectivity, service responsiveness, throttling, proxy/firewall conditions, and Microsoft 365 service health.'
            }
        }

        return [pscustomobject]@{
            Priority = 'MONITOR'
            Assessment = 'ISOLATED GRAPH/API TIMEOUT - MONITOR'
            Rationale = ('Only {0} timeout failure(s) were seen in the selected window.' -f $Occurrences)
            Recommendation = 'Monitor for recurrence. A timeout by itself does not indicate that the configured Exchange mailbox is invalid.'
        }
    }

    $IsMailboxLookupFailure = (
        $Failure.Classification -eq 'GRAPH MAILBOX/FOLDER LOOKUP FAILED' -or
        $Failure.Classification -eq 'GRAPH/CSE NULL REFERENCE - CHECK ACCOUNT'
    )

    if ($IsMailboxLookupFailure) {
        if ($IsPersistent) {
            if (-not $HasSuccessfulSync) {
                return [pscustomobject]@{
                    Priority = 'VERIFY'
                    Assessment = 'PERSISTENT GRAPH MAILBOX FAILURE - VERIFY EXCHANGE/M365'
                    Rationale = ('{0} mailbox/folder lookup failures were seen at approximately {1:N2}/hour, and no successful Inbox sync is recorded.' -f $Occurrences,$RatePerHour)
                    Recommendation = 'Verify that the configured Exchange/M365 mailbox exists, is mailbox-enabled, and is accessible to the IX Messaging Graph application. If the user has left, remove or update the stale Graph/email configuration.'
                }
            }

            if ($SyncAfterLastFailure) {
                return [pscustomobject]@{
                    Priority = 'MONITOR'
                    Assessment = 'REPEATED GRAPH FAILURE - SUCCESSFUL SYNC RECORDED AFTERWARD'
                    Rationale = ('{0} failures occurred, but LastInboxSync ({1}) is at or after the last observed failure.' -f $Occurrences,$LastSync.ToString('MM/dd/yyyy hh:mm:ss tt'))
                    Recommendation = 'The mailbox appears to have synchronized after the failure sequence. Monitor for recurrence before treating the account as invalid.'
                }
            }

            if ($SyncDuringWindow) {
                return [pscustomobject]@{
                    Priority = 'INVESTIGATE'
                    Assessment = 'REPEATED GRAPH FAILURE WITH INTERMITTENT SUCCESS'
                    Rationale = ('{0} failures occurred, but a successful Inbox sync was also recorded during the failure window.' -f $Occurrences)
                    Recommendation = 'Investigate intermittent Graph/mailbox access, service throttling, permissions, or mailbox availability rather than assuming the email account is invalid.'
                }
            }

            return [pscustomobject]@{
                Priority = 'INVESTIGATE'
                Assessment = 'REPEATED GRAPH SYNC FAILURE - INVESTIGATE ACCOUNT / ACCESS'
                Rationale = ('{0} mailbox/folder lookup failures were seen; the last successful Inbox sync ({1}) predates the current failure sequence.' -f $Occurrences,$LastSync.ToString('MM/dd/yyyy hh:mm:ss tt'))
                Recommendation = 'Verify the Exchange/M365 mailbox and Graph access. The prior successful sync proves the account worked previously, but repeated current failures warrant investigation.'
            }
        }

        # This is the important distinction for users such as extension 2631:
        # one Graph folder lookup failure plus a previous successful Inbox sync
        # is evidence of an isolated failure, not evidence that the email address
        # or Exchange mailbox is invalid.
        if ($HasSuccessfulSync) {
            return [pscustomobject]@{
                Priority = 'MONITOR'
                Assessment = 'ISOLATED / TRANSIENT GRAPH FAILURE - MONITOR'
                Rationale = ('Only {0} mailbox/folder lookup failure(s) were seen, and a previous successful Inbox sync is recorded at {1}.' -f $Occurrences,$LastSync.ToString('MM/dd/yyyy hh:mm:ss tt'))
                Recommendation = 'Monitor for recurrence. Do not treat the configured email address or Exchange mailbox as invalid based on this isolated failure.'
            }
        }

        return [pscustomobject]@{
            Priority = 'REVIEW'
            Assessment = 'ISOLATED GRAPH FAILURE / NO SYNC RECORDED'
            Rationale = ('Only {0} mailbox/folder lookup failure(s) were seen, but no successful Inbox sync is recorded.' -f $Occurrences)
            Recommendation = 'Verify the mailbox if the user is reporting a problem or if failures recur. One isolated NullReferenceException is not enough to prove that the email account is invalid.'
        }
    }

    if ($IsPersistent) {
        $Priority = 'INVESTIGATE'
        $Assessment = 'REPEATED GRAPH/CSE FAILURE - INVESTIGATE'
        $Rationale = ('{0} failure(s) were seen at approximately {1:N2}/hour.' -f $Occurrences,$RatePerHour)
    }
    elseif ($IsIsolated) {
        $Priority = 'MONITOR'
        $Assessment = 'ISOLATED GRAPH/CSE FAILURE - MONITOR'
        $Rationale = ('Only {0} failure(s) were seen in the selected window.' -f $Occurrences)
    }
    else {
        $Priority = 'REVIEW'
        $Assessment = 'GRAPH/CSE FAILURE - REVIEW'
        $Rationale = ('{0} failure(s) were seen in the selected window.' -f $Occurrences)
    }

    return [pscustomobject]@{
        Priority = $Priority
        Assessment = $Assessment
        Rationale = $Rationale
        Recommendation = $Recommendation
    }
}

function Merge-CseGraphFailuresWithIxm {
    param(
        [Parameter(Mandatory)][object[]]$Events,
        [Parameter(Mandatory)][double]$CoverageHours,
        [AllowEmptyCollection()][object[]]$Directory = @(),
        [bool]$DatabaseAvailable = $true
    )

    $MailboxMap = @{}

    foreach ($Mailbox in @($Directory)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$Mailbox.Extension)) {
            $MailboxMap[[string]$Mailbox.Extension] = $Mailbox
        }
    }

    $Groups = @(
        $Events |
        Group-Object -Property Extension,ErrorCode,Classification,Operation,Folder
    )

    $Rows = New-Object System.Collections.Generic.List[object]

    foreach ($Group in $Groups) {
        $Items = @($Group.Group | Sort-Object EventTime)
        if ($Items.Count -eq 0) { continue }

        $FirstItem = $Items[0]
        $LastItem = $Items[$Items.Count - 1]
        $Extension = [string]$FirstItem.Extension

        $Mailbox = $null
        if ($MailboxMap.ContainsKey($Extension)) {
            $Mailbox = $MailboxMap[$Extension]
        }

        $Count = $Items.Count
        $Rate = 0.0

        if ($CoverageHours -gt 0) {
            $Rate = $Count / $CoverageHours
        }

        if ($DatabaseAvailable) {
            $Name = '<not in IXM database>'
        }
        else {
            $Name = '<database unavailable>'
        }

        $MailboxID = $null
        $FeatureGroup = ''
        $EmailAddress = ''
        $GraphConfigured = 'Unknown'
        $LastInboxSync = ''

        if ($null -ne $Mailbox) {
            $Name = $Mailbox.Name
            $MailboxID = $Mailbox.MailboxID
            $FeatureGroup = $Mailbox.FeatureGroup
            $EmailAddress = $Mailbox.EmailAddress
            $GraphConfigured = $Mailbox.GraphIMAPConfigured
            $LastInboxSync = $Mailbox.LastInboxSync
        }

        $AssessmentInfo = Get-GraphFailureAssessment `
            -Failure $FirstItem `
            -Mailbox $Mailbox `
            -Occurrences $Count `
            -RatePerHour $Rate `
            -FirstSeen $FirstItem.EventTime `
            -LastSeen $LastItem.EventTime `
            -DatabaseAvailable $DatabaseAvailable

        $Rows.Add([pscustomobject]@{
            Extension = $Extension
            Name = $Name
            MailboxID = $MailboxID
            EmailAddress = $EmailAddress
            FeatureGroup = $FeatureGroup
            GraphConfigured = $GraphConfigured
            LastInboxSync = if ([string]::IsNullOrWhiteSpace([string]$LastInboxSync)) { 'Blank / not recorded' } else { $LastInboxSync }
            LastInboxSyncDateTime = if ($null -ne $Mailbox) { $Mailbox.LastInboxSyncDateTime } else { $null }
            Folder = $FirstItem.Folder
            ErrorCode = $FirstItem.ErrorCode
            Operation = $FirstItem.Operation
            Classification = $FirstItem.Classification
            Occurrences = $Count
            RatePerHour = [math]::Round($Rate,2)
            FirstSeen = $FirstItem.EventTime
            LastSeen = $LastItem.EventTime
            Priority = $AssessmentInfo.Priority
            Assessment = $AssessmentInfo.Assessment
            Rationale = $AssessmentInfo.Rationale
            Recommendation = $AssessmentInfo.Recommendation
            LastSourceFile = $LastItem.SourceFile
        })
    }

    return @(
        $Rows |
        Sort-Object `
            @{Expression={
                switch ($_.Priority) {
                    'VERIFY'      { 0 }
                    'INVESTIGATE' { 1 }
                    'REVIEW'      { 2 }
                    'MONITOR'     { 3 }
                    default       { 4 }
                }
            }}, `
            @{Expression='Occurrences';Descending=$true}, `
            Extension
    )
}

function Show-GraphFailureAudit {
    param(
        [Parameter(Mandatory)][object[]]$Rows,
        [Parameter(Mandatory)][object[]]$Files,
        [Parameter(Mandatory)]$Coverage,
        [Parameter(Mandatory)][string]$CseRoot,
        [Parameter(Mandatory)][string]$PeriodDescription,
        [bool]$DatabaseAvailable
    )

    Write-Host ''
    Write-Host 'AUDIT SUMMARY' -ForegroundColor Cyan
    Write-Host ('-' * 78)
    Write-Host ('CSE directory           : {0}' -f $CseRoot)
    Write-Host ('Audit period            : {0}' -f $PeriodDescription)
    Write-Host ('Hourly error files      : {0}' -f @($Files).Count)
    Write-Host ('Scanned log window      : {0} through {1}' -f $Coverage.Start.ToString('MM/dd/yyyy HH:mm:ss'),$Coverage.End.ToString('MM/dd/yyyy HH:mm:ss'))
    Write-Host ('Scanned hours           : {0:N2}' -f $Coverage.Hours)
    Write-Host ('Database correlation    : {0}' -f $(if ($DatabaseAvailable) { 'Available' } else { 'Unavailable - log-only results' }))

    $TotalEvents = 0
    foreach ($Row in @($Rows)) {
        $TotalEvents += [int]$Row.Occurrences
    }

    $AffectedExtensions = @($Rows | Select-Object -ExpandProperty Extension -Unique).Count

    Write-Host ('Failure events          : {0}' -f $TotalEvents)
    Write-Host ('Affected extensions     : {0}' -f $AffectedExtensions)

    if (@($Rows).Count -eq 0) {
        Write-Host ''
        Write-Host 'No Graph/CSE mailbox failure events were found in the selected ERR.SESGRFM logs.' -ForegroundColor Green
        return
    }

    Write-Host ''
    Write-Host 'FAILURE TYPES' -ForegroundColor Cyan
    Write-Host ('-' * 78)

    @(
        $Rows |
        Group-Object Classification |
        ForEach-Object {
            $EventCount = 0
            foreach ($Item in $_.Group) {
                $EventCount += [int]$Item.Occurrences
            }

            [pscustomobject]@{
                Classification = $_.Name
                Extensions = @($_.Group | Select-Object -ExpandProperty Extension -Unique).Count
                Events = $EventCount
            }
        } |
        Sort-Object Events -Descending
    ) | Format-Table -AutoSize -Wrap | Out-Host

    Write-Host ''
    Write-Host 'ASSESSMENT SUMMARY' -ForegroundColor Cyan
    Write-Host ('-' * 78)

    @(
        $Rows |
        Group-Object Priority |
        ForEach-Object {
            [pscustomobject]@{
                Priority = $_.Name
                Findings = $_.Count
                Extensions = @($_.Group | Select-Object -ExpandProperty Extension -Unique).Count
            }
        } |
        Sort-Object @{
            Expression = {
                switch ($_.Priority) {
                    'VERIFY'      { 0 }
                    'INVESTIGATE' { 1 }
                    'REVIEW'      { 2 }
                    'MONITOR'     { 3 }
                    default       { 4 }
                }
            }
        }
    ) | Format-Table -AutoSize | Out-Host

    $ActionRows = @(
        $Rows |
        Where-Object { $_.Priority -in @('VERIFY','INVESTIGATE','REVIEW') }
    )

    $MonitorRows = @(
        $Rows |
        Where-Object { $_.Priority -eq 'MONITOR' }
    )

    if ($ActionRows.Count -gt 0) {
        Write-Host ''
        Write-Host 'ACTION / INVESTIGATION' -ForegroundColor Cyan
        Write-Host ('-' * 78)

        $ActionRows |
            Select-Object `
                Priority,
                Extension,
                Name,
                EmailAddress,
                GraphConfigured,
                LastInboxSync,
                ErrorCode,
                Occurrences,
                RatePerHour,
                @{Name='LastSeen';Expression={$_.LastSeen.ToString('MM/dd/yyyy HH:mm:ss')}},
                Assessment |
            Format-Table -AutoSize -Wrap |
            Out-Host
    }

    if ($MonitorRows.Count -gt 0) {
        Write-Host ''
        Write-Host 'ISOLATED / TRANSIENT - MONITOR' -ForegroundColor Cyan
        Write-Host ('-' * 78)

        $MonitorRows |
            Select-Object `
                Extension,
                Name,
                EmailAddress,
                LastInboxSync,
                ErrorCode,
                Occurrences,
                @{Name='LastSeen';Expression={$_.LastSeen.ToString('MM/dd/yyyy HH:mm:ss')}},
                Assessment |
            Format-Table -AutoSize -Wrap |
            Out-Host
    }

    Write-Host ''
    Write-Host 'Interpretation:' -ForegroundColor Yellow
    Write-Host '  A single NullReferenceException does not mean the configured email address is invalid.' -ForegroundColor Yellow
    Write-Host '  The audit now considers occurrence count, approximate failure rate, and LastInboxSync.' -ForegroundColor Yellow
    Write-Host '  One or two mailbox/folder lookup failures with a previous successful Inbox sync are' -ForegroundColor Yellow
    Write-Host '  classified as isolated/transient and placed in the MONITOR section.' -ForegroundColor Yellow
    Write-Host '  High-rate repeated mailbox/folder failures with no successful Inbox sync are treated' -ForegroundColor Yellow
    Write-Host '  as persistent and should be verified with the Exchange/M365 administrator.' -ForegroundColor Yellow
    Write-Host '  Explicit Graph mailbox/user-not-found errors remain actionable regardless of history.' -ForegroundColor Yellow
    Write-Host '  ServiceException timeouts are evaluated separately from mailbox lookup failures.' -ForegroundColor Yellow
    Write-Host '  Rate/hour is an approximate rate over the scanned hourly-log window.' -ForegroundColor Yellow
}

function Convert-GraphFailureRowsForExport {
    param([Parameter(Mandatory)][object[]]$Rows)

    return @(
        $Rows | ForEach-Object {
            [pscustomobject]@{
                Extension = $_.Extension
                Name = $_.Name
                MailboxID = $_.MailboxID
                EmailAddress = $_.EmailAddress
                FeatureGroup = $_.FeatureGroup
                GraphConfigured = $_.GraphConfigured
                LastInboxSync = $_.LastInboxSync
                Folder = $_.Folder
                ErrorCode = $_.ErrorCode
                Operation = $_.Operation
                Classification = $_.Classification
                Occurrences = $_.Occurrences
                RatePerHour = $_.RatePerHour
                FirstSeen = $_.FirstSeen.ToString('MM/dd/yyyy HH:mm:ss')
                LastSeen = $_.LastSeen.ToString('MM/dd/yyyy HH:mm:ss')
                Priority = $_.Priority
                Assessment = $_.Assessment
                Rationale = $_.Rationale
                Recommendation = $_.Recommendation
                LastSourceFile = $_.LastSourceFile
            }
        }
    )
}

function Invoke-GraphEmailFailureAudit {
    Write-Section 'Graph / Exchange Mailbox Failure Audit'

    Write-Host 'This audit reads CSE Graph error logs and correlates them with the IX Messaging database.' -ForegroundColor Cyan
    Write-Host 'This tool submits SELECT-only database queries. Effective database permissions are controlled by the configured SQL Anywhere DSN/account.' -ForegroundColor Green
    Write-Host 'No Exchange/M365 or Graph changes are made.' -ForegroundColor Green

    $CseRoot = Resolve-CseGraphRoot
    $Inventory = @(Get-CseGraphErrorFileInventory -CseRoot $CseRoot)

    if ($Inventory.Count -eq 0) {
        Write-Host ''
        Write-Host ('No ERR.SESGRFM.YYYY-MM-DDTHH.csv files were found under {0}.' -f $CseRoot) -ForegroundColor Yellow
        return
    }

    $Range = Read-CseGraphAuditRange -Inventory $Inventory

    $Files = @(
        $Inventory |
        Where-Object {
            $_.FileHour.Date -ge $Range.Start.Date -and
            $_.FileHour.Date -le $Range.End.Date
        } |
        Sort-Object FileHour
    )

    if ($Files.Count -eq 0) {
        Write-Host ''
        Write-Host 'No ERR.SESGRFM hourly files exist in the selected date range.' -ForegroundColor Yellow
        return
    }

    $Coverage = Get-CseAuditCoverageWindow -Files $Files

    Write-Host ''
    Write-Host ('Scanning {0} hourly Graph/CSE error file(s)...' -f $Files.Count) -ForegroundColor DarkGray
    $Events = @(Get-CseGraphFailureEvents -Files $Files)

    if ($Events.Count -eq 0) {
        Write-Host ''
        Write-Host 'No Graph/CSE mailbox failure events were found in the selected files.' -ForegroundColor Green
        return
    }

    Write-Host ('Found {0} Graph/CSE failure event(s).' -f $Events.Count) -ForegroundColor Yellow
    Write-Host ''
    Write-Host 'Correlating failing extensions with IX Messaging mailbox configuration...' -ForegroundColor DarkGray

    $Directory = @()
    $DatabaseAvailable = $false
    $Connection = $null

    try {
        try {
            $Selected = Select-IxmDatabaseDsn

            Write-Host ('Using DSN: {0}' -f $Selected.Name) -ForegroundColor Cyan
            $Connection = Open-IxmDsnConnection -Name $Selected.Name

            if (-not (Test-IxmMailboxSchema -Connection $Connection)) {
                throw "DSN '$($Selected.Name)' connected, but the expected IX Messaging MAILBOX/FGROUP schema was not found."
            }

            $Directory = @(Get-IxmGraphAuditDirectory -Connection $Connection)
            $DatabaseAvailable = $true
        }
        catch {
            Write-Host ''
            Write-Host ('WARNING: IX Messaging database correlation is unavailable: {0}' -f $_.Exception.Message) -ForegroundColor Yellow
            Write-Host 'The audit will continue with extensions and CSE error data only.' -ForegroundColor Yellow
            $Directory = @()
            $DatabaseAvailable = $false
        }

        $Rows = @(
            Merge-CseGraphFailuresWithIxm `
                -Events $Events `
                -CoverageHours $Coverage.Hours `
                -Directory $Directory `
                -DatabaseAvailable $DatabaseAvailable
        )

        Show-GraphFailureAudit `
            -Rows $Rows `
            -Files $Files `
            -Coverage $Coverage `
            -CseRoot $CseRoot `
            -PeriodDescription $Range.Description `
            -DatabaseAvailable $DatabaseAvailable

        if ($Rows.Count -gt 0) {
            $ExportRows = @(Convert-GraphFailureRowsForExport -Rows $Rows)
            Export-ResultSet -Data $ExportRows -BaseName 'graph_exchange_failure_audit'
        }
    }
    finally {
        if ($null -ne $Connection) {
            try { $Connection.Close() } catch { Write-Verbose ('ODBC connection Close() cleanup failed: {0}' -f $_.Exception.Message) }
            try { $Connection.Dispose() } catch { Write-Verbose ('ODBC connection Dispose() cleanup failed: {0}' -f $_.Exception.Message) }
        }
    }
}


# -----------------------------------------------------------------------------
# IX Messaging system health check
# -----------------------------------------------------------------------------

function New-IxmHealthFinding {
    param(
        [Parameter(Mandatory)][string]$Section,
        [Parameter(Mandatory)][string]$Check,
        [AllowEmptyString()][string]$Value = '',
        [ValidateSet('OK','WARNING','ATTENTION','INFO','NOT APPLICABLE')]
        [string]$Status = 'INFO',
        [AllowEmptyString()][string]$Details = ''
    )

    return [pscustomobject]@{
        Section = $Section
        Check = $Check
        Value = $Value
        Status = $Status
        Details = $Details
    }
}

function Add-IxmHealthFinding {
    param(
        [Parameter(Mandatory)]$List,
        [Parameter(Mandatory)][string]$Section,
        [Parameter(Mandatory)][string]$Check,
        [AllowEmptyString()][string]$Value = '',
        [ValidateSet('OK','WARNING','ATTENTION','INFO','NOT APPLICABLE')]
        [string]$Status = 'INFO',
        [AllowEmptyString()][string]$Details = ''
    )

    $List.Add((New-IxmHealthFinding -Section $Section -Check $Check -Value $Value -Status $Status -Details $Details))
}

function Resolve-IxmUcRoot {
    # First preference: derive it from the already-selected VServer log root.
    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$LogRoot)) {
            $VServerParent = Split-Path -Path $LogRoot -Parent
            $Candidate = Split-Path -Path $VServerParent -Parent

            if ($Candidate -and (Test-Path -LiteralPath $Candidate)) {
                return $Candidate
            }
        }
    }
    catch {
        Write-Verbose ('Unable to derive the UC root from the selected VServer path: {0}' -f $_.Exception.Message)
    }

    # Second preference: UC Services uninstall metadata.
    $RegistryPaths = @(
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    foreach ($RegistryPath in $RegistryPaths) {
        try {
            foreach ($App in @(Get-ItemProperty $RegistryPath -ErrorAction SilentlyContinue)) {
                if ([string]$App.DisplayName -eq 'UC Services' -and
                    -not [string]::IsNullOrWhiteSpace([string]$App.InstallLocation)) {

                    $Install = ([string]$App.InstallLocation).TrimEnd('\')
                    if (Test-Path -LiteralPath $Install) {
                        return $Install
                    }
                }
            }
        }
        catch {
            Write-Verbose ('Unable to read UC Services uninstall metadata from {0}: {1}' -f $RegistryPath,$_.Exception.Message)
        }
    }

    # Third preference: derive the UC root from installed Avaya/SQL Anywhere
    # service executable paths. This helps when the VServer log root or uninstall
    # metadata is unavailable or the UC drive letter differs from the default.
    try {
        $ServiceRoots = New-Object System.Collections.Generic.List[string]
        foreach ($Svc in @(Get-CimInstance Win32_Service -ErrorAction SilentlyContinue)) {
            $Display = [string]$Svc.DisplayName
            $Name = [string]$Svc.Name
            $Path = [string]$Svc.PathName

            if ([string]::IsNullOrWhiteSpace($Path)) { continue }
            if ($Display -notmatch '(?i)^(UC|SQL Anywhere|MobiLink|DBWatcher)' -and
                $Name -notmatch '(?i)^(UC|SQL Anywhere|MobiLink|DBWatcher)') {
                continue
            }

            if ($Path -match '(?i)(?<Root>[A-Z]:\\(?:[^\\"]+\\)*UC)(?:\\|")') {
                $Candidate = $Matches.Root
                if ((Test-Path -LiteralPath $Candidate) -and -not $ServiceRoots.Contains($Candidate)) {
                    $ServiceRoots.Add($Candidate)
                }
            }
        }

        if ($ServiceRoots.Count -eq 1) {
            return $ServiceRoots[0]
        }
    }
    catch {
        Write-Verbose ('Unable to derive the UC root from Windows service paths: {0}' -f $_.Exception.Message)
    }

    # Last automatic attempt: local fixed drives.
    $Found = New-Object System.Collections.Generic.List[string]

    try {
        $Drives = @(
            Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop |
            Select-Object -ExpandProperty DeviceID
        )
    }
    catch {
        $Drives = @(
            Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Root.TrimEnd('\') }
        )
    }

    foreach ($Drive in @($Drives | Sort-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace([string]$Drive)) { continue }

        $Candidate = Join-Path ($Drive.TrimEnd('\') + '\') 'UC'
        if (Test-Path -LiteralPath $Candidate) {
            if (-not $Found.Contains($Candidate)) {
                $Found.Add($Candidate)
            }
        }
    }

    if ($Found.Count -eq 1) {
        return $Found[0]
    }

    if ($Found.Count -gt 1) {
        Write-Host ''
        Write-Host 'Multiple IX Messaging UC directories were detected:' -ForegroundColor Yellow
        for ($i = 0; $i -lt $Found.Count; $i++) {
            Write-Host ('  {0}. {1}' -f ($i + 1),$Found[$i])
        }

        do {
            $Selection = (Read-Host ('Select 1-{0}' -f $Found.Count)).Trim()
            $Number = 0
        } until ([int]::TryParse($Selection,[ref]$Number) -and $Number -ge 1 -and $Number -le $Found.Count)

        return $Found[$Number - 1]
    }

    return $null
}

function Get-IxmInstalledUcInfo {
    $RegistryPaths = @(
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    foreach ($RegistryPath in $RegistryPaths) {
        try {
            foreach ($App in @(Get-ItemProperty $RegistryPath -ErrorAction SilentlyContinue)) {
                if ([string]$App.DisplayName -eq 'UC Services') {
                    return [pscustomobject]@{
                        DisplayName = [string]$App.DisplayName
                        Version = [string]$App.DisplayVersion
                        InstallLocation = [string]$App.InstallLocation
                    }
                }
            }
        }
        catch {
            Write-Verbose ('Unable to read UC Services installation metadata from {0}: {1}' -f $RegistryPath,$_.Exception.Message)
        }
    }

    return $null
}

function Convert-IxmHealthDate {
    param($Value)

    if ($null -eq $Value -or $Value -is [System.DBNull]) {
        return $null
    }

    if ($Value -is [datetime]) {
        return [datetime]$Value
    }

    $Text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    foreach ($Format in @(
        'yyyyMMdd',
        'yyyyMMddHHmmss',
        'yyyy-MM-dd HH:mm:ss.fff',
        'yyyy-MM-dd HH:mm:ss',
        'MM/dd/yyyy HH:mm:ss',
        'MM/dd/yyyy hh:mm:ss tt'
    )) {
        $Parsed = [datetime]::MinValue

        if ([datetime]::TryParseExact(
            $Text,
            $Format,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None,
            [ref]$Parsed
        )) {
            return $Parsed
        }
    }

    $General = [datetime]::MinValue
    if ([datetime]::TryParse($Text,[ref]$General)) {
        return $General
    }

    return $null
}

function Get-IxmHealthStatusColor {
    param([Parameter(Mandatory)][string]$Status)

    switch ($Status) {
        'OK'             { return 'Green' }
        'WARNING'        { return 'Yellow' }
        'ATTENTION'      { return 'Red' }
        'INFO'           { return 'Gray' }
        'NOT APPLICABLE' { return 'DarkGray' }
        default          { return 'Gray' }
    }
}

function Show-IxmHealthFindings {
    param([Parameter(Mandatory)][object[]]$Findings)

    if (@($Findings).Count -eq 0) {
        Write-Host 'No health-check findings were generated.' -ForegroundColor Yellow
        return
    }

    $Summary = @(
        $Findings |
        Group-Object Status |
        ForEach-Object {
            [pscustomobject]@{
                Status = $_.Name
                Count = $_.Count
            }
        }
    )

    Write-Host ''
    Write-Host 'HEALTH SUMMARY' -ForegroundColor Cyan
    Write-Host ('-' * 78)

    foreach ($Status in @('ATTENTION','WARNING','OK','INFO','NOT APPLICABLE')) {
        $Item = $Summary | Where-Object { $_.Status -eq $Status } | Select-Object -First 1
        $Count = if ($null -eq $Item) { 0 } else { [int]$Item.Count }
        Write-Host ('{0,-16}: {1}' -f $Status,$Count) -ForegroundColor (Get-IxmHealthStatusColor -Status $Status)
    }

    $Sections = @($Findings | Select-Object -ExpandProperty Section -Unique)

    foreach ($Section in $Sections) {
        Write-Host ''
        Write-Host $Section.ToUpperInvariant() -ForegroundColor Cyan
        Write-Host ('-' * 78)

        foreach ($Finding in @($Findings | Where-Object { $_.Section -eq $Section })) {
            $Color = Get-IxmHealthStatusColor -Status $Finding.Status
            Write-Host ('[{0,-10}] {1,-30} {2}' -f $Finding.Status,$Finding.Check,$Finding.Value) -ForegroundColor $Color

            if (-not [string]::IsNullOrWhiteSpace([string]$Finding.Details)) {
                Write-Host ('             {0}' -f $Finding.Details) -ForegroundColor DarkGray
            }
        }
    }
}

function Get-IxmHealthRegistryAndSystem {
    param(
        [Parameter(Mandatory)]$Findings,
        [AllowNull()][string]$UcRoot
    )

    $Section = 'System'

    try {
        $HostName = [System.Net.Dns]::GetHostName()
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Hostname' -Value $HostName -Status 'INFO'
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Hostname' -Value 'Unavailable' -Status 'WARNING' -Details $_.Exception.Message
    }

    try {
        $Os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Windows' -Value ([string]$Os.Caption) -Status 'INFO' -Details ('Version {0}' -f $Os.Version)

        if ($Os.LastBootUpTime -is [datetime]) {
            $LastBoot = [datetime]$Os.LastBootUpTime
        }
        else {
            $LastBoot = [System.Management.ManagementDateTimeConverter]::ToDateTime([string]$Os.LastBootUpTime)
        }

        $Uptime = (Get-Date) - $LastBoot
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Uptime' -Value ('{0}d {1}h {2}m' -f $Uptime.Days,$Uptime.Hours,$Uptime.Minutes) -Status 'INFO' -Details ('Last boot: {0}' -f $LastBoot)
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Windows / uptime' -Value 'Unavailable' -Status 'WARNING' -Details $_.Exception.Message
    }

    $UcInfo = Get-IxmInstalledUcInfo
    if ($null -ne $UcInfo) {
        $VersionText = if ([string]::IsNullOrWhiteSpace($UcInfo.Version)) { 'Installed' } else { $UcInfo.Version }
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'UC Services version' -Value $VersionText -Status 'INFO' -Details $UcInfo.InstallLocation
    }
    else {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'UC Services version' -Value 'Not found in uninstall registry' -Status 'WARNING'
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$UcRoot)) {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'UC root' -Value $UcRoot -Status 'INFO'
    }
    else {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'UC root' -Value 'Could not resolve' -Status 'WARNING'
    }

    try {
        $CpuValue = (
            Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop
        ).CounterSamples[0].CookedValue

        $Cpu = [math]::Round([double]$CpuValue,2)
        $Status = 'OK'
        if ($Cpu -ge 90) { $Status = 'ATTENTION' }
        elseif ($Cpu -ge 80) { $Status = 'WARNING' }

        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Current CPU' -Value ('{0:N2}%' -f $Cpu) -Status $Status
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Current CPU' -Value 'Unavailable' -Status 'INFO' -Details $_.Exception.Message
    }

    try {
        $OsMem = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $TotalKb = [double]$OsMem.TotalVisibleMemorySize
        $FreeKb = [double]$OsMem.FreePhysicalMemory
        $UsedPct = if ($TotalKb -gt 0) { (($TotalKb - $FreeKb) / $TotalKb) * 100 } else { 0 }
        $UsedPct = [math]::Round($UsedPct,2)

        $Status = 'OK'
        if ($UsedPct -ge 95) { $Status = 'ATTENTION' }
        elseif ($UsedPct -ge 85) { $Status = 'WARNING' }

        $TotalGb = ($TotalKb * 1KB) / 1GB
        $FreeGb = ($FreeKb * 1KB) / 1GB
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Physical memory used' -Value ('{0:N2}%' -f $UsedPct) -Status $Status -Details ('Total={0:N1} GB; Free={1:N1} GB' -f $TotalGb,$FreeGb)
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Physical memory used' -Value 'Unavailable' -Status 'INFO' -Details $_.Exception.Message
    }

    try {
        $Disks = @(
            Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop |
            Sort-Object DeviceID
        )

        foreach ($Disk in $Disks) {
            $Size = [double]$Disk.Size
            $Free = [double]$Disk.FreeSpace

            if ($Size -le 0) { continue }

            $FreePct = ($Free / $Size) * 100
            $FreeGb = $Free / 1GB
            $TotalGb = $Size / 1GB

            $Status = 'OK'
            if ($FreePct -lt 10 -or $FreeGb -lt 5) {
                $Status = 'ATTENTION'
            }
            elseif ($FreePct -lt 15 -or $FreeGb -lt 10) {
                $Status = 'WARNING'
            }

            Add-IxmHealthFinding `
                -List $Findings `
                -Section $Section `
                -Check ('Drive {0}' -f $Disk.DeviceID) `
                -Value ('{0:N1} GB free / {1:N1} GB ({2:N1}% free)' -f $FreeGb,$TotalGb,$FreePct) `
                -Status $Status
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Drive space' -Value 'Unavailable' -Status 'WARNING' -Details $_.Exception.Message
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$UcRoot)) {
        $DbFile = Join-Path $UcRoot 'db\eeam21.db'
        if (Test-Path -LiteralPath $DbFile) {
            try {
                $DbItem = Get-Item -LiteralPath $DbFile -ErrorAction Stop
                Add-IxmHealthFinding -List $Findings -Section $Section -Check 'SQL database file size' -Value ('{0:N2} GB' -f ($DbItem.Length / 1GB)) -Status 'INFO' -Details $DbFile
            }
            catch {
                Add-IxmHealthFinding -List $Findings -Section $Section -Check 'SQL database file size' -Value 'Unavailable' -Status 'INFO'
            }
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'SQL database file' -Value 'Not found at expected path' -Status 'INFO' -Details $DbFile
        }
    }
}

function Get-IxmServiceInventory {
    try {
        return @(
            Get-CimInstance Win32_Service -ErrorAction Stop |
            Sort-Object DisplayName,Name
        )
    }
    catch {
        Write-Verbose ('Unable to inventory Win32_Service through CIM: {0}' -f $_.Exception.Message)
        return @()
    }
}

function Test-IxmBuiltInServiceAccount {
    param([AllowNull()][string]$StartName)

    if ([string]::IsNullOrWhiteSpace([string]$StartName)) { return $false }

    return ($StartName -match '(?i)^(LocalSystem|NT AUTHORITY\\(?:LocalService|NetworkService)|LocalService|NetworkService)$')
}

function Get-IxmServiceProcessStartTime {
    param($Service)

    if ($null -eq $Service -or [int]$Service.ProcessId -le 0) { return $null }

    try {
        return (Get-Process -Id ([int]$Service.ProcessId) -ErrorAction Stop).StartTime
    }
    catch {
        return $null
    }
}

function New-IxmHaContext {
    param([Parameter(Mandatory)][object[]]$ServiceInventory)

    $HostName = try { [System.Net.Dns]::GetHostName() } catch { $env:COMPUTERNAME }

    $Consolidated = @(
        $ServiceInventory | Where-Object {
            [string]$_.DisplayName -match '(?i)^MobiLink\s*-\s*Consolidated$' -or
            [string]$_.Name -match '(?i)^MobiLink\s*-\s*Consolidated$'
        }
    )

    $Remote = @(
        $ServiceInventory | Where-Object {
            [string]$_.DisplayName -match '(?i)^SQL Anywhere[-\s]*MobiLink\s+Remote$' -or
            [string]$_.Name -match '(?i)^SQL Anywhere[-\s]*MobiLink\s+Remote$'
        }
    )

    $Listener = @(
        $ServiceInventory | Where-Object {
            [string]$_.DisplayName -match '(?i)^SQL Anywhere[-\s]*MobiLinkListener$' -or
            [string]$_.Name -match '(?i)^SQL Anywhere[-\s]*MobiLinkListener$'
        }
    )

    $Voice = @(
        $ServiceInventory | Where-Object {
            [string]$_.DisplayName -match '(?i)^UC\s*Voice\s*Server$' -or
            [string]$_.Name -match '(?i)^UCVoiceServer$'
        }
    )

    $Role = 'Unknown'
    $RoleSource = 'Installed Windows services'

    if ($Consolidated.Count -gt 0 -and $Remote.Count -eq 0) {
        $Role = 'Consolidated Server'
    }
    elseif ($Remote.Count -gt 0 -or $Listener.Count -gt 0 -or $Voice.Count -gt 0) {
        $Role = 'Voice Server (Primary/Secondary)'
    }
    elseif ($Consolidated.Count -gt 0) {
        $Role = 'Consolidated Server'
    }

    return [pscustomobject]@{
        HostName = $HostName
        Role = $Role
        RoleSource = $RoleSource
        ServiceInventory = @($ServiceInventory)
        RelevantServices = @()
        RequiredServiceIssues = (New-Object System.Collections.Generic.List[string])
        AccountWarnings = (New-Object System.Collections.Generic.List[string])
        ServiceEvents = (New-Object System.Collections.Generic.List[object])
        MobiclientLogPath = ''
        MobiclientLogFound = $false
        MobiclientFileLastWrite = $null
        SyncMarkerCount = 0
        LastSyncSuccess = $null
        SyncAge = $null
        SyncIsRecent = $false
        RecentSyncLines = @()
        RecentLogErrors = @()
        LastLogFailure = $null
        OverallStatus = 'UNKNOWN'
        OverallReason = ''
    }
}

function Get-IxmHaServiceDefinitions {
    param([Parameter(Mandatory)][string]$Role)

    $CommonDb = [pscustomobject]@{
        Label = 'SQL Anywhere IXM database'
        Regex = '(?i)^SQL Anywhere[-\s]*(?:ASADB_UC|USADB_UC)$'
        Critical = $true
    }

    if ($Role -match 'Consolidated') {
        return @(
            [pscustomobject]@{ Label='MobiLink - Consolidated'; Regex='(?i)^MobiLink\s*-\s*Consolidated$'; Critical=$true },
            $CommonDb,
            [pscustomobject]@{ Label='DBWatcher'; Regex='(?i)^DBWatcher$'; Critical=$true },
            [pscustomobject]@{ Label='UC VPIMServer'; Regex='(?i)^UC\s*VPIMServer$|^UCVPIMServer$'; Critical=$false },
            [pscustomobject]@{ Label='UC Unified Messaging System Tasks Service'; Regex='(?i)^UC\s*Unified Messaging System Tasks Service$'; Critical=$false },
            [pscustomobject]@{ Label='UC Background Task Manager'; Regex='(?i)^UC\s*Background Task Manager$'; Critical=$false },
            [pscustomobject]@{ Label='UC Background File Organizer'; Regex='(?i)^UC\s*Background File Organizer$'; Critical=$false },
            [pscustomobject]@{ Label='UC Business Layer Service'; Regex='(?i)^UC\s*Business Layer Service$'; Critical=$false },
            [pscustomobject]@{ Label='UC Service Recovery Manager'; Regex='(?i)^UC\s*Service Recovery Manager$'; Critical=$false }
        )
    }

    if ($Role -match 'Voice Server') {
        return @(
            [pscustomobject]@{ Label='SQL Anywhere-MobiLink Remote'; Regex='(?i)^SQL Anywhere[-\s]*MobiLink\s+Remote$'; Critical=$true },
            [pscustomobject]@{ Label='SQL Anywhere-MobiLinkListener'; Regex='(?i)^SQL Anywhere[-\s]*MobiLinkListener$'; Critical=$true },
            $CommonDb,
            [pscustomobject]@{ Label='DBWatcher'; Regex='(?i)^DBWatcher$'; Critical=$true },
            [pscustomobject]@{ Label='UC Voice Server'; Regex='(?i)^UC\s*Voice\s*Server$|^UCVoiceServer$'; Critical=$false },
            [pscustomobject]@{ Label='UC Background Task Manager'; Regex='(?i)^UC\s*Background Task Manager$'; Critical=$false },
            [pscustomobject]@{ Label='UC Background File Organizer'; Regex='(?i)^UC\s*Background File Organizer$'; Critical=$false },
            [pscustomobject]@{ Label='UC Business Layer Service'; Regex='(?i)^UC\s*Business Layer Service$'; Critical=$false }
        )
    }

    # Unknown role: discover both MobiLink families but do not mark role-specific
    # services missing. Database and DBWatcher are still worth reporting.
    return @(
        [pscustomobject]@{ Label='MobiLink - Consolidated'; Regex='(?i)^MobiLink\s*-\s*Consolidated$'; Critical=$false },
        [pscustomobject]@{ Label='SQL Anywhere-MobiLink Remote'; Regex='(?i)^SQL Anywhere[-\s]*MobiLink\s+Remote$'; Critical=$false },
        [pscustomobject]@{ Label='SQL Anywhere-MobiLinkListener'; Regex='(?i)^SQL Anywhere[-\s]*MobiLinkListener$'; Critical=$false },
        [pscustomobject]@{ Label='SQL Anywhere IXM database'; Regex='(?i)^SQL Anywhere[-\s]*(?:ASADB_UC|USADB_UC)$'; Critical=$false },
        [pscustomobject]@{ Label='DBWatcher'; Regex='(?i)^DBWatcher$'; Critical=$false },
        [pscustomobject]@{ Label='UC Voice Server'; Regex='(?i)^UC\s*Voice\s*Server$|^UCVoiceServer$'; Critical=$false }
    )
}

function Confirm-IxmHaCriticalServicesForRefinedRole {
    param(
        [Parameter(Mandatory)]$Findings,
        [Parameter(Mandatory)]$HaContext
    )

    if ([string]$HaContext.Role -eq 'Unknown') { return }

    $Section = 'HA / MobiLink Services'
    $Services = @($HaContext.ServiceInventory)
    $Definitions = @(Get-IxmHaServiceDefinitions -Role ([string]$HaContext.Role) | Where-Object { $_.Critical })

    foreach ($Definition in $Definitions) {
        $Found = @(
            $Services | Where-Object {
                [string]$_.DisplayName -match $Definition.Regex -or [string]$_.Name -match $Definition.Regex
            }
        )

        if ($Found.Count -eq 0) {
            $Issue = '{0}: MISSING' -f $Definition.Label
            if (-not $HaContext.RequiredServiceIssues.Contains($Issue)) {
                $HaContext.RequiredServiceIssues.Add($Issue)
                Add-IxmHealthFinding -List $Findings -Section $Section -Check ('{0} (role refinement)' -f $Definition.Label) -Value 'MISSING' -Status 'ATTENTION' -Details ('Required after the local role was refined to {0}.' -f $HaContext.Role)
            }
            continue
        }

        foreach ($Svc in $Found) {
            $State = [string]$Svc.State
            $StartMode = [string]$Svc.StartMode
            $DisplayState = $State.ToUpperInvariant()
            if ($StartMode -match '(?i)disabled') { $DisplayState = 'DISABLED' }

            if ($StartMode -match '(?i)disabled' -or $State -ne 'Running') {
                $Issue = '{0}: {1}' -f $Definition.Label,$DisplayState
                if (-not $HaContext.RequiredServiceIssues.Contains($Issue)) {
                    $HaContext.RequiredServiceIssues.Add($Issue)
                    Add-IxmHealthFinding -List $Findings -Section $Section -Check ('{0} (role refinement)' -f $Definition.Label) -Value $DisplayState -Status 'ATTENTION' -Details ('Required after the local role was refined to {0}; DisplayName="{1}"; ServiceName="{2}"; Startup={3}; LogOnAs={4}' -f $HaContext.Role,$Svc.DisplayName,$Svc.Name,$StartMode,$Svc.StartName)
                }
            }
        }
    }
}

function Update-IxmHaRoleFromDatabase {
    param(
        [Parameter(Mandatory)]$Findings,
        [Parameter(Mandatory)]$Connection,
        [Parameter(Mandatory)]$HaContext
    )

    $Result = Invoke-IxmHealthSql -Connection $Connection -Sql @"
SELECT ServerName, ServerType
FROM DBA.LocationNodes
ORDER BY ServerName
"@

    if (-not $Result.Success) {
        Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink' -Check 'Role refinement' -Value 'Database lookup unavailable' -Status 'INFO' -Details $Result.Error
        return
    }

    $ShortHost = ([string]$HaContext.HostName).Split('.')[0]
    $Candidates = @(
        $Result.Rows | Where-Object {
            if ($_.IsNull('ServerName')) { return $false }
            $ServerName = [string]$_.ServerName
            $ServerShort = $ServerName.Split('.')[0]
            return ($ServerName -ieq [string]$HaContext.HostName -or $ServerShort -ieq $ShortHost)
        }
    )

    if ($Candidates.Count -eq 0) {
        Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink' -Check 'Role refinement' -Value 'Local hostname not found in LocationNodes' -Status 'INFO' -Details ('Host: {0}; service-derived role remains {1}' -f $HaContext.HostName,$HaContext.Role)
        return
    }

    $Type = $null
    try { $Type = [int]$Candidates[0].ServerType } catch { $Type = $null }

    $RoleMap = @{
        1 = 'Primary Voice Server'
        2 = 'Secondary Voice Server'
        4 = 'Primary Consolidated Server'
        8 = 'Secondary Consolidated Server'
    }

    $PrimaryConsolidatedCount = @(
        $Result.Rows | Where-Object {
            -not $_.IsNull('ServerType') -and [int]$_.ServerType -eq 4
        }
    ).Count

    if ($Type -eq 1 -and $PrimaryConsolidatedCount -eq 0) {
        $HaContext.Role = 'Single Server (Primary Voice)'
        $HaContext.RoleSource = 'DBA.LocationNodes / original topology rule'
        Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink' -Check 'Role refinement' -Value $HaContext.Role -Status 'INFO' -Details ('Matched local host {0} to ServerType 1 and found no ServerType 4 Primary Consolidated node. HA/MobiLink synchronization requirements are not applied to this topology.' -f $HaContext.HostName)
    }
    elseif ($null -ne $Type -and $RoleMap.ContainsKey($Type)) {
        $HaContext.Role = $RoleMap[$Type]
        $HaContext.RoleSource = 'DBA.LocationNodes'
        Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink' -Check 'Role refinement' -Value $HaContext.Role -Status 'OK' -Details ('Matched local host {0} to ServerType {1} in DBA.LocationNodes.' -f $HaContext.HostName,$Type)
    }
    else {
        Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink' -Check 'Role refinement' -Value 'LocationNodes match found, role not mapped' -Status 'INFO' -Details ('ServerType={0}' -f $Type)
    }
}

function Get-IxmTimestampFromLogLine {
    param([AllowEmptyString()][string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) { return $null }

    $Patterns = @(
        '(?<Stamp>\d{4}[-/]\d{2}[-/]\d{2}[ T]\d{2}:\d{2}:\d{2}(?:\.\d{1,6})?)',
        '(?<Stamp>\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}(?:\.\d{1,6})?)',
        '(?<Stamp>\d{8}\s+\d{2}:\d{2}:\d{2}(?:\.\d{1,6})?)'
    )

    foreach ($Pattern in $Patterns) {
        if ($Line -match $Pattern) {
            $Stamp = $Matches.Stamp
            $Parsed = [datetime]::MinValue
            if ([datetime]::TryParse($Stamp,[ref]$Parsed)) {
                return $Parsed
            }

            if ($Stamp -match '^\d{8}\s') {
                foreach ($Format in @('yyyyMMdd HH:mm:ss.fff','yyyyMMdd HH:mm:ss')) {
                    if ([datetime]::TryParseExact(
                        $Stamp,
                        $Format,
                        [System.Globalization.CultureInfo]::InvariantCulture,
                        [System.Globalization.DateTimeStyles]::None,
                        [ref]$Parsed
                    )) {
                        return $Parsed
                    }
                }
            }
        }
    }

    return $null
}

function Get-IxmHaServiceEvents {
    param(
        [Parameter(Mandatory)]$Findings,
        [Parameter(Mandatory)]$HaContext,
        [int]$Days = 7
    )

    $Section = 'HA / Service Events'
    $Since = (Get-Date).AddDays(-$Days)

    try {
        $Events = @(
            Get-WinEvent -FilterHashtable @{
                LogName = 'System'
                ProviderName = 'Service Control Manager'
                StartTime = $Since
            } -MaxEvents 2000 -ErrorAction Stop
        )
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Service Control Manager' -Value 'Event-log check unavailable' -Status 'INFO' -Details $_.Exception.Message
        return
    }

    $RelevantTokens = New-Object System.Collections.Generic.List[string]
    foreach ($Svc in $HaContext.RelevantServices) {
        foreach ($Token in @([string]$Svc.DisplayName,[string]$Svc.Name)) {
            if (-not [string]::IsNullOrWhiteSpace($Token) -and -not $RelevantTokens.Contains($Token)) {
                $RelevantTokens.Add($Token)
            }
        }
    }

    $FailureIdSet = @(7000,7001,7009,7011,7022,7023,7024,7031,7034,7038,7041)
    $Matches = New-Object System.Collections.Generic.List[object]

    foreach ($Event in $Events) {
        $Message = [string]$Event.Message
        if ([string]::IsNullOrWhiteSpace($Message)) { continue }

        $Relevant = ($Message -match '(?i)MobiLink|SQL Anywhere|DBWatcher|UC\s+(?:Voice|Background|Business|VPIM|Unified Messaging|Service Recovery)')
        if (-not $Relevant) {
            foreach ($Token in $RelevantTokens) {
                if ($Message.IndexOf($Token,[System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $Relevant = $true
                    break
                }
            }
        }
        if (-not $Relevant) { continue }

        $FailureText = ($Message -match '(?i)logon|log on|password|user name|account|failed|failure|could not start|cannot start|dependency|timed out|timeout|terminated unexpectedly|unexpectedly terminated')
        if (($FailureIdSet -notcontains [int]$Event.Id) -and -not $FailureText) { continue }

        $Credential = ([int]$Event.Id -in @(7038,7041) -or $Message -match '(?i)logon|log on|password|user name|account.*(?:failed|failure)|failed.*account')
        $StartFailure = ([int]$Event.Id -in @(7000,7001,7009,7011,7022,7023,7024) -or $Message -match '(?i)failed to start|could not start|cannot start|dependency|timed out|timeout')

        $ServiceLabel = 'IX Messaging service'
        foreach ($Token in $RelevantTokens) {
            if ($Message.IndexOf($Token,[System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $ServiceLabel = $Token
                break
            }
        }

        $OneLine = ($Message -replace '[\r\n]+',' ' -replace '\s+',' ').Trim()
        if ($OneLine.Length -gt 320) { $OneLine = $OneLine.Substring(0,320) + '...' }

        $Obj = [pscustomobject]@{
            TimeCreated = $Event.TimeCreated
            Id = [int]$Event.Id
            Service = $ServiceLabel
            CredentialFailure = $Credential
            StartFailure = $StartFailure
            Message = $OneLine
        }
        $Matches.Add($Obj)
        $HaContext.ServiceEvents.Add($Obj)
    }

    $Recent = @($Matches | Sort-Object TimeCreated -Descending | Select-Object -First 10)
    if ($Recent.Count -eq 0) {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'SCM startup / logon failures' -Value ('0 relevant in last {0} day(s)' -f $Days) -Status 'OK'
        return
    }

    foreach ($Event in $Recent) {
        $Kind = if ($Event.CredentialFailure) { 'START FAILURE / CREDENTIAL' } elseif ($Event.StartFailure) { 'START FAILURE' } else { 'SERVICE FAILURE' }
        Add-IxmHealthFinding -List $Findings -Section $Section -Check ('SCM {0} - {1}' -f $Event.Id,$Event.Service) -Value $Kind -Status 'WARNING' -Details ('{0:MM/dd/yyyy HH:mm:ss} - {1}' -f $Event.TimeCreated,$Event.Message)
    }
}

function Add-IxmHaOverallFinding {
    param(
        [Parameter(Mandatory)]$Findings,
        [Parameter(Mandatory)]$HaContext
    )

    $Now = Get-Date
    $RequiredFailed = ($HaContext.RequiredServiceIssues.Count -gt 0)
    $LastSync = $HaContext.LastSyncSuccess

    $FailureTimes = New-Object System.Collections.Generic.List[datetime]
    foreach ($Event in $HaContext.ServiceEvents) {
        if (($Event.CredentialFailure -or $Event.StartFailure) -and $null -ne $Event.TimeCreated) {
            $FailureTimes.Add([datetime]$Event.TimeCreated)
        }
    }
    if ($null -ne $HaContext.LastLogFailure) {
        $FailureTimes.Add([datetime]$HaContext.LastLogFailure)
    }

    $LastFailure = $null
    if ($FailureTimes.Count -gt 0) {
        $LastFailure = @($FailureTimes | Sort-Object -Descending | Select-Object -First 1)[0]
    }

    $FailureAfterSuccess = ($null -ne $LastFailure -and ($null -eq $LastSync -or $LastFailure -gt $LastSync))

    if ([string]$HaContext.Role -match '^Single Server') {
        $HaContext.OverallStatus = 'NOT APPLICABLE'
        $HaContext.OverallReason = 'No Primary Consolidated node was detected; this server is treated as a single-server topology and HA/MobiLink synchronization status does not apply.'
    }
    elseif ($RequiredFailed) {
        $HaContext.OverallStatus = 'FAILED'
        $HaContext.OverallReason = ('Required HA/MobiLink service issue(s): {0}' -f ($HaContext.RequiredServiceIssues -join '; '))
    }
    elseif ($FailureAfterSuccess -and $null -ne $LastFailure) {
        $HaContext.OverallStatus = 'FAILED'
        $HaContext.OverallReason = ('A service/log failure at {0:MM/dd/yyyy HH:mm:ss} is newer than the last confirmed successful synchronization.' -f $LastFailure)
    }
    elseif ($HaContext.MobiclientLogFound -and $HaContext.SyncIsRecent) {
        $HaContext.OverallStatus = 'HEALTHY'
        if ($null -ne $LastFailure) {
            $HaContext.OverallReason = ('Required services are running and synchronization succeeded after the most recent recorded failure ({0:MM/dd/yyyy HH:mm:ss}).' -f $LastFailure)
        }
        else {
            $HaContext.OverallReason = 'Required services are running and recent successful MobiLink synchronization is confirmed from Mobiclient.log.'
        }
    }
    elseif ($HaContext.MobiclientLogFound -and $HaContext.SyncMarkerCount -gt 0) {
        $HaContext.OverallStatus = 'WARNING'
        $HaContext.OverallReason = 'Successful synchronization markers exist, but the most recent success is stale or its timestamp could not be parsed.'
    }
    elseif (-not $HaContext.MobiclientLogFound -and $HaContext.Role -ne 'Unknown') {
        $HaContext.OverallStatus = 'WARNING'
        $HaContext.OverallReason = 'Required services do not show a current failure, but sync status cannot be verified because the documented Mobiclient.log was not found.'
    }
    elseif ($HaContext.Role -eq 'Unknown') {
        $HaContext.OverallStatus = 'UNKNOWN'
        $HaContext.OverallReason = 'The local IX Messaging HA role could not be determined with enough confidence to prove synchronization health.'
    }
    else {
        $HaContext.OverallStatus = 'UNKNOWN'
        $HaContext.OverallReason = 'There is not enough evidence to prove current HA synchronization status.'
    }

    $StatusMap = @{
        HEALTHY = 'OK'
        WARNING = 'WARNING'
        FAILED = 'ATTENTION'
        UNKNOWN = 'INFO'
        'NOT APPLICABLE' = 'NOT APPLICABLE'
    }

    Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink Summary' -Check 'Server' -Value ([string]$HaContext.HostName) -Status 'INFO'
    Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink Summary' -Check 'Detected role' -Value ([string]$HaContext.Role) -Status 'INFO' -Details ('Source: {0}' -f $HaContext.RoleSource)
    Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink Summary' -Check 'OVERALL HA STATUS' -Value $HaContext.OverallStatus -Status $StatusMap[$HaContext.OverallStatus] -Details $HaContext.OverallReason

    # Avaya Messaging Server Installation Guide 11.0 SP2, HA chapter.
    # Page 203 documents the role-specific synchronization services, DB\Logs\Mobiclient.log,
    # and the successful-sync marker below. Page 138 documents a 10-day recovery window
    # for loss of Primary-to-Consolidated synchronization. These are guidance only and are
    # deliberately not used as automatic health thresholds by this tool.
    Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink Guidance' -Check 'Avaya documented sync marker' -Value 'Completed processing of download stream' -Status 'INFO' -Details 'Avaya Messaging 11.0 SP2 Server Installation Guide, Verifying File Sync (page 203), uses this Mobiclient.log message to confirm that synchronization has finished.'

    if ([string]$HaContext.Role -ne 'Unknown' -and [string]$HaContext.Role -notmatch '^Single Server' -and $HaContext.SyncMarkerCount -eq 0) {
        Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink Guidance' -Check 'Initial HA synchronization' -Value 'Full synchronization not confirmed from Mobiclient.log' -Status 'WARNING' -Details 'For a new HA installation, Avaya warns to complete the full Primary/Consolidated synchronization, and then each Secondary synchronization, before logging in or proceeding with additional Secondary deployment. This warning does not prove the current system is a new installation.'
    }

    if ([string]$HaContext.Role -ne 'Unknown' -and [string]$HaContext.Role -notmatch '^Single Server' -and $HaContext.OverallStatus -ne 'HEALTHY') {
        Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink Guidance' -Check 'Primary-to-Consolidated sync recovery window' -Value '10 days documented for Avaya Messaging 11.0 SP2' -Status 'WARNING' -Details 'The Avaya Messaging 11.0 SP2 HA introduction states that if synchronization between the Primary voice server and Consolidated server fails, the connection should be restored within 10 days before all servers revert to Demo Mode. Treat this as release-specific operational guidance, not as a reason to delay troubleshooting.'
    }

    Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink Guidance' -Check 'Consolidated server failure impact' -Value 'Voice traffic may continue while UM services are unavailable' -Status 'INFO' -Details 'The Avaya Messaging 11.0 SP2 HA introduction states that if the Consolidated server fails, remaining voice servers can continue voice processing, while UM services such as calendar sync, email integration, and transcription are unavailable.'

    $CredentialFailures = @($HaContext.ServiceEvents | Where-Object { $_.CredentialFailure } | Sort-Object TimeCreated -Descending)
    if ($HaContext.OverallStatus -eq 'FAILED' -and $CredentialFailures.Count -gt 0) {
        $Latest = $CredentialFailures[0]
        Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink Summary' -Check 'Likely cause' -Value 'Service-account startup failure' -Status 'ATTENTION' -Details ('SCM {0} at {1:MM/dd/yyyy HH:mm:ss}. Verify the configured Log On credentials for {2} and other IX Messaging services using the same account.' -f $Latest.Id,$Latest.TimeCreated,$Latest.Service)
    }

    Add-IxmHealthFinding -List $Findings -Section 'HA / MobiLink Summary' -Check 'Voicemail-to-email note' -Value 'MobiLink is upstream of SMTP task processing in HA' -Status 'INFO' -Details 'MobiLink is not the SMTP client. A Voice-to-Consolidated synchronization failure can prevent Consolidated-side SMTP tasks from receiving/processing new voicemail-to-email work.'
}

function Show-IxmHaLogDetail {
    param([Parameter(Mandatory)]$HaContext)

    if (-not $HaContext.MobiclientLogFound) { return }
    if (@($HaContext.RecentSyncLines).Count -eq 0 -and @($HaContext.RecentLogErrors).Count -eq 0) { return }

    Write-Host ''
    $Answer = (Read-Host 'Show recent MobiLink success/error log entries? [y/N]').Trim()
    if ($Answer -notmatch '^(?i)y(?:es)?$') { return }

    Write-Host ''
    Write-Host 'RECENT MOBILINK SUCCESS MARKERS' -ForegroundColor Cyan
    Write-Host ('-' * 78)
    if (@($HaContext.RecentSyncLines).Count -eq 0) {
        Write-Host 'No successful download-stream markers found.' -ForegroundColor Yellow
    }
    else {
        foreach ($Line in @($HaContext.RecentSyncLines)) {
            Write-Host $Line -ForegroundColor DarkGray
        }
    }

    Write-Host ''
    Write-Host 'RECENT MOBILINK ERROR / FAILURE LINES' -ForegroundColor Cyan
    Write-Host ('-' * 78)
    if (@($HaContext.RecentLogErrors).Count -eq 0) {
        Write-Host 'No recent failure-pattern lines found in the inspected log tail.' -ForegroundColor Green
    }
    else {
        foreach ($Line in @($HaContext.RecentLogErrors)) {
            Write-Host $Line -ForegroundColor DarkGray
        }
    }
}

function Get-IxmHealthServices {
    param(
        [Parameter(Mandatory)]$Findings,
        [Parameter(Mandatory)]$HaContext
    )

    $Section = 'HA / MobiLink Services'
    $Services = @($HaContext.ServiceInventory)
    $Definitions = @(Get-IxmHaServiceDefinitions -Role ([string]$HaContext.Role))

    Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Role used for service checks' -Value ([string]$HaContext.Role) -Status 'INFO' -Details ('Derived from {0}; Primary/Secondary is refined from LocationNodes later when database access is available.' -f $HaContext.RoleSource)

    $Relevant = New-Object System.Collections.Generic.List[object]

    foreach ($Definition in $Definitions) {
        $Found = @(
            $Services | Where-Object {
                [string]$_.DisplayName -match $Definition.Regex -or [string]$_.Name -match $Definition.Regex
            }
        )

        if ($Found.Count -eq 0) {
            if ($Definition.Critical) {
                $HaContext.RequiredServiceIssues.Add(('{0}: MISSING' -f $Definition.Label))
                Add-IxmHealthFinding -List $Findings -Section $Section -Check $Definition.Label -Value 'MISSING' -Status 'ATTENTION' -Details 'Required for this detected HA/MobiLink role.'
            }
            else {
                Add-IxmHealthFinding -List $Findings -Section $Section -Check $Definition.Label -Value 'Not installed / not detected' -Status 'NOT APPLICABLE' -Details 'Related service; not treated as an HA synchronization failure by itself.'
            }
            continue
        }

        foreach ($Svc in $Found) {
            if (-not @($Relevant | Where-Object { $_.Name -eq $Svc.Name }).Count) {
                $Relevant.Add($Svc)
            }

            $State = [string]$Svc.State
            $StartMode = [string]$Svc.StartMode
            $DisplayState = $State.ToUpperInvariant()
            if ($StartMode -match '(?i)disabled') { $DisplayState = 'DISABLED' }

            $Status = if ($Definition.Critical) { 'OK' } else { 'INFO' }
            if ($StartMode -match '(?i)disabled') {
                $Status = if ($Definition.Critical) { 'ATTENTION' } else { 'WARNING' }
                if ($Definition.Critical) { $HaContext.RequiredServiceIssues.Add(('{0}: DISABLED' -f $Definition.Label)) }
            }
            elseif ($State -ne 'Running') {
                $Status = if ($Definition.Critical) { 'ATTENTION' } else { 'WARNING' }
                if ($Definition.Critical) { $HaContext.RequiredServiceIssues.Add(('{0}: {1}' -f $Definition.Label,$DisplayState)) }
            }

            $StartTime = Get-IxmServiceProcessStartTime -Service $Svc
            $StartText = if ($null -eq $StartTime) { 'n/a' } else { $StartTime.ToString('MM/dd/yyyy HH:mm:ss') }
            $Path = ([string]$Svc.PathName -replace '[\r\n]+',' ').Trim()
            $Details = 'DisplayName="{0}"; ServiceName="{1}"; Startup={2}; LogOnAs={3}; PID={4}; ProcessStart={5}; Path={6}' -f $Svc.DisplayName,$Svc.Name,$StartMode,$Svc.StartName,$Svc.ProcessId,$StartText,$Path

            Add-IxmHealthFinding -List $Findings -Section $Section -Check $Definition.Label -Value $DisplayState -Status $Status -Details $Details
        }
    }

    # Include all UC / SQL Anywhere / MobiLink / DBWatcher services so account and
    # event correlation does not miss non-UC-prefixed HA components.
    foreach ($Svc in @(
        $Services | Where-Object {
            [string]$_.DisplayName -match '(?i)^(UC|SQL Anywhere|MobiLink|DBWatcher)' -or
            [string]$_.Name -match '(?i)^(UC|SQL Anywhere|MobiLink|DBWatcher)'
        }
    )) {
        if (-not @($Relevant | Where-Object { $_.Name -eq $Svc.Name }).Count) {
            $Relevant.Add($Svc)
        }
    }
    $HaContext.RelevantServices = $Relevant.ToArray()

    # Service-account summary for all Avaya-related services.
    $AccountSection = 'Services Using IX Messaging Service Accounts'
    $Accounts = @(
        $HaContext.RelevantServices |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.StartName) } |
        Group-Object StartName |
        Sort-Object -Property @{ Expression = 'Count'; Descending = $true }, @{ Expression = 'Name'; Descending = $false }
    )

    if ($Accounts.Count -eq 0) {
        Add-IxmHealthFinding -List $Findings -Section $AccountSection -Check 'Account summary' -Value 'No service logon accounts discovered' -Status 'INFO'
    }
    else {
        foreach ($Group in $Accounts) {
            $Account = [string]$Group.Name
            Add-IxmHealthFinding -List $Findings -Section $AccountSection -Check $Account -Value ('{0} IX Messaging service(s)' -f $Group.Count) -Status 'INFO' -Details ((@($Group.Group | ForEach-Object { [string]$_.DisplayName }) | Sort-Object -Unique) -join '; ')

            $SameAccount = @($Services | Where-Object { [string]$_.StartName -ieq $Account })
            $Names = @($SameAccount | ForEach-Object { '{0} [{1}]' -f $_.DisplayName,$_.Name } | Sort-Object)
            $Preview = @($Names | Select-Object -First 30)
            $Suffix = if ($Names.Count -gt 30) { ' ... +{0} more' -f ($Names.Count - 30) } else { '' }
            Add-IxmHealthFinding -List $Findings -Section $AccountSection -Check ('All Windows services using {0}' -f $Account) -Value ('{0} service(s)' -f $SameAccount.Count) -Status 'INFO' -Details (($Preview -join '; ') + $Suffix)
        }
    }

    # Account-consistency heuristic within the SQL Anywhere/MobiLink family. This
    # is intentionally a warning rather than a failure because legitimate service
    # account layouts can vary by IX Messaging release and deployment.
    $SqlMobi = @(
        $HaContext.RelevantServices | Where-Object {
            [string]$_.DisplayName -match '(?i)^(SQL Anywhere|MobiLink)' -or
            [string]$_.Name -match '(?i)^(SQL Anywhere|MobiLink)'
        }
    )
    $SqlAccountGroups = @($SqlMobi | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.StartName) } | Group-Object StartName | Sort-Object Count -Descending)
    if ($SqlAccountGroups.Count -gt 1 -and $SqlAccountGroups[0].Count -ge 2) {
        $Dominant = [string]$SqlAccountGroups[0].Name
        foreach ($Group in @($SqlAccountGroups | Select-Object -Skip 1)) {
            foreach ($Svc in $Group.Group) {
                $Warning = '{0} uses {1}; most discovered SQL Anywhere/MobiLink services use {2}' -f $Svc.DisplayName,$Svc.StartName,$Dominant
                $HaContext.AccountWarnings.Add($Warning)
                Add-IxmHealthFinding -List $Findings -Section $AccountSection -Check 'UNEXPECTED LOGON ACCOUNT (heuristic)' -Value ([string]$Svc.DisplayName) -Status 'WARNING' -Details $Warning
            }
        }
    }

    try {
        $StartPending = @($Services | Where-Object { [string]$_.State -eq 'Start Pending' -or [string]$_.State -eq 'StartPending' })
        if ($StartPending.Count -eq 0) {
            Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'Services stuck StartPending' -Value '0' -Status 'OK'
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'Services stuck StartPending' -Value ($StartPending.Count.ToString()) -Status 'ATTENTION' -Details (($StartPending | Select-Object -ExpandProperty DisplayName) -join '; ')
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'Services stuck StartPending' -Value 'Check failed' -Status 'INFO'
    }

    # Preserve the original Avaya health-check intent: DBWatcher and UCArchiver
    # are explicitly reviewed as core services expected to be running on IXM servers.
    foreach ($CoreService in @(
        [pscustomobject]@{ Label='DBWatcher'; Regex='(?i)^DBWatcher$' },
        [pscustomobject]@{ Label='UCArchiver'; Regex='(?i)^UC\s*Archiver$|^UCArchiver$' }
    )) {
        try {
            $Core = @(
                $Services | Where-Object {
                    [string]$_.DisplayName -match $CoreService.Regex -or
                    [string]$_.Name -match $CoreService.Regex
                }
            )

            if ($Core.Count -eq 0) {
                Add-IxmHealthFinding -List $Findings -Section 'Services' -Check $CoreService.Label -Value 'Not detected' -Status 'WARNING' -Details 'The original Avaya health check identifies this as a service that should be running on all messaging servers.'
            }
            else {
                foreach ($Svc in $Core) {
                    $State = [string]$Svc.State
                    $Status = if ($State -eq 'Running') { 'OK' } else { 'ATTENTION' }
                    $Details = 'DisplayName="{0}"; ServiceName="{1}"; Startup={2}; LogOnAs={3}' -f $Svc.DisplayName,$Svc.Name,$Svc.StartMode,$Svc.StartName
                    Add-IxmHealthFinding -List $Findings -Section 'Services' -Check $CoreService.Label -Value $State -Status $Status -Details $Details
                }
            }
        }
        catch {
            Add-IxmHealthFinding -List $Findings -Section 'Services' -Check $CoreService.Label -Value 'Check failed' -Status 'WARNING' -Details $_.Exception.Message
        }
    }

    # The source health check displayed every SQL Anywhere service. Keep a compact
    # equivalent so a technician can see unexpected stopped SQL components even
    # when they are not part of the role-specific HA requirements above.
    try {
        $SqlAnywhereServices = @(
            $Services | Where-Object {
                [string]$_.DisplayName -like 'SQL Anywhere*' -or
                [string]$_.Name -like 'SQL Anywhere*'
            }
        )

        if ($SqlAnywhereServices.Count -eq 0) {
            Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'SQL Anywhere services' -Value 'None detected' -Status 'WARNING'
        }
        else {
            $SqlNonRunning = @($SqlAnywhereServices | Where-Object { [string]$_.State -ne 'Running' })
            $SqlDetails = @(
                $SqlAnywhereServices | ForEach-Object {
                    '{0}={1}' -f $_.DisplayName,$_.State
                }
            ) -join '; '
            $SqlStatus = if ($SqlNonRunning.Count -gt 0) { 'WARNING' } else { 'OK' }
            Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'SQL Anywhere services' -Value ('{0} running / {1} non-running / {2} total' -f ($SqlAnywhereServices.Count - $SqlNonRunning.Count),$SqlNonRunning.Count,$SqlAnywhereServices.Count) -Status $SqlStatus -Details $SqlDetails
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'SQL Anywhere services' -Value 'Check failed' -Status 'INFO' -Details $_.Exception.Message
    }

    # Restore the original running/stopped UC-service review, but summarize it so
    # Option 14 remains readable. Non-running UC services are warnings because some
    # services are legitimately role-dependent and must be interpreted by topology.
    try {
        $UcServices = @(
            $Services | Where-Object {
                [string]$_.DisplayName -match '(?i)^UC(?:\s|$)' -or
                [string]$_.Name -match '(?i)^UC'
            }
        )
        $UcRunning = @($UcServices | Where-Object { [string]$_.State -eq 'Running' })
        $UcNonRunning = @($UcServices | Where-Object { [string]$_.State -ne 'Running' })

        if ($UcServices.Count -eq 0) {
            Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'UC services running' -Value 'No UC services detected' -Status 'WARNING'
        }
        else {
            $RunningNames = @($UcRunning | ForEach-Object { [string]$_.DisplayName } | Sort-Object -Unique)
            $NonRunningNames = @($UcNonRunning | ForEach-Object { '{0}={1}' -f $_.DisplayName,$_.State } | Sort-Object -Unique)

            Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'UC services running' -Value ('{0} of {1}' -f $UcRunning.Count,$UcServices.Count) -Status 'INFO' -Details ($RunningNames -join '; ')

            if ($UcNonRunning.Count -eq 0) {
                Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'UC services non-running' -Value '0' -Status 'OK'
            }
            else {
                Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'UC services non-running' -Value ($UcNonRunning.Count.ToString()) -Status 'WARNING' -Details (('Review against the local server role; stopped UC services can be intentional. ' + ($NonRunningNames -join '; ')))
            }
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'UC service overview' -Value 'Check failed' -Status 'INFO' -Details $_.Exception.Message
    }

    foreach ($Optional in @(
        [pscustomobject]@{Label='Dialogic services'; Pattern='Dialogic*'},
        [pscustomobject]@{Label='Nuance services'; Pattern='*Nuance*'},
        [pscustomobject]@{Label='RealSpeak services'; Pattern='RealSpeak*'}
    )) {
        try {
            $OptionalServices = @(
                $Services | Where-Object { [string]$_.DisplayName -like $Optional.Pattern -or [string]$_.Name -like $Optional.Pattern }
            )

            if ($OptionalServices.Count -eq 0) {
                Add-IxmHealthFinding -List $Findings -Section 'Services' -Check $Optional.Label -Value 'Not installed' -Status 'NOT APPLICABLE'
            }
            else {
                $Stopped = @($OptionalServices | Where-Object { [string]$_.State -ne 'Running' })
                if ($Stopped.Count -gt 0) {
                    Add-IxmHealthFinding -List $Findings -Section 'Services' -Check $Optional.Label -Value ('{0} installed / {1} non-running' -f $OptionalServices.Count,$Stopped.Count) -Status 'WARNING' -Details (($Stopped | ForEach-Object { '{0}={1}' -f $_.DisplayName,$_.State }) -join '; ')
                }
                else {
                    Add-IxmHealthFinding -List $Findings -Section 'Services' -Check $Optional.Label -Value ('{0} running' -f $OptionalServices.Count) -Status 'OK'
                }
            }
        }
        catch {
            Add-IxmHealthFinding -List $Findings -Section 'Services' -Check $Optional.Label -Value 'Check failed' -Status 'INFO'
        }
    }

    try {
        $W3 = $Services | Where-Object { [string]$_.Name -eq 'W3SVC' } | Select-Object -First 1
        if ($null -eq $W3) {
            Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'IIS World Wide Web service' -Value 'Not installed' -Status 'NOT APPLICABLE'
        }
        elseif ([string]$W3.State -eq 'Running') {
            Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'IIS World Wide Web service' -Value 'Running' -Status 'OK'
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'IIS World Wide Web service' -Value ([string]$W3.State) -Status 'ATTENTION'
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'IIS World Wide Web service' -Value 'Check failed' -Status 'INFO'
    }

    try {
        if (Get-Module -ListAvailable -Name WebAdministration -ErrorAction SilentlyContinue) {
            Import-Module WebAdministration -ErrorAction Stop
            $Pools = @(Get-ChildItem IIS:\AppPools -ErrorAction Stop)

            if ($Pools.Count -eq 0) {
                Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'IIS application pools' -Value 'No application pools found' -Status 'INFO'
            }
            else {
                $NonRunning = New-Object System.Collections.Generic.List[string]
                foreach ($Pool in $Pools) {
                    $State = (Get-WebAppPoolState -Name $Pool.Name -ErrorAction Stop).Value
                    if ($State -ne 'Started') { $NonRunning.Add(('{0}={1}' -f $Pool.Name,$State)) }
                }

                if ($NonRunning.Count -eq 0) {
                    Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'IIS application pools' -Value ('{0} started / 0 stopped' -f $Pools.Count) -Status 'OK'
                }
                else {
                    Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'IIS application pools' -Value ('{0} non-started of {1}' -f $NonRunning.Count,$Pools.Count) -Status 'WARNING' -Details ($NonRunning -join '; ')
                }
            }
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'IIS application pools' -Value 'WebAdministration module not installed' -Status 'NOT APPLICABLE'
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section 'Services' -Check 'IIS application pools' -Value 'Check failed' -Status 'WARNING' -Details $_.Exception.Message
    }
}

function Invoke-IxmHealthSql {
    param(
        [Parameter(Mandatory)]$Connection,
        [Parameter(Mandatory)][string]$Sql
    )

    try {
        $Result = Invoke-IxmReadOnlyQuery -Connection $Connection -Sql $Sql
        return [pscustomobject]@{
            Success = $true
            Rows = @(Get-IxmResultRows -Result $Result)
            Error = ''
        }
    }
    catch {
        return [pscustomobject]@{
            Success = $false
            Rows = @()
            Error = $_.Exception.Message
        }
    }
}

function Get-IxmHealthDatabase {
    param(
        [Parameter(Mandatory)]$Findings,
        [Parameter(Mandatory)]$Connection,
        [AllowNull()]$HaContext
    )

    $Section = 'Database'

    # LocationNodes / server inventory.
    $LocationResult = Invoke-IxmHealthSql -Connection $Connection -Sql @"
SELECT LocNodePath, ServerName, ServerType
FROM DBA.LocationNodes
ORDER BY ServerType, ServerName
"@

    if ($LocationResult.Success) {
        $Rows = @($LocationResult.Rows)

        $NodeDetail = @(
            $Rows | ForEach-Object {
                $ServerName = if ($_.IsNull('ServerName')) { '' } else { [string]$_.ServerName }
                $ServerType = if ($_.IsNull('ServerType')) { '' } else { [string]$_.ServerType }
                $NodePath = if ($_.IsNull('LocNodePath')) { '' } else { [string]$_.LocNodePath }
                '{0} (Type {1}) {2}' -f $ServerName,$ServerType,$NodePath
            }
        ) -join '; '

        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'LocationNodes' -Value ('{0} node(s)' -f $Rows.Count) -Status 'INFO' -Details $NodeDetail

        $TypeZero = @($Rows | Where-Object { -not $_.IsNull('ServerType') -and [int]$_.ServerType -eq 0 })
        if ($TypeZero.Count -gt 0) {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Invalid ServerType 0 nodes' -Value ($TypeZero.Count.ToString()) -Status 'ATTENTION' -Details (($TypeZero | ForEach-Object { [string]$_.ServerName }) -join '; ')
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Invalid ServerType 0 nodes' -Value '0' -Status 'OK'
        }

        $RoleMap = @{
            1  = 'Primary Voice'
            2  = 'Secondary Voice'
            4  = 'Primary Consolidated'
            8  = 'Secondary Consolidated'
            16 = 'Dedicated CSE'
            32 = 'Report Server'
        }

        foreach ($RoleType in @(1,2,4,8,16,32)) {
            $RoleRows = @($Rows | Where-Object { -not $_.IsNull('ServerType') -and [int]$_.ServerType -eq $RoleType })
            if ($RoleRows.Count -gt 0) {
                Add-IxmHealthFinding -List $Findings -Section $Section -Check $RoleMap[$RoleType] -Value (($RoleRows | ForEach-Object { [string]$_.ServerName }) -join ', ') -Status 'INFO'
            }
        }

        $PrimaryConsolidatedRows = @($Rows | Where-Object { -not $_.IsNull('ServerType') -and [int]$_.ServerType -eq 4 })
        if ($PrimaryConsolidatedRows.Count -eq 0) {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Topology mode' -Value 'No Primary Consolidated node detected' -Status 'INFO' -Details 'The original Avaya health check treated the absence of ServerType 4 as a single-server topology.'
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Topology mode' -Value 'HA / Consolidated topology detected' -Status 'INFO' -Details ('Primary Consolidated: {0}' -f (($PrimaryConsolidatedRows | ForEach-Object { [string]$_.ServerName }) -join ', '))
        }
    }
    else {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'LocationNodes' -Value 'Unavailable' -Status 'WARNING' -Details $LocationResult.Error
    }

    # Tasks table.
    $TaskResult = Invoke-IxmHealthSql -Connection $Connection -Sql @"
SELECT TASKTYPE, COUNT(*) AS TaskCount
FROM DBA.TASKS
GROUP BY TASKTYPE
ORDER BY TASKTYPE
"@

    if ($TaskResult.Success) {
        $TaskRows = @($TaskResult.Rows)
        if ($TaskRows.Count -gt 0) {
            $TaskDetail = @(
                $TaskRows | ForEach-Object {
                    '{0}={1}' -f ([string]$_.TASKTYPE),([string]$_.TaskCount)
                }
            ) -join '; '

            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Task counts' -Value ('{0} task type(s)' -f $TaskRows.Count) -Status 'INFO' -Details $TaskDetail
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Task counts' -Value 'No rows' -Status 'INFO'
        }
    }
    else {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Task counts' -Value 'Unavailable' -Status 'INFO' -Details $TaskResult.Error
    }

    # Message count.
    $MessageCount = $null
    $MsgResult = Invoke-IxmHealthSql -Connection $Connection -Sql 'SELECT COUNT(*) AS MessageCount FROM DBA.MESSAGES'
    if ($MsgResult.Success -and @($MsgResult.Rows).Count -gt 0) {
        $MessageCount = [int64]$MsgResult.Rows[0].MessageCount
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Messages table count' -Value ([string]$MessageCount) -Status 'INFO'
    }
    else {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Messages table count' -Value 'Unavailable' -Status 'INFO' -Details $MsgResult.Error
    }

    $RunLoopScans = $true
    if ($null -ne $MessageCount -and $MessageCount -gt $script:HealthMessageLoopScanThreshold) {
        Write-Host ''
        Write-Host ('Performance guard: the MESSAGES table contains {0:N0} rows.' -f $MessageCount) -ForegroundColor Yellow
        Write-Host 'The two notification-loop checks perform full subject searches and may add database I/O.' -ForegroundColor Yellow
        $Approval = (Read-Host ('Run the extended message-loop scans anyway? [y/N] (threshold {0:N0})' -f $script:HealthMessageLoopScanThreshold)).Trim()
        $RunLoopScans = ($Approval -match '^(?i)y(?:es)?$')
    }

    if ($RunLoopScans) {
        # Looping-notification indicators copied from the source health-check logic.
        $LoopFw = Invoke-IxmHealthSql -Connection $Connection -Sql @"
SELECT COUNT(*) AS LoopCount
FROM DBA.MESSAGES
WHERE SUBJECT LIKE '%FW: FW: FW: FW: FW: FW: FW: FW: FW: FW:%'
"@

        if ($LoopFw.Success -and @($LoopFw.Rows).Count -gt 0) {
            $Count = [int]$LoopFw.Rows[0].LoopCount
            $Status = if ($Count -gt 100) { 'ATTENTION' } elseif ($Count -gt 0) { 'WARNING' } else { 'OK' }
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'FW: FW: notification loops' -Value ($Count.ToString()) -Status $Status -Details 'A high count can indicate looping/misconfigured message notification.'
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'FW: FW: notification loops' -Value 'Unavailable' -Status 'INFO' -Details $LoopFw.Error
        }

        $LoopSubject = Invoke-IxmHealthSql -Connection $Connection -Sql @"
SELECT COUNT(*) AS LoopCount
FROM DBA.MESSAGES
WHERE SUBJECT LIKE '%FW: Message with subject ''FW: Message with subject%'
"@

        if ($LoopSubject.Success -and @($LoopSubject.Rows).Count -gt 0) {
            $Count = [int]$LoopSubject.Rows[0].LoopCount
            $Status = if ($Count -gt 100) { 'ATTENTION' } elseif ($Count -gt 0) { 'WARNING' } else { 'OK' }
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Nested FW subject loops' -Value ($Count.ToString()) -Status $Status -Details 'A high count can indicate looping/misconfigured message notification.'
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Nested FW subject loops' -Value 'Unavailable' -Status 'INFO' -Details $LoopSubject.Error
        }
    }
    else {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'FW: FW: notification loops' -Value 'Skipped by production performance guard' -Status 'INFO' -Details ('MESSAGES table exceeded the {0:N0}-row automatic scan threshold.' -f $script:HealthMessageLoopScanThreshold)
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Nested FW subject loops' -Value 'Skipped by production performance guard' -Status 'INFO' -Details ('MESSAGES table exceeded the {0:N0}-row automatic scan threshold.' -f $script:HealthMessageLoopScanThreshold)
    }

    # Active MobiLink subscriptions. The original health check scoped this view
    # to the Primary Consolidated server. Preserve that intent while retaining the
    # newer role detection and structured findings.
    $LocalRole = if ($null -ne $HaContext) { [string]$HaContext.Role } else { 'Unknown' }
    $RunActiveSubscriptions = ($LocalRole -eq 'Primary Consolidated Server')

    if ($RunActiveSubscriptions) {
        $SubResult = Invoke-IxmHealthSql -Connection $Connection -Sql @"
SELECT name, last_upload_time, last_download_time
FROM DBA.vw_ml_ActiveSubscriptions
ORDER BY name
"@

        if ($SubResult.Success) {
            $Rows = @($SubResult.Rows)
            if ($Rows.Count -gt 0) {
                $Detail = @(
                    $Rows | ForEach-Object {
                        '{0}: upload={1}, download={2}' -f ([string]$_.name),([string]$_.last_upload_time),([string]$_.last_download_time)
                    }
                ) -join '; '

                $TopologyNote = 'Original Avaya guidance: ml_remote_consol_0, CSE, Web, and Report-server entries can legitimately show 1900-01-01 00:00:00.0 for last_upload_time. Do not treat that timestamp alone as a synchronization failure.'
                Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Active subscriptions' -Value ('{0} subscription(s)' -f $Rows.Count) -Status 'INFO' -Details ($TopologyNote + ' ' + $Detail)
            }
            else {
                Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Active subscriptions' -Value 'No active-subscription rows' -Status 'INFO' -Details 'Local role is Primary Consolidated Server.'
            }
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Active subscriptions' -Value 'Unavailable' -Status 'INFO' -Details $SubResult.Error
        }
    }
    else {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Active subscriptions' -Value ('Not run on local role: {0}' -f $LocalRole) -Status 'NOT APPLICABLE' -Details 'The original Avaya health check displays DBA.vw_ml_ActiveSubscriptions on the Primary Consolidated server.'
    }

    # WebLM license expiration.
    $LicenseResult = Invoke-IxmHealthSql -Connection $Connection -Sql 'SELECT szExpirationDate FROM DBA.LicenseInfo'
    if ($LicenseResult.Success -and @($LicenseResult.Rows).Count -gt 0) {
        $Raw = if ($LicenseResult.Rows[0].IsNull('szExpirationDate')) { '' } else { [string]$LicenseResult.Rows[0].szExpirationDate }
        $Expiration = Convert-IxmHealthDate -Value $Raw

        if ($null -ne $Expiration) {
            $Days = [math]::Floor(($Expiration.Date - (Get-Date).Date).TotalDays)
            $Status = 'OK'
            if ($Days -lt 0) { $Status = 'ATTENTION' }
            elseif ($Days -le 30) { $Status = 'ATTENTION' }
            elseif ($Days -le 90) { $Status = 'WARNING' }

            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'WebLM license expiration' -Value $Expiration.ToString('MM/dd/yyyy') -Status $Status -Details ('{0} day(s) from today' -f $Days)
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'WebLM license expiration' -Value $Raw -Status 'INFO' -Details 'Could not parse the expiration value as a date.'
        }
    }
    else {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'WebLM license expiration' -Value 'Unavailable' -Status 'INFO' -Details $LicenseResult.Error
    }
}

function Get-IxmHealthTodayActivity {
    param(
        [Parameter(Mandatory)]$Findings
    )

    $Section = 'Today Activity'

    if (-not (Test-Path -LiteralPath $LogRoot)) {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'VServer STATUS logs' -Value 'Log directory unavailable' -Status 'NOT APPLICABLE'
        return
    }

    $Today = (Get-Date).Date
    $StatusLogs = @(Get-DatedLogs -Type STATUS -StartDate $Today -EndDate $Today)

    if ($StatusLogs.Count -eq 0) {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'VServer STATUS logs' -Value 'No STATUS log for today' -Status 'WARNING'
        return
    }

    $Patterns = @(
        [pscustomobject]@{Check='Answered-call markers (IDMS!)'; Pattern='IDMS!'; Regex=$false; Status='INFO'; Detail='Count of IDMS! markers in today''s STATUS log.'},
        [pscustomobject]@{Check='Voicemail deposits'; Pattern='MessageAdd succeeded'; Regex=$false; Status='INFO'; Detail='Count of MessageAdd succeeded markers in today''s STATUS log.'},
        [pscustomobject]@{Check='TUI "you have" prompt occurrences'; Pattern='you have'; Regex=$false; Status='INFO'; Detail='This is a prompt-occurrence estimate, not a unique-user login count.'},
        [pscustomobject]@{Check='Short-message events'; Pattern='msg too short'; Regex=$false; Status='INFO'; Detail='Can represent messages below minimum length or RTP/media issues.'}
    )

    foreach ($PatternItem in $Patterns) {
        $Count = 0
        try {
            foreach ($Log in $StatusLogs) {
                foreach ($Match in @(
                    Select-String -LiteralPath $Log.Path -SimpleMatch -Pattern $PatternItem.Pattern -AllMatches -ErrorAction SilentlyContinue
                )) {
                    $Count += @($Match.Matches).Count
                }
            }

            $Status = $PatternItem.Status
            if ($PatternItem.Check -eq 'Short-message events' -and $Count -gt 0) {
                $Status = 'WARNING'
            }

            Add-IxmHealthFinding -List $Findings -Section $Section -Check $PatternItem.Check -Value ($Count.ToString()) -Status $Status -Details $PatternItem.Detail
        }
        catch {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check $PatternItem.Check -Value 'Check failed' -Status 'INFO' -Details $_.Exception.Message
        }
    }

    try {
        $DtmfLines = New-Object System.Collections.Generic.List[string]

        foreach ($Log in $StatusLogs) {
            foreach ($Match in @(
                Select-String -LiteralPath $Log.Path -SimpleMatch -Pattern "dtmfs returned: '" -ErrorAction SilentlyContinue
            )) {
                $DtmfLines.Add(([string]$Match.Line).Trim())
            }
        }

        $Last = @($DtmfLines | Select-Object -Last 20)
        if ($Last.Count -gt 0) {
            $EntryLabel = if ($Last.Count -eq 1) { 'entry' } else { 'entries' }
            $Detail = 'Last {0} {1}: {2}' -f $Last.Count,$EntryLabel,($Last -join ' || ')
        }
        else {
            $Detail = 'No DTMF buffer entries found today.'
        }

        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'DTMF buffer entries' -Value ($DtmfLines.Count.ToString()) -Status 'INFO' -Details $Detail
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'DTMF buffer entries' -Value 'Check failed' -Status 'INFO'
    }
}

function Get-IxmHealthDatabaseSyncLog {
    param(
        [Parameter(Mandatory)]$Findings,
        [AllowNull()][string]$UcRoot,
        [Parameter(Mandatory)]$HaContext
    )

    $Section = 'HA / MobiLink Log'

    if ([string]$HaContext.Role -match '^Single Server') {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Mobiclient.log' -Value 'Not required for single-server topology' -Status 'NOT APPLICABLE' -Details 'No Primary Consolidated node was detected, so HA MobiLink synchronization health is not evaluated.'
        return
    }

    if ([string]::IsNullOrWhiteSpace([string]$UcRoot)) {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Mobiclient.log' -Value 'UC root unavailable' -Status 'WARNING' -Details 'Sync status cannot be verified from the documented log.'
        return
    }

    $ExpectedDirs = @(
        (Join-Path $UcRoot 'DB\Logs'),
        (Join-Path $UcRoot 'Logs\DB')
    ) | Select-Object -Unique

    foreach ($Dir in $ExpectedDirs) {
        if (Test-Path -LiteralPath $Dir) {
            $Files = @(Get-ChildItem -LiteralPath $Dir -File -Filter '*.log' -ErrorAction SilentlyContinue)
            if ($Files.Count -eq 0) {
                Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Expected log directory' -Value 'WARNING: directory exists but contains no log files' -Status 'WARNING' -Details $Dir
            }
            else {
                Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Expected log directory' -Value ('{0} log file(s)' -f $Files.Count) -Status 'INFO' -Details $Dir
            }
        }
    }

    $Candidates = New-Object System.Collections.Generic.List[object]
    foreach ($Path in @(
        (Join-Path $UcRoot 'DB\Logs\Mobiclient.log'),
        (Join-Path $UcRoot 'Logs\DB\Mobiclient.log')
    )) {
        if (Test-Path -LiteralPath $Path) {
            try { $Candidates.Add((Get-Item -LiteralPath $Path -ErrorAction Stop)) } catch { }
        }
    }

    if ($Candidates.Count -eq 0) {
        try {
            foreach ($Item in @(Get-ChildItem -LiteralPath $UcRoot -File -Recurse -Filter 'Mobiclient.log' -ErrorAction SilentlyContinue)) {
                if (-not @($Candidates | Where-Object { $_.FullName -ieq $Item.FullName }).Count) {
                    $Candidates.Add($Item)
                }
            }
        }
        catch {
            Write-Verbose ('Recursive Mobiclient.log discovery failed: {0}' -f $_.Exception.Message)
        }
    }

    # Also surface related MobiLink/SQL Anywhere logs without dumping their contents.
    try {
        $Related = @(
            Get-ChildItem -LiteralPath $UcRoot -File -Recurse -Filter '*.log' -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -match '(?i)mobi|mobilink|sql.*anywhere|dbml|mlclient' -or
                $_.DirectoryName -match '(?i)\\DB\\Logs$|\\Logs\\DB$'
            } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 20
        )
        if ($Related.Count -gt 0) {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Related MobiLink / SQL logs' -Value ('{0} discovered (showing up to 20)' -f $Related.Count) -Status 'INFO' -Details (($Related | ForEach-Object { '{0} [{1:MM/dd/yyyy HH:mm:ss}]' -f $_.FullName,$_.LastWriteTime }) -join '; ')
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Related MobiLink / SQL logs' -Value 'Discovery incomplete' -Status 'INFO' -Details $_.Exception.Message
    }

    if ($Candidates.Count -eq 0) {
        $HaContext.MobiclientLogFound = $false
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Mobiclient.log' -Value 'NOT FOUND' -Status 'WARNING' -Details 'Sync status cannot be verified from the documented log. Services and Windows events are still evaluated.'
        return
    }

    $Item = @($Candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1)[0]
    $SyncLog = $Item.FullName
    $HaContext.MobiclientLogFound = $true
    $HaContext.MobiclientLogPath = $SyncLog
    $HaContext.MobiclientFileLastWrite = $Item.LastWriteTime

    Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Mobiclient.log' -Value $SyncLog -Status 'INFO' -Details ('Size={0:N2} MB; LastWriteTime={1:MM/dd/yyyy HH:mm:ss}' -f ($Item.Length / 1MB),$Item.LastWriteTime)

    try {
        $Matches = @(
            Select-String -LiteralPath $SyncLog -SimpleMatch -Pattern 'Completed processing of download stream' -ErrorAction SilentlyContinue
        )
        $HaContext.SyncMarkerCount = $Matches.Count
        $Recent = @($Matches | Select-Object -Last 10)
        $HaContext.RecentSyncLines = @($Recent | ForEach-Object { ([string]$_.Line).Trim() })

        $LastSuccess = $null
        for ($i = $Recent.Count - 1; $i -ge 0; $i--) {
            $CandidateTime = Get-IxmTimestampFromLogLine -Line ([string]$Recent[$i].Line)
            if ($null -ne $CandidateTime) {
                $LastSuccess = $CandidateTime
                break
            }
        }
        $HaContext.LastSyncSuccess = $LastSuccess

        if ($Matches.Count -eq 0) {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Successful sync markers' -Value '0' -Status 'WARNING' -Details 'No "Completed processing of download stream" marker was found.'
        }
        elseif ($null -eq $LastSuccess) {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Successful sync markers' -Value ('{0} found' -f $Matches.Count) -Status 'WARNING' -Details 'Markers are present, but a timestamp could not be parsed from the most recent entries; recency cannot be proven.'
        }
        else {
            $Age = (Get-Date) - $LastSuccess
            $HaContext.SyncAge = $Age
            # Diagnostic threshold only; it is intentionally not presented as Avaya policy.
            $HaContext.SyncIsRecent = ($Age.TotalMinutes -le 30 -and $Age.TotalMinutes -ge -5)
            $AgeText = '{0}d {1}h {2}m {3}s' -f [math]::Floor($Age.TotalDays),$Age.Hours,$Age.Minutes,$Age.Seconds
            $Status = if ($HaContext.SyncIsRecent) { 'OK' } else { 'WARNING' }
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Last successful sync' -Value ($LastSuccess.ToString('MM/dd/yyyy HH:mm:ss')) -Status $Status -Details ('Sync age: {0}; marker: Completed processing of download stream; freshness threshold used by this tool: 30 minutes.' -f $AgeText)
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Successful sync markers' -Value 'Check failed' -Status 'WARNING' -Details $_.Exception.Message
    }

    try {
        $Tail = @(Get-Content -LiteralPath $SyncLog -Tail 5000 -ErrorAction Stop)
        $ErrorLines = @(
            $Tail | Where-Object {
                $_ -match '(?i)\b(error|failed|failure|exception|unable|authentication|disconnect(?:ed)?|timeout|timed out)\b' -or
                $_ -match '(?i)connection\s+(?:failed|lost|refused|error|closed|terminated)'
            } | Select-Object -Last 10
        )
        $HaContext.RecentLogErrors = @($ErrorLines | ForEach-Object { ([string]$_).Trim() })

        $LastLogFailure = $null
        for ($i = $ErrorLines.Count - 1; $i -ge 0; $i--) {
            $CandidateTime = Get-IxmTimestampFromLogLine -Line ([string]$ErrorLines[$i])
            if ($null -ne $CandidateTime) {
                $LastLogFailure = $CandidateTime
                break
            }
        }
        $HaContext.LastLogFailure = $LastLogFailure

        if ($ErrorLines.Count -eq 0) {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Recent failure-pattern lines' -Value '0 in last 5000 lines' -Status 'OK'
        }
        else {
            $TimeDetail = if ($null -eq $LastLogFailure) { 'Timestamp of newest failure-pattern line could not be parsed.' } else { 'Newest parsed failure-pattern timestamp: {0:MM/dd/yyyy HH:mm:ss}.' -f $LastLogFailure }
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Recent failure-pattern lines' -Value ($ErrorLines.Count.ToString()) -Status 'WARNING' -Details ($TimeDetail + ' Use the detail prompt to review the recent lines; normal connection messages are not treated as failures unless they contain failure language.')
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Recent failure-pattern lines' -Value 'Check failed' -Status 'INFO' -Details $_.Exception.Message
    }
}

function Get-IxmHealthVpimEvents {
    param([Parameter(Mandatory)]$Findings)

    $Section = 'Windows Events'

    try {
        $Since = (Get-Date).AddDays(-30)

        $Events = @(
            Get-WinEvent -FilterHashtable @{
                LogName = 'System'
                Id = @(7031,7036)
                ProviderName = 'Service Control Manager'
                StartTime = $Since
            } -ErrorAction Stop |
            Where-Object {
                $_.Message -match 'UC VPIMServer' -and
                $_.Message -match '(?i)terminated'
            } |
            Sort-Object TimeCreated -Descending |
            Select-Object -First 5
        )

        if ($Events.Count -eq 0) {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'UC VPIMServer terminations' -Value '0 in last 30 days' -Status 'OK'
        }
        else {
            $Recent24 = @($Events | Where-Object { $_.TimeCreated -ge (Get-Date).AddHours(-24) }).Count
            $Status = if ($Recent24 -gt 0) { 'WARNING' } else { 'INFO' }
            $Detail = @($Events | ForEach-Object { $_.TimeCreated.ToString('MM/dd/yyyy HH:mm:ss') }) -join '; '

            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'UC VPIMServer terminations' -Value ('{0} recent event(s)' -f $Events.Count) -Status $Status -Details ('Most recent event times: {0}' -f $Detail)
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'UC VPIMServer terminations' -Value 'Event-log check unavailable' -Status 'INFO' -Details $_.Exception.Message
    }
}

function Get-IxmHealthTiffConverter {
    param(
        [Parameter(Mandatory)]$Findings,
        [AllowNull()][string]$UcRoot
    )

    $Section = 'Fax / Attachments'

    if ([string]::IsNullOrWhiteSpace([string]$UcRoot)) {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'TiffConverter exceptions' -Value 'UC root unavailable' -Status 'NOT APPLICABLE'
        return
    }

    $TiffLog = Join-Path $UcRoot 'logs\efsp\TiffConverter.log'
    if (-not (Test-Path -LiteralPath $TiffLog)) {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'TiffConverter exceptions' -Value 'TiffConverter.log not present' -Status 'NOT APPLICABLE'
        return
    }

    try {
        $Pattern = 'EndpointNotFoundException: There was no endpoint listening at https'
        $Count = @(
            Select-String -LiteralPath $TiffLog -SimpleMatch -Pattern $Pattern -ErrorAction SilentlyContinue
        ).Count

        $Status = if ($Count -gt 0) { 'WARNING' } else { 'OK' }
        $Detail = if ($Count -gt 0) {
            'The source health-check associates this with unsupported attachment types / outbound fax jobs that may be stuck in Initial status.'
        }
        else {
            'No matching EndpointNotFoundException entries found.'
        }

        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'TiffConverter exceptions' -Value ($Count.ToString()) -Status $Status -Details $Detail
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'TiffConverter exceptions' -Value 'Check failed' -Status 'INFO'
    }
}

function Get-IxmHealthMutare {
    param([Parameter(Mandatory)]$Findings)

    $Section = 'Mutare / IIS'

    $IisDate = Get-Date -Format 'yyMMdd'
    $TodayLog = Join-Path 'C:\inetpub\logs\LogFiles\W3SVC1' ('u_ex{0}.log' -f $IisDate)

    if (-not (Test-Path -LiteralPath $TodayLog)) {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Mutare POST activity' -Value 'Today IIS log not found' -Status 'NOT APPLICABLE' -Details $TodayLog
        return
    }

    try {
        $MutareMatches = @(
            Select-String -LiteralPath $TodayLog -SimpleMatch -Pattern 'Mutare' -ErrorAction SilentlyContinue
        )

        if ($MutareMatches.Count -gt 0) {
            $LastMatch = $MutareMatches | Select-Object -Last 1
            $Detail = ('Last matching IIS entry is at line {0}. Raw request content is intentionally not echoed by the health check.' -f $LastMatch.LineNumber)
        }
        else {
            $Detail = 'No Mutare entries were found in today''s IIS log.'
        }

        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Mutare POST activity' -Value ('{0} matching IIS entr{1}' -f $MutareMatches.Count,$(if ($MutareMatches.Count -eq 1) { 'y' } else { 'ies' })) -Status 'INFO' -Details $Detail
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Mutare POST activity' -Value 'Check failed' -Status 'INFO'
    }
}

function Get-IxmHealthTcp {
    param([Parameter(Mandatory)]$Findings)

    $Section = 'TCP Connections'

    try {
        $All = @(Get-NetTCPConnection -ErrorAction Stop)
        $Established = @($All | Where-Object { $_.State -eq 'Established' }).Count
        $Listening = @($All | Where-Object { $_.State -eq 'Listen' }).Count
        $TimeWait = @($All | Where-Object { $_.State -eq 'TimeWait' }).Count
        $Bound = @($All | Where-Object { $_.State -eq 'Bound' }).Count

        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'TCP totals' -Value ('{0} total' -f $All.Count) -Status 'INFO' -Details ('Established={0}; Listen={1}; TimeWait={2}; Bound={3}' -f $Established,$Listening,$TimeWait,$Bound)
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'TCP totals' -Value 'Unavailable' -Status 'INFO' -Details $_.Exception.Message
    }
}

function Invoke-IxmSystemHealthCheck {
    Write-Section 'IX Messaging System Health Check + HA / MobiLink'

    Write-Host 'This health check performs local read operations; database statements are restricted to SELECT by the tool.' -ForegroundColor Green
    Write-Host 'It does not restart services, modify service accounts/passwords, change startup types, modify IIS, change the IX Messaging database, or alter mailbox configuration.' -ForegroundColor Green
    Write-Host ''

    $Findings = New-Object System.Collections.Generic.List[object]
    $UcRoot = Resolve-IxmUcRoot
    $ServiceInventory = @(Get-IxmServiceInventory)
    $HaContext = New-IxmHaContext -ServiceInventory $ServiceInventory

    Write-Host 'Checking Windows, resources, and IX Messaging installation...' -ForegroundColor DarkGray
    Get-IxmHealthRegistryAndSystem -Findings $Findings -UcRoot $UcRoot

    Write-Host 'Checking IX Messaging database and refining local topology...' -ForegroundColor DarkGray

    $Connection = $null
    try {
        try {
            $Selected = Select-IxmDatabaseDsn
            Write-Host ('Using DSN: {0}' -f $Selected.Name) -ForegroundColor Cyan

            $Connection = Open-IxmDsnConnection -Name $Selected.Name
            Update-IxmHaRoleFromDatabase -Findings $Findings -Connection $Connection -HaContext $HaContext
            Get-IxmHealthDatabase -Findings $Findings -Connection $Connection -HaContext $HaContext
        }
        catch {
            Add-IxmHealthFinding -List $Findings -Section 'Database' -Check 'Database health checks' -Value 'Unavailable' -Status 'WARNING' -Details $_.Exception.Message
        }

        Write-Host 'Checking role-aware services, core services, and service accounts...' -ForegroundColor DarkGray
        Get-IxmHealthServices -Findings $Findings -HaContext $HaContext

        Write-Host 'Checking current-day VServer activity...' -ForegroundColor DarkGray
        Get-IxmHealthTodayActivity -Findings $Findings

        Write-Host 'Checking Service Control Manager for IXM/MobiLink startup and credential failures...' -ForegroundColor DarkGray
        Get-IxmHaServiceEvents -Findings $Findings -HaContext $HaContext

        Write-Host 'Discovering and checking MobiLink/Mobiclient synchronization logs...' -ForegroundColor DarkGray
        Get-IxmHealthDatabaseSyncLog -Findings $Findings -UcRoot $UcRoot -HaContext $HaContext

        Add-IxmHaOverallFinding -Findings $Findings -HaContext $HaContext

        Write-Host 'Checking VPIM event history, fax exceptions, Mutare IIS activity, and TCP state...' -ForegroundColor DarkGray
        Get-IxmHealthVpimEvents -Findings $Findings
        Get-IxmHealthTiffConverter -Findings $Findings -UcRoot $UcRoot
        Get-IxmHealthMutare -Findings $Findings
        Get-IxmHealthTcp -Findings $Findings

        $Rows = $Findings.ToArray()
        Show-IxmHealthFindings -Findings $Rows

        Write-Host ''
        Write-Host 'Notes:' -ForegroundColor Yellow
        Write-Host '  - HA status is not declared HEALTHY merely because a MobiLink service is running; the tool also requires a recent successful sync marker from Mobiclient.log.' -ForegroundColor Yellow
        Write-Host '  - The 30-minute synchronization freshness threshold is a diagnostic threshold used by this tool, not an Avaya support-policy threshold.' -ForegroundColor Yellow
        Write-Host '  - Avaya Messaging 11.0 SP2 page 203 identifies DB\Logs\Mobiclient.log and "Completed processing of download stream" as the file-sync completion check.' -ForegroundColor Yellow
        Write-Host '  - The Avaya Messaging 11.0 SP2 HA chapter documents a 10-day recovery window for Primary-to-Consolidated sync loss; this is release-specific guidance, not a troubleshooting delay threshold.' -ForegroundColor Yellow
        Write-Host '  - On a new HA deployment, Avaya warns not to log into Primary/Secondary servers until the initial full synchronization is complete.' -ForegroundColor Yellow
        Write-Host '  - MobiLink synchronization is upstream of Consolidated-side SMTP task processing in HA; MobiLink itself is not the SMTP client.' -ForegroundColor Yellow
        Write-Host '  - Account-consistency warnings are heuristic because legitimate service-account layouts can vary by release and deployment.' -ForegroundColor Yellow
        Write-Host '  - Stopped related UC services can be role-dependent; the broad UC-service list is for technician review, while only role-required HA services drive the HA result.' -ForegroundColor Yellow
        Write-Host '  - DBWatcher and UCArchiver are also shown explicitly to preserve the original Avaya health-check review.' -ForegroundColor Yellow
        Write-Host '  - "TUI you have" is counted as prompt occurrences, not unique subscriber logins.' -ForegroundColor Yellow
        Write-Host '  - ActiveSubscriptions is shown on the Primary Consolidated server, matching the original health-check intent; certain 1900-01-01 upload timestamps can be normal for topology-specific entries.' -ForegroundColor Yellow
        Write-Host '  - CPU, memory, disk, license-expiration, and loop-count thresholds are diagnostic thresholds used by this tool, not Avaya support policy.' -ForegroundColor Yellow
        Write-Host '  - Run under the least-privileged Windows account that can read the required IXM logs, IIS/service state, Windows System event log, and database DSN.' -ForegroundColor Yellow
        Write-Host '  - Enterprise PowerShell transcription can capture console output containing customer information.' -ForegroundColor Yellow

        Show-IxmHaLogDetail -HaContext $HaContext
        Export-ResultSet -Data $Rows -BaseName 'ixm_system_health'
    }
    finally {
        if ($null -ne $Connection) {
            try { $Connection.Close() } catch { Write-Verbose ('ODBC connection Close() cleanup failed: {0}' -f $_.Exception.Message) }
            try { $Connection.Dispose() } catch { Write-Verbose ('ODBC connection Dispose() cleanup failed: {0}' -f $_.Exception.Message) }
        }
    }
}


# -----------------------------------------------------------------------------
# Main menu actions
# -----------------------------------------------------------------------------

function Invoke-MailboxActivity {
    $Mailbox = Read-NumericValue -Prompt 'Mailbox / extension'
    $Range = Read-DateRange -CoverageMode LIFECYCLE

    Write-Section ("Mailbox Activity - {0} - {1}" -f $Mailbox,$Range.Description)

    $AllDeposits = @(Get-VoicemailDeposits -StartDate $Range.Start -EndDate $Range.End)
    $Deposits = @($AllDeposits | Where-Object { $_.Mailbox -eq $Mailbox })
    $MWI = @(Get-MWIEvents -Extension $Mailbox -StartDate $Range.Start -EndDate $Range.End)

    $Timeline = @(Show-MailboxTimeline -Mailbox $Mailbox -Deposits $Deposits -MWI $MWI)
    Export-ResultSet -Data $Timeline -BaseName ("mailbox_{0}_activity" -f $Mailbox)
}

function Invoke-MailboxVoicemails {
    $Mailbox = Read-NumericValue -Prompt 'Mailbox / extension'
    $Range = Read-DateRange -CoverageMode STATUS

    Write-Section ("Voicemails for Mailbox {0} - {1}" -f $Mailbox,$Range.Description)

    $Deposits = @(
        Get-VoicemailDeposits -StartDate $Range.Start -EndDate $Range.End |
        Where-Object { $_.Mailbox -eq $Mailbox }
    )

    Show-Deposits -Data $Deposits
    Export-ResultSet -Data $Deposits -BaseName ("mailbox_{0}_voicemails" -f $Mailbox)
}

function Invoke-MailboxClearHistory {
    $Mailbox = Read-NumericValue -Prompt 'Mailbox / extension'
    $Range = Read-DateRange -CoverageMode MWI

    Write-Section ("Voicemail / MWI Clear History - {0} - {1}" -f $Mailbox,$Range.Description)

    $Events = @(Get-MailboxClearEvents -Mailbox $Mailbox -StartDate $Range.Start -EndDate $Range.End)
    Show-ClearEvents -Data $Events

    if ($Events.Count -gt 0) {
        $Last = $Events | Select-Object -Last 1
        Write-Host ''
        Write-Host ('Last clear event: {0} {1}' -f $Last.Date,$Last.Time) -ForegroundColor Cyan
        Write-Host ('  {0}' -f $Last.Meaning)
        Write-Host ('  Notification acknowledged: {0}' -f $Last.Acknowledged)
    }

    Export-ResultSet -Data $Events -BaseName ("mailbox_{0}_clear_history" -f $Mailbox)
}

function Invoke-AllVoicemails {
    $Range = Read-DateRange -CoverageMode STATUS
    Write-Section ("All Successfully Stored Voicemails - {0}" -f $Range.Description)

    $Deposits = @(Get-VoicemailDeposits -StartDate $Range.Start -EndDate $Range.End)
    Show-Deposits -Data $Deposits
    Export-ResultSet -Data $Deposits -BaseName 'all_voicemails'
}

function Invoke-AllClearEvents {
    $Range = Read-DateRange -CoverageMode MWI
    Write-Section ("All Mailbox MWI Clear Events - {0}" -f $Range.Description)

    $Events = @(Get-AllClearEvents -StartDate $Range.Start -EndDate $Range.End)
    Show-ClearEvents -Data $Events
    Export-ResultSet -Data $Events -BaseName 'all_mailbox_clear_events'
}

function Invoke-MailboxSummary {
    $Range = Read-DateRange -CoverageMode STATUS
    Write-Section ("Voicemail Summary by Mailbox - {0}" -f $Range.Description)

    $Deposits = @(Get-VoicemailDeposits -StartDate $Range.Start -EndDate $Range.End)
    if ($Deposits.Count -eq 0) {
        Write-Host 'No successfully stored voicemail deposits were found.' -ForegroundColor Yellow
        return
    }

    $Summary = foreach ($Group in ($Deposits | Group-Object Mailbox)) {
        $Items = @($Group.Group | Sort-Object EventTime)
        $Last = $Items[-1]
        $Callers = @($Items | ForEach-Object { $_.CallerID } | Where-Object { $_ } | Sort-Object -Unique)

        [pscustomobject]@{
            Mailbox = $Group.Name
            'IXM-ID' = ($Items | Select-Object -Last 1).'IXM-ID'
            Voicemails = $Items.Count
            LastDeposit = if ($Last.EventTime) { $Last.EventTime.ToString('MM/dd/yyyy HH:mm:ss') } else { '' }
            Callers = ($Callers -join ', ')
        }
    }

    $Summary = @($Summary | Sort-Object @{Expression='Voicemails';Descending=$true},Mailbox)
    $Summary | Format-Table -Wrap -AutoSize
    Export-ResultSet -Data $Summary -BaseName 'voicemail_summary'
}

function Invoke-CallerSearch {
    $Caller = (Read-Host 'Caller ID / number to search').Trim()
    if ([string]::IsNullOrWhiteSpace($Caller)) { return }

    $Range = Read-DateRange -CoverageMode STATUS
    Write-Section ("Voicemails from Caller {0} - {1}" -f $Caller,$Range.Description)

    $Deposits = @(
        Get-VoicemailDeposits -StartDate $Range.Start -EndDate $Range.End |
        Where-Object { $_.CallerID -like "*$Caller*" }
    )

    Show-Deposits -Data $Deposits
    Export-ResultSet -Data $Deposits -BaseName ("caller_{0}_voicemails" -f $Caller)
}

function Invoke-MWIHistory {
    $Mailbox = Read-NumericValue -Prompt 'Mailbox / extension'
    $Range = Read-DateRange -CoverageMode MWI

    Write-Section ("MWI History - {0} - {1}" -f $Mailbox,$Range.Description)

    $Events = @(Get-MWIEvents -Extension $Mailbox -StartDate $Range.Start -EndDate $Range.End)
    Show-MWIEvents -Data $Events

    if ($Events.Count -gt 0) {
        $On = @($Events | Where-Object { $_.MWI -eq 'ON' }).Count
        $Off = @($Events | Where-Object { $_.MWI -eq 'OFF' }).Count
        $Last = $Events | Select-Object -Last 1

        Write-Host ''
        Write-Host ('MWI ON events : {0}' -f $On)
        Write-Host ('MWI OFF events: {0}' -f $Off)
        Write-Host ('Last event    : {0} {1} - MWI {2}' -f $Last.Date,$Last.Time,$Last.MWI)
    }

    Export-ResultSet -Data $Events -BaseName ("mailbox_{0}_mwi" -f $Mailbox)
}

function Invoke-LogCoverage {
    $Range = Read-DateRange -CoverageMode ALL
    Write-Section ("Log Coverage - {0}" -f $Range.Description)
    Show-LogCoverage -StartDate $Range.Start -EndDate $Range.End
}

# -----------------------------------------------------------------------------
# Startup validation
# -----------------------------------------------------------------------------

function Resolve-VServerLogRoot {
    param(
        [Parameter(Mandatory)][string]$PreferredPath
    )

    # Normal/default IX Messaging installation path.
    if (Test-Path -LiteralPath $PreferredPath) {
        return $PreferredPath
    }

    Write-Host ('Default IX Messaging log directory was not found: {0}' -f $PreferredPath) -ForegroundColor Yellow
    Write-Host ''
    Write-Host 'IX Messaging normally stores the VServer logs under:' -ForegroundColor Cyan
    Write-Host '  <drive>:\UC\logs\VServer' -ForegroundColor Cyan
    Write-Host 'For example: X:\UC\logs\VServer' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'Checking other local drives for \UC\logs\VServer...' -ForegroundColor DarkGray

    $Detected = New-Object System.Collections.Generic.List[string]

    try {
        $Drives = @(
            Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop |
            Select-Object -ExpandProperty DeviceID
        )
    }
    catch {
        # Fallback if CIM is unavailable.
        $Drives = @(
            Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Root.TrimEnd('\') }
        )
    }

    foreach ($Drive in @($Drives | Sort-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace([string]$Drive)) { continue }

        $Candidate = Join-Path ($Drive.TrimEnd('\') + '\') 'UC\logs\VServer'
        if (Test-Path -LiteralPath $Candidate) {
            if (-not $Detected.Contains($Candidate)) {
                $Detected.Add($Candidate)
            }
        }
    }

    if ($Detected.Count -eq 1) {
        Write-Host ('Detected IX Messaging VServer logs: {0}' -f $Detected[0]) -ForegroundColor Green
        return $Detected[0]
    }

    if ($Detected.Count -gt 1) {
        Write-Host ''
        Write-Host 'Multiple IX Messaging VServer log directories were detected:' -ForegroundColor Yellow

        for ($i = 0; $i -lt $Detected.Count; $i++) {
            Write-Host ('  {0}. {1}' -f ($i + 1),$Detected[$i])
        }

        Write-Host ''
        do {
            $Selection = (Read-Host ('Select 1-{0}, or press Enter to enter a different location' -f $Detected.Count)).Trim()

            if ([string]::IsNullOrWhiteSpace($Selection)) {
                break
            }

            $Number = 0
            if ([int]::TryParse($Selection,[ref]$Number) -and
                $Number -ge 1 -and
                $Number -le $Detected.Count) {

                return $Detected[$Number - 1]
            }

            Write-Host 'Invalid selection.' -ForegroundColor Yellow
        } while ($true)
    }
    else {
        Write-Host 'No other \UC\logs\VServer directory was detected automatically.' -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host 'Enter the drive letter or full VServer log path.' -ForegroundColor Cyan
    Write-Host 'Examples:' -ForegroundColor DarkGray
    Write-Host '  F:' -ForegroundColor DarkGray
    Write-Host '  H:' -ForegroundColor DarkGray
    Write-Host '  F:\UC\logs\VServer' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'If you enter only a drive letter, the script will automatically look for:' -ForegroundColor DarkGray
    Write-Host '  <drive>:\UC\logs\VServer' -ForegroundColor DarkGray
    Write-Host 'Press Enter with no value to continue with database-backed options 11 and 12; options 13 and 14 can locate or use other IX Messaging data separately.' -ForegroundColor DarkGray

    while ($true) {
        Write-Host ''
        $Entered = (Read-Host 'IX Messaging drive or VServer log directory').Trim()

        if ([string]::IsNullOrWhiteSpace($Entered)) {
            return $PreferredPath
        }

        $Candidate = $Entered

        # Allow F, F:, or F:\ and expand it to the normal IXM log path.
        if ($Entered -match '^(?<Drive>[A-Za-z])(?::)?(?:\\)?$') {
            $Candidate = ('{0}:\UC\logs\VServer' -f $Matches.Drive.ToUpper())
        }

        if (Test-Path -LiteralPath $Candidate) {
            Write-Host ('Using IX Messaging VServer logs: {0}' -f $Candidate) -ForegroundColor Green
            return $Candidate
        }

        Write-Host ('Directory not found: {0}' -f $Candidate) -ForegroundColor Yellow
        Write-Host 'Typical IX Messaging location: <drive>:\UC\logs\VServer' -ForegroundColor DarkGray
        Write-Host 'Try another drive/path, or press Enter to continue with database-backed options 11 and 12; options 13 and 14 can locate or use other IX Messaging data separately.' -ForegroundColor DarkGray
    }
}

Clear-Host
Write-Section 'Avaya IX Messaging - Mailbox Activity & MWI Log Search'

$LogRoot = Resolve-VServerLogRoot -PreferredPath $LogRoot

if (-not (Test-Path -LiteralPath $LogRoot)) {
    Write-Host ''
    Write-Host 'No VServer log directory is currently available.' -ForegroundColor Yellow
    Write-Host 'VServer log options may return no data, but database-backed options 11 and 12 remain available; options 13 and 14 can use other IX Messaging data separately.' -ForegroundColor Yellow
}


# Option 15 - STATUS based inbound call analysis
function Invoke-IxmInboundCallAnalysis {
    $Range = Read-DateRange -CoverageMode STATUS
    $CallerFilter = (Read-Host 'Caller ID filter (Enter for all)').Trim()
    $MailboxFilter = (Read-Host 'Mailbox filter (Enter for all)').Trim()
    Write-Host '  1. All calls'
    Write-Host '  2. Unsuccessful / uncertain only'
    $View = (Read-Host 'View [1]').Trim()
    $Rows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($Log in @(Get-DatedLogs -Type STATUS -StartDate $Range.Start -EndDate $Range.End)) {
        Write-Host ('Reading {0}...' -f $Log.Name) -ForegroundColor DarkGray
        $Active = @{}
        $PendingAdd = $null
        $Handle = $null
        try {
            $Handle = New-SharedReader -Path $Log.Path
            while ($null -ne ($Line = $Handle.Reader.ReadLine())) {
                if ($Line -match '\[F:FastMessageAdd\]\s+start,.*?Channel:\s*(\d+),\s*CallerIDNumber:\s*([^,]*)') {
                    $AddChannel = [string]([int]$Matches[1])
                    $AddCaller = $Matches[2].Trim()
                    $PendingAdd = $null
                    if ($Active.ContainsKey($AddChannel) -and $Active[$AddChannel].CallerID -eq $AddCaller) {
                        $PendingAdd = $AddChannel
                    }
                }
                elseif ($Line -match 'XEEAM_MessageAdd succeeded') {
                    if ($null -ne $PendingAdd -and $Active.ContainsKey($PendingAdd)) {
                        $Active[$PendingAdd].Saved = $true
                        $Active[$PendingAdd].Evidence.Add('FastMessageAdd success')
                    }
                    $PendingAdd = $null
                }
                elseif ($Line -match '\[F:FastMessageAdd\] end, retval:' -or
                        $Line -match 'MessageAdd returned\s*=\s*(?!0\b)') {
                    $PendingAdd = $null
                }
                $Ch = Get-ChannelFromLine -Line $Line
                if ($null -eq $Ch -and $Line -match '<CHAN>(\d+)</CHAN>') { $Ch = [int]$Matches[1] }
                if ($null -eq $Ch) { continue }
                $Key = [string]$Ch
                if ($Line -match 'IDMS!' -and $Line -match '<CALLERID>([^<]+)</CALLERID>') {
                    if ($Active.ContainsKey($Key)) {
                        $Prior = $Active[$Key]
                        $Prior.Incomplete = $true
                        $Rows.Add((ConvertTo-IxmInboundRow $Prior))
                    }
                    $Number = ($Matches[1] -split '[\x00-\x1f]')[0].Trim()
                    $Called = ''
                    if ($Line -match '<CALLEDID>([^<]+)</CALLEDID>') { $Called = ($Matches[1] -split '[\x00-\x1f]')[0].Trim() }
                    $Name = ''
                    if ($Line -match '<CALLERNAME>([^<]*)</CALLERNAME>') { $Name = $Matches[1] }
                    $Active[$Key] = [pscustomobject]@{
                        Channel=$Ch; CallerID=$Number; CallerName=$Name; Mailbox=$Called; MailboxID=''
                        Start=(Get-LineTime -Line $Line -FileDate $Log.Date); End=$null
                        Greeting=$false; RecordingAttempts=0; TooShort=0; Saved=$false
                        HangupDuringGreeting=$false; Ended=$false; Incomplete=$false
                        Log=$Log.Name; Evidence=(New-Object 'System.Collections.Generic.List[string]')
                    }
                    $Active[$Key].Evidence.Add('IDMS')
                    continue
                }
                if (-not $Active.ContainsKey($Key)) { continue }
                $S = $Active[$Key]
                if ($Line -match 'ProcessSharedExtension\], mailboxID:\s*(\d+)') { $S.MailboxID=$Matches[1] }
                if ($Line -match 'State 70 Data: Play Greeting') {
                    $S.Greeting=$true
                    $S.Evidence.Add('Greeting started')
                }
                if ($Line -match 'Recording Message Mbx\s+(\d+)' -or $Line -match 'Re-Recording Message Mailbox\s+(\d+)') {
                    $S.Mailbox=$Matches[1]
                }
                if ($Line -match 'UMST sckOpen or sckConnected, send command:' -and $Line -match '<CMD>INMSGSTART</CMD>' -and $Line -match ('<CHAN>{0}</CHAN>' -f $Ch)) {
                    $S.RecordingAttempts++
                    $S.Evidence.Add('Recording start')
                }
                if ($Line -match 'Data: Message too Short Mbx') {
                    $S.TooShort++
                    $S.Evidence.Add('Message too Short')
                }
                if ($Line -match ('\b{0} FROM:70 TO:304\b' -f $Ch)) {
                    $S.HangupDuringGreeting=$true
                    $S.Evidence.Add('70 to 304')
                }
                # Background MessageAdd successes cannot safely be assigned to a
                # channel by proximity; evidence of saved messages is UNKNOWN here.
                if ($Line -match 'State 304 Data: User hanging up' -or
                    ($Line -match '<CMD>CALLENDED</CMD>' -and $Line -match ('<CHAN>{0}</CHAN>' -f $Ch))) {
                    $S.End = Get-LineTime -Line $Line -FileDate $Log.Date
                    $S.Ended=$true
                    $S.Evidence.Add('Call ended')
                    $Rows.Add((ConvertTo-IxmInboundRow $S))
                    $Active.Remove($Key)
                }
            }
            foreach ($S in $Active.Values) {
                $S.Incomplete=$true
                $Rows.Add((ConvertTo-IxmInboundRow $S))
            }
        }
        finally {
            if ($null -ne $Handle) {
                if ($null -ne $Handle.Reader) { $Handle.Reader.Dispose() }
                if ($null -ne $Handle.Stream) { $Handle.Stream.Dispose() }
            }
        }
    }
    $Filtered = @($Rows | Where-Object {
        ($CallerFilter -eq '' -or $_.CallerID -like ('*'+$CallerFilter+'*')) -and
        ($MailboxFilter -eq '' -or $_.Mailbox -like ('*'+$MailboxFilter+'*'))
    } | Sort-Object Date,Start,Channel)
    if ($View -eq '2') { $Filtered = @($Filtered | Where-Object { $_.Result -ne 'VOICEMAIL_SAVED' }) }
    Write-Section 'Inbound / Abandoned Call Analysis (STATUS evidence)'
    if ($Filtered.Count -eq 0) { Write-Host 'No matching inbound calls.' -ForegroundColor Yellow; return }
    Write-Host 'Outcome summary:' -ForegroundColor Cyan
    $Filtered | Group-Object Result | Sort-Object Count -Descending | ForEach-Object { Write-Host ('  {0,-38} {1,6}' -f $_.Name,$_.Count) }
    $Filtered | Select-Object Date,Start,CallerID,Mailbox,Channel,DurationSec,RecordingAttempts,TooShort,Result,Confidence |
        Format-Table -AutoSize | Out-Host
    Write-Host 'No saved-message claim is made without channel-specific proof. Silent audio cannot be verified from STATUS alone.' -ForegroundColor Yellow
    Export-ResultSet -Data $Filtered -BaseName 'ixm_inbound_call_analysis'
}
function ConvertTo-IxmInboundRow {
    param([object]$S)
    $Result='NO_MESSAGE_UNDETERMINED'; $Confidence='LOW'
    if ($S.Incomplete -or -not $S.Ended) { $Result='INCOMPLETE_LOG_EVIDENCE' }
    elseif ($S.Saved) { $Result='VOICEMAIL_SAVED'; $Confidence='HIGH' }
    elseif ($S.HangupDuringGreeting -and $S.RecordingAttempts -eq 0) {
        $Result='ABANDONED_DURING_GREETING'; $Confidence='HIGH'
    }
    elseif ($S.TooShort -gt 0) {
        $Result='RECORDING_TOO_SHORT'; $Confidence='HIGH'
    }
    elseif ($S.RecordingAttempts -gt 0) {
        $Result='RECORDING_OUTCOME_UNVERIFIED'; $Confidence='LOW'
    }
    elseif ($S.Greeting) { $Result='GREETING_ENDED_NO_RECORDING'; $Confidence='MEDIUM' }
    $Duration=$null
    if ($S.Start -and $S.End) { $Duration=[math]::Round(($S.End-$S.Start).TotalSeconds,1) }
    return [pscustomobject]@{
        Date=if ($S.Start) {$S.Start.ToString('MM/dd/yyyy')} else {''}
        Start=if ($S.Start) {$S.Start.ToString('HH:mm:ss')} else {''}
        End=if ($S.End) {$S.End.ToString('HH:mm:ss')} else {''}
        DurationSec=$Duration; CallerID=$S.CallerID; CallerName=$S.CallerName
        Mailbox=$S.Mailbox; MailboxID=$S.MailboxID; Channel=$S.Channel
        RecordingAttempts=$S.RecordingAttempts; TooShort=$S.TooShort; Saved=$S.Saved
        Result=$Result; Confidence=$Confidence; Evidence=($S.Evidence -join ' | ')
        Log=$S.Log
    }
}

# -----------------------------------------------------------------------------
# Main menu
# -----------------------------------------------------------------------------

do {
    Clear-Host
    Write-Section 'Avaya IX Messaging - Mailbox Activity & MWI Log Search'

    Write-Host ('Version      : {0}' -f $ToolVersion) -ForegroundColor DarkGray
    Write-Host ('Log directory: {0}' -f $LogRoot) -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  1. Mailbox lifecycle (voicemail deposits + MWI + clear events)'
    Write-Host '  2. Voicemails left for one mailbox'
    Write-Host '  3. Voicemail / MWI clear history for one mailbox'
    Write-Host '  4. All voicemails left for all mailboxes'
    Write-Host '  5. All mailbox MWI clear events'
    Write-Host '  6. Voicemail summary by mailbox'
    Write-Host '  7. Search voicemails by caller ID'
    Write-Host '  8. Raw MWI history for one mailbox'
    Write-Host '  9. Log coverage / diagnostics'
    Write-Host ' 10. Extension Graph / email sync history'
    Write-Host ' 11. Export mailboxes / email addresses'
    Write-Host ' 12. Current mailbox status / health'
    Write-Host ' 13. Graph / Exchange mailbox failure audit'
    Write-Host ' 14. IX Messaging system health check + HA / MobiLink'
    Write-Host ' 15. Inbound calls / abandoned voicemail analysis'
    Write-Host '  0. Exit'
    Write-Host ''

    $Choice = (Read-Host 'Select an option').Trim()

    try {
        switch ($Choice) {
            '1' { Invoke-MailboxActivity; Pause-Tool }
            '2' { Invoke-MailboxVoicemails; Pause-Tool }
            '3' { Invoke-MailboxClearHistory; Pause-Tool }
            '4' { Invoke-AllVoicemails; Pause-Tool }
            '5' { Invoke-AllClearEvents; Pause-Tool }
            '6' { Invoke-MailboxSummary; Pause-Tool }
            '7' { Invoke-CallerSearch; Pause-Tool }
            '8' { Invoke-MWIHistory; Pause-Tool }
            '9' { Invoke-LogCoverage; Pause-Tool }
            '10' { Invoke-ExtensionGraphSyncHistory; Pause-Tool }
            '11' { Invoke-MailboxDirectoryExport; Pause-Tool }
            '12' { Invoke-CurrentMailboxHealth; Pause-Tool }
            '13' { Invoke-GraphEmailFailureAudit; Pause-Tool }
            '14' { Invoke-IxmSystemHealthCheck; Pause-Tool }
            '15' { Invoke-IxmInboundCallAnalysis; Pause-Tool }
            '0' { }
            default {
                Write-Host 'Invalid selection.' -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
        }
    }
    catch [System.OperationCanceledException] {
        Write-Host ''
        Write-Host ('Canceled: {0}' -f $_.Exception.Message) -ForegroundColor Yellow
        Pause-Tool
    }
    catch {
        Write-Host ''
        Write-Host ('ERROR: {0}' -f $_.Exception.Message) -ForegroundColor Red
        Pause-Tool
    }

} until ($Choice -eq '0')

Write-Host 'Goodbye.' -ForegroundColor Cyan
