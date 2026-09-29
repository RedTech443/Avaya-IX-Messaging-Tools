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
    mailbox and Graph access. SQL access remains read-only.

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

$ToolVersion = '2.5.0'
$LogRoot = 'X:\UC\logs\VServer'
$ExportRoot = 'C:\Temp\IXM-Reports'

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

        $Result = New-Object PSObject -Property @{
            Start = $Today.AddDays(-($Days - 1))
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
                # Ignore malformed dated filenames.
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

            while (($Line = $Reader.ReadLine()) -ne $null) {
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

            while (($Line = $Reader.ReadLine()) -ne $null) {
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
    foreach ($Event in @(Get-MWIEvents -Extension $Mailbox -StartDate $StartDate -EndDate $EndDate)) {
        if ($Event.MWI -eq 'OFF') {
            $Results.Add((Convert-ToClearEvent -MWIEvent $Event))
        }
    }
    return @($Results | Sort-Object EventTime)
}

function Get-AllClearEvents {
    param(
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate    )

    $Results = New-Object System.Collections.Generic.List[object]
    foreach ($Event in @(Get-AllMWIEvents -StartDate $StartDate -EndDate $EndDate)) {
        if ($Event.MWI -eq 'OFF') {
            $Results.Add((Convert-ToClearEvent -MWIEvent $Event))
        }
    }
    return @($Results | Sort-Object EventTime,Mailbox)
}

# -----------------------------------------------------------------------------
# Display / export helpers
# -----------------------------------------------------------------------------

function Export-ResultSet {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Data,
        [Parameter(Mandatory)][string]$BaseName
    )

    if (@($Data).Count -eq 0) { return }

    Write-Host ''
    $Answer = (Read-Host 'Export these results to CSV? [y/N]').Trim()
    if ($Answer -notmatch '^(?i)y(?:es)?$') { return }

    $SafeName = $BaseName -replace '[^A-Za-z0-9_.-]','_'
    $DefaultFileName = ('{0}_{1}.csv' -f $SafeName,(Get-Date -Format 'yyyyMMdd_HHmmss'))
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
        # If the GUI dialog cannot be opened, fall back to the legacy export folder.
        Write-Host ('Save As dialog unavailable ({0}). Using default export folder.' -f $_.Exception.Message) -ForegroundColor Yellow

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
        $Data | Select-Object * -ExcludeProperty EventTime | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
        Write-Host ('Saved: {0}' -f $Path) -ForegroundColor Green
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
                # Ignore malformed dated filenames.
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

            while (($Line = $Reader.ReadLine()) -ne $null) {
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

            while (($Line = $Reader.ReadLine()) -ne $null) {
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

            while (($Line = $Reader.ReadLine()) -ne $null) {
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
                # Optional display attributes may be unavailable on some hosts.
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
    $Candidates = @(Get-SystemSqlAnywhereDsns)

    if ($Candidates.Count -eq 0) {
        throw 'No SQL Anywhere System DSNs were found on this server.'
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
                try { $Connection.Close() } catch {}
                try { $Connection.Dispose() } catch {}
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

function Invoke-IxmReadOnlyQuery {
    param(
        [Parameter(Mandatory)]$Connection,
        [Parameter(Mandatory)][string]$Sql
    )

    if ($Sql -match '(?i)\b(INSERT|UPDATE|DELETE|ALTER|DROP|CREATE|TRUNCATE|MERGE|GRANT|REVOKE)\b') {
        throw 'Only read-only SELECT queries are permitted by the mailbox exporter.'
    }

    $Command = $Connection.CreateCommand()
    $Command.CommandText = $Sql
    $Command.CommandTimeout = 120

    $Adapter = New-Object System.Data.Odbc.OdbcDataAdapter $Command
    $Table = New-Object System.Data.DataTable
    [void]$Adapter.Fill($Table)

    # DataTable implements IEnumerable. Prevent PowerShell from silently
    # converting the table into DataRow objects when returning it.
    Write-Output -NoEnumerate $Table
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
        $SourceRows = @($Table.Rows)    }
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

    Write-Host 'This database operation is READ ONLY.' -ForegroundColor Green
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
            try { $Connection.Close() } catch {}
            try { $Connection.Dispose() } catch {}
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

    Write-Host 'This database operation is READ ONLY.' -ForegroundColor Green
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
            try { $Connection.Close() } catch {}
            try { $Connection.Dispose() } catch {}
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

        return [pscustomobject]@{
            Start = $Today.AddDays(-($Days - 1))
            End = $Today
            Description = ('last {0} day(s)' -f $Days)
        }
    }

    if ($Choice -eq '3') {
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
            Write-Host 'Invalid date.' -ForegroundColor Yellow        }
    }

    if ($EndDate -lt $StartDate) {
        $Temp = $StartDate
        $StartDate = $EndDate
        $EndDate = $Temp
    }

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
        [Parameter(Mandatory)][string]$FirstLine,
        [Parameter(Mandatory)][string]$BlockText
    )

    $Lower = $BlockText.ToLowerInvariant()

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
            $Info = Get-CseGraphErrorInfo -FirstLine $FirstLine -BlockText $BlockText
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
    Write-Host 'Database access is READ ONLY.' -ForegroundColor Green
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
            try { $Connection.Close() } catch {}
            try { $Connection.Dispose() } catch {}
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
        [Parameter(Mandatory)][string]$Section,        [Parameter(Mandatory)][string]$Check,
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
    catch {}

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
        catch {}
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
        catch {}
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

        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Physical memory used' -Value ('{0:N2}%' -f $UsedPct) -Status $Status
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

function Get-IxmHealthServices {
    param([Parameter(Mandatory)]$Findings)

    $Section = 'Services'

    $ServiceChecks = @(
        [pscustomobject]@{ Label='SQL Anywhere'; Pattern='SQL Anywhere*'; Required=$true },
        [pscustomobject]@{ Label='DBWatcher'; Pattern='DBWatcher'; Required=$true },
        [pscustomobject]@{ Label='UCArchiver'; Pattern='UCArchiver'; Required=$true }
    )

    foreach ($Check in $ServiceChecks) {
        try {
            $Services = @(
                Get-Service -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -like $Check.Pattern -or $_.Name -like $Check.Pattern }
            )

            if ($Services.Count -eq 0) {
                $Status = if ($Check.Required) { 'ATTENTION' } else { 'NOT APPLICABLE' }
                Add-IxmHealthFinding -List $Findings -Section $Section -Check $Check.Label -Value 'Not found' -Status $Status
                continue
            }

            $Stopped = @($Services | Where-Object { $_.Status -ne 'Running' })
            if ($Stopped.Count -eq 0) {
                Add-IxmHealthFinding -List $Findings -Section $Section -Check $Check.Label -Value ('{0} service(s) running' -f $Services.Count) -Status 'OK'
            }
            else {
                Add-IxmHealthFinding -List $Findings -Section $Section -Check $Check.Label -Value ('{0} non-running' -f $Stopped.Count) -Status 'ATTENTION' -Details (($Stopped | ForEach-Object { '{0}={1}' -f $_.DisplayName,$_.Status }) -join '; ')
            }
        }
        catch {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check $Check.Label -Value 'Check failed' -Status 'WARNING' -Details $_.Exception.Message
        }
    }

    try {
        $StartPending = @(
            Get-Service -ErrorAction SilentlyContinue |
            Where-Object { $_.Status -eq 'StartPending' }
        )

        if ($StartPending.Count -eq 0) {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Services stuck StartPending' -Value '0' -Status 'OK'
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Services stuck StartPending' -Value ($StartPending.Count.ToString()) -Status 'ATTENTION' -Details (($StartPending | Select-Object -ExpandProperty DisplayName) -join '; ')
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Services stuck StartPending' -Value 'Check failed' -Status 'INFO'
    }

    try {
        $VoiceService = Get-Service -Name 'UCVoiceServer' -ErrorAction SilentlyContinue
        if ($null -ne $VoiceService) {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Local IXM role' -Value 'Voice Server' -Status 'INFO' -Details ('UCVoiceServer service state: {0}' -f $VoiceService.Status)
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Local IXM role' -Value 'Not identified as Voice Server' -Status 'INFO' -Details 'UCVoiceServer service is not installed on this Windows host.'
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Local IXM role' -Value 'Unable to determine' -Status 'INFO'
    }

    try {
        $UcServices = @(
            Get-Service -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like 'UC *' }
        )

        if ($UcServices.Count -gt 0) {
            $Running = @($UcServices | Where-Object { $_.Status -eq 'Running' })
            $Stopped = @($UcServices | Where-Object { $_.Status -eq 'Stopped' })

            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'UC services' -Value ('{0} running / {1} stopped / {2} total' -f $Running.Count,$Stopped.Count,$UcServices.Count) -Status 'INFO' -Details 'Stopped UC services can be role-dependent; review rather than assuming every stopped service is a fault.'

            if ($Stopped.Count -gt 0) {
                Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Stopped UC service list' -Value ($Stopped.Count.ToString()) -Status 'INFO' -Details (($Stopped | Select-Object -ExpandProperty DisplayName) -join '; ')
            }
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'UC services' -Value 'No "UC *" display names found' -Status 'INFO'
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'UC services' -Value 'Check failed' -Status 'WARNING' -Details $_.Exception.Message
    }

    foreach ($Optional in @(
        [pscustomobject]@{Label='Dialogic services'; Pattern='Dialogic*'},
        [pscustomobject]@{Label='Nuance services'; Pattern='*Nuance*'},
        [pscustomobject]@{Label='RealSpeak services'; Pattern='RealSpeak*'}
    )) {
        try {
            $Services = @(
                Get-Service -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -like $Optional.Pattern -or $_.Name -like $Optional.Pattern }
            )

            if ($Services.Count -eq 0) {
                Add-IxmHealthFinding -List $Findings -Section $Section -Check $Optional.Label -Value 'Not installed' -Status 'NOT APPLICABLE'
            }
            else {
                $Stopped = @($Services | Where-Object { $_.Status -ne 'Running' })
                if ($Stopped.Count -gt 0) {
                    Add-IxmHealthFinding -List $Findings -Section $Section -Check $Optional.Label -Value ('{0} installed / {1} non-running' -f $Services.Count,$Stopped.Count) -Status 'WARNING' -Details (($Stopped | ForEach-Object { '{0}={1}' -f $_.DisplayName,$_.Status }) -join '; ')
                }
                else {
                    Add-IxmHealthFinding -List $Findings -Section $Section -Check $Optional.Label -Value ('{0} running' -f $Services.Count) -Status 'OK'
                }
            }
        }
        catch {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check $Optional.Label -Value 'Check failed' -Status 'INFO'
        }
    }

    try {
        $W3 = Get-Service -Name W3SVC -ErrorAction SilentlyContinue
        if ($null -eq $W3) {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'IIS World Wide Web service' -Value 'Not installed' -Status 'NOT APPLICABLE'
        }
        elseif ($W3.Status -eq 'Running') {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'IIS World Wide Web service' -Value 'Running' -Status 'OK'
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'IIS World Wide Web service' -Value ([string]$W3.Status) -Status 'ATTENTION'
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'IIS World Wide Web service' -Value 'Check failed' -Status 'INFO'
    }

    try {
        if (Get-Module -ListAvailable -Name WebAdministration -ErrorAction SilentlyContinue) {
            Import-Module WebAdministration -ErrorAction Stop
            $Pools = @(Get-ChildItem IIS:\AppPools -ErrorAction Stop)

            if ($Pools.Count -eq 0) {
                Add-IxmHealthFinding -List $Findings -Section $Section -Check 'IIS application pools' -Value 'No application pools found' -Status 'INFO'
            }
            else {
                $NonRunning = New-Object System.Collections.Generic.List[string]

                foreach ($Pool in $Pools) {
                    $State = (Get-WebAppPoolState -Name $Pool.Name -ErrorAction Stop).Value
                    if ($State -ne 'Started') {
                        $NonRunning.Add(('{0}={1}' -f $Pool.Name,$State))
                    }
                }

                if ($NonRunning.Count -eq 0) {
                    Add-IxmHealthFinding -List $Findings -Section $Section -Check 'IIS application pools' -Value ('{0} started / 0 stopped' -f $Pools.Count) -Status 'OK'
                }
                else {
                    Add-IxmHealthFinding -List $Findings -Section $Section -Check 'IIS application pools' -Value ('{0} non-started of {1}' -f $NonRunning.Count,$Pools.Count) -Status 'WARNING' -Details ($NonRunning -join '; ')
                }
            }
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'IIS application pools' -Value 'WebAdministration module not installed' -Status 'NOT APPLICABLE'
        }
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'IIS application pools' -Value 'Check failed' -Status 'WARNING' -Details $_.Exception.Message
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
        [Parameter(Mandatory)]$Connection
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
    $MsgResult = Invoke-IxmHealthSql -Connection $Connection -Sql 'SELECT COUNT(*) AS MessageCount FROM DBA.MESSAGES'
    if ($MsgResult.Success -and @($MsgResult.Rows).Count -gt 0) {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Messages table count' -Value ([string]$MsgResult.Rows[0].MessageCount) -Status 'INFO'
    }
    else {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Messages table count' -Value 'Unavailable' -Status 'INFO' -Details $MsgResult.Error
    }

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

    # Active MobiLink subscriptions.
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

            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Active subscriptions' -Value ('{0} subscription(s)' -f $Rows.Count) -Status 'INFO' -Details $Detail
        }
        else {
            Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Active subscriptions' -Value 'No active-subscription rows' -Status 'INFO'
        }
    }
    else {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Active subscriptions' -Value 'Unavailable / not applicable' -Status 'INFO' -Details $SubResult.Error
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

        $Last = @($DtmfLines | Select-Object -Last 5)
        $Detail = if ($Last.Count -gt 0) { $Last -join ' || ' } else { 'No DTMF buffer entries found today.' }

        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'DTMF buffer entries' -Value ($DtmfLines.Count.ToString()) -Status 'INFO' -Details $Detail
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'DTMF buffer entries' -Value 'Check failed' -Status 'INFO'
    }
}

function Get-IxmHealthDatabaseSyncLog {
    param(
        [Parameter(Mandatory)]$Findings,
        [AllowNull()][string]$UcRoot
    )

    $Section = 'Database Sync'

    if ([string]::IsNullOrWhiteSpace([string]$UcRoot)) {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Mobiclient.log' -Value 'UC root unavailable' -Status 'NOT APPLICABLE'
        return
    }

    $SyncLog = Join-Path $UcRoot 'logs\db\Mobiclient.log'

    if (-not (Test-Path -LiteralPath $SyncLog)) {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Mobiclient.log' -Value 'Not present' -Status 'NOT APPLICABLE' -Details $SyncLog
        return
    }

    try {
        $Item = Get-Item -LiteralPath $SyncLog -ErrorAction Stop
        $Count = @(
            Select-String -LiteralPath $SyncLog -SimpleMatch -Pattern 'completed processing of download stream' -ErrorAction SilentlyContinue
        ).Count

        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Mobiclient.log' -Value ('{0} completed download-stream marker(s)' -f $Count) -Status 'INFO' -Details ('Last write: {0}; file: {1}. Interpret activity in the context of the server topology.' -f $Item.LastWriteTime,$SyncLog)
    }
    catch {
        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Mobiclient.log' -Value 'Check failed' -Status 'INFO' -Details $_.Exception.Message
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
        $Matches = @(
            Select-String -LiteralPath $TodayLog -SimpleMatch -Pattern 'Mutare' -ErrorAction SilentlyContinue
        )

        if ($Matches.Count -gt 0) {
            $LastMatch = $Matches | Select-Object -Last 1
            $Detail = ('Last matching IIS entry is at line {0}. Raw request content is intentionally not echoed by the health check.' -f $LastMatch.LineNumber)
        }
        else {
            $Detail = 'No Mutare entries were found in today''s IIS log.'
        }

        Add-IxmHealthFinding -List $Findings -Section $Section -Check 'Mutare POST activity' -Value ('{0} matching IIS entr{1}' -f $Matches.Count,$(if ($Matches.Count -eq 1) { 'y' } else { 'ies' })) -Status 'INFO' -Details $Detail
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
    Write-Section 'IX Messaging System Health Check'

    Write-Host 'This health check is read-only.' -ForegroundColor Green
    Write-Host 'It does not restart services, modify IIS, change the IX Messaging database, or alter mailbox configuration.' -ForegroundColor Green
    Write-Host ''

    $Findings = New-Object System.Collections.Generic.List[object]
    $UcRoot = Resolve-IxmUcRoot

    Write-Host 'Checking Windows, resources, and IX Messaging installation...' -ForegroundColor DarkGray
    Get-IxmHealthRegistryAndSystem -Findings $Findings -UcRoot $UcRoot

    Write-Host 'Checking services and IIS...' -ForegroundColor DarkGray
    Get-IxmHealthServices -Findings $Findings

    Write-Host 'Checking IX Messaging database...' -ForegroundColor DarkGray

    $Connection = $null
    try {
        try {
            $Selected = Select-IxmDatabaseDsn
            Write-Host ('Using DSN: {0}' -f $Selected.Name) -ForegroundColor Cyan

            $Connection = Open-IxmDsnConnection -Name $Selected.Name
            Get-IxmHealthDatabase -Findings $Findings -Connection $Connection
        }
        catch {
            Add-IxmHealthFinding -List $Findings -Section 'Database' -Check 'Database health checks' -Value 'Unavailable' -Status 'WARNING' -Details $_.Exception.Message
        }

        Write-Host 'Checking current-day VServer activity...' -ForegroundColor DarkGray
        Get-IxmHealthTodayActivity -Findings $Findings

        Write-Host 'Checking MobiLink/Mobiclient database synchronization log...' -ForegroundColor DarkGray
        Get-IxmHealthDatabaseSyncLog -Findings $Findings -UcRoot $UcRoot
        Write-Host 'Checking Windows event log, fax exceptions, Mutare IIS activity, and TCP state...' -ForegroundColor DarkGray
        Get-IxmHealthVpimEvents -Findings $Findings
        Get-IxmHealthTiffConverter -Findings $Findings -UcRoot $UcRoot
        Get-IxmHealthMutare -Findings $Findings
        Get-IxmHealthTcp -Findings $Findings

        $Rows = $Findings.ToArray()
        Show-IxmHealthFindings -Findings $Rows

        Write-Host ''
        Write-Host 'Notes:' -ForegroundColor Yellow
        Write-Host '  - Stopped UC services can be role-dependent; the report lists them for review rather than assuming every stopped service is a fault.' -ForegroundColor Yellow
        Write-Host '  - "TUI you have" is counted as prompt occurrences, not unique subscriber logins.' -ForegroundColor Yellow
        Write-Host '  - ActiveSubscriptions timestamps are displayed for review because normal behavior depends on the IX Messaging topology.' -ForegroundColor Yellow
        Write-Host '  - CPU, memory, disk, license-expiration, and loop-count thresholds are diagnostic thresholds used by this tool, not Avaya support policy.' -ForegroundColor Yellow

        Export-ResultSet -Data $Rows -BaseName 'ixm_system_health'
    }
    finally {
        if ($null -ne $Connection) {
            try { $Connection.Close() } catch {}
            try { $Connection.Dispose() } catch {}
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
    Write-Host ' 14. IX Messaging system health check'
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
            '0' { }
            default {
                Write-Host 'Invalid selection.' -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
        }
    }
    catch {
        Write-Host ''
        Write-Host ('ERROR: {0}' -f $_.Exception.Message) -ForegroundColor Red
        Pause-Tool
    }

} until ($Choice -eq '0')

Write-Host 'Goodbye.' -ForegroundColor Cyan