#Requires -Version 7.0

$ErrorActionPreference = "Continue"

#=========================================
# Validate EXO Module
#=========================================

if (-not (Get-Module -ListAvailable ExchangeOnlineManagement))
{
    Write-Host ""
    Write-Host "ExchangeOnlineManagement module not found." -ForegroundColor Red
    Write-Host ""
    Write-Host "Run:"
    Write-Host "Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    Write-Host ""
    return
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

#=========================================
# Date Parser
#=========================================

function Convert-ToDateTime {

    param(
        [string]$InputDate
    )

    $InputDate = $InputDate.Trim()

    if($InputDate -match '^(\d{2})-(\d{2})-(\d{4})$')
    {
        return (
            Get-Date `
                -Year $matches[3] `
                -Month $matches[2] `
                -Day $matches[1] `
                -Hour 0 `
                -Minute 0 `
                -Second 0
        )
    }

    if($InputDate -match '^(\d{2})-(\d{2})-(\d{4})\s(\d{2}):(\d{2})$')
    {
        return (
            Get-Date `
                -Year $matches[3] `
                -Month $matches[2] `
                -Day $matches[1] `
                -Hour $matches[4] `
                -Minute $matches[5] `
                -Second 0
        )
    }

    throw "Invalid format. Use DD-MM-YYYY or DD-MM-YYYY HH:MM"
}

#=========================================
# Message Trace Query Function
#=========================================

function Invoke-TraceQuery {

    param(
        [datetime]$StartDate,
        [datetime]$EndDate,
        [int]$ResultSize = 1000
    )

    $RetryDelays = @(15,30,60)

    foreach($Delay in $RetryDelays)
    {
        try
        {
            return Get-MessageTraceV2 `
                -StartDate $StartDate `
                -EndDate $EndDate `
                -ResultSize $ResultSize `
                -ErrorAction Stop
        }
        catch
        {
            Write-Warning $_.Exception.Message
            Write-Host "Retrying in $Delay seconds..."

            Start-Sleep -Seconds $Delay
        }
    }

    return $null
}

#=========================================
# Input Banner
#=========================================

Clear-Host

Write-Host ""
Write-Host "======================================================"
Write-Host " Exchange Online Message Trace Report"
Write-Host "======================================================"
Write-Host ""
Write-Host "Maximum Lookup Duration : 10 Days"
Write-Host ""
Write-Host "Examples:"
Write-Host ""
Write-Host "20-08-2026"
Write-Host "20-08-2026 06:00"
Write-Host ""
Write-Host "If time is omitted, 00:00 is assumed."
Write-Host ""

#=========================================
# Read Start Date
#=========================================

do {

    try {

        $StartInput = Read-Host "Enter Start Date"

        $StartDate = Convert-ToDateTime $StartInput

        $ValidStart = $true
    }
    catch {

        Write-Warning $_.Exception.Message
        $ValidStart = $false
    }

} until ($ValidStart)

#=========================================
# Read End Date
#=========================================

do {

    try {

        $EndInput = Read-Host "Enter End Date"

        $EndDate = Convert-ToDateTime $EndInput

        if($EndInput -match '^\d{2}-\d{2}-\d{4}$')
        {
            $EndDate = $EndDate.AddDays(1)
        }

        $ValidEnd = $true
    }
    catch {

        Write-Warning $_.Exception.Message
        $ValidEnd = $false
    }

} until ($ValidEnd)

#=========================================
# Validation
#=========================================

if($EndDate -le $StartDate)
{
    Write-Error "End Date must be greater than Start Date."
    return
}

if(($EndDate - $StartDate).TotalDays -gt 10)
{
    Write-Error "Get-MessageTraceV2 supports only 10 days per query."
    return
}

#=========================================
# Connect Exchange Online
#=========================================

try
{
    Connect-ExchangeOnline -ShowBanner:$false
}
catch
{
    Write-Error "Failed to connect to Exchange Online."
    return
}

#=========================================
# Settings
#=========================================

$WindowMinutes = 10
$ResultSize = 1000

$AllResults = [System.Collections.Generic.List[object]]::new()
$StillLimited = @()

#=========================================
# Collection
#=========================================

$currentStart = $StartDate

while ($currentStart -lt $EndDate)
{
    $currentEnd = $currentStart.AddMinutes($WindowMinutes)

    if($currentEnd -gt $EndDate)
    {
        $currentEnd = $EndDate
    }

    Write-Host ""
    Write-Host "Processing: $currentStart --> $currentEnd"

    $Results = Invoke-TraceQuery `
        -StartDate $currentStart `
        -EndDate $currentEnd `
        -ResultSize $ResultSize

    if($null -eq $Results)
    {
        $currentStart = $currentEnd
        continue
    }

    if(($Results | Measure-Object).Count -lt 1000)
    {
        foreach($Item in $Results)
        {
            $AllResults.Add($Item)
        }
    }
    else
    {
        Write-Warning "1000 rows detected. Splitting into 15-minute windows."

        $subStart = $currentStart

        while($subStart -lt $currentEnd)
        {
            $subEnd = $subStart.AddMinutes(15)

            if($subEnd -gt $currentEnd)
            {
                $subEnd = $currentEnd
            }

            $SubResults = Invoke-TraceQuery `
                -StartDate $subStart `
                -EndDate $subEnd `
                -ResultSize $ResultSize

            if($null -ne $SubResults)
            {
                foreach($Item in $SubResults)
                {
                    $AllResults.Add($Item)
                }

                if(($SubResults | Measure-Object).Count -ge 1000)
                {
                    $StillLimited += [PSCustomObject]@{
                        StartTime = $subStart
                        EndTime   = $subEnd
                        Records   = ($SubResults | Measure-Object).Count
                    }
                }
            }

            $subStart = $subEnd
        }
    }

    $currentStart = $currentEnd
}

#=========================================
# Deduplicate
#=========================================

$AllResults = $AllResults |
Sort-Object `
    MessageTraceId,
    RecipientAddress,
    Received -Unique

#=========================================
# Build Report
#=========================================

$Report = $AllResults |
Group-Object {

    "{0}|{1}|{2}" -f `
    $_.MessageTraceId,
    $_.SenderAddress,
    [string]$_.Subject

} |
ForEach-Object {

    $Recipients = $_.Group |
        Select-Object -ExpandProperty RecipientAddress -Unique

    if($Recipients.Count -gt 20)
    {
        $First = $_.Group | Select-Object -First 1

        [PSCustomObject]@{
            DateReceived   = $First.Received
            SenderAddress  = $First.SenderAddress
            Subject        = $First.Subject
            RecipientCount = $Recipients.Count
            Recipients     = ($Recipients -join ";")
            Status         = $First.Status
            MessageTraceId = $First.MessageTraceId
        }
    }
}

#=========================================
# Export
#=========================================

$OutputFolder = Join-Path $env:USERPROFILE "Documents\MessageTraceReports"

if(!(Test-Path $OutputFolder))
{
    New-Item `
        -ItemType Directory `
        -Path $OutputFolder `
        -Force | Out-Null
}

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

$RawFile = Join-Path $OutputFolder "RawTrace_$Timestamp.csv"
$ReportFile = Join-Path $OutputFolder "GT20Recipients_$Timestamp.csv"
$LimitFile = Join-Path $OutputFolder "PotentiallyIncompleteWindows_$Timestamp.csv"

$AllResults |
Export-Csv $RawFile `
-NoTypeInformation `
-Encoding utf8BOM

$Report |
Export-Csv $ReportFile `
-NoTypeInformation `
-Encoding utf8BOM

if($StillLimited.Count -gt 0)
{
    $StillLimited |
    Export-Csv $LimitFile `
    -NoTypeInformation `
    -Encoding utf8BOM
}

#=========================================
# Summary
#=========================================

Write-Host ""
Write-Host "======================================================"
Write-Host "REPORT COMPLETE"
Write-Host "======================================================"
Write-Host ""
Write-Host "Total Trace Records : $($AllResults.Count)"
Write-Host "Messages >20 Recipients : $($Report.Count)"
Write-Host ""
Write-Host "Raw Trace:"
Write-Host $RawFile
Write-Host ""
Write-Host "Final Report:"
Write-Host $ReportFile

if($StillLimited.Count -gt 0)
{
    Write-Host ""
    Write-Host "Potentially Incomplete Windows:"
    Write-Host $LimitFile -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Investigation Confidence:"
Write-Host "- 10-minute windows used"
Write-Host "- Auto-refined to 15-minute windows if 1000 records returned"
Write-Host "- Deduplication applied"
Write-Host "- Remaining windows at 1000 rows may require manual review"
Write-Host ""

Disconnect-ExchangeOnline -Confirm:$false