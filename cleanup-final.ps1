# ============================================================
# Script: site-cleanup-pro-enterprise-v6.4.1.ps1
# Purpose: Unified SharePoint Cleanup (Lists + Libraries)
# Version: 6.4.1 - Logs every deleted item/file
# Author: Venkat's Optimized Edition ⚡
# ============================================================

[CmdletBinding()]
param(
    [switch]$DryRun,
    [string]$ConfigFile = "cleanup-config.json",
    [switch]$UseRecycleBin,
    [int]$MaxRetries = 3,
    [int]$ThrottleDelaySeconds = 30
)

# ============================================================
# Configuration
# ============================================================

$script:config = @{
    MasterSiteUrl   = "https://.sharepoint.com/sites/Dev"
    MasterListTitle = "Cleanup Site Lists Config"
    DaysThreshold   = 1
    CustomDateField = "Stamped"
    PageSize        = 2000
    MaxPages        = 100
    BatchDeleteSize = 30
    RequestsPerMinute = 60
    BaseVerificationDelay = 10
    LargeVerificationDelay = 15
    ExtraLargeVerificationDelay = 20
}

# Load from config JSON if provided
if (Test-Path $ConfigFile) {
    Write-Host "📋 Loading configuration from $ConfigFile" -ForegroundColor Cyan
    try {
        $fileConfig = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        foreach ($key in $fileConfig.PSObject.Properties.Name) {
            $script:config[$key] = $fileConfig.$key
        }
    } catch {
        Write-Host "⚠️ Could not load config file, using defaults" -ForegroundColor Yellow
    }
}

$cutoffDate = (Get-Date).AddDays(-$script:config.DaysThreshold)

# ============================================================
# Logging setup
# ============================================================

$scriptPath = $PSScriptRoot
if (-not $scriptPath) { $scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
if (-not $scriptPath) { $scriptPath = Get-Location }

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDir = Join-Path $scriptPath "sp-cleanup-logs"
New-Item -ItemType Directory -Path $logDir -Force | Out-Null

$logFile = Join-Path $logDir "cleanup-$timestamp.csv"
$errorLog = Join-Path $logDir "errors-$timestamp.log"
$transcriptLog = Join-Path $logDir "transcript-$timestamp.txt"
$batchTempDir = Join-Path $logDir "batch-temp-$timestamp"
New-Item -ItemType Directory -Path $batchTempDir -Force | Out-Null

$script:totalProcessed = 0
$script:totalDeleted = 0
$script:totalErrors = 0
$script:verificationFailures = 0

Start-Transcript -Path $transcriptLog -Append
"SiteUrl,ListOrLibrary,ItemType,ItemPath,AbsoluteUrl,Action" | Out-File $logFile -Encoding utf8

Write-Host "`n🚀 SharePoint Cleanup v6.4.1 - Unified Mode" -ForegroundColor Cyan
Write-Host "Cutoff Date: $($cutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Green
Write-Host "Batch Size: $($script:config.BatchDeleteSize)" -ForegroundColor Green

# ============================================================
# Helper Functions
# ============================================================

function Write-AuditLog {
    param([string]$Level, [string]$Message)
    $entry = "$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ss') [$Level] $Message"
    Add-Content -Path $errorLog -Value $entry -ErrorAction SilentlyContinue
    $color = switch ($Level) {
        "ERROR"   { "Red" }
        "WARNING" { "Yellow" }
        "SUCCESS" { "Green" }
        default   { "White" }
    }
    Write-Host $entry -ForegroundColor $color
}

function Get-VerificationDelay {
    param([int]$Count)
    if ($Count -ge 45) { return $script:config.ExtraLargeVerificationDelay }
    elseif ($Count -ge 35) { return $script:config.LargeVerificationDelay }
    elseif ($Count -ge 25) { return 12 }
    else { return $script:config.BaseVerificationDelay }
}

function Invoke-WithRetry {
    param([scriptblock]$Action)
    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try { return & $Action }
        catch {
            $attempt++
            Start-Sleep -Seconds ($ThrottleDelaySeconds * $attempt)
        }
    }
}

function Ensure-M365Connection {
    Write-Host "`n🔐 Checking Microsoft 365 connection..." -ForegroundColor Cyan
    try {
        $status = m365 status --output json | ConvertFrom-Json
        if (-not $status.connectedAs) { throw "Not connected" }
        Write-Host "✅ Connected as $($status.connectedAs)" -ForegroundColor Green
    } catch {
        throw "Not authenticated. Run 'm365 login' first."
    }
}

# ============================================================
# CAML Query for both Lists & Libraries
# ============================================================

function Get-OldListItemsCAML {
    param(
        [string]$SiteUrl,
        [string]$ListTitle,
        [string]$CustomDateField,
        [datetime]$CutoffDate
    )

    $dateStr = $CutoffDate.ToString("yyyy-MM-ddTHH:mm:ss")
    $pageSize = $script:config.PageSize
    $maxPages = $script:config.MaxPages
    $allItems = @()
    $lastId = 0
    $page = 0

    Write-Host "🔍 Fetching items/files older than $dateStr ..." -ForegroundColor Cyan

    while ($page -lt $maxPages) {
        $page++
        $caml = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <And>
        <Gt><FieldRef Name='ID'/><Value Type='Counter'>$lastId</Value></Gt>
        <Lt><FieldRef Name='$CustomDateField'/><Value Type='DateTime' IncludeTimeValue='TRUE'>$dateStr</Value></Lt>
      </And>
    </Where>
    <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
  </Query>
  <RowLimit>$pageSize</RowLimit>
</View>
"@

        try {
            $json = Invoke-WithRetry { m365 spo listitem list --webUrl $SiteUrl --listTitle "$ListTitle" --camlQuery "$caml" --output json }
            if (-not $json) { break }
            $batch = $json | ConvertFrom-Json
            if (-not $batch -or $batch.Count -eq 0) { break }

            # Exclude folders
            $filtered = $batch | Where-Object { $_.FileSystemObjectType -ne 1 }
            $allItems += $filtered
            $lastId = ($batch | Measure-Object -Property Id -Maximum).Maximum
            Write-Host "📦 Page $page : $($filtered.Count) deletable items (Total: $($allItems.Count))" -ForegroundColor Gray
        }
        catch {
            Write-AuditLog "ERROR" "CAML query failed on $ListTitle (page $page): $($_.Exception.Message)"
            break
        }
    }

    return $allItems
}

# ============================================================
# Batch Deletion + Logging
# ============================================================

function Remove-ListItemsInBatch {
    param(
        [string]$SiteUrl,
        [string]$ListTitle,
        [array]$Items,
        [switch]$DryRun
    )

    $Items = @($Items)
    $total = $Items.Count
    if ($total -eq 0) { Write-Host "✅ No items to delete" -ForegroundColor Green; return }

    Write-Host "`n🗂️ Processing $ListTitle ($total items)" -ForegroundColor Cyan
    $list = (m365 spo list get --webUrl $SiteUrl --title $ListTitle --output json) | ConvertFrom-Json
    if (-not $list) { Write-AuditLog "ERROR" "Cannot read list metadata for $ListTitle at $SiteUrl"; return }

    $listId = $list.Id
    $itemType = if ($list.BaseTemplate -eq 101) { "File" } else { "ListItem" }
    $batchSize = [int]$script:config.BatchDeleteSize
    $batches = [math]::Ceiling($total / $batchSize)

    for ($b = 0; $b -lt $batches; $b++) {
        $start = [int]($b * $batchSize)
        $end = [int][math]::Min($start + $batchSize - 1, $total - 1)
        if ($start -gt $end) { continue }

        $chunk = if ($start -eq $end) { @($Items[$start]) } else { $Items[$start..$end] }
        $ids = $chunk | ForEach-Object { $_.Id }
        $csvFile = Join-Path $batchTempDir "batch-$($b+1).csv"

        "ID" | Out-File $csvFile -Encoding utf8NoBOM -Force
        $ids | Out-File $csvFile -Encoding utf8NoBOM -Append

        if ($DryRun) {
            foreach ($item in $chunk) {
                $path = $item.ListItemAllFields.FileRef ?? $item.FileRef ?? "ID:$($item.Id)"
                $absUrl = "$SiteUrl$path"
                "$SiteUrl,$ListTitle,$itemType,$path,$absUrl,DryRun" | Out-File $logFile -Append -Encoding utf8
            }
            Write-Host "🧪 [DRY RUN] Would delete $($ids.Count) $itemType(s)" -ForegroundColor Yellow
            continue
        }

        try {
            Write-Host "🗑️ Batch $($b+1)/$batches - $($ids.Count) items" -ForegroundColor White
            if ($UseRecycleBin) {
                m365 spo listitem batch remove --webUrl $SiteUrl --listId $listId --filePath $csvFile --recycle --force | Out-Null
            } else {
                m365 spo listitem batch remove --webUrl $SiteUrl --listId $listId --filePath $csvFile --force | Out-Null
            }

            $delay = Get-VerificationDelay -Count $ids.Count
            Start-Sleep -Seconds $delay

            foreach ($item in $chunk) {
                $path = $item.ListItemAllFields.FileRef ?? $item.FileRef ?? "ID:$($item.Id)"
                $absUrl = "$SiteUrl$path"
                "$SiteUrl,$ListTitle,$itemType,$path,$absUrl,Deleted" | Out-File $logFile -Append -Encoding utf8
            }

            $script:totalDeleted += $ids.Count
            Write-Host "✅ Deleted batch $($b+1)" -ForegroundColor Green

        } catch {
            Write-AuditLog "ERROR" "Batch delete failed for $ListTitle (batch $($b+1)): $($_.Exception.Message)"
            $script:totalErrors++
            foreach ($item in $chunk) {
                $path = $item.ListItemAllFields.FileRef ?? $item.FileRef ?? "ID:$($item.Id)"
                $absUrl = "$SiteUrl$path"
                "$SiteUrl,$ListTitle,$itemType,$path,$absUrl,Failed" | Out-File $logFile -Append -Encoding utf8
            }
        }
    }

    # Final sweep
    Write-Host "`n🔁 Final sweep check..." -ForegroundColor Cyan
    $remaining = Get-OldListItemsCAML -SiteUrl $SiteUrl -ListTitle $ListTitle -CustomDateField $script:config.CustomDateField -CutoffDate $cutoffDate
    if ($remaining.Count -gt 0 -and -not $DryRun) {
        Write-Host "🧹 Deleting remaining $($remaining.Count) files..." -ForegroundColor Yellow
        Remove-ListItemsInBatch -SiteUrl $SiteUrl -ListTitle $ListTitle -Items @($remaining) -DryRun:$DryRun
    } else {
        Write-Host "✅ No remaining items" -ForegroundColor Green
    }
}

# ============================================================
# MAIN EXECUTION
# ============================================================

try {
    Ensure-M365Connection
    $targets = (m365 spo listitem list --webUrl $script:config.MasterSiteUrl --listTitle $script:config.MasterListTitle --output json) | ConvertFrom-Json
    if (-not $targets) { Write-Host "No cleanup targets found"; exit }

    $grouped = $targets | Group-Object -Property SiteURL
    $startTime = Get-Date

    foreach ($group in $grouped) {
        $siteUrl = $group.Name
        Write-Host "`n🌐 Site: $siteUrl" -ForegroundColor Cyan

        foreach ($target in $group.Group) {
            $listTitle = $target.ListOrLibraryName
            Write-Host "`n📁 Target: $listTitle" -ForegroundColor White

            $list = (m365 spo list get --webUrl $siteUrl --title $listTitle --output json) | ConvertFrom-Json
            if (-not $list) { Write-AuditLog "ERROR" "Cannot access $listTitle"; continue }

            if ($list.BaseTemplate -eq 101) {
                Write-Host "📦 Document Library Detected - Recursive CAML Cleanup" -ForegroundColor Cyan
            } else {
                Write-Host "📋 List Detected" -ForegroundColor Green
            }

            $items = Get-OldListItemsCAML -SiteUrl $siteUrl -ListTitle $listTitle -CustomDateField $script:config.CustomDateField -CutoffDate $cutoffDate
            if ($items.Count -eq 0) { Write-Host "✅ No old items/files to delete" -ForegroundColor Green; continue }

            Remove-ListItemsInBatch -SiteUrl $siteUrl -ListTitle $listTitle -Items $items -DryRun:$DryRun
        }
    }

    $duration = (Get-Date) - $startTime
    Write-Host "`n✅ Cleanup complete in $([math]::Round($duration.TotalMinutes, 2)) minutes" -ForegroundColor Green
    Write-AuditLog "SUCCESS" "Cleanup finished. Total deleted: $($script:totalDeleted)"

} finally {
    Stop-Transcript
    if (Test-Path $batchTempDir) { Remove-Item -Recurse -Force $batchTempDir -ErrorAction SilentlyContinue }
    Write-Host "`n📜 Logs saved in: $logDir" -ForegroundColor Cyan
}
