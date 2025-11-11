# ============================================================
# Script: populate-dummy-listitems.ps1
# Purpose: Create dummy list items with random Title and date field
# ============================================================

# --- Configuration ---
$siteUrl = "https://.sharepoint.com/sites/Dev2"
$listName = "Cleanup Test List"   # Must be a custom list
$itemCount = 10000                  # Total number of dummy items to create
$customField = "Stamped"          # Custom date column name
$daysRange = 90                   # Random date range (past X days)

Write-Host "🔐 Logging in to Microsoft 365..."
#m365 login --authType deviceCode

Write-Host "🚀 Creating $itemCount dummy items in list '$listName'..."

# --- Function to generate a random title ---
function Get-RandomTitle {
    $prefixes = @("Alpha", "Beta", "Gamma", "Delta", "Omega", "Zeta", "Nova", "Luna", "Apex", "Zenith")
    $suffixes = @("One", "Prime", "Force", "Edge", "Flow", "Wave", "Core", "Link", "Shift", "Spark")
    $prefix = Get-Random -InputObject $prefixes
    $suffix = Get-Random -InputObject $suffixes
    $num = Get-Random -Minimum 100 -Maximum 999
    return "$prefix-$suffix-$num"
}

# --- Loop to create dummy items ---
for ($i = 1; $i -le $itemCount; $i++) {
    $title = Get-RandomTitle
    $daysOld = Get-Random -Minimum 1 -Maximum $daysRange
    $dummyDate = (Get-Date).AddDays(-$daysOld).ToString("yyyy-MM-dd HH:mm:ss")

    Write-Host "📝 Creating item #$i → Title: '$title', '$customField': '$dummyDate'"

    try {
        # Create item with Title first
        $itemJson = m365 spo listitem add --webUrl $siteUrl --listTitle "$listName" --Title "$title" --Stamped $dummyDate --output json
        $item = $itemJson | ConvertFrom-Json

        Write-Host "   ✅ Created item ID: $($item.Id)"
    }
    catch {
        Write-Host "⚠️ Failed to create item #$i"
    }
}

Write-Host "`n✅ Dummy list item creation complete!"
