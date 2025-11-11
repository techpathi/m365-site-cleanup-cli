# ============================================================
# Script: populate-dummy-files-nested.ps1
# Purpose: Upload dummy files into nested folders and set custom date field
# ============================================================

# --- Configuration ---
$siteUrl = "https://.sharepoint.com/sites/Dev"
$libraryName = "Cleanup Test v2"      # Must be a document library
$rootFolder = "TestData"        # Folder under the library where everything goes
$folderDepth = 10                # Number of nested folder levels
$foldersPerLevel = 10            # How many folders per level
$filesPerFolder = 100             # Number of files per folder
$customField = "Stamped"  # Custom date column to simulate age
$daysRange = 90                # Random date range (past X days)
$tempFolder = "sp-dummy"

Write-Host "🔐 Logging in to Microsoft 365..."
#m365 login --authType deviceCode

# --- Ensure local temp folder ---
if (-not (Test-Path $tempFolder)) {
    New-Item -ItemType Directory -Path $tempFolder | Out-Null
}

Write-Host "🚀 Preparing nested folder structure in $libraryName/$rootFolder..."

# --- Recursive function to create folders and upload files ---
function Add-DummyFilesInFolder($parentPath, $level) {
    if ($level -gt $folderDepth) { return }

    for ($i = 1; $i -le $foldersPerLevel; $i++) {
        $folderName = "Folder_Level${level}_$i"
        $currentPath = "$parentPath/$folderName"

        Write-Host "📁 Creating folder: $currentPath"
        m365 spo folder add --webUrl $siteUrl --parentFolderUrl "$parentPath" --name "$folderName" --output none

        # --- Upload dummy files ---
        for ($j = 1; $j -le $filesPerFolder; $j++) {
            $fileName = "Dummy_${level}_${i}_${j}.txt"
            $filePath = "$tempFolder/$fileName"
            "This is dummy file L$level F$i N$j" | Out-File -FilePath $filePath -Encoding utf8

            # Construct the full server-relative folder path
            # Example: /sites/TestSite/Shared Documents/TestData/Folder_Level1_1
            $serverRelativeFolder = "/sites/$(($siteUrl.Split('/')[-1]))/$currentPath"

            Write-Host "📤 Uploading file: $filePath -> $fileName → $serverRelativeFolder"
            try {
                m365 spo file add --webUrl $siteUrl --folder $serverRelativeFolder --path $filePath --output none
            } catch {
                Write-Host "⚠️ Failed to upload $fileName"
                continue
            }

            # Randomize custom date
            $daysOld = Get-Random -Minimum 1 -Maximum $daysRange
            $dummyDate = (Get-Date).AddDays(-$daysOld).ToString("yyyy-MM-dd HH:mm:ss")

            # Lookup uploaded file by FileLeafRef
            $itemsJson = m365 spo listitem list --webUrl $siteUrl --listTitle $libraryName --output json --filter "FileLeafRef eq '$fileName'"
            $item = ($itemsJson | ConvertFrom-Json)[0]

            if ($item) {
                m365 spo listitem set --webUrl $siteUrl --listTitle $libraryName --id $item.Id --Stamped $dummyDate --output none
                Write-Host "🕒 Set $customField = $dummyDate"
            }
        }

        # Recurse deeper
        Add-DummyFilesInFolder $currentPath ($level + 1)
    }
}

# --- Start process ---
Write-Host "`n🌲 Creating folder tree starting at: $libraryName/$rootFolder"
m365 spo folder add --webUrl $siteUrl --parentFolderUrl "$libraryName" --name "$rootFolder" --output none
Add-DummyFilesInFolder "$libraryName/$rootFolder" 1

Write-Host "`n✅ Nested folder and file creation complete!"
