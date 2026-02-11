# ============================================================
# Paperless-ngx Document Uploader
# Place this script on each remote site scanner PC
# It watches a folder for new files and uploads them to
# Paperless-ngx via the REST API over the VPN
# ============================================================

# --- CONFIGURATION (change per site) ---
$PaperlessUrl    = "http://192.168.x.x:8010"     # Paperless server IP over VPN
$ApiToken        = "YOUR_API_TOKEN_HERE"         # Per-site API token from Paperless
$SiteName        = "Site-Birmingham"             # Tag for this site (created in Paperless first)
$WatchFolder     = "C:\PaperlessScans"           # Folder where scanner saves files
$ArchiveFolder   = "C:\PaperlessScans\Uploaded"  # Successfully uploaded files move here
$FailedFolder    = "C:\PaperlessScans\Failed"    # Failed uploads move here
$LogFile         = "C:\PaperlessScans\upload.log"
$FileTypes       = @("*.pdf", "*.png", "*.jpg", "*.jpeg", "*.tiff", "*.tif")
$MaxRetries      = 3
$RetryDelaySec   = 10
$PollIntervalSec = 15  # How often to check for new files
# --- END CONFIGURATION ---

# Create folders if they don't exist
foreach ($folder in @($WatchFolder, $ArchiveFolder, $FailedFolder)) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $entry
    Write-Host $entry
}

function Test-PaperlessConnection {
    try {
        $headers = @{ "Authorization" = "Token $ApiToken" }
        $response = Invoke-RestMethod -Uri "$PaperlessUrl/api/" -Headers $headers -Method Get -TimeoutSec 10
        return $true
    }
    catch {
        Write-Log "Cannot reach Paperless at $PaperlessUrl - $_" "ERROR"
        return $false
    }
}

function Get-SiteTagId {
    # Look up the tag ID for this site name
    try {
        $headers = @{ "Authorization" = "Token $ApiToken" }
        $response = Invoke-RestMethod -Uri "$PaperlessUrl/api/tags/?name__iexact=$SiteName" -Headers $headers -Method Get
        if ($response.count -gt 0) {
            return $response.results[0].id
        }
        else {
            Write-Log "Tag '$SiteName' not found in Paperless. Create it first." "WARN"
            return $null
        }
    }
    catch {
        Write-Log "Failed to look up site tag: $_" "ERROR"
        return $null
    }
}

function Upload-Document {
    param(
        [string]$FilePath,
        [int]$TagId
    )

    $fileName = Split-Path $FilePath -Leaf

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            Write-Log "Uploading '$fileName' (attempt $attempt/$MaxRetries)..."

            # Build multipart form data
            $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
            $boundary = [System.Guid]::NewGuid().ToString()

            $bodyLines = @()
            # Document file
            $bodyLines += "--$boundary"
            $bodyLines += "Content-Disposition: form-data; name=`"document`"; filename=`"$fileName`""
            $bodyLines += "Content-Type: application/octet-stream"
            $bodyLines += ""

            # Tag (if available)
            $tagPart = ""
            if ($TagId) {
                $tagPart = "--$boundary`r`nContent-Disposition: form-data; name=`"tags`"`r`n`r`n$TagId`r`n"
            }

            $enc = [System.Text.Encoding]::GetEncoding("iso-8859-1")
            $headerBytes = $enc.GetBytes(($bodyLines -join "`r`n") + "`r`n")
            $footerText = "`r`n$tagPart--$boundary--`r`n"
            $footerBytes = $enc.GetBytes($footerText)

            $bodyBytes = New-Object byte[] ($headerBytes.Length + $fileBytes.Length + $footerBytes.Length)
            [System.Buffer]::BlockCopy($headerBytes, 0, $bodyBytes, 0, $headerBytes.Length)
            [System.Buffer]::BlockCopy($fileBytes, 0, $bodyBytes, $headerBytes.Length, $fileBytes.Length)
            [System.Buffer]::BlockCopy($footerBytes, 0, $bodyBytes, $headerBytes.Length + $fileBytes.Length, $footerBytes.Length)

            $headers = @{
                "Authorization" = "Token $ApiToken"
            }

            $response = Invoke-RestMethod `
                -Uri "$PaperlessUrl/api/documents/post_document/" `
                -Method Post `
                -Headers $headers `
                -ContentType "multipart/form-data; boundary=$boundary" `
                -Body $bodyBytes `
                -TimeoutSec 120

            Write-Log "Successfully uploaded '$fileName' (Task: $response)" "INFO"

            # Move to archive folder
            $archivePath = Join-Path $ArchiveFolder $fileName
            # Handle duplicate filenames
            if (Test-Path $archivePath) {
                $base = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
                $ext = [System.IO.Path]::GetExtension($fileName)
                $archivePath = Join-Path $ArchiveFolder "$base`_$(Get-Date -Format 'yyyyMMdd_HHmmss')$ext"
            }
            Move-Item -Path $FilePath -Destination $archivePath -Force
            Write-Log "Archived to '$archivePath'"
            return $true
        }
        catch {
            Write-Log "Upload failed for '$fileName': $_" "ERROR"
            if ($attempt -lt $MaxRetries) {
                Write-Log "Retrying in $RetryDelaySec seconds..."
                Start-Sleep -Seconds $RetryDelaySec
            }
        }
    }

    # All retries exhausted - move to failed folder
    Write-Log "All retries exhausted for '$fileName'. Moving to failed folder." "ERROR"
    $failedPath = Join-Path $FailedFolder $fileName
    if (Test-Path $failedPath) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        $ext = [System.IO.Path]::GetExtension($fileName)
        $failedPath = Join-Path $FailedFolder "$base`_$(Get-Date -Format 'yyyyMMdd_HHmmss')$ext"
    }
    Move-Item -Path $FilePath -Destination $failedPath -Force
    return $false
}

# ============================================================
# MAIN LOOP
# ============================================================
Write-Log "========================================="
Write-Log "Paperless Uploader started"
Write-Log "Watching: $WatchFolder"
Write-Log "Server:   $PaperlessUrl"
Write-Log "Site:     $SiteName"
Write-Log "========================================="

# Verify connection on startup
if (-not (Test-PaperlessConnection)) {
    Write-Log "Initial connection failed. Will keep retrying..." "WARN"
}

# Get site tag ID
$siteTagId = Get-SiteTagId

while ($true) {
    try {
        # Find all matching files
        $files = @()
        foreach ($type in $FileTypes) {
            $files += Get-ChildItem -Path $WatchFolder -Filter $type -File -ErrorAction SilentlyContinue
        }

        if ($files.Count -gt 0) {
            Write-Log "Found $($files.Count) file(s) to process"

            foreach ($file in $files) {
                # Wait briefly to ensure file is fully written
                Start-Sleep -Seconds 2

                # Check file is not locked (still being written)
                try {
                    $stream = [System.IO.File]::Open($file.FullName, 'Open', 'Read', 'None')
                    $stream.Close()
                }
                catch {
                    Write-Log "File '$($file.Name)' is locked, skipping for now" "WARN"
                    continue
                }

                Upload-Document -FilePath $file.FullName -TagId $siteTagId
            }
        }
    }
    catch {
        Write-Log "Error in main loop: $_" "ERROR"
    }

    Start-Sleep -Seconds $PollIntervalSec
}