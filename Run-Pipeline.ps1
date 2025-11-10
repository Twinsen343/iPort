# ====================================================================================================
# iPort: iPhone / MTP Media Automation Toolkit
# ----------------------------------------------------------------------------------------------------
# Main Pipeline Controller: Run-Pipeline.ps1
#
# This script orchestrates the full iPort workflow:
#   1. Copy photos and videos from iPhone via MTP (CopyFromIPhone.ps1)
#   2. Convert images (HEIC/HEIF → JPG/PNG) using ImageMagick (convertImages.ps1)
#   3. Convert videos (MOV/MP4 → MP4 H.264) using FFmpeg (convertVideos.ps1)
#
# Requirements:
#   - PowerShell 5.1+ or 7+
#   - iPhone connected via USB (MTP mode)
#   - ffmpeg + ffprobe available in PATH or defined in FfmpegPath
#   - ImageMagick (or compatible tool) installed
#
# Repository: https://github.com/yourname/iPort
# Author: Mason Walker
# ====================================================================================================


# ==================================================================
# Enhanced Console Output Helpers (timestamps + emojis + colors)
# ==================================================================

function Write-Info ($m) {
    # General informational message (cyan background)
    $ts = (Get-Date).ToString("HH:mm:ss")
    Write-Host "[INFO  $ts] 💬 $m" -ForegroundColor White -BackgroundColor DarkCyan
}

function Write-Success ($m) {
    # Success message (green background)
    $ts = (Get-Date).ToString("HH:mm:ss")
    Write-Host "[OK    $ts] ✅ $m" -ForegroundColor White -BackgroundColor DarkGreen
}

function Write-Warn ($m) {
    # Warning message (yellow background)
    $ts = (Get-Date).ToString("HH:mm:ss")
    Write-Host "[WARN  $ts] ⚠️  $m" -ForegroundColor Black -BackgroundColor DarkYellow
}

function Write-Err ($m) {
    # Error message (red background)
    $ts = (Get-Date).ToString("HH:mm:ss")
    Write-Host "[ERROR $ts] ❌ $m" -ForegroundColor White -BackgroundColor Red
}

function Write-VerboseMsg ($m) {
    # Verbose log (only prints when -Verbose or $VerbosePreference = 'Continue')
    if ($VerbosePreference -eq 'Continue') {
        $ts = (Get-Date).ToString("HH:mm:ss")
        Write-Host "[VERB  $ts] 🪶 $m" -ForegroundColor Gray -BackgroundColor DarkGray
    }
}

# ------------------------------------------------------------------
# Enable verbose logging globally
# ------------------------------------------------------------------
$VerbosePreference = 'Continue'

# Startup message
Write-Host "===== START: iPort iPhone Import and Conversion Pipeline =====" -ForegroundColor White -BackgroundColor Green

# Normalize $PSScriptRoot if run interactively
if (-not $PSScriptRoot) {
    $PSScriptRoot = Split-Path -LiteralPath $MyInvocation.MyCommand.Path -Parent
}

# Begin total runtime measurement
$scriptTimer = [System.Diagnostics.Stopwatch]::StartNew()


try {
    # ==================================================================
    # STEP 1: Copy from iPhone (MTP Import)
    # ==================================================================
    Write-VerboseMsg "Resolving CopyFromIPhone.ps1 path..."
    $copyScript = Join-Path $PSScriptRoot 'CopyFromIPhone.ps1'
    if (-not (Test-Path $copyScript)) { throw "Copy script not found at: $copyScript" }

    Write-Info "`n[1/3] Running CopyFromIPhone.ps1 ..."
    $copyTimer = [System.Diagnostics.Stopwatch]::StartNew()

    # Configure import parameters
    $params = @{
        Destination     = "$env:USERPROFILE\Pictures\from iphone\Raw" # Target path for raw media
        DeviceName      = 'iPhone'                                    # MTP device name as seen in Explorer
        SkipByNameOnly  = $true                                        # Skip duplicates by filename only
        NoDateFolders   = $true                                        # Store all in flat folder (no date structure)
        Progress        = $true                                        # Show live progress
        DeleteAfterCopy = $false                                       # Keep originals on device
        TimeoutSeconds  = 20                                           # File read timeout safeguard

        # Optional UI / structure output
        TextPulse       = $true
        Heartbeat       = $true
        ListStructure   = $true
        # ListOnly       = $true      # Enable if testing scan only (no copy)
    }

    # Ensure no host-injected parameters interfere (fix for PowerShell host quirk)
    if ($params.ContainsKey('ProgressAction')) { $null = $params.Remove('ProgressAction') }
    if ($PSDefaultParameterValues) { $null = $PSDefaultParameterValues.Remove('*:ProgressAction') }

    # Execute the import script
    & $copyScript @params -Verbose

    $copyTimer.Stop()
    Write-Success "CopyFromIPhone.ps1 completed in $([math]::Round($copyTimer.Elapsed.TotalSeconds,2)) seconds."


    # ==================================================================
    # STEP 2: Convert Images (HEIC/HEIF → JPG)
    # ==================================================================
    Write-VerboseMsg "Resolving convertImages.ps1 path..."
    $conv = Join-Path $PSScriptRoot 'convertImages.ps1'
    if (-not (Test-Path $conv)) { throw "Image conversion script not found at: $conv" }

    Write-Info "`n[2/3] Running convertImages.ps1 ..."
    $imgTimer = [System.Diagnostics.Stopwatch]::StartNew()

    # Run image conversion process
    & $conv `
        -InputDir  "$env:USERPROFILE\Pictures\from iphone\Raw\Images" `
        -OutputDir "$env:USERPROFILE\Pictures\from iphone\Processed\Images" `
        -Format jpg -Quality 85 `
        -MaxWidth 1920 -MaxHeight 1080 `
        -SkipExisting -Recurse `
        -ToolPath "C:\Program Files\ImageMagick-7.1.2-Q16-HDRI\magick.exe" `
        -Verbose

    $imgTimer.Stop()
    Write-Success "convertImages.ps1 completed in $([math]::Round($imgTimer.Elapsed.TotalSeconds,2)) seconds."


    # ==================================================================
    # STEP 3: Convert Videos (MOV/MP4 → MP4 H.264)
    # ==================================================================
    Write-VerboseMsg "Resolving convertVideos.ps1 path..."
    $vid = Join-Path $PSScriptRoot 'convertVideos.ps1'
    if (-not (Test-Path $vid)) { throw "Video conversion script not found at: $vid" }

    Write-Info "`n[3/3] Running convertVideos.ps1 ..."
    $vidTimer = [System.Diagnostics.Stopwatch]::StartNew()

    # Configure video conversion parameters
    $vidArgs = @{
        InputDir      = "$env:USERPROFILE\Pictures\from iphone\Raw\Videos"   # Source video folder
        OutputDir     = "$env:USERPROFILE\Pictures\from iphone\Processed\Videos" # Output folder
        Recurse       = $true        # Process all subfolders
        SkipExisting  = $true        # Skip already processed videos

        # Quality & performance settings
        Crf           = 28           # Constant Rate Factor: 26–29 good balance; lower = higher quality
        Preset        = 'slower'     # slower = smaller output size
        Fps           = 24           # Cap FPS (reduce only)
        AudioBitrateK = 96           # Audio bitrate in Kbps (mono AAC)
        LiveCutoffSec = 4.5          # Skip clips ≤ this length

        # Video filters
        Denoise       = $true        # Apply hqdn3d filter for cleaner visuals
        Decimate      = $true        # Remove near-duplicate frames (mpdecimate)

        # Encoder tool location
        FfmpegPath    = "C:\Scripts\Tools\ffmpeg.exe"

        Verbose       = $true
    }

    # Execute video conversion
    & $vid @vidArgs

    $vidTimer.Stop()
    Write-Success "convertVideos.ps1 completed in $([math]::Round($vidTimer.Elapsed.TotalSeconds,2)) seconds."


} catch {
    # Exception handler for any stage
    Write-Err "$($_.Exception.Message)"

} finally {
    # Final block — always runs, even on error
    $scriptTimer.Stop()
    Write-Success  "`n===== ALL TASKS COMPLETE ====="
    Write-VerboseMsg  "Total elapsed time: $($scriptTimer.Elapsed.ToString())"
}
