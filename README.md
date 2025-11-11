# iPort — iPhone & MTP Windows Media Automation Toolkit

iPort is a modular PowerShell toolkit that automates importing, organizing, and converting iPhone or MTP photos and videos on Windows — without iTunes, iCloud, or manual drag-and-drop. It connects directly to MTP (Media Transfer Protocol) devices and uses FFmpeg and ImageMagick to convert and optimize media for storage or business use.

## Why iPort Exists
MTP is notoriously unreliable on Windows. It’s not a true filesystem — it’s a COM-based object interface that constantly drops connections, hides extensions, and times out during long copies. iPort was built to fix that.
- Reliable transfers – Opens a fresh stream for every file, retries on timeouts, and rebinds when the device disconnects.
- Accurate naming & metadata – Derives filenames and timestamps from Shell, EXIF, and file info automatically.
- Duplicate protection – Skips or verifies existing files using hashes and safe temp-file staging.
- Clear progress – Clean, non-ghosting progress bars with explicit “skipped / failed / copied” states.

### Key Features
- Safe, resumable import from iPhone or MTP (Internal Storage/DCIM/...)
- Organized output (Images\YYYY-MM and Videos\YYYY-MM)
- Conversion via FFmpeg & ImageMagick (H.264/H.265, HEIF to JPG/PNG)
- Read-only by default (optional -DeleteAfterCopy)
- Rich on-screen logs; no external dependencies beyond PowerShell + tools


---

##  Core Components

| Script | Purpose | Key Features |
|--------|----------|---------------|
| **CopyFromIPhone.ps1** | Robust MTP importer | Transfers from iPhone, skips duplicates, supports resume/checkpoint, handles timeouts |
| **convertImages.ps1** | HEIC / HEIF → JPG or PNG converter | EXIF + Shell + FS date detection, second-precise filenames, duplicate skip, left-aligned progress bar |
| **convertVideos.ps1** | MOV / MP4 → H.264 MP4 transcoder | Smart duration detection, live-photo skip, downscale-only conversion, accurate ETA |
| **Run-Pipeline.ps1** | Unified controller (runs all above) | Centralized config, enhanced console output, per-stage timers, unified success/error handling |




##  Run-Pipeline.ps1

Run-Pipeline.ps1 is the main entry point that executes all three stages sequentially:

1 Copy media from iPhone
2 Convert images
3 Convert videos

Example
.\Run-Pipeline.ps1

##  Configuration Overview
All configurable parameters are centralized in Run-Pipeline.ps1.

Step 1 — Import from iPhone
```
$params = @{
  Destination     = "$env:USERPROFILE\Pictures\from iphone\Raw"
  DeviceName      = 'iPhone' # searches wildcard
  SkipByNameOnly  = $true
  NoDateFolders   = $true
  Progress        = $true
  TimeoutSeconds  = 20
}
```
| Parameter | Description |
|------------|-------------|
| **Destination** | The local folder where all photos and videos copied from the iPhone will be stored. Typically defaults to `"$env:USERPROFILE\Pictures\from iphone\Raw"`. The script will automatically create this folder if it doesn’t exist. |
| **DeviceName** | The name of the MTP device as it appears in Windows Explorer. For most users this will be `'iPhone'`, but it can be changed if the connected device reports a different name. |
| **SkipByNameOnly** | When enabled (`$true`), the script compares only filenames (ignoring timestamps and metadata) to detect duplicates. Files with the same name will be skipped even if their metadata differs. |
| **NoDateFolders** | Prevents the script from creating year/month subfolders inside the destination. If `$true`, all files go directly into the destination folder instead of being grouped by date. |
| **Progress** | Displays a live progress bar in the PowerShell console showing copy progress. Set to `$false` to disable progress output for silent or automated runs. |
| **TimeoutSeconds** | Defines how long (in seconds) the script will wait when accessing the device or reading a file before skipping it and moving on. Prevents hangs if the phone connection stalls or a file is inaccessible. |

---

## Step 2 — Convert Images
```
& $conv `
  -InputDir  "$env:USERPROFILE\Pictures\from iphone\Raw\Images" `
  -OutputDir "$env:USERPROFILE\Pictures\from iphone\Processed\Images" `
  -Format jpg -Quality 85 `
  -MaxWidth 1920 -MaxHeight 1080 `
  -SkipExisting -Recurse `
  -ToolPath "C:\Program Files\ImageMagick-7.1.2-Q16-HDRI\magick.exe"
```
| Parameter | Description |
|------------|-------------|
| **-InputDir** | The source folder containing the original images to be processed. Typically this is the path where images were copied from the iPhone (e.g. `"$env:USERPROFILE\Pictures\from iphone\Raw\Images"`). |
| **-OutputDir** | The target folder where converted images will be saved. It’s usually a structured location such as `"$env:USERPROFILE\Pictures\from iphone\Processed\Images"`. The script will create the folder if it doesn’t exist. |
| **-Format** | Specifies the desired output image format, such as `jpg` or `png`. JPG is typically used for smaller, compressed files. |
| **-Quality** | Controls image compression quality (1–100). A value around 85 gives a good balance between file size and visual quality. |
| **-MaxWidth** | The maximum allowed width of the output image in pixels. Larger images are automatically downscaled to this width if needed. |
| **-MaxHeight** | The maximum allowed height of the output image in pixels. Used together with `-MaxWidth` to maintain aspect ratio without upscaling. |
| **-SkipExisting** | If enabled, the script checks whether a converted file already exists in the output directory and skips it to save processing time. |
| **-Recurse** | Enables recursive folder scanning — includes subfolders under the input directory. Useful if the iPhone import process includes multiple nested folders. |
| **-ToolPath** | The full path to the image conversion tool executable. In this case, it’s pointing to **ImageMagick** (e.g., `"C:\Program Files\ImageMagick-7.1.2-Q16-HDRI\magick.exe"`), which performs the format conversion and resizing. |
---
## Step 3 — Convert Videos
```
$vidArgs = @{
  InputDir     = "$env:USERPROFILE\Pictures\from iphone\Raw\Videos"
  OutputDir    = "$env:USERPROFILE\Pictures\from iphone\Processed\Videos"
  Recurse      = $true
  SkipExisting = $true
  Crf          = 28
  Preset       = 'slower'
  Fps          = 24
  AudioBitrateK= 96
  LiveCutoffSec= 4.5
  Denoise      = $true
  Decimate     = $true
  FfmpegPath   = "C:\Scripts\Tools\ffmpeg.exe"
}
```


| Parameter | Description |
|------------|-------------|
| **InputDir** | The source directory containing the original video files to process. Usually points to the folder where iPhone videos are imported, such as `"$env:USERPROFILE\Pictures\from iphone\Raw\Videos"`. |
| **OutputDir** | The destination directory for converted videos. The script saves all processed `.mp4` files here (e.g., `"$env:USERPROFILE\Pictures\from iphone\Processed\Videos"`). If the directory doesn’t exist, it’s created automatically. |
| **Recurse** | Enables recursive scanning of all subfolders under the input directory, ensuring every video file within nested directories is processed. |
| **SkipExisting** | Prevents reprocessing of files that already exist in the output directory. Useful for incremental runs or large libraries — saves time by skipping duplicates. |
| **Crf** | The **Constant Rate Factor** for video quality. Lower numbers increase quality (and file size), higher numbers reduce size (and quality). A value around `28` gives good quality while keeping file sizes small. *(Typical range: 18–30)* |
| **Preset** | Controls the encoding speed/efficiency trade-off. Slower presets use more CPU but produce smaller files. Common options: `ultrafast`, `fast`, `medium`, `slow`, `slower`. |
| **Fps** | Sets the target frames per second. The script will not increase FPS — it only reduces frame rates if the source exceeds this value, helping keep file sizes down. |
| **AudioBitrateK** | Sets the audio bitrate (in kilobits per second). `96` kbps mono AAC provides clear voice and ambient sound at compact sizes. |
| **LiveCutoffSec** | Automatically skips short video clips (e.g., Live Photos or accidental captures) shorter than this duration in seconds. Default `4.5` means anything ≤4.5 seconds is ignored. |
| **Denoise** | Enables video denoising using the `hqdn3d` filter for smoother visuals in low-light or noisy footage. |
| **Decimate** | Removes near-duplicate frames with the `mpdecimate` filter to reduce file size (especially useful for static scenes or time-lapse footage). |
| **FfmpegPath** | The full path to the **FFmpeg** executable. Used for all encoding, filtering, and media conversion (e.g., `"C:\Scripts\Tools\ffmpeg.exe"`). |

---

## Design Philosophy

- No iCloud, no Photos app. Direct filesystem control.
- Predictable filenames. Every file includes seconds + unique suffixes.
- Transparent. Every step visible and modifiable.
- Efficient. No re-encode if duplicate or already converted.
- Safe. No overwriting or upscaling.

## QuickStart
```
git clone https://github.com/Twinsen343/iPort.git
cd iPort
.\Run-Pipeline.ps1
```


##  Requirements
- Windows 10 / 11
- PowerShell 5.1+ (or PowerShell 7+)
- ffmpeg + ffprobe in PATH
- ImageMagick (or heif-convert) installed
- iPhone connected via USB (MTP mode)
