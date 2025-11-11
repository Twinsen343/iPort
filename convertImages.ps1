<#
convertImages.ps1  —  iPort Toolkit
(Seconds-Guaranteed Names + Skip-Existing + Pinned Bottom Progress Bar)

WHAT THIS VERSION DOES
- Draws the progress bar on the **last console row** (pinned with RawUI), so no smearing.
- Updates the bar **for every file**, including when it is **skipped** for duplicates.
- Debounced redraws (BarDebounceMs) to keep the UI smooth.

FEATURES
- Converts HEIC/HEIF/JPG → JPG or PNG (downscale-only).
- Stable, second-precise filenames: yyyy-MM-dd_HH-mm-ss[_NN].ext
- Shell + EXIF + FS date fusion with a seconds patch when needed.
- Duplicate-safe: skips outputs that already exist.

USAGE
  .\convertImages.ps1 `
    -InputDir  "$env:USERPROFILE\Pictures\from iphone\Raw\Images" `
    -OutputDir "$env:USERPROFILE\Pictures\from iphone\Processed\Images" `
    -Format jpg -Quality 85 `
    -MaxWidth 1920 -MaxHeight 1080 `
    -SkipExisting -Recurse `
    -ToolPath "C:\Program Files\ImageMagick-7.1.2-Q16-HDRI\magick.exe"
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)] [string]$InputDir,
  [string]$OutputDir,
  [ValidateSet('jpg','png')] [string]$Format = 'jpg',
  [ValidateRange(1,100)] [int]$Quality = 85,
  [int]$MaxWidth = 1920,
  [int]$MaxHeight = 1080,
  [switch]$Recurse,
  [switch]$SkipExisting = $true,
  [switch]$DeleteOriginal,
  [switch]$DryRun,
  [switch]$ShowCandidates,
  [switch]$NoProgress,
  [string]$ToolPath,

  # UI tuning
  [ValidateSet('Black','DarkBlue','DarkGreen','DarkCyan','DarkRed','DarkMagenta','DarkYellow','Gray','DarkGray','Blue','Green','Cyan','Red','Magenta','Yellow','White')]
  [string]$BarColor = 'Yellow',
  [int]$BarDebounceMs = 120
)

$ErrorActionPreference = 'Stop'

# ---------------- Bottom bar helpers ----------------
$script:__barClock   = [System.Diagnostics.Stopwatch]::StartNew()
$script:__barLastMs  = -99999
$script:__barLastPct = -1

function _ConsoleWidth { try { [Math]::Max(40, $Host.UI.RawUI.WindowSize.Width) } catch { 120 } }
function Format-TS([TimeSpan]$ts){ if ($ts.TotalHours -ge 1) { "{0:hh\:mm\:ss}" -f $ts } else { "{0:mm\:ss}" -f $ts } }

function _DrawWideBar {
  param([int]$pct,[string]$note)
  if($NoProgress){return}
  $pct  = [Math]::Max(0,[Math]::Min(100,$pct))
  $w    = _ConsoleWidth
  $note = $note -replace "`r"," " -replace "`n"," "

  # Left-aligned: fixed bar width (40% of console width) for stability
  $barW = [Math]::Max(10,[Math]::Floor($w * 0.4))
  $fill = [Math]::Floor(($pct/100.0)*$barW)
  $bar  = '▏'+('█'*$fill)+('─'*($barW-$fill))+'▕'

  Write-Host ("`r{0} {1}" -f $bar,$note) -NoNewline -ForegroundColor $BarColor
}
function _ClearWideBar {
  if($NoProgress){return}
  $w=_ConsoleWidth
  Write-Host ("`r"+(' '*$w)+"`r") -NoNewline
}
function _UpdateBarCore {
  param([int]$done,[int]$total,[Diagnostics.Stopwatch]$timer,[string]$status)
  if($NoProgress){return}
  $done=[Math]::Max(0,[Math]::Min($total,$done))
  $pct=[int]([Math]::Floor(($done/[double][Math]::Max(1,$total))*100))
  $elapsed=$timer.Elapsed
  $rate  = if($done -gt 0){ $elapsed.TotalSeconds / $done } else { 0 }
  $remain= if($rate -gt 0){ [TimeSpan]::FromSeconds([int]($rate*($total-$done))) } else { [TimeSpan]::Zero }
  $note=("[{0,3}%] elapsed {1} • ETA {2} • {3}" -f $pct,(Format-TS $elapsed),(Format-TS $remain),$status)
  _DrawWideBar -pct $pct -note $note
  $script:__barLastPct = $pct
}
function _UpdateBarSmart {
  param([int]$done,[int]$total,[Diagnostics.Stopwatch]$timer,[string]$status,[switch]$Force)
  if($NoProgress){return}
  $now = $script:__barClock.ElapsedMilliseconds
  $pctNow=[int]([Math]::Floor(([Math]::Min($total,[Math]::Max(0,$done)) / [double][Math]::Max(1,$total))*100))
  if($Force -or ($pctNow -ne $script:__barLastPct) -or ($now -ge ($script:__barLastMs + $BarDebounceMs))){
    _UpdateBarCore -done $done -total $total -timer $timer -status $status
    $script:__barLastMs = $now
  }
}

# --- colored output helpers (clear bar before printing) ---
function Write-Info    { param([string]$m) _ClearWideBar; Write-Host $m -ForegroundColor Cyan }
function Write-Success { param([string]$m) _ClearWideBar; Write-Host $m -ForegroundColor Green }
function Write-Skip    { param([string]$m) _ClearWideBar; Write-Host $m -ForegroundColor DarkYellow }
function Write-Fail    { param([string]$m) _ClearWideBar; Write-Host $m -ForegroundColor Red }

# --- detect seconds in time string ---
function Test-HasSecondsInString {
  param([string]$s)
  if (-not $s) { return $false }
  $s = $s -replace '[\u200E\u200F\u202A-\u202E]', '' ; $s = $s.Trim()
  $m = [regex]::Match($s, '(\d{1,2}):(\d{2})(?::(\d{2}))')
  return ($m.Success -and $m.Groups[3].Success)
}

# --- robust date parsing ---
function Parse-Date {
  param([string]$s)
  if (-not $s) { return $null }
  $s = $s -replace '^(\d{4}):(\d{2}):(\d{2})','${1}-${2}-${3}'
  $s = $s -replace '[\u200E\u200F\u202A-\u202E]', ''
  $s = $s.Trim()
  $formats = @(
    'yyyy-MM-dd HH:mm:ssK','yyyy-MM-dd HH:mm:ss',
    "yyyy-MM-dd'T'HH:mm:ssK","yyyy-MM-dd'T'HH:mm:ss"
  )
  [datetime]$dt = [datetime]::MinValue
  foreach($f in $formats){
    if([datetime]::TryParseExact($s,$f,[Globalization.CultureInfo]::InvariantCulture,
      [Globalization.DateTimeStyles]::AssumeLocal,[ref]$dt)){ return $dt }
  }
  if([datetime]::TryParse($s,[Globalization.CultureInfo]::CurrentCulture,
      [Globalization.DateTimeStyles]::AssumeLocal,[ref]$dt)){ return $dt }
  return $null
}

# --- converter discovery ---
function Resolve-ConverterTool {
  param([string]$ToolPath)
  if ($ToolPath) {
    if (-not (Test-Path $ToolPath)) { throw "ToolPath not found: $ToolPath" }
    $exe = [IO.Path]::GetFileName($ToolPath).ToLowerInvariant()
    if ($exe -eq 'magick.exe' -or $exe -eq 'magick') { return @{Type='magick'; Path=$ToolPath} }
    if ($exe -eq 'heif-convert.exe' -or $exe -eq 'heif-convert') { return @{Type='heif'; Path=$ToolPath} }
    throw "Unsupported ToolPath: $ToolPath"
  }
  $mag = Get-Command magick -ErrorAction SilentlyContinue
  if ($mag) { return @{Type='magick'; Path=$mag.Path} }
  $heif = Get-Command heif-convert -ErrorAction SilentlyContinue
  if ($heif) { return @{Type='heif'; Path=$heif.Path} }
  throw "No converter found. Install ImageMagick (recommended) or supply -ToolPath."
}

function Get-IM-Dates {
  param([Parameter(Mandatory=$true)][string]$ToolPath,
        [Parameter(Mandatory=$true)][string]$FullName)
  $fmt = '%[EXIF:DateTimeOriginal]\n%[EXIF:DateTime]\n%[QuickTime:CreateDate]\n%[QuickTime:CreationDate]\n%[date:create]'
  $psi = New-Object Diagnostics.ProcessStartInfo
  $psi.FileName = $ToolPath
  $psi.Arguments = ('identify -quiet -format "{0}" "{1}"' -f $fmt, $FullName)
  $psi.UseShellExecute = $false
  $psi.RedirectStandardError = $true
  $psi.RedirectStandardOutput = $true
  $p=[Diagnostics.Process]::Start($psi); $p.WaitForExit()
  $stdout = $p.StandardOutput.ReadToEnd()
  if (-not $stdout) { return @() }
  return $stdout -split "`n" | ForEach-Object { $_.Trim() }
}

# --- improved DateTaken ---
function Get-DateTaken {
  param([IO.FileInfo]$File, [hashtable]$ToolOrNull = $null)
  try {
    $shell  = New-Object -ComObject Shell.Application
    $folder = $shell.Namespace($File.DirectoryName)
    $item   = $folder.ParseName($File.Name)
    for ($i=0; $i -lt 200; $i++) {
      $name = $folder.GetDetailsOf($null, $i)
      if ($name -match 'Date\s*taken') {
        $valRaw = $folder.GetDetailsOf($item, $i)
        if ($valRaw) {
          $val = $valRaw -replace '[\u200E\u200F\u202A-\u202E]', '' ; $val = $val.Trim()
          [datetime]$dtShell=[datetime]::MinValue
          if([datetime]::TryParse($val,[Globalization.CultureInfo]::CurrentCulture,
            [Globalization.DateTimeStyles]::AssumeLocal,[ref]$dtShell)){
            if (Test-HasSecondsInString $val) { return $dtShell }

            # try EXIF via ImageMagick
            $toolLocal = $ToolOrNull
            if (-not $toolLocal) { try { $toolLocal = Resolve-ConverterTool -ToolPath $ToolPath } catch { $toolLocal=$null } }
            if ($toolLocal -and $toolLocal.Type -eq 'magick') {
              try {
                $candidates = Get-IM-Dates -ToolPath $toolLocal.Path -FullName $File.FullName
                foreach($raw in $candidates){ $dtExif=Parse-Date $raw; if($dtExif){return $dtExif}}
              } catch {}
            }
            # patch seconds if shell time has no seconds
            $fs = $File.LastWriteTime
            return (Get-Date -Year $dtShell.Year -Month $dtShell.Month -Day $dtShell.Day `
                    -Hour $dtShell.Hour -Minute $dtShell.Minute -Second $fs.Second)
          }
        }
      }
    }
  } catch {}
  $toolLocal2=$ToolOrNull
  if (-not $toolLocal2) { try { $toolLocal2=Resolve-ConverterTool -ToolPath $ToolPath } catch { $toolLocal2=$null } }
  if ($toolLocal2 -and $toolLocal2.Type -eq 'magick') {
    try { foreach($raw in (Get-IM-Dates -ToolPath $toolLocal2.Path -FullName $File.FullName)){ $dt=Parse-Date $raw; if($dt){return $dt}} } catch {}
  }
  return $File.LastWriteTimeUtc
}

# --- detect canonical base from name ---
function Try-GetBaseFromName {
  param([string]$NameNoExt)
  $re='^(?<Y>\d{4})-(?<Mo>\d{2})-(?<D>\d{2})_(?<H>\d{2})-(?<Mi>\d{2})-(?<S>\d{2})(?:_(?<N>\d{2}))?$'
  $m=[regex]::Match($NameNoExt,$re)
  if($m.Success){
    if($m.Groups['N'].Success){
      return '{0}-{1}-{2}_{3}-{4}-{5}_{6}' -f $m.Groups['Y'].Value,$m.Groups['Mo'].Value,$m.Groups['D'].Value,$m.Groups['H'].Value,$m.Groups['Mi'].Value,$m.Groups['S'].Value,$m.Groups['N'].Value
    } else {
      return '{0}-{1}-{2}_{3}-{4}-{5}' -f $m.Groups['Y'].Value,$m.Groups['Mo'].Value,$m.Groups['D'].Value,$m.Groups['H'].Value,$m.Groups['Mi'].Value,$m.Groups['S'].Value
    }
  }
  return $null
}

# --- resolve paths ---
$InputDir  = (Resolve-Path $InputDir).Path
if (-not $OutputDir) { $OutputDir = $InputDir } 
elseif (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

# --- scan inputs ---
$patterns=@('*.heic','*.heif','*.jpg','*.jpeg','*.HEIC','*.HEIF','*.JPG','*.JPEG')
$gci=@{LiteralPath=$InputDir;File=$true;ErrorAction='SilentlyContinue'}
if($Recurse){$gci.Recurse=$true}
$files=Get-ChildItem @gci -Include $patterns | Sort-Object FullName
if(-not $files){ Write-Skip "No HEIC/JPG files in '$InputDir'."; return }

# --- process ---
$sw=[Diagnostics.Stopwatch]::StartNew()
[int]$ok=0;[int]$skip=0;[int]$fail=0
$total=$files.Count
$seen=@{}
$toolEnc=$null

_UpdateBarSmart -done 0 -total $total -timer $sw -status "[0/$total] starting..." -Force

for($i=0;$i -lt $total;$i++){
  $idx=$i+1
  $src=$files[$i]
  $nameNoExt=[IO.Path]::GetFileNameWithoutExtension($src.Name)
  $base=Try-GetBaseFromName -NameNoExt $nameNoExt
  if(-not $base){ $dt=Get-DateTaken -File $src; $base=$dt.ToLocalTime().ToString('yyyy-MM-dd_HH-mm-ss') }

  if(-not $seen.ContainsKey($base)){$seen[$base]=0}else{$seen[$base]++}
  $suffix=if($seen[$base]-gt 0){'_{0:D2}' -f $seen[$base]}else{''}
  $targetName="$base$suffix.$Format"; $outPath=Join-Path $OutputDir $targetName

  _UpdateBarSmart -done ($idx-1) -total $total -timer $sw -status ("[{0}/{1}] {2} → {3}" -f $idx,$total,$src.Name,$targetName)

  if($SkipExisting -and (Test-Path $outPath)){
    $skip++
    Write-Skip ("[{0}/{1}] Skip (exists) {2} → {3}" -f $idx,$total,$src.Name,$outPath)
    _UpdateBarSmart -done $idx -total $total -timer $sw -status ("[{0}/{1}] skipped (exists) {2}" -f $idx,$total,(Split-Path -Leaf $outPath)) -Force
    continue
  }

  try{
    if(-not $toolEnc){$toolEnc=Resolve-ConverterTool -ToolPath $ToolPath}
    if(-not $DryRun){
      if($toolEnc.Type -eq 'magick'){
        $args=@(
          '"{0}"' -f $src.FullName,
          '-auto-orient','-strip','-colorspace','sRGB',
          ('-resize {0}x{1}>' -f $MaxWidth,$MaxHeight)
        )
        if($Format -eq 'jpg'){
          $args+=('-quality {0}' -f $Quality);$args+='-sampling-factor 4:2:0'
        } else {
          $args+='-define png:compression-level=9'
          $args+='-define png:exclude-chunk=all'
        }
        $args+=('"{0}"' -f $outPath)
        $psi=New-Object Diagnostics.ProcessStartInfo
        $psi.FileName=$toolEnc.Path
        $psi.Arguments=($args -join ' ')
        $psi.UseShellExecute=$false
        $psi.RedirectStandardError=$true
        $psi.RedirectStandardOutput=$true
        $p=[Diagnostics.Process]::Start($psi);$p.WaitForExit()
        if($p.ExitCode -ne 0 -or -not (Test-Path $outPath)){throw "magick failed (code $($p.ExitCode))"}
      } elseif($toolEnc.Type -eq 'heif') {
        $psi=New-Object Diagnostics.ProcessStartInfo
        $psi.FileName=$toolEnc.Path
        $psi.Arguments=('"{0}" "{1}"' -f $src.FullName,$outPath)
        $psi.UseShellExecute=$false
        $psi.RedirectStandardError=$true
        $psi.RedirectStandardOutput=$true
        $p=[Diagnostics.Process]::Start($psi);$p.WaitForExit()
        if($p.ExitCode -ne 0 -or -not (Test-Path $outPath)){throw "heif-convert failed (code $($p.ExitCode))"}
      }
    }
    if($DeleteOriginal -and -not $DryRun){
      try { Remove-Item -LiteralPath $src.FullName -Force } catch {}
    }
    $ok++; Write-Success ("[OK] {0}/{1} {2}" -f $idx,$total,$targetName)
  } catch {
    $fail++; Write-Fail ("Failed: {0}" -f $src.FullName)
    Write-Fail ("Reason : {0}" -f $_.Exception.Message)
  }

  _UpdateBarSmart -done $idx -total $total -timer $sw -status ("[{0}/{1}] done {2}" -f $idx,$total,$targetName)
}

_UpdateBarSmart -done $total -total $total -timer $sw -status ("All images processed") -Force
_ClearWideBar

$sw.Stop()
$elapsed=[math]::Round($sw.Elapsed.TotalSeconds,1)
$summary="Done. Converted:{0} Skipped:{1} Failed:{2} Elapsed:{3:N1}s" -f $ok,$skip,$fail,$elapsed
if($fail -gt 0){Write-Fail $summary}elseif($skip -gt 0){Write-Skip $summary}else{Write-Success $summary}
