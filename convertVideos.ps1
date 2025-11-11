<#
convertVideos.ps1  —  iPort Toolkit
(Seconds-Guaranteed Names + Skip-Existing + Pinned Bottom Progress Bar)


#>
<#
convertVideos.ps1
(H.264 — numeric ffprobe gate + bottom [X/Y] progress + ffmpeg machine progress above)

- MOV/MP4/M4V → MP4 (H.264 + AAC), downscale-only ≤1920x1080
- Output names from capture date: yyyy-MM-dd_HH-mm-ss[_NN].mp4
- Strict skip for clips ≤ LiveCutoffSec (numeric ffprobe)
- Uses ffmpeg -progress pipe:1 for accurate within-file ETA
- ffmpeg one-liners print ABOVE; global [X/Y] bar pinned at bottom (Yellow)
- NO ghost bars; minimal flashing via smart redraw + batched skip prints
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)] [string]$InputDir,
  [Parameter(Mandatory=$true)] [string]$OutputDir,
  [switch]$Recurse,
  [switch]$SkipExisting = $true,

  [int]$MaxWidth  = 1920,
  [int]$MaxHeight = 1080,

  [ValidateRange(0,51)] [int]$Crf = 26,
  [ValidateSet('ultrafast','superfast','veryfast','faster','fast','medium','slow','slower','veryslow')]
  [string]$Preset = 'slower',

  [double]$Fps = 24,
  [int]$AudioBitrateK = 64,
  [string]$FfmpegPath,

  [switch]$Denoise,
  [switch]$Decimate,
  [switch]$VerboseOutput,
  [switch]$NoProgress,

  [double]$LiveCutoffSec = 3.0,

  # Bar color (default Yellow)
  [ValidateSet('Black','DarkBlue','DarkGreen','DarkCyan','DarkRed','DarkMagenta','DarkYellow','Gray','DarkGray','Blue','Green','Cyan','Red','Magenta','Yellow','White')]
  [string]$BarColor = 'Yellow',

  # UI tuning
  [int]$BarDebounceMs = 120,
  [int]$SkipPrintIntervalMs = 600
)

$ErrorActionPreference = 'Stop'

# ---------- Bottom bar helpers ----------
$script:__barClock = [System.Diagnostics.Stopwatch]::StartNew()
$script:__barLastMs = -99999
$script:__barLastPct = -1

function _ConsoleWidth { try { [Math]::Max(40, $Host.UI.RawUI.WindowSize.Width) } catch { 120 } }
function _FmtTimeSpan([TimeSpan]$ts){ if($ts.TotalHours -ge 1){"{0:hh\:mm\:ss}" -f $ts}else{"{0:mm\:ss}" -f $ts} }
function _DrawWideBar {
  param([int]$pct,[string]$note)
  if($NoProgress){return}
  $pct=[Math]::Max(0,[Math]::Min(100,$pct))
  $w=_ConsoleWidth
  $note=$note -replace "`r"," " -replace "`n"," "
  $barW=[Math]::Max(10,$w-($note.Length+3))
  if($barW -gt ($w-10)){$barW=$w-10}
  $fill=[Math]::Floor(($pct/100.0)*$barW)
  $bar='▏'+('█'*$fill)+('─'*($barW-$fill))+'▕'
  Write-Host ("`r{0} {1}" -f $bar,$note).PadRight($w) -NoNewline -ForegroundColor $BarColor
}
function _ClearWideBar {
  if($NoProgress){return}
  $w=_ConsoleWidth
  Write-Host ("`r"+(' '*$w)+"`r") -NoNewline
}
function _UpdateBarCore {
  param([double]$done,[int]$total,[Diagnostics.Stopwatch]$timer,[string]$fileNote)
  if($NoProgress){return}
  $done  = [Math]::Max(0.0,[Math]::Min([double]$total,[double]$done))
  $pctD  = ($done/[double][Math]::Max(1,$total))*100.0
  $pct   = [int]([Math]::Floor($pctD))   # floor for stable display
  $elapsed=$timer.Elapsed
  $rateSec= if($done -gt 0){ $elapsed.TotalSeconds / $done } else { 0 }
  $remain= if($rateSec -gt 0){ [TimeSpan]::FromSeconds([int]($rateSec*($total-$done))) } else { [TimeSpan]::Zero }
  $note=("[{0,3}%] elapsed {1} • ETA {2} • {3}" -f $pct,(_FmtTimeSpan $elapsed),(_FmtTimeSpan $remain),$fileNote)
  _DrawWideBar -pct $pct -note $note
  $script:__barLastPct = $pct
}
function _UpdateBarSmart {
  param([double]$done,[int]$total,[Diagnostics.Stopwatch]$timer,[string]$fileNote,[int]$MinMs,[switch]$Force)
  if($NoProgress){return}
  $now = $script:__barClock.ElapsedMilliseconds
  # Only redraw if:
  #  - forced, OR
  #  - debounce window elapsed, OR
  #  - integer pct changed
  $pctNow = [int]([Math]::Floor( [Math]::Min(100.0,[Math]::Max(0.0, ($done/[Math]::Max(1,$total))*100.0 )) ))
  if($Force -or ($pctNow -ne $script:__barLastPct) -or ($now -ge ($script:__barLastMs + $MinMs))){
    _UpdateBarCore -done $done -total $total -timer $timer -fileNote $fileNote
    $script:__barLastMs = $now
  }
}

# ---------- Colour helpers (bar-safe; clear before printing) ----------
function Write-Info    { param([string]$m) _ClearWideBar; Write-Host $m -ForegroundColor Cyan }
function Write-Success { param([string]$m) _ClearWideBar; Write-Host $m -ForegroundColor Green }
function Write-SkipMsg { param([string]$m) _ClearWideBar; Write-Host $m -ForegroundColor DarkYellow }
function Write-Fail    { param([string]$m) _ClearWideBar; Write-Host $m -ForegroundColor Red }

# ---------- Tools ----------
function Get-FFmpeg([string]$Path){
  if($Path){ if(-not(Test-Path $Path)){throw "ffmpeg not found: $Path"}; return $Path }
  $c=Get-Command ffmpeg -ErrorAction SilentlyContinue
  if($c){return $c.Path}; throw "ffmpeg.exe not found — install it or pass -FfmpegPath"
}
function Get-FFprobe([string]$ffmpegPath){
  if($ffmpegPath){
    $dir=Split-Path -Parent $ffmpegPath
    $probe=Join-Path $dir 'ffprobe.exe'
    if(Test-Path $probe){return $probe}
  }
  $c=Get-Command ffprobe -ErrorAction SilentlyContinue
  if($c){return $c.Path}; throw "ffprobe.exe not found — install it or keep it next to ffmpeg"
}

# ---------- Metadata helpers ----------
function Remove-InvisibleMarks { param([string]$s) if(-not $s){return $s}; ($s -replace "[\u200E\u200F\u202A-\u202E]",'').Trim() }
function Parse-Date { param([string]$s)
  if(-not $s){return $null}
  $s=$s -replace '^(\d{4}):(\d{2}):(\d{2})','${1}-${2}-${3}'
  $s=Remove-InvisibleMarks $s
  $strict=@('yyyy-MM-dd HH:mm:ssK','yyyy-MM-dd HH:mm:ss',"yyyy-MM-dd'T'HH:mm:ssK","yyyy-MM-dd'T'HH:mm:ss",'yyyy-MM-ddTHH:mm:ssK','yyyy-MM-ddTHH:mm:ss')
  foreach($f in $strict){ try { return [datetime]::ParseExact($s,$f,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeLocal) } catch {} }
  try { return [datetime]::Parse($s,[Globalization.CultureInfo]::CurrentCulture) } catch {}
  return $null
}

# ---------- ffprobe (numeric) ----------
function ffprobe-NumLine {
  param([string]$ArgsJoined)
  $psi=New-Object Diagnostics.ProcessStartInfo
  $psi.FileName=$script:ffprobe
  $psi.Arguments=$ArgsJoined
  $psi.UseShellExecute=$false
  $psi.RedirectStandardError=$true
  $psi.RedirectStandardOutput=$true
  $p=[Diagnostics.Process]::Start($psi); $p.WaitForExit()
  $out=$p.StandardOutput.ReadToEnd().Trim()
  return $out
}
function TryParse-DoubleInvariant { param([string]$s)
  $d=[double]::NaN
  if([double]::TryParse($s,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$d)){return $d}
  return $null
}
function Probe-StrictNumeric {
  param([string]$Path)
  $d1=ffprobe-NumLine ('-v error -select_streams v:0 -show_entries stream=duration -of default=nk=1:nw=1 -i "{0}"' -f $Path)
  $dur=TryParse-DoubleInvariant $d1
  if(-not $dur){
    $d2=ffprobe-NumLine ('-v error -show_entries format=duration -of default=nk=1:nw=1 -i "{0}"' -f $Path)
    $dur=TryParse-DoubleInvariant $d2
  }
  if(-not $dur){
    $lines=ffprobe-NumLine ('-v error -select_streams v:0 -count_frames 1 -show_entries stream=nb_read_frames,avg_frame_rate -of default=nk=1:nw=1 -i "{0}"' -f $Path)
    $parts=($lines -split "`r?`n") | ? {$_}
    if($parts.Count -ge 2){
      $frames=TryParse-DoubleInvariant $parts[0]
      $afr=$parts[1]; if($afr -match '^\d+/\d+$'){
        $a=$afr -split '/'
        $num=TryParse-DoubleInvariant $a[0]; $den=TryParse-DoubleInvariant $a[1]
        if($num -and $den -and $den -ne 0){$fps=$num/$den}
      }
      if($frames -and $fps -gt 0){$dur=$frames/$fps}
    }
  }
  $rawS=ffprobe-NumLine ('-v error -select_streams v:0 -show_entries stream_tags=creation_time -of default=nk=1:nw=1 -i "{0}"' -f $Path)
  $rawF=$null; if(-not $rawS){$rawF=ffprobe-NumLine ('-v error -show_entries format_tags=creation_time -of default=nk=1:nw=1 -i "{0}"' -f $Path)}
  $dt=$null; if($rawS){$dt=Parse-Date $rawS}elseif($rawF){$dt=Parse-Date $rawF}
  if(-not $dt){try{$dt=(Get-Item -LiteralPath $Path).LastWriteTimeUtc}catch{$dt=Get-Date}}
  @{Duration=$dur;CreateDate=$dt}
}

# ---------- Scaling chain ----------
function Get-ScaleLandscape([int]$MaxW){ "scale='min(iw,${MaxW})':-2:flags=lanczos,setsar=1" }
function Get-ScalePortrait ([int]$MaxH){ "scale=-2:'min(ih,${MaxH})':flags=lanczos,setsar=1" }
function Compose-Vf([string]$base,[double]$fps,[bool]$useDecimate,[bool]$useDenoise){
  $p=@(); if($base){$p+=$base}; if($useDecimate){$p+='mpdecimate'}; if($useDenoise){$p+='hqdn3d=1.5:1.5:3:3'}
  if($fps -gt 0){$p+=("fps={0}" -f $fps)}; ($p -join ',')
}

# ---------- Paths / tools ----------
$InputDir=(Resolve-Path $InputDir).Path
if(-not(Test-Path $OutputDir)){New-Item -ItemType Directory -Force -Path $OutputDir|Out-Null}
$OutputDir=(Resolve-Path $OutputDir).Path
$ffmpeg=Get-FFmpeg $FfmpegPath
$ffprobe=Get-FFprobe $FfmpegPath

# ---------- Scan inputs ----------
$patterns=@('*.mov','*.mp4','*.m4v','*.MOV','*.MP4','*.M4V')
$gci=@{LiteralPath=$InputDir;File=$true;ErrorAction='SilentlyContinue'}
if($Recurse){$gci.Recurse=$true}
$files=Get-ChildItem @gci -Include $patterns | Sort-Object FullName
if(-not $files){ Write-SkipMsg "No video files found in '$InputDir'."; return }

# ---------- Skip batching state ----------
$script:__skipFlushClock = [System.Diagnostics.Stopwatch]::StartNew()
$script:__skipLastFlushMs = -99999
$script:__skipExist = 0
$script:__skipShort = 0

function _FlushSkipBatch([switch]$Force){
  $now = $script:__skipFlushClock.ElapsedMilliseconds
  if(-not $Force -and ($now -lt ($script:__skipLastFlushMs + $SkipPrintIntervalMs))){ return }
  if(($script:__skipExist + $script:__skipShort) -gt 0){
    _ClearWideBar
    $msg = "[Skips] exists={0} short={1}" -f $script:__skipExist, $script:__skipShort
    Write-SkipMsg $msg
    $script:__skipExist = 0
    $script:__skipShort = 0
    $script:__skipLastFlushMs = $now
  }
}

# ---------- Main ----------
$sw=[Diagnostics.Stopwatch]::StartNew()
[int]$ok=0;[int]$skip=0;[int]$fail=0
$total=$files.Count
$seen=@{}
$gop=[int]([Math]::Max(1,[Math]::Round([double]$Fps*5)))

for($i=0;$i -lt $total;$i++){
  $idx=$i+1; $src=$files[$i]
  $p=Probe-StrictNumeric -Path $src.FullName
  $dur=$p.Duration

  # ---- Skips ----
  if($dur -and $dur -le $LiveCutoffSec){
    $skip++; $script:__skipShort++
    _FlushSkipBatch
    _UpdateBarSmart -done $idx -total $total -timer $sw -fileNote ("[{0}/{1}] skipped (short) {2}" -f $idx,$total,$src.Name) -MinMs $BarDebounceMs
    continue
  }

  # ---- Name preflight (+suffix) ----
  $localDt=$p.CreateDate.ToLocalTime()
  $stamp=$localDt.ToString('yyyy-MM-dd_HH-mm-ss')
  if(-not $seen.ContainsKey($stamp)){$seen[$stamp]=0}else{$seen[$stamp]++}
  $suffix=if($seen[$stamp]-gt 0){'_{0:D2}' -f $seen[$stamp]}else{''}
  $finalOut=Join-Path $OutputDir ($stamp+$suffix+'.mp4')

  if($SkipExisting -and (Test-Path $finalOut)){
    $skip++; $script:__skipExist++
    _FlushSkipBatch
    _UpdateBarSmart -done $idx -total $total -timer $sw -fileNote ("[{0}/{1}] skipped (exists) {2}" -f $idx,$total,(Split-Path -Leaf $finalOut)) -MinMs $BarDebounceMs
    continue
  }
  $tempOut=$finalOut+'.tmp'

  # ---- Show overall bar BEFORE encode (only completed counted) ----
  _UpdateBarSmart -done ($idx-1) -total $total -timer $sw -fileNote ("[{0}/{1}] {2} → {3}" -f $idx,$total,$src.Name,(Split-Path -Leaf $finalOut)) -MinMs $BarDebounceMs -Force

  # filters: landscape → portrait → noscale
  $tryVfs=@(
    (Compose-Vf (Get-ScaleLandscape $MaxWidth) $Fps $Decimate $Denoise),
    (Compose-Vf (Get-ScalePortrait  $MaxHeight) $Fps $Decimate $Denoise),
    (Compose-Vf $null               $Fps $Decimate $Denoise)
  )

  $exit=1
  foreach($vf in $tryVfs){
    $args=@(
      '-hide_banner','-nostdin','-y',
      '-loglevel','error','-nostats','-progress','pipe:1',
      '-i', $src.FullName,
      '-c:v','libx264','-preset',$Preset,'-crf',"$Crf",
      '-pix_fmt','yuv420p','-profile:v','high','-level:v','4.1',
      '-g',"$gop",'-keyint_min',"$gop",'-sc_threshold','40'
    )
    if($vf){ $args += @('-vf',$vf) }
    $args += @('-c:a','aac','-b:a',("$AudioBitrateK"+"k"),'-ac','1',
               '-movflags','+faststart','-f','mp4', $tempOut)

    # PS5.1-safe arg string
    $argString = ($args | ForEach-Object {
      if ($_ -match '[\s"]') { '"' + ($_ -replace '"','\"') + '"' } else { $_ }
    }) -join ' '

    # ---- ffmpeg run; parse its progress ABOVE; bar smart-redraw ----
    $psi=New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName=$ffmpeg
    $psi.Arguments=$argString
    $psi.UseShellExecute=$false
    $psi.RedirectStandardOutput=$true   # progress stream
    $psi.RedirectStandardError=$false
    $psi.CreateNoWindow=$true

    $p2=[System.Diagnostics.Process]::Start($psi)
    $out=$p2.StandardOutput

    $fileDurSec = if($dur){ [double]$dur } else { $null }
    $curTimeSec = 0.0
    $lastSpeed=$null; $lastBitrateK=$null

    while(-not $p2.HasExited){
      while(-not $out.EndOfStream){
        $line=$out.ReadLine()
        if(-not $line){ continue }
        $kv = $line -split '=', 2
        if($kv.Count -ne 2){ if($VerboseOutput){Write-Host $line}; continue }
        $key=$kv[0]; $val=$kv[1]

        switch ($key) {
          'out_time_ms' {
            $us=0L; [void][long]::TryParse($val,[ref]$us)
            if($us -gt 0){ $curTimeSec = [double]$us / 1000000.0 }
            $within = 0.0
            if($fileDurSec -and $fileDurSec -gt 0){ $within = [Math]::Min(1.0,[Math]::Max(0.0,$curTimeSec/$fileDurSec)) }
            $overallDone = ($idx-1) + $within

            _ClearWideBar
            Write-Host ("time={0}  speed={1}x  bitrate={2}k" -f `
              ([TimeSpan]::FromSeconds([int]$curTimeSec).ToString()), `
              ($lastSpeed ?? ''), ($lastBitrateK ?? ''))

            _UpdateBarSmart -done $overallDone -total $total -timer $sw `
              -fileNote ("[{0}/{1}] {2} → {3}" -f $idx,$total,$src.Name,(Split-Path -Leaf $finalOut)) `
              -MinMs $BarDebounceMs
          }
          'speed'   { $lastSpeed = ($val -replace 'x$','') }
          'bitrate' {
            $m=[regex]::Match($val,'([0-9.]+)kbits/s'); if($m.Success){ $lastBitrateK = $m.Groups[1].Value }
          }
          default { if($VerboseOutput){ Write-Host "$key=$val" } }
        }
      }
      Start-Sleep -Milliseconds 60
    }
    $p2.WaitForExit()
    $exit=$p2.ExitCode

    if(($exit -eq 0) -and (Test-Path $tempOut) -and ((Get-Item $tempOut).Length -gt 0)){ break }
    Write-Info "[Retry] Trying next VF chain..."
  }

  if(($exit -ne 0) -or -not (Test-Path $tempOut) -or ((Get-Item $tempOut).Length -eq 0)){
    Write-Fail ("Failed: {0}" -f $src.FullName)
    Write-Fail ("Reason: ffmpeg failed or empty file (exit {0})" -f $exit)
    if(Test-Path $tempOut){ Remove-Item $tempOut -Force -ErrorAction SilentlyContinue }
    $fail++
  } else {
    if(Test-Path $finalOut){ Remove-Item $finalOut -Force -ErrorAction SilentlyContinue }
    Move-Item $tempOut $finalOut -Force
    Write-Success ("[OK] {0}" -f (Split-Path -Leaf $finalOut))
    $ok++
  }

  # Progress update as we move to next file
  $progressDone = if($idx -lt $total){ [double]$idx } else { [double]$total - 0.01 }
  _UpdateBarSmart -done $progressDone -total $total -timer $sw `
    -fileNote ("[{0}/{1}] done {2}" -f $idx,$total,(Split-Path -Leaf $finalOut)) -MinMs $BarDebounceMs

  # Ensure any skips buffered are printed occasionally
  _FlushSkipBatch
}

# Final 100% redraw + flush any remaining skip summary
_UpdateBarSmart -done $total -total $total -timer $sw -fileNote ("All conversions complete") -MinMs 0 -Force
_FlushSkipBatch -Force

# Final clear + summary
_ClearWideBar
$sw.Stop()
$summary="Done. Converted: {0}, Skipped: {1}, Failed: {2}. Elapsed: {3:N1}s" -f $ok,$skip,$fail,$sw.Elapsed.TotalSeconds
if($fail -gt 0){Write-Fail $summary}elseif($skip -gt 0){Write-SkipMsg $summary}else{Write-Success $summary}
