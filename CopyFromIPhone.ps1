# ====================================================================================================
# iPort: iPhone / MTP Media Automation Toolkit
# ----------------------------------------------------------------------------------------------------
# Component: CopyFromIPhone.ps1 — v3.6.9-MAX+ROBUST+STREAM (PS5.1-safe)
#
# Purpose
#   High-reliability copier for iPhone (MTP) media. Streams folder-by-folder, guards against
#   duplicate prompts via a LOCAL STAGING directory, resumes via checkpoints, and keeps UI alive.
#
# Highlights
#   - Duplicate-safe staging (precheck stage path, purge stale partials, no "Copy and replace?" UI)
#   - Newest-first folder ordering; per-folder warmups to stabilize MTP views
#   - Strict/By-Name duplicate detection; optional no date folder layout
#   - Checkpoint resume per folder; free-space guard; configurable yields
#   - PS5.1-safe; auto-relaunch in STA when required on PowerShell 7+
#
# Requirements
#   - Windows with PowerShell 5.1+ (or 7+), Shell.Application COM
#   - iPhone connected via USB (unlocked; “Trust this computer” accepted)
#   - Run from a non-elevated console for best MTP behaviour
#
# Usage (typical)
#   .\CopyFromIPhone.ps1 -Destination "$env:USERPROFILE\Pictures\from iphone\Raw" -Progress -Checkpoint
#
# Repository
#   https://github.com/yourname/iPort
# ====================================================================================================

<# 
CopyFromIPhone.ps1 — v3.6.9-MAX+ROBUST+STREAM (PS5.1-safe)
- Duplicate-safe staging: pre-check stage path, skip on match, purge stale partials
- Prevents "Copy and replace?" prompts originating from Stage collisions
- Keeps v3.6.8 streaming, warmups, checkpoint, newest-first sorting, free space guard
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)]
  [string]$Destination,

  [string]$DeviceName = 'iPhone',
  [string]$StorageFolderName = 'Internal Storage',

  [switch]$SkipByNameOnly,
  [switch]$NoDateFolders,
  [switch]$Progress,
  [switch]$DeleteAfterCopy,
  [int]$TimeoutSeconds = 20,

  # Structure / diagnostics
  [switch]$ListStructure,
  [switch]$ListOnly,

  # UI keep-alive
  [switch]$Heartbeat,
  [switch]$TextPulse,

  # Local staging to avoid OneDrive/KFM quirks
  [string]$StageDir = (Join-Path $env:TEMP 'CopyFromIPhone_Stage'),

  # Confirm/poll tuning
  [int]$ConfirmMaxWaitSec = 30,
  [int]$ConfirmPollMs     = 250,

  # --- New streaming / scale knobs ---
  [string]$OnlyFolderName,
  [int]$MaxFolders = 0,

  [switch]$Checkpoint,
  [string]$CheckpointPath = (Join-Path $env:TEMP 'CopyFromIPhone.checkpoint.json'),

  [int]$StageFlushEvery = 50,
  [int]$YieldMs = 50,

  [switch]$StrictExtensions,
  [int]$MinFreeGB = 2
)

# --------------------------- MAX-VERBOSE UX ---------------------------
$ErrorActionPreference = 'Stop'
$VerbosePreference     = 'Continue'
$global:ProgressPreference = 'Continue'
$Progress  = $true
$Heartbeat = $true
$TextPulse = $true

function Write-Info    ($m){ Write-Host $m -ForegroundColor Cyan }
function Write-Success ($m){ Write-Host $m -ForegroundColor Green }
function Write-Warn    ($m){ Write-Host $m -ForegroundColor Yellow }
function Write-Err     ($m){ Write-Host $m -ForegroundColor Red }

$__savedProgressPref = $global:ProgressPreference
function Restore-ProgressPreference { try { $global:ProgressPreference = $__savedProgressPref } catch {} }

# Spinner
$script:__spinIdx = 0
$script:__spinSeq = @('|','/','-','\')
function Spin-Start([string]$label){ if($TextPulse){ $script:__spinIdx=0; Write-Host ("{0} {1}" -f $script:__spinSeq[$script:__spinIdx],$label) -NoNewline } }
function Spin-Tick ([string]$label){ if($TextPulse){ $script:__spinIdx=($script:__spinIdx+1)%$script:__spinSeq.Count; $c=$script:__spinSeq[$script:__spinIdx]; Write-Host ("`r{0} {1}" -f $c,$label) -NoNewline } }
function Spin-Stop(){ if($TextPulse){ Write-Host "`r   " -NoNewline; Write-Host "" } }

# Heartbeat
$HB_ID = 777
function HB-Start([string]$Activity,[string]$Status){ if($Heartbeat){ Write-Progress -Id $HB_ID -Activity $Activity -Status $Status -PercentComplete 0 } }
function HB-Tick([int]$Step,[int]$Total,[string]$Status){
  if($Heartbeat){
    $pct = if($Total -gt 0){ [int]([Math]::Min(99,[Math]::Max(0, ($Step/[double]$Total)*100))) } else { ($Step % 100) }
    Write-Progress -Id $HB_ID -Activity "Enumerating device (MTP)" -Status $Status -PercentComplete $pct
  }
}
function HB-Stop(){ if($Heartbeat){ Write-Progress -Id $HB_ID -Activity "Enumerating device (MTP)" -Status "Ready" -Completed } }

# --- Ensure STA for Shell.Application on PowerShell 7+ ---
try { $isCore = ($PSVersionTable.PSEdition -eq 'Core') } catch { $isCore = $false }
try { $apt   = [System.Threading.Thread]::CurrentThread.ApartmentState } catch { $apt = 'Unknown' }

if ($isCore -and $apt -ne [System.Threading.ApartmentState]::STA -and $PSCommandPath) {
  Write-Host "[Info] Relaunching in STA for Shell COM..." -ForegroundColor Yellow

  $exeCmd = $null
  try { $exeCmd = Get-Command pwsh -ErrorAction SilentlyContinue } catch {}
  $exe = $null
  if ($exeCmd) { $exe = $exeCmd.Source }
  else {
    $exe = "$env:ProgramFiles\PowerShell\7\pwsh.exe"
    if (-not (Test-Path $exe)) { $exe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" }
  }

  $args = @('-STA','-NoLogo','-NoProfile','-File',"`"$PSCommandPath`"")
  $skipCommons = @('Verbose','Debug','WarningAction','ErrorAction','InformationAction','ProgressAction','VerbosePreference','ErrorVariable','WarningVariable','InformationVariable','OutVariable','OutBuffer','PipelineVariable')
  foreach($kv in $PSBoundParameters.GetEnumerator()){
    $k = [string]$kv.Key
    if ($skipCommons -contains $k) { continue }
    if ($kv.Value -is [switch]) { if ($kv.Value.IsPresent) { $args += "-$k" }; continue }
    if ($kv.Value -is [System.Collections.IEnumerable] -and -not ($kv.Value -is [string])) {
      foreach($v in $kv.Value){ $args += "-$k"; $args += "`"$v`"" }
    } else {
      $args += "-$k"; $args += "`"$($kv.Value)`""
    }
  }
  Start-Process -FilePath $exe -ArgumentList $args
  exit
}

# --------------------------- Shell / MTP utilities ---------------------------
$IMAGE_EXT = @('jpg','jpeg','heic','png','gif','bmp','tif','tiff','webp','aae','dng')
$VIDEO_EXT = @('mp4','mov','m4v','avi','mkv','hevc','3gp','caf')
$KNOWN_EXT_ORDER = @(
  'jpg','jpeg','heic','png',
  'mp4','mov','m4v','hevc','3gp',
  'gif','webp','bmp','tif','tiff','avi','mkv','aae','dng','caf'
)

function Test-IsFolder { param($Item) try { return [bool]$Item.IsFolder } catch { return $false } }
function Test-Name     { param($Item) try { return [string]$Item.Name } catch { return '(unknown)' } }

function Normalize-Folder {
  param([Parameter(Mandatory)][object]$Folder)
  if ($Folder -is [__ComObject]) { return $Folder }
  if ($Folder -is [object[]]) {
    foreach($e in $Folder){ if($e -is [__ComObject]){ return $e } }
    return $null
  }
  try {
    if ($Folder.PSObject -and ($Folder.PSObject.TypeNames -contains 'System.__ComObject')) { return $Folder }
  } catch {}
  return $null
}

function Safe-Items {
  param([Parameter(Mandatory)][object]$Folder,[int]$Retries=3,[int]$DelayMs=150)
  $f = Normalize-Folder $Folder
  if (-not $f) { return @() }
  for($r=1;$r -le $Retries;$r++){
    try{
      $items = @($f.Items())
      if($items){ return $items }
    }catch{}
    Start-Sleep -Milliseconds $DelayMs
  }
  @()
}
function Safe-GetFolder {
  param([Parameter(Mandatory)][object]$FolderItem,[int]$Retries=3,[int]$DelayMs=150)
  $candidate = $FolderItem
  if ($candidate -is [object[]]) {
    foreach($e in $candidate){ if($e -is [__ComObject]){ $candidate = $e; break } }
  }
  for($r=1;$r -le $Retries;$r++){
    try{
      if($candidate -is [__ComObject]){
        $f = $candidate.GetFolder()
        if($f){ return $f }
      } elseif ($candidate -is [hashtable] -and $candidate.ContainsKey('Folder')) {
        if($candidate.Folder){ return (Normalize-Folder $candidate.Folder) }
      } elseif ($candidate.PSObject -and $candidate.PSObject.Properties['Folder']){
        $f = $candidate.Folder
        $f = Normalize-Folder $f
        if($f){ return $f }
      }
    }catch{}
    Start-Sleep -Milliseconds $DelayMs
  }
  $null
}
function Get-ShellFolder {
  param([Parameter(Mandatory)][string]$Path)
  if(-not (Test-Path -LiteralPath $Path)){ New-Item -ItemType Directory -Path $Path | Out-Null }
  $resolved = $Path; try { $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path } catch {}
  $shell  = New-Object -ComObject Shell.Application
  $folder = $shell.NameSpace($resolved)
  if(-not $folder -or $folder -isnot [__ComObject]){ throw ("Shell NameSpace failed for '{0}'." -f $resolved) }
  $folder
}
function _CountFiles($ns){
  $ns = Normalize-Folder $ns
  if(-not $ns){ return -1 }
  try { return (@($ns.Items() | Where-Object { -not (Test-IsFolder $_) })).Count } catch { return -1 }
}
function WarmUp-NameSpace { param([object]$ns,[int]$Passes=1)
  $ns = Normalize-Folder $ns
  if(-not $ns){ return }
  try { for($i=1;$i -le $Passes;$i++){ $null = @($ns.Items()) | Out-Null } } catch {}
}

# >>> VERBOSE WarmUp-MTPFolder <<<
function WarmUp-MTPFolder {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][object] $ShellFolder,
    [int]$Passes = 2,[int]$DelayMs = 200,[switch]$Recurse,[int]$MaxDepth = 2,[int]$ShowEvery = 200,[switch]$Quiet
  )
  $ShellFolder = Normalize-Folder $ShellFolder
  if(-not $ShellFolder){ Write-Warn "[WarmUp] ShellFolder is not a COM folder (null/array without COM). Skipping."; return [pscustomobject]@{Files=0;Folders=0;Errors=0;Touched=0;Passes=$Passes;Elapsed=0} }

  function _Out([string]$s,[ConsoleColor]$c=[ConsoleColor]::Gray){ if(-not $Quiet){ Write-Host $s -ForegroundColor $c } }
  function _TouchDetails([__ComObject]$folder,$item){ try{ $null=$folder.GetDetailsOf($item,0) }catch{}; try{ $null=$folder.GetDetailsOf($item,1) }catch{}; try{ $null=$folder.GetDetailsOf($item,3) }catch{} }
  function _EnumerateOnce([__ComObject]$folder,[int]$pass,[int]$depth,[ref]$stats,[int]$showEvery){
    $sw=[System.Diagnostics.Stopwatch]::StartNew(); $name=try{$folder.Title}catch{"<unknown>"}; _Out ("[WarmUp] Pass {0} • Depth {1} • Folder: {2}" -f $pass,$depth,$name) DarkCyan
    $items = Safe-Items -Folder $folder
    $count = $items.Count; $stats.Value.Touched += $count
    for($i=0;$i -lt $count;$i++){
      $it=$items[$i]; try{ _TouchDetails $folder $it; if(Test-IsFolder $it){$stats.Value.Folders++}else{$stats.Value.Files++} }catch{$stats.Value.Errors++}
      if( ((($i+1)%$showEvery) -eq 0) -or $i -eq ($count-1)){
        if(-not $Quiet){ Write-Host ("  → Pass {0}/{1} [{2}/{3}] | Files:{4} Folders:{5} Errors:{6}" -f $pass,$Passes,($i+1),$count,$stats.Value.Files,$stats.Value.Folders,$stats.Value.Errors) -ForegroundColor DarkGray }
        Write-Progress -Activity "Enumerating MTP (Pass $pass/$Passes)" -Status ("{0} items scanned in '{1}'" -f ($i+1),$name) -PercentComplete ((($i+1)*100.0)/[Math]::Max(1,$count))
      }
    }
    $sw.Stop(); _Out ("[WarmUp] Done • {0} items • {1:N0} ms" -f $count,$sw.Elapsed.TotalMilliseconds) DarkCyan; return $items
  }
  $total=[pscustomobject]@{Files=0;Folders=0;Errors=0;Touched=0;Passes=$Passes;Elapsed=0}; $overall=[System.Diagnostics.Stopwatch]::StartNew()
  _Out ("=== WarmUp-MTPFolder: start • {0} passes • Recurse:{1} (MaxDepth={2}) ===" -f $Passes,([bool]$Recurse),$MaxDepth) Cyan
  HB-Start "Enumerating device (MTP)" "Warming folder..."
  for($p=1;$p -le $Passes;$p++){
    _Out ("[WarmUp] Attempt {0}/{1}" -f $p,$Passes) Green
    $queue = New-Object System.Collections.Generic.Queue[object]; $queue.Enqueue([pscustomobject]@{Folder=$ShellFolder;Depth=0})
    while($queue.Count -gt 0){
      $node=$queue.Dequeue(); $folder=$node.Folder; $depth=[int]$node.Depth
      $folder = Normalize-Folder $folder
      if(-not $folder){ _Out "[WarmUp] (skip null/non-COM folder node)" DarkYellow; continue }
      $items=_EnumerateOnce -folder $folder -pass $p -depth $depth -stats ([ref]$total) -showEvery $ShowEvery
      if($Recurse -and $depth -lt $MaxDepth){
        foreach($it in $items){
          if(Test-IsFolder $it){ $sub=Safe-GetFolder -FolderItem $it; if($sub){ $queue.Enqueue([pscustomobject]@{Folder=$sub;Depth=$depth+1}) } }
        }
      }
    }
    if($DelayMs -gt 0 -and $p -lt $Passes){ Write-Progress -Activity "Cooling off between passes" -Status "$DelayMs ms" -PercentComplete 0; Start-Sleep -Milliseconds $DelayMs; Write-Progress -Activity "Cooling off between passes" -Completed }
  }
  $overall.Stop(); $total.Elapsed=[int][Math]::Round($overall.Elapsed.TotalMilliseconds)
  _Out ("=== WarmUp-MTPFolder: done • {0} ms total • Touched:{1} Files:{2} Folders:{3} Errors:{4} ===" -f $total.Elapsed,$total.Touched,$total.Files,$total.Folders,$total.Errors) Cyan
  HB-Stop; Write-Progress -Activity "Enumerating MTP" -Completed; return $total
}
# <<< END WarmUp-MTPFolder >>>

function Get-StableMtpItemCount{
  param([object]$ShellFolder,[int]$MaxRounds=6,[int]$WaitSec=2)
  $ShellFolder = Normalize-Folder $ShellFolder
  if(-not $ShellFolder){ return -1 }
  $prev=-1
  for($i=1;$i -le $MaxRounds;$i++){
    $items = Safe-Items -Folder $ShellFolder
    $count = $items.Count
    Write-Verbose ("[MTP] Round {0}/{1}: {2} items visible" -f $i,$MaxRounds,$count)
    if($count -ge 0 -and $count -eq $prev){ return $count }
    $prev = $count
    WarmUp-MTPFolder -ShellFolder $ShellFolder -Passes 1 -ShowEvery 100 -Recurse -MaxDepth 2
    Start-Sleep -Seconds $WaitSec
  }
  return $prev
}

# Diagnostics: list devices
function List-MTPDevices {
  try { $shell = New-Object -ComObject Shell.Application } catch { Write-Err "Shell.Application failed: $($_.Exception.Message)"; return }
  $pc = $shell.NameSpace('shell:MyComputerFolder'); if (-not $pc) { Write-Err "Cannot open 'This PC'."; return }
  Write-Info "---- Visible portable devices under 'This PC' ----"
  $items = @($pc.Items()); if (-not $items -or $items.Count -eq 0) { Write-Warn "[None]"; return }
  foreach($it in $items){
    try{
      $name = Test-Name $it
      if(-not (Test-IsFolder $it)){ continue }
      if ($name -notmatch 'iPhone|Apple|Phone|Android|Galaxy|Pixel') { continue }
      Write-Host ("• {0}" -f $name) -ForegroundColor Magenta
      try {
        $f=$it.GetFolder()
        if($f){
          foreach($k in (Safe-Items -Folder $f)){ Write-Host ("   - {0}" -f (Test-Name $k)) -ForegroundColor DarkGray }
        } else { Write-Host "   (GetFolder() returned null)" -ForegroundColor DarkYellow }
      } catch { Write-Host ("   (error reading children: {0})" -f $_.Exception.Message) -ForegroundColor DarkYellow }
    } catch {}
  }
  Write-Info "-----------------------------------------------"
}

# Robust CopyHere/MoveHere with warmups (stay no-UI)
function CopyTo-ShellFolder {
  param(
    [Parameter(Mandatory)][object]$ParentShellFolder,
    [Parameter(Mandatory)]$ShellItem,
    [Parameter(Mandatory)][string]$DestFolder,
    [string]$Label = "(unknown)",
    [switch]$DeleteSource
  )
  $ParentShellFolder = Normalize-Folder $ParentShellFolder
  if(-not $ParentShellFolder){ throw "[Copy] ParentShellFolder is not a COM folder." }

  $destNS = Get-ShellFolder -Path $DestFolder
  Write-Verbose ("[Copy] Dest NameSpace: {0}" -f $destNS.Self.Path)

  $srcItem = $ShellItem
  if ($srcItem -isnot [__ComObject]) {
    $srcItem = $ParentShellFolder.ParseName([string]$ShellItem)
    if (-not $srcItem) { throw ("Could not parse source item '{0}'." -f [string]$ShellItem) }
  }

  $method = if($DeleteSource){ 'MoveHere' } else { 'CopyHere' }

  $beforeCount = _CountFiles $destNS
  $beforeBytes = 0
  try { $beforeBytes = (@($destNS.Items() | Where-Object { -not (Test-IsFolder $_) }) | Measure-Object Length -Sum).Sum } catch {}
  if (-not $beforeBytes) { $beforeBytes = 0 }
  Write-Verbose ("[Copy] Pre-count: {0} | Pre-bytes: {1} | File: {2}" -f $beforeCount, $beforeBytes, $Label)

  # Always prefer no-UI flags
  $FOF_SILENT=0x0004; $FOF_NOCONFIRMATION=0x0010; $FOF_NOCONFIRMMKDIR=0x0200; $FOF_NOERRORUI=0x0400; $FOF_NOCOPYSECURITYATTRIBS=0x1000; $FOF_SIMPLEPROGRESS=0x0100
  $NOUI_FLAGS = [int]($FOF_SILENT -bor $FOF_NOCONFIRMATION -bor $FOF_NOCONFIRMMKDIR -bor $FOF_NOERRORUI -bor $FOF_NOCOPYSECURITYATTRIBS)
  $ALT_FLAGS  = [int]($FOF_SILENT -bor $FOF_NOCONFIRMATION -bor $FOF_NOERRORUI -bor $FOF_NOCOPYSECURITYATTRIBS -bor $FOF_SIMPLEPROGRESS)

  $copied=$false
  foreach($flags in @($NOUI_FLAGS,$ALT_FLAGS)){
    Write-Verbose ("[Copy] Attempt with flags 0x{0:X} ({1})" -f $flags,$Label)
    $attemptStart=Get-Date
    try{
      [void]$destNS.GetType().InvokeMember($method,'InvokeMethod',$null,$destNS,@($srcItem,$flags))
      $perAttemptSec=[Math]::Min([Math]::Max(3,[int]($ConfirmMaxWaitSec/3)),10)
      $deadline=(Get-Date).AddSeconds($perAttemptSec)
      while((Get-Date) -lt $deadline -and -not $copied){
        $current=_CountFiles $destNS
        if($current -gt $beforeCount){ $copied=$true; break }
        $ms=[int]((Get-Date)-$attemptStart).TotalMilliseconds
        Write-Verbose ("[Copy] Waiting ({0}, elapsed={1}ms)... dest files: {2} → {3}" -f $Label,$ms,$current,$beforeCount)
        Start-Sleep -Milliseconds ([Math]::Max(100,$ConfirmPollMs))
      }
      if($copied){ break }
    }catch{
      Write-Verbose ("[Copy] Shell copy failed for {0} with 0x{1:X}: {2}" -f $Label,$flags,$_.Exception.Message)
    }
  }

  if(-not $copied){ throw ("Shell {0} returned but no items appeared in '{1}' for {2}. Likely elevated session, device prompt not granted, or MTP quirk." -f $method,$DestFolder,$Label) }
}

# --------------------------- Metadata helpers ---------------------------
function Get-ItemFileNameWithExt { param([__ComObject]$Parent,[__ComObject]$Item)
  try{ $n=[string]$Item.Name; if($n -and $n.Contains('.')){ return $n } }catch{}
  try{ $d0=$Parent.GetDetailsOf($Item,0); if($d0 -and $d0.Contains('.')){ return $d0 } }catch{}
  try{ return [string]$Item.Name }catch{ return $null }
}
function Get-DisplayNameWithExtension {
  param([__ComObject]$Parent,[__ComObject]$Item)
  try {
    $nm = $null; try { $nm = [string]$Item.Name } catch {}
    $disp = $null; try { $disp = [string]$Parent.GetDetailsOf($Item,0) } catch {}
    $ext = $null
    try { $ext = [string]$Item.ExtendedProperty('System.FileExtension') } catch {}
    if ([string]::IsNullOrWhiteSpace($ext)) {
      try {
        $itype = [string]$Item.ExtendedProperty('System.ItemType')
        if ($itype -and $itype.StartsWith('.')) { $ext = $itype }
      } catch {}
    }
    if ($nm -and $nm.Contains('.')) { return $nm }
    if ($disp -and $disp.Contains('.')) { return $disp }
    if ($nm -and $ext) { $ext = $ext.Trim(); if ($ext.StartsWith('.')) { return "$nm$ext" } else { return "$nm.$ext" } }
    if ($disp -and $ext) { $ext = $ext.Trim(); if ($ext.StartsWith('.')) { return "$disp$ext" } else { return "$disp.$ext" } }
    if ($nm) { return $nm }
    if ($disp) { return $disp }
  } catch {}
  return $null
}
function Get-LowerExtFromName{ param([string]$n)
  if([string]::IsNullOrWhiteSpace($n)){ return $null }
  $p=$n.LastIndexOf('.'); if($p -ge 0 -and $p -lt ($n.Length-1)){ return $n.Substring($p+1).ToLowerInvariant() }
  $null
}
function Classify-Item{ param([__ComObject]$i,[string]$nameWithExt)
  $e = Get-LowerExtFromName $nameWithExt
  if($IMAGE_EXT -contains $e){ return 'Image' }
  if($VIDEO_EXT -contains $e){ return 'Video' }
  'Image'
}
function Get-ItemDate{ param([__ComObject]$p,[__ComObject]$i)
  try{ $d=$p.GetDetailsOf($i,12); if([string]::IsNullOrWhiteSpace($d)){ $d=$p.GetDetailsOf($i,3) }; if(-not [string]::IsNullOrWhiteSpace($d)){ return [datetime]::Parse($d) } }catch{}
  (Get-Date)
}
function Get-MediaBasePath { param([string]$Root,[string]$Kind) if($Kind -eq 'Video'){ Join-Path $Root 'Videos' } else { Join-Path $Root 'Images' } }
function Get-DestinationFolder{
  param([string]$Root,[string]$Kind,[datetime]$When,[switch]$NoDateFolders)
  $base = Get-MediaBasePath -Root $Root -Kind $Kind
  $path = if($NoDateFolders){ $base } else { Join-Path (Join-Path $base $When.ToString('yyyy')) $When.ToString('yyyy-MM') }
  if(-not (Test-Path $path)){ New-Item -ItemType Directory -Path $path | Out-Null }
  $path
}
function Test-AlreadyExists{
  param([string]$DestFolder,[string]$FileName,[Nullable[long]]$SizeBytes,[switch]$ByNameOnly)
  if(-not $FileName){ return $false }
  $target=Join-Path $DestFolder $FileName
  if(-not (Test-Path $target)){ return $false }
  if($ByNameOnly){ return $true }
  try{ $fi=Get-Item -LiteralPath $target -ErrorAction Stop; if($SizeBytes.HasValue -and $fi.Length -eq $SizeBytes.Value){ return $true } }catch{}
  $true
}
function Get-ItemSizeBytes{ param([__ComObject]$p,[__ComObject]$i)
  try{
    $s=$p.GetDetailsOf($i,1)
    if(-not [string]::IsNullOrWhiteSpace($s)){
      $num = ($s -replace '[^\d\,\.]',''); if($num -match ','){ $num = $num -replace ',','.' }
      if($s -match 'KB'){ return [long]([double]$num*1KB) }
      if($s -match 'MB'){ return [long]([double]$num*1MB) }
      if($s -match 'GB'){ return [long]([double]$num*1GB) }
      return [long]$num
    }
  }catch{}
}

# --------------------------- Resolver ---------------------------
function Resolve-DeviceRoot {
  param([string]$DeviceName,[string]$StorageFolderName = 'Internal Storage')

  Write-Info "[Resolver] DeviceName='$DeviceName' | StorageFolderName='$StorageFolderName'"

  try { $shell = New-Object -ComObject Shell.Application } catch { throw "Shell.Application COM failed: $($_.Exception.Message)" }
  $pc = $shell.NameSpace('shell:MyComputerFolder'); if (-not $pc) { throw "Cannot open 'This PC'." }

  $all = @($pc.Items()); if (-not $all -or $all.Count -eq 0) { throw "No items in 'This PC'." }

  $dev = $all | Where-Object { (Test-Name $_) -match [Regex]::Escape($DeviceName) } | Select-Object -First 1
  if (-not $dev) { $dev = $all | Where-Object { (Test-Name $_) -match 'iPhone|Apple' } | Select-Object -First 1 }
  if (-not $dev) { List-MTPDevices; throw "Device '$DeviceName' not found. See list above." }

  $devName = Test-Name $dev; Write-Info ("[Resolver] Matched device: {0}" -f $devName)
  $folder = $null; try{ $folder = $dev.GetFolder() }catch{}
  if(-not $folder){ List-MTPDevices; throw "Device '$devName' detected, but GetFolder() returned null. Unlock/tap **Trust**." }

  $children = Safe-Items -Folder $folder
  if($children.Count -eq 0){ List-MTPDevices; throw "Device '$devName' has no visible children. Keep phone unlocked." }

  $internal = $children | Where-Object { (Test-IsFolder $_) -and ((Test-Name $_) -eq $StorageFolderName) } | Select-Object -First 1
  if (-not $internal) { $internal = $children | Where-Object { (Test-IsFolder $_) -and ((Test-Name $_) -match 'Internal|Storage') } | Select-Object -First 1 }
  if (-not $internal) { $internal = $children | Where-Object { (Test-IsFolder $_) } | Select-Object -First 1 }
  if (-not $internal) { List-MTPDevices; throw "No storage folder under '$devName'." }

  $internalName = Test-Name $internal; Write-Info ("[Resolver] Using storage folder: {0}" -f $internalName)
  $internalFolder = Safe-GetFolder -FolderItem $internal
  if (-not $internalFolder) { throw "'$internalName' handle is null (trust/lock/elevation)." }

  try { Write-Verbose "[Resolver] Priming storage (warmup)..." ; WarmUp-MTPFolder -ShellFolder $internalFolder -Passes 3 -ShowEvery 100 -Recurse -MaxDepth 2 } catch { Write-Verbose "[Resolver] Warmup error (continuing): $($_.Exception.Message)" }

  $dcim = $null; try { $dcim = (Safe-Items -Folder $internalFolder) | Where-Object { (Test-IsFolder $_) -and ((Test-Name $_) -eq 'DCIM') } | Select-Object -First 1 } catch {}
  if ($dcim) {
    $dcimFolder = Safe-GetFolder -FolderItem $dcim
    if ($dcimFolder) { Write-Info "[Resolver] Found DCIM (fast path)."; return [pscustomobject]@{ Root = $dcimFolder; Mode = 'DCIM'; Candidates = @() } }
  }

  $cands = @(); try { $cands = (Safe-Items -Folder $internalFolder) | Where-Object { Test-IsFolder $_ } } catch {}
  Write-Info ("[Resolver] DCIM not found; falling back to '{0}' root with {1} candidate folder(s)." -f $internalName,$cands.Count)
  [pscustomobject]@{ Root = $internalFolder; Mode = 'Internal'; Candidates = $cands }
}

function Get-FolderSummary { param([object]$Folder,[int]$TimeoutSec)
  $Folder = Normalize-Folder $Folder
  if(-not $Folder){ return @{SubFolderCount=0;FileCount=0} }
  WarmUp-MTPFolder -ShellFolder $Folder -Passes 2 -ShowEvery 100 -Recurse -MaxDepth 2
  [void](Get-StableMtpItemCount -ShellFolder $Folder -MaxRounds 4 -WaitSec ([Math]::Max([int]($TimeoutSec/8),1)))
  $subFolders=@(Safe-Items -Folder $Folder | Where-Object { Test-IsFolder $_ })
  $files=@(Safe-Items -Folder $Folder | Where-Object { -not (Test-IsFolder $_) })
  @{ SubFolderCount=$subFolders.Count; FileCount=$files.Count }
}
function Show-Structure { param([object]$RootFolder,[string]$Mode,[object[]]$Candidates,[int]$TimeoutSec=20)
  $RootFolder = Normalize-Folder $RootFolder
  Write-Host ("----- STRUCTURE (mode:  {0}) -----" -f $Mode) -ForegroundColor Magenta
  $subs = if($Mode -eq 'DCIM'){ @(Safe-Items -Folder $RootFolder | Where-Object { Test-IsFolder $_ }) }
          elseif ($null -ne $Candidates){ @($Candidates | Where-Object { $_ -and (Test-IsFolder $_) }) }
          else { @(Safe-Items -Folder $RootFolder | Where-Object { Test-IsFolder $_ }) }
  $rootFiles=@(Safe-Items -Folder $RootFolder | Where-Object { -not (Test-IsFolder $_) })
  if($rootFiles.Count -gt 0){ Write-Host ("[root-files] {0} item(s) at root" -f $rootFiles.Count) -ForegroundColor DarkCyan }
  if(-not $subs -or $subs.Count -eq 0){ Write-Host "[no subfolders]" -ForegroundColor DarkYellow; return }
  for($i=0;$i -lt $subs.Count;$i++){
    $it=$subs[$i]; $name=Test-Name $it; $f=Safe-GetFolder -FolderItem $it; if(-not $f){ Write-Warn ("[diag] cannot GetFolder() for '{0}', skipping" -f $name); continue }
    $sum=Get-FolderSummary -Folder $f -TimeoutSec $TimeoutSeconds
    Write-Host ("[{0}/{1}] {2}  -  {3} files, {4} subfolders" -f ($i+1),$subs.Count,$name,$sum.FileCount,$sum.SubFolderCount) -ForegroundColor Gray
  }
  Write-Host "----- END STRUCTURE -----" -ForegroundColor Magenta
}

# --------------------------- NEW: Reopen + name list + smart ParseName ---------------------------
function Reopen-ChildFolderByName {
  param(
    [Parameter(Mandatory)][object]$ParentFolder,
    [Parameter(Mandatory)][string]$ChildName
  )
  $ParentFolder = Normalize-Folder $ParentFolder
  if(-not $ParentFolder){ return $null }
  try {
    $kids = Safe-Items -Folder $ParentFolder
    foreach($k in $kids){
      if(Test-IsFolder $k){
        $nm = Test-Name $k
        if($nm -eq $ChildName){
          $f = Safe-GetFolder -FolderItem $k
          if($f){ return $f }
        }
      }
    }
  } catch {}
  $null
}
function Get-SafeFileNameList {
  param([Parameter(Mandatory)][object]$Folder)
  $Folder = Normalize-Folder $Folder
  if(-not $Folder){ return @() }
  $list = New-Object System.Collections.Generic.List[string]
  try {
    $items = Safe-Items -Folder $Folder
    foreach($it in $items){
      if(-not (Test-IsFolder $it)){
        $nm = Get-DisplayNameWithExtension -Parent $Folder -Item $it
        if([string]::IsNullOrWhiteSpace($nm)){ $nm = Get-ItemFileNameWithExt -Parent $Folder -Item $it }
        if([string]::IsNullOrWhiteSpace($nm)){ $nm = Test-Name $it }
        if(-not [string]::IsNullOrWhiteSpace($nm)){ [void]$list.Add($nm) }
      }
    }
  } catch {}
  ,$list
}
function Try-ParseNameSmart {
  param(
    [Parameter(Mandatory)][object]$Folder,
    [Parameter(Mandatory)][string]$NameMaybeStem
  )
  $Folder = Normalize-Folder $Folder
  if(-not $Folder){ return $null }
  if ($NameMaybeStem.Contains('.')) {
    try { $it = $Folder.ParseName($NameMaybeStem); if($it){ return $it } } catch {}
    return $null
  }
  foreach($e in $KNOWN_EXT_ORDER){
    $cand = "$NameMaybeStem.$e"
    try { $it = $Folder.ParseName($cand); if($it){ return $it } } catch {}
  }
  $null
}

# --------------------------- NEW: Free-space + checkpoint helpers ---------------------------
function Get-DriveRoot([string]$path){
  try { return [IO.Path]::GetPathRoot((Resolve-Path $path -ErrorAction Stop).Path) } catch { return $null }
}
function Assert-FreeSpace([string]$path,[int]$minGb){
  try{
    $root = Get-DriveRoot $path
    if(-not $root){ return }
    $d = New-Object System.IO.DriveInfo($root)
    $freeGb = [math]::Floor($d.AvailableFreeSpace/1GB)
    if($freeGb -lt $minGb){ Write-Warn "[Warn] Less than $minGb GB free on $($d.Name) (available: $freeGb GB) — continuing but may fail." }
  }catch{}
}
function To-Hashtable([object]$obj){
  if ($obj -is [hashtable]) { return $obj }
  if ($null -eq $obj) { return @{} }
  $ht=@{}
  foreach($p in $obj.PSObject.Properties){
    $val = $p.Value
    if ($val -is [System.Collections.IDictionary]) { $ht[$p.Name] = To-Hashtable $val }
    elseif ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string])) { $ht[$p.Name] = @($val) }
    else { $ht[$p.Name] = $val }
  }
  return $ht
}
function Load-Checkpoint { param($p)
  if(Test-Path $p){
    try{
      $o = Get-Content $p -Raw | ConvertFrom-Json
      $ht = To-Hashtable $o
      if(-not $ht.ContainsKey('Folders')){ $ht['Folders'] = @{} }
      if(-not ($ht['Folders'] -is [hashtable])){ $ht['Folders'] = To-Hashtable $ht['Folders'] }
      return $ht
    }catch{}
  }
  return @{ Folders = @{} }
}
function Save-Checkpoint { param($p,$obj)
  try{ $obj | ConvertTo-Json -Depth 6 | Set-Content -Path $p -Encoding UTF8 }catch{}
}

# --------------------------- Main ---------------------------
Write-Info "===== START: CopyFromIPhone.ps1 (MAX+ROBUST+STREAM) ====="

try {
  $IsElevated = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
  ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if ($IsElevated) { Write-Warn "[Warn] Session is elevated (Admin). MTP copies can fail in elevated windows. Use a normal console if possible." }
} catch {}

$imgRoot = Join-Path $Destination 'Images'
$vidRoot = Join-Path $Destination 'Videos'
foreach($p in @($Destination,$imgRoot,$vidRoot,$StageDir)){ if(-not (Test-Path $p)){ New-Item -ItemType Directory -Path $p | Out-Null } }

Assert-FreeSpace -path $StageDir -minGb $MinFreeGB
Assert-FreeSpace -path $Destination -minGb $MinFreeGB

$t = [Diagnostics.Stopwatch]::StartNew()

$cp = if($Checkpoint){ Load-Checkpoint -p $CheckpointPath } else { @{ Folders = @{} } }

try{
  $root = $null
  try { $root = Resolve-DeviceRoot -DeviceName $DeviceName -StorageFolderName $StorageFolderName }
  catch {
    Write-Err ("[Resolver] {0}" -f $_.Exception.Message)
    Write-Err "Hints: unlock iPhone, tap **Trust**, use a non-admin PowerShell, direct USB port/original cable."
    throw
  }

  if (-not $root -or -not $root.Root) {
    List-MTPDevices
    throw "[Resolver] Device resolved but no usable Root folder was returned. This is typically trust/lock/elevation. See list above."
  }

  $rf   = Normalize-Folder $root.Root
  if(-not $rf){ throw "[Resolver] Root is not a COM folder handle." }

  $modeText = if ($root.Mode) { [string]$root.Mode } else { '(unknown)' }
  Write-Host ("[Info] Root mode:  {0}" -f $modeText) -ForegroundColor Magenta

  $subs = @()
  if($modeText -eq 'DCIM'){
    $subs = @(Safe-Items -Folder $rf | Where-Object { Test-IsFolder $_ })
  } elseif ($root.PSObject.Properties['Candidates'] -and $root.Candidates) {
    $allKids = @($root.Candidates)
    $subs = @($allKids | Where-Object { $_ -and (Test-IsFolder $_) })
  } else {
    $allKids = @(Safe-Items -Folder $rf)
    $subs = @($allKids | Where-Object { Test-IsFolder $_ })
  }

  # Newest-first by name; optional filter/limit
  $subs = $subs | Sort-Object { Test-Name $_ } -Descending
  if ($OnlyFolderName) { $subs = $subs | Where-Object { (Test-Name $_) -eq $OnlyFolderName } }
  if ($MaxFolders -gt 0) { $subs = $subs | Select-Object -First $MaxFolders }

  if($ListStructure -or $ListOnly){
    Show-Structure -RootFolder $rf -Mode $modeText -Candidates $subs -TimeoutSec $TimeoutSeconds
    if($ListOnly){ return }
  }

  if(-not $subs -or $subs.Count -eq 0){ Write-Warn "[MTP] No subfolders; using root."; $subs=@() }
  $rootFiles=@(Safe-Items -Folder $rf | Where-Object { -not (Test-IsFolder $_) })
  if($rootFiles.Count -gt 0){
    Write-Verbose ("[MTP] Found {0} files at root." -f $rootFiles.Count)
    $subs = @(@{Name='(root)'; Folder=$rf; Direct=$true}) + $subs
  }

  # ======================= PER-FOLDER STREAMING =======================
  $tot=0; $ok=0; $skip=0; $fail=0
  $folderIdx=0; $folderCount=[Math]::Max(1,$subs.Count)
  $emptyRounds=0

  foreach($sf in $subs){
    try { $null = Safe-Items -Folder $rf } catch { throw "Device removed or locked (MTP root became unavailable)." }

    $folderIdx++
    $sfName=''; $folder=$null; $isRoot=$false

    if(($sf -is [hashtable] -and $sf.ContainsKey('Direct') -and $sf.Direct) -or ($sf.PSObject.Properties['Direct'] -and $sf.Direct)){
      $sfName='(root)'; $folder=$rf; $isRoot=$true
    } else {
      $sfName = try { Test-Name $sf } catch { "(unknown)" }
      if($sf -and (Test-IsFolder $sf)){ $folder = Safe-GetFolder -FolderItem $sf }
    }

    $folder = Normalize-Folder $folder
    if(-not $folder){
      Write-Warn ("[Scan] Skip: cannot open folder for '{0}' (GetFolder() null)" -f $sfName)
      continue
    }

    # Checkpoint restore for this folder
    $doneSet = @{}
    if($Checkpoint){
      $folderMap = $cp.Folders
      if ($folderMap -and $folderMap.ContainsKey($sfName)) {
        foreach($nm in @($folderMap[$sfName])){ $doneSet[$nm] = $true }
      }
    }

    # ===== PER-FOLDER WARMUP =====
    Write-Verbose ("[Scan] Folder {0}/{1}: {2}" -f $folderIdx,$folderCount,$sfName)
    $pct=[int]( ($folderIdx/[double]$folderCount)*100 )
    Write-Progress -Id 10 -Activity "Scanning device folders" -Status ("Folder: {0}" -f $sfName) -PercentComplete $pct

    HB-Start "Enumerating device (MTP)" ("Reading: {0}" -f $sfName); Spin-Start ("Reading: {0}" -f $sfName)
    WarmUp-MTPFolder -ShellFolder $folder -Passes 2 -ShowEvery 100 -Recurse -MaxDepth 2
    [void](Get-StableMtpItemCount -ShellFolder $folder -MaxRounds 4 -WaitSec ([Math]::Max([int]($TimeoutSeconds/8),1)))
    Spin-Stop; HB-Stop

    # ===== BUILD NAME LIST =====
    $fileNames = Get-SafeFileNameList -Folder $folder
    if($fileNames.Count -eq 0){ $emptyRounds++ } else { $emptyRounds=0 }
    if($emptyRounds -ge 3){ throw "Device appears unavailable (no items visible multiple rounds). Is it unplugged or locked?" }

    Write-Verbose ("[Scan] {0}: discovered {1} files" -f $sfName,$fileNames.Count)

    # ===== COPY BY NAME; RE-OPEN THE FOLDER HANDLE EACH TIME =====
    $fileIdx=0
    foreach($fn in $fileNames){
      $fileIdx++; $tot++

      if(-not $isRoot){
        $folder = Reopen-ChildFolderByName -ParentFolder $rf -ChildName $sfName
        if(-not $folder){
          Write-Err ("[Error] Folder '{0}' vanished (cannot reopen). Skipping remaining files in this folder." -f $sfName)
          $fail += ($fileNames.Count - $fileIdx + 1)
          break
        }
      } else {
        WarmUp-MTPFolder -ShellFolder $folder -Passes 1 -ShowEvery 200 -Recurse -MaxDepth 1 -Quiet
      }

      # Resolve COM item
      $it = $null
      try { $it = $folder.ParseName($fn) } catch {}
      if(-not $it){ $it = Try-ParseNameSmart -Folder $folder -NameMaybeStem $fn }
      if(-not $it){
        Write-Warn ("[Skip] Could not ParseName() '{0}' in '{1}' (stale view or missing extension)." -f $fn,$sfName)
        $skip++; continue
      }

      # Final name with extension
      $fnResolved = $fn
      if(-not $fnResolved.Contains('.')){
        try { $disp = Get-DisplayNameWithExtension -Parent $folder -Item $it; if($disp){ $fnResolved = $disp } } catch {}
      }
      if ($StrictExtensions -and -not $fnResolved.Contains('.')) {
        Write-Warn ("[Skip] No extension resolved for {0} in '{1}' (StrictExtensions on)." -f $fn,$sfName)
        $skip++; if($YieldMs -gt 0){ Start-Sleep -Milliseconds $YieldMs }; continue
      }

      # Checkpoint guard
      if($Checkpoint -and $doneSet.ContainsKey($fnResolved)){
        Write-Host ("[Skip] {0} (checkpoint)" -f $fnResolved) -ForegroundColor DarkYellow
        $skip++; if($YieldMs -gt 0){ Start-Sleep -Milliseconds $YieldMs }; continue
      }

      $k = Classify-Item -i $it -nameWithExt $fnResolved
      $d = Get-ItemDate $folder $it
      $finalDest = Get-DestinationFolder -Root $Destination -Kind $k -When $d -NoDateFolders:$NoDateFolders
      $s = Get-ItemSizeBytes $folder $it

      $pctFile=[int]( ($fileIdx/[Math]::Max(1,$fileNames.Count))*100 )
      $msg = "[{0}/{1}] {2}: {3}/{4}  {5}" -f $folderIdx,$folderCount,$sfName,$fileIdx,$fileNames.Count,$fnResolved
      Write-Progress -Id 11 -Activity ("Copying {0} files" -f $k) -Status $msg -PercentComplete $pctFile

      # -------- Final destination duplicate guard (unchanged) --------
      if(Test-AlreadyExists -DestFolder $finalDest -FileName $fnResolved -SizeBytes $s -ByNameOnly:$SkipByNameOnly){
        $existsSuffix = if($SkipByNameOnly){ '' } else { '/size' }
        Write-Host ("[Skip] {0} → {1} (exists{2})" -f $fnResolved,$finalDest,$existsSuffix) -ForegroundColor DarkYellow
        $skip++; if($YieldMs -gt 0){ Start-Sleep -Milliseconds $YieldMs }; continue
      }

      # Stage dir per kind
      $stageKindDir = Get-MediaBasePath -Root $StageDir -Kind $k
      if(-not (Test-Path $stageKindDir)){ New-Item -ItemType Directory -Path $stageKindDir | Out-Null }

      # -------- NEW: Stage duplicate guard to prevent prompts --------
      $stageCandidate = Join-Path $stageKindDir $fnResolved
      if(Test-Path $stageCandidate){
        if($SkipByNameOnly){
          Write-Host ("[Skip] (staged) {0} → {1} (exists)" -f $fnResolved,$stageKindDir) -ForegroundColor DarkYellow
          $skip++; if($YieldMs -gt 0){ Start-Sleep -Milliseconds $YieldMs }; continue
        } else {
          try{
            $st=Get-Item -LiteralPath $stageCandidate -ErrorAction Stop
            if($s.HasValue -and $st.Length -eq $s.Value){
              Write-Host ("[Skip] (staged/size) {0} → {1} (exists/size)" -f $fnResolved,$stageKindDir) -ForegroundColor DarkYellow
              $skip++; if($YieldMs -gt 0){ Start-Sleep -Milliseconds $YieldMs }; continue
            } else {
              # stale partial from prior run — clean it so shell copy won't prompt
              Write-Verbose "[Stage] Removing stale staged file with mismatched size."
              Remove-Item -LiteralPath $stageCandidate -Force -ErrorAction SilentlyContinue
            }
          }catch{
            # can't read size — safest is remove and proceed
            Remove-Item -LiteralPath $stageCandidate -Force -ErrorAction SilentlyContinue
          }
        }
      }

      try {
        # Shell copy to Stage (no-UI flags inside)
        CopyTo-ShellFolder -ParentShellFolder $folder -ShellItem $it -DestFolder $stageKindDir -Label $fnResolved -DeleteSource:$DeleteAfterCopy

        # Confirm arrival in Stage
        $deadline    = (Get-Date).AddSeconds([Math]::Max($TimeoutSeconds,$ConfirmMaxWaitSec))
        $arrived     = $false
        $stageFile   = $null
        $poll        = [Math]::Max(100,$ConfirmPollMs)
        $beforeList  = @(Get-ChildItem -LiteralPath $stageKindDir -File -ErrorAction SilentlyContinue)
        $beforeCount = $beforeList.Count
        $beforeBytes = ($beforeList | Measure-Object Length -Sum).Sum; if(-not $beforeBytes){ $beforeBytes = 0 }

        $t0    = Get-Date
        $ticks = 0

        while((Get-Date) -lt $deadline -and -not $arrived){
          Start-Sleep -Milliseconds $poll
          $ticks++
          $elapsed = [int]((Get-Date) - $t0).TotalMilliseconds

          $candidate = Join-Path $stageKindDir $fnResolved
          if (Test-Path $candidate) { $stageFile = $candidate; $arrived = $true; Write-Verbose ("[Stage] ({0}) name-hit: {1}" -f $fnResolved, $candidate); break }

          $nowList   = @(Get-ChildItem -LiteralPath $stageKindDir -File -ErrorAction SilentlyContinue)
          $nowCount  = $nowList.Count
          $nowBytes  = ($nowList | Measure-Object Length -Sum).Sum; if(-not $nowBytes){ $nowBytes = 0 }
          $countDelta = $nowCount - $beforeCount
          $byteDelta  = $nowBytes - $beforeBytes
          Write-Verbose ("[Stage] ({0}) tick {1}, {2}ms — deltas: count {3}, bytes {4}" -f $fnResolved,$ticks,$elapsed,$countDelta,$byteDelta)

          if ($countDelta -gt 0 -or $byteDelta -gt 0) {
            $newestAny = $nowList | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($newestAny) { $stageFile = $newestAny.FullName; $arrived = $true; Write-Verbose ("[Stage] ({0}) folder-delta confirm: picked newest {1}" -f $fnResolved, $stageFile); break }
          }
        }

        if(-not $arrived){
          Write-Warn ("[Warn] Could not confirm STAGED copy for {0}" -f $fnResolved)
          $fail++; if($YieldMs -gt 0){ Start-Sleep -Milliseconds $YieldMs }; continue
        }

        $finalName = [IO.Path]::GetFileName($stageFile)
        $finalPath = Join-Path $finalDest $finalName

        # Final duplicate check again (belt & braces)
        if($SkipByNameOnly){
          if(Test-Path $finalPath){
            Write-Host ("[Skip] {0} → {1} (exists)" -f $finalName,$finalDest) -ForegroundColor DarkYellow
            Remove-Item -LiteralPath $stageFile -Force -ErrorAction SilentlyContinue
            $skip++; if($YieldMs -gt 0){ Start-Sleep -Milliseconds $YieldMs }; continue
          }
        } else {
          if(Test-Path $finalPath){
            try{
              $fi=Get-Item -LiteralPath $finalPath -ErrorAction Stop
              $st=Get-Item -LiteralPath $stageFile -ErrorAction Stop
              if($fi.Length -eq $st.Length){
                Write-Host ("[Skip] {0} → {1} (exists/size)" -f $finalName,$finalDest) -ForegroundColor DarkYellow
                Remove-Item -LiteralPath $stageFile -Force -ErrorAction SilentlyContinue
                $skip++; if($YieldMs -gt 0){ Start-Sleep -Milliseconds $YieldMs }; continue
              }
            }catch{}
          }
        }

        Move-Item -LiteralPath $stageFile -Destination $finalPath -Force
        Write-Host ("[File] {0} → {1}" -f $finalName,$finalDest) -ForegroundColor Green
        $ok++

        if($Checkpoint){
          if(-not $cp.Folders.ContainsKey($sfName)){ $cp.Folders[$sfName] = @() }
          $cp.Folders[$sfName] += $finalName
          Save-Checkpoint -p $CheckpointPath -obj $cp
          $doneSet[$finalName] = $true
        }

        if(($ok + $skip + $fail) % [Math]::Max(1,$StageFlushEvery) -eq 0){
          try{
            Get-ChildItem -LiteralPath $StageDir -Recurse -File -ErrorAction SilentlyContinue |
              Where-Object { $_.LastWriteTime -lt (Get-Date).AddMinutes(-10) } |
              Remove-Item -Force -ErrorAction SilentlyContinue
          }catch{}
          [GC]::Collect()
          Start-Sleep -Milliseconds 50
        }
        if($YieldMs -gt 0){ Start-Sleep -Milliseconds $YieldMs }

      }
      catch{
        Write-Err ("[Error] {0} → {1} : {2}" -f $fnResolved,$finalDest,$_.Exception.Message)
        $fail++
        try{
          $pattern = ([IO.Path]::GetFileNameWithoutExtension($fnResolved) + '.*')
          $cands = Get-ChildItem -LiteralPath $stageKindDir -File -Filter $pattern -ErrorAction SilentlyContinue
          foreach($c in $cands){ if($c.LastWriteTime -gt (Get-Date).AddMinutes(-5)){ Remove-Item -LiteralPath $c.FullName -Force -ErrorAction SilentlyContinue } }
        }catch{}
        if($YieldMs -gt 0){ Start-Sleep -Milliseconds $YieldMs }
      }
    } # file loop
  } # folder loop

  Write-Success ("Done. Copied: {0}, Skipped: {1}, Failed: {2}, Total: {3}, Elapsed: {4:N1}s" -f $ok,$skip,$fail,$tot,$t.Elapsed.TotalSeconds)

}catch{
  Write-Err ("[FATAL] {0}" -f $_.Exception.Message)
}finally{
  $t.Stop()
  Write-Progress -Id 10 -Activity "Scanning device folders" -Completed
  Write-Progress -Id 11 -Activity "Copying files" -Completed
  HB-Stop
  Restore-ProgressPreference
  Write-Info "===== END: CopyFromIPhone.ps1 (MAX+ROBUST+STREAM) ====="
}
