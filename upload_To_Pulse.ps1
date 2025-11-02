<# upload_To_Pulse.ps1 (v2)
 - Runs Python generator
 - Waits for the expected monthly Excel to appear and be unlocked
 - Uploads to /work-essentials/apps/pulse-intranet/Documents
 - Updates /work-essentials/apps/pulse-intranet/Pages/Pulse-page-analytics.aspx
 - Logs to Optional\upload_To_Pulse.log
#>

[CmdletBinding()]
param(
  [string]$SiteUrl             = "https://pulse/work-essentials/apps/pulse-intranet",
  [string]$LibraryServerRelUrl = "/work-essentials/apps/pulse-intranet/Documents",
  [string]$PageServerRelUrl    = "/work-essentials/apps/pulse-intranet/Pages/Pulse-page-analytics.aspx",
  [string]$PythonScript        = "generate_Report.py",
  [int]$MonthsToUpdate         = 3,
  [switch]$PublishAfterUpload,
  [switch]$Overwrite
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$optionalDir = Join-Path $here "Optional"; New-Item -ItemType Directory -Force -Path $optionalDir | Out-Null
if (-not $PSBoundParameters.ContainsKey('PublishAfterUpload')) { $PublishAfterUpload = $true }
if (-not $PSBoundParameters.ContainsKey('Overwrite'))          { $Overwrite          = $true }

# ---- log
$logPath = Join-Path $optionalDir "upload_To_Pulse.log"
Start-Transcript -Path $logPath -Append | Out-Null
Write-Host "=== Pulse GA4 Upload Runner ===" -ForegroundColor Cyan
Write-Host "Working dir: $here"

# ---- previous full month & expected filename
$today = Get-Date
$firstThisMonth = Get-Date -Year $today.Year -Month $today.Month -Day 1
$lastPrev = $firstThisMonth.AddDays(-1)
$firstPrev = Get-Date -Year $lastPrev.Year -Month $lastPrev.Month -Day 1
$dateTag = "{0:yyyyMMdd}-{1:yyyyMMdd}" -f $firstPrev,$lastPrev
$expectedName = "Pulse-all-pages-report-$dateTag.xlsx"
$expectedRoot = Join-Path $here $expectedName
$expectedOptional = Join-Path $optionalDir $expectedName

# ---- helper: wait for file and ensure not locked
function Wait-FileReady {
  param([string]$Path,[int]$TimeoutSec=180)
  $sw=[Diagnostics.Stopwatch]::StartNew()
  while($sw.Elapsed.TotalSeconds -lt $TimeoutSec){
    if(Test-Path -LiteralPath $Path){
      try{
        $fs=[System.IO.File]::Open($Path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::ReadWrite,[System.IO.FileShare]::None)
        $fs.Close()
        return $true
      }catch{ Start-Sleep -Seconds 2 }
    }else{ Start-Sleep -Seconds 2 }
  }
  return $false
}

# ---- 1) run python generator
try{
  Push-Location $here
  Write-Host "Running: py -u $PythonScript" -ForegroundColor Yellow
  & py -u $PythonScript
} finally { Pop-Location }

# ---- 2) pick a file to upload (prefer expected name, check root then Optional)
$localFile = $null
if (Wait-FileReady -Path $expectedRoot -TimeoutSec 180) {
  $localFile = $expectedRoot
  Write-Host "Found expected report in ROOT: $localFile" -ForegroundColor Green
} elseif (Wait-FileReady -Path $expectedOptional -TimeoutSec 180) {
  $localFile = $expectedOptional
  Write-Host "Found expected report in OPTIONAL: $localFile" -ForegroundColor Green
} else {
  # Fallback: newest unlocked file
  $cand = Get-ChildItem -Path $here -Filter "Pulse-all-pages-report-*.xlsx" -File |
          Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if ($cand -and (Wait-FileReady -Path $cand.FullName -TimeoutSec 60)) {
    $localFile = $cand.FullName
    Write-Host "Using newest unlocked file: $localFile" -ForegroundColor Yellow
  } else {
    throw "Could not find an unlocked 'Pulse-all-pages-report-*.xlsx'. Close Excel and re-run."
  }
}
$targetName = $expectedName  # always upload with canonical name

# ---- 3) CSOM setup
$csomPath = "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI"
Add-Type -Path (Join-Path $csomPath "Microsoft.SharePoint.Client.dll")
Add-Type -Path (Join-Path $csomPath "Microsoft.SharePoint.Client.Runtime.dll")

$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$ctx.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

function Ensure-FolderPath {
  param([Microsoft.SharePoint.Client.ClientContext]$Ctx,[string]$ServerRelativeUrl)
  $web=$Ctx.Web; $Ctx.Load($web); $Ctx.ExecuteQuery()
  $parts=$ServerRelativeUrl.Trim('/').Split('/'); if($parts.Length -lt 2){throw "Bad ServerRelativeUrl: $ServerRelativeUrl"}
  $cur='/' + ($parts[0..1] -join '/')
  for($i=2;$i -lt $parts.Length;$i++){
    $cur="$cur/$($parts[$i])"
    try{ $f=$web.GetFolderByServerRelativeUrl($cur); $Ctx.Load($f); $Ctx.ExecuteQuery() }catch{
      $parentUrl=$cur.Substring(0,$cur.LastIndexOf('/')); $p=$web.GetFolderByServerRelativeUrl($parentUrl)
      $Ctx.Load($p); $Ctx.ExecuteQuery(); $nf=$p.Folders.Add($cur); $Ctx.Load($nf); $Ctx.ExecuteQuery()
    }
  }
  return $web.GetFolderByServerRelativeUrl($ServerRelativeUrl.TrimEnd('/'))
}

# validate lib
$lists=$ctx.Web.Lists; $ctx.Load($lists); $ctx.ExecuteQuery()
$docLibs=$lists | Where-Object{ $_.BaseTemplate -eq 101 -and -not $_.Hidden }
foreach($l in $docLibs){ $ctx.Load($l.RootFolder) }; $ctx.ExecuteQuery()
$roots=$docLibs | ForEach-Object{ $_.RootFolder.ServerRelativeUrl.TrimEnd('/') }
if(-not ($roots | Where-Object{ $LibraryServerRelUrl.TrimEnd('/').StartsWith($_,'OrdinalIgnoreCase') })){
  throw "Document library not found for: $LibraryServerRelUrl. Use one of:`n$($roots -join "`n")"
}

try{
  $folder=$ctx.Web.GetFolderByServerRelativeUrl($LibraryServerRelUrl.TrimEnd('/'))
  $ctx.Load($folder); $ctx.ExecuteQuery()
}catch{
  Write-Host "Creating library/folder path: $LibraryServerRelUrl" -ForegroundColor DarkYellow
  $folder=Ensure-FolderPath -Ctx $ctx -ServerRelativeUrl $LibraryServerRelUrl
}

# ---- 4) upload (after an extra quick unlock wait)
if (-not (Wait-FileReady -Path $localFile -TimeoutSec 60)) {
  throw "Local file is still locked: $localFile"
}
$targetUrl = ($folder.ServerRelativeUrl.TrimEnd('/') + '/' + $targetName)
$ctx.Load($folder.Files); $ctx.ExecuteQuery()
$existing = $folder.Files | Where-Object { $_.ServerRelativeUrl -ieq $targetUrl }
if ($existing -and -not $Overwrite) {
  Write-Host "File exists and -Overwrite not specified: $targetUrl" -ForegroundColor Yellow
} else {
  $bytes=[System.IO.File]::ReadAllBytes($localFile)
  $ms = New-Object System.IO.MemoryStream(,$bytes)
  $fi = New-Object Microsoft.SharePoint.Client.FileCreationInformation
  $fi.ContentStream=$ms; $fi.Url=$targetName; $fi.Overwrite=$true
  $uploaded=$folder.Files.Add($fi); $ctx.Load($uploaded); $ctx.ExecuteQuery()
  $ms.Dispose()
  Write-Host "Uploaded: $targetUrl" -ForegroundColor Green
}
try{
  $uploaded.CheckIn("Automated upload",[Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn); $ctx.ExecuteQuery()
  if($PublishAfterUpload){ $uploaded.Publish("Automated publish"); $ctx.ExecuteQuery() }
}catch{ Write-Host "Check-in/Publish skipped: $($_.Exception.Message)" -ForegroundColor DarkGray }

# ---- 5) page update (last N full months)
Add-Type -AssemblyName System.Web | Out-Null

function Get-PageHtmlOrNull {
  param([Microsoft.SharePoint.Client.ClientContext]$Ctx,[string]$ServerRelUrl)
  $file=$Ctx.Web.GetFileByServerRelativeUrl($ServerRelUrl); $item=$file.ListItemAllFields
  $Ctx.Load($item); $Ctx.ExecuteQuery()
  $html=$null; $field=$null
  if($item.FieldValues.ContainsKey("PublishingPageContent") -and $item["PublishingPageContent"]){ $html=[string]$item["PublishingPageContent"]; $field="PublishingPageContent" }
  elseif($item.FieldValues.ContainsKey("WikiField") -and $item["WikiField"]){ $html=[string]$item["WikiField"]; $field="WikiField" }
  return @($html,$field,$item,$file)
}
function Set-PageHtml { param($Ctx,$Item,[string]$FieldName,[string]$Html)
  $Item[$FieldName]=$Html; $Item.Update(); $Ctx.ExecuteQuery()
  try{ $Item.File.CheckIn("Auto update monthly links",[Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn); $Ctx.ExecuteQuery(); $Item.File.Publish("Auto publish"); $Ctx.ExecuteQuery() }catch{}
}
function MonthLabel([datetime]$d){ return ("{0} {1}" -f (Get-Culture).DateTimeFormat.AbbreviatedMonthNames[$d.Month-1].TrimEnd('.'), $d.Year) }

# Build link objects for last N full months IF files exist in the library
$now=Get-Date; $cursor=$now; $links=@()
for($i=0;$i -lt [Math]::Max(1,$MonthsToUpdate);$i++){
  $firstThis=Get-Date -Year $cursor.Year -Month $cursor.Month -Day 1
  $lastPrev=$firstThis.AddDays(-1)
  $firstPrev=Get-Date -Year $lastPrev.Year -Month $lastPrev.Month -Day 1
  $lbl=MonthLabel $firstPrev
  $nm="Pulse-all-pages-report-{0:yyyyMMdd}-{1:yyyyMMdd}.xlsx" -f $firstPrev,$lastPrev
  $href=$LibraryServerRelUrl.TrimEnd('/') + '/' + $nm
  $ok=$true; try{ $f=$ctx.Web.GetFileByServerRelativeUrl($href); $ctx.Load($f); $ctx.ExecuteQuery() }catch{ $ok=$false }
  if($ok){ $links += [pscustomobject]@{Label=$lbl;Href=$href;Year=$firstPrev.Year} }
  $cursor=(Get-Date -Year $firstPrev.Year -Month $firstPrev.Month -Day 1).AddDays(-1)
}

function Insert-Links-IntoHtml {
  param([string]$Html,[object[]]$Links)
  if([string]::IsNullOrWhiteSpace($Html)){ $Html = "" }
  foreach($lnk in $Links){
    $label=[regex]::Escape($lnk.Label); $safeHref=[System.Web.HttpUtility]::HtmlAttributeEncode($lnk.Href)
    $pat="<a\b[^>]*>(\s*$label\s*)</a>"
    if([regex]::IsMatch($Html,$pat,'IgnoreCase')){
      $Html=[regex]::Replace($Html,$pat,"<a href=""$safeHref"">$($lnk.Label)</a>",1,'IgnoreCase')
      Write-Host "Updated link: $($lnk.Label)" -ForegroundColor Green
    } else {
      $yearToken=">$($lnk.Year)<"
      $idxYear=$Html.IndexOf($yearToken,[StringComparison]::OrdinalIgnoreCase)
      if($idxYear -ge 0){
        $idxUl=$Html.IndexOf("<ul",$idxYear,[StringComparison]::OrdinalIgnoreCase)
        if($idxUl -ge 0){
          $idxClose=$Html.IndexOf(">",$idxUl)
          if($idxClose -ge 0){
            $ins="`n<li><a href=""$safeHref"">$($lnk.Label)</a></li>`n"
            $Html=$Html.Insert($idxClose+1,$ins)
            Write-Host "Inserted under $($lnk.Year): $($lnk.Label)" -ForegroundColor Green
            continue
          }
        }
      }
      $Html="<p><a href=""$safeHref"">$($lnk.Label)</a></p>`n$Html"
      Write-Host "Prepended: $($lnk.Label)" -ForegroundColor Green
    }
  }
  return $Html
}

# Fallback: update first Content Editor–style web part (property "Content")
function Update-First-ContentEditor-WebPart {
  param($Ctx,$File,[object[]]$Links)
  $scope=[Microsoft.SharePoint.Client.WebParts.PersonalizationScope]::Shared
  $wpm=$File.GetLimitedWebPartManager($scope)
  $Ctx.Load($wpm.WebParts); $Ctx.ExecuteQuery()
  foreach($def in $wpm.WebParts){ $Ctx.Load($def); $Ctx.Load($def.WebPart) }
  $Ctx.ExecuteQuery()

  if($wpm.WebParts.Count -eq 0){
    Write-Host "No web parts found on page." -ForegroundColor DarkYellow
    return $false
  }

  $preferred = @('Pulse','Analytics','Reports','Links')
  $cands=@()
  foreach($def in $wpm.WebParts){
    $wp=$def.WebPart
    $t=$wp.Title
    Write-Host ("WebPart: Title='{0}' Type='{1}'" -f $t,$wp.GetType().FullName) -ForegroundColor Gray
    $props=$wp.Properties
    $hasContent=$false
    try{ $hasContent = $props -and $props.FieldValues.ContainsKey("Content") }catch{ $hasContent=$false }
    if($hasContent){
      $score=100; foreach($tok in $preferred){ if($t -and ($t -match $tok)){ $score-=50 } }
      $cands += [pscustomobject]@{Def=$def;WP=$wp;Score=$score}
    }
  }
  if($cands.Count -eq 0){
    Write-Host "No CEWP-like web part with editable 'Content' found." -ForegroundColor DarkYellow
    return $false
  }

  $chosen = $cands | Sort-Object Score | Select-Object -First 1
  $wp = $chosen.WP; $def = $chosen.Def
  $old=""; try{ $old=[string]$wp.Properties["Content"] }catch{}
  $new = Insert-Links-IntoHtml -Html $old -Links $Links
  if($new -eq $old){ Write-Host ("No changes needed in '" + $wp.Title + "'.") -ForegroundColor DarkGray; return $true }
  $wp.Properties["Content"]=$new
  $def.SaveWebPartChanges(); $Ctx.ExecuteQuery()
  Write-Host ("Updated web part '" + $wp.Title + "' content.") -ForegroundColor Green
  return $true
}

if($links.Count -gt 0){
  $ph = Get-PageHtmlOrNull -Ctx $ctx -ServerRelUrl $PageServerRelUrl
  $html=$ph[0]; $field=$ph[1]; $item=$ph[2]; $file=$ph[3]

  $updated = $false
  if($field){
    $newHtml = Insert-Links-IntoHtml -Html $html -Links $links
    if($newHtml -ne $html){ Set-PageHtml -Ctx $ctx -Item $item -FieldName $field -Html $newHtml }
    $updated = $true
    Write-Host "Page updated (field: $field)." -ForegroundColor Green
  } else {
    Write-Host "Page has no Publishing/Wiki field; attempting Web Part update..." -ForegroundColor Yellow
    $updated = Update-First-ContentEditor-WebPart -Ctx $ctx -File $file -Links $links
  }

  if(-not $updated){
    Write-Host "Could not update page content automatically. See log for web part list." -ForegroundColor DarkYellow
  }
} else {
  Write-Host "No page updates (no matching monthly files in library)." -ForegroundColor DarkGray
}

# ---- end / logging tidy ----
try { Stop-Transcript | Out-Null } catch {}
Write-Host ("Done. Log: " + (Join-Path $optionalDir "upload_To_Pulse.log")) -ForegroundColor Cyan


# Stop-Transcript | Out-Null

# ---- logging
$logPath = Join-Path $optionalDir "upload_To_Pulse.log"

$transcriptStarted = $false
try {
  Start-Transcript -Path $logPath -Append -Force | Out-Null
  $transcriptStarted = $true
} catch {
  Write-Host "Transcript not started (host limitation): $($_.Exception.Message)" -ForegroundColor DarkGray
}

try {
  # >>> all your existing script body goes here <<<
}
finally {
  if ($transcriptStarted) {
    try { Stop-Transcript | Out-Null } catch {}
  }
}

Write-Host "Done. Log: $logPath" -ForegroundColor Cyan
