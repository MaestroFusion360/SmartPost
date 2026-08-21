param(
  [Parameter(Mandatory = $true)]
  [string]$PostExe,
  [Parameter(Mandatory = $true)]
  [string]$InputFile,
  [Parameter(Mandatory = $true)]
  [string]$DownstreamPost,
  [Parameter(Mandatory = $true)]
  [string]$OutputDirectory,
  [string[]]$XmlPost
)

$ErrorActionPreference = "Continue"
$repositoryRoot = Split-Path -Parent $PSScriptRoot

if (-not $XmlPost) {
  $XmlPost = @(
    (Join-Path $repositoryRoot "commands\smart_post_dialog\xml_last.cps"),
    (Join-Path $repositoryRoot "commands\smart_post_dialog\xml.cps")
  )
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null

$results = foreach ($post in $XmlPost) {
  $name = [System.IO.Path]::GetFileNameWithoutExtension($post)
  $postOutput = Join-Path $OutputDirectory $name
  New-Item -ItemType Directory -Force -Path $postOutput | Out-Null

  $stem = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
  $xml = Join-Path $postOutput ($stem + ".xml")
  $nc = Join-Path $postOutput ($stem + ".nc")

  $xmlLog = & $PostExe $post $InputFile $xml --noeditor --property useTimeStamp false --property highAccuracy true 2>&1 | Out-String
  $xmlExit = $LASTEXITCODE
  [System.IO.File]::WriteAllText(($xml + ".log"), $xmlLog)

  $importLog = ""
  $importExit = $null
  if ($xmlExit -eq 0) {
    $importLog = & $PostExe $DownstreamPost $xml $nc --format xml --noeditor --debugall --property programName 1001 2>&1 | Out-String
    $importExit = $LASTEXITCODE
    [System.IO.File]::WriteAllText(($nc + ".log"), $importLog)
  }

  $ncText = if (Test-Path -LiteralPath $nc) {
    [System.IO.File]::ReadAllText($nc)
  } else {
    ""
  }
  $xmlText = if (Test-Path -LiteralPath $xml) {
    [System.IO.File]::ReadAllText($xml)
  } else {
    ""
  }

  [pscustomobject]@{
    XmlPost = [System.IO.Path]::GetFileName($post)
    Groups = ([regex]::Matches($xmlText, "<group id='([^']+)'") | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique) -join ","
    XmlExit = $xmlExit
    ImportExit = $importExit
    InternalError = $importLog.Contains("Internal error")
    OnCycle = $ncText.Contains("!DEBUG: onCycle()")
    OnCyclePoint = $ncText.Contains("!DEBUG: onCyclePoint(")
    OnCycleEnd = $ncText.Contains("!DEBUG: onCycleEnd()")
    CannedCodes = ([regex]::Matches($ncText, "(?m)^N\d+.*\bG(8[1-9])\b") | ForEach-Object { "G" + $_.Groups[1].Value } | Select-Object -Unique) -join ","
  }
}

$results | Export-Csv -LiteralPath (Join-Path $OutputDirectory "results.csv") -NoTypeInformation -Encoding UTF8
$results | Format-Table -AutoSize

if ($results.Where({ $_.XmlExit -ne 0 -or $_.ImportExit -ne 0 }).Count -gt 0) {
  exit 1
}
