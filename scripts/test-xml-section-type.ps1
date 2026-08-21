param(
  [Parameter(Mandatory = $true)]
  [string]$PostExe,
  [Parameter(Mandatory = $true)]
  [string]$InputFile,
  [Parameter(Mandatory = $true)]
  [string]$DumpPost,
  [Parameter(Mandatory = $true)]
  [string]$TurningPost,
  [Parameter(Mandatory = $true)]
  [string]$OutputDirectory,
  [string]$XmlPost
)

$ErrorActionPreference = "Continue"
$repositoryRoot = Split-Path -Parent $PSScriptRoot

if (-not $XmlPost) {
  $XmlPost = Join-Path $repositoryRoot "commands\smart_post_dialog\xml_last.cps"
}

foreach ($path in @($PostExe, $InputFile, $DumpPost, $TurningPost, $XmlPost)) {
  if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
    throw "Required file not found: $path"
  }
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null

$stem = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
$directDump = Join-Path $OutputDirectory ($stem + "-direct.dmp")
$xml = Join-Path $OutputDirectory ($stem + ".xml")
$xmlDump = Join-Path $OutputDirectory ($stem + "-xml.dmp")
$turningOutput = Join-Path $OutputDirectory ($stem + "-turning.nc")

$directLog = & $PostExe $DumpPost $InputFile $directDump --noeditor --debugall 2>&1 | Out-String
$directExit = $LASTEXITCODE
[System.IO.File]::WriteAllText(($directDump + ".log"), $directLog)

$exportLog = & $PostExe $XmlPost $InputFile $xml --noeditor `
  --property useTimeStamp false `
  --property highAccuracy true `
  --property diagnosticSectionType true 2>&1 | Out-String
$exportExit = $LASTEXITCODE
[System.IO.File]::WriteAllText(($xml + ".log"), $exportLog)

$xmlDumpLog = ""
$xmlDumpExit = $null
$turningLog = ""
$turningExit = $null
if ($exportExit -eq 0) {
  $xmlDumpLog = & $PostExe $DumpPost $xml $xmlDump --format xml --noeditor --debugall 2>&1 | Out-String
  $xmlDumpExit = $LASTEXITCODE
  [System.IO.File]::WriteAllText(($xmlDump + ".log"), $xmlDumpLog)

  $turningLog = & $PostExe $TurningPost $xml $turningOutput --format xml --noeditor 2>&1 | Out-String
  $turningExit = $LASTEXITCODE
  [System.IO.File]::WriteAllText(($turningOutput + ".log"), $turningLog)
}

$directText = if (Test-Path -LiteralPath $directDump) {
  [System.IO.File]::ReadAllText($directDump)
} else {
  ""
}
$xmlDumpText = if (Test-Path -LiteralPath $xmlDump) {
  [System.IO.File]::ReadAllText($xmlDump)
} else {
  ""
}

$directTurning = $directText.Contains("currentSection.type=1 (TYPE_TURNING)")
$xmlMilling = $xmlDumpText.Contains("currentSection.type=0 (TYPE_MILLING)")
$turningRejectedAsMilling = $turningLog.Contains("Milling toolpath is not supported by the post configuration.")

$result = [pscustomobject]@{
  DirectDumpExit = $directExit
  XmlExportExit = $exportExit
  XmlDumpExit = $xmlDumpExit
  TurningPostExit = $turningExit
  DirectTypeTurning = $directTurning
  XmlTypeMilling = $xmlMilling
  TurningPostRejectedAsMilling = $turningRejectedAsMilling
}

$result | Format-List
$result | Export-Csv -LiteralPath (Join-Path $OutputDirectory "section-type-result.csv") -NoTypeInformation -Encoding UTF8

$reproduced = (
  $directExit -eq 0 -and
  $exportExit -eq 0 -and
  $xmlDumpExit -eq 0 -and
  $turningExit -ne 0 -and
  $directTurning -and
  $xmlMilling -and
  $turningRejectedAsMilling
)

if (-not $reproduced) {
  Write-Error "The expected Autodesk XML Section.type limitation was not reproduced."
  exit 1
}

Write-Host "Reproduced: TYPE_TURNING in CNC becomes TYPE_MILLING after XML import."
exit 0
