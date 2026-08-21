# XML postprocessor tests

These PowerShell scripts test the complete SmartPost intermediate XML path:

```text
Autodesk .cnc dataset
-> xml_last.cps or xml.cps
-> XML
-> post.exe --format xml
-> unchanged downstream .cps
-> NC output
```

The scripts never edit the downstream postprocessor. Generated XML, NC, log,
and CSV files are written below the selected output directory.

## Test one dataset

```powershell
.\scripts\test-xml-roundtrip.ps1 `
  -PostExe "C:\path\to\post.exe" `
  -InputFile "C:\path\to\deep drilling.cnc" `
  -DownstreamPost "C:\path\to\fanuc.cps" `
  -OutputDirectory ".\tmp\roundtrip"
```

By default, the script tests both
`commands/smart_post_dialog/xml_last.cps` and
`commands/smart_post_dialog/xml.cps`. Use `-XmlPost` to test a specific file.

## Test every dataset in a directory

```powershell
.\scripts\test-xml-cycle-matrix.ps1 `
  -PostExe "C:\path\to\post.exe" `
  -InputDirectory "C:\path\to\CNC files\Milling\Drilling" `
  -DownstreamPost "C:\path\to\fanuc.cps" `
  -OutputDirectory ".\tmp\cycle-matrix"
```

The matrix writes `results.csv` and reports XML export status, XML import
status, restored cycle callbacks, internal errors, and detected G81-G89 codes.
