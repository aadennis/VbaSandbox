$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$basPath = Join-Path $scriptDir "generator.bas"

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $true

$txtFiles = Get-ChildItem "sample_lyrics.txt" -Filter *.txt

foreach ($file in $txtFiles) {
    $pres = $ppt.Presentations.Add()
    $pres.VBProject.VBComponents.Import($basPath)

    $ppt.Run("CreatePresentationFromText", $file.FullName)

    $outputPath = Join-Path $scriptDir + ($file.BaseName + ".pptm")
    $pres.SaveAs($outputPath, 25) # ppSaveAsOpenXMLMacroEnabled
    $pres.Close()
}

$ppt.Quit()