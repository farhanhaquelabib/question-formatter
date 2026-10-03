param (
    [string]$docxPath,
    [string]$pdfPath
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($docxPath)
    $wdFormatPDF = 17
    $doc.SaveAs([ref]$pdfPath, [ref]$wdFormatPDF)
    $doc.Close()
    Write-Host "Success: $pdfPath"
} catch {
    Write-Host "Error: $_"
} finally {
    $word.Quit()
}
