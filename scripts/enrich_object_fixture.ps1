param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath
)

$ErrorActionPreference = "Stop"
$resolvedDocument = (Resolve-Path -LiteralPath $DocumentPath).Path

Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 300

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

function Find-EndRange($document, [string]$text) {
    $range = $document.Content.Duplicate
    $range.Find.ClearFormatting()
    $range.Find.Text = $text
    $range.Find.Forward = $true
    $range.Find.Wrap = 0
    if (-not $range.Find.Execute()) {
        throw "Object fixture marker not found: $text"
    }
    $range.Collapse(0)
    return $range
}

try {
    $document = $word.Documents.Open($resolvedDocument, $false, $false)
    try {
        $hasTextBox = $false
        for ($index = 1; $index -le $document.Shapes.Count; $index++) {
            if ($document.Shapes.Item($index).Type -eq 17) {
                $hasTextBox = $true
                break
            }
        }
        if (-not $hasTextBox) {
            $range = Find-EndRange $document "Text boxes |"
            $shape = $document.Shapes.AddTextbox(1, 370, 160, 100, 14, $range)
            $shape.TextFrame.TextRange.Text = "REMOVABLE TEXT BOX"
            $shape.TextFrame.TextRange.Font.Size = 7
            $shape.TextFrame.MarginTop = 0
            $shape.TextFrame.MarginBottom = 0
            $shape.TextFrame.MarginLeft = 2
            $shape.TextFrame.MarginRight = 2
            $shape.RelativeHorizontalPosition = 0
            $shape.RelativeVerticalPosition = 2
            $shape.Left = 370
            $shape.Top = -1
            $shape.LockAnchor = $true
        }

        $hasHorizontalLine = $false
        for ($index = 1; $index -le $document.InlineShapes.Count; $index++) {
            if ($document.InlineShapes.Item($index).Type -eq 6) {
                $hasHorizontalLine = $true
                break
            }
        }
        if (-not $hasHorizontalLine) {
            $range = Find-EndRange $document "Horizontal lines |"
            $document.InlineShapes.AddHorizontalLineStandard($range) | Out-Null
        }

        $hasFixtureControl = $false
        for ($index = 1; $index -le $document.ContentControls.Count; $index++) {
            if ($document.ContentControls.Item($index).Tag -eq "CleanupSuiteFixtureFormControl") {
                $hasFixtureControl = $true
                break
            }
        }
        if (-not $hasFixtureControl) {
            $range = Find-EndRange $document "Form controls |"
            $control = $document.ContentControls.Add(8, $range)
            $control.Title = "Removable fixture form control"
            $control.Tag = "CleanupSuiteFixtureFormControl"
            $control.Checked = $true
        }

        $document.Save()
    }
    finally {
        $document.Close(0)
    }
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
}
