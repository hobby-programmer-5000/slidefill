# Checks SmartArt filling against the real PowerPoint, which is the only judge of whether it shows the values.
#   -Make <template.pptx>  : PowerPoint draws a SmartArt list whose nodes hold the sample's {{sales}} and {{rate}} names
#   -Read <report.pptx> [-OutDir <png folder>] : prints each SmartArt node's text as PowerPoint reads it, and exports PNGs
# Usage: make, fill with samples\sample_data.xlsx (core.run), then read. Exit 0 = PowerPoint opened and read it.
param(
    [string]$Make = "",
    [string]$Read = "",
    [string]$OutDir = ""
)
$ErrorActionPreference = "Stop"
$ppt = New-Object -ComObject PowerPoint.Application
$hadOpen = $ppt.Presentations.Count -gt 0
$exit = 0
try {
    if ($Make) {
        $deck = $ppt.Presentations.Add(-1)  # with a window: without one PowerPoint lists no SmartArt layouts
        $slide = $deck.Slides.Add(1, 12)  # ppLayoutBlank
        $layout = $ppt.SmartArtLayouts.Item(1)  # Basic Block List
        $shape = $slide.Shapes.AddSmartArt($layout, 50, 50, 600, 300)
        $nodes = $shape.SmartArt.AllNodes
        while ($nodes.Count -gt 2) { $nodes.Item($nodes.Count).Delete() }
        $sales = [string][char]0xB9E4 + [char]0xCD9C  # the sample names, spelled in code points: Windows PowerShell reads a script without BOM as ANSI
        $rate = [string][char]0xB2EC + [char]0xC131 + [char]0xB960
        $nodes.Item(1).TextFrame2.TextRange.Text = "$sales {{$sales}}"
        $nodes.Item(2).TextFrame2.TextRange.Text = "{{$rate}}"
        $deck.SaveAs([System.IO.Path]::GetFullPath($Make))
        $deck.Close()
        "Made: $Make"
    }
    if ($Read) {
        $deck = $ppt.Presentations.Open((Resolve-Path $Read).Path, -1, 0, 0)
        foreach ($slide in $deck.Slides) {
            foreach ($shape in $slide.Shapes) {
                if ($shape.HasSmartArt) {
                    foreach ($node in $shape.SmartArt.AllNodes) { "Slide {0} SmartArt node: {1}" -f $slide.SlideIndex, $node.TextFrame2.TextRange.Text }
                }
            }
        }
        if ($OutDir) {
            New-Item -ItemType Directory -Force $OutDir | Out-Null
            $deck.Export((Resolve-Path $OutDir).Path, "PNG", 1280, 720)
            "Exported PNGs to $OutDir"
        }
        $deck.Close()
    }
    "OK"
}
catch {
    "FAILED: $($_.Exception.Message)"
    $exit = 1
}
finally {
    if (-not $hadOpen -and $ppt.Presentations.Count -eq 0) { $ppt.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
exit $exit
