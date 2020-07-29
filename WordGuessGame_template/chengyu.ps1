
Param(
[Parameter(Mandatory=$false)]
    [string]$FilePath = '.\data\chengyu-sample.csv',
[Parameter(Mandatory=$false)]    
    [int]$Size = 10,
[Parameter(Mandatory=$false)]    
    [string]$OutputFile = 'output.pptx'
)

Add-type -AssemblyName office
$Application = New-Object -ComObject powerpoint.application
$application.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$slideType = "microsoft.office.interop.powerpoint.ppSlideLayout" -as [type]
$templatePresentation = Join-Path $PSScriptRoot 'template.pptx'

$newline = [System.Environment]::NewLine
$msoCTrue = 1 
$msoTextOrientationHorizontal = 1  
$horiz = $msoTextOrientationHorizontal 

$presentation = $application.Presentations.open($templatePresentation)
Import-Csv $FilePath -Encoding "UTF8" | Get-Random -Count $Size | ForEach-Object { `
    $customLayout = $presentation.Slides.item(1).customLayout
    $slide = $presentation.slides.addSlide(1, $customLayout)
    $slide.layout = $slideType::ppLayoutTitle
    $slide.Shapes.title.TextFrame.TextRange.Text = $_.idiom

    $left = 80
    $top = 200
    $width = 1000
    $height = 550
    $tb = $slide.Shapes.AddTextbox($horiz,$left,$top,$width,$height) 
    $tb.TextFrame.TextRange.Text = $_.pinyin + $newline +  $_.explanation
}

$outPath = Join-Path $PSScriptRoot $OutputFile
$presentation.SavecopyAs($outPath)
$presentation.Close()
$application.quit()
$application = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()