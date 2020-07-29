
Param(
[Parameter(Mandatory=$false)]
    [string]$FilePath = '.\data\THUOCL_chengyu.txt',
[Parameter(Mandatory=$false)]    
    [int]$Size = 50,
[Parameter(Mandatory=$false)]    
    [string]$OutputFile = 'output.pptx'
)

Add-type -AssemblyName office
$Application = New-Object -ComObject powerpoint.application
$application.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$slideType = "microsoft.office.interop.powerpoint.ppSlideLayout" -as [type]
$templatePresentation = Join-Path $PSScriptRoot 'template.pptx'

$presentation = $application.Presentations.open($templatePresentation)
[string[]]$arrayFromFile = Get-Content $FilePath -Encoding "UTF8" 
$output = $arrayFromFile | Get-Random -Count $Size 
$output | ForEach-Object { `
    $customLayout = $presentation.Slides.item(1).customLayout
    $slide = $presentation.slides.addSlide(1, $customLayout)
    $slide.layout = $slideType::ppLayoutTitle
    $slide.Shapes.title.TextFrame.TextRange.Text = $_.split()[0]
    $slide.Shapes.title.TextFrame.TextRange.Font.Bold = $true
}

$outPath = Join-Path $PSScriptRoot $OutputFile
$presentation.SavecopyAs($outPath)
$presentation.Close()
$application.quit()
$application = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()