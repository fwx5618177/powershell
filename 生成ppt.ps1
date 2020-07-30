Add-type -AssemblyName office
$ppt = New-Object -ComObject Powerpoint.Application;

# 设置可视化
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoCTrue;
#弹窗
$ppt.displayalerts = [Microsoft.Office.Core.MsoTriState]::msoCTrue;

#打开ppt
$PPTDocument = $ppt.Presentations.open("D:\code data\Powershell\pptgenerator\local\test.pptx");

#获取选定页（第一页）
$PPTSlideLayout = $PPTDocument.Slides(1).Layout

#指定版式
$PPTSlideLayout = [microsoft.office.interop.powerpoint.ppSlideLayout]::ppLayoutTitle;

#添加slip
$slip = $ppt.Presentations.Add();

#幻灯片动画效果
$SlideAnimationEffects = 257, 258, 513, 769, 770, 1025, 1026, 1281, 1282, 1283, 1284, 1285, 1286, 1287, 1288, 1537, 1793, 2049, 2050, 2051, 2052, 2053, 2054, 2055, 2056, 2305, 2306, 2561, 2562, 2563, 2564, 2565, 2566, 2567, 2568, 2817, 2818, 2819, 2820, 3073, 3074, 3585, 3586, 3587, 3588, 3844, 3845, 3846, 3847, 3848, 3849, 3850, 3851, 3852, 3853, 3854, 3855, 3856, 3857, 3858, 3859, 3860, 3861
#文字动画效果
$TextAnimationEffects = 257, 258, 513, 769, 770, 1025, 1026, 1281, 1282, 1283, 1284, 1285, 1286, 1287, 1288, 1537, 1793, 2049, 2050, 2051, 2052, 2053, 2054, 2055, 2056, 2305, 2306, 2561, 2562, 2563, 2564, 2565, 2566, 2567, 2568, 2817, 2818, 2819, 2820, 3073, 3074, 3585, 3586, 3587, 3588, 3844

#从效果中选择随机动画效果
$RandomAnimation = ($SlideAnimationEffects | Get-Random);
#设置效果
$slip.SlideShowTransition.EntryEffect = $RandomAnimation;

#标题
        $slide.Shapes.Title.TextFrame.TextRange.Text = $TitleText
        $slide.Shapes.Title.AnimationSettings.TextLevelEffect = 1
        $RandomAnimation = ($TextAnimationEffects | Get-Random)
        $slide.Shapes.Title.AnimationSettings.EntryEffect = $RandomAnimation
        $slide.Shapes.Title.AnimationSettings.AdvanceMode = 2
        $slide.Shapes.Title.AnimationSettings.Animate = $msoTrue
                $slide.Shapes.Item(2).TextFrame2.Column.Number = 3 # This is where you might wish to change the number of the item if you have the text field on a different index
        $slide.Shapes.Item(2).TextFrame.TextRange.Text = $SlideText # This is where you might wish to change the number of the item if you have the text field on a different index
        $slide.Shapes.Item(2).AnimationSettings.TextLevelEffect = 1
        $RandomAnimation = ($TextAnimationEffects | Get-Random)
        $slide.Shapes.Item(2).AnimationSettings.EntryEffect = $RandomAnimation
        $slide.Shapes.Item(2).AnimationSettings.AdvanceMode = 2
        $slide.Shapes.Item(2).AnimationSettings.Animate = $msoTrue

#添加图片
$pic = $slide.Shapes.AddPicture(($RandomPicture.FilePath),$msoFalse, $msoTrue,$PicturesLeft,$PicturesTop, $RandomPicture.Width, $RandomPicture.Height)




#另存为
$slip.SaveAs("D:\code data\Powershell\pptgenerator\local\模板生成.pptx");

#退出
$slip.Application.quit();

#完整关闭
 [gc]::collect()
    [gc]::WaitForPendingFinalizers()
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
    $PPTDocument.Saved = $msoTrue
    $PPTDocument.Close()
    $null = [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($PPTDocument)
    $PPTDocument = $null
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
    $PPTApplication.quit()
    $null = [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($PPTApplication)
    $PPTApplication = $null
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()

$ppt = $null;
[GC]::Collect();