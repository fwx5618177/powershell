$AnyWindow=$Chrome.WindowHandles.Item(2) 
$Chrome=$Chrome.SwitchTo().Window($AnyWindow)
Write-Host $Chrome.url : $Chrome.Title
$Chrome.Close()