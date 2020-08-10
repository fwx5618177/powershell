$chromes = Get-Process -name 360chrome;

foreach($chrome in $360chromes){
    #$element = [System.Windows.Automation.AutomationProperties]::NameProperty;
    $element = [System.Windows.Automation.AutomationElement]::FromHandle($chrome.MainWindowHandle);
    #$condition = [System.Windows.Automation.PropertyCondition]::New([System.Windows.Automation.AutomationElement]::ClassNameProperty, "Chrome_OmniboxView")
    $condition = [System.Windows.Automation.AndCondition]::new([System.Windows.Automation.AutomationElement]::ProcessIdProperty);
    Write-Host $element
   
}