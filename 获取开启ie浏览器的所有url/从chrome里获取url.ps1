Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
#UIAutomationClientsideProviders をロードするためのヘルパー
$source = @"
using System;
using System.Windows.Automation;
using System.Diagnostics;
namespace UIAutomationHelper {
  public class Element {
    public static AutomationElement RootElement {
      get { return AutomationElement.RootElement; }
    }
    public static AutomationElementCollection FromProcess(Process process) {
      PropertyCondition condProc = new PropertyCondition(AutomationElement.ProcessIdProperty, process.Id);
      PropertyCondition condClass = new PropertyCondition(AutomationElement.ClassNameProperty, "Chrome_WidgetWin_1");
      return RootElement.FindAll(TreeScope.Element | TreeScope.Children, new AndCondition(condProc, condClass));
    }
  }
}
"@
Add-Type -TypeDefinition $source -ReferencedAssemblies("UIAutomationClient", "UIAutomationTypes");

$chromes = Get-Process -name 360chrome
foreach ($chrome in $chromes) {
  foreach ($elChromeMain in [UIAutomationHelper.Element]::FromProcess($chrome)) {
    Write-Host $elChromeMain -BackgroundColor Cyan;
    $cond = New-Object -TypeName System.Windows.Automation.PropertyCondition(
      [System.Windows.Automation.AutomationElement]::NameProperty, "");
    Write-Host $cond -BackgroundColor Cyan;   
    $elChromeSub = $elChromeMain.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond);
    Write-Host "el:" $elChromeSub -BackgroundColor Cyan;

    $cond = New-Object -TypeName System.Windows.Automation.PropertyCondition(
      [System.Windows.Automation.AutomationElement]::NameProperty, "Address and search bar");
    $elChromeUrl = $elChromeSub.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond);

    $url = $elChromeUrl.GetCurrentPropertyValue([System.Windows.Automation.ValuePatternIdentifiers]::ValueProperty);

    'Proc: 0x{0:X8}({0})' -f $chrome.Id;
    'HWND: 0x{0:X8}({0})' -f $chrome.MainWindowHandle.ToInt32();
    '{0}' -f $elChromeMain.Current.Name;
    $url;

    #ついでにタブも列挙してみる
    $cond = New-Object -TypeName System.Windows.Automation.PropertyCondition(
      [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
      [System.Windows.Automation.ControlType]::Tab);
    $elChromeTab = $elChromeSub.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond);
    $cond = New-Object -TypeName System.Windows.Automation.PropertyCondition(
      [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
      [System.Windows.Automation.ControlType]::TabItem);
    $elChromeTabItems = $elChromeTab.FindAll([System.Windows.Automation.TreeScope]::Children, $cond);
    $idx = 0;
    foreach ($elChromeTabItem in $elChromeTabItems) {
      $idx += 1;
      '  #{0,-2}: {1}' -f $idx, $elChromeTabItem.Current.Name;
    }
    '';
  }
}