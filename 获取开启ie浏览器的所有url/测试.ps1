
   $IEObjs = @()
   $ShellWindows = (New-Object -ComObject Shell.Application).Windows()
   $360chromes = (Get-Process -Name 360chrome).MainWindowTitle;

   #foreach($chrome in $360chromes){
    #$element = [System.Windows.Automation.AutomationProperties]::NameProperty;
    #Write-Host $element
   
   #}
 
   Write-Host $ShellWindows;
   Write-Host $360chromes -BackgroundColor Cyan;



    Foreach($IE in $ShellWindows)
    {
        $FullName = $IE.FullName;
        Write-Host $IE;
        Write-Host " Fullname: $FullName" -BackgroundColor Red;

        If($FullName -ne $NULL)
        {
            $FileName = Split-Path -Path $FullName -Leaf
            Write-Host $FileName
            If($FileName.ToLower() -eq "360chrome.exe")
            {
                $Title = $IE.LocationName
                $URL = $IE.LocationURL
                $IEObj = New-Object -TypeName PSObject -Property @{Title = $Title; URL = $URL}
                $IEObjs += $IEObj
            }
        }
    }