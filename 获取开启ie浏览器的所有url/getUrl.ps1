

[CmdletBinding(DefaultParameterSetName='Export')]
Param
(
    [Parameter(Mandatory=$true,Position=0,ParameterSetName='Open')]
    [Alias('Open')][String]$OpenURL,
    [Parameter(Mandatory=$true,Position=0,ParameterSetName='Export')]
    [Alias('path')][String]$Export
)


#这里我们设计了一个主方法
#GetIEURL用来提取当前IE游览器中所有的Tab栏目地址
#这里有一点需要提到的是，我们采用了，Com接口里的Shell.Application的对象，然后调用旗下的Windows方法来寻找当前打开的iexplore.exe进程，从而获取每一个IE进程的locationURL与LocationName

Function GetIEURL
{
    $IEObjs = @()
    $ShellWindows = (New-Object -ComObject Shell.Application).Windows()

    Foreach($IE in $ShellWindows)
    {
        $FullName = $IE.FullName
        If($FullName -ne $NULL)
        {
            $FileName = Split-Path -Path $FullName -Leaf

            If($FileName.ToLower() -eq "360chrome.exe")
            {
                $Title = $IE.LocationName;
                $URL = $IE.LocationURL
                $IEObj = New-Object -TypeName PSObject -Property @{Title = $Title; URL = $URL}
                $IEObjs += $IEObj
            }
        }
    }

    $IEObjs
}

#导出提取的URL

If($Export)
{
    $CurrentIEURL = GetIEURL
    If($CurrentIEURL -ne $null)
    {
        $CurrentIEURL|Export-Csv -Path "$Export" -NoTypeInformation
        Write-Host "成功导出Csv文件到 '$Export'"
    }
    Else
    {
        Write-Warning "当前没有打开的Tab栏。"
    }
}

#导出的URL文件里自动打开所有的URL

If($OpenURL)
{
    $URLs = (Import-Csv -Path $OpenURL).URL
            
    $IEApplication = New-Object -ComObject InternetExplorer.Application
    $navOpenInBackgroundTab = 0x1000
    $IEApplication.Visible = $true 

    ForEach($URL in $URLs)
    {
        Try
        {
            $IEApplication.Navigate($URL, $navOpenInBackgroundTab)
            While($IEApplication.Busy)
            {
                Start-Sleep -Millisecond 100
            }
            Write-Host "成功打开 '$URL' 到IE游览器中。"
        }
        Catch
        {
            Write-Host "打开'$URL'失败。"
        }
    }
}