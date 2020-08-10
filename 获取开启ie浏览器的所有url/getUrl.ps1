

[CmdletBinding(DefaultParameterSetName='Export')]
Param
(
    [Parameter(Mandatory=$true,Position=0,ParameterSetName='Open')]
    [Alias('Open')][String]$OpenURL,
    [Parameter(Mandatory=$true,Position=0,ParameterSetName='Export')]
    [Alias('path')][String]$Export
)


#�������������һ��������
#GetIEURL������ȡ��ǰIE�����������е�Tab��Ŀ��ַ
#������һ����Ҫ�ᵽ���ǣ����ǲ����ˣ�Com�ӿ����Shell.Application�Ķ���Ȼ��������µ�Windows������Ѱ�ҵ�ǰ�򿪵�iexplore.exe���̣��Ӷ���ȡÿһ��IE���̵�locationURL��LocationName

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

#������ȡ��URL

If($Export)
{
    $CurrentIEURL = GetIEURL
    If($CurrentIEURL -ne $null)
    {
        $CurrentIEURL|Export-Csv -Path "$Export" -NoTypeInformation
        Write-Host "�ɹ�����Csv�ļ��� '$Export'"
    }
    Else
    {
        Write-Warning "��ǰû�д򿪵�Tab����"
    }
}

#������URL�ļ����Զ������е�URL

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
            Write-Host "�ɹ��� '$URL' ��IE�������С�"
        }
        Catch
        {
            Write-Host "��'$URL'ʧ�ܡ�"
        }
    }
}