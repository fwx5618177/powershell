#测试路径是否存在，如果不存在，则创建一个
function Resolve-Directory {
    param (
        [Parameter(Mandatory)]
        [string]
        $Path
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -ErrorAction SilentlyContinue
    }
}

#url
#https://xxgk.eic.sh.cn/jsp/view/list.jsp
#https://xxgk.eic.sh.cn/jsp/view/info.jsp?id=1632
#/jsp/view/info.jsp?id=
#https://xxgk.eic.sh.cn/jsp/view/info.jsp?id=1632

function GetALLInformation() {
        param (
        [Parameter(ValueFromPipeline=$true)]
        [int]
        $Page,
        [Parameter(Mandatory)]
        [string]
        $Path,
        [Parameter(Mandatory)]
        [string]
        $DateTime
        )

        begin {
            $headers = @{
            'Accept'                    = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3'
            'Accept-Encoding'           = 'gzip, deflate, br'
            'Accept-Language'           = 'zh-CN,zh;q=0.9'
            'Host'                      = 'xxgk.eic.sh.cn'
            'Upgrade-Insecure-Requests' = '1'
            'User-Agent'                = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36'
            }

            Resolve-Directory -Path $Path

        }

        process {
           $body = @{
             'pageSize'                 = ''
             'currentPage'              = $Page
             'selField'                 = 'ST_EXT1'
             'selValue'                 = ''
             'selFieldShowStr'          = '污染源名称'
             'district'                 = ''
             'nd'                       = $DateTime
            }


            #Set Url
            $url = "https://xxgk.eic.sh.cn/jsp/view/list.jsp"
            #Get Web Page
            $web = Invoke-WebRequest -Uri $url -Method Post -Body $body -Headers $header
            #Regex, and Get id data and page in first page
            $table = ([regex]"(?s)<table[^<].+/table>").Matches($web.content);

            ([regex]"(?s)<th.+/th>").Matches($table) | ForEach-Object {
                #title, column name
                $ColumnList = (([regex]"th[^<]+?</th").Matches($_.value)) -replace "(th|>|<|/)", ""
                $TitleSequence = $ColumnList[0]
                $TitleName = $ColumnList[1]
                $TitleDistrict = $ColumnList[2]
                $TitleType = $ColumnList[3]
                $TitleManage = $ColumnList[4]

            }


            ([regex]"(?s)<tr.+?/tr>").Matches($table) | ForEach-Object {
                #id
                $_.value | Select-String -Pattern "\(.+?\)" -AllMatches | ForEach-Object {$_.matches} | ForEach-Object {
                    $id = $_.value -replace "(\('|\'\))", ""
                    
                }
                
                #Sequence, Company, Distriction, Type
                #$_.value
                $list = (([regex]"<td.+?/td>").Matches($_.value))  -replace "(<|/|>|td|&nbsp;)", ""
                if([string]::IsNullOrEmpty($list)){
                    return
                }
                $Sequence = $list[0];
                $Company = $list[1].Trim();
                #Write-Host $Company -BackgroundColor red
                $Distriction = $list[2];
                $Type = $list[3];

                #Write-Host "$Sequence - $Company - $Distriction - $Type "  -BackgroundColor DarkYellow
                #Write-Output "$Sequence - $Company - $Distriction - $Type" | Out-File -Append List.txt


                #进入子网页爬取数据
                #https://xxgk.eic.sh.cn/jsp/view/list.jsp
                #https://xxgk.eic.sh.cn/jsp/view/info.jsp?id=1632
                #/jsp/view/info.jsp?id=
                #https://xxgk.eic.sh.cn/jsp/view/info.jsp?id=1632
                if([string]::IsNullOrEmpty($id)){
                    return
                }

                $ChildUrl = "https://xxgk.eic.sh.cn/jsp/view/info.jsp?id=$id"
                #Write-Host $ChildUrl
                
                
                #获取网页
                $ChildWeb = Invoke-WebRequest -Uri $ChildUrl -Method GET -Headers $header
                
                #Validate $validate_title
                $validate = ([regex]"(?s)<div.+?limit.+?/div>").Matches($ChildWeb.content);
                
                $validate_title = (([regex]"p.+?<a").Matches($validate)) -replace "(p>|<a)", ""
                #验证正确后继续执行
                #Write-Host $validate_title -BackgroundColor Red
                #Write-Host $Company.trim() -BackgroundColor Cyan

 
                    #联系人，电话
                    $ChildContent = ([regex]"(?s)联系人.+?/tr>").Matches($ChildWeb.content);
                    
                    $INF = (([regex]"<th.+?/th>").Matches($ChildContent)) -replace "(<th>|</th>)",""

                    $ContactPersonTitle = $INF[0];
                    $PhoneNumTitle = $INF[1];

                    $Text = (([regex]"<td.+?/td>").Matches($ChildContent)) -replace "(<td>|</td>|&nbsp;)",""

                    $ContactPerson = $Text[0]
                    $PhoneNum = $Text[1]

                    #Write-Host "$ContactPersonTitle - $ContactPerson, $PhoneNumTitle - $PhoneNum" -BackgroundColor Green

                    #注册号/统一社会信用代码, 法定代表人
                    $LegalPresent = ([regex]"(?s)注册号.+?/tr>").Matches($ChildWeb.content);

                    $INF = (([regex]"<th.+?/th>").Matches($LegalPresent)) -replace "(<th>|</th>)",""

                    $RegisNumTitle = $INF[0];
                    $PresentTitle = $INF[1];

                    $Text = (([regex]"<td.+?/td>").Matches($LegalPresent)) -replace "(<td>|</td>|&nbsp;)",""
                    $RegisNum = $Text[0]
                    $Present = $Text[1]
                    #Write-Host "$RegisNumTitle - $RegisNum, $PresentTitle - $Present" -BackgroundColor red

                    #单位名称
                    $Enterprise = ([regex]"(?s)<th>单位名称.+?/tr>").Matches($ChildWeb.content);

                    $INF = (([regex]"<th.+?/th>").Matches($Enterprise)) -replace "(<th>|</th>)",""

                    $EnterpriseName = $INF[0];

                    $Text = (([regex]"<td.+?/td>").Matches($Enterprise)) -replace '(<td|</td|&nbsp;|colspan|=|3|>|")',""
                    $CompanyName = $Text[0].trim();

                    #Write-Host "$EnterpriseName  -  $CompanyName" -BackgroundColor DarkCyan

                    #污染源类别
                    $Pollution = ([regex]"(?s)<th>污染源编码.+?/tr>").Matches($ChildWeb.content);

                    $INF = (([regex]"<th.+?/th>").Matches($Pollution)) -replace "(<th>|</th>)",""

                    $PollutionTypeName = $INF[1]

                    $Text = (([regex]"<td.+?/td>").Matches($Pollution)) -replace '(<td|</td|&nbsp;|>)',""

                    $PollutionType = $Text[1].trim();

                    #Write-Host "$PollutionTypeName -- $PollutionType" -BackgroundColor red 

                    #生产地址
                    $Product = ([regex]"(?s)<th>生产地址.+?/tr>").Matches($ChildWeb.content);

                    $INF = (([regex]"<th.+?/th>").Matches($Product)) -replace "(<th>|</th>)",""

                    $ProductAddressTitle = $INF[0]
                    
                    $Text = (([regex]"<td.+?/td>").Matches($Product)) -replace '(<td|</td|&nbsp;|colspan|=|3|>|")',""

                    $ProductAddress = $Text[0]

                    #Write-Host "$ProductAddressTitle --- $ProductAddress" -BackgroundColor Red

                    #Write-Host $Sequence
                    #Write-Host $TitleSequence, $TitleName, $TitleDistrict, $TitleType, $ContactPersonTitle, $PhoneNumTitle, $RegisNumTitle, $PresentTitle, $EnterpriseName, $PollutionTypeName, $ProductAddressTitle -BackgroundColor Cyan

                    Write-Host "序号，企业名称，地区，污染类型，联系人名称，联系电话，注册号，法定代表人，单位名称，污染源类别，生产地址" -BackgroundColor DarkGray -ForegroundColor Cyan
                    Write-Host "$Sequence, $Company, $Distriction, $Type, $ContactPerson, $PhoneNum, $RegisNum, $Present, $CompanyName, $PollutionType, $ProductAddress" -BackgroundColor red -ForegroundColor green
                    "$Sequence* $Company* $Distriction* $Type* $ContactPerson* $PhoneNum* $RegisNum* $Present* $CompanyName* $PollutionType* $ProductAddress" >> "$Path\tmp.txt" 


            }

        }




}

function WriteToExcel() {
       param (
        [Parameter(Mandatory)]
        [string]
        $Path
        )


    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false;
    $workbook = $excel.Workbooks.add()
    $sheet = $workbook.worksheets.Item(1)
    $workbook.Worksheets.item(1).name="List"
    $sheet = $workbook.WorkSheets.Item("List")

    $row = 1


    foreach($content in Get-Content "$Path\tmp.txt" ){
            $col = 1
            $content -split "\*" | foreach {
                $sheet.cells.item($row, $col) = $_
                $col++
            }
            
            $row++
    }

    $date = Get-Date -Format 'yyyy-mm-dd_HH-mm-ss'
    $workbook.saveas("$Path\List_$date.xlsx");
    $workbook.application.quit();

}


#$Path = "D:\codedata\Powershell\爬虫\回忆-爬取网页数据"
$Path = $args[0]
$DateTime = "2020"
#$PAGE_SET_START = $args[0] -as [int]
#$PAGE_SET_END = $args[1] -as [int]

#第一种办法：
$Page = Invoke-Expression $args[1]
#第二种办法：
#$Page = powershell -command $args[1]

#Write-Host $args[0], $args[1], $args

"序号* 企业名称* 地区* 污染类型* 联系人名称* 联系电话* 注册号* 法定代表人* 单位名称* 污染源类别* 生产地址" > "$Path\tmp.txt"
# foreach ($i in $PAGE_SET) {
 #   $i | GetALLInformation -Path $Path -DateTime $DateTime
 #}

 #1..3 | GetALLInformation -Path $Path -DateTime $DateTime


#更改1..10，修改为你想要的页数
#暂时无法更改
$Page | GetALLInformation -Path $Path -DateTime $DateTime

#写入excel
WriteToExcel -Path $Path

Write-Host "Over" -BackgroundColor DarkGray -ForegroundColor Magenta