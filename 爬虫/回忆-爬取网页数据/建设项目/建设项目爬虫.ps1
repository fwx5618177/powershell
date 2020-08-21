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


function GetAllBody {
    param(
        [Parameter(mandatory)]
        [string]
        $Path,
        [Parameter(ValueFromPipeline=$true)]
        [int]
        $Page,
        [Parameter(Mandatory)]
        [int]
        $row,
        [Parameter(Mandatory)]
        [string]
        $FileName
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
           'selField'                = 'ST_XM_NAME'
           'selValue'                 = '' 
           'selFieldShowStr'          = '建设项目名称'
           'district'                 = ''
           'hptype'                 = ''
        }


        #url : https://xxgk.eic.sh.cn/jsp/view/eiaReportList.jsp
        
        
        #获取外面的数据列表
        $url = "https://xxgk.eic.sh.cn/jsp/view/eiaReportList.jsp"
        #获取网页内容
        $web = Invoke-WebRequest -Uri $url -Method Post -Body $body -Headers $headers

        #获取表单数据
        $table = ([regex]"(?s)<table.+Llist limit.+/table>").Matches($web.content) 
        
        #获取标题头
        ([regex]"(?s)<th.+th>").Matches($table)| ForEach-Object {
            $TitleHeaders = (([regex]">[^<].+</th").Matches($_.value)) -replace "(th|<|>|/)", ""
            
            $TitleSequence = $TitleHeaders[0]
            $TitleName = $TitleHeaders[1]
            $TitleSector = $TitleHeaders[2]
            $TitleDistrict = $TitleHeaders[3]
            $TitleType = $TitleHeaders[4]
            $TitlePublication = $TitleHeaders[5]
            $TitleDeadline = $TitleHeaders[6]

        }



        

        #获取每行数据
        ([regex]"(?S)<tr.+?>[^<].+?</tr>").Matches($table) | ForEach-Object {
            #$_.value

            #获取openinfo中的数据
            ([regex]"openInfo[^\)].+?\)").Matches($_.value) -replace "(openInfo|'|\(|\))","" | ForEach-Object {
                $info = $_ -split ","
                #id, 报文
                $id = $info[0]
                $msg = $info[1]
            }


            #获取行列的数据
            #([regex]"\>[^<].+?</td").Matches($_.value)
            $list = ([regex]"\>.+?</td").Matches($_.value) -replace "(>|</td|&nbsp;)", ""

            $Sequence = $list[0];
            $Name = $list[1];
            $Sector = $list[2];
            $ConstructionLocation = $list[3];
            $Type = $list[4];
            $Pulication = $list[5];
            $Deadline = $list[6];

            #Write-Host $list

            #进入子网抓取数据
            #url-deatil page: url="/jsxmxxgk/eiareport/action/jsxm_eiaReportDetail.do?from=jsxm&stEiaId="+id+"&type="+encodeURI(type); 
            #https://xxgk.eic.sh.cn/jsxmxxgk/eiareport/action/jsxm_eiaReportDetail.do?from=jsxm&stEiaId=2dcbffb1-a530-4c6d-81f7-7af87f4843b9&type=%E6%8A%A5%E5%91%8A%E8%A1%A8
            #$msg = encodeURL($msg);
            #Write-Host $msg
            if([string]::IsNullOrEmpty($id)){
                return;
            }

            $ChildUrl = "https://xxgk.eic.sh.cn/jsxmxxgk/eiareport/action/jsxm_eiaReportDetail.do?from=jsxm&stEiaId=$id&type=%E6%8A%A5%E5%91%8A%E8%A1%A8";

            $ChildWeb = Invoke-WebRequest -Uri $ChildUrl -Method Get -Headers $headers
            #$ChildWeb.content
            #获取页面的数据
            $DetailInformation = ([regex]"(?s)<div.+bCon active.+</div>").Matches($ChildWeb.content);

            #2.建设单位名称， 3.环评编制单位名称， 4.项目建设地点， 联 系 人， 联系电话
            #([regex]"(?S)<li>.+?2.建设单位名称.+?</li>").Matches($DetailInformation);
            #获取第二大类和第三大类
            $INFLIST = ([regex]"(?s)<ul.+?blist.+?</ul").Matches($DetailInformation) -replace "ul",""
            #第二大类
            $CompanyList = (([regex]"2.建设单位名称.+").Matches($INFLIST[0]) -replace '(<span class="colorBlue">|</span>)',"") -split "："
            $CompanyNameTitle = $CompanyList[0]
            $CompanyNameValue = $CompanyList[1]

            #Write-Host "$CompanyNameTitle  --- $CompanyNameValue"

            $SectorList = (([regex]"3.环评编制单位名称.+").Matches($INFLIST[0]) -replace '(<span class="colorBlue">|</span>)',"") -split "："
            $SectorTitle = $SectorList[0]
            $SectorValue = $SectorList[1]
            #Write-Host "$SectorTitle  --- $SectorValue "


            $ConstructList = (([regex]"4.项目建设地点.+").Matches($INFLIST[0]) -replace '(<span class="colorBlue">|</span>)',"") -split "："
            $ConstructTitle = $ConstructList[0]
            $ConstructValue = $ConstructList[1]
            #Write-Host "$ConstructTitle  --- $ConstructValue "

            #第三大类
            
            $Contact = (([regex]"联 系 人.+").Matches($INFLIST[1]) -replace '(<span class="colorBlue">|</span>)',"") -split "："
            $ContactName = $Contact[0]
            $ContactNum = $Contact[1]
            #Write-Host "$ContactName  --- $ContactNum "

            $ContactPhone = (([regex]"联系电话.+").Matches($INFLIST[1]) -replace '(<span class="colorBlue">|</span>)',"") -split "："
            $ContactPhoneName = $ContactPhone[0]
            $ContactPhoneNum = $ContactPhone[1]
            #Write-Host "$ContactPhoneName  --- $ContactPhoneNum "


            Write-Host "$TitleSequence, $TitleName, $TitleSector, $TitleDistrict, $TitleType, $TitlePublication, $TitleDeadline, $CompanyNameTitle, $SectorTitle, $ConstructTitle, $ContactName, $ContactPhoneName" -BackgroundColor Gray -ForegroundColor Green

            Write-Host "$Sequence* $Name* $Sector* $ConstructionLocation* $Type* $Pulication* $Deadline* $CompanyNameValue* $SectorValue* $ConstructValue* $ContactNum* $ContactPhoneNum" -BackgroundColor DarkRed -ForegroundColor Cyan

            #"$Sequence* $Name* $Sector* $ConstructionLocation* $Pulication* $Deadline* $CompanyNameValue* $SectorValue* $ConstructValue* $ContactNum* $ContactPhoneNum" >> "$Path\tmp.txt"
            
            $DataList = ("$Sequence", "$Name", "$Sector", "$ConstructionLocation", "$Type", "$Pulication", "$Deadline", "$CompanyNameValue", "$SectorValue", "$ConstructValue", "$ContactNum", "$ContactPhoneNum")

            if([string]::IsNullOrEmpty($DataList[0])){
                return
            }

            WriteToExcel -Path $Path -row $row -DataList $DataList -FileName $FileName
            $row++

            
        }


    }

}

function WriteToExcel() {
       param (
        [Parameter(Mandatory)]
        [string]
        $Path,
        [Parameter(Mandatory)]
        [int]
        $row,
        [Parameter(Mandatory)]
        [array]
        $DataList,
        [Parameter(Mandatory)]
        [string]
        $FileName
        )


    $sheet = $workbook.WorkSheets.Item("List")



#    foreach($content in Get-Content "$Path\tmp.txt" ){
#            $col = 1
#            $content -split "\*" | foreach {
#                $sheet.cells.item($row, $col) = $_
#                $col++
#            }
#            
#    }

    $col = 1
    foreach($item in $DataList) {
        $sheet.cells.item($row, $col) = $item
        $col++
    }


}


#$Path = "D:\codedata\Powershell\爬虫\回忆-爬取网页数据\建设项目"

$Path = $args[0]
$Page = Invoke-Expression $args[1]

#创建文件
$date = Get-Date -Format 'yyyy-mm-dd_HH-mm-ss'
$FileName = "$Path\List_$date.xlsx"

#"序号* 建设项目名称* 审核部门* 建设地点* 环评类别* 发布时间* 公示截止日期* 2.建设单位名称* 3.环评编制单位名称* 4.项目建设地点* 联 系 人* 联系电话" > "$Path\tmp.txt"
#("序号", "建设项目名称", "审核部门", "建设地点", "环评类别", "发布时间", "公示截止日期", "2.建设单位名称", "3.环评编制单位名称", "4.项目建设地点", "联 系 人", "联系电话")

$ExcelHeaders = ("序号", "建设项目名称", "审核部门", "建设地点", "环评类别", "发布时间", "公示截止日期", "2.建设单位名称", "3.环评编制单位名称", "4.项目建设地点", "联 系 人", "联系电话");


#Excel 设置
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true;
    $excel.DisplayAlerts = $false;
    $workbook = $excel.Workbooks.add()
    $workbook.Worksheets.item(1).name="List"
#写入标题头
$row = 1
WriteToExcel -Path $Path -row $row -DataList $ExcelHeaders -FileName $FileName
$row++
$Page | GetAllBody -Path $Path -row $row -FileName $FileName
#写入excel
#WriteToExcel -Path $Path

#最后关闭excel
$workbook.saveas($FileName);
$workbook.application.quit();