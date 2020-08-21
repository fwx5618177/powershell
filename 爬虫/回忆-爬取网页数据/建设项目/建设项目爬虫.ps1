#����·���Ƿ���ڣ���������ڣ��򴴽�һ��
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
           'selFieldShowStr'          = '������Ŀ����'
           'district'                 = ''
           'hptype'                 = ''
        }


        #url : https://xxgk.eic.sh.cn/jsp/view/eiaReportList.jsp
        
        
        #��ȡ����������б�
        $url = "https://xxgk.eic.sh.cn/jsp/view/eiaReportList.jsp"
        #��ȡ��ҳ����
        $web = Invoke-WebRequest -Uri $url -Method Post -Body $body -Headers $headers

        #��ȡ������
        $table = ([regex]"(?s)<table.+Llist limit.+/table>").Matches($web.content) 
        
        #��ȡ����ͷ
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



        

        #��ȡÿ������
        ([regex]"(?S)<tr.+?>[^<].+?</tr>").Matches($table) | ForEach-Object {
            #$_.value

            #��ȡopeninfo�е�����
            ([regex]"openInfo[^\)].+?\)").Matches($_.value) -replace "(openInfo|'|\(|\))","" | ForEach-Object {
                $info = $_ -split ","
                #id, ����
                $id = $info[0]
                $msg = $info[1]
            }


            #��ȡ���е�����
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

            #��������ץȡ����
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
            #��ȡҳ�������
            $DetailInformation = ([regex]"(?s)<div.+bCon active.+</div>").Matches($ChildWeb.content);

            #2.���赥λ���ƣ� 3.�������Ƶ�λ���ƣ� 4.��Ŀ����ص㣬 �� ϵ �ˣ� ��ϵ�绰
            #([regex]"(?S)<li>.+?2.���赥λ����.+?</li>").Matches($DetailInformation);
            #��ȡ�ڶ�����͵�������
            $INFLIST = ([regex]"(?s)<ul.+?blist.+?</ul").Matches($DetailInformation) -replace "ul",""
            #�ڶ�����
            $CompanyList = (([regex]"2.���赥λ����.+").Matches($INFLIST[0]) -replace '(<span class="colorBlue">|</span>)',"") -split "��"
            $CompanyNameTitle = $CompanyList[0]
            $CompanyNameValue = $CompanyList[1]

            #Write-Host "$CompanyNameTitle  --- $CompanyNameValue"

            $SectorList = (([regex]"3.�������Ƶ�λ����.+").Matches($INFLIST[0]) -replace '(<span class="colorBlue">|</span>)',"") -split "��"
            $SectorTitle = $SectorList[0]
            $SectorValue = $SectorList[1]
            #Write-Host "$SectorTitle  --- $SectorValue "


            $ConstructList = (([regex]"4.��Ŀ����ص�.+").Matches($INFLIST[0]) -replace '(<span class="colorBlue">|</span>)',"") -split "��"
            $ConstructTitle = $ConstructList[0]
            $ConstructValue = $ConstructList[1]
            #Write-Host "$ConstructTitle  --- $ConstructValue "

            #��������
            
            $Contact = (([regex]"�� ϵ ��.+").Matches($INFLIST[1]) -replace '(<span class="colorBlue">|</span>)',"") -split "��"
            $ContactName = $Contact[0]
            $ContactNum = $Contact[1]
            #Write-Host "$ContactName  --- $ContactNum "

            $ContactPhone = (([regex]"��ϵ�绰.+").Matches($INFLIST[1]) -replace '(<span class="colorBlue">|</span>)',"") -split "��"
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


#$Path = "D:\codedata\Powershell\����\����-��ȡ��ҳ����\������Ŀ"

$Path = $args[0]
$Page = Invoke-Expression $args[1]

#�����ļ�
$date = Get-Date -Format 'yyyy-mm-dd_HH-mm-ss'
$FileName = "$Path\List_$date.xlsx"

#"���* ������Ŀ����* ��˲���* ����ص�* �������* ����ʱ��* ��ʾ��ֹ����* 2.���赥λ����* 3.�������Ƶ�λ����* 4.��Ŀ����ص�* �� ϵ ��* ��ϵ�绰" > "$Path\tmp.txt"
#("���", "������Ŀ����", "��˲���", "����ص�", "�������", "����ʱ��", "��ʾ��ֹ����", "2.���赥λ����", "3.�������Ƶ�λ����", "4.��Ŀ����ص�", "�� ϵ ��", "��ϵ�绰")

$ExcelHeaders = ("���", "������Ŀ����", "��˲���", "����ص�", "�������", "����ʱ��", "��ʾ��ֹ����", "2.���赥λ����", "3.�������Ƶ�λ����", "4.��Ŀ����ص�", "�� ϵ ��", "��ϵ�绰");


#Excel ����
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true;
    $excel.DisplayAlerts = $false;
    $workbook = $excel.Workbooks.add()
    $workbook.Worksheets.item(1).name="List"
#д�����ͷ
$row = 1
WriteToExcel -Path $Path -row $row -DataList $ExcelHeaders -FileName $FileName
$row++
$Page | GetAllBody -Path $Path -row $row -FileName $FileName
#д��excel
#WriteToExcel -Path $Path

#���ر�excel
$workbook.saveas($FileName);
$workbook.application.quit();