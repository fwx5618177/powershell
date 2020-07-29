function searchExcelStr([string]$DIRPATH, [string]$FILEFIX, [string]$SEARCHSTRING){
    cd $DIRPATH;

    #GET File-List
    $FILELIST = Get-ChildItem $DIRPATH -Recurse $FILEFIX | %{$_.FullName};

    #Operation Excel
    $excel = New-Object -ComObject excel.application;
    $excel.visible = $false;
    $excel.displayalerts = $false;

    foreach($FILE in $FILELIST){
        
        #Open EXCEL
        $wb = $excel.Workbooks.open($FILE);

        #Output sheet name
        $info = $wb.sheets.count();

        $nameList = $wb.sheets | Select-Object -Property name

        for($i = 1; $i -le $info; $i++){
            
            #select sheet
            $sheet = $wb.sheets.item($i);

            #Get sheets' name
            $sheetName_now = $sheet | Select-Object -Property name

            $sheet.select();

            #Search range
            $searchRange = $sheet.UsedRange;

            $target = $searchRange.find($SEARCHSTRING);

            if($target -eq $null){
                continue;
            }
            $target.select();

            #First Search Result
            $First = $target;

            do{
                
                #Get row, columns
                $row_num = $target.row;
                $column_num = $target.column;

                Write-Host "String: $SEARCHSTRING , CELL($row_num, $column_num), Sheet name: $sheetName_now, File: $FILE" -BackgroundColor DarkCyan
                Write-Output "String: $SEARCHSTRING , CELL($row_num, $column_num), Sheet name: $sheetName_now, File: $FILE" | Out-File -Append -FilePath "$DIRPATH\result.txt";

                $target = $searchRange.findnext($target);
            }while($target -ne $null -and $target.row -ne $First.row)
        }

        $wb.application.quit();
        Write-Output "-------------------文件分割线-----------" | Out-File -Append -FilePath "$DIRPATH\result.txt";
        Write-Host "-------------------文件分割线-----------" -BackgroundColor DarkGray
    }
}

$DIRPATH = Read-Host "path:"
$FILEFIX = Read-Host "file-fix:"
$SEARCHSTRING = Read-Host "search str:"

if ($DIRPATH -eq $null){
    $DIRPATH = Split-Path -Parent $MyInvocation.MyCommand.Definition;
}

searchExcelStr -DIRPATH $DIRPATH -FILEFIX $FILEFIX -SEARCHSTRING $SEARCHSTRING
Pause