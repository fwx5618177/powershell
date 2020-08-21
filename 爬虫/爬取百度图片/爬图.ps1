#This is a crawler for baidu images
<#
All these funuctions would be used to implementation crawler's target
#>

#获取随机字符，然后作为图片的名字
function Get-Random-String {
    $fileName = -join ([char[]](65..90 + 97..122) | Get-Random -Count 6)
    $fileName
}

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

#获取目录下的所有文件，并且MD5计算
function Invoke-MD5 {
    param (
        # Parameter Path
        [Parameter(Mandatory)]
        [string]
        $Path
    )
    begin {
        $global:hashTable = @{ }
    }
    process {
        Get-ChildItem -Path $Path | Where-Object {
            $hash = Get-FileHash -Path $_.FullName -Algorithm MD5
            $hashTable[$hash.Hash] = $hash.Path
        }
    }
    end { }
}

#获取照片
function Get-Images {
    param (
        [Parameter(ValueFromPipeline)]
        [int]
        $page = 1,
        [Parameter(Mandatory)]
        [string]
        $Path,
        [Parameter(Mandatory)]
        [string]
        $keyword
    )
    begin {
        Resolve-Directory -Path $Path
        #获取已知所有文件的MD5，并且记录
        Invoke-MD5 -Path $Path
        $headers = @{
            'Accept'                    = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3'
            'Accept-Encoding'           = 'gzip, deflate, br'
            'Accept-Language'           = 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7'
            'Host'                      = 'image.baidu.com'
            'Upgrade-Insecure-Requests' = '1'
            'User-Agent'                = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36'
        }
    }
    process {
        #页码
        $n = ($page * 20)
        #关键字
        $word = [uri]::EscapeDataString($keyword)
        
        #url
        #$url = "https://image.baidu.com/search/index?tn=baiduimage&ps=1&ct=201326592&lm=-1&cl=2&nc=1&ie=utf-8&word=%E7%BE%8E%E5%A5%B3"
        $url = "https://image.baidu.com/search/flip?tn=baiduimage&ie=utf-8&word=${word}&pn=$n"
        #处理结果
        Write-Host "Handling $url`n"
        
        #打开网页，获取内容
        $web = (Invoke-WebRequest -Uri $url -Method GET -Headers $headers)
        #正则抽取，获取网页地址
        #获取地址后，进行筛选。如果有符合规则的在进行抽取图片地址
        #选中图片地址
        $web | Select-String '"objURL":"https?://+[^\s]+[\w]' -AllMatches | ForEach-Object { $_.Matches } | Foreach-Object {
            $_ -match 'https?://.+.'
            #匹配的值进行遍历和再筛选
            $Matches.Values | ForEach-Object {
                Write-Host "Fetching from $_" -ForegroundColor 3
                #筛选出图片的地址
                $ext=([regex]'\.(jpe?g|png|gif|tif|bmp)').Match($_).Value
                if ([String]::IsNullOrEmpty($ext)) {
                    $ext=".jpg"
                }

                #对图片进行重新命名
                $fileFullName = (Get-Random-String) + $ext
                #指定文件目录+文件名             
                $TargetPath = Join-Path -Path $Path -ChildPath $fileFullName
                #下载图片并写入
                Invoke-WebRequest -Uri $_ -PassThru -TimeoutSec 20000 -OutFile $TargetPath -ErrorAction SilentlyContinue

                # 计算md5的值，校准
                if ((Test-Path $TargetPath)) {
                    $hashValue = (Get-FileHash -Path $TargetPath -Algorithm MD5).Hash
                    #如果文件已经存在了，就删除
                    if($hashValue -and $hashTable.ContainsKey($hashValue)){
                        Remove-Item -Path $TargetPath -Force -ErrorAction SilentlyContinue
                    }
                }
                #暂停1秒
                Start-Sleep -Milliseconds 50000
            }
        }
        #ii $Path
    }
}

$DIRPATH = $args[0]
$KeyWord = $args[1]
$Pages = Invoke-Expression $args[2]

#指定前10页
$Pages |Get-Images -Path "$DIRPATH\$KeyWord" -keyword "$KeyWord"