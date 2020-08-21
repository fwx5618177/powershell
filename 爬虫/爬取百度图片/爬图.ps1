#This is a crawler for baidu images
<#
All these funuctions would be used to implementation crawler's target
#>

#��ȡ����ַ���Ȼ����ΪͼƬ������
function Get-Random-String {
    $fileName = -join ([char[]](65..90 + 97..122) | Get-Random -Count 6)
    $fileName
}

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

#��ȡĿ¼�µ������ļ�������MD5����
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

#��ȡ��Ƭ
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
        #��ȡ��֪�����ļ���MD5�����Ҽ�¼
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
        #ҳ��
        $n = ($page * 20)
        #�ؼ���
        $word = [uri]::EscapeDataString($keyword)
        
        #url
        #$url = "https://image.baidu.com/search/index?tn=baiduimage&ps=1&ct=201326592&lm=-1&cl=2&nc=1&ie=utf-8&word=%E7%BE%8E%E5%A5%B3"
        $url = "https://image.baidu.com/search/flip?tn=baiduimage&ie=utf-8&word=${word}&pn=$n"
        #������
        Write-Host "Handling $url`n"
        
        #����ҳ����ȡ����
        $web = (Invoke-WebRequest -Uri $url -Method GET -Headers $headers)
        #�����ȡ����ȡ��ҳ��ַ
        #��ȡ��ַ�󣬽���ɸѡ������з��Ϲ�����ڽ��г�ȡͼƬ��ַ
        #ѡ��ͼƬ��ַ
        $web | Select-String '"objURL":"https?://+[^\s]+[\w]' -AllMatches | ForEach-Object { $_.Matches } | Foreach-Object {
            $_ -match 'https?://.+.'
            #ƥ���ֵ���б�������ɸѡ
            $Matches.Values | ForEach-Object {
                Write-Host "Fetching from $_" -ForegroundColor 3
                #ɸѡ��ͼƬ�ĵ�ַ
                $ext=([regex]'\.(jpe?g|png|gif|tif|bmp)').Match($_).Value
                if ([String]::IsNullOrEmpty($ext)) {
                    $ext=".jpg"
                }

                #��ͼƬ������������
                $fileFullName = (Get-Random-String) + $ext
                #ָ���ļ�Ŀ¼+�ļ���             
                $TargetPath = Join-Path -Path $Path -ChildPath $fileFullName
                #����ͼƬ��д��
                Invoke-WebRequest -Uri $_ -PassThru -TimeoutSec 20000 -OutFile $TargetPath -ErrorAction SilentlyContinue

                # ����md5��ֵ��У׼
                if ((Test-Path $TargetPath)) {
                    $hashValue = (Get-FileHash -Path $TargetPath -Algorithm MD5).Hash
                    #����ļ��Ѿ������ˣ���ɾ��
                    if($hashValue -and $hashTable.ContainsKey($hashValue)){
                        Remove-Item -Path $TargetPath -Force -ErrorAction SilentlyContinue
                    }
                }
                #��ͣ1��
                Start-Sleep -Milliseconds 50000
            }
        }
        #ii $Path
    }
}

$DIRPATH = $args[0]
$KeyWord = $args[1]
$Pages = Invoke-Expression $args[2]

#ָ��ǰ10ҳ
$Pages |Get-Images -Path "$DIRPATH\$KeyWord" -keyword "$KeyWord"