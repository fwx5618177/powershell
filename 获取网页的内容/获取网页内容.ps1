$web = Invoke-WebRequest -Uri "http://www.baidu.com"
$content = $web.Content | ConvertFrom-Json 

#Write-Host $web
Write-Host $content