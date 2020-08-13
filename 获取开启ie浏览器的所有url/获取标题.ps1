$proce = (Get-Process 360chrome);

foreach ($str in $proce) {
    Write-Host $str.MainWindowTitle.ToString();
}