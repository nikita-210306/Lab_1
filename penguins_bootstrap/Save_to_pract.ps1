param(
    [string]$Destination = 'D:\ФБИУКС\R\pract'
)
$ErrorActionPreference = 'Stop'
$BaseUrl = 'https://raw.githubusercontent.com/nikita-210306/Lab_1/main/penguins_bootstrap'
$Files = @(
    @{ Name = 'Bootstrap_Markin.Rmd'; SHA256 = '259f1372abd895d73077d9def18d3af4b63a8bcf608065a83f621015f35f0a23' }
    @{ Name = 'Bootstrap_Andreus.Rmd'; SHA256 = 'd7f94315901f456c0934f5b474d808ccde0cff61cbaaaef49ae1fe5ddb1d983b' }
    @{ Name = 'Bootstrap_Markin.html'; SHA256 = '0c40d83af2e240d8a5b5704fb11140355dfca1c3f57b58681993caff52d582b7' }
    @{ Name = 'Bootstrap_Andreus.html'; SHA256 = 'd6d3c87e16996b74fc8bf150c586136ebafd917ba4099a3497dcdc906ceae3eb' }
    @{ Name = 'Explanation_and_speech.md'; SHA256 = 'c4d77834ab5de9257f89146dac70e8aa282c17124dbe08f586963490097b42be' }
)

$Root = [System.IO.Path]::GetPathRoot($Destination)
if (-not (Test-Path -LiteralPath $Root)) {
    throw "Диск назначения недоступен: $Root"
}
New-Item -ItemType Directory -Path $Destination -Force | Out-Null
$TemporaryFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('penguins-bootstrap-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $TemporaryFolder | Out-Null
try {
    # Сначала загрузить и проверить все файлы, не трогая существующие.
    foreach ($Item in $Files) {
        $Downloaded = Join-Path $TemporaryFolder $Item.Name
        Invoke-WebRequest -UseBasicParsing -Uri ($BaseUrl + '/' + $Item.Name) -OutFile $Downloaded
        $ActualHash = (Get-FileHash -LiteralPath $Downloaded -Algorithm SHA256).Hash
        if ($ActualHash -ine $Item.SHA256) {
            throw "Не совпала контрольная сумма: $($Item.Name)"
        }
    }
    $Stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    foreach ($Item in $Files) {
        $Target = Join-Path $Destination $Item.Name
        if (Test-Path -LiteralPath $Target) {
            Copy-Item -LiteralPath $Target -Destination ($Target + '.backup_' + $Stamp)
        }
        Copy-Item -LiteralPath (Join-Path $TemporaryFolder $Item.Name) -Destination $Target -Force
        Write-Host ('Сохранен: ' + $Target)
    }
    Write-Host 'Готово. Rmd открываются в RStudio, HTML — в браузере.'
} finally {
    Remove-Item -LiteralPath $TemporaryFolder -Recurse -Force
}
