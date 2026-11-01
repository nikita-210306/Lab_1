param(
    [string]$Destination = 'D:\ФБИУКС\R\pract'
)
$ErrorActionPreference = 'Stop'
$BaseUrl = 'https://raw.githubusercontent.com/nikita-210306/Lab_1/main/penguins_bootstrap'
$Files = @(
    @{ Name = 'Bootstrap_Markin.Rmd'; SHA256 = 'fbaa7e3d135b0b443487b7b84516c5ed2dd61314869576faf1d9bfcb2eb97e84' }
    @{ Name = 'Bootstrap_Andreus.Rmd'; SHA256 = 'e36afc1a71c8e87324918eb861168a586a1d437083ce901b90c5796dc22611bd' }
    @{ Name = 'Bootstrap_Markin.html'; SHA256 = 'c039e6552e2aced265ff8d3c6445bd08041998187db9ac90cd390cef6e2ec613' }
    @{ Name = 'Bootstrap_Andreus.html'; SHA256 = '4b01fa34a5b7d0ef3e7aeb5b5c48f6a1dde615e9c5c472d70441152712bbc8fa' }
    @{ Name = 'Explanation_and_speech.md'; SHA256 = '2f7cba3f77d4398b5fad01154b1a1d6c670b9c022fe7e7439e90085759f7a8a5' }
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
