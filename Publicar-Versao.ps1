<#
  Publicar-Versao.ps1 - publica uma nova versao do manifest do XLerate Pessoal.

  Use quando mudar algo no manifest (atalho novo, ribbon, permissoes).
  Mudanca so de codigo NAO precisa disto: basta git push (o Pages atualiza).

  O que faz, em ordem:
    1. Confere que o Excel esta fechado e que o git esta limpo
    2. Sobe a <Version> do XLerate\manifest.xml (padrao: 1.1.2.0 -> 1.1.3.0)
    3. Roda npm run ci:all
    4. Commit + push para "mine" (dispara o deploy do GitHub Pages)
    5. Atualiza a versao dentro de Ativar-XLerate.xlsx (webextension1.xml)
    6. Copia manifest + xlsx para:
         %LOCALAPPDATA%\XLerate Pessoal      (o que o Excel carrega)
         OneDrive\Documents\Add-in's         (backup)
         OneDrive\...\XLerate-para-compartilhar  e refaz XLerate-Pessoal.zip
    7. Limpa o cache de add-ins do Office (Wef)

  Exemplos:
    .\Publicar-Versao.ps1                    # sobe o 3o numero
    .\Publicar-Versao.ps1 -Versao 1.2.0.0
    .\Publicar-Versao.ps1 -SemPush           # tudo local, sem commit/push
    .\Publicar-Versao.ps1 -SoLimparCache     # so limpa o cache do XLerate
#>
param(
  [string]$Versao,
  [switch]$SemPush,
  [switch]$SoLimparCache
)

$ErrorActionPreference = 'Stop'
$repo      = $PSScriptRoot
$manifest  = Join-Path $repo 'XLerate\manifest.xml'
$addinsDir = Join-Path $env:USERPROFILE "OneDrive - BFCAPITAL\Documents\Add-in's"
$shareDir  = Join-Path $addinsDir 'XLerate-para-compartilhar'
$localDir  = Join-Path $env:LOCALAPPDATA 'XLerate Pessoal'
$gh        = Join-Path $env:LOCALAPPDATA 'gh-cli\bin\gh.exe'
$masterXlsx = Join-Path $addinsDir 'Ativar-XLerate.xlsx'

function Step($msg) { Write-Host "`n==> $msg" -ForegroundColor Cyan }

function Clear-XLerateCache {
  # 7. Cache do Office - SO o que e do XLerate.
  #    NUNCA apagar a pasta Wef inteira: Wef\webview2 guarda os logins dos
  #    outros add-ins (Claude, GPT...) e a pasta e compartilhada com Word/PPT.
  #    Mapeado em 2026-11-17:
  #      Wef\{GUID}\<hash>\Manifests\<id>_<versao>
  #      Wef\{GUID}\<hash>\ExtendedManifest\<id>_<versao>_<lang>
  #      Wef\AggregatedCache\<v>\<hash>\ShortcutsMapping.Excel.*  (mapa de atalhos
  #        agregado de todos os add-ins; o Excel recria, nao afeta login)
  #    Caminhos passam de 260 chars -> usar prefixo \\?\.
  $addinId = ([xml](Get-Content $manifest -Raw -Encoding UTF8)).OfficeApp.Id
  $wef = '\\?\' + (Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef')
  $alvos = @()
  if (Test-Path -LiteralPath $wef) {
    $alvos += Get-ChildItem -LiteralPath $wef -Directory -Filter '{*}' -ErrorAction SilentlyContinue |
      ForEach-Object { Get-ChildItem -LiteralPath $_.FullName -Directory -ErrorAction SilentlyContinue } |
      ForEach-Object { foreach ($sub in 'Manifests', 'ExtendedManifest') {
        $d = "$($_.FullName)\$sub"
        if (Test-Path -LiteralPath $d) { Get-ChildItem -LiteralPath $d -Force | Where-Object Name -like "$addinId*" } } }
    $agg = "$wef\AggregatedCache"
    if (Test-Path -LiteralPath $agg) {
      $alvos += Get-ChildItem -LiteralPath $agg -Recurse -File -Force -ErrorAction SilentlyContinue |
        Where-Object Name -like 'ShortcutsMapping.Excel.*'
    }
  }
  $falhas = 0
  foreach ($a in $alvos) {
    if ($a.FullName -match '\\webview2\\') { continue }   # trava de seguranca
    try { Remove-Item -LiteralPath $a.FullName -Recurse -Force -ErrorAction Stop; Write-Host "  apagado: $($a.Name)" }
    catch { $falhas++; Write-Host "  em uso: $($a.Name)" -ForegroundColor Yellow }
  }
  if (-not $alvos) { Write-Host '  nada do XLerate no cache (ok)' }
  if ($falhas) { Write-Host 'Aviso: alguns arquivos estavam em uso. Feche o Excel e rode: .\Publicar-Versao.ps1 -SoLimparCache.' -ForegroundColor Yellow }
}

if ($SoLimparCache) {
  if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Feche o Excel antes.' }
  Step 'Limpando cache do XLerate'; Clear-XLerateCache; return
}

# Node (winget) nem sempre esta no PATH da sessao
$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
            [Environment]::GetEnvironmentVariable('Path', 'User')

# 1. Pre-checagens
Step 'Pre-checagens'
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Feche o Excel antes de publicar.' }
Push-Location $repo
try {
  $dirty = git status --porcelain
  if ($dirty) { throw "Ha mudancas nao commitadas no repo. Commite antes:`n$dirty" }
} finally { Pop-Location }

# 2. Nova versao
[xml]$xml = Get-Content $manifest -Raw -Encoding UTF8
$atual = $xml.OfficeApp.Version
if (-not $Versao) {
  $p = $atual.Split('.') | ForEach-Object { [int]$_ }
  $p[2]++; $p[3] = 0
  $Versao = $p -join '.'
}
if ($Versao -notmatch '^\d+\.\d+\.\d+\.\d+$') { throw "Versao invalida: $Versao" }
Step "Versao $atual -> $Versao"
$texto = [IO.File]::ReadAllText($manifest)
$texto = $texto -replace "<Version>$([regex]::Escape($atual))</Version>", "<Version>$Versao</Version>"
[IO.File]::WriteAllText($manifest, $texto, (New-Object Text.UTF8Encoding($false)))

# 3. Testes + build
Step 'npm run ci:all'
Push-Location (Join-Path $repo 'XLerate')
try {
  npm run ci:all
  if ($LASTEXITCODE -ne 0) {
    git checkout -- manifest.xml
    throw 'ci:all falhou. Versao do manifest revertida.'
  }
} finally { Pop-Location }

# 4. Commit + push
if (-not $SemPush) {
  Step 'Commit + push para mine/master'
  Push-Location $repo
  try {
    git add XLerate/manifest.xml
    git commit -m "Bump manifest to $Versao" -q
    $token = & $gh auth token
    $b64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("x-access-token:$token"))
    git -c "http.https://github.com/.extraheader=AUTHORIZATION: basic $b64" push mine master
    if ($LASTEXITCODE -ne 0) { throw 'git push falhou.' }
  } finally { Pop-Location }
} else {
  Write-Host '(-SemPush: manifest alterado mas nao commitado)' -ForegroundColor Yellow
}

# 5. Versao dentro do Ativar-XLerate.xlsx
Step 'Atualizando Ativar-XLerate.xlsx'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::Open($masterXlsx, 'Update')
try {
  $name  = 'xl/webextensions/webextension1.xml'
  $entry = $zip.GetEntry($name)
  $r = New-Object IO.StreamReader($entry.Open()); $we = $r.ReadToEnd(); $r.Close()
  # so a tag we:reference (nao mexer no <?xml version="1.0"?>)
  $we = $we -replace '(<we:reference [^>]*?version=")[\d\.]+"', "`${1}$Versao"""
  if ($we -notmatch "version=""$([regex]::Escape($Versao))""") { throw 'Nao achei a versao em webextension1.xml.' }
  $entry.Delete()
  $w = New-Object IO.StreamWriter($zip.CreateEntry($name).Open(), (New-Object Text.UTF8Encoding($false)))
  $w.Write($we); $w.Close()
} finally { $zip.Dispose() }

# 6. Copias + zip para colegas
Step 'Copiando manifest e xlsx'
New-Item -ItemType Directory -Force $localDir | Out-Null
foreach ($d in $localDir, $addinsDir, $shareDir) {
  Copy-Item $manifest (Join-Path $d 'manifest.xml') -Force
  if ($d -ne $addinsDir) { Copy-Item $masterXlsx (Join-Path $d 'Ativar-XLerate.xlsx') -Force }
}
$zipOut = Join-Path $addinsDir 'XLerate-Pessoal.zip'
Compress-Archive -Path (Join-Path $shareDir '*') -DestinationPath $zipOut -Force

# 7. Cache do Office (so o do XLerate - ver Clear-XLerateCache)
Step 'Limpando cache do XLerate (sem tocar nos outros add-ins)'
Clear-XLerateCache

Write-Host "`nPronto: versao $Versao publicada." -ForegroundColor Green
if (-not $SemPush) { Write-Host 'Espere ~1 min o deploy do Pages antes de abrir o Excel.' }
Write-Host "Colegas: mandar o novo $zipOut (precisam reinstalar)."
