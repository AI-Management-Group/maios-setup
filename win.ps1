# MAIOS オリエン用の一括セットアップ（Windows）
#
# オリエンの場で、PowerShell（管理者ではない普通の PowerShell）に次の1行を貼り付けて実行する：
#   iex ([Text.Encoding]::UTF8.GetString((New-Object Net.WebClient).DownloadData('<配布 URL>/maios_setup_win.ps1')))
#   （文字化けしないよう、UTF-8 として読み込んでから実行する形にしている）
#
# やること（何度実行しても同じ結果になる・管理者権限は要らない）：
#   1. Python 3.12 を「この利用者だけ」に入れる（%LOCALAPPDATA%\Programs\Python\Python312。MAIOS のフックはここを PATH に無くても拾う）
#      ARM 版 Windows でも x64 版を入れる（部品の配布物が x64 にしか無いものがあるため。Windows 11 は x64 をそのまま動かせる）
#   2. スキルが使う部品7つを入れる（python-pptx openpyxl pandas PyYAML jpholiday Pillow playwright）
#   3. Claude アプリと Google ドライブ（パソコン版）があるかを確かめる。どちらもこのスクリプトでは入れない（事前案内で各自が入れる）
#      Google ドライブは必須＝無ければ ×。Claude アプリはこの1行を走らせている間に入れてもよいので、無くても △ にとどめる
# Git は入れない（Claude Code は Git が無ければ PowerShell で動く・公式 docs）。
# 最後に ○／× の一覧を出す（日本語の Windows の画面は絵文字を出せないことがあるため記号にする）。記録は %LOCALAPPDATA%\maios_setup.log。
#
# 環境変数（AMG・テスト用）：MAIOS_SETUP_FORCE_PYTHON=1 で既存の Python があっても入れる／MAIOS_SETUP_SKIP_APPS=1 でアプリの確認を省く

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}

$PyVer = '3.12.10'
$Pkgs = @('python-pptx', 'openpyxl', 'pandas', 'PyYAML', 'jpholiday', 'Pillow', 'playwright')
$Imports = 'import pptx, openpyxl, pandas, yaml, jpholiday, PIL, playwright'
$Log = Join-Path $env:LOCALAPPDATA 'maios_setup.log'

function Log([string]$m) { try { Add-Content -Path $Log -Value ("{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m) -Encoding UTF8 } catch {} }
function Step([string]$m) { Write-Host ''; Write-Host ('▶ ' + $m); Log ('== ' + $m) }

function Test-Python([string]$Exe, [string[]]$Pre) {
  # 3.9 以上で動けば実行ファイルの場所を返す（Microsoft Store の身代わり python.exe は弾く）
  if (-not $Exe) { return $null }
  if ($Exe -like '*\Microsoft\WindowsApps\*') { return $null }
  try {
    $out = & $Exe @Pre -c "import sys; print('maios-ok|' + sys.executable if sys.version_info >= (3, 9) else 'old')" 2>$null
    $line = @($out | Where-Object { $_ -like 'maios-ok|*' })[0]
    if ($line) { return $line.Substring(9).Trim() }
  } catch {}
  return $null
}

function Find-Python {
  $cands = New-Object System.Collections.ArrayList
  $base = Join-Path $env:LOCALAPPDATA 'Programs\Python'
  if (Test-Path $base) {
    # 3.12 を先に、ほかは新しい順（名前の並びだと Python39 が Python312 より先に来るので数字で並べる）
    Get-ChildItem -Path $base -Filter 'Python3*' -Directory -ErrorAction SilentlyContinue |
      Sort-Object @{ Expression = { if ($_.Name -eq 'Python312') { 0 } else { 1 } } }, @{ Expression = { [int](($_.Name -replace '\D', '') + '0') }; Descending = $true } |
      ForEach-Object { [void]$cands.Add(@((Join-Path $_.FullName 'python.exe'), @())) }
  }
  if (-not $env:MAIOS_SETUP_FORCE_PYTHON) {
    foreach ($n in @('py', 'python', 'python3')) {
      $c = Get-Command $n -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
      if ($c) { if ($n -eq 'py') { [void]$cands.Add(@($c.Source, @('-3'))) } else { [void]$cands.Add(@($c.Source, @())) } }
    }
  }
  foreach ($c in $cands) { $e = Test-Python $c[0] $c[1]; if ($e) { return $e } }
  return $null
}

$R_PY = '×'; $R_PKG = '×'; $R_APP = '－'; $R_DRIVE = '－'
Write-Host ''
Write-Host 'MAIOS のセットアップを始めます（5〜10分）。終わるまで、この画面を閉じないでください。'
Log ("start {0} {1} ps{2}" -f [Environment]::OSVersion.VersionString, $env:PROCESSOR_ARCHITECTURE, $PSVersionTable.PSVersion)

# 1. Python
Step '1/3 Python 3.12'
$py = Find-Python
if ($py -and $env:MAIOS_SETUP_FORCE_PYTHON -and ($py -notlike (Join-Path $env:LOCALAPPDATA 'Programs\Python\*'))) { $py = $null }
if (-not $py) {
  $url = "https://www.python.org/ftp/python/$PyVer/python-$PyVer-amd64.exe"
  $exe = Join-Path $env:TEMP "python-$PyVer-amd64.exe"
  Write-Host '  ダウンロードしています…'
  try {
    (New-Object Net.WebClient).DownloadFile($url, $exe)
    Log "downloaded $url"
    Write-Host '  入れています…（1〜3分）'
    $instArgs = '/quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1 InstallLauncherAllUsers=0 Include_test=0 Include_doc=0 AssociateFiles=0 Shortcuts=0'
    $p = Start-Process -FilePath $exe -ArgumentList $instArgs -Wait -PassThru
    Log ("installer exit {0}" -f $p.ExitCode)
  } catch { Log ("python install error: {0}" -f $_) }
  $py = Find-Python
}
if ($py) {
  $R_PY = '○'
  $v = & $py -c 'import platform; print(platform.python_version())' 2>$null
  Write-Host ("  OK（{0}）" -f $v)
  Log "python $py"
} else { Write-Host '  入れられませんでした' }

# 2. 部品7つ
Step '2/3 スキルが使う部品（7つ）'
if ($py) {
  $o = & $py -m pip install --disable-pip-version-check --no-warn-script-location @Pkgs 2>&1
  $o | ForEach-Object { Log ([string]$_) }
  & $py -c $Imports 2>$null
  if ($LASTEXITCODE -eq 0) { $R_PKG = '○'; Write-Host '  OK' } else { Write-Host '  一部が入りませんでした' }
} else { Write-Host '  Python が無いので飛ばしました' }

# 3. アプリの確認（入れるのはオリエンの画面の手順で）
Step '3/3 仕上げ'
if (-not $env:MAIOS_SETUP_SKIP_APPS) {
  $claude = @(
    (Join-Path $env:LOCALAPPDATA 'AnthropicClaude\claude.exe'),
    (Join-Path $env:LOCALAPPDATA 'Programs\Claude\Claude.exe'),
    (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Anthropic\Claude.lnk'),
    (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Claude.lnk')
  ) | Where-Object { Test-Path $_ }
  $appx = $null; try { $appx = Get-AppxPackage -Name '*Claude*' -ErrorAction SilentlyContinue } catch {}
  if ($claude -or $appx) { $R_APP = '○' } else { $R_APP = '△' }
  $drive = @((Join-Path $env:ProgramFiles 'Google\Drive File Stream')) | Where-Object { Test-Path $_ }
  if ($drive) { $R_DRIVE = '○' } else { $R_DRIVE = '×' }
}
Write-Host '  OK'
Log "result py=$R_PY pkg=$R_PKG app=$R_APP drive=$R_DRIVE"

Write-Host ''
Write-Host '=============================='
Write-Host ' MAIOS セットアップの結果'
Write-Host '=============================='
Write-Host (' {0} Python' -f $R_PY)
Write-Host (' {0} 部品7つ' -f $R_PKG)
if ($R_APP -ne '－') { Write-Host (' {0} Claude アプリ' -f $R_APP) }
if ($R_DRIVE -ne '－') { Write-Host (' {0} Google ドライブ' -f $R_DRIVE) }
Write-Host '=============================='
if ($R_PY -eq '○' -and $R_PKG -eq '○' -and $R_DRIVE -ne '×') {
  Write-Host ' 完了です。Claude アプリを一度終了して開き直し、次の手順へ進んでください。'
  if ($R_APP -eq '△') { Write-Host ' Claude アプリがまだ見つかりません。入れていなければ、画面の手順で入れてください。' }
  $global:LASTEXITCODE = 0
} else {
  if ($R_DRIVE -eq '×') { Write-Host ' Google ドライブ（パソコン版）が入っていません。事前案内の手順で入れてから、もう一度この1行を貼ってください。' }
  Write-Host ' × があります。手を挙げて、この画面を AMG に見せてください。'
  $global:LASTEXITCODE = 1
}
