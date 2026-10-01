# MAIOS オリエン用の一括セットアップ（Windows）
#
# オリエンの場で、PowerShell（管理者ではない普通の PowerShell）に次の1行を貼り付けて実行する：
#   iex ([Text.Encoding]::UTF8.GetString((New-Object Net.WebClient).DownloadData('https://raw.githubusercontent.com/AI-Management-Group/maios-setup/main/win.ps1')))
#   （文字化けしないよう、UTF-8 として読み込んでから実行する形にしている）
#
# やること（何度実行しても同じ結果になる・管理者権限は要らない）：
#   1. Python 3.12 を「この利用者だけ」に入れる（%LOCALAPPDATA%\Programs\Python\Python312。MAIOS のフックはここを PATH に無くても拾う）
#      ARM 版 Windows でも x64 版を入れる（部品の配布物が x64 にしか無いものがあるため。Windows 11 は x64 をそのまま動かせる）
#   2. スキルが使う Python の部品を入れる（同梱スキルのスクリプトが読み込むもの全部。一覧は下の $Pkgs。
#      動画・文字認識など、別のソフトも要る部品は入れない＝tools/setup/README.md の「条件付きの部品」）
#      あわせて、この利用者の環境変数 PYTHONUTF8=1 を設定する（Windows の Python は既定の文字コードが cp932 で、
#      スキルのスクリプトが日本語を読み書きする時に化けたり落ちたりするため。Python が常に UTF-8 を使うようになる）
#   3. Node.js（JavaScript のプログラムを動かすソフト）と、その部品2つ（pptxgenjs playwright）を入れる
#      リッチなデックを組む経路と、ブラウザで表示を測るスクリプトが使う。Node.js 20 以上が既にあればそれを使い、
#      無ければ %LOCALAPPDATA%\maios\node に入れて、この利用者の PATH に足す（公式の配布物を、チェックサムで中身を確かめてから展開する）。
#      部品は %LOCALAPPDATA%\maios\node_parts に入れる（既にある Node.js の中身には触れない）。
#      どのフォルダに置いたスクリプトからも部品が見つかるよう、この利用者の環境変数 NODE_PATH にその場所を入れる
#   4. 画面の確認に使うブラウザ（Playwright の Chromium）を入れる。普段使いの Chrome とは別物で、スクリプトだけが使う
#   5. Claude アプリと Google ドライブ（パソコン版）があるかを確かめる。どちらもこのスクリプトでは入れない（事前案内で各自が入れる）
#      Google ドライブは必須＝無ければ ×。Claude アプリはこの1行を走らせている間に入れてもよいので、無くても △ にとどめる
# Git は入れない（Claude Code は Git が無ければ PowerShell で動く・公式 docs）。
# 最後に ○／× の一覧を出す（日本語の Windows の画面は絵文字を出せないことがあるため記号にする）。記録は %LOCALAPPDATA%\maios_setup.log。
#
# 環境変数（AMG・テスト用）：MAIOS_SETUP_FORCE_PYTHON=1 で既存の Python があっても入れる／MAIOS_SETUP_FORCE_NODE=1 で既存の Node.js があっても入れる／MAIOS_SETUP_SKIP_APPS=1 でアプリの確認を省く

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}

$PyVer = '3.12.10'
# $Pkgs を変えたら $Imports と Mac 版（maios_setup_mac.sh の PKGS・IMPORTS）も同じにする（tests/test_scripts_portable.py が照らす）
$Pkgs = @('python-pptx', 'openpyxl', 'pandas', 'PyYAML', 'jpholiday', 'Pillow', 'playwright', 'reportlab', 'numpy', 'matplotlib', 'seaborn', 'tabulate', 'beautifulsoup4', 'lxml', 'requests')
$Imports = 'import pptx, openpyxl, pandas, yaml, jpholiday, PIL, playwright, reportlab, numpy, matplotlib, seaborn, tabulate, bs4, lxml, requests'
$NodeVer = 'v24.21.0'
$NodeSha = '158f7685b44de51f6c0df1d153526cbcd3e1bc739a8dfc607721cef75de9e541'
$NodePkgs = @('pptxgenjs', 'playwright')
$NodeRequires = "require('pptxgenjs'); require('playwright')"
$MaiosDir = Join-Path $env:LOCALAPPDATA 'maios'
$NodeHome = Join-Path $MaiosDir 'node'
$NodeParts = Join-Path $MaiosDir 'node_parts'
$NodeModules = Join-Path $NodeParts 'node_modules'
# ブラウザが無くて起動に失敗した時も、playwright を必ず閉じてから終わる（閉じないと終わらないことがある）
$BrowserPy = "from playwright.sync_api import sync_playwright as s; c = s(); p = c.__enter__(); r = [0]; exec('try:\n p.chromium.launch().close()\nexcept Exception:\n r[0] = 1'); c.__exit__(None, None, None); raise SystemExit(r[0])"
$BrowserJs = "require('playwright').chromium.launch().then(b => b.close()).catch(e => { console.error(String(e)); process.exit(1); })"
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

function Install-Parts([string]$Exe) {
  # 部品を入れて、全部が読み込めれば $true
  $o = & $Exe -m pip install --disable-pip-version-check --no-warn-script-location @Pkgs 2>&1
  $o | ForEach-Object { Log ([string]$_) }
  & $Exe -c $Imports 2>$null | Out-Null
  return ($LASTEXITCODE -eq 0)
}

function Repair-Pip([string]$Exe) {
  # Python に同梱されている pip（Lib\ensurepip\_bundled の中）を使って、pip を入れ直す。
  # 同梱の pip を PYTHONPATH の先頭に置き、python -m pip の形で動かす（Windows の pip は「python -m pip」以外の形では自分自身を入れ直さない）
  $keep = $env:PYTHONPATH
  try {
    $dir = Join-Path (Split-Path $Exe) 'Lib\ensurepip\_bundled'
    $whl = Get-ChildItem -Path $dir -Filter 'pip-*.whl' -ErrorAction Stop | Select-Object -First 1
    if ($whl) {
      $env:PYTHONPATH = $whl.FullName
      $o = & $Exe -m pip install --disable-pip-version-check --no-warn-script-location --force-reinstall --no-index --find-links $dir pip 2>&1
      $o | ForEach-Object { Log ([string]$_) }
    } else { Log 'repair pip: 同梱の pip が見つからない' }
  } catch { Log ('repair pip error: ' + $_) }
  $env:PYTHONPATH = $keep
}

function Test-Node([string]$Exe) {
  # 20 以上で、隣に npm（部品を入れる道具）があれば使える
  if (-not $Exe) { return $false }
  if (-not (Test-Path $Exe)) { return $false }
  if (-not (Test-Path (Join-Path (Split-Path $Exe) 'npm.cmd'))) { return $false }
  try {
    & $Exe -e 'process.exit(parseInt(process.versions.node) >= 20 ? 0 : 1)' 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0)
  } catch { return $false }
}

function Find-Node {
  $own = Join-Path $NodeHome 'node.exe'
  if (Test-Node $own) { return $own }
  if ($env:MAIOS_SETUP_FORCE_NODE) { return $null }
  $c = Get-Command node -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($c) { if (Test-Node $c.Source) { return $c.Source } }
  $pf = Join-Path $env:ProgramFiles 'nodejs\node.exe'
  if (Test-Node $pf) { return $pf }
  return $null
}

function Install-Node {
  # 公式の zip を取り、チェックサムで中身を確かめてから %LOCALAPPDATA%\maios\node に展開する
  $name = "node-$NodeVer-win-x64"
  $url = "https://nodejs.org/dist/$NodeVer/$name.zip"
  $zip = Join-Path $env:TEMP "$name.zip"
  try {
    (New-Object Net.WebClient).DownloadFile($url, $zip)
    Log "downloaded $url"
    $hash = (Get-FileHash -Path $zip -Algorithm SHA256).Hash.ToLower()
    if ($hash -ne $NodeSha) { Log "node: checksum mismatch $hash"; return }
    if (Test-Path $NodeHome) { Remove-Item $NodeHome -Recurse -Force }
    New-Item -ItemType Directory -Path $NodeHome -Force | Out-Null
    $done = $false
    $tar = Join-Path $env:SystemRoot 'System32\tar.exe'
    if (Test-Path $tar) {
      # Windows 付属の tar は zip も展開できる（Expand-Archive より速い）
      $o = & $tar -xf $zip -C $NodeHome --strip-components 1 2>&1
      $o | ForEach-Object { Log ([string]$_) }
      $done = Test-Path (Join-Path $NodeHome 'node.exe')
    }
    if (-not $done) {
      Log 'node: expand with Expand-Archive'
      Remove-Item $NodeHome -Recurse -Force
      Expand-Archive -Path $zip -DestinationPath $MaiosDir -Force
      Rename-Item -Path (Join-Path $MaiosDir $name) -NewName 'node'
    }
  } catch { Log ("node install error: {0}" -f $_) }
  Remove-Item $zip -Force -ErrorAction SilentlyContinue
}

function Add-UserPath([string]$Dir) {
  # この利用者の PATH に足す（既にあれば何もしない）。元の書き方（%USERPROFILE% などが入った形）は保つ
  try {
    $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Environment', $true)
    $cur = [string]$key.GetValue('Path', '', [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
    $parts = @($cur -split ';' | Where-Object { $_ })
    if ($parts -notcontains $Dir) {
      $key.SetValue('Path', ((@($Dir) + $parts) -join ';'), [Microsoft.Win32.RegistryValueKind]::ExpandString)
      Log "PATH += $Dir (User)"
    }
    $key.Close()
  } catch { Log ("PATH を設定できませんでした: {0}" -f $_) }
  if (($env:Path -split ';') -notcontains $Dir) { $env:Path = "$Dir;$env:Path" }
}

function Install-NodeParts([string]$Exe, [string[]]$Names) {
  # 部品を入れて、全部が読み込めれば $true
  $npm = Join-Path (Split-Path $Exe) 'npm.cmd'
  New-Item -ItemType Directory -Path $NodeParts -Force | Out-Null
  $o = & $npm install --prefix $NodeParts --no-audit --no-fund --loglevel=error @Names 2>&1
  $o | ForEach-Object { Log ([string]$_) }
  & $Exe -e $NodeRequires 2>$null | Out-Null
  return ($LASTEXITCODE -eq 0)
}

$R_PY = '×'; $R_PKG = '×'; $R_NODE = '×'; $R_BROWSER = '△'; $R_APP = '－'; $R_DRIVE = '－'
$OldNode = $false
Write-Host ''
Write-Host 'MAIOS のセットアップを始めます（10分ほど）。終わるまで、この画面を閉じないでください。'
Log ("start {0} {1} ps{2}" -f [Environment]::OSVersion.VersionString, $env:PROCESSOR_ARCHITECTURE, $PSVersionTable.PSVersion)

# 1. Python
Step '1/5 Python 3.12'
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

# 2. スキルが使う Python の部品
Step '2/5 スキルが使う部品'
if ($py) {
  $ok = Install-Parts $py
  if (-not $ok) {
    # Python を入れた直後は、pip（部品を入れる道具）の準備が終わっていなかったり、欠けていたりすることがある。
    # 少し待ってやり直し、それでも駄目なら Python に同梱の pip で入れ直してから、もう一度試す
    Write-Host '  やり直しています…'
    Log 'retry: wait'
    Start-Sleep -Seconds 8
    $ok = Install-Parts $py
  }
  if (-not $ok) {
    Log 'retry: repair pip'
    Repair-Pip $py
    $ok = Install-Parts $py
  }
  if ($ok) { $R_PKG = '○'; Write-Host '  OK' } else { Write-Host '  一部が入りませんでした' }
} else { Write-Host '  Python が無いので飛ばしました' }
# Python の文字コードを UTF-8 に揃える（この利用者だけ。次に開くアプリから効く）
try {
  [Environment]::SetEnvironmentVariable('PYTHONUTF8', '1', 'User')
  $env:PYTHONUTF8 = '1'
  Log 'PYTHONUTF8=1 (User)'
} catch { Log ('PYTHONUTF8 を設定できませんでした: ' + $_) }

# 3. Node.js と部品2つ
Step '3/5 Node.js と部品（2つ）'
$node = Find-Node
if (-not $node) {
  Write-Host '  ダウンロードしています…'
  Install-Node
  $node = Find-Node
}
if ($node) {
  $nodeDir = Split-Path $node
  if ($nodeDir -eq $NodeHome) {
    Add-UserPath $nodeDir
    # 全利用者向けに古い Node.js が入っている PC は、Windows がそちらを先に使う（利用者の PATH より先に見るため）
    $sysNode = Join-Path $env:ProgramFiles 'nodejs\node.exe'
    if ((Test-Path $sysNode) -and -not (Test-Node $sysNode)) { $OldNode = $true; Log "old node: $sysNode" }
  }
  if (($env:Path -split ';') -notcontains $nodeDir) { $env:Path = "$nodeDir;$env:Path" }
  # どのフォルダのスクリプトからも部品が見つかるよう、探す場所を利用者の環境変数に入れる（次に開くアプリから効く）
  try {
    $curNp = [Environment]::GetEnvironmentVariable('NODE_PATH', 'User')
    if (-not $curNp) { $newNp = $NodeModules } elseif (($curNp -split ';') -contains $NodeModules) { $newNp = $curNp } else { $newNp = "$NodeModules;$curNp" }
    [Environment]::SetEnvironmentVariable('NODE_PATH', $newNp, 'User')
    Log "NODE_PATH=$newNp (User)"
  } catch { Log ('NODE_PATH を設定できませんでした: ' + $_) }
  $env:NODE_PATH = $NodeModules
  # Node.js の playwright は、Python の playwright と同じ系列の版にそろえる（同じブラウザを共用できる）
  $names = $NodePkgs
  if ($py) {
    $pwv = $null
    try { $pwv = & $py -c "from importlib.metadata import version; print('.'.join(version('playwright').split('.')[:2]))" 2>$null } catch {}
    $pwv = ([string](@($pwv)[0])).Trim()
    if ($pwv -match '^\d+\.\d+$') { $names = @('pptxgenjs', ('playwright@~{0}.0' -f $pwv)) }
  }
  $ok = Install-NodeParts $node $names
  if (-not $ok) {
    Log 'retry: npm'
    Start-Sleep -Seconds 5
    $ok = Install-NodeParts $node $NodePkgs
  }
  if ($ok) {
    $R_NODE = '○'
    $v = & $node -v 2>$null
    Write-Host ("  OK（{0}）" -f $v)
    Log "node $node"
  } else { Write-Host '  部品が入りませんでした' }
} else { Write-Host '  入れられませんでした（ネットワークを確認してください）' }

# 4. 画面の確認に使うブラウザ（Playwright の Chromium）。無くても Chrome で代わりが利くスキルが多いので、入らなくても △ にとどめる
Step '4/5 画面の確認に使うブラウザ'
$bPy = $false; $bJs = $false
if ($R_PKG -eq '○') {
  & $py -c $BrowserPy 2>$null | Out-Null
  if ($LASTEXITCODE -ne 0) {
    Write-Host '  ダウンロードしています…（1〜3分）'
    $o = & $py -m playwright install chromium 2>&1
    $o | ForEach-Object { Log ([string]$_) }
    & $py -c $BrowserPy 2>$null | Out-Null
  }
  $bPy = ($LASTEXITCODE -eq 0)
}
if ($R_NODE -eq '○') {
  & $node -e $BrowserJs 2>$null | Out-Null
  if ($LASTEXITCODE -ne 0) {
    $o = & $node (Join-Path $NodeModules 'playwright\cli.js') install chromium 2>&1
    $o | ForEach-Object { Log ([string]$_) }
    & $node -e $BrowserJs 2>$null | Out-Null
  }
  $bJs = ($LASTEXITCODE -eq 0)
}
if ($bPy -and $bJs) { $R_BROWSER = '○'; Write-Host '  OK' } else { Log "browser py=$bPy js=$bJs"; Write-Host '  入りませんでした（後からでも入れられます）' }

# 5. アプリの確認（入れるのはオリエンの画面の手順で）
Step '5/5 仕上げ'
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
Log "result py=$R_PY pkg=$R_PKG node=$R_NODE browser=$R_BROWSER app=$R_APP drive=$R_DRIVE"

Write-Host ''
Write-Host '=============================='
Write-Host ' MAIOS セットアップの結果'
Write-Host '=============================='
Write-Host (' {0} Python' -f $R_PY)
Write-Host (' {0} スキルが使う部品' -f $R_PKG)
Write-Host (' {0} Node.js と部品2つ' -f $R_NODE)
Write-Host (' {0} 画面確認用のブラウザ' -f $R_BROWSER)
if ($R_APP -ne '－') { Write-Host (' {0} Claude アプリ' -f $R_APP) }
if ($R_DRIVE -ne '－') { Write-Host (' {0} Google ドライブ' -f $R_DRIVE) }
Write-Host '=============================='
if ($R_PY -eq '○' -and $R_PKG -eq '○' -and $R_NODE -eq '○' -and $R_DRIVE -ne '×') {
  Write-Host ' 完了です。Claude アプリを一度終了して開き直し、次の手順へ進んでください。'
  if ($R_APP -eq '△') { Write-Host ' Claude アプリがまだ見つかりません。入れていなければ、画面の手順で入れてください。' }
  if ($OldNode) { Write-Host ' 古い Node.js が入っている PC です。入れ替えが要るので、AMG に伝えてください（ほかは使えます）。' }
  if ($R_BROWSER -eq '△') { Write-Host ' 画面確認用のブラウザが入りませんでした。このままでも始められます。後でもう一度この1行を貼ると入ります。' }
  $global:LASTEXITCODE = 0
} else {
  if ($R_DRIVE -eq '×') { Write-Host ' Google ドライブ（パソコン版）が入っていません。事前案内の手順で入れてから、もう一度この1行を貼ってください。' }
  Write-Host ' × があります。手を挙げて、この画面を AMG に見せてください。'
  $global:LASTEXITCODE = 1
}
