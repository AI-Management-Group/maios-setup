# maios-setup

MAIOS の導入で使う一括セットアップ（Python 3.12・Node.js と、スキルが使う部品を、管理者権限なしで利用者のフォルダに入れる）。

- Mac（ターミナル）：`t=$(mktemp) && curl -fsSL https://raw.githubusercontent.com/AI-Management-Group/maios-setup/main/mac.sh -o "$t" && { [ "$(tail -n 1 "$t")" = "# maios-setup:end" ] || { echo "取得が途中で切れました。もう一度貼ってください" >&2; false; }; } && bash "$t"; s=$?; rm -f "$t"; (exit $s)`
- Windows（PowerShell）：`iex ([Text.Encoding]::UTF8.GetString((New-Object Net.WebClient).DownloadData('https://raw.githubusercontent.com/AI-Management-Group/maios-setup/main/win.ps1')))`

正本は非公開のリポジトリにあり、ここは写し。直接直さない。
