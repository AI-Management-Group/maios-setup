#!/bin/bash
# MAIOS オリエン用の一括セットアップ（Mac）
#
# オリエンの場で、ターミナルに次の1行を貼り付けて実行する（管理者のパスワードは要らない）：
#   curl -fsSL https://raw.githubusercontent.com/AI-Management-Group/maios-setup/main/mac.sh | bash
#
# やること（何度実行しても同じ結果になる）：
#   1. uv（Python を入れる道具）を ~/.local/bin に入れる
#   2. Python 3.12 を ~/.local/bin/python3 に入れる（MAIOS のフックはここを PATH に無くても拾う）
#   3. スキルが使う部品7つを入れる（python-pptx openpyxl pandas PyYAML jpholiday Pillow playwright）
#   4. ターミナルから python3 が使えるよう ~/.zprofile と ~/.zshrc に1行足す
#   5. Claude アプリと Google ドライブ（パソコン版）があるかを確かめる。どちらもこのスクリプトでは入れない（事前案内で各自が入れる）
#      Google ドライブは必須＝無ければ ❌。Claude アプリはこの1行を走らせている間に入れてもよいので、無くても ⚠️ にとどめる
# 最後に ✅／❌ の一覧を出す。記録は ~/Library/Logs/maios_setup.log。
#
# 環境変数（AMG・テスト用）：MAIOS_SETUP_SKIP_APPS=1 でアプリの確認を省く
set -u

PKGS="python-pptx openpyxl pandas PyYAML jpholiday Pillow playwright"
IMPORTS="import pptx, openpyxl, pandas, yaml, jpholiday, PIL, playwright"
BIN="$HOME/.local/bin"
PY="$BIN/python3"
LOG="$HOME/Library/Logs/maios_setup.log"
mkdir -p "$BIN" "$(dirname "$LOG")" 2>/dev/null
: >>"$LOG" 2>/dev/null || LOG=/dev/null
export PATH="$BIN:$PATH"

log() { printf '%s %s\n' "$(date '+%F %T')" "$*" >>"$LOG"; }
step() { printf '\n▶ %s\n' "$*"; log "== $*"; }
run() { log "\$ $*"; "$@" >>"$LOG" 2>&1; }

R_PY="❌" R_PKG="❌" R_APP="－" R_DRIVE="－"

printf '\nMAIOS のセットアップを始めます（3〜5分）。終わるまで、この画面を閉じないでください。\n'
log "start $(sw_vers -productVersion 2>/dev/null) $(uname -m)"

# 1. uv
step "1/4 Python を入れる道具（uv）"
UV="$(command -v uv 2>/dev/null || true)"
if [ -z "$UV" ]; then
  if curl -LsSf https://astral.sh/uv/install.sh 2>>"$LOG" | env UV_INSTALL_DIR="$BIN" INSTALLER_NO_MODIFY_PATH=1 sh >>"$LOG" 2>&1; then
    UV="$BIN/uv"
  fi
fi
if [ -n "$UV" ] && [ -x "$UV" ]; then echo "  OK"; else echo "  入れられませんでした（ネットワークを確認してください）"; fi

# 2. Python
step "2/4 Python 3.12"
py_ok() { [ -x "$PY" ] && "$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' >/dev/null 2>&1; }
if ! py_ok && [ -n "$UV" ] && [ -x "$UV" ]; then
  run "$UV" python install 3.12 --default
fi
if py_ok; then
  # uv の Python は「外から部品を足さない」印（EXTERNALLY-MANAGED）が付いている。MAIOS 専用の Python なので外して、スキルの pip も通るようにする
  MARK="$("$PY" -c 'import os, sysconfig; print(os.path.join(sysconfig.get_path("stdlib"), "EXTERNALLY-MANAGED"))' 2>/dev/null)"
  [ -n "$MARK" ] && [ -f "$MARK" ] && run rm -f "$MARK"
  R_PY="✅"
  echo "  OK（$("$PY" -c 'import platform; print(platform.python_version())' 2>/dev/null)）"
else
  echo "  入れられませんでした"
fi

# 3. 部品7つ
step "3/4 スキルが使う部品（7つ）"
if [ "$R_PY" = "✅" ]; then
  # shellcheck disable=SC2086
  if [ -n "$UV" ] && [ -x "$UV" ]; then
    run "$UV" pip install --python "$PY" --system --break-system-packages $PKGS
  else
    run "$PY" -m pip install --disable-pip-version-check --break-system-packages $PKGS
  fi
  if "$PY" -c "$IMPORTS" >>"$LOG" 2>&1; then R_PKG="✅"; echo "  OK"; else echo "  一部が入りませんでした"; fi
else
  echo "  Python が無いので飛ばしました"
fi

# 4. ターミナルの PATH（新しく開くターミナルから python3 が ~/.local/bin のものになる）
step "4/4 仕上げ"
LINE='export PATH="$HOME/.local/bin:$PATH"  # MAIOS'
for rc in "$HOME/.zprofile" "$HOME/.zshrc"; do
  grep -qs 'export PATH="$HOME/.local/bin:$PATH"' "$rc" || printf '\n%s\n' "$LINE" >>"$rc"
done
echo "  OK"

# アプリの確認（入れるのはオリエンの画面の手順で）
if [ -z "${MAIOS_SETUP_SKIP_APPS:-}" ]; then
  if [ -d /Applications/Claude.app ] || [ -d "$HOME/Applications/Claude.app" ]; then R_APP="✅"; else R_APP="⚠️"; fi
  if [ -d "/Applications/Google Drive.app" ] || [ -d "$HOME/Applications/Google Drive.app" ]; then R_DRIVE="✅"; else R_DRIVE="❌"; fi
fi
log "result py=$R_PY pkg=$R_PKG app=$R_APP drive=$R_DRIVE python=$PY"

printf '\n==============================\n'
printf ' MAIOS セットアップの結果\n'
printf '==============================\n'
printf ' %s Python\n' "$R_PY"
printf ' %s 部品7つ\n' "$R_PKG"
[ "$R_APP" = "－" ] || printf ' %s Claude アプリ\n' "$R_APP"
[ "$R_DRIVE" = "－" ] || printf ' %s Google ドライブ\n' "$R_DRIVE"
printf '==============================\n'
if [ "$R_PY" = "✅" ] && [ "$R_PKG" = "✅" ] && [ "$R_DRIVE" != "❌" ]; then
  printf ' 完了です。Claude アプリを一度終了して開き直し、次の手順へ進んでください。\n'
  [ "$R_APP" = "⚠️" ] && printf ' Claude アプリがまだ見つかりません。入れていなければ、画面の手順で入れてください。\n'
  exit 0
fi
[ "$R_DRIVE" = "❌" ] && printf ' Google ドライブ（パソコン版）が入っていません。事前案内の手順で入れてから、もう一度この1行を貼ってください。\n'
printf ' ❌ があります。手を挙げて、この画面を AMG に見せてください。\n'
exit 1
